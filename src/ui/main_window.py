"""
Main Window UI - PyQt5
"""
import sys
from pathlib import Path
from typing import Optional
import logging
import shutil
import os
import json

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
    QLabel, QLineEdit, QFileDialog, QProgressBar, QTextEdit,
    QTabWidget, QTableWidget, QTableWidgetItem, QComboBox,
    QSpinBox, QGroupBox, QMessageBox, QStatusBar, QDesktopWidget,
    QListWidget, QListWidgetItem, QAbstractItemView, QCheckBox, QGridLayout,
    QRadioButton, QHeaderView, QGraphicsDropShadowEffect, QStackedWidget,
    QSizePolicy, QScrollArea, QFormLayout, QFrame, QPlainTextEdit
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QEvent, QRect, QSize
from PyQt5.QtGui import QFont, QIcon, QColor, QFontMetricsF, QTextOption, QPainter, QTextFormat

from ..parsers import AlertLogParser, AWRParser, DatabaseInfoParser
from ..generators.comprehensive_report_generator import ComprehensiveHealthcareReportGenerator
from ..config import (
    APP_NAME, APP_VERSION, WINDOW_WIDTH, WINDOW_HEIGHT,
    NUM_DAYS_ALERT, OUTPUT_DIR, APPENDIX_OUTPUT_DIR, REPORT_OUTPUT_DIR,
    COLLECT_TOOL_DIR, GITHUB_TOOLS_API_URL, AUTO_SYNC_TOOLS
)
from ..utils import setup_logger, sanitize_filename
from ..utils.github_sync_worker import GitHubSyncWorker
from ..utils.generator_worker import GeneratorWorker
from ..utils.report_worker import ReportWorker
from ..utils.exawatcher_runner import ExaWatcherGraphGenerator
from ..utils.rules_manager import RulesManager

logger = setup_logger(__name__)

from ..parsers.alert_parser import AlertLogParser
from ..parsers.awr_parser import AWRParser
from ..parsers.database_info_parser import DatabaseInfoParser

# ─────────────────────────────────────────────────────────────────
#  CUSTOM WIDGETS: Code Editor with Line Numbers
# ─────────────────────────────────────────────────────────────────

class LineNumberArea(QWidget):
    def __init__(self, editor):
        super().__init__(editor)
        self.codeEditor = editor

    def sizeHint(self):
        return QSize(self.codeEditor.lineNumberAreaWidth(), 0)

    def paintEvent(self, event):
        self.codeEditor.lineNumberAreaPaintEvent(event)

class CodeEditor(QPlainTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.lineNumberArea = LineNumberArea(self)
        self.blockCountChanged.connect(self.updateLineNumberAreaWidth)
        self.updateRequest.connect(self.updateLineNumberArea)
        self.cursorPositionChanged.connect(self.highlightCurrentLine)
        self.updateLineNumberAreaWidth(0)

    def lineNumberAreaWidth(self):
        digits = 1
        max_value = max(1, self.blockCount())
        while max_value >= 10:
            max_value /= 10
            digits += 1
        # Added padding and margin
        space = 8 + self.fontMetrics().horizontalAdvance('9') * digits
        return space

    def updateLineNumberAreaWidth(self, _):
        self.setViewportMargins(self.lineNumberAreaWidth(), 0, 0, 0)

    def updateLineNumberArea(self, rect, dy):
        if dy:
            self.lineNumberArea.scroll(0, dy)
        else:
            self.lineNumberArea.update(0, rect.y(), self.lineNumberArea.width(), rect.height())
        if rect.contains(self.viewport().rect()):
            self.updateLineNumberAreaWidth(0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        cr = self.contentsRect()
        self.lineNumberArea.setGeometry(QRect(cr.left(), cr.top(), self.lineNumberAreaWidth(), cr.height()))

    def lineNumberAreaPaintEvent(self, event):
        painter = QPainter(self.lineNumberArea)
        # Sligthly darker background for the line numbers area
        painter.fillRect(event.rect(), QColor("#f5f5f5"))

        block = self.firstVisibleBlock()
        blockNumber = block.blockNumber()
        top = int(self.blockBoundingGeometry(block).translated(self.contentOffset()).top())
        bottom = top + int(self.blockBoundingRect(block).height())

        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                number = str(blockNumber + 1)
                painter.setPen(QColor("#888888"))
                painter.drawText(0, top, self.lineNumberArea.width() - 5, self.fontMetrics().height(),
                                 Qt.AlignRight, number)
            block = block.next()
            top = bottom
            bottom = top + int(self.blockBoundingRect(block).height())
            blockNumber += 1

    def highlightCurrentLine(self):
        extraSelections = []
        if not self.isReadOnly():
            selection = QTextEdit.ExtraSelection()
            lineColor = QColor(Qt.yellow).lighter(160)
            selection.format.setBackground(lineColor)
            selection.format.setProperty(QTextFormat.FullWidthSelection, True)
            selection.cursor = self.textCursor()
            selection.cursor.clearSelection()
            extraSelections.append(selection)
        self.setExtraSelections(extraSelections)

def standalone_parse_node(node_id: int, data_dir: str, num_days: int):
    """Standalone function for multiprocessing compatibility"""
    try:
        # 1. Alert Log
        parser_alert = AlertLogParser(data_dir, num_days)
        parser_alert.parse()
        
        # 2. AWR
        parser_awr = AWRParser(data_dir)
        parser_awr.parse()
        
        # 3. DB Info
        parser_db_info = DatabaseInfoParser(data_dir)
        parser_db_info.parse()
        
        db_info_data = parser_db_info.get_all_data()
        db_info_data['backup_schedule'] = parser_db_info.get_backup_schedule()
        
        return {
            "node_id": node_id,
            "data_dir": data_dir,
            "alert_data": parser_alert.get_data(),
            "awr_data": parser_awr.get_data(),
            "db_info_data": db_info_data
        }
    except Exception as e:
        return {"error": str(e)}

class ParseWorker(QThread):
    """Worker thread for parsing data"""
    
    progress = pyqtSignal(str, int)  # Progress message and percentage
    finished = pyqtSignal(dict)  # Finished with data
    error = pyqtSignal(str)  # Error message
    
    def __init__(self, log_folders: list, num_days: int):
        super().__init__()
        self.log_folders = log_folders
        self.num_days = num_days
        self.parsed_data = {}
    
    def run(self):
        """Parse data in parallel across all nodes"""
        from concurrent.futures import ThreadPoolExecutor
        try:
            total_nodes = len(self.log_folders)
            if total_nodes == 0:
                self.error.emit("No folders selected.")
                return
            
            self.progress.emit("Initializing parallel parsing...", 5)
            
            results = []
            from concurrent.futures import ProcessPoolExecutor
            
            with ProcessPoolExecutor(max_workers=min(4, total_nodes)) as executor:
                futures = []
                for idx, folder in enumerate(self.log_folders):
                    futures.append(executor.submit(standalone_parse_node, idx + 1, folder, self.num_days))
                
                # Monitor progress more dynamically
                completed = 0
                for f in futures:
                    res = f.result()
                    results.append(res)
                    completed += 1
                    self.progress.emit(f"Parsed Node {completed}/{total_nodes}...", 5 + int((completed / total_nodes) * 90))
            
            nodes = []
            for idx, res in enumerate(results):
                if res.get('error'):
                    self.error.emit(f"Node {idx + 1}: {res['error']}")
                    return
                # Compile unified structure
                alrt = res.get('alert_data', {})
                awr = res.get('awr_data', {})
                instance_name = alrt.get('instance_name') or awr.get('instance_name') or f"NODE{idx+1}"
                
                nodes.append({
                    'node_id': idx + 1,
                    'data_dir': self.log_folders[idx],
                    'instance_name': instance_name,
                    'alerts': res['alert_data'],
                    'awr': res['awr_data'],
                    'database_info': res['db_info_data']
                })
            
            # Find DB name from first available node that has it
            db_name = 'Unknown'
            for n in nodes:
                name = n.get('awr', {}).get('db_name')
                if name and name != 'Unknown':
                    db_name = name
                    break
            
            self.parsed_data = {
                'db_name': db_name,
                'nodes': nodes
            }
            
            self.progress.emit("All data parsed successfully!", 100)
            self.finished.emit(self.parsed_data)
        
        except Exception as e:
            self.error.emit(f"Parallel parsing failed: {str(e)}")



class MainWindow(QMainWindow):
    """Main Application Window"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.resize(max(WINDOW_WIDTH, 1280), max(WINDOW_HEIGHT, 860))
        self.setMinimumSize(1100, 780)
        self._center()
        
        # Set Application Icon
        icon_path = self._get_resource_path("styles/app_icon.ico")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        
        self.log_folders = []
        self.oswbb_push_folders = []
        self.exa_push_folders = [] # New list for ExaWatcher
        self.exa_db_input_dir = None
        self.exa_cell_input_dir = None
        self.exa_output_dir = None
        self.parsed_data = None
        self.parse_worker = None
        self.github_worker = None
        self.has_auto_synced = False
        self.gen_mode = 'appendix' # 'appendix' or 'report'
        
        # Explicitly initialize Status Bar to prevent UI jumping
        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Ready")
        
        self._init_ui()
        self._apply_pointing_cursor(self) # Apply hand cursor to all buttons
        self._apply_modern_shadows()
        self._load_stylesheet()
        logger.info(f"Application started: {APP_NAME} v{APP_VERSION}")
    
    def _center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def _get_resource_path(self, relative_path: str) -> Path:
        """Get absolute path to resource, works for dev and for PyInstaller"""
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            return Path(sys._MEIPASS) / relative_path
        return Path(__file__).resolve().parent.parent.parent / relative_path

    def _load_stylesheet(self):
        """Load external QSS stylesheet"""
        try:
            qss_path = self._get_resource_path("styles/main.qss")
            logger.info(f"Loading stylesheet from: {qss_path}")
            if qss_path.exists():
                with open(qss_path, "r", encoding="utf-8") as f:
                    content = f.read()
                    self.setStyleSheet(content)
                    logger.info(f"Stylesheet loaded successfully ({len(content)} bytes)")
            else:
                logger.warning(f"Stylesheet not found: {qss_path}")
        except Exception as e:
            logger.error(f"Failed to load stylesheet: {e}")
    
    def _init_ui(self):
        central_widget = QWidget()
        # Set explicitly a background role class/name for CSS targeting if needed
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # --- LEFT SIDEBAR ---
        sidebar_container = QWidget()
        sidebar_container.setObjectName("sidebarContainer")
        sidebar_container.setFixedWidth(320) # Mở rộng bề ngang như yêu cầu
        
        sidebar_layout = QVBoxLayout(sidebar_container)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)
        
        self.sidebar = QListWidget()
        self.sidebar.setObjectName("sidebarMenu")
        self.sidebar.setFocusPolicy(Qt.NoFocus) # Remove dotted border on selection
        self.sidebar.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff) # Ngăn scrollbar ngang hiện ra
        
        # Add Logo/Branding widget
        logo_widget = QWidget()
        logo_layout = QGridLayout(logo_widget)
        logo_layout.setContentsMargins(16, 24, 16, 16)
        logo_layout.setSpacing(12)
        
        # Icon Label
        icon_lbl = QLabel()
        icon_path = self._get_resource_path("styles/app_icon.ico")
        if icon_path.exists():
            from PyQt5.QtGui import QPixmap
            pixmap = QPixmap(str(icon_path))
            if not pixmap.isNull():
                icon_lbl.setPixmap(pixmap.scaled(54, 54, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        
        # App Name: Oracle HC [newline] Generator
        # APP_NAME is "Oracle HC Generator", we want a newline before the last word
        display_name = APP_NAME.rsplit(" ", 1)
        name_text = "\n".join(display_name) if len(display_name) > 1 else APP_NAME
        
        app_name_lbl = QLabel(name_text)
        app_name_lbl.setWordWrap(True)
        app_name_lbl.setStyleSheet("font-size: 26px; font-weight: 800; color: #1A1A1A; line-height: 1.0;")
        
        version_lbl = QLabel(f"v{APP_VERSION}")
        version_lbl.setStyleSheet("font-size: 11px; font-weight: 600; color: #0067C0; background: rgba(0,103,192,0.1); padding: 2px 6px; border-radius: 4px;")
        
        # Layout arrangement: [Icon] [Name] [Version]
        # Icon spans 2 rows to center against the multi-line text
        logo_layout.addWidget(icon_lbl, 0, 0, 2, 1, Qt.AlignLeft | Qt.AlignVCenter)
        logo_layout.addWidget(app_name_lbl, 0, 1, 2, 1, Qt.AlignLeft | Qt.AlignVCenter)
        logo_layout.addWidget(version_lbl, 0, 2, 1, 1, Qt.AlignRight | Qt.AlignTop)
        
        # Set column stretch to keep logo and name together
        logo_layout.setColumnStretch(1, 1)
        
        sidebar_layout.addWidget(logo_widget)
        
        # Main sidebar for core features
        self.sidebar.addItem(QListWidgetItem("⚙ Appendix Generator"))
        self.sidebar.addItem(QListWidgetItem("📊 OSWBB Graph Generator"))
        self.sidebar.addItem(QListWidgetItem("📈 ExaWatcher Graph Generator"))
        self.sidebar.addItem(QListWidgetItem("📄 Merge Documents"))
        
        # Bottom Sidebar for support tools
        self.sidebar_footer = QListWidget()
        self.sidebar_footer.setObjectName("sidebarMenu")
        self.sidebar_footer.setFocusPolicy(Qt.NoFocus)
        self.sidebar_footer.setFixedHeight(135) # Tăng thêm để chắc chắn không bị cuộn
        self.sidebar_footer.setFrameShape(QListWidget.NoFrame) # Bỏ viền để bớt chiếm diện tích
        self.sidebar_footer.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.sidebar_footer.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.sidebar_footer.setContentsMargins(0, 0, 0, 0)
        self.sidebar_footer.setSpacing(0)
        
        self.sidebar_footer.addItem(QListWidgetItem("⚙ Recommendation Setting"))
        self.sidebar_footer.addItem(QListWidgetItem("📦 Collection Tools"))
        self.sidebar_footer.currentRowChanged.connect(self._on_footer_tab_changed)
        
        # Connect main sidebar to also support switching logic
        self.sidebar.currentRowChanged.connect(self._on_main_tab_changed)
        
        sidebar_layout.addWidget(self.sidebar)
        sidebar_layout.addStretch() 
        sidebar_layout.addWidget(self.sidebar_footer)
        
        # Connect signals correctly with unique handlers
        self.sidebar.currentRowChanged.connect(self._on_main_tab_changed)
        self.sidebar_footer.currentRowChanged.connect(self._on_footer_tab_changed)
        
        # Chữ ký người tạo
        signature_label = QLabel("Developed by Victor Le")
        signature_label.setObjectName("signatureLabel")
        signature_label.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(signature_label)
        
        main_layout.addWidget(sidebar_container)
        
        # --- RIGHT CONTENT AREA ---
        content_container = QWidget()
        content_layout = QVBoxLayout(content_container)
        content_layout.setContentsMargins(30, 20, 30, 10) 
        
        # Header Area with Title
        self.section_title = QLabel("APPENDIX GENERATOR")
        self.section_title.setObjectName("sectionTitle")
        content_layout.addWidget(self.section_title)
        
        # Main Stack for different pages
        self.stack = QStackedWidget()
        self.stack.addWidget(self._create_main_tab())       # Index 0
        self.stack.addWidget(self._create_oswbb_tab())      # Index 1
        self.stack.addWidget(self._create_exawatcher_tab()) # Index 2 (ExaWatcher)
        self.stack.addWidget(self._create_merge_tab())          # Index 3
        self.stack.addWidget(self._create_recommendation_tab()) # Index 4 (New)
        self.stack.addWidget(self._create_tools_tab())          # Index 5 (Shifted)
        
        content_layout.addWidget(self.stack)
        main_layout.addWidget(content_container)
        
        # Set initial selection AFTER connections
        self.sidebar.setCurrentRow(0)
        
        # Initial Title Update
        self._on_sidebar_changed(0)

    def _apply_pointing_cursor(self, widget):
        """Recursively apply pointing hand cursor to all buttons"""
        for child in widget.findChildren(QPushButton):
            child.setCursor(Qt.PointingHandCursor)

    def _on_sidebar_changed(self, index):
        """Redundant method, keeping for stability or removing completely"""
        pass

    def _on_main_tab_changed(self, index):
        """Handle switching from the main sidebar items"""
        if index != -1:
            # Block signals temporarily to prevent infinite loop of deselecting
            self.sidebar_footer.blockSignals(True)
            self.sidebar_footer.clearSelection()
            self.sidebar_footer.setCurrentRow(-1)
            self.sidebar_footer.blockSignals(False)
            
            # Index 0-2 in main sidebar corresponds directly to index 0-3 in stack? No, wait.
            # Main sidebar indices: 0 (Main), 1 (OSWBB), 2 (ExaWatcher), 3 (Merge)
            self.stack.setCurrentIndex(index)
            # Update title
            titles = ["APPENDIX GENERATOR", "OSWBB GRAPH GENERATOR", "EXAWATCHER GRAPH GENERATOR", "MERGE DOCUMENTS"]
            if index < len(titles):
                self.section_title.setText(titles[index])

    def _on_footer_tab_changed(self, index):
        """Handle switching from the footer sidebar items"""
        if index != -1:
            self.sidebar.blockSignals(True)
            self.sidebar.clearSelection()
            self.sidebar.setCurrentRow(-1)
            self.sidebar.blockSignals(False)
            
            if index == 0:
                # "Recommendation Setting" is index 4 in stack
                self.stack.setCurrentIndex(4)
                self.section_title.setText("RECOMMENDATION SETTING")
            elif index == 1:
                # "Collection Tools" is index 5 in stack
                self.stack.setCurrentIndex(5)
                self.section_title.setText("COLLECTION TOOLS")
                
                # Auto-sync logic for Tools tab
                if AUTO_SYNC_TOOLS and not self.has_auto_synced:
                    self._on_sync_github_clicked()

    def _create_main_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(15, 10, 15, 10)
        
        # 1. Folders Group
        folder_group = QGroupBox("1. DATA SOURCES (Multi-Node)")
        folder_layout = QVBoxLayout()
        
        list_action_box = QHBoxLayout()
        btn_add_node = QPushButton("➕ Add Node Folder")
        btn_add_node.setObjectName("browse_btn")
        btn_add_node.clicked.connect(self._add_node_folder)
        
        btn_clear_nodes = QPushButton("✖ Clear All")
        btn_clear_nodes.setObjectName("clear_btn")
        btn_clear_nodes.clicked.connect(self._clear_nodes)
        
        list_action_box.addWidget(btn_add_node)
        list_action_box.addWidget(btn_clear_nodes)
        list_action_box.addStretch()
        folder_layout.addLayout(list_action_box)
        
        self.node_list_widget = QListWidget()
        self.node_list_widget.setSelectionMode(QAbstractItemView.NoSelection)
        self.node_list_widget.setAcceptDrops(True)
        self.node_list_widget.installEventFilter(self)
        folder_layout.addWidget(self.node_list_widget)
        
        folder_group.setLayout(folder_layout)
        layout.addWidget(folder_group)
        
        # 2. Settings Group - Refined Layout
        settings_group = QGroupBox("2. DOCUMENT SETTINGS")
        settings_layout = QGridLayout()
        settings_layout.setSpacing(10)
        
        # Column 0: Alert Logs
        settings_layout.addWidget(QLabel("Alert Logs (Days):"), 0, 0)
        self.num_days_spin = QSpinBox()
        self.num_days_spin.setButtonSymbols(QSpinBox.NoButtons) # Remove arrows
        self.num_days_spin.setValue(NUM_DAYS_ALERT)
        self.num_days_spin.setRange(1, 365)
        self.num_days_spin.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.num_days_spin.setAlignment(Qt.AlignCenter) # Center text to look better without buttons
        settings_layout.addWidget(self.num_days_spin, 1, 0)
        
        # Column 1: Document Font
        settings_layout.addWidget(QLabel("Document Font:"), 0, 1)
        self.font_combo = QComboBox()
        self.font_combo.addItems(["Times NR", "Calibri"])
        self.font_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        settings_layout.addWidget(self.font_combo, 1, 1)

        # Column 2: Database Role
        settings_layout.addWidget(QLabel("Database Role:"), 0, 2)
        self.role_combo = QComboBox()
        self.role_combo.addItems(["Primary", "Standby"])
        self.role_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        settings_layout.addWidget(self.role_combo, 1, 2)
        
        # Column 3: Report Language (New)
        settings_layout.addWidget(QLabel("Report Language:"), 0, 3)
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["Vietnamese", "English"])
        self.lang_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        settings_layout.addWidget(self.lang_combo, 1, 3)
        
        # Column 4: Custom Filename (Giving more width)
        settings_layout.addWidget(QLabel("Custom Filename (Optional):"), 0, 4)
        
        file_row = QHBoxLayout()
        self.filename_input = QLineEdit()
        self.filename_input.setPlaceholderText("Auto-generated if empty")
        
        self.clear_filename_btn = QPushButton("✖")
        self.clear_filename_btn.setObjectName("clear_btn")
        self.clear_filename_btn.clicked.connect(self.filename_input.clear)
        
        file_row.addWidget(self.filename_input)
        file_row.addWidget(self.clear_filename_btn)
        settings_layout.addLayout(file_row, 1, 4)
        
        # Adjusting column stretches
        settings_layout.setColumnStretch(0, 1) # Alert Logs (narrow)
        settings_layout.setColumnStretch(1, 3) # Font
        settings_layout.setColumnStretch(2, 3) # Role
        settings_layout.setColumnStretch(3, 3) # Language
        settings_layout.setColumnStretch(4, 5) # Filename
        
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # 3. Actions
        action_layout = QHBoxLayout()
        action_layout.setSpacing(15)
        
        self.generate_btn = QPushButton("🚀 GENERATE APPENDIX")
        self.generate_btn.setObjectName("main_action_btn")
        self.generate_btn.clicked.connect(self._on_generate_clicked)
        
        self.generate_report_btn = QPushButton("📝 GENERATE REPORT")
        self.generate_report_btn.setObjectName("main_action_btn")
        self.generate_report_btn.clicked.connect(self._on_generate_report_clicked)
        
        action_layout.addWidget(self.generate_btn, stretch=1)
        action_layout.addWidget(self.generate_report_btn, stretch=1)
        layout.addLayout(action_layout)
        
        # --- FIXED STATUS AREA ---
        status_layout = QHBoxLayout()
        self.appendix_progress = QProgressBar()
        self.appendix_progress.setObjectName("globalProgressBar")
        self.appendix_progress.setFixedHeight(20)
        self.appendix_progress.setTextVisible(True)
        self.appendix_progress.setFormat(" %p% ")
        self.appendix_progress.setAlignment(Qt.AlignCenter)
        self.appendix_progress.setValue(0)
        
        self.appendix_status_lbl = QLabel("READY")
        self.appendix_status_lbl.setObjectName("status_ready")
        self.appendix_status_lbl.setAlignment(Qt.AlignCenter)
        self.appendix_status_lbl.setFixedHeight(20)
        
        status_layout.setSpacing(5) # Narrow gap between bar and box
        status_layout.addWidget(self.appendix_progress, stretch=1)
        status_layout.addWidget(self.appendix_status_lbl)
        layout.addLayout(status_layout)
        
        # 4. Logs
        log_group = QGroupBox("CONSOLE LOG")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        widget.setLayout(layout)
        return widget

    def _create_oswbb_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8) # Reduced from 15
        layout.setContentsMargins(15, 10, 15, 10)
        
        # 1. SETUP GENERATOR
        config_group = QGroupBox("1. SETUP OSWBB SOURCES")
        config_layout = QGridLayout()
        config_layout.setContentsMargins(12, 12, 12, 10) 
        config_layout.setSpacing(10)
        
        config_layout.addWidget(QLabel("Input Log Folder (OSWBB):"), 0, 0)
        self.oswbb_input_dir = QLineEdit()
        self.oswbb_input_dir.setPlaceholderText("Select the OSWBB archive folder (oswcpuinfo, etc)...")
        self.oswbb_input_dir.setAcceptDrops(True)
        self.oswbb_input_dir.installEventFilter(self)
        config_layout.addWidget(self.oswbb_input_dir, 0, 1)
        
        btn_box_in = QHBoxLayout()
        btn_browse_in = QPushButton("📂")
        btn_browse_in.setObjectName("browse_btn")
        btn_browse_in.clicked.connect(self._browse_oswbb_in)
        
        btn_clear_in = QPushButton("✖")
        btn_clear_in.setObjectName("clear_btn")
        btn_clear_in.clicked.connect(self.oswbb_input_dir.clear)
        
        btn_box_in.addWidget(btn_browse_in)
        btn_box_in.addWidget(btn_clear_in)
        config_layout.addLayout(btn_box_in, 0, 2)
        
        config_layout.addWidget(QLabel("Output Images Folder:"), 1, 0)
        self.oswbb_output_dir = QLineEdit()
        self.oswbb_output_dir.setPlaceholderText("Empty = Default 'generated_files' in application folder")
        self.oswbb_output_dir.setAcceptDrops(True)
        self.oswbb_output_dir.installEventFilter(self)
        config_layout.addWidget(self.oswbb_output_dir, 1, 1)
        
        btn_box_out = QHBoxLayout()
        btn_browse_out = QPushButton("📂")
        btn_browse_out.setObjectName("browse_btn")
        btn_browse_out.clicked.connect(self._browse_oswbb_out)
        
        btn_clear_out = QPushButton("✖")
        btn_clear_out.setObjectName("clear_btn")
        btn_clear_out.clicked.connect(self.oswbb_output_dir.clear)
        
        btn_box_out.addWidget(btn_browse_out)
        btn_box_out.addWidget(btn_clear_out)
        config_layout.addLayout(btn_box_out, 1, 2)
        
        config_layout.addWidget(QLabel("Analysis Tool (JAR):"), 2, 0)
        self.oswbb_jar_select = QComboBox()
        self.oswbb_jar_select.addItems(["oswbba9020.jar", "oswbba.jar"])
        config_layout.addWidget(self.oswbb_jar_select, 2, 1)
        
        config_layout.setColumnStretch(1, 1)
        config_group.setLayout(config_layout)
        layout.addWidget(config_group)
        
        # 2. TARGET FOLDERS FOR SYNC (New Linear Style)
        push_group = QGroupBox("2. TARGET FOLDERS FOR SYNC")
        push_layout = QVBoxLayout()
        push_layout.setSpacing(10)
        
        # Action row for targets
        push_btn_row = QHBoxLayout()
        btn_add_target = QPushButton("➕ Add Target Folders")
        btn_add_target.setObjectName("browse_btn")
        btn_add_target.clicked.connect(self._on_oswbb_add_push_folder)
        
        btn_clear_target = QPushButton("✖ Clear All")
        btn_clear_target.setObjectName("clear_btn")
        btn_clear_target.clicked.connect(self._on_oswbb_clear_push_folders)
        
        self.push_mode_overwrite = QRadioButton("Overwrite")
        self.push_mode_timestamp = QRadioButton("Timestamp Subfolder")
        self.push_mode_overwrite.setMinimumWidth(100)
        self.push_mode_timestamp.setMinimumWidth(170)
        self.push_mode_overwrite.setChecked(True)
        
        push_btn_row.addWidget(btn_add_target)
        push_btn_row.addWidget(btn_clear_target)
        push_btn_row.addStretch()
        push_btn_row.addWidget(QLabel("Push Mode:"))
        push_btn_row.addWidget(self.push_mode_overwrite)
        push_btn_row.addWidget(self.push_mode_timestamp)
        push_layout.addLayout(push_btn_row)
        
        # Target list - Now using QListWidget for better reliability and drag-drop
        self.push_target_list = QListWidget()
        self.push_target_list.setObjectName("merge_list") # Reuse style
        self.push_target_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.push_target_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.push_target_list.setMinimumHeight(120) # Allow expansion
        # Enable OS-level drag-drop for folders
        self.push_target_list.setAcceptDrops(True)
        self.push_target_list.installEventFilter(self) # We will handle drops here
        push_layout.addWidget(self.push_target_list)
        
        # Cấu hình gọn gàng hơn
        push_layout.setContentsMargins(12, 10, 12, 12) 
        push_layout.setSpacing(10)
        
        push_group.setLayout(push_layout)
        layout.addWidget(push_group, 2) # Added stretch factor 2
        
        # 3. ACTIONS
        action_layout = QHBoxLayout()
        action_layout.setSpacing(15)
        self.gen_oswbb_btn = QPushButton("🖻 GENERATE IMAGES")
        self.gen_oswbb_btn.setObjectName("secondary_action_btn")
        self.gen_oswbb_btn.clicked.connect(lambda: self._on_generate_oswbb_clicked(push=False))
        
        self.gen_push_oswbb_btn = QPushButton("🚀 GENERATE & PUSH")
        self.gen_push_oswbb_btn.setObjectName("main_action_btn")
        self.gen_push_oswbb_btn.clicked.connect(lambda: self._on_generate_oswbb_clicked(push=True))
        
        self.stop_oswbb_btn = QPushButton("🛑 STOP")
        self.stop_oswbb_btn.setObjectName("clear_btn")
        self.stop_oswbb_btn.setEnabled(False)
        self.stop_oswbb_btn.setFixedWidth(100)
        self.stop_oswbb_btn.clicked.connect(self._on_stop_oswbb_clicked)
        
        action_layout.addWidget(self.gen_oswbb_btn)
        action_layout.addWidget(self.gen_push_oswbb_btn)
        action_layout.addWidget(self.stop_oswbb_btn)
        layout.addLayout(action_layout)
        
        # --- FIXED STATUS AREA ---
        oswbb_status_layout = QHBoxLayout()
        self.oswbb_progress = QProgressBar()
        self.oswbb_progress.setObjectName("globalProgressBar")
        self.oswbb_progress.setFixedHeight(20)
        self.oswbb_progress.setTextVisible(True)
        self.oswbb_progress.setFormat(" %p% ")
        self.oswbb_progress.setAlignment(Qt.AlignCenter)
        self.oswbb_progress.setValue(0)
        
        self.oswbb_status_lbl = QLabel("READY")
        self.oswbb_status_lbl.setObjectName("status_ready")
        self.oswbb_status_lbl.setAlignment(Qt.AlignCenter)
        self.oswbb_status_lbl.setFixedHeight(20)
        
        oswbb_status_layout.setSpacing(5)
        oswbb_status_layout.addWidget(self.oswbb_progress, stretch=1)
        oswbb_status_layout.addWidget(self.oswbb_status_lbl)
        layout.addLayout(oswbb_status_layout)
        
        # 4. CONSOLE LOG
        log_group = QGroupBox("JAVA CONSOLE LIVE")
        log_layout = QVBoxLayout()
        log_layout.setContentsMargins(12, 8, 12, 10)
        self.oswbb_log_text = QTextEdit()
        self.oswbb_log_text.setReadOnly(True)
        self.oswbb_log_text.setMinimumHeight(100) 
        log_layout.addWidget(self.oswbb_log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group, 1)
        
        widget.setLayout(layout)
        return widget
    
    def _create_exawatcher_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(15, 10, 15, 10)
        
        # 1. SETUP GENERATOR
        config_group = QGroupBox("1. SETUP EXAWATCHER SOURCES")
        config_layout = QGridLayout()
        config_layout.setContentsMargins(12, 12, 12, 10) 
        config_layout.setSpacing(10)
        
        # CPU/Mem Source (DB/VM)
        config_layout.addWidget(QLabel("DB/VM Log (CPU/Mem):"), 0, 0)
        self.exa_db_input_dir = QLineEdit()
        self.exa_db_input_dir.setPlaceholderText("Folder containing _mp.html and _meminfo.html...")
        self.exa_db_input_dir.setAcceptDrops(True)
        self.exa_db_input_dir.installEventFilter(self)
        config_layout.addWidget(self.exa_db_input_dir, 0, 1)
        
        btn_box_db = QHBoxLayout()
        btn_browse_db = QPushButton("📂")
        btn_browse_db.setObjectName("browse_btn")
        btn_browse_db.clicked.connect(self._browse_exa_db)
        
        btn_clear_db = QPushButton("✖")
        btn_clear_db.setObjectName("clear_btn")
        btn_clear_db.clicked.connect(self.exa_db_input_dir.clear)
        
        btn_box_db.addWidget(btn_browse_db)
        btn_box_db.addWidget(btn_clear_db)
        config_layout.addLayout(btn_box_db, 0, 2)
        
        # IO Source (Cell)
        config_layout.addWidget(QLabel("Cell Log (IO):"), 1, 0)
        self.exa_cell_input_dir = QLineEdit()
        self.exa_cell_input_dir.setPlaceholderText("Folder containing _iosummary.html...")
        self.exa_cell_input_dir.setAcceptDrops(True)
        self.exa_cell_input_dir.installEventFilter(self)
        config_layout.addWidget(self.exa_cell_input_dir, 1, 1)
        
        btn_box_cell = QHBoxLayout()
        btn_browse_cell = QPushButton("📂")
        btn_browse_cell.setObjectName("browse_btn")
        btn_browse_cell.clicked.connect(self._browse_exa_cell)
        
        btn_clear_cell = QPushButton("✖")
        btn_clear_cell.setObjectName("clear_btn")
        btn_clear_cell.clicked.connect(self.exa_cell_input_dir.clear)
        
        btn_box_cell.addWidget(btn_browse_cell)
        btn_box_cell.addWidget(btn_clear_cell)
        config_layout.addLayout(btn_box_cell, 1, 2)

        # Output
        config_layout.addWidget(QLabel("Output Images Folder:"), 2, 0)
        self.exa_output_dir = QLineEdit()
        self.exa_output_dir.setPlaceholderText("Empty = Default 'exawatcher_files' in application folder")
        self.exa_output_dir.setAcceptDrops(True)
        self.exa_output_dir.installEventFilter(self)
        config_layout.addWidget(self.exa_output_dir, 2, 1)
        
        btn_box_out = QHBoxLayout()
        btn_browse_out = QPushButton("📂")
        btn_browse_out.setObjectName("browse_btn")
        btn_browse_out.clicked.connect(self._browse_exa_out)
        
        btn_clear_out = QPushButton("✖")
        btn_clear_out.setObjectName("clear_btn")
        btn_clear_out.clicked.connect(self.exa_output_dir.clear)
        
        btn_box_out.addWidget(btn_browse_out)
        btn_box_out.addWidget(btn_clear_out)
        config_layout.addLayout(btn_box_out, 2, 2)
        
        config_layout.setColumnStretch(1, 1)
        config_group.setLayout(config_layout)
        layout.addWidget(config_group)
        
        push_group = QGroupBox("2. TARGET FOLDERS FOR SYNC")
        push_layout = QVBoxLayout()
        push_layout.setSpacing(10)
        
        # Action row for targets
        push_btn_row = QHBoxLayout()
        btn_add_target = QPushButton("➕ Add Target Folders")
        btn_add_target.setObjectName("browse_btn")
        btn_add_target.clicked.connect(self._on_exa_add_push_folder)
        
        btn_clear_target = QPushButton("✖ Clear All")
        btn_clear_target.setObjectName("clear_btn")
        btn_clear_target.clicked.connect(self._on_exa_clear_push_folders)
        
        self.exa_push_mode_overwrite = QRadioButton("Overwrite")
        self.exa_push_mode_timestamp = QRadioButton("Timestamp Subfolder")
        self.exa_push_mode_overwrite.setMinimumWidth(100)
        self.exa_push_mode_timestamp.setMinimumWidth(170)
        self.exa_push_mode_overwrite.setChecked(True)
        
        push_btn_row.addWidget(btn_add_target)
        push_btn_row.addWidget(btn_clear_target)
        push_btn_row.addStretch()
        push_btn_row.addWidget(QLabel("Push Mode:"))
        push_btn_row.addWidget(self.exa_push_mode_overwrite)
        push_btn_row.addWidget(self.exa_push_mode_timestamp)
        push_layout.addLayout(push_btn_row)
        
        # Target list
        self.exa_push_target_list = QListWidget()
        self.exa_push_target_list.setObjectName("merge_list") 
        self.exa_push_target_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.exa_push_target_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.exa_push_target_list.setMinimumHeight(120) 
        self.exa_push_target_list.setAcceptDrops(True)
        self.exa_push_target_list.installEventFilter(self) 
        push_layout.addWidget(self.exa_push_target_list)
        
        push_layout.setContentsMargins(12, 10, 12, 12) 
        push_layout.setSpacing(10)
        
        push_group.setLayout(push_layout)
        layout.addWidget(push_group, 2)
        
        # 3. ACTIONS
        action_layout = QHBoxLayout()
        action_layout.setSpacing(15)
        self.gen_exa_btn = QPushButton("🖻 GENERATE IMAGES")
        self.gen_exa_btn.setObjectName("secondary_action_btn")
        self.gen_exa_btn.clicked.connect(lambda: self._on_generate_exawatcher_clicked(push=False))
        
        self.gen_push_exa_btn = QPushButton("🚀 GENERATE & PUSH")
        self.gen_push_exa_btn.setObjectName("main_action_btn")
        self.gen_push_exa_btn.clicked.connect(lambda: self._on_generate_exawatcher_clicked(push=True))
        
        self.stop_exa_btn = QPushButton("🛑 STOP")
        self.stop_exa_btn.setObjectName("clear_btn")
        self.stop_exa_btn.setEnabled(False)
        self.stop_exa_btn.setFixedWidth(100)
        self.stop_exa_btn.clicked.connect(self._on_stop_exawatcher_clicked)
        
        action_layout.addWidget(self.gen_exa_btn)
        action_layout.addWidget(self.gen_push_exa_btn)
        action_layout.addWidget(self.stop_exa_btn)
        layout.addLayout(action_layout)
        
        # --- FIXED STATUS AREA ---
        exa_status_layout = QHBoxLayout()
        self.exa_progress = QProgressBar()
        self.exa_progress.setObjectName("globalProgressBar")
        self.exa_progress.setFixedHeight(20)
        self.exa_progress.setTextVisible(True)
        self.exa_progress.setFormat(" %p% ")
        self.exa_progress.setAlignment(Qt.AlignCenter)
        self.exa_progress.setValue(0)
        
        self.exa_status_lbl = QLabel("READY")
        self.exa_status_lbl.setObjectName("status_ready")
        self.exa_status_lbl.setAlignment(Qt.AlignCenter)
        self.exa_status_lbl.setFixedHeight(20)
        
        exa_status_layout.setSpacing(5)
        exa_status_layout.addWidget(self.exa_progress, stretch=1)
        exa_status_layout.addWidget(self.exa_status_lbl)
        layout.addLayout(exa_status_layout)
        
        # 4. CONSOLE LOG
        log_group = QGroupBox("EXAWATCHER PROCESSING LOG")
        log_layout = QVBoxLayout()
        log_layout.setContentsMargins(12, 8, 12, 10)
        self.exa_log_text = QTextEdit()
        self.exa_log_text.setReadOnly(True)
        self.exa_log_text.setMinimumHeight(100) 
        log_layout.addWidget(self.exa_log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group, 1)
        
        widget.setLayout(layout)
        return widget
    

    def _add_node_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Node Folder")
        if folder:
            folder = Path(folder).as_posix()
            if folder in self.log_folders:
                QMessageBox.information(self, "Duplicate", "This folder is already added.")
                return
            if len(self.log_folders) >= 8:
                QMessageBox.warning(self, "Limit Reached", "Maximum 8 nodes supported currently to prevent memory issues.")
                return
            
            self.log_folders.append(folder)
            display_text = f"Node {len(self.log_folders)}: {folder}"
            self.node_list_widget.addItem(display_text)
            self._log(f"Added {display_text}")

    def _clear_nodes(self):
        self.log_folders.clear()
        self.node_list_widget.clear()
        self._log("Cleared all node folders.")

    def _apply_modern_shadows(self):
        """Apply very subtle elevation for Windows 11 card style."""
        for gbox in self.findChildren(QGroupBox):
            shadow = QGraphicsDropShadowEffect()
            shadow.setBlurRadius(8) # Windows 11 uses very tight shadow if any
            shadow.setXOffset(0)
            shadow.setYOffset(2)
            shadow.setColor(QColor(0, 0, 0, 15)) # Very light black (15 alpha out of 255)
            gbox.setGraphicsEffect(shadow)

    def _on_oswbb_add_push_folder(self):
        # Mở popup Select Native của Windows thay vì Popup dạng Find Directory cũ
        folder = QFileDialog.getExistingDirectory(self, "Select Target Folder", "", QFileDialog.ShowDirsOnly)
        if folder:
            path = Path(folder).as_posix()
            if path not in self.oswbb_push_folders:
                self.oswbb_push_folders.append(path)
                self._update_push_target_list()

    def _on_oswbb_clear_push_folders(self):
        self.oswbb_push_folders.clear()
        self._update_push_target_list()

    def _update_push_target_list(self):
        self.push_target_list.clear()
        for folder in self.oswbb_push_folders:
            item = QListWidgetItem(folder)
            item.setToolTip(folder)
            self.push_target_list.addItem(item)

    def _on_exa_add_push_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Target Folder", "", QFileDialog.ShowDirsOnly)
        if folder:
            path = Path(folder).as_posix()
            if path not in self.exa_push_folders:
                self.exa_push_folders.append(path)
                self._update_exa_push_target_list()

    def _on_exa_clear_push_folders(self):
        self.exa_push_folders.clear()
        self._update_exa_push_target_list()

    def _update_exa_push_target_list(self):
        self.exa_push_target_list.clear()
        for folder in self.exa_push_folders:
            item = QListWidgetItem(folder)
            # Add tooltip for long paths
            item.setToolTip(folder)
            self.exa_push_target_list.addItem(item)

    # ── Drag & Drop Handling ─────────────────────────────────────────
    def eventFilter(self, watched, event):
        # Safely identify if the watched widget is one of our drag & drop targets
        targets = [
            getattr(self, 'node_list_widget', None),
            getattr(self, 'push_target_list', None),
            getattr(self, 'exa_push_target_list', None), # Added new target
            getattr(self, 'merge_file_list', None),
            getattr(self, 'oswbb_input_dir', None),
            getattr(self, 'oswbb_output_dir', None),
            getattr(self, 'exa_db_input_dir', None),
            getattr(self, 'exa_cell_input_dir', None),
            getattr(self, 'exa_output_dir', None),
            getattr(self, 'merge_output_path', None)
        ]
        
        targets = [t for t in targets if t is not None]
        if watched in targets and event.type() == QEvent.DragEnter:
            if event.mimeData().hasUrls():
                event.accept()
                return True
        if watched in targets and event.type() == QEvent.Drop:
            urls = event.mimeData().urls()
            if not urls: return False
            path = urls[0].toLocalFile()
            
            if watched == self.node_list_widget:
                for u in urls:
                    p = Path(u.toLocalFile()).as_posix()
                    if Path(p).is_dir() and p not in self.log_folders:
                        self.log_folders.append(p)
                        self.node_list_widget.addItem(f"Node {len(self.log_folders)}: {p}")
                return True
            elif watched == self.push_target_list:
                for u in urls:
                    p = Path(u.toLocalFile()).as_posix()
                    if Path(p).is_dir() and p not in self.oswbb_push_folders:
                        self.oswbb_push_folders.append(p)
                        self._update_push_target_list()
                return True
            elif watched == self.exa_push_target_list: # Added logic for ExaWatcher
                for u in urls:
                    p = Path(u.toLocalFile()).as_posix()
                    if Path(p).is_dir() and p not in self.exa_push_folders:
                        self.exa_push_folders.append(p)
                        self._update_exa_push_target_list()
                return True
            elif watched == self.merge_file_list:
                for u in urls:
                    p = Path(u.toLocalFile()).as_posix()
                    if p.lower().endswith(".docx"): self._merge_add_direct(p)
                return True
            elif isinstance(watched, QLineEdit):
                watched.setText(Path(path).as_posix())
                return True
            return True
        if watched in targets and event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_Delete:
                if watched == self.merge_file_list:
                    self._merge_remove_file()
                elif watched == self.node_list_widget:
                    row = self.node_list_widget.currentRow()
                    if row >= 0:
                        self.node_list_widget.takeItem(row)
                        self.log_folders.pop(row)
                        self._log(f"Removed item at row {row+1}")
                return True

        return super().eventFilter(watched, event)

    def _merge_add_direct(self, path):
        existing = [self.merge_file_list.item(i).data(Qt.UserRole)
                    for i in range(self.merge_file_list.count())]
        if path not in existing:
            item = QListWidgetItem(f"{self.merge_file_list.count() + 1}. {Path(path).name}")
            item.setData(Qt.UserRole, path)
            self.merge_file_list.addItem(item)
            self._merge_refresh_numbers()

    def _browse_oswbb_in(self):
        folder = QFileDialog.getExistingDirectory(self, "Select OSWBB Input Folder")
        if folder:
            self.oswbb_input_dir.setText(Path(folder).as_posix())
            
    def _browse_oswbb_out(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.oswbb_output_dir.setText(Path(folder).as_posix())

    def _on_generate_oswbb_clicked(self, push=False):
        in_dir = self.oswbb_input_dir.text().strip()
        out_dir = self.oswbb_output_dir.text().strip()
        
        if not in_dir:
            QMessageBox.warning(self, "Missing fields", "Please select input OSWBB directory.")
            return

        # Default output path if empty
        if not out_dir:
            from ..config import BASE_DIR
            out_dir = (BASE_DIR / "generated_files").as_posix()
            self.oswbb_output_dir.setText(out_dir)
            self._log(f"[INFO] Automatically selected OSWBB output folder: {out_dir}")
            
        if push and not self.oswbb_push_folders:
            QMessageBox.warning(self, "No Targets", "Please add at least one target folder to push results.")
            return
            
        from ..utils.oswbb_runner import OSWBBGraphGenerator
        self.gen_oswbb_btn.setEnabled(False)
        self.gen_push_oswbb_btn.setEnabled(False)
        self.stop_oswbb_btn.setEnabled(True)
        self.oswbb_progress.setRange(0, 100)
        self.oswbb_progress.setValue(0)
        self.oswbb_progress.setVisible(True)
        self.oswbb_status_lbl.setText("RUNNING")
        self.oswbb_status_lbl.setObjectName("status_ready") # Keep gray box style for running
        self.oswbb_status_lbl.setStyle(self.oswbb_status_lbl.style())
        
        push_targets = self.oswbb_push_folders if push else []
        push_mode = "timestamp" if self.push_mode_timestamp.isChecked() else "overwrite"
        selected_jar = self.oswbb_jar_select.currentText()
        
        # Option D is now the default integrated behavior
        # Chạy trên Worker Thread để tránh lag GUI
        self.oswbb_thread = QThread()
        self.oswbb_worker = OSWBBGraphGenerator(
            in_dir, out_dir, 
            gen_dashboard=True, 
            push_targets=push_targets, 
            push_mode=push_mode,
            jar_filename=selected_jar
        )
        self.oswbb_worker.moveToThread(self.oswbb_thread)
        
        self.oswbb_thread.started.connect(self.oswbb_worker.run)
        self.oswbb_worker.progress.connect(self._on_oswbb_log)
        self.oswbb_worker.progress_val.connect(self.oswbb_progress.setValue)
        self.oswbb_worker.finished.connect(self._on_oswbb_finished)
        self.oswbb_worker.finished.connect(self.oswbb_thread.quit)
        
        # Đảm bảo tự dọn dẹp sau khi xong
        self.oswbb_worker.finished.connect(self.oswbb_worker.deleteLater)
        self.oswbb_thread.finished.connect(self.oswbb_thread.quit) # Corrected finish logic
        
        self.oswbb_thread.start()
        
    def _on_oswbb_log(self, text: str):
        self.oswbb_log_text.append(text)
        
    def _on_oswbb_finished(self, success: bool):
        self.gen_oswbb_btn.setEnabled(True)
        self.gen_push_oswbb_btn.setEnabled(True)
        self.stop_oswbb_btn.setEnabled(False)
        self.oswbb_progress.setRange(0, 100)
        self.oswbb_progress.setValue(100 if success else 0)
        
        if success:
            self.oswbb_status_lbl.setText("FINISHED")
            self.oswbb_status_lbl.setObjectName("status_finished")
            self.oswbb_status_lbl.setStyle(self.oswbb_status_lbl.style()) # Refresh style
        else:
            self.oswbb_status_lbl.setText("FAILED")
            self.oswbb_status_lbl.setObjectName("status_failed")
            self.oswbb_status_lbl.setStyle(self.oswbb_status_lbl.style())

    def _on_stop_oswbb_clicked(self):
        if hasattr(self, 'oswbb_worker') and self.oswbb_worker:
            self.oswbb_worker.stop()
            self.stop_oswbb_btn.setEnabled(False)

    def _on_parse_clicked(self):
        num_days = self.num_days_spin.value()
        self.parse_worker = ParseWorker(self.log_folders, num_days)
        self.parse_worker.progress.connect(self._on_parse_progress)
        self.parse_worker.finished.connect(self._on_parse_finished)
        self.parse_worker.error.connect(self._on_parse_error)
        self.parse_worker.start()
    
    def _on_parse_progress(self, message: str, value: int):
        self._log(message)
        self.appendix_progress.setValue(value)
    
    def _on_parse_finished(self, data: dict):
        self.parsed_data = data
        self._log("[SUCCESS] All data parsed successfully!")
        if self.gen_mode == 'report':
            self._run_report_generation_and_finalize()
        else:
            self._run_generation_and_finalize()
    
    def _on_parse_error(self, error_msg: str):
        self._log(f"[ERROR] {error_msg}")
        self.appendix_progress.setVisible(True) # Keep visible to show error on status
        self.appendix_progress.setValue(0)
        self.appendix_status_lbl.setText("FAILED")
        self.appendix_status_lbl.setObjectName("status_failed")
        self.appendix_status_lbl.setStyle(self.appendix_status_lbl.style())
        self.generate_btn.setEnabled(True)
    

        self.appendix_progress.setRange(0, 100)
        self.appendix_progress.setValue(0)
        self._log(f"Starting parsing for {len(self.log_folders)} nodes...")
        self._on_parse_clicked()

    def _on_generate_report_clicked(self):
        if not self.log_folders:
            self.appendix_status_lbl.setText("MISSING INPUT")
            self.appendix_status_lbl.setObjectName("status_failed")
            self.appendix_status_lbl.setStyle(self.appendix_status_lbl.style())
            return
            
        self.gen_mode = 'report'
        self.appendix_status_lbl.setText("RUNNING")
        self.appendix_status_lbl.setObjectName("status_ready")
        self.appendix_status_lbl.setStyle(self.appendix_status_lbl.style())
        self.generate_btn.setEnabled(False)
        self.generate_report_btn.setEnabled(False)
        self.appendix_progress.setVisible(True)
        self.appendix_progress.setRange(0, 100)
        self.appendix_progress.setValue(0)
        self._log("Starting Final Report workflow (Node 1)...")
        self._on_parse_clicked()

    def _on_generate_clicked(self):
        if not self.log_folders:
            self.appendix_status_lbl.setText("MISSING INPUT")
            self.appendix_status_lbl.setObjectName("status_failed")
            self.appendix_status_lbl.setStyle(self.appendix_status_lbl.style())
            return

        self.gen_mode = 'appendix'
        self.appendix_status_lbl.setText("RUNNING")
        self.appendix_status_lbl.setObjectName("status_ready")
        self.appendix_status_lbl.setStyle(self.appendix_status_lbl.style())
        self.generate_btn.setEnabled(False)
        self.generate_report_btn.setEnabled(False)
        self.appendix_progress.setVisible(True)
        self.appendix_progress.setRange(0, 100)
        self.appendix_progress.setValue(0)
        self._log(f"Starting workflow for {len(self.log_folders)} nodes...")
        self._on_parse_clicked()

    def _get_calculated_db_name(self) -> str:
        """Helper to calculate the consistent database name across all report types"""
        if not self.parsed_data: return "UNKNOWN"
        
        nodes = self.parsed_data.get('nodes', [])
        db_role = self.role_combo.currentText().lower()
        
        base_db_name = "Unknown"
        if nodes:
            inst = str(nodes[0].get('instance_name', '')).strip()
            if not inst or inst == 'NODE1':
                # Fallback to DB Name from AWR/Config
                base_db_name = self.parsed_data.get('db_name', 'Unknown')
            else:
                # Use instance name but strip trailing node digits (e.g. MISDB1 -> MISDB)
                base_db_name = inst
                while base_db_name and base_db_name[-1].isdigit():
                    base_db_name = base_db_name[:-1]
        
        base_db_name = base_db_name.upper()
        if db_role == 'standby':
            base_db_name = f"{base_db_name}-STB"
            
        return base_db_name

    def _run_generation_and_finalize(self):
        self._log("Initializing dynamic multi-node report generation...")
        try:
            font_choice = self.font_combo.currentText()
            db_role = self.role_combo.currentText().lower()
            
            base_db_name = self._get_calculated_db_name()
            
            default_filename = f"{base_db_name}_appendix"
            filename = self.filename_input.text() or default_filename
            filename = sanitize_filename(filename)
            
            docx_path = str(APPENDIX_OUTPUT_DIR / f"{filename}.docx")
            font_token = 'times' if 'times' in font_choice.lower() else 'calibri'
            
            # Create and start Generator Worker in background
            self.gen_thread = QThread()
            self.gen_worker = GeneratorWorker(self.parsed_data, docx_path, font_token, filename, db_role)
            self.gen_worker.moveToThread(self.gen_thread)
            
            self.gen_thread.started.connect(self.gen_worker.run)
            self.gen_worker.progress.connect(self._on_parse_progress)
            self.gen_worker.finished.connect(self._on_generation_finished)
            self.gen_worker.finished.connect(self.gen_thread.quit)
            self.gen_worker.finished.connect(self.gen_worker.deleteLater)
            self.gen_thread.finished.connect(self.gen_thread.deleteLater)
            
            self.gen_thread.start()
            
        except Exception as e:
            self._log(f"[ERROR] Fault initiating generator: {str(e)}")
            logger.exception("Generator initiation error")
            self.appendix_progress.setVisible(False)
            self.generate_btn.setEnabled(True)

    def _on_generation_finished(self, success: bool, docx_path: str, filename: str):
        """Callback when background generation completes"""
        self.generate_btn.setEnabled(True)
        self.generate_report_btn.setEnabled(True)
        self.appendix_progress.setValue(100 if success else 0)
        self.statusBar().showMessage("Workflow Finished")
        
        if success:
            self._log(f"[SUCCESS] Report saved: {docx_path}")
            self.appendix_status_lbl.setText("FINISHED")
            self.appendix_status_lbl.setObjectName("status_finished")
        else:
            # If docx_path contains error message (from GeneratorWorker)
            error_msg = docx_path if docx_path else "Unknown error"
            self._log(f"[ERROR] Report generation failed: {error_msg}")
            self.appendix_status_lbl.setText("FAILED")
            self.appendix_status_lbl.setObjectName("status_failed")
        
        self.appendix_status_lbl.setStyle(self.appendix_status_lbl.style())

    def _run_report_generation_and_finalize(self):
        self._log("Initializing summary report generation...")
        try:
            font_choice = self.font_combo.currentText()
            lang = 'vi' if "Vietnamese" in self.lang_combo.currentText() else 'en'
            
            base_db_name = self._get_calculated_db_name()
            
            default_filename = f"{base_db_name}_report"
            filename = self.filename_input.text() or default_filename
            filename = sanitize_filename(filename)
            
            docx_path = str(REPORT_OUTPUT_DIR / f"{filename}.docx")
            font_token = 'times' if 'times' in font_choice.lower() else 'calibri'
            
            self.report_thread = QThread()
            self.report_worker = ReportWorker(self.parsed_data, docx_path, font_token, filename, lang, db_name=base_db_name)
            self.report_worker.moveToThread(self.report_thread)
            
            self.report_thread.started.connect(self.report_worker.run)
            self.report_worker.progress.connect(self._on_parse_progress)
            self.report_worker.finished.connect(self._on_report_finished)
            self.report_worker.finished.connect(self.report_thread.quit)
            self.report_worker.finished.connect(self.report_worker.deleteLater)
            self.report_thread.finished.connect(self.report_thread.deleteLater)
            
            self.report_thread.start()
        except Exception as e:
            self._log(f"[ERROR] Fault initiating report: {str(e)}")
            self.generate_btn.setEnabled(True)
            self.generate_report_btn.setEnabled(True)

    def _on_report_finished(self, success: bool, docx_path: str, filename: str):
        self.generate_btn.setEnabled(True)
        self.generate_report_btn.setEnabled(True)
        self.appendix_progress.setValue(100 if success else 0)
        self.statusBar().showMessage("Final Report Finished")
        if success:
            self._log(f"[SUCCESS] Final Report saved: {docx_path}")
            self.appendix_status_lbl.setText("FINISHED")
            self.appendix_status_lbl.setObjectName("status_finished")
        else:
            self._log(f"[ERROR] Final Report failed: {docx_path}")
            self.appendix_status_lbl.setText("FAILED")
            self.appendix_status_lbl.setObjectName("status_failed")
        self.appendix_status_lbl.setStyle(self.appendix_status_lbl.style())
    
    def _log(self, message: str):
        self.log_text.append(message)
        logger.info(message)

    # ─────────────────────────────────────────────────────────────────
    #  TAB: Merge Documents
    # ─────────────────────────────────────────────────────────────────
    def _create_merge_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(15, 10, 15, 10)

        # ── 1. File Queue ──────────────────────────────────────────
        queue_group = QGroupBox("1. DOCUMENTS TO MERGE")
        queue_layout = QVBoxLayout()
        queue_layout.setSpacing(8) # Tighten internal list area

        # Action buttons row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        btn_add = QPushButton("➕ Add Files")
        btn_add.setObjectName("browse_btn")
        btn_add.clicked.connect(self._merge_add_files)

        btn_clear_all = QPushButton("✖ Clear All")
        btn_clear_all.setObjectName("clear_btn")
        btn_clear_all.clicked.connect(self._merge_clear_all)

        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_clear_all)
        btn_row.addStretch()
        queue_layout.addLayout(btn_row)

        # Draggable list widget - no alternating colors to prevent ghost rows
        self.merge_file_list = QListWidget()
        self.merge_file_list.setObjectName("merge_list")
        self.merge_file_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.merge_file_list.setDragDropOverwriteMode(False)
        self.merge_file_list.setDefaultDropAction(Qt.MoveAction)
        self.merge_file_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.merge_file_list.setAcceptDrops(True)
        self.merge_file_list.installEventFilter(self)
        self.merge_file_list.setMinimumHeight(200)
        queue_layout.addWidget(self.merge_file_list)

        queue_group.setLayout(queue_layout)

        # ── 2. DB ORDER LIST ──────────────────────────────────────────
        sort_group = QGroupBox("2. DB ORDER LIST")
        sort_layout = QVBoxLayout()
        sort_layout.setSpacing(8)

        self.db_order_input = QTextEdit()
        self.db_order_input.setObjectName("merge_list")
        self.db_order_input.setPlaceholderText(
            "Dán danh sách DB theo\nthứ tự mong muốn:\n\nDCFNGTB\nPRODDB\nFINDB\n\n"
            "Quy tắc: tên file\nDCFNGTB_appendix.docx\n→ db_name = DCFNGTB"
        )
        sort_layout.addWidget(self.db_order_input, stretch=1)

        btn_sort = QPushButton("⚡ Auto Sort by List")
        btn_sort.setObjectName("main_action_btn")
        btn_sort.setToolTip("Sắp xếp lại danh sách file theo thứ tự DB ở trên")
        btn_sort.clicked.connect(self._merge_sort_by_db_list)
        sort_layout.addWidget(btn_sort)

        sort_group.setLayout(sort_layout)

        # ── Đặt 2 GroupBox nằm ngang nhau ────────────────────────────
        top_row = QHBoxLayout()
        top_row.setSpacing(10)
        top_row.addWidget(queue_group, stretch=3)   # 60% chiều rộng
        top_row.addWidget(sort_group, stretch=2)    # 40% chiều rộng
        layout.addLayout(top_row, stretch=1)        # Chiếm toàn bộ không gian dọc còn lại

        # ── 3. Output Settings ─────────────────────────────────────
        output_group = QGroupBox("3. OUTPUT SETTINGS")
        output_grid = QGridLayout()
        output_grid.setSpacing(10)

        output_grid.addWidget(QLabel("Save Merged File As:"), 0, 0)
        self.merge_output_path = QLineEdit()
        self.merge_output_path.setPlaceholderText("Empty = Default 'output/merged_appendix.docx' in application folder")
        self.merge_output_path.setAcceptDrops(True)
        self.merge_output_path.installEventFilter(self)
        output_grid.addWidget(self.merge_output_path, 0, 1)

        btn_box_m = QHBoxLayout()
        btn_browse_merge = QPushButton("📂")
        btn_browse_merge.setObjectName("browse_btn")
        btn_browse_merge.clicked.connect(self._browse_merge_output)
        
        btn_clear_merge = QPushButton("✖")
        btn_clear_merge.setObjectName("clear_btn")
        btn_clear_merge.clicked.connect(self.merge_output_path.clear)
        
        btn_box_m.addWidget(btn_browse_merge)
        btn_box_m.addWidget(btn_clear_merge)
        output_grid.addLayout(btn_box_m, 0, 2)

        output_grid.setColumnStretch(1, 1)
        output_group.setLayout(output_grid)
        layout.addWidget(output_group)

        # ── 3. Execute ─────────────────────────────────────────────
        self.merge_btn = QPushButton("🔗 MERGE DOCUMENTS")
        self.merge_btn.setObjectName("main_action_btn")
        self.merge_btn.clicked.connect(self._on_merge_clicked)
        layout.addWidget(self.merge_btn)
        
        # --- FIXED STATUS AREA ---
        merge_status_layout = QHBoxLayout()
        self.merge_progress_bar = QProgressBar()
        self.merge_progress_bar.setObjectName("globalProgressBar")
        self.merge_progress_bar.setFixedHeight(20)
        self.merge_progress_bar.setTextVisible(True)
        self.merge_progress_bar.setFormat(" %p% ")
        self.merge_progress_bar.setAlignment(Qt.AlignCenter)
        self.merge_progress_bar.setValue(0)
        
        self.merge_status_lbl = QLabel("READY")
        self.merge_status_lbl.setObjectName("status_ready")
        self.merge_status_lbl.setAlignment(Qt.AlignCenter)
        self.merge_status_lbl.setFixedHeight(20)
        
        merge_status_layout.setSpacing(5)
        merge_status_layout.addWidget(self.merge_progress_bar, stretch=1)
        merge_status_layout.addWidget(self.merge_status_lbl)
        layout.addLayout(merge_status_layout)

        # ── 4. Live log ────────────────────────────────────────────
        log_group = QGroupBox("MERGE CONSOLE")
        log_layout = QVBoxLayout()
        self.merge_log_text = QTextEdit()
        self.merge_log_text.setReadOnly(True)
        self.merge_log_text.setFixedHeight(100)
        log_layout.addWidget(self.merge_log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        widget.setLayout(layout)
        return widget

    def _create_tools_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(20)
        
        header_lbl = QLabel("Danh sách công cụ hỗ trợ thu thập dữ liệu (HC Tools)")
        header_lbl.setStyleSheet("font-size: 16px; font-weight: 600; color: #555;")
        layout.addWidget(header_lbl)
        
        self.tools_list = QListWidget()
        self.tools_list.setObjectName("merge_list")
        self.tools_list.setSelectionMode(QAbstractItemView.NoSelection)
        layout.addWidget(self.tools_list)
        
        # Populate Tools
        self._refresh_tools_list()
        
        # Action Buttons Row
        btn_layout = QHBoxLayout()
        
        btn_refresh = QPushButton("🔄 Refresh Local")
        btn_refresh.setObjectName("secondary_action_btn")
        btn_refresh.clicked.connect(self._on_refresh_tools_clicked)
        
        btn_sync_github = QPushButton("🌐 Sync from GitHub")
        btn_sync_github.setObjectName("main_action_btn")
        btn_sync_github.clicked.connect(self._on_sync_github_clicked)
        
        btn_open_folder = QPushButton("📂 Open Folder")
        btn_open_folder.setObjectName("secondary_action_btn")
        btn_open_folder.clicked.connect(self._on_open_tools_folder)
        
        btn_layout.addWidget(btn_refresh)
        btn_layout.addWidget(btn_sync_github)
        btn_layout.addWidget(btn_open_folder)
        layout.addLayout(btn_layout)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget

    def _refresh_tools_list(self):
        self.tools_list.clear()
        if COLLECT_TOOL_DIR.exists():
            # Sorted file listing for better UX
            files = sorted([f for f in COLLECT_TOOL_DIR.iterdir() if f.is_file()])
            for f in files:
                item = QListWidgetItem(f"📄 {f.name} ({round(f.stat().st_size/1024, 1)} KB)")
                self.tools_list.addItem(item)
        else:
            self.tools_list.addItem("⚠️ Thư mục công cụ không tồn tại hoặc chưa được giải nén.")

    def _on_open_tools_folder(self):
        import os
        if COLLECT_TOOL_DIR.exists():
            os.startfile(str(COLLECT_TOOL_DIR))
        else:
            self.statusBar().showMessage("⚠️ Lỗi: Thư mục công cụ không tồn tại.", 5000)

    def _on_refresh_tools_clicked(self):
        """Manual check and refresh for tool list"""
        self._refresh_tools_list()
        self.statusBar().showMessage("Đã cập nhật danh sách công cụ cục bộ.", 3000)

    def _on_sync_github_clicked(self):
        """Khởi chạy đồng bộ từ GitHub"""
        if self.github_worker and self.github_worker.isRunning():
            return
            
        self.has_auto_synced = True # Mark as حاول (even if fails, we don't spam)
        self.statusBar().showMessage("Đang đồng bộ công cụ từ GitHub...")
        
        self.github_worker = GitHubSyncWorker()
        self.github_worker.progress.connect(lambda p, m: self.statusBar().showMessage(f"{m} ({p}%)"))
        self.github_worker.finished.connect(self._on_github_sync_finished)
        self.github_worker.error.connect(lambda e: self.statusBar().showMessage(f"❌ Lỗi: {e}", 5000))
        self.github_worker.start()

    def _on_github_sync_finished(self, success: bool, message: str):
        if success:
            self._refresh_tools_list()
            self.statusBar().showMessage("Đồng bộ GitHub hoàn tất!", 5000)
        else:
            self.statusBar().showMessage(f"Đầu bộ thất bại: {message}", 5000)


    # ── Merge tab slots ────────────────────────────────────────────
    def _merge_add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select .docx Files", "", "Word Documents (*.docx)"
        )
        for path in paths:
            # Avoid duplicates
            existing = [self.merge_file_list.item(i).data(Qt.UserRole)
                        for i in range(self.merge_file_list.count())]
            if path not in existing:
                from pathlib import Path as _P
                from PyQt5.QtWidgets import QListWidgetItem
                item = QListWidgetItem(
                    f"{self.merge_file_list.count() + 1}. {_P(path).name}"
                )
                item.setData(Qt.UserRole, path)
                self.merge_file_list.addItem(item)
        self._merge_refresh_numbers()

    def _merge_remove_file(self):
        row = self.merge_file_list.currentRow()
        if row >= 0:
            self.merge_file_list.takeItem(row)
            self._merge_refresh_numbers()

    def _merge_move_up(self):
        row = self.merge_file_list.currentRow()
        if row > 0:
            item = self.merge_file_list.takeItem(row)
            self.merge_file_list.insertItem(row - 1, item)
            self.merge_file_list.setCurrentRow(row - 1)
            self._merge_refresh_numbers()

    def _merge_move_down(self):
        row = self.merge_file_list.currentRow()
        if row < self.merge_file_list.count() - 1:
            item = self.merge_file_list.takeItem(row)
            self.merge_file_list.insertItem(row + 1, item)
            self.merge_file_list.setCurrentRow(row + 1)
            self._merge_refresh_numbers()

    def _merge_clear_all(self):
        self.merge_file_list.clear()

    def _merge_refresh_numbers(self):
        """Re-label all list items so numbering stays consistent."""
        from pathlib import Path as _P
        for i in range(self.merge_file_list.count()):
            item = self.merge_file_list.item(i)
            path = item.data(Qt.UserRole)
            item.setText(f"{i + 1}. {_P(path).name}")

    def _merge_sort_by_db_list(self):
        """Sắp xếp lại queue file theo thứ tự danh sách DB người dùng nhập."""
        raw_text = self.db_order_input.toPlainText().strip()
        if not raw_text:
            self.statusBar().showMessage("⚠️ Vui lòng nhập danh sách tên DB trước khi sắp xếp.", 3000)
            return

        if self.merge_file_list.count() == 0:
            self.statusBar().showMessage("⚠️ Vui lòng thêm file vào danh sách trước khi sắp xếp.", 3000)
            return

        # Đọc danh sách DB theo thứ tự (bỏ dòng trống, strip whitespace)
        db_order = [line.strip().upper() for line in raw_text.splitlines() if line.strip()]

        # Thu thập tất cả file hiện có trong queue
        all_paths = [
            self.merge_file_list.item(i).data(Qt.UserRole)
            for i in range(self.merge_file_list.count())
        ]

        # Tạo dict: db_name -> path (lấy phần trước '_' đầu tiên, uppercase)
        from pathlib import Path as _P
        db_to_path = {}
        unmatched = []
        for path in all_paths:
            filename = _P(path).stem  # Bỏ đuôi .docx
            db_key = filename.split('_')[0].upper()
            db_to_path[db_key] = path

        # Sắp xếp theo thứ tự trong danh sách
        sorted_paths = []
        not_found_dbs = []
        for db_name in db_order:
            if db_name in db_to_path:
                sorted_paths.append(db_to_path[db_name])
            else:
                not_found_dbs.append(db_name)

        # Thêm các file không khớp vào cuối
        matched_paths = set(sorted_paths)
        for path in all_paths:
            if path not in matched_paths:
                sorted_paths.append(path)
                unmatched.append(_P(path).name)

        # Cập nhật lại list widget theo thứ tự mới
        self.merge_file_list.clear()
        for path in sorted_paths:
            item = QListWidgetItem(f"temp. {_P(path).name}")
            item.setData(Qt.UserRole, path)
            self.merge_file_list.addItem(item)
        self._merge_refresh_numbers()

        # Sắp xếp hoàn tất -> cập nhật status bar
        msg = f"✅ Đã sắp xếp {len(sorted_paths)} file."
        if not_found_dbs: msg += f" (Thiếu: {len(not_found_dbs)})"
        self.statusBar().showMessage(msg, 5000)

    def _browse_merge_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Merged Document As", "merged_appendix.docx",
            "Word Documents (*.docx)"
        )
        if path:
            if not path.lower().endswith('.docx'):
                path += '.docx'
            self.merge_output_path.setText(Path(path).as_posix())

    def _on_merge_clicked(self):
        count = self.merge_file_list.count()
        if count < 2:
            self.merge_status_lbl.setText("NOT ENOUGH FILES")
            self.merge_status_lbl.setObjectName("status_failed")
            self.merge_status_lbl.setStyle(self.merge_status_lbl.style())
            return

        output = self.merge_output_path.text().strip()
        if not output:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output = (OUTPUT_DIR / "merged_appendix.docx").as_posix()
            self.merge_output_path.setText(output)
            self._log(f"[INFO] Automatically selected merge output: {output}")

        ordered_paths = [
            self.merge_file_list.item(i).data(Qt.UserRole)
            for i in range(count)
        ]

        self.merge_btn.setEnabled(False)
        self.merge_log_text.clear()
        self.merge_progress_bar.setValue(0)
        self.merge_progress_bar.setRange(0, 100)
        self.merge_progress_bar.setVisible(True)
        self.merge_log_text.append(f"[SYSTEM] Starting merge of {count} documents...")
        self.merge_status_lbl.setText("RUNNING") # Clear previous
        self.merge_status_lbl.setObjectName("status_ready")
        self.merge_status_lbl.setStyle(self.merge_status_lbl.style())

        from ..utils.merge_worker import MergeWorker
        self.merge_thread = QThread()
        self.merge_worker = MergeWorker(ordered_paths, output)
        self.merge_worker.moveToThread(self.merge_thread)

        self.merge_thread.started.connect(self.merge_worker.run)
        self.merge_worker.progress.connect(self._on_merge_progress)
        self.merge_worker.finished.connect(self._on_merge_finished)
        self.merge_worker.finished.connect(self.merge_thread.quit)
        self.merge_worker.finished.connect(self.merge_worker.deleteLater)
        self.merge_thread.finished.connect(self.merge_thread.deleteLater)

        self.merge_thread.start()

    def _on_merge_progress(self, percent: int, message: str):
        """Update progress bar and append the log line."""
        self.merge_progress_bar.setValue(percent)
        self.merge_log_text.append(f"[{percent:3d}%] {message}")

    def _on_merge_finished(self, success: bool, message: str):
        self.merge_btn.setEnabled(True)
        self.merge_progress_bar.setValue(100 if success else 0)
        self.merge_log_text.append(f"\n{'[SUCCESS]' if success else '[ERROR]'} {message}")
        
        if success:
            self.merge_status_lbl.setText("FINISHED")
            self.merge_status_lbl.setObjectName("status_finished")
        else:
            self.merge_status_lbl.setText("FAILED")
            self.merge_status_lbl.setObjectName("status_failed")
        
        self.merge_status_lbl.setStyle(self.merge_status_lbl.style())
    def _browse_exa_db(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select DB Log Source", "", "Archives (*.tar.bz2);;All Files (*)")
        if file_path:
            self.exa_db_input_dir.setText(Path(file_path).as_posix())

    def _browse_exa_cell(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Cell Log Source", "", "Archives (*.tar.bz2);;All Files (*)")
        if path:
            self.exa_cell_input_dir.setText(Path(path).as_posix())

    def _browse_exa_out(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder: self.exa_output_dir.setText(Path(folder).as_posix())

    def _on_generate_exawatcher_clicked(self, push=False):
        db_in = self.exa_db_input_dir.text().strip()
        cell_in = self.exa_cell_input_dir.text().strip()
        out_dir = self.exa_output_dir.text().strip()
        
        if not db_in or not cell_in:
            self.exa_status_lbl.setText("MISSING FIELDS")
            self.exa_status_lbl.setObjectName("status_failed")
            self.exa_status_lbl.setStyle(self.exa_status_lbl.style())
            return

        # Default output path if empty
        if not out_dir:
            from ..config import BASE_DIR
            out_dir = (BASE_DIR / "exawatcher_files").as_posix()
            self.exa_output_dir.setText(out_dir)
            self.exa_log_text.append(f"[INFO] Automatically selected ExaWatcher output folder: {out_dir}")

        if push and not self.exa_push_folders:
            self.exa_status_lbl.setText("NO TARGETS")
            self.exa_status_lbl.setObjectName("status_failed")
            self.exa_status_lbl.setStyle(self.exa_status_lbl.style())
            return

        self.gen_exa_btn.setEnabled(False)
        self.gen_push_exa_btn.setEnabled(False)
        self.stop_exa_btn.setEnabled(True)
        self.exa_progress.setRange(0, 100)
        self.exa_progress.setValue(0)
        self.exa_progress.setVisible(True)
        self.exa_log_text.clear()
        self.exa_log_text.append("[SYSTEM] Khởi chạy bộ xử lý ExaWatcher...")
        self.exa_status_lbl.setText("RUNNING") # Clear previous
        self.exa_status_lbl.setObjectName("status_ready")
        self.exa_status_lbl.setStyle(self.exa_status_lbl.style())

        push_targets = self.exa_push_folders if push else []
        push_mode = "timestamp" if self.exa_push_mode_timestamp.isChecked() else "overwrite"

        self.exa_thread = QThread()
        from ..utils.exawatcher_runner import ExaWatcherGraphGenerator
        self.exa_worker = ExaWatcherGraphGenerator(
            db_in, cell_in, out_dir,
            push_targets=push_targets,
            push_mode=push_mode
        )
        self.exa_worker.moveToThread(self.exa_thread)

        self.exa_thread.started.connect(self.exa_worker.run)
        self.exa_worker.progress.connect(self._on_exawatcher_log)
        self.exa_worker.progress_val.connect(self.exa_progress.setValue)
        self.exa_worker.finished.connect(self._on_exawatcher_finished)
        self.exa_worker.finished.connect(self.exa_thread.quit)
        
        self.exa_worker.finished.connect(self.exa_worker.deleteLater)
        self.exa_thread.finished.connect(self.exa_thread.quit)

        self.exa_thread.start()

    def _on_exawatcher_log(self, text: str):
        self.exa_log_text.append(text)

    def _on_exawatcher_finished(self, success: bool):
        self.gen_exa_btn.setEnabled(True)
        self.gen_push_exa_btn.setEnabled(True)
        self.stop_exa_btn.setEnabled(False)
        self.exa_progress.setRange(0, 100)
        self.exa_progress.setValue(100 if success else 0)
        
        if success:
            self.exa_status_lbl.setText("FINISHED")
            self.exa_status_lbl.setObjectName("status_finished")
        else:
            self.exa_status_lbl.setText("FAILED")
            self.exa_status_lbl.setObjectName("status_failed")
            
        self.exa_status_lbl.setStyle(self.exa_status_lbl.style())

    def _on_stop_exawatcher_clicked(self):
        if hasattr(self, 'exa_worker') and self.exa_worker:
            self.exa_worker.stop()
            self.stop_exa_btn.setEnabled(False)

    def _create_recommendation_tab(self) -> QWidget:
        """Create the configuration tab for recommendation rules with internal sidebar navigation"""
        widget = QWidget()
        main_hbox = QHBoxLayout()
        main_hbox.setContentsMargins(0, 0, 0, 0)
        main_hbox.setSpacing(0)
        
        # --- INTERNAL SIDEBAR (LEFT) ---
        self.rules_sidebar = QListWidget()
        self.rules_sidebar.setObjectName("sidebarMenu")
        self.rules_sidebar.setFixedWidth(240)
        
        # --- CONTENT AREA (RIGHT) ---
        content_scroll = QScrollArea()
        content_scroll.setWidgetResizable(True)
        content_scroll.setFrameShape(QFrame.NoFrame)
        # REMOVED: setStyleSheet causing black dropdowns and lost button colors
        
        content_pane = QWidget()
        content_pane_layout = QVBoxLayout(content_pane)
        # Horizontal Alignment: Left 20, right 40, bottom 20
        content_pane_layout.setContentsMargins(20, 20, 40, 20) 
        content_pane_layout.setSpacing(15)
        
        # PUSH EVERYTHING TO THE BOTTOM REMOVED to allow editor to expand
        
        # Stacks for each section
        self.rules_stack = QStackedWidget()
        # Expanding policy is safer for interaction
        self.rules_stack.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        content_pane_layout.addWidget(self.rules_stack)
        
        # Action Buttons (Attached directly to content area)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch() # Push to right
        
        btn_save = QPushButton("💾 SAVE ALL SETTINGS")
        btn_save.setObjectName("main_action_btn")
        btn_save.setFixedWidth(200)
        btn_save.clicked.connect(self._on_save_rules_clicked)
        
        btn_reset = QPushButton("🔄 Reset to Default")
        btn_reset.setObjectName("secondary_action_btn")
        btn_reset.setFixedWidth(180)
        btn_reset.clicked.connect(self._on_reset_rules_clicked)
        
        btn_layout.addWidget(btn_reset)
        btn_layout.addWidget(btn_save)
        content_pane_layout.addLayout(btn_layout)
        
        # Removed stretch from here (moved to top)
        
        content_scroll.setWidget(content_pane)
        
        # Load current rules
        self.current_rules = RulesManager.load_rules()
        self.rule_inputs = {}
        
        # Sort by section ID
        def sort_key(s):
            parts = s.split('.')
            return [int(p) if p.isdigit() else p for p in parts]
            
        sorted_sections = sorted(self.current_rules.keys(), key=sort_key)
        
        for sid in sorted_sections:
            rule = self.current_rules[sid]
            
            # 1. Add to sidebar
            item = QListWidgetItem(f"{sid}. {rule.get('title', 'Unknown')}")
            self.rules_sidebar.addItem(item)
            
            # 2. Add to stack - JSON Editor Page
            section_page = QWidget()
            page_layout = QVBoxLayout(section_page)
            page_layout.setContentsMargins(0, 0, 0, 10)
            
            header_lbl = QLabel(f"Configuration Rule {sid}: {rule.get('title')}")
            header_lbl.setStyleSheet("font-weight: bold; font-size: 15px; color: #1565C0; margin-bottom: 2px;")
            page_layout.addWidget(header_lbl)
            
            # Create JSON Editor with LINE NUMBERS
            editor = CodeEditor()
            editor.setObjectName("jsonEditor")
            # Modern Code Styling for the Editor
            editor.setStyleSheet("""
                QPlainTextEdit#jsonEditor {
                    background-color: #fcfcfc;
                    border: 1px solid #d1d9e6;
                    border-radius: 15px;
                    padding: 15px;
                    color: #333333;
                    selection-background-color: #CCE5FF;
                    selection-color: #004085;
                }
            """)
            
            # Set Monospaced Font
            font = QFont("Consolas", 11)
            if not font.fixedPitch(): # Fallback
                font = QFont("Courier New", 10)
            editor.setFont(font)
            editor.setTabStopDistance(QFontMetricsF(editor.font()).horizontalAdvance(' ') * 4)
            editor.setLineWrapMode(QPlainTextEdit.WidgetWidth) # Enable wrapping
            
            # Initialize with JSON
            json_text = json.dumps(rule, indent=4, ensure_ascii=False)
            editor.setPlainText(json_text)
            
            page_layout.addWidget(editor)
            self.rule_inputs[sid] = editor
            
            self.rules_stack.addWidget(section_page)
        
        # Sidebar connection
        self.rules_sidebar.currentRowChanged.connect(self.rules_stack.setCurrentIndex)
        self.rules_sidebar.setCurrentRow(0)
        
        main_hbox.addWidget(self.rules_sidebar)
        main_hbox.addWidget(content_scroll, stretch=1)
        
        widget.setLayout(main_hbox)
        return widget

    def _on_save_rules_clicked(self):
        """Collect all JSON values, validate, and save"""
        new_rules = {}
        errors = []
        
        for sid, editor in self.rule_inputs.items():
            try:
                text = editor.toPlainText()
                rule_data = json.loads(text)
                new_rules[sid] = rule_data
            except json.JSONDecodeError as e:
                errors.append(f"Section {sid}: {str(e)}")
        
        if errors:
            error_msg = "JSON Syntax Errors found:\n\n" + "\n".join(errors)
            QMessageBox.critical(self, "Invalid JSON", error_msg)
            return
            
        if RulesManager.save_rules(new_rules):
            self.current_rules = new_rules
            QMessageBox.information(self, "Success", "All recommendation settings have been saved successfully.")
            self.statusBar().showMessage("Rules saved.", 3000)
        else:
            QMessageBox.critical(self, "Error", "Failed to write rules to file.")

    def _on_reset_rules_clicked(self):
        """Reset everything to factory defaults and refresh editors"""
        reply = QMessageBox.question(self, 'Reset Rules', 
                                   "Restore ALL rules to factory defaults?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            defaults = RulesManager.reset_rules()
            self.current_rules = defaults
            
            # Update all editors with new default values
            for sid, editor in self.rule_inputs.items():
                if sid in defaults:
                    json_text = json.dumps(defaults[sid], indent=4, ensure_ascii=False)
                    editor.setPlainText(json_text)
            
            QMessageBox.information(self, "Success", "Rules have been reset to default values.")
            self.statusBar().showMessage("Rules reset to default.", 3000)
