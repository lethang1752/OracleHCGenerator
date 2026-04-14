"""
Main Window UI - PyQt5
"""
import sys
from pathlib import Path
from typing import Optional
import logging
import shutil
import os

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
    QLabel, QLineEdit, QFileDialog, QProgressBar, QTextEdit,
    QTabWidget, QTableWidget, QTableWidgetItem, QComboBox,
    QSpinBox, QGroupBox, QMessageBox, QStatusBar, QDesktopWidget,
    QListWidget, QListWidgetItem, QAbstractItemView, QCheckBox, QGridLayout,
    QRadioButton, QHeaderView, QGraphicsDropShadowEffect, QStackedWidget,
    QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QEvent
from PyQt5.QtGui import QFont, QIcon, QColor

from ..parsers import AlertLogParser, AWRParser, DatabaseInfoParser
from ..generators.comprehensive_report_generator import ComprehensiveHealthcareReportGenerator
from ..config import (
    APP_NAME, APP_VERSION, WINDOW_WIDTH, WINDOW_HEIGHT,
    NUM_DAYS_ALERT, OUTPUT_DIR, COLLECT_TOOL_DIR
)
from ..utils import setup_logger, sanitize_filename

logger = setup_logger(__name__)

from ..parsers.alert_parser import AlertLogParser
from ..parsers.awr_parser import AWRParser
from ..parsers.database_info_parser import DatabaseInfoParser

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
        
        return {
            "node_id": node_id,
            "data_dir": data_dir,
            "alert_data": parser_alert.get_data(),
            "awr_data": parser_awr.get_data(),
            "db_info_data": parser_db_info.get_all_data()
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
        self.parsed_data = None
        self.parse_worker = None
        
        self._init_ui()
        self._apply_modern_shadows()
        self._load_stylesheet()
        self._ensure_tools_extracted()
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
        self.sidebar.addItem(QListWidgetItem("📄 Merge Documents"))
        self.sidebar.addItem(QListWidgetItem("🔍 Data Preview"))
        
        # Bottom Sidebar for support tools
        self.sidebar_footer = QListWidget()
        self.sidebar_footer.setObjectName("sidebarMenu")
        self.sidebar_footer.setFocusPolicy(Qt.NoFocus)
        self.sidebar_footer.setFixedHeight(60) # Nới rộng để không bị mất text
        self.sidebar_footer.setContentsMargins(0, 0, 0, 0)
        self.sidebar_footer.setSpacing(0)
        
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
        self.stack.addWidget(self._create_merge_tab())      # Index 2
        self.stack.addWidget(self._create_preview_tab())    # Index 3
        self.stack.addWidget(self._create_tools_tab())      # Index 4
        
        content_layout.addWidget(self.stack)
        main_layout.addWidget(content_container)
        
        # Set initial selection AFTER connections
        self.sidebar.setCurrentRow(0)
        
        # Initial Title Update
        self._on_sidebar_changed(0)

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
            
            # Index 0-3 in main sidebar corresponds directly to index 0-3 in stack
            self.stack.setCurrentIndex(index)
            # Update title
            titles = ["APPENDIX GENERATOR", "OSWBB GRAPH GENERATOR", "MERGE DOCUMENTS", "DATA PREVIEW"]
            if index < len(titles):
                self.section_title.setText(titles[index])

    def _on_footer_tab_changed(self, index):
        """Handle switching from the footer sidebar items"""
        if index != -1:
            self.sidebar.blockSignals(True)
            self.sidebar.clearSelection()
            self.sidebar.setCurrentRow(-1)
            self.sidebar.blockSignals(False)
            
            # "Collection Tools" is the only item in footer, corresponds to index 4 in stack
            self.stack.setCurrentIndex(4)
            self.section_title.setText("COLLECTION TOOLS")

    def _create_main_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(15, 10, 15, 10)
        
        # 1. Folders Group
        folder_group = QGroupBox("1. DATA SOURCES (Multi-Node)")
        folder_layout = QVBoxLayout()
        
        list_action_box = QHBoxLayout()
        btn_add_node = QPushButton("+ Add Node Folder")
        btn_add_node.setObjectName("browse_btn")
        btn_add_node.clicked.connect(self._add_node_folder)
        
        btn_clear_nodes = QPushButton("Clear All")
        btn_clear_nodes.setObjectName("clear_btn")
        btn_clear_nodes.clicked.connect(self._clear_nodes)
        
        list_action_box.addWidget(btn_add_node)
        list_action_box.addWidget(btn_clear_nodes)
        list_action_box.addStretch()
        folder_layout.addLayout(list_action_box)
        
        self.node_list_widget = QListWidget()
        self.node_list_widget.setSelectionMode(QAbstractItemView.NoSelection)
        folder_layout.addWidget(self.node_list_widget)
        
        folder_group.setLayout(folder_layout)
        layout.addWidget(folder_group)
        
        # 2. Settings Group - Refined Layout
        settings_group = QGroupBox("2. APPENDIX SETTINGS")
        settings_layout = QGridLayout()
        settings_layout.setSpacing(10)
        
        # Column 0: Alert Logs
        settings_layout.addWidget(QLabel("Alert Logs (Days):"), 0, 0)
        self.num_days_spin = QSpinBox()
        self.num_days_spin.setValue(NUM_DAYS_ALERT)
        self.num_days_spin.setRange(1, 365)
        self.num_days_spin.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.num_days_spin.setMinimumHeight(32) # Match QLineEdit height
        settings_layout.addWidget(self.num_days_spin, 1, 0)
        
        # Column 1: Document Font
        settings_layout.addWidget(QLabel("Document Font:"), 0, 1)
        self.font_combo = QComboBox()
        self.font_combo.addItems(["Times New Roman", "Calibri"])
        self.font_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.font_combo.setMinimumHeight(32) # Match QLineEdit height
        settings_layout.addWidget(self.font_combo, 1, 1)
        
        # Column 2-3: Custom Filename (Giving more width)
        settings_layout.addWidget(QLabel("Custom Filename (Optional):"), 0, 2)
        self.filename_input = QLineEdit()
        self.filename_input.setPlaceholderText("Auto-generated if empty")
        self.filename_input.setMinimumHeight(32)
        settings_layout.addWidget(self.filename_input, 1, 2)
        
        # Adjusting column stretches to make Column 0 and 1 narrower
        settings_layout.setColumnStretch(0, 1)
        settings_layout.setColumnStretch(1, 1)
        settings_layout.setColumnStretch(2, 2) # Filename takes 50% width relatively
        
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # 3. Actions
        action_layout = QVBoxLayout()
        self.generate_btn = QPushButton("🚀 GENERATE APPENDIX (ALL NODES)")
        self.generate_btn.setObjectName("main_action_btn")
        self.generate_btn.clicked.connect(self._on_generate_clicked)
        action_layout.addWidget(self.generate_btn)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        action_layout.addWidget(self.progress_bar)
        layout.addLayout(action_layout)
        
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
        config_group = QGroupBox("1. SETUP GENERATOR")
        config_layout = QGridLayout()
        config_layout.setContentsMargins(12, 12, 12, 10) 
        config_layout.setSpacing(10)
        
        config_layout.addWidget(QLabel("Input Log Folder (OSWBB):"), 0, 0)
        self.oswbb_input_dir = QLineEdit()
        self.oswbb_input_dir.setPlaceholderText("Select the OSWBB archive folder...")
        config_layout.addWidget(self.oswbb_input_dir, 0, 1)
        btn_browse_in = QPushButton("Browse")
        btn_browse_in.setObjectName("browse_btn")
        btn_browse_in.clicked.connect(self._browse_oswbb_in)
        config_layout.addWidget(btn_browse_in, 0, 2)
        
        config_layout.addWidget(QLabel("Output Images Folder:"), 1, 0)
        self.oswbb_output_dir = QLineEdit()
        self.oswbb_output_dir.setPlaceholderText("Select where to save generated graphs...")
        config_layout.addWidget(self.oswbb_output_dir, 1, 1)
        btn_browse_out = QPushButton("Browse")
        btn_browse_out.setObjectName("browse_btn")
        btn_browse_out.clicked.connect(self._browse_oswbb_out)
        config_layout.addWidget(btn_browse_out, 1, 2)
        
        config_layout.addWidget(QLabel("Analysis Tool (JAR):"), 2, 0)
        self.oswbb_jar_select = QComboBox()
        self.oswbb_jar_select.addItems(["oswbba.jar", "oswbba9020.jar"])
        self.oswbb_jar_select.setMinimumHeight(32)
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
        
        btn_clear_target = QPushButton("🗑 Clear All")
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
        
        action_layout.addWidget(self.gen_oswbb_btn)
        action_layout.addWidget(self.gen_push_oswbb_btn)
        layout.addLayout(action_layout)
        
        # 4. CONSOLE LOG
        log_group = QGroupBox("JAVA CONSOLE LIVE")
        log_layout = QVBoxLayout()
        log_layout.setContentsMargins(12, 8, 12, 10)
        self.oswbb_log_text = QTextEdit()
        self.oswbb_log_text.setReadOnly(True)
        self.oswbb_log_text.setMinimumHeight(80) 
        log_layout.addWidget(self.oswbb_log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group, 1)
        
        widget.setLayout(layout)
        return widget
    
    def _create_preview_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        
        # Title label removed as requested
        
        preview_tabs = QTabWidget()
        # Đảm bảo tab bar đủ cao và không rút gọn chữ (elide)
        preview_tabs.tabBar().setFixedHeight(60) 
        preview_tabs.tabBar().setElideMode(Qt.ElideNone)
        
        self.alerts_table = QTableWidget()
        self.alerts_table.setColumnCount(3)
        self.alerts_table.setHorizontalHeaderLabels(["Node ID", "Timestamp", "Error Code"])
        self.alerts_table.horizontalHeader().setStretchLastSection(True)
        self.alerts_table.horizontalHeader().setFixedHeight(55) # Header bảng cũng cần cao ráo
        self.alerts_table.setColumnWidth(0, 100) # Node ID
        self.alerts_table.setColumnWidth(1, 250) # Timestamp (tăng chiều rộng)
        preview_tabs.addTab(self.alerts_table, "Alert Logs")
        
        self.awr_table = QTableWidget()
        self.awr_table.setColumnCount(2)
        self.awr_table.setHorizontalHeaderLabels(["Node ID", "Tables Found"])
        self.awr_table.horizontalHeader().setStretchLastSection(True)
        self.awr_table.horizontalHeader().setFixedHeight(55)
        self.awr_table.setColumnWidth(0, 100)
        preview_tabs.addTab(self.awr_table, "AWR Summary")
        
        self.db_info_table = QTableWidget()
        self.db_info_table.setColumnCount(2)
        self.db_info_table.setHorizontalHeaderLabels(["Node ID", "Sections Found"])
        self.db_info_table.horizontalHeader().setStretchLastSection(True)
        self.db_info_table.horizontalHeader().setFixedHeight(55)
        self.db_info_table.setColumnWidth(0, 100)
        preview_tabs.addTab(self.db_info_table, "HTML Summary")
        
        layout.addWidget(preview_tabs)
        widget.setLayout(layout)
        return widget

    def _add_node_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Node Folder")
        if folder:
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
            path = str(Path(folder))
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
            # Add tooltip for long paths
            item.setToolTip(folder)
            self.push_target_list.addItem(item)

    # ── Drag & Drop Handling ─────────────────────────────────────────
    def eventFilter(self, watched, event):
        if (watched == self.push_target_list or watched == self.merge_file_list) and event.type() == QEvent.DragEnter:
            if event.mimeData().hasUrls():
                event.accept()
                return True
        if (watched == self.push_target_list or watched == self.merge_file_list) and event.type() == QEvent.Drop:
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if watched == self.push_target_list:
                    if Path(path).is_dir() and path not in self.oswbb_push_folders:
                        self.oswbb_push_folders.append(path)
                        self._update_push_target_list()
                else: # merge_file_list
                    if path.lower().endswith(".docx"):
                        self._merge_add_direct(path)
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
            self.oswbb_input_dir.setText(folder)
            
    def _browse_oswbb_out(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.oswbb_output_dir.setText(folder)

    def _on_generate_oswbb_clicked(self, push=False):
        in_dir = self.oswbb_input_dir.text()
        out_dir = self.oswbb_output_dir.text()
        if not in_dir or not out_dir:
            QMessageBox.warning(self, "Missing fields", "Please select input and output directories.")
            return
            
        if push and not self.oswbb_push_folders:
            QMessageBox.warning(self, "No Targets", "Please add at least one target folder to push results.")
            return
            
        from ..utils.oswbb_runner import OSWBBGraphGenerator
        self.gen_oswbb_btn.setEnabled(False)
        self.gen_push_oswbb_btn.setEnabled(False)
        self.oswbb_log_text.append("[SYSTEM] Khởi chạy máy chủ đồ họa OSWBB Java...")
        
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
        if success:
            QMessageBox.information(self, "Thành công", "Tiến trình OSWBB hoàn tất. Vui lòng kiểm tra thư mục đích!")
        else:
            QMessageBox.warning(self, "Lỗi Java", "Tiến trình thu thập ảnh gặp lỗi hoặc quá trình Push thất bại, kiểm tra console log.")

    def _on_parse_clicked(self):
        num_days = self.num_days_spin.value()
        self.parse_worker = ParseWorker(self.log_folders, num_days)
        self.parse_worker.progress.connect(self._on_parse_progress)
        self.parse_worker.finished.connect(self._on_parse_finished)
        self.parse_worker.error.connect(self._on_parse_error)
        self.parse_worker.start()
    
    def _on_parse_progress(self, message: str, value: int):
        self._log(message)
        self.progress_bar.setValue(value)
    
    def _on_parse_finished(self, data: dict):
        self.parsed_data = data
        self._log("[SUCCESS] All data parsed successfully!")
        self._show_preview(data)
        self._run_generation_and_finalize()
    
    def _on_parse_error(self, error_msg: str):
        self._log(f"[ERROR] {error_msg}")
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        QMessageBox.critical(self, "Parse Error", error_msg)
    
    def _show_preview(self, data: dict):
        nodes = data.get('nodes', [])
        
        # Calculate total alerts
        total_alerts = sum(len(n['alerts'].get('alerts', [])) for n in nodes)
        self.alerts_table.setRowCount(total_alerts)
        
        self.awr_table.setRowCount(len(nodes))
        self.db_info_table.setRowCount(len(nodes))
        
        row_alert = 0
        for i, node in enumerate(nodes):
            node_id = str(node['node_id'])
            # Alerts
            for alert in node['alerts'].get('alerts', []):
                self.alerts_table.setItem(row_alert, 0, QTableWidgetItem(f"Node {node_id}"))
                self.alerts_table.setItem(row_alert, 1, QTableWidgetItem(alert.get('timestamp', '')))
                self.alerts_table.setItem(row_alert, 2, QTableWidgetItem(alert.get('error_code', '')))
                row_alert += 1
            
            # AWR summary
            table_count = str(node['awr'].get('table_count', 0))
            self.awr_table.setItem(i, 0, QTableWidgetItem(f"Node {node_id}"))
            self.awr_table.setItem(i, 1, QTableWidgetItem(table_count))
            
            # DB Info summary
            db_sections = str(len(node.get('database_info', {})))
            self.db_info_table.setItem(i, 0, QTableWidgetItem(f"Node {node_id}"))
            self.db_info_table.setItem(i, 1, QTableWidgetItem(db_sections))

    def _on_generate_clicked(self):
        if not self.log_folders:
            QMessageBox.warning(self, "Missing Input", "Please add at least one Node data folder.")
            return

        self.generate_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self._log(f"Starting workflow for {len(self.log_folders)} nodes...")
        self._on_parse_clicked()

    def _run_generation_and_finalize(self):
        self.progress_bar.setValue(92)
        self._log("Initializing dynamic multi-node report generation...")
        try:
            font_choice = self.font_combo.currentText()
            self.progress_bar.setValue(95)
            
            # Calculate Base DB Name
            nodes = self.parsed_data.get('nodes', [])
            base_db_name = "Unknown"
            if nodes:
                inst = str(nodes[0].get('instance_name', '')).strip()
                if not inst or inst == 'NODE1':
                    base_db_name = self.parsed_data.get('db_name', 'Unknown')
                else:
                    base_db_name = inst
                    while base_db_name and base_db_name[-1].isdigit():
                        base_db_name = base_db_name[:-1]
            base_db_name = base_db_name.upper()
            
            default_filename = f"{base_db_name}_appendix"
            filename = self.filename_input.text() or default_filename
            filename = sanitize_filename(filename)
            
            docx_path = OUTPUT_DIR / f"{filename}.docx"
            self._log(f"Generating DOCX with {font_choice} font...")
            
            gen = ComprehensiveHealthcareReportGenerator(str(docx_path), font_name=font_choice)
            if gen.generate_from_parsed_data(self.parsed_data):
                self._log(f"[SUCCESS] Report saved: {docx_path}")
                QMessageBox.information(self, "Complete", f"Workflow completed successfully!\nFile: {filename}.docx")
            else:
                self._log("[ERROR] Report generation failed.")
                QMessageBox.warning(self, "Failed", "Data was parsed but document generation failed.")
        
        except Exception as e:
            self._log(f"[ERROR] Fault: {str(e)}")
            logger.exception("Generator error")
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)
            self.generate_btn.setEnabled(True)
            self.statusBar().showMessage("Workflow Finished")
    
    def _log(self, message: str):
        self.log_text.append(message)
        logger.info(message)

    # ─────────────────────────────────────────────────────────────────
    #  TAB: Merge Documents
    # ─────────────────────────────────────────────────────────────────
    def _create_merge_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)

        # ── 1. File Queue ──────────────────────────────────────────
        queue_group = QGroupBox("1. DOCUMENTS TO MERGE")
        queue_layout = QVBoxLayout()

        # Action buttons row
        btn_row = QHBoxLayout()

        btn_add = QPushButton("➕ Add Files")
        btn_add.setObjectName("browse_btn")
        btn_add.clicked.connect(self._merge_add_files)

        btn_remove = QPushButton("🗑 Remove")
        btn_remove.setObjectName("clear_btn")
        btn_remove.clicked.connect(self._merge_remove_file)

        btn_up = QPushButton("▲ Move Up")
        btn_up.setObjectName("browse_btn")
        btn_up.clicked.connect(self._merge_move_up)

        btn_down = QPushButton("▼ Move Down")
        btn_down.setObjectName("browse_btn")
        btn_down.clicked.connect(self._merge_move_down)

        btn_clear_all = QPushButton("✖ Clear All")
        btn_clear_all.setObjectName("clear_btn")
        btn_clear_all.clicked.connect(self._merge_clear_all)

        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_remove)
        btn_row.addWidget(btn_up)
        btn_row.addWidget(btn_down)
        btn_row.addStretch()
        btn_row.addWidget(btn_clear_all)
        queue_layout.addLayout(btn_row)

        # Draggable list widget - no alternating colors to prevent ghost rows
        self.merge_file_list = QListWidget()
        self.merge_file_list.setObjectName("merge_list")
        self.merge_file_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.merge_file_list.setDragDropOverwriteMode(False)
        self.merge_file_list.setDefaultDropAction(Qt.MoveAction)
        self.merge_file_list.setSelectionMode(QAbstractItemView.SingleSelection)
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
        top_row.setSpacing(12)
        top_row.addWidget(queue_group, stretch=3)   # 60% chiều rộng
        top_row.addWidget(sort_group, stretch=2)    # 40% chiều rộng
        layout.addLayout(top_row, stretch=1)        # Chiếm toàn bộ không gian dọc còn lại

        # ── 3. Output Settings ─────────────────────────────────────
        output_group = QGroupBox("3. OUTPUT SETTINGS")
        output_grid = QGridLayout()
        output_grid.setSpacing(10)

        output_grid.addWidget(QLabel("Save Merged File As:"), 0, 0)
        self.merge_output_path = QLineEdit()
        self.merge_output_path.setPlaceholderText("Select destination file path (.docx)...")
        output_grid.addWidget(self.merge_output_path, 0, 1)

        btn_browse_merge = QPushButton("Browse")
        btn_browse_merge.setObjectName("browse_btn")
        btn_browse_merge.clicked.connect(self._browse_merge_output)
        output_grid.addWidget(btn_browse_merge, 0, 2)

        output_grid.setColumnStretch(1, 1)
        output_group.setLayout(output_grid)
        layout.addWidget(output_group)

        # ── 3. Execute ─────────────────────────────────────────────
        self.merge_btn = QPushButton("🔗 MERGE DOCUMENTS")
        self.merge_btn.setObjectName("main_action_btn")
        self.merge_btn.clicked.connect(self._on_merge_clicked)
        layout.addWidget(self.merge_btn)

        # ── 4. Progress Bar ────────────────────────────────────────
        self.merge_progress = QProgressBar()
        self.merge_progress.setRange(0, 100)
        self.merge_progress.setValue(0)
        self.merge_progress.setVisible(False)
        layout.addWidget(self.merge_progress)

        # ── 5. Live log ────────────────────────────────────────────
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
        
        btn_refresh = QPushButton("🔄 Refresh List")
        btn_refresh.setObjectName("secondary_action_btn")
        btn_refresh.clicked.connect(self._on_refresh_tools_clicked)
        
        btn_open_folder = QPushButton("📂 Open Collection Folder")
        btn_open_folder.setObjectName("main_action_btn")
        btn_open_folder.clicked.connect(self._on_open_tools_folder)
        
        btn_layout.addWidget(btn_refresh)
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
            QMessageBox.warning(self, "Lỗi", "Thư mục công cụ không tồn tại.")

    def _on_refresh_tools_clicked(self):
        """Manual check and refresh for tool list"""
        self._ensure_tools_extracted()
        self._refresh_tools_list()
        self.statusBar().showMessage("Đã cập nhật danh sách công cụ.", 3000)

    def _ensure_tools_extracted(self):
        """Nếu chạy từ EXE, tự động bung thư mục tool ra ngoài nếu chưa có"""
        if getattr(sys, 'frozen', False):
            # Trong chế độ Bundle, PyInstaller giải nén vào sys._MEIPASS
            bundle_tools = Path(getattr(sys, '_MEIPASS')) / "HC_collect_tool"
            if bundle_tools.exists() and not COLLECT_TOOL_DIR.exists():
                try:
                    logger.info(f"Extracting tools from {bundle_tools} to {COLLECT_TOOL_DIR}")
                    shutil.copytree(str(bundle_tools), str(COLLECT_TOOL_DIR))
                except Exception as e:
                    logger.error(f"Failed to auto-extract tools: {e}")

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
        """Sắp xếp lại queue file theo thứ tự danh sách DB người dùng nhập.
        
        Quy tắc: db_name = phần trước dấu '_' đầu tiên trong tên file.
        Ví dụ: DCFNGTB_appendix.docx -> db_name = DCFNGTB
        """
        raw_text = self.db_order_input.toPlainText().strip()
        if not raw_text:
            QMessageBox.warning(self, "Thiếu danh sách",
                                "Vui lòng nhập danh sách tên DB vào ô bên trên trước khi sắp xếp.")
            return

        if self.merge_file_list.count() == 0:
            QMessageBox.warning(self, "Chưa có file",
                                "Vui lòng thêm file vào danh sách trước khi sắp xếp.")
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

        # Thông báo kết quả
        msg_parts = [f"✅ Đã sắp xếp {len(sorted_paths)} file theo danh sách DB."]
        if not_found_dbs:
            msg_parts.append(f"\n⚠️ Không tìm thấy file cho: {', '.join(not_found_dbs)}")
        if unmatched:
            msg_parts.append(f"\n📌 Đặt xuống cuối (không khớp DB): {', '.join(unmatched)}")
        QMessageBox.information(self, "Sắp xếp hoàn tất", "\n".join(msg_parts))

    def _browse_merge_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Merged Document As", "merged_appendix.docx",
            "Word Documents (*.docx)"
        )
        if path:
            if not path.lower().endswith('.docx'):
                path += '.docx'
            self.merge_output_path.setText(path)

    def _on_merge_clicked(self):
        count = self.merge_file_list.count()
        if count < 2:
            QMessageBox.warning(self, "Not Enough Files",
                                "Please add at least 2 .docx files to merge.")
            return

        output = self.merge_output_path.text().strip()
        if not output:
            QMessageBox.warning(self, "Missing Output",
                                "Please specify a destination file path.")
            return

        ordered_paths = [
            self.merge_file_list.item(i).data(Qt.UserRole)
            for i in range(count)
        ]

        self.merge_btn.setEnabled(False)
        self.merge_log_text.clear()
        self.merge_progress.setValue(0)
        self.merge_progress.setVisible(True)
        self.merge_log_text.append(f"[SYSTEM] Starting merge of {count} documents...")

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
        self.merge_progress.setValue(percent)
        self.merge_log_text.append(f"[{percent:3d}%] {message}")

    def _on_merge_finished(self, success: bool, message: str):
        self.merge_btn.setEnabled(True)
        self.merge_progress.setValue(100 if success else 0)
        self.merge_progress.setVisible(False)
        self.merge_log_text.append(f"\n{'[SUCCESS]' if success else '[ERROR]'} {message}")
        if success:
            QMessageBox.information(self, "Merge Complete", message)
        else:
            QMessageBox.critical(self, "Merge Failed", message)
