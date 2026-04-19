"""
Final Report Generator - AI Recommendations Summary
Generates the summary report with specific segments requested by the user.
"""
import logging
from pathlib import Path
from typing import Dict, Any, List
from datetime import datetime

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml, OxmlElement

from ..config import TEMPLATE_DIR

logger = logging.getLogger(__name__)

class FinalReportGenerator:
    """Generates the summary Final Report (DOCX)"""
    
    def __init__(self, output_path: str, font_option: str = 'times', language: str = 'vi'):
        self.output_path = output_path
        self.font_option = font_option.lower()
        self.language = language.lower() # 'vi' or 'en'
        
        # Select template
        template_name = 'timenr_template_report.docx' if self.font_option == 'times' else 'calibri_template_report.docx'
        template_path = TEMPLATE_DIR / template_name
        
        try:
            if template_path.exists():
                self.doc = Document(str(template_path))
                logger.info(f"Initialized Final Report with template: {template_path}")
            else:
                logger.warning(f"Template {template_path} not found. Using default document.")
                self.doc = Document()
        except Exception as e:
            logger.error(f"Error loading report template: {e}")
            self.doc = Document()

        self.font_name = 'Times New Roman' if self.font_option == 'times' else 'Calibri'
        self._setup_styles()

    def _setup_styles(self):
        """Configure default fonts and styles"""
        try:
            style = self.doc.styles['Normal']
            style.font.name = self.font_name
            style.font.size = Pt(11)
            
            def force_font(style_obj, name):
                rPr = style_obj._element.rPr
                if rPr is None:
                    rPr = style_obj._element.make_element(qn('w:rPr'))
                    style_obj._element.append(rPr)
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{name}" w:hAnsi="{name}" w:cs="{name}" w:eastAsia="{name}"/>')
                    rPr.append(rFonts)
                else:
                    rFonts.set(qn('w:ascii'), name)
                    rFonts.set(qn('w:hAnsi'), name)
            
            force_font(style, self.font_name)
            
            headings = [('Heading 1', 16), ('Heading 2', 13), ('Heading 3', 11)]
            for name, size in headings:
                if name in self.doc.styles:
                    h = self.doc.styles[name]
                    h.font.name = self.font_name
                    h.font.size = Pt(size)
                    h.font.bold = True
                    h.font.color.rgb = RGBColor(0, 0, 0)
                    force_font(h, self.font_name)
        except Exception as e:
            logger.warning(f"Report style setup failed: {e}")

    def generate(self, parsed_data: Dict[str, Any], db_name: str = None) -> bool:
        """Main entry point to build the report"""
        try:
            # We only use data from Node 1 (index 0)
            nodes = parsed_data.get('nodes', [])
            if not nodes:
                logger.error("No node data found for report generation")
                return False
            
            node1 = nodes[0]
            if not db_name:
                db_name = str(parsed_data.get('db_name', 'UNKNOWN')).upper()
            else:
                db_name = db_name.upper()
            
            # --- Remove potential blank line at start (common in templates) ---
            if self.doc.paragraphs and not self.doc.paragraphs[0].text.strip():
                p = self.doc.paragraphs[0]._element
                p.getparent().remove(p)

            # --- 3.1 DB NAME ---
            self.doc.add_heading(db_name, level=2)
            
            # --- 3.1.1 General Information ---
            self._add_general_info_section(node1)
            
            # --- 3.1.2 Evaluation / Summary ---
            self._add_evaluation_section()
            
            # --- Page Break ---
            self.doc.add_page_break()
            
            # --- 3.1.3 Recommendation ---
            self._add_recommendation_section()
            
            # --- 3.1.4 Appendix Reference ---
            self._add_appendix_ref_section(db_name)
            
            # Save
            self.doc.save(self.output_path)
            logger.info(f"Final report saved to: {self.output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to generate final report: {e}")
            return False

    def _add_general_info_section(self, node_data: Dict[str, Any]):
        title = "Thông tin tổng quan" if self.language == 'vi' else "General information"
        self.doc.add_heading(title, level=3)
        
        # Table with 3 columns, 12 rows (header + 11 data?) 
        # User said: "bảng với 3 cột và 12 dòng, header row là 'Mục', 'Thông tin', 'Thông tin (Báo cáo trước đây)'"
        # 1 header + 11 data = 12 rows.
        table = self.doc.add_table(rows=12, cols=3)
        table.style = 'Table Grid'
        table.allow_autofit = False
        
        # Header
        headers = ["Mục", "Thông tin", "Thông tin (Báo cáo trước đây)"] if self.language == 'vi' else \
                  ["ITEM", "VALUE", "VALUE (PREVIOUS REPORT)"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
        self._format_header_row(table.rows[0])
        
        # Fixed widths: 1 - 4.4 cm, 2 - 6.0 cm, 3 - 6.0 cm (Total 16.4cm)
        self._set_column_widths(table, [Cm(4.4), Cm(6.0), Cm(6.0)])
        
        # Data from REPORT_DETAILS (Node 1)
        report_details = node_data.get('database_info', {}).get('REPORT_DETAILS', [])
        # report_details[0] is likely header, data starts from index 1
        data_rows = report_details[1:] if len(report_details) > 1 else []
        
        for i in range(1, 12):
            idx = i - 1
            if idx < len(data_rows):
                row_data = data_rows[idx]
                if len(row_data) >= 2:
                    table.rows[i].cells[0].text = str(row_data[0])
                    table.rows[i].cells[1].text = str(row_data[1])
            
            # Column 3 is always empty
            table.rows[i].cells[2].text = ""

            # Formatting: Vertical center for all, Column 1 is Left-aligned
            for j in range(3):
                align = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.LEFT
                # User specifically asked: "cột 1 ở mục 1.1.1. cần set align center left"
                # Center vertical is already handled globally if I add it to _set_column_widths or here
                cell = table.rows[i].cells[j]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    p.alignment = align

    def _add_evaluation_section(self):
        title = "Đánh giá" if self.language == 'vi' else "SUMMARY"
        self.doc.add_heading(title, level=3)
        
        # Table 4 cols, 11 rows. Headers: "STT", "Kiểm tra", "Đánh giá", "Đánh giá (Báo cáo trước đây)"
        headers = ["STT", "Kiểm tra", "Đánh giá", "Đánh giá (Báo cáo trước đây)"] if self.language == 'vi' else \
                  ["NO", "CRITERIA", "EVALUATION SCORE", "EVALUATION SCORE (PREVIOUS REPORT)"]
        
        table = self.doc.add_table(rows=11, cols=4)
        table.style = 'Table Grid'
        table.allow_autofit = False
        
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
        self._format_header_row(table.rows[0])
        
        # Widths: 1.4, 10.0, 2.5, 2.5 cm
        self._set_column_widths(table, [Cm(1.4), Cm(10.0), Cm(2.5), Cm(2.5)])
        
        if self.language == 'vi':
            criteria = [
                "Kiểm tra trạng thái",
                "Kiểm tra hiệu năng - CPU",
                "Kiểm tra hiệu năng - Bộ nhớ",
                "Kiểm tra hiệu năng - Storage (còn trống) trên mỗi vùng đĩa",
                "Kiểm tra hiệu năng - Storage (I/O Wait)",
                "Kiểm tra hiệu năng – SQL",
                "Kiểm tra Firmware và bản vá",
                "Kiểm tra sao lưu",
                "Kiểm tra đồng bộ Active Dataguard",
                "Total"
            ]
        else:
            criteria = [
                "Status check",
                "Performance check – CPU",
                "Performance check – Memory",
                "Performance check – Storage (Free capacity) for each volume",
                "Performance check – Storage (I/O check) (Average/max)",
                "Performance check – SQL Execution",
                "Firmware and Patch check",
                "Backup check",
                "Standby check",
                "Total"
            ]
            
        for i in range(1, 11):
            table.rows[i].cells[0].text = str(i) if i < 10 else "" # Last row first cell empty
            table.rows[i].cells[1].text = criteria[i-1]
            
            # Formatting: Vertical center for all, Column 1 is Centered
            for j in range(4):
                align = WD_ALIGN_PARAGRAPH.CENTER if j == 0 else WD_ALIGN_PARAGRAPH.LEFT
                cell = table.rows[i].cells[j]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    p.alignment = align
            
        # Final row formatting: background D9D9D9, bold
        self._format_table_row_bg(table.rows[10], "D9D9D9", bold=True)

    def _add_recommendation_section(self):
        title = "Khuyến nghị" if self.language == 'vi' else "Recommendation"
        self.doc.add_heading(title, level=3)
        
        # Table 5 cols, 2 rows (header + 1 empty). 
        # Headers: STT, Khuyến nghị, Rủi ro/ảnh hưởng, Mức độ, Tham khảo
        headers = ["STT", "Khuyến nghị", "Rủi ro/ảnh hưởng", "Mức độ", "Tham khảo"] if self.language == 'vi' else \
                  ["NO", "RECOMMENDATION", "RISK/EFFECT", "SEVERITY", "REFERENCE"]
                  
        table = self.doc.add_table(rows=2, cols=5)
        table.style = 'Table Grid'
        table.allow_autofit = False
        
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
        self._format_header_row(table.rows[0])
        
        # Widths: 1.2, 6.55, 4.5, 1.9, 2.25 cm
        self._set_column_widths(table, [Cm(1.2), Cm(6.55), Cm(4.5), Cm(1.9), Cm(2.25)])

    def _add_appendix_ref_section(self, db_name: str):
        title = "Thông tin chi tiết của cơ sở dữ liệu" if self.language == 'vi' else \
                "Information of DB Health check"
        self.doc.add_heading(title, level=3)
        
        if self.language == 'vi':
            text = f"Tham khảo phụ lục báo cáo {db_name} đính kèm."
        else:
            text = f"Please refer to Section {db_name} in the Appendix attached along this Report."
        
        self.doc.add_paragraph(text)

    def _format_header_row(self, row):
        """Standard navy blue header formatting for consistency with Generator"""
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            shading_elm = parse_xml(r'<w:shd {} w:fill="1F3864"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(shading_elm)
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.bold = True
                    run.font.name = self.font_name

    def _format_table_row_bg(self, row, bg_color, bold=False):
        """Set background color for a whole row"""
        for cell in row.cells:
            shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), bg_color))
            cell._tc.get_or_add_tcPr().append(shading_elm)
            for para in cell.paragraphs:
                for run in para.runs:
                    if bold: run.font.bold = True
                    run.font.name = self.font_name

    def _set_column_widths(self, table, widths):
        """Force absolute fixed widths using internal XML manipulation (Twips/DXA)"""
        table.allow_autofit = False
        
        # 1. Set Table-level fixed layout and width
        tbl = table._tbl
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        
        # Fixed layout: Forces Word to respect column sizes exactly
        layout = tblPr.xpath('w:tblLayout')
        if not layout:
            layout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
            tblPr.append(layout)
        else:
            layout[0].set(qn('w:type'), 'fixed')
            
        # Total Table Width in DXA (Twips)
        # EMU / 635 = Twips
        total_emu = sum(w for w in widths)
        total_dxa = int(total_emu / 635)
        
        tblW = tblPr.xpath('w:tblW')
        if not tblW:
            tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_dxa}" w:type="dxa"/>')
            tblPr.append(tblW)
        else:
            tblW[0].set(qn('w:w'), str(total_dxa))
            tblW[0].set(qn('w:type'), 'dxa')

        # 2. Set Cell-level fixed widths for each column
        for i, width in enumerate(widths):
            width_dxa = int(width / 635)
            # Apply to column object first
            table.columns[i].width = width
            # Apply to each cell in the column via XML
            for cell in table.columns[i].cells:
                tcPr = cell._tc.get_or_add_tcPr()
                tcW = tcPr.xpath('w:tcW')
                if not tcW:
                    tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{width_dxa}" w:type="dxa"/>')
                    tcPr.append(tcW)
                else:
                    tcW[0].set(qn('w:w'), str(width_dxa))
                    tcW[0].set(qn('w:type'), 'dxa')

