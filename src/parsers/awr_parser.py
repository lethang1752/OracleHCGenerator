"""
AWR Parser - Parse Oracle AWR HTML Reports
"""
import os
from pathlib import Path
from typing import List, Dict, Any
import logging
import re
from html.parser import HTMLParser

from .base_parser import BaseParser
from ..config import (
    AWR_REPORT_PATTERN,
    AWR_TABLES_12C,
    AWR_TABLES_11G
)

logger = logging.getLogger(__name__)


class AWRTable:
    """Represents an extracted AWR table"""
    
    def __init__(self, title: str, rows: List[List[str]]):
        self.title = title
        self.rows = rows
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            "title": self.title,
            "rows": self.rows,
            "row_count": len(self.rows)
        }


class SimpleHTMLTableExtractor(HTMLParser):
    """Extract tables from HTML"""
    
    def __init__(self):
        super().__init__()
        self.tables = []
        self.current_table = None
        self.current_row = None
        self.current_cell = None
        self.table_summary = None
    
    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            attrs_dict = dict(attrs)
            self.table_summary = attrs_dict.get('summary', '')
            self.current_table = []
        elif tag == 'tr' and self.current_table is not None:
            self.current_row = []
        elif tag in ('td', 'th') and self.current_row is not None:
            self.current_cell = []
    
    def handle_endtag(self, tag):
        if tag == 'table' and self.current_table is not None:
            if self.current_table:
                self.tables.append({
                    'summary': self.table_summary,
                    'rows': self.current_table
                })
            self.current_table = None
            self.table_summary = None
        elif tag == 'tr' and self.current_row is not None:
            if self.current_row:
                self.current_table.append(self.current_row)
            self.current_row = None
        elif tag in ('td', 'th') and self.current_cell is not None:
            text = ''.join(self.current_cell).strip()
            self.current_row.append(text)
            self.current_cell = None
    
    def handle_data(self, data):
        if self.current_cell is not None:
            self.current_cell.append(data)


from bs4 import BeautifulSoup

class AWRParser(BaseParser):
    """Parse Oracle AWR Reports (HTML)"""
    
    def __init__(self, log_dir: str):
        """
        Initialize AWR Parser
        
        Args:
            log_dir: Directory containing alert log files
        """
        super().__init__(log_dir)
        self.awr_tables: List[AWRTable] = []
        self.db_version = None
        self.db_name = None
        self.instance_name = None
    
    def parse(self) -> bool:
        """
        Parse alert log files
        
        Returns:
            True if successful, False if failed
        """
        try:
            # Find alert log file
            awr_files = list(Path(self.log_dir).glob(AWR_REPORT_PATTERN))
            
            if not awr_files:
                self.add_error(f"No AWR report found in {self.log_dir}")
                return False
            
            awr_file = awr_files[0]
            logger.info(f"Parsing AWR: {awr_file}")
            
            # Read HTML
            with open(awr_file, 'r', encoding='utf-8', errors='ignore') as f:
                html_content = f.read()
            
            # Extract tables
            soup = BeautifulSoup(html_content, 'lxml')
            self._extract_tables_bs(soup)
            
            # Extract metadata
            self._extract_metadata()
            
            logger.info(f"Found {len(self.awr_tables)} tables in AWR")
            return True
        
        except Exception as e:
            self.add_error("Failed to parse AWR report", e)
            return False
            
    def _extract_tables_bs(self, soup):
        """Extract tables using BeautifulSoup - much faster and more reliable"""
        tables = soup.find_all('table')
        filters = [f.lower().strip() for f in self._get_table_filters()]
        
        for table in tables:
            summary = table.get('summary', '').lower().strip()
            if not summary:
                # Fallback: check caption if any
                caption = table.find('caption')
                if caption:
                    summary = caption.get_text().lower().strip()
            
            for filter_text in filters:
                if filter_text in summary:
                    rows_data = []
                    for tr in table.find_all('tr'):
                        cells = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
                        if cells:
                            rows_data.append(cells)
                    
                    if rows_data:
                        # Use title/summary as the title
                        self.awr_tables.append(AWRTable(table.get('summary') or summary, rows_data))
                        break

    def _get_table_filters(self) -> List[str]:
        """Get table filter based on all known versions to be robust"""
        # Return unique combined list of all tables we ever care about
        return list(set(AWR_TABLES_12C + AWR_TABLES_11G))
    
    def _extract_metadata(self):
        """Extract metadata like DB name, instance name"""
        if self.awr_tables and len(self.awr_tables[0].rows) > 1:
            first_table = self.awr_tables[0].rows
            try:
                if len(first_table[1]) > 0:
                    self.db_name = first_table[1][0]
                if len(first_table[1]) > 1:
                    self.instance_name = first_table[1][1]
                if len(first_table[1]) > 2:
                    self.db_version = first_table[1][2]
            except:
                pass
    
    def get_data(self) -> Dict[str, Any]:
        """Return parsed AWR data"""
        return {
            "tables": [table.to_dict() for table in self.awr_tables],
            "table_count": len(self.awr_tables),
            "db_name": self.db_name,
            "instance_name": self.instance_name,
            "db_version": self.db_version
        }
