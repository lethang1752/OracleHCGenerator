"""
Database Information Parser - Parse database_information.html
Extracts tables marked with +ASM, +TABLESPACE, etc.
"""
import logging
from pathlib import Path
from typing import Dict, Any, List
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)


class DatabaseInfoParser:
    """Parse database_information.html files"""
    
    # Map of section markers to table names
    SECTION_MARKERS = {
        'ASM': '+ASM',
        'TABLESPACE': '+TABLESPACE',
        'INDEX_FRAGMENT': '+INDEX_FRAGMENT',
        'INDEX_PARTITION_FRAGMENT': '+INDEX_PARTITION_FRAGMENT',
        'TABLE_FRAGMENT': '+TABLE_FRAGMENT',
        'TABLE_PARTITION_FRAGMENT': '+TABLE_PARTITION_FRAGMENT',
        'INVALID_OBJECT': '+INVALID_OBJECT',
        'TABLE_STATISTICS': '+TABLE_STATISTICS',
        'INDEX_STATISTICS': '+INDEX_STATISTICS',
        'CHECK_CLUSTER': '+CHECK_CLUSTER',
        'RESOURCE_CRS': '+RESOURCE_CRS',
        'DISK_USAGE': '+DISK_USAGE',
        'CHECK_BACKUP': '+CHECK_BACKUP',
        'BACKUP_POLICY': '+BACKUP_POLICY',
        'DBA_ROLE': '+DBA_ROLE',
        'OBJECT_IN_SYSTEM': '+OBJECT_IN_SYSTEM/SYSAUX',
        'CHECK_PATCHES': '+CHECK_PATCHES',
    }
    
    def __init__(self, log_dir: str):
        """Initialize parser"""
        self.log_dir = log_dir
        self.data = {}
        self.errors = []
    
    def parse(self) -> bool:
        """Parse database_information.html"""
        try:
            db_info_files = list(Path(self.log_dir).glob('database_information.html'))
            
            if not db_info_files:
                logger.warning(f"No database_information.html found in {self.log_dir}")
                return False
            
            db_info_file = db_info_files[0]
            logger.info(f"Parsing: {db_info_file}")
            
            with open(db_info_file, 'r', encoding='latin-1') as f:
                # html.parser is more robust for legacy HTML scripts
                soup = BeautifulSoup(f.read(), 'html.parser')
            
            # Extract tables by section markers
            self._extract_tables(soup)
            
            success = len(self.data) > 0
            logger.info(f"Extracted {len(self.data)} sections. Success: {success}")
            return success
        
        except Exception as e:
            logger.error(f"Error parsing database_information.html: {e}")
            return False
    
    def _extract_tables(self, soup):
        """Extract all tables using robust marker-to-table matching"""
        # Find all elements that could be markers
        tags = ['b', 'p', 'span', 'h1', 'h2', 'h3', 'font', 'center', 'td', 'div']
        potential_markers = soup.find_all(tags)
        
        seen_keys = set()
        
        for elem in potential_markers:
            text = elem.get_text(strip=True)
            if not text.startswith('+'):
                continue
                
            normalized_text = text.replace('+ ', '+').strip()
            section_key = None
            for key, marker in self.SECTION_MARKERS.items():
                if normalized_text == marker or normalized_text.startswith(marker):
                    section_key = key
                    break
            
            if section_key and section_key not in seen_keys:
                # Find the very next table relative to this marker element
                table = elem.find_next('table')
                if table:
                    rows = table.find_all('tr')
                    if not rows: continue
                    
                    table_data = []
                    for row in rows:
                        cols = row.find_all(['td', 'th'])
                        if cols:
                            row_data = [col.get_text(separator='\n', strip=True) for col in cols]
                            table_data.append(row_data)
                    
                    if table_data:
                        # Clean up trailing empty rows
                        while table_data and all(not str(cell).strip() for cell in table_data[-1]):
                            table_data.pop()
                        
                        if table_data:
                            self.data[section_key] = table_data
                            seen_keys.add(section_key)
                            logger.info(f"Found section: {section_key} ({len(table_data)} rows)")
    
    def get_table(self, section_key: str) -> List[List[str]]:
        """Get a specific table by section key"""
        return self.data.get(section_key, [])
    
    def get_all_data(self) -> Dict[str, Any]:
        """Return all extracted data"""
        return self.data
    
    def has_section(self, section_key: str) -> bool:
        """Check if section exists"""
        return section_key in self.data
