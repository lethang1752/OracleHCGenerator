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
        'REPORT_DETAILS': '+REPORT DETAILS',
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
                            # Standardize headers to uppercase for easier lookup
                            table_data[0] = [h.upper().replace('\n', ' ') for h in table_data[0]]
                            self.data[section_key] = table_data
                            seen_keys.add(section_key)
                            logger.info(f"Found section: {section_key} ({len(table_data)} rows)")

    def get_backup_schedule(self) -> Dict[str, str]:
        """
        Analyze backup patterns and return formatted strings for Level 0, Level 1, and Archive.
        Returns: { 'level_0': '...', 'level_1': '...', 'archive': '...' }
        """
        results = {'level_0': 'NOT FOUND', 'level_1': 'NOT FOUND', 'archive': 'NOT FOUND'}
        
        backup_table = self.get_table('CHECK_BACKUP')
        if not backup_table or len(backup_table) < 2:
            return results

        # 1. Calculate Reference Size (Used Data)
        db_used_gb = 0
        ts_table = self.get_table('TABLESPACE')
        if ts_table and len(ts_table) > 1:
            headers = ts_table[0]
            used_idx = next((i for i, h in enumerate(headers) if 'USED (MB)' in h), None)
            if used_idx is not None:
                for row in ts_table[1:]:
                    try:
                        db_used_gb += float(row[used_idx].replace(',', '')) / 1024
                    except: continue

        # 2. Parse and Classify Backups
        headers = backup_table[0]
        try:
            input_idx = next(i for i, h in enumerate(headers) if 'INPUT GBYTES' in h)
            output_idx = next(i for i, h in enumerate(headers) if 'OUTPUT GBYTES' in h)
            status_idx = next(i for i, h in enumerate(headers) if 'STATUS' in h)
            time_idx = next(i for i, h in enumerate(headers) if 'START TIME' in h)
            dow_idx = next(i for i, h in enumerate(headers) if 'DOW' in h)
        except StopIteration:
            logger.warning("Required backup columns missing")
            return results

        classified = {'level_0': [], 'level_1': [], 'archive': []}
        
        # Get unique daily slots to detect patterns
        for row in backup_table[1:]:
            if len(row) <= max(input_idx, output_idx, status_idx, time_idx, dow_idx): continue
            
            status = row[status_idx].upper()
            if status != 'COMPLETED': continue
            
            try:
                inp = float(row[input_idx].replace(',', ''))
                outp = float(row[output_idx].replace(',', ''))
                dow = row[dow_idx].upper()
                time_str = row[time_idx].split(' ')[1] # HH:MM
                
                # Heuristic Logic
                is_full_scan = (inp > db_used_gb * 0.6) or (inp > 100 and db_used_gb < 100)
                
                if is_full_scan and (outp > 100 or outp > inp * 0.2): 
                    lvl = 'level_0'
                elif is_full_scan:
                    lvl = 'level_1'
                elif inp > 10: # Archive logs but with some volume
                    lvl = 'archive'
                else: 
                    lvl = 'archive'
                
                classified[lvl].append({'dow': dow, 'time': time_str})
            except: continue

        # 3. Format Strings
        def format_schedule(entries):
            if not entries: return 'NOT FOUND'
            
            # Group by unique time slots
            time_to_days = {}
            for e in entries:
                t = e['time']
                if t not in time_to_days: time_to_days[t] = set()
                time_to_days[t].add(e['dow'])
            
            parts = []
            for t, days in time_to_days.items():
                # Detect standard ranges
                day_str = ""
                all_days = {'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'}
                if all_days.issubset(days):
                    day_str = "Monday to Sunday"
                elif days == {'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY'}:
                    day_str = "Monday to Friday"
                elif len(days) == 1:
                    day_str = list(days)[0].capitalize()
                else:
                    day_str = ", ".join([d.capitalize() for d in sorted(list(days))])
                
                parts.append(f"{day_str} (about time: {t})")
            
            return " / ".join(parts)

        # Handle Archive specially if it's very frequent
        archive_entries = classified['archive']
        if archive_entries:
            # Check if it's bi-hourly or similar (frequency > 4/day)
            times = sorted(list(set([e['time'] for e in archive_entries])))
            if len(times) >= 4:
                # Merge into one line if consistent
                # Find days it happens on
                days = set([e['dow'] for e in archive_entries])
                day_str = "Monday to Sunday" if len(days) >= 7 else "Daily"
                results['archive'] = f"{day_str} (about time: {'; '.join(times)})"
            else:
                results['archive'] = format_schedule(archive_entries)
        
        results['level_0'] = format_schedule(classified['level_0'])
        results['level_1'] = format_schedule(classified['level_1'])
        
        return results
    
    def get_table(self, section_key: str) -> List[List[str]]:
        """Get a specific table by section key"""
        return self.data.get(section_key, [])
    
    def get_all_data(self) -> Dict[str, Any]:
        """Return all extracted data"""
        return self.data
    
    def has_section(self, section_key: str) -> bool:
        """Check if section exists"""
        return section_key in self.data
