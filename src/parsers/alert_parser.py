"""
Alert Log Parser - Parse Oracle Alert Logs
"""
import os
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Any
import logging
import re

from .base_parser import BaseParser
from ..config import (
    ALERT_LOG_PATTERN, 
    NUM_DAYS_ALERT,
    DATETIME_FORMAT_12C,
    DATETIME_FORMAT_11G
)

logger = logging.getLogger(__name__)


class AlertError:
    """Represents an alert log error"""
    
    def __init__(self, timestamp: str, message: str, error_code: str = None, full_text: str = None):
        self.timestamp = timestamp
        self.message = message  # First line of error
        self.error_code = error_code or self._extract_error_code(message)
        self.full_text = full_text or message  # Complete error block
    
    def _extract_error_code(self, message: str) -> str:
        """Extract ORA error code from message"""
        if not message: return None
        match = re.search(r'ORA-\d+', message)
        return match.group(0) if match else None
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            "timestamp": self.timestamp,
            "message": self.message,
            "error_code": self.error_code,
            "full_text": self.full_text
        }


class AlertLogParser(BaseParser):
    """Parse Oracle Alert Logs"""
    
    # Pre-compiled regex for performance
    TIMESTAMP_12C_RE = re.compile(r'^[2][0][0-9]{2}[-][0-9]{2}[-][0-9]{2}[T][0-9]{2}[:][0-9]{2}[:][0-9]{2}')
    TIMESTAMP_11G_RE = re.compile(r'^\w{3} \w{3} \d{2} \d{2}:\d{2}:\d{2} \d{4}$')
    
    def __init__(self, log_dir: str, num_days: int = NUM_DAYS_ALERT, max_lines: int = None):
        """
        Initialize Alert Log Parser
        
        Args:
            log_dir: Directory containing alert log files
            num_days: Filter alerts from last N days
            max_lines: Max lines to scan back
        """
        super().__init__(log_dir)
        self.num_days = num_days
        from ..config import ALERT_MAX_LINES
        self.max_lines = max_lines or ALERT_MAX_LINES
        self.alerts: List[AlertError] = []
        self.db_name = None
        self.instance_name = None
    
    def parse(self) -> bool:
        """
        Parse alert log files efficiently by reading backward from the end
        """
        try:
            alert_files = list(Path(self.log_dir).glob(ALERT_LOG_PATTERN))
            if not alert_files:
                self.add_error(f"No alert log found in {self.log_dir}")
                return False
            
            alert_file = alert_files[0]
            logger.info(f"Parsing alert log (Reverse, limit {self.max_lines} lines): {alert_file}")
            
            # Extract instance name
            filename = alert_file.stem
            self.instance_name = filename.replace('alert_', '') if filename.startswith('alert_') else filename
            
            # Calculate cutoff date
            mtime = os.path.getmtime(alert_file)
            file_time = datetime.fromtimestamp(mtime)
            cutoff_date = file_time - timedelta(days=self.num_days)
            
            # Read and parse backward
            self._parse_backward(alert_file, cutoff_date)
            
            logger.info(f"Found {len(self.alerts)} errors in scan")
            return True
        except Exception as e:
            self.add_error("Failed to parse alert log", e)
            return False

    def _parse_backward(self, file_path: Path, cutoff_date: datetime):
        """Read file from end in chunks and stop when date or line limit reached"""
        chunk_size = 65536 # 64KB
        line_count = 0
        
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            f.seek(0, os.SEEK_END)
            file_size = f.tell()
            pointer = file_size
            
            buffer = ""
            current_message_lines = []
            
            while pointer > 0:
                step = min(pointer, chunk_size)
                pointer -= step
                f.seek(pointer)
                
                new_chunk = f.read(step)
                buffer = new_chunk + buffer
                lines = buffer.splitlines()
                
                # Keep the first line fragment in the buffer
                if pointer > 0:
                    buffer = lines[0]
                    lines_to_process = lines[1:]
                else:
                    buffer = ""
                    lines_to_process = lines
                
                # Process lines from this chunk in reverse
                for line in reversed(lines_to_process):
                    line_count += 1
                    if line_count > self.max_lines:
                        logger.info(f"Reached max scan lines limit: {self.max_lines}")
                        return

                    # 1. Check for 12c Timestamp
                    is_12c = self._is_timestamp_12c(line)
                    if is_12c:
                        try:
                            ts = datetime.strptime(line[:19], DATETIME_FORMAT_12C)
                            if ts < cutoff_date:
                                return # STOP: Reached date threshold
                            
                            # Valid timestamp found, finalize previous message block
                            self._check_and_save_reverse_block(line[:19], current_message_lines)
                            current_message_lines = []
                        except:
                            pass
                        continue
                    
                    # 2. Check for 11g Timestamp
                    is_11g = self._is_timestamp_11g(line)
                    if is_11g:
                        try:
                            ts = datetime.strptime(line, DATETIME_FORMAT_11G)
                            if ts < cutoff_date:
                                return # STOP
                            
                            ts_str = ts.strftime("%Y-%m-%dT%H:%M:%S")
                            self._check_and_save_reverse_block(ts_str, current_message_lines)
                            current_message_lines = []
                        except:
                            pass
                        continue
                        
                    # 3. Accumulated non-timestamp line
                    current_message_lines.insert(0, line)
    
    def _check_and_save_reverse_block(self, timestamp: str, lines: List[str]):
        """Check if block contains ORA error and save"""
        is_ora = any(l.startswith('ORA-') for l in lines)
        if is_ora:
            first_ora = next((l for l in lines if l.startswith('ORA-')), "")
            full_text = '\n'.join(lines)
            self.alerts.append(AlertError(timestamp, first_ora, first_ora, full_text))
            # Sort later since we are getting them newest first
    
    def _parse_stream(self, stream, cutoff_date: datetime):
        """Parse alert log stream"""
        current_timestamp = None
        current_message = []
        is_ora = False
        
        for line in stream:
            line = line.rstrip('\n')
            
            # Check for timestamp (12c format: YYYY-MM-DDTHH:MM:SS)
            if self._is_timestamp_12c(line):
                # Save previous error if exists
                if is_ora and current_timestamp:
                    self._save_alert(current_timestamp, current_message)
                
                try:
                    timestamp_str = line[:19]
                    ts = datetime.strptime(timestamp_str, DATETIME_FORMAT_12C)
                    
                    if ts > cutoff_date:
                        current_timestamp = timestamp_str
                        current_message = []
                        is_ora = False
                    else:
                        current_timestamp = None
                except:
                    pass
            
            # Check for 11g format: DDD MMM DD HH:MM:SS YYYY
            elif self._is_timestamp_11g(line):
                if is_ora and current_timestamp:
                    self._save_alert(current_timestamp, current_message)
                
                try:
                    ts = datetime.strptime(line, DATETIME_FORMAT_11G)
                    if ts > cutoff_date:
                        current_timestamp = ts.strftime("%Y-%m-%dT%H:%M:%S")
                        current_message = []
                        is_ora = False
                    else:
                        current_timestamp = None
                except:
                    pass
            
            # Handle non-timestamp lines
            else:
                if current_timestamp is not None:
                    current_message.append(line)
                    if line.startswith('ORA-'):
                        is_ora = True
        
        # Save last error
        if is_ora and current_timestamp:
            self._save_alert(current_timestamp, current_message)
    
    def _is_timestamp_12c(self, line: str) -> bool:
        """Fast check then Regex for 12c timestamp"""
        if len(line) < 19 or line[0] != '2': # Quick filter
            return False
        return bool(self.TIMESTAMP_12C_RE.match(line))
    
    def _is_timestamp_11g(self, line: str) -> bool:
        """Fast check then Regex for 11g timestamp"""
        if len(line) < 24 or not line[0].isalpha(): # Quick filter
            return False
        return bool(self.TIMESTAMP_11G_RE.match(line))
    
    def _save_alert(self, timestamp: str, message_lines: List[str]):
        """Save alert to list - capture only ORA- code for identification"""
        if message_lines:
            first_ora_line = ""
            error_code_only = "ORA-UNKNOWN"
            
            import re
            ora_pattern = re.compile(r'(ORA-\d+)')
            
            for line in message_lines:
                if line.startswith('ORA-'):
                    first_ora_line = line
                    match = ora_pattern.search(line)
                    if match:
                        error_code_only = match.group(1)
                    else:
                        error_code_only = line[:9].strip(':') # Fallback
                    break
            
            # Join with newline to keep full block context
            full_text = '\n'.join(message_lines)
            
            alert = AlertError(timestamp, first_ora_line, error_code=error_code_only, full_text=full_text)
            self.alerts.append(alert)
    
    def get_data(self) -> Dict[str, Any]:
        """Return parsed alerts as dictionary"""
        return {
            "alerts": [alert.to_dict() for alert in self.alerts],
            "count": len(self.alerts),
            "num_days": self.num_days,
            "instance_name": self.instance_name
        }
