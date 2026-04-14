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
    
    def __init__(self, log_dir: str, num_days: int = NUM_DAYS_ALERT):
        """
        Initialize Alert Log Parser
        
        Args:
            log_dir: Directory containing alert log files
            num_days: Filter alerts from last N days
        """
        super().__init__(log_dir)
        self.num_days = num_days
        self.alerts: List[AlertError] = []
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
            alert_files = list(Path(self.log_dir).glob(ALERT_LOG_PATTERN))
            
            if not alert_files:
                self.add_error(f"No alert log found in {self.log_dir}")
                return False
            
            alert_file = alert_files[0]
            logger.info(f"Parsing alert log: {alert_file}")
            
            # Extract instance name from filename (alert_<instance_name>.log)
            filename = alert_file.stem  # Gets 'alert_drewallet1' from 'alert_drewallet1.log'
            if filename.startswith('alert_'):
                self.instance_name = filename.replace('alert_', '')
            else:
                self.instance_name = filename
            
            logger.info(f"Instance name: {self.instance_name}")
            
            # Calculate cutoff date based on file's last modified time (like PS1)
            import os
            mtime = os.path.getmtime(alert_file)
            file_time = datetime.fromtimestamp(mtime)
            cutoff_date = file_time - timedelta(days=self.num_days)
            
            with open(alert_file, 'r', encoding='utf-8', errors='ignore') as f:
                self._parse_stream(f, cutoff_date)
            
            logger.info(f"Found {len(self.alerts)} errors in last {self.num_days} days")
            return True
        
        except Exception as e:
            self.add_error("Failed to parse alert log", e)
            return False
    
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
        """Check if line is 12c timestamp"""
        if len(line) != 32:
            return False
        pattern = r'^[2][0][0-9]{2}[-][0-9]{2}[-][0-9]{2}[T][0-9]{2}[:][0-9]{2}[:][0-9]{2}'
        return bool(re.match(pattern, line))
    
    def _is_timestamp_11g(self, line: str) -> bool:
        """Check if line is 11g timestamp"""
        if len(line) != 24:
            return False
        pattern = r'^\w{3} \w{3} \d{2} \d{2}:\d{2}:\d{2} \d{4}$'
        return bool(re.match(pattern, line))
    
    def _save_alert(self, timestamp: str, message_lines: List[str]):
        """Save alert to list - capture full error block"""
        if message_lines:
            first_ora_line = ""
            for line in message_lines:
                if line.startswith('ORA-'):
                    first_ora_line = line
                    break
            
            # Join with newline to keep full block context
            full_text = '\n'.join(message_lines)
            
            alert = AlertError(timestamp, first_ora_line, error_code=first_ora_line, full_text=full_text)
            self.alerts.append(alert)
    
    def get_data(self) -> Dict[str, Any]:
        """Return parsed alerts as dictionary"""
        return {
            "alerts": [alert.to_dict() for alert in self.alerts],
            "count": len(self.alerts),
            "num_days": self.num_days,
            "instance_name": self.instance_name
        }
