"""
Base Parser - Abstract base class for all parsers
"""
from abc import ABC, abstractmethod
from typing import Any, Dict, List
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


class BaseParser(ABC):
    """Abstract base class for data parsers"""
    
    def __init__(self, log_dir: str):
        """
        Initialize parser
        
        Args:
            log_dir: Directory containing log files
        """
        self.log_dir = log_dir
        self.data = {}
        self.errors = []
    
    @abstractmethod
    def parse(self) -> bool:
        """Parse data from files"""
        pass
    
    @abstractmethod
    def get_data(self) -> Dict[str, Any]:
        """Return parsed data"""
        pass
    
    def add_error(self, message: str, exception: Exception = None):
        """Log error"""
        error_msg = message
        if exception:
            error_msg = f"{message}: {str(exception)}"
        self.errors.append(error_msg)
        logger.error(error_msg)
    
    def get_errors(self) -> List[str]:
        """Get all errors"""
        return self.errors
    
    def has_errors(self) -> bool:
        """Check if has errors"""
        return len(self.errors) > 0
    
    def clear_errors(self):
        """Clear errors"""
        self.errors = []
