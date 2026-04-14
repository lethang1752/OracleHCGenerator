"""
Helper functions
"""
from pathlib import Path
from typing import List


def get_log_files(directory: str) -> List[str]:
    """Get all log files from directory"""
    path = Path(directory)
    return [str(f) for f in path.glob("*.log")]


def get_html_files(directory: str) -> List[str]:
    """Get all HTML files from directory"""
    path = Path(directory)
    return [str(f) for f in path.glob("*.html")]


def format_file_size(bytes_size: int) -> str:
    """Format bytes to human readable format"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if bytes_size < 1024:
            return f"{bytes_size:.1f} {unit}"
        bytes_size /= 1024
    return f"{bytes_size:.1f} TB"


def sanitize_filename(filename: str) -> str:
    """Remove invalid characters from filename"""
    invalid_chars = r'<>:"/\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename
