"""
Configuration settings for Oracle RAC Report Generator
"""
import os
import sys
from pathlib import Path

# Detect if running as a bundled executable (PyInstaller)
if getattr(sys, 'frozen', False):
    # Path of the .exe file
    EXE_DIR = Path(sys.executable).resolve().parent
    # For internal assets/logs, we might still need the temp dir (sys._MEIPASS)
    BUNDLE_DIR = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
    
    # Use EXE_DIR for output and user-facing files
    BASE_DIR = EXE_DIR
    OUTPUT_DIR = BASE_DIR / "output"
    APPENDIX_OUTPUT_DIR = OUTPUT_DIR / "appendix"
    REPORT_OUTPUT_DIR = OUTPUT_DIR / "report"
    COLLECT_TOOL_DIR = BASE_DIR / "HC_collect_tool"
    TEMPLATE_DIR = BUNDLE_DIR / "template_docx"
else:
    # Running as a regular script
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR.parent / "output"
    APPENDIX_OUTPUT_DIR = OUTPUT_DIR / "appendix"
    REPORT_OUTPUT_DIR = OUTPUT_DIR / "report"
    COLLECT_TOOL_DIR = BASE_DIR.parent / "HC_collect_tool"
    TEMPLATE_DIR = BASE_DIR.parent / "template_docx"

# Logs and DB in the same directory as EXE for portability
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(APPENDIX_OUTPUT_DIR, exist_ok=True)
os.makedirs(REPORT_OUTPUT_DIR, exist_ok=True)
os.makedirs(COLLECT_TOOL_DIR, exist_ok=True)

# Database
DB_PATH = BASE_DIR / "data" / "reports.db"
os.makedirs(DB_PATH.parent, exist_ok=True)


# Logging
LOG_DIR = BASE_DIR / "logs"
os.makedirs(LOG_DIR, exist_ok=True)

LOG_FILE = LOG_DIR / "app.log"

# Application Settings
APP_NAME = "Oracle HC Generator"
APP_VERSION = "2.5.2"
WINDOW_WIDTH = 1440
WINDOW_HEIGHT = 900

# Parser Settings
NUM_DAYS_ALERT = 30  # Default: last 30 days
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100 MB

# Database Settings
DATABASE_URL = f"sqlite:///{DB_PATH}"

# Report Templates
REPORT_LOGO = None  # Optional: path to logo image
REPORT_COMPANY = "Your Company"

# Parser filters
ALERT_LOG_PATTERN = "alert_*.log"
AWR_REPORT_PATTERN = "awrrpt_*.html"

# AWR Table Filters
AWR_TABLES_12C = [
    "This table displays host information",
    "This table displays instance efficiency percentages",
    "This table displays top 10 wait events by total wait time",
    "Top 10 Foreground Events by Total Wait Time",
    "top 10 foreground events",
    "this table displays wait events",
    "This table displays wait class statistics ordered by total wait time",
    "This table displays top SQL by elapsed time",
    "This table displays the text of the SQL statements"
]

AWR_TABLES_11G = [
    "This table displays host information",
    "This table displays instance efficiency percentages",
    "Top 5 Timed Foreground Events",
    "Top 10 Foreground Events by Total Wait Time",
    "top 10 foreground events",
    "timed foreground events",
    "This table displays foreground wait class statistics",
    "This table displays top SQL by elapsed time",
    "This table displays the text of the SQL statements"
]

# Datetime Formats
DATETIME_FORMAT_12C = "%Y-%m-%dT%H:%M:%S"  # 2025-12-01T10:30:45
DATETIME_FORMAT_11G = "%a %b %d %H:%M:%S %Y"  # Mon Dec 01 10:30:45 2025

# Node Names
NODE_1_NAME = "NODE 1"
NODE_2_NAME = "NODE 2"

# Report Fields
REPORT_FIELDS = {
    "db_name": "Database Name",
    "instance_node1": "Instance Node 1",
    "instance_node2": "Instance Node 2",
    "backup_policy": "Backup Policy",
}

# GitHub Sync
GITHUB_TOOLS_API_URL = "https://api.github.com/repos/lethang1752/github_work/contents/HC_collect_tool?ref=master"
AUTO_SYNC_TOOLS = True
