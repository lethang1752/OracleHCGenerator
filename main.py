#!/usr/bin/env python
"""
Oracle RAC Report Generator - Desktop Application
Main Entry Point
"""
import sys
import logging
from pathlib import Path

from PyQt5.QtWidgets import QApplication

# Add src to path
sys.path.insert(0, str(Path(__file__).parent))

from src.utils import setup_logger
from src.ui import MainWindow
from src.config import APP_NAME, APP_VERSION

logger = setup_logger('main')


import multiprocessing

def main():
    """Main application entry point"""
    # Link code to handle multiprocessing in frozen executables
    multiprocessing.freeze_support()
    
    logger.info(f"Starting {APP_NAME} v{APP_VERSION}")
    
    # Fix for Windows Taskbar Icon (AppUserModelID)
    try:
        import ctypes
        myappid = f"victorle.oracle.reportgen.{APP_VERSION}"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass
    
    app = QApplication(sys.argv)
    
    # Set application info
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    
    # Create and show main window
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
