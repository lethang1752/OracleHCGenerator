"""
report_worker.py
================
QThread worker that runs FinalReportGenerator in the background.
"""

import logging
from typing import Dict, Any
from PyQt5.QtCore import QThread, pyqtSignal
from ..generators.final_report_generator import FinalReportGenerator

logger = logging.getLogger(__name__)

class ReportWorker(QThread):
    """
    Background thread for generating the final summary report.

    Signals
    -------
    progress(str, int)  – (message, percent 0-100)
    finished(bool, str, str) – (success, docx_path, filename)
    """

    progress = pyqtSignal(str, int)
    finished = pyqtSignal(bool, str, str)

    def __init__(self, parsed_data: Dict[str, Any], output_path: str, font_option: str, filename: str, language: str, db_name: str = None):
        super().__init__()
        self.parsed_data = parsed_data
        self.output_path = output_path
        self.font_option = font_option
        self.filename = filename
        self.language = language
        self.db_name = db_name

    def run(self):
        try:
            self.progress.emit("Initializing Final Report generator...", 10)
            
            gen = FinalReportGenerator(
                self.output_path, 
                font_option=self.font_option,
                language=self.language
            )
            
            self.progress.emit("Analyzing Node 1 data items...", 50)
            
            success = gen.generate(self.parsed_data, db_name=self.db_name)
            
            if success:
                self.progress.emit("Report successfully saved!", 100)
                self.finished.emit(True, self.output_path, self.filename)
            else:
                self.finished.emit(False, "", "")
                
        except Exception as e:
            logger.exception("ReportWorker error")
            self.finished.emit(False, str(e), "")
