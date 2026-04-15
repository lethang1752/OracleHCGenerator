"""
generator_worker.py
===================
QThread worker that runs ComprehensiveHealthcareReportGenerator in the background
so the GUI stays responsive during the potentially slow file-generation operation.
"""

import logging
from typing import Dict, Any
from pathlib import Path

from PyQt5.QtCore import QThread, pyqtSignal
from ..generators.comprehensive_report_generator import ComprehensiveHealthcareReportGenerator

logger = logging.getLogger(__name__)


class GeneratorWorker(QThread):
    """
    Background thread for generating the comprehensive .docx report.

    Signals
    -------
    progress(str, int)  – emitted at each step: (message, percent 0-100)
    finished(bool, str, str) – emitted when done: (success, docx_path, filename)
    """

    progress = pyqtSignal(str, int)
    finished = pyqtSignal(bool, str, str)

    def __init__(self, parsed_data: Dict[str, Any], output_path: str, font_option: str, filename: str):
        super().__init__()
        self.parsed_data = parsed_data
        self.output_path = output_path
        self.font_option = font_option
        self.filename = filename

    def run(self):
        try:
            self.progress.emit("Initializing report generator...", 92)
            
            gen = ComprehensiveHealthcareReportGenerator(
                self.output_path, 
                font_option=self.font_option
            )
            
            # We don't have a direct callback into generate_from_parsed_data currently
            # but we can wrap the main call. If the generator is updated to support
            # progress callbacks, we should pass them here.
            
            self.progress.emit(f"Generating items for {len(self.parsed_data.get('nodes', []))} nodes...", 95)
            
            success = gen.generate_from_parsed_data(self.parsed_data)
            
            if success:
                self.progress.emit("Finalizing and saving document...", 100)
                self.finished.emit(True, self.output_path, self.filename)
            else:
                self.finished.emit(False, "", "")
                
        except Exception as e:
            logger.exception("GeneratorWorker unhandled error")
            self.finished.emit(False, str(e), "")
