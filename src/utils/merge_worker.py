"""
merge_worker.py
===============
QThread worker that runs merge_docx_reports() in the background so the
GUI stays responsive during the (potentially slow) file-merge operation.
"""

import logging
from typing import List

from PyQt5.QtCore import QThread, pyqtSignal

logger = logging.getLogger(__name__)


class MergeWorker(QThread):
    """
    Background thread for merging .docx files.

    Signals
    -------
    progress(int, str)   – emitted at each step: (percent 0-100, message)
    finished(bool, str)  – emitted when done: (success, status message)
    """

    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)

    def __init__(self, ordered_paths: List[str], output_path: str):
        super().__init__()
        self.ordered_paths = ordered_paths
        self.output_path   = output_path

    def run(self):
        from .docx_merger import merge_docx_reports

        def _on_progress(pct: int, msg: str):
            self.progress.emit(pct, msg)

        try:
            success, message = merge_docx_reports(
                self.ordered_paths,
                self.output_path,
                progress_callback=_on_progress,
            )
            self.finished.emit(success, message)
        except Exception as exc:
            logger.exception("MergeWorker unhandled error")
            self.finished.emit(False, f"Unexpected error: {exc}")
