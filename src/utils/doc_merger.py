"""
doc_merger.py
=============
Merges multiple .docx files in a user-specified order using Microsoft Word
automation (pywin32). Ensures continuous multi-level list numbering across
all merged documents so that section headings continue seamlessly.

Usage (from your service layer):
    from src.utils.doc_merger import merge_documents_ordered

    success, message = merge_documents_ordered(
        ordered_file_paths=[r"C:\docs\file_A.docx", r"C:\docs\file_B.docx"],
        output_destination=r"C:\output\merged_report.docx"
    )
"""

import os
import logging
from typing import List, Tuple

logger = logging.getLogger(__name__)

# Word automation constants (mirrors win32com Word enums)
_WD_DISPLAY_ALERTS_NONE = 0
_WD_INSERT_FILE_VISIBLE_IN_TASK_PANE = 0


def _continue_all_lists(doc) -> None:
    """
    Scan every paragraph in the document. For any paragraph that belongs
    to a numbered/multi-level list, set ContinuePreviousList = True so that
    numbering flows uninterrupted across the seam between inserted files.

    Why:  InsertFile physically appends content but does NOT recalculate list
          numbering context. Without this step a new file always resets its
          own lists back to "1." even when we want them to carry on from the
          previous file's final number.
    """
    for para in doc.Paragraphs:
        try:
            # ListFormat.CountNumberedItems > 0 means the paragraph is in a list
            lf = para.Range.ListFormat
            if lf.CountNumberedItems > 0:
                lf.ApplyListTemplate(
                    ListTemplate=lf.ListTemplate,   # keep the existing template
                    ContinuePreviousList=True,       # ← the magic flag
                    DefaultListBehavior=1            # wdWord10ListBehavior
                )
        except Exception:
            # Skip paragraphs that cannot be reformatted (images, empty, etc.)
            pass


def merge_documents_ordered(
    ordered_file_paths: List[str],
    output_destination: str
) -> Tuple[bool, str]:
    """
    Merge a list of .docx files (in the given order) into a single document.

    Parameters
    ----------
    ordered_file_paths : list of str
        Absolute paths to .docx files, already sorted to the user's preference.
    output_destination : str
        Absolute path where the merged document will be saved.

    Returns
    -------
    (success: bool, message: str)
        success  – True if the file was saved without error.
        message  – Human-readable status or error description.
    """
    if not ordered_file_paths:
        return False, "No files provided to merge."

    # Validate all paths exist before starting Word
    missing = [p for p in ordered_file_paths if not os.path.isfile(p)]
    if missing:
        return False, f"Files not found: {missing}"

    word_app = None
    master_doc = None

    try:
        import win32com.client as win32
    except ImportError as exc:
        return False, (
            f"pywin32 import failed: {exc}\n"
            "Please run: pip install pywin32  and then:\n"
            "python Scripts/pywin32_postinstall.py -install"
        )

    try:
        logger.info("Starting Word automation for document merge...")

        # --- 1. Initialise a hidden Word instance ---
        # Use Dispatch (not gencache.EnsureDispatch) for maximum compatibility
        word_app = win32.Dispatch("Word.Application")
        word_app.Visible = False                    # keep Word hidden
        word_app.DisplayAlerts = _WD_DISPLAY_ALERTS_NONE  # suppress pop-ups
        word_app.ScreenUpdating = False             # skip redraws → faster

        # --- 2. Open the FIRST file as the master document ---
        first_path = os.path.abspath(ordered_file_paths[0])
        logger.info(f"Opening master document: {first_path}")
        master_doc = word_app.Documents.Open(first_path, ReadOnly=False)

        # Move the insertion cursor to the very end of the document
        insert_range = master_doc.Content
        insert_range.Collapse(Direction=0)  # wdCollapseEnd = 0

        # --- 3. Append every subsequent file ---
        for file_path in ordered_file_paths[1:]:
            abs_path = os.path.abspath(file_path)
            logger.info(f"Inserting: {abs_path}")

            # InsertFile appends the content at the current range position.
            # It preserves the source file's styles, images, and tables.
            insert_range.InsertFile(
                FileName=abs_path,
                ConfirmConversions=False,
                Link=False,
                Attachment=False
            )

            # After insertion the range has expanded; collapse back to the end
            # so the next file is appended after this one (not inside it).
            insert_range = master_doc.Content
            insert_range.Collapse(Direction=0)

        # --- 4. Synchronise all list numbering across the merged content ---
        logger.info("Synchronising multi-level list numbering...")
        _continue_all_lists(master_doc)

        # --- 5. Save the merged document to the destination path ---
        abs_dest = os.path.abspath(output_destination)
        os.makedirs(os.path.dirname(abs_dest), exist_ok=True)

        master_doc.SaveAs2(abs_dest, FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
        logger.info(f"Merged document saved to: {abs_dest}")

        return True, f"Merge successful. Saved to: {abs_dest}"

    except Exception as exc:
        logger.exception("Document merge failed.")
        return False, f"Merge failed: {exc}"

    finally:
        # --- 6. Always clean up to avoid 'ghost' WINWORD.EXE processes ---
        try:
            if master_doc is not None:
                master_doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word_app is not None:
                word_app.Quit()
        except Exception:
            pass
        logger.info("Word instance closed.")
