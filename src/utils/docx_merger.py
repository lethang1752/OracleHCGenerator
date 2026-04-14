"""
docx_merger.py
==============
Merges multiple .docx report files into a single output using python-docx.

KEY FEATURE – Heading Renumbering:
  Every source file starts its top-level sections at "1." (e.g. "1. Database",
  "1.1 Status Check", "1.3.7 Disk group"). When merging, this module
  automatically increments the first-level number so sections stay sequential:

      File 0 (first)  → unchanged:   1. Database …   1.1 Status …
      File 1 (second) → renumbered:  2. Database …   2.1 Status …
      File 2 (third)  → renumbered:  3. Database …   3.1 Status …

KEY FEATURE – Image Transfer:
  Images in .docx files are stored as binary blobs in word/media/ and are
  referenced by relationship IDs (r:embed). A plain XML deep-copy keeps the
  old rIds but does NOT transfer the actual image data, resulting in broken
  images. This module:
    1. Finds all image references (DrawingML a:blip + legacy VML v:imagedata).
    2. Copies the binary blob from the source document's part to the master.
    3. Creates a new relationship in the master document's package.
    4. Rewrites the r:embed / r:id attribute with the new rId.
"""

import re
import copy
import posixpath
import logging
from pathlib import Path
from typing import Callable, List, Optional, Tuple

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.part import Part
from docx.opc.packuri import PackURI

logger = logging.getLogger(__name__)

# ─── Namespaces ───────────────────────────────────────────────────────────────
_NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_NS_V = 'urn:schemas-microsoft-com:vml'

# ─── Heading regex ────────────────────────────────────────────────────────────
# Matches the dotted-number prefix at the START of a heading / section title.
#   Group 1 → top-level number         e.g. "1"
#   Group 2 → remaining dotted path    e.g. ".2.3"  (may be empty)
#   Group 3 → separator                e.g. ". "  or  " "
#
# Matched examples:
#   "1. Database"        → ('1', '',      '. ')
#   "1.1 Status Check"   → ('1', '.1',    ' ')
#   "1.3.7 Disk group"   → ('1', '.3.7',  ' ')
_HEADING_RE = re.compile(r'^(\d+)((?:\.\d+)*)(\.?\s)')


# ════════════════════════════════════════════════════════════════════════════
# IMAGE TRANSFER
# ════════════════════════════════════════════════════════════════════════════

def _transfer_image_part(source_doc, master_doc, old_rid: str) -> Optional[str]:
    """
    Copy one embedded binary part (image, chart, …) from *source_doc* to
    *master_doc* and return the new rId string in *master_doc*.

    Returns None if the relationship is external, missing, or transfer fails.
    """
    source_rel = source_doc.part.rels.get(old_rid)
    if not source_rel or source_rel.is_external:
        return None

    source_part = source_rel.target_part

    # Build a unique URI so the image does not collide with existing media
    base_name = posixpath.basename(str(source_part.partname))
    stem, ext  = posixpath.splitext(base_name)

    existing_uris = {str(p.partname) for p in master_doc.part.package.iter_parts()}
    candidate = f'/word/media/{base_name}'
    counter   = 1
    while candidate in existing_uris:
        candidate = f'/word/media/{stem}_{counter}{ext}'
        counter += 1

    new_part = Part(
        PackURI(candidate),
        source_part.content_type,
        source_part.blob,
        master_doc.part.package,
    )

    # Register the new part and its relationship in the master document
    new_rid = master_doc.part.relate_to(new_part, source_rel.reltype)
    logger.debug("Image transferred: %s  rId %s → %s", base_name, old_rid, new_rid)
    return new_rid


def _copy_images_to_master(xml_elem, source_doc, master_doc) -> None:
    """
    Scan an XML element (paragraph or table) for all embedded image
    references and transfer each image binary to *master_doc*, then
    rewrite the reference attribute so the copied element points to the
    correct image in the master document.

    Handles:
      • Modern DrawingML:  <a:blip r:embed="rIdN"/>
      • Legacy VML:        <v:imagedata r:id="rIdN"/>
    """
    r_embed = f'{{{_NS_R}}}embed'
    r_id    = f'{{{_NS_R}}}id'

    # ── DrawingML blip references ────────────────────────────────────────────
    for blip in xml_elem.iter(f'{{{_NS_A}}}blip'):
        old = blip.get(r_embed)
        if old:
            new = _transfer_image_part(source_doc, master_doc, old)
            if new:
                blip.set(r_embed, new)

    # ── Legacy VML imagedata references ─────────────────────────────────────
    for imgdata in xml_elem.iter(f'{{{_NS_V}}}imagedata'):
        old = imgdata.get(r_id)
        if old:
            new = _transfer_image_part(source_doc, master_doc, old)
            if new:
                imgdata.set(r_id, new)


# ════════════════════════════════════════════════════════════════════════════
# HEADING RENUMBERING
# ════════════════════════════════════════════════════════════════════════════

def _replace_heading_prefix(text: str, file_index: int) -> str:
    """
    Increment the first-level number of a dotted heading by *file_index*.

    file_index 0 → first file; no change.
    file_index 1 → "1.x.y …" becomes "2.x.y …".
    file_index 2 → "1.x.y …" becomes "3.x.y …".
    """
    if file_index == 0:
        return text

    def _sub(m: re.Match) -> str:
        new_first = int(m.group(1)) + file_index
        return f"{new_first}{m.group(2)}{m.group(3)}"

    return _HEADING_RE.sub(_sub, text, count=1)


def _get_para_text(xml_elem) -> str:
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    return ''.join((t.text or '') for t in xml_elem.findall(f'.//{{{NS}}}t'))


def _rewrite_heading_text(xml_elem, file_index: int) -> None:
    """
    Rewrite the numbered prefix inside a paragraph's <w:t> nodes.
    Puts the full updated text into the FIRST <w:t>, blanks the rest.
    This preserves all run-level formatting (bold, colour, font size).
    """
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    t_nodes  = xml_elem.findall(f'.//{{{NS}}}t')
    if not t_nodes:
        return

    full_text = ''.join((t.text or '') for t in t_nodes)
    new_text  = _replace_heading_prefix(full_text, file_index)

    if new_text != full_text:
        t_nodes[0].text = new_text
        for t in t_nodes[1:]:
            t.text = ''
        logger.debug("Renumbered: '%s' → '%s'", full_text, new_text)


def _is_heading_or_numbered(xml_elem) -> bool:
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    pStyle = xml_elem.find(f'.//{{{NS}}}pStyle')
    if pStyle is not None:
        val = (pStyle.get(f'{{{NS}}}val') or '').lower()
        if 'heading' in val:
            return True
    return bool(_HEADING_RE.match(_get_para_text(xml_elem).lstrip()))


# ════════════════════════════════════════════════════════════════════════════
# PAGE BREAK HELPER
# ════════════════════════════════════════════════════════════════════════════

def _make_page_break():
    p  = OxmlElement('w:p')
    r  = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br)
    p.append(r)
    return p


# ════════════════════════════════════════════════════════════════════════════
# MAIN MERGE FUNCTION
# ════════════════════════════════════════════════════════════════════════════

def merge_docx_reports(
    ordered_file_paths: List[str],
    output_path: str,
    progress_callback: Optional[Callable[[int, str], None]] = None,
) -> Tuple[bool, str]:
    """
    Merge an ordered list of .docx files into a single output document.

    • All paragraph formatting is preserved via lxml deep-copy.
    • Images are transferred (binary + relationship) so they appear correctly.
    • Tables keep column widths, shading, and borders.
    • Numbered heading prefixes are auto-incremented per file.
    • A page break is inserted between files.
    • Trailing <w:sectPr> of intermediate files is dropped.

    Parameters
    ----------
    ordered_file_paths : list of absolute path strings (merge order).
    output_path        : destination path for the merged .docx file.
    progress_callback  : optional callable(percent: int, message: str).

    Returns
    -------
    (success: bool, message: str)
    """
    if not ordered_file_paths:
        return False, "No files provided."

    missing = [p for p in ordered_file_paths if not Path(p).is_file()]
    if missing:
        return False, "Files not found:\n" + "\n".join(missing)

    def _prog(pct: int, msg: str) -> None:
        if progress_callback:
            progress_callback(pct, msg)
        logger.info("[%d%%] %s", pct, msg)

    try:
        total = len(ordered_file_paths)
        _prog(5, f"Opening base document: {Path(ordered_file_paths[0]).name}")

        master      = Document(ordered_file_paths[0])
        master_body = master.element.body

        for file_index, file_path in enumerate(ordered_file_paths[1:], start=1):
            pct = 10 + int(file_index / total * 80)
            _prog(pct, f"[{file_index + 1}/{total}] Appending: {Path(file_path).name}")

            source = Document(file_path)

            # Separate each report with a page break
            master_body.append(_make_page_break())

            for child in source.element.body:
                # Drop trailing section properties of intermediate docs
                if child.tag == qn('w:sectPr'):
                    continue

                # Deep-copy preserves all XML: fonts, shading, widths, etc.
                new_elem = copy.deepcopy(child)

                # ── Transfer images BEFORE appending ────────────────────
                # Without this, r:embed attributes point at rIds that don't
                # exist in the master document → images appear as red X.
                _copy_images_to_master(new_elem, source, master)

                # ── Renumber headings ────────────────────────────────────
                if child.tag == qn('w:p') and _is_heading_or_numbered(new_elem):
                    _rewrite_heading_text(new_elem, file_index)

                master_body.append(new_elem)

        _prog(92, "Saving merged document...")
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        master.save(str(out))

        _prog(100, f"Saved → {out}")
        return True, f"Merge successful!\nSaved to: {out}"

    except Exception as exc:
        logger.exception("merge_docx_reports failed")
        return False, f"Merge error: {exc}"
