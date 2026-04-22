#!/usr/bin/env python3
#header_footer_fixed.py
"""
Header/Footer Formatter + Blank-Line Cleanup (GUI) with Optional Backups + PDF Run Report
+ Working-copy editing + Force-editability + Step logging + Heartbeat + Watchdog timeout
+ SUBSTEP diagnostics + Optional Shapes/TextBoxes processing

Enhancements added (Jan 2026):
- Optional per-document PDF export (Word ExportAsFixedFormat) via GUI checkbox
- Optional output naming by Document ID read from header table via GUI checkbox
  Example: "Document ID" -> "FRM-176" => outputs FRM-176.docx and FRM-176.pdf
- DOCX ZIP/XML post-processing remains present but is DISABLED by default to avoid
  Word "found unreadable content" repair prompts (COM-only output is canonical).

Fix implemented previously:
- Replace iterative paragraph-by-paragraph blank deletion after tables (hang-prone)
  with a single-range delete that is story-correct for headers/footers.

Also:
- When Show Word (debug)=YES, open docs Visible=True and Activate() to avoid "empty shell" window.
"""

from __future__ import annotations

import os
import re
import time
import stat
import shutil
import socket
import getpass
import platform
import traceback
import threading
import queue
import subprocess
import zipfile
from io import BytesIO
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pythoncom
    import win32com.client as win32
    try:
        import win32process
    except Exception:
        win32process = None
except Exception:
    win32 = None
    pythoncom = None
    win32process = None

try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Preformatted
    )
    from reportlab.pdfgen import canvas
except Exception:
    LETTER = None
    inch = None
    colors = None
    getSampleStyleSheet = None
    ParagraphStyle = None
    SimpleDocTemplate = None
    Paragraph = None
    Spacer = None
    Table = None
    TableStyle = None
    PageBreak = None
    Preformatted = None
    canvas = None


# -----------------------------
# Word constants
# -----------------------------
wdHeaderFooterPrimary = 1
wdHeaderFooterFirstPage = 2
wdHeaderFooterEvenPages = 3

wdAlignParagraphCenter = 1
wdLineSpaceSingle = 0
wdAlertsNone = 0
wdParagraph = 4
wdAlignTableCenter = 1
wdAlignRowCenter = 1

msoAutomationSecurityForceDisable = 3

# ExportAsFixedFormat
wdExportFormatPDF = 17

HF_KINDS = [
    (wdHeaderFooterPrimary, "PRIMARY"),
    (wdHeaderFooterFirstPage, "FIRSTPAGE"),
    (wdHeaderFooterEvenPages, "EVENPAGES"),
]

_WS_RE = re.compile(r"\s+", re.UNICODE)


# -----------------------------
# Stats
# -----------------------------
@dataclass
class RunStats:
    run_id: str
    start_ts: float
    start_dt: datetime
    end_dt: datetime | None = None

    total_files: int = 0
    processed: int = 0
    succeeded: int = 0
    failed: int = 0
    skipped: int = 0

    total_actions: int = 0

    header_post_table_deleted_total: int = 0
    footer_post_table_deleted_total: int = 0
    tables_centered_total: int = 0
    text_shapes_formatted_total: int = 0
    format_apps_total: int = 0

    backups_enabled: bool = False
    backups_attempted: int = 0
    backups_ok: int = 0
    backups_failed: int = 0

    # NEW: per-document PDF export
    pdf_exports_enabled: bool = False
    pdf_exports_attempted: int = 0
    pdf_exports_ok: int = 0
    pdf_exports_failed: int = 0

    # NEW: naming by Document ID
    docid_naming_enabled: bool = False
    docid_found_count: int = 0
    docid_used_count: int = 0
    docid_missing_count: int = 0

    user: str = ""
    computer: str = ""
    os_info: str = ""
    python_info: str = ""
    word_version: str = "UNKNOWN"
    tz_name: str = ""
    tz_offset: str = ""


# -----------------------------
# Helpers
# -----------------------------
def make_run_id() -> str:
    return datetime.now().strftime("%d%m%Y_%H%M%S")


def ts_line() -> str:
    return datetime.now().strftime("[%H:%M:%S]")


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def is_word_file(filename: str) -> bool:
    fn = filename.lower()
    if fn.startswith("~$"):
        return False
    return fn.endswith(".docx") or fn.endswith(".docm") or fn.endswith(".doc")


def iter_word_files(input_root: str):
    for root, _, files in os.walk(input_root):
        for f in files:
            if not is_word_file(f):
                continue
            abs_path = os.path.join(root, f)
            rel_path = os.path.relpath(abs_path, input_root)
            yield abs_path, rel_path


def safe_strip_word_text(s: str) -> str:
    if s is None:
        return ""
    # Word artifacts and invisible whitespace variants
    s = (
        s.replace("\r", "")      # paragraph mark
         .replace("\x07", "")    # end-of-cell marker
         .replace("\xa0", " ")   # NBSP
         .replace("\u200b", "")  # zero-width space
         .replace("\v", "")      # vertical tab
         .replace("\f", "")      # form feed
         .replace("\t", "")      # tabs
    )
    s = _WS_RE.sub("", s)
    return s


def is_blank_paragraph(paragraph) -> bool:
    try:
        txt = paragraph.Range.Text
    except Exception:
        return False
    return safe_strip_word_text(txt) == ""


def apply_paragraph_format_to_range(rng) -> int:
    pf = rng.ParagraphFormat
    pf.SpaceBefore = 0
    pf.SpaceAfter = 0
    try:
        pf.SpaceBeforeAuto = False
    except Exception:
        pass
    try:
        pf.SpaceAfterAuto = False
    except Exception:
        pass
    pf.LineSpacingRule = wdLineSpaceSingle
    pf.Alignment = wdAlignParagraphCenter
    return 1


def apply_paragraph_format_to_shapes_in_headerfooter(hf) -> tuple[int, int, int]:
    shapes_count = 0
    text_shapes = 0
    formatted = 0
    try:
        shapes = hf.Shapes
        shapes_count = int(shapes.Count)
        for i in range(1, shapes.Count + 1):
            shp = shapes(i)
            try:
                has_text = bool(shp.TextFrame.HasText)
            except Exception:
                has_text = False
            if not has_text:
                continue
            text_shapes += 1
            try:
                tr = shp.TextFrame.TextRange
                formatted += apply_paragraph_format_to_range(tr)
            except Exception:
                continue
    except Exception:
        shapes_count = 0
        text_shapes = 0
        formatted = 0
    return shapes_count, text_shapes, formatted


def center_tables_in_range(rng) -> tuple[int, int]:
    tables_found = 0
    centered = 0
    try:
        tables_found = int(rng.Tables.Count)
        for i in range(1, tables_found + 1):
            tbl = rng.Tables(i)
            try:
                tbl.Alignment = wdAlignTableCenter
            except Exception:
                pass
            try:
                tbl.Rows.Alignment = wdAlignRowCenter
            except Exception:
                pass
            try:
                apply_paragraph_format_to_range(tbl.Range)
            except Exception:
                pass
            centered += 1
    except Exception:
        return tables_found, centered
    return tables_found, centered


def range_is_empty(hf_range, table_count: int, text_shapes_count: int) -> bool:
    if table_count > 0 or text_shapes_count > 0:
        return False
    try:
        txt = hf_range.Text or ""
    except Exception:
        return True
    return safe_strip_word_text(txt) == ""


def clear_readonly_attribute(path: str) -> None:
    try:
        os.chmod(path, stat.S_IWRITE)
    except Exception:
        pass
    try:
        subprocess.run(["attrib", "-R", path], capture_output=True, text=True, check=False)
    except Exception:
        pass


def make_working_copy(abs_in: str, work_abs: str) -> None:
    ensure_dir(os.path.dirname(work_abs))
    shutil.copy2(abs_in, work_abs)
    clear_readonly_attribute(work_abs)


def force_document_editable(doc) -> dict:
    result = {
        "doc_readonly_initial": "UNKNOWN",
        "unprotect_success": "NO",
        "final_cleared_success": "NO",
        "readonly_recommended_cleared_success": "NO",
        "changefileaccess_success": "NO",
        "doc_readonly_final": "UNKNOWN",
    }

    try:
        result["doc_readonly_initial"] = "YES" if bool(doc.ReadOnly) else "NO"
    except Exception:
        pass

    try:
        doc.Unprotect("")
        result["unprotect_success"] = "YES"
    except Exception:
        result["unprotect_success"] = "NO"

    try:
        doc.Final = False
        result["final_cleared_success"] = "YES"
    except Exception:
        result["final_cleared_success"] = "NO"

    try:
        doc.ReadOnlyRecommended = False
        result["readonly_recommended_cleared_success"] = "YES"
    except Exception:
        result["readonly_recommended_cleared_success"] = "NO"

    try:
        doc.ChangeFileAccess(2)  # wdReadWrite (best-effort)
        result["changefileaccess_success"] = "YES"
    except Exception:
        result["changefileaccess_success"] = "NO"

    try:
        result["doc_readonly_final"] = "YES" if bool(doc.ReadOnly) else "NO"
    except Exception:
        pass

    return result


# -----------------------------
# NEW: Filename + Document ID helpers
# -----------------------------
_FILENAME_BAD_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1F]+')
_FILENAME_DOTS_SPACES = re.compile(r"[\. ]+$")

def sanitize_filename_component(name: str, *, max_len: int = 120) -> str:
    """
    Produce a Windows-safe filename base (no extension).
    Keeps letters, digits, spaces, hyphen, underscore, dot; collapses whitespace.
    """
    if not name:
        return ""
    s = name.strip()
    s = s.replace("\r", " ").replace("\n", " ")
    s = s.replace("\x07", " ")
    s = _WS_RE.sub(" ", s).strip()
    s = _FILENAME_BAD_CHARS.sub("_", s)
    s = _FILENAME_DOTS_SPACES.sub("", s)
    s = s.strip()
    if len(s) > max_len:
        s = s[:max_len].rstrip()
    return s

def _normalize_key_text(s: str) -> str:
    s = (s or "").lower()
    s = s.replace("\r", "").replace("\x07", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _cell_text(cell) -> str:
    try:
        return str(cell.Range.Text or "")
    except Exception:
        return ""

def extract_document_id_from_headers(doc) -> str | None:
    """
    Look in section headers (all kinds) for a table cell containing 'Document ID'
    and return the adjacent cell value (e.g., FRM-176). Returns None if not found.
    """
    target = "document id"
    try:
        sec_count = int(doc.Sections.Count)
    except Exception:
        return None

    for sidx in range(1, sec_count + 1):
        sec = doc.Sections(sidx)
        for kind, _kind_name in HF_KINDS:
            try:
                hdr = sec.Headers(kind)
                rng = hdr.Range
                tbl_count = int(rng.Tables.Count)
            except Exception:
                continue

            for t in range(1, tbl_count + 1):
                try:
                    tbl = rng.Tables(t)
                    rows = int(tbl.Rows.Count)
                    cols = int(tbl.Columns.Count)
                except Exception:
                    continue

                # Scan by row/col so "adjacent cell" logic is reliable
                for r in range(1, rows + 1):
                    for c in range(1, cols + 1):
                        try:
                            cell = tbl.Cell(r, c)
                            key_txt = _normalize_key_text(_cell_text(cell))
                        except Exception:
                            continue

                        if target == key_txt or target in key_txt:
                            # Prefer right-adjacent cell if exists
                            val = ""
                            if c < cols:
                                try:
                                    val_cell = tbl.Cell(r, c + 1)
                                    val = _cell_text(val_cell)
                                except Exception:
                                    val = ""
                            # Fallback: next cell in same row by Cells collection (handles odd merges)
                            if not sanitize_filename_component(safe_strip_word_text(val)):
                                try:
                                    row = tbl.Rows(r)
                                    rc = int(row.Cells.Count)
                                    # locate this cell index in row.Cells by comparing start positions
                                    this_start = int(cell.Range.Start)
                                    best_idx = None
                                    for i in range(1, rc + 1):
                                        try:
                                            cc = row.Cells(i)
                                            if int(cc.Range.Start) == this_start:
                                                best_idx = i
                                                break
                                        except Exception:
                                            continue
                                    if best_idx is not None and best_idx < rc:
                                        val = _cell_text(row.Cells(best_idx + 1))
                                except Exception:
                                    pass

                            candidate = val or ""
                            candidate = candidate.replace("\r", "").replace("\x07", "")
                            candidate = re.sub(r"\s+", " ", candidate).strip()
                            candidate = sanitize_filename_component(candidate)

                            if candidate:
                                return candidate

    return None

def pick_unique_base_name(out_dir: str, base: str, extensions: list[str]) -> str:
    """
    Choose a unique base name within out_dir such that all out_dir/base+ext do not already exist.
    If conflicts exist, append _2, _3, ...
    """
    base = sanitize_filename_component(base)
    if not base:
        base = "Document"

    def conflict(b: str) -> bool:
        for ext in extensions:
            p = os.path.join(out_dir, b + ext)
            if os.path.exists(p):
                return True
        return False

    if not conflict(base):
        return base

    k = 2
    while k < 10_000:
        b2 = f"{base}_{k}"
        if not conflict(b2):
            return b2
        k += 1

    # last-resort timestamp
    return f"{base}_{int(time.time())}"

def export_doc_to_pdf(doc, pdf_path: str) -> None:
    """
    Export the active document to PDF using Word COM.
    """
    ensure_dir(os.path.dirname(pdf_path))
    # Use named args for clarity/compatibility
    doc.ExportAsFixedFormat(
        OutputFileName=pdf_path,
        ExportFormat=wdExportFormatPDF,
        OpenAfterExport=False
    )


# -----------------------------
# SAFE blank cleanup (single-range delete)
# -----------------------------
def count_consecutive_blank_paragraphs(paras, max_scan: int = 50) -> int:
    """Count leading blank paragraphs in a Paragraphs collection, up to max_scan."""
    blank = 0
    try:
        n = int(paras.Count)
    except Exception:
        return 0
    lim = min(n, max_scan)
    for i in range(1, lim + 1):
        try:
            p = paras(i)
        except Exception:
            break
        if is_blank_paragraph(p):
            blank += 1
        else:
            break
    return blank


def delete_blank_paragraph_block(paras, start_idx: int, end_idx: int) -> int:
    """
    Delete paragraphs start_idx..end_idx as a single contiguous Range.Delete(),
    using paragraph ranges (story-correct for headers/footers).
    Returns number deleted.
    """
    if end_idx < start_idx:
        return 0
    try:
        p_start = paras(start_idx)
        p_end = paras(end_idx)

        rng_del = p_start.Range.Duplicate
        rng_del.End = p_end.Range.End
        rng_del.Delete()

        return end_idx - start_idx + 1
    except Exception:
        return 0


def _first_paragraph_index_at_or_after(paras, start_pos: int, max_scan: int = 500) -> int | None:
    """
    Find the first paragraph index i such that paras(i).Range.Start >= start_pos.
    Returns 1-based index or None if not found within max_scan.
    """
    try:
        n = int(paras.Count)
    except Exception:
        return None

    lim = min(n, max_scan)
    for i in range(1, lim + 1):
        try:
            p = paras(i)
            p_start = int(p.Range.Start)
        except Exception:
            continue
        if p_start >= start_pos:
            return i
    return None


def enforce_max_blank_paragraphs_after_table_safe(hf_range, table, max_blanks: int) -> tuple[int, int]:
    """
    Robust COM-only enforcement (header/footer story-correct):
      - Build a container range INSIDE the header/footer story via hf_range.Duplicate
      - Find first paragraph at/after table.Range.End
      - Delete excess blanks as a single contiguous range delete
    """
    try:
        container_start = int(table.Range.Start)
        container_end = int(hf_range.End)
        if container_start >= container_end:
            return 0, 0

        container_rng = hf_range.Duplicate
        container_rng.Start = container_start
        container_rng.End = container_end
        paras = container_rng.Paragraphs

        tbl_end = int(table.Range.End)

        first_i = _first_paragraph_index_at_or_after(paras, tbl_end, max_scan=500)
        if first_i is None:
            return 0, 0

        blank_found = 0
        try:
            n = int(paras.Count)
        except Exception:
            n = 0

        lim = min(n, first_i + 50 - 1)
        for i in range(first_i, lim + 1):
            try:
                p = paras(i)
            except Exception:
                break
            if is_blank_paragraph(p):
                blank_found += 1
            else:
                break

        if blank_found <= max_blanks:
            return blank_found, 0

        del_start_i = first_i + max_blanks
        del_end_i = first_i + blank_found - 1

        deleted = delete_blank_paragraph_block(paras, del_start_i, del_end_i)
        return blank_found, deleted

    except Exception:
        return 0, 0


def enforce_max_blank_paragraphs_after_phrase_safe(hf_range, phrase: str, max_blanks: int = 0) -> tuple[int, int]:
    """
    In the header/footer story, locate the first paragraph containing `phrase`,
    then enforce at most `max_blanks` blank paragraphs immediately AFTER it.

    Returns: (blank_found_after_phrase, deleted)
    """
    try:
        phrase_norm = (phrase or "").strip().lower()
        if not phrase_norm:
            return 0, 0

        paras = hf_range.Paragraphs
        try:
            n = int(paras.Count)
        except Exception:
            return 0, 0

        # Find target paragraph index containing phrase
        target_i = None
        for i in range(1, min(n, 500) + 1):
            try:
                p = paras(i)
                t = (p.Range.Text or "")
            except Exception:
                continue
            if phrase_norm in t.lower():
                target_i = i
                break

        if target_i is None:
            return 0, 0

        # Count consecutive blanks after target paragraph
        first_blank_i = target_i + 1
        blank_found = 0
        lim = min(n, first_blank_i + 50 - 1)
        for i in range(first_blank_i, lim + 1):
            try:
                p = paras(i)
            except Exception:
                break
            if is_blank_paragraph(p):
                blank_found += 1
            else:
                break

        if blank_found <= max_blanks:
            return blank_found, 0

        del_start_i = first_blank_i + max_blanks
        del_end_i = first_blank_i + blank_found - 1

        deleted = delete_blank_paragraph_block(paras, del_start_i, del_end_i)
        return blank_found, deleted

    except Exception:
        return 0, 0


# -----------------------------
# DOCX post-save XML cleanup (deterministic) — PRESENT BUT NOT USED BY DEFAULT
# (Left in place; disabled in worker to avoid Word repair prompts.)
# -----------------------------
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_VML_NS = "urn:schemas-microsoft-com:vml"

ET.register_namespace("w", _W_NS)

def _qn(tag: str) -> str:
    return f"{{{_W_NS}}}{tag}"

def _vqn(tag: str) -> str:
    return f"{{{_VML_NS}}}{tag}"

def _p_visible_text(p: ET.Element) -> str:
    parts: list[str] = []
    for t in p.findall(".//" + _qn("t")):
        if t.text:
            parts.append(t.text)
    return "".join(parts).strip()

def _p_has_fields_or_objects(p: ET.Element) -> bool:
    if p.find(".//" + _qn("instrText")) is not None:
        return True
    if p.find(".//" + _qn("fldChar")) is not None:
        return True
    if p.find(".//" + _qn("drawing")) is not None:
        return True
    if p.find(".//" + _qn("pict")) is not None:
        return True
    if p.find(".//" + _qn("object")) is not None:
        return True
    if p.find(".//" + _qn("br")) is not None:
        return True
    if p.find(".//" + _qn("tab")) is not None:
        return True
    return False

def _is_empty_w_p(p: ET.Element) -> bool:
    if _p_visible_text(p):
        return False
    if _p_has_fields_or_objects(p):
        return False
    return True

def _is_watermark_paragraph(p: ET.Element) -> bool:
    if _p_visible_text(p):
        return False

    for shp in p.findall(".//" + _vqn("shape")):
        sid = (shp.attrib.get("id", "") or "")
        sname = (shp.attrib.get("name", "") or "")
        blob = (sid + " " + sname).lower()

        if "powerpluswatermarkobject" in blob or "watermark" in blob or "watermarkobject" in blob:
            return True

        tp = (shp.attrib.get("type", "") or "").lower()
        st = (shp.attrib.get("style", "") or "").lower()

        if tp == "#_x0000_t136":
            if ("mso-position-vertical:center" in st) and ("mso-position-horizontal:center" in st):
                return True

    return False

def _iter_block_containers(root: ET.Element) -> list[ET.Element]:
    containers = [root]
    containers.extend(root.findall(".//" + _qn("sdtContent")))
    containers.extend(root.findall(".//" + _qn("txbxContent")))
    return containers

def _remove_watermarks_in_container(container: ET.Element) -> int:
    removed = 0
    for ch in list(container):
        if ch.tag == _qn("p") and _is_watermark_paragraph(ch):
            container.remove(ch)
            removed += 1
    return removed

def _cleanup_tbl_following_empty_paragraphs(container: ET.Element, keep_after_tbl: int) -> int:
    removed = 0
    children = list(container)
    i = 0
    while i < len(children):
        el = children[i]
        if el.tag == _qn("tbl"):
            j = i + 1
            empties: list[ET.Element] = []
            while j < len(children) and children[j].tag == _qn("p") and _is_empty_w_p(children[j]):
                empties.append(children[j])
                j += 1

            if len(empties) > keep_after_tbl:
                for p_el in empties[keep_after_tbl:]:
                    container.remove(p_el)
                    removed += 1
                children = list(container)
        i += 1
    return removed

def _enforce_footer_adjacency(container: ET.Element, phrase: str) -> int:
    removed = 0
    phrase_norm = (phrase or "").strip().lower()
    if not phrase_norm:
        return 0

    children = list(container)

    target_idx = None
    for i, el in enumerate(children):
        if el.tag != _qn("p"):
            continue
        if phrase_norm in _p_visible_text(el).lower():
            target_idx = i
            break
    if target_idx is None:
        return 0

    last_tbl_idx = None
    for k in range(target_idx - 1, -1, -1):
        if children[k].tag == _qn("tbl"):
            last_tbl_idx = k
            break
    if last_tbl_idx is None:
        return 0

    children = list(container)
    to_remove: list[ET.Element] = []
    for i in range(last_tbl_idx + 1, target_idx):
        if 0 <= i < len(children) and children[i].tag == _qn("p") and _is_empty_w_p(children[i]):
            to_remove.append(children[i])

    for el in reversed(to_remove):
        container.remove(el)
        removed += 1

    return removed

def _cleanup_empty_paragraphs_after_phrase(container: ET.Element, phrase: str, keep_after: int = 0) -> int:
    removed = 0
    phrase_norm = (phrase or "").strip().lower()
    if not phrase_norm:
        return 0

    children = list(container)

    target_idx = None
    for i, el in enumerate(children):
        if el.tag != _qn("p"):
            continue
        if phrase_norm in _p_visible_text(el).lower():
            target_idx = i
            break
    if target_idx is None:
        return 0

    j = target_idx + 1
    empties: list[ET.Element] = []
    while j < len(children) and children[j].tag == _qn("p") and _is_empty_w_p(children[j]):
        empties.append(children[j])
        j += 1

    if len(empties) <= keep_after:
        return 0

    for p_el in empties[keep_after:]:
        container.remove(p_el)
        removed += 1

    return removed

def docx_postprocess_header_footer_spacing(
    docx_path: str,
    keep_header_blanks: int = 1,
    keep_footer_blanks: int = 0,
    confidential_phrase: str = "Neurotech Pharmaceuticals, Inc."
) -> dict:
    lower = docx_path.lower()
    if not (lower.endswith(".docx") or lower.endswith(".docm")):
        return {
            "skipped": True,
            "reason": "not_docx_or_docm",
            "header_watermarks_removed": 0,
            "footer_watermarks_removed": 0,
            "header_empty_p_removed": 0,
            "footer_empty_p_removed": 0,
            "footer_gap_empty_p_removed": 0,
            "footer_trailing_empty_p_removed": 0,
        }

    metrics = {
        "skipped": False,
        "reason": "",
        "header_watermarks_removed": 0,
        "footer_watermarks_removed": 0,
        "header_empty_p_removed": 0,
        "footer_empty_p_removed": 0,
        "footer_gap_empty_p_removed": 0,
        "footer_trailing_empty_p_removed": 0,
    }

    with zipfile.ZipFile(docx_path, "r") as zin:
        file_map = {name: zin.read(name) for name in zin.namelist()}

    for name, data in list(file_map.items()):
        if not (name.startswith("word/header") or name.startswith("word/footer")):
            continue
        if not name.endswith(".xml"):
            continue

        try:
            root = ET.fromstring(data)
        except Exception:
            continue

        if root.tag == _qn("hdr"):
            for container in _iter_block_containers(root):
                metrics["header_watermarks_removed"] += _remove_watermarks_in_container(container)
                metrics["header_empty_p_removed"] += _cleanup_tbl_following_empty_paragraphs(container, keep_after_tbl=keep_header_blanks)
            file_map[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        elif root.tag == _qn("ftr"):
            for container in _iter_block_containers(root):
                metrics["footer_watermarks_removed"] += _remove_watermarks_in_container(container)
                metrics["footer_empty_p_removed"] += _cleanup_tbl_following_empty_paragraphs(container, keep_after_tbl=keep_footer_blanks)
                metrics["footer_gap_empty_p_removed"] += _enforce_footer_adjacency(container, confidential_phrase)

                metrics["footer_trailing_empty_p_removed"] += _cleanup_empty_paragraphs_after_phrase(
                    container, confidential_phrase, keep_after=0
                )

            file_map[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    tmp_buf = BytesIO()
    with zipfile.ZipFile(tmp_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in file_map.items():
            zout.writestr(name, data)

    with open(docx_path, "wb") as f:
        f.write(tmp_buf.getvalue())

    return metrics


# -----------------------------
# Header/footer processing
# -----------------------------
def process_headerfooter(
    hf,
    doc,
    mode: str,
    *,
    substep: Optional[Callable[[str], None]] = None,
    enable_shapes: bool = False,
    enable_table_centering: bool = True,
    enable_blank_cleanup: bool = True,
) -> tuple[dict, int, dict]:
    actions = 0
    agg = {"format_apps": 0, "post_table_deleted": 0, "tables_centered": 0, "text_shapes_formatted": 0}

    res = {
        "present": "YES",
        "tables_found": "NO",
        "table_count": 0,
        "anchor": "N/A",
        "format_applied": "NO",
        "spec_max": 0 if mode == "header" else 0,
        "tables_centering_applied": "NO",
        "tables_centered_count": 0,
        "post_table_blanks_found": "N/A",
        "post_table_blank_count": "N/A",
        "post_table_deletions_required": "N/A",
        "post_table_deleted": 0,
        "post_table_remaining": "N/A",
        "shapes_found": "NO",
        "shapes_count": 0,
        "text_shapes_found": "NO",
        "text_shapes_count": 0,
        "text_shapes_formatted": 0,
        "range_empty": "YES",
        "blank_cleanup_method": "SAFE_RANGE_DELETE",
    }

    try:
        if substep:
            substep("range_format")
        actions += apply_paragraph_format_to_range(hf.Range)
        agg["format_apps"] += 1
        res["format_applied"] = "YES"

        if substep:
            substep("tables_count")
        try:
            tbl_count = int(hf.Range.Tables.Count)
        except Exception:
            tbl_count = 0

        res["table_count"] = tbl_count
        res["tables_found"] = "YES" if tbl_count > 0 else "NO"
        res["anchor"] = "LAST_TABLE" if tbl_count > 0 else "N/A"

        if enable_shapes:
            if substep:
                substep("shapes")
            shapes_count, text_shapes_count, text_shapes_formatted = apply_paragraph_format_to_shapes_in_headerfooter(hf)
            res["shapes_count"] = shapes_count
            res["shapes_found"] = "YES" if shapes_count > 0 else "NO"
            res["text_shapes_count"] = text_shapes_count
            res["text_shapes_found"] = "YES" if text_shapes_count > 0 else "NO"
            res["text_shapes_formatted"] = text_shapes_formatted
            agg["text_shapes_formatted"] += int(text_shapes_formatted)
        else:
            res["shapes_found"] = "SKIPPED"
            res["text_shapes_found"] = "SKIPPED"
            res["text_shapes_formatted"] = 0

        if enable_table_centering:
            if substep:
                substep("table_centering")
            t_found, t_centered = center_tables_in_range(hf.Range)
            if t_found > 0:
                res["tables_centering_applied"] = "YES"
                res["tables_centered_count"] = int(t_centered)
                agg["tables_centered"] += int(t_centered)
        else:
            res["tables_centering_applied"] = "SKIPPED"
            res["tables_centered_count"] = 0

        if substep:
            substep("range_empty_check")
        text_shapes_for_empty = 0 if not enable_shapes else int(res.get("text_shapes_count", 0))
        res["range_empty"] = "YES" if range_is_empty(hf.Range, tbl_count, text_shapes_for_empty) else "NO"

        # SAFE blank cleanup (single delete)
        if enable_blank_cleanup and tbl_count > 0:
            if substep:
                substep("blank_cleanup")
            tbl = hf.Range.Tables(tbl_count)
            blank_count, deleted = enforce_max_blank_paragraphs_after_table_safe(
                hf.Range, tbl, max_blanks=res["spec_max"]
            )
            res["post_table_blanks_found"] = "YES" if blank_count > 0 else "NO"
            res["post_table_blank_count"] = blank_count
            res["post_table_deleted"] = deleted
            res["post_table_deletions_required"] = "YES" if blank_count > res["spec_max"] else "NO"
            res["post_table_remaining"] = max(blank_count - deleted, 0)
            actions += int(deleted)
            agg["post_table_deleted"] += int(deleted)

            # Footer-specific: remove any blank lines AFTER the Confidential line
            if mode == "footer":
                phrase = "Neurotech Pharmaceuticals"
                _, deleted_after_phrase = enforce_max_blank_paragraphs_after_phrase_safe(hf.Range, phrase, max_blanks=0)
                actions += int(deleted_after_phrase)
                agg["post_table_deleted"] += int(deleted_after_phrase)

        elif not enable_blank_cleanup:
            res["post_table_blanks_found"] = "SKIPPED"
            res["post_table_deletions_required"] = "SKIPPED"

    except Exception:
        res["present"] = "NO"

    return res, actions, agg


def process_headers_footers_only(
    doc,
    *,
    substep: Optional[Callable[[str], None]] = None,
    enable_shapes: bool = False,
    enable_table_centering: bool = True,
    enable_blank_cleanup: bool = True,
) -> tuple[dict, int, dict]:
    details = {"sections": 0, "sections_data": []}
    aggregates = {
        "format_apps": 0,
        "header_post_table_deleted": 0,
        "footer_post_table_deleted": 0,
        "tables_centered": 0,
        "text_shapes_formatted": 0,
    }
    actions_total = 0

    sec_count = int(doc.Sections.Count)
    details["sections"] = sec_count

    for sidx in range(1, sec_count + 1):
        sec = doc.Sections(sidx)
        sec_block = {"section_index": sidx, "headers": {}, "footers": {}}

        for kind, kind_name in HF_KINDS:
            # Header
            try:
                hdr = sec.Headers(kind)

                def hs(step: str, s=sidx, k=kind_name):
                    if substep:
                        substep(f"S{s} HEADER {k} | {step}")

                hdr_res, hdr_actions, hdr_agg = process_headerfooter(
                    hdr, doc, mode="header",
                    substep=hs,
                    enable_shapes=enable_shapes,
                    enable_table_centering=enable_table_centering,
                    enable_blank_cleanup=enable_blank_cleanup,
                )
            except Exception:
                hdr_res, hdr_actions, hdr_agg = (
                    {"present": "NO"}, 0,
                    {"format_apps": 0, "post_table_deleted": 0, "tables_centered": 0, "text_shapes_formatted": 0}
                )

            sec_block["headers"][kind_name] = hdr_res
            actions_total += hdr_actions
            aggregates["format_apps"] += int(hdr_agg.get("format_apps", 0))
            aggregates["header_post_table_deleted"] += int(hdr_agg.get("post_table_deleted", 0))
            aggregates["tables_centered"] += int(hdr_agg.get("tables_centered", 0))
            aggregates["text_shapes_formatted"] += int(hdr_agg.get("text_shapes_formatted", 0))

            # Footer
            try:
                ftr = sec.Footers(kind)

                def fs(step: str, s=sidx, k=kind_name):
                    if substep:
                        substep(f"S{s} FOOTER {k} | {step}")

                ftr_res, ftr_actions, ftr_agg = process_headerfooter(
                    ftr, doc, mode="footer",
                    substep=fs,
                    enable_shapes=enable_shapes,
                    enable_table_centering=enable_table_centering,
                    enable_blank_cleanup=enable_blank_cleanup,
                )
            except Exception:
                ftr_res, ftr_actions, ftr_agg = (
                    {"present": "NO"}, 0,
                    {"format_apps": 0, "post_table_deleted": 0, "tables_centered": 0, "text_shapes_formatted": 0}
                )

            sec_block["footers"][kind_name] = ftr_res
            actions_total += ftr_actions
            aggregates["format_apps"] += int(ftr_agg.get("format_apps", 0))
            aggregates["footer_post_table_deleted"] += int(ftr_agg.get("post_table_deleted", 0))
            aggregates["tables_centered"] += int(ftr_agg.get("tables_centered", 0))
            aggregates["text_shapes_formatted"] += int(ftr_agg.get("text_shapes_formatted", 0))

        details["sections_data"].append(sec_block)

    return details, actions_total, aggregates


# -----------------------------
# PDF report generation
# -----------------------------
class NumberedCanvas(canvas.Canvas):
    """
    Canvas that renders "Run Report #: <id>" and "Page n of N" on every page.
    IMPORTANT: Do NOT pass pagesize twice. ReportLab passes pagesize in kwargs.
    """
    def __init__(self, *args, run_id: str, left_margin: float, right_margin: float, **kwargs):
        super().__init__(*args, **kwargs)  # pagesize comes via kwargs from ReportLab
        self._run_id = run_id
        self._left_margin = left_margin
        self._right_margin = right_margin
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        super().showPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_header(num_pages)
            super().showPage()
        super().save()

    def _draw_header(self, page_count: int):
        width, height = self._pagesize  # ReportLab sets this
        y = height - 0.50 * inch
        self.saveState()
        self.setFont("Helvetica", 9)
        self.drawString(self._left_margin, y, f"Run Report #: {self._run_id}")
        self.drawRightString(width - self._right_margin, y, f"Page {self.getPageNumber()} of {page_count}")
        self.restoreState()


def _wrap_line(line: str, max_chars: int) -> list[str]:
    if len(line) <= max_chars:
        return [line]
    out = []
    s = line
    while len(s) > max_chars:
        cut = max_chars
        for sep in [" | ", " : ", "\\", "/", " "]:
            pos = s.rfind(sep, 0, max_chars)
            if pos > 20:
                cut = pos + len(sep) if sep.strip() else pos
                break
        out.append(s[:cut].rstrip())
        s = s[cut:].lstrip()
    if s:
        out.append(s)
    return out


def _wrap_lines(lines: list[str], max_chars: int) -> list[str]:
    wrapped: list[str] = []
    for ln in lines:
        wrapped.extend(_wrap_line(ln, max_chars))
    return wrapped


def generate_pdf_report(
    pdf_path: str,
    run_id: str,
    stats,
    input_root: str,
    output_root: str,
    backups_enabled: bool,
    backup_root: str,
    error_rows: list[tuple[str, str, str]],
    doc_blocks: list[list[str]],
) -> None:
    ensure_dir(os.path.dirname(pdf_path))

    styles = getSampleStyleSheet()
    title = ParagraphStyle("title", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=16, spaceAfter=12)
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, spaceBefore=12, spaceAfter=6)
    body = ParagraphStyle("body", parent=styles["BodyText"], fontName="Helvetica", fontSize=10, leading=12)
    mono = ParagraphStyle("mono", parent=styles["BodyText"], fontName="Courier", fontSize=8, leading=9)

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=LETTER,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.90 * inch,
        bottomMargin=0.75 * inch,
        title=f"Run Report {run_id}",
    )

    story: list = []
    story.append(Paragraph("Header/Footer Formatting Run Report", title))
    story.append(Paragraph(f"Run ID: {run_id}", body))
    story.append(Spacer(1, 10))

    story.append(Paragraph("Run Metadata", h2))
    start_local = stats.start_dt
    end_local = stats.end_dt or datetime.now().astimezone()
    tz_line = f"{stats.tz_name} ({stats.tz_offset})" if stats.tz_name or stats.tz_offset else "UNKNOWN"

    meta_rows = [
        ["Run Start (Local)", start_local.strftime("%d-%b-%Y %H:%M:%S")],
        ["Run End (Local)", end_local.strftime("%d-%b-%Y %H:%M:%S")],
        ["Timezone", tz_line],
        ["Elapsed", time.strftime("%H:%M:%S", time.gmtime(max(0, int((end_local.timestamp() - start_local.timestamp())))))],
        ["User", stats.user],
        ["Computer", stats.computer],
        ["OS", stats.os_info],
        ["Python", stats.python_info],
        ["MS Word", stats.word_version],
        ["Input Root", os.path.abspath(input_root)],
        ["Output Root", os.path.abspath(output_root)],
        ["Backups Enabled", "YES" if backups_enabled else "NO"],
        ["Backup Root", os.path.abspath(backup_root) if backups_enabled else "N/A"],
        ["PDF Export Enabled", "YES" if getattr(stats, "pdf_exports_enabled", False) else "NO"],
        ["DocID Naming Enabled", "YES" if getattr(stats, "docid_naming_enabled", False) else "NO"],
    ]
    meta_tbl = Table(meta_rows, colWidths=[2.0 * inch, 5.5 * inch])
    meta_tbl.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 9),
        ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.whitesmoke, colors.white]),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(meta_tbl)
    story.append(Spacer(1, 12))

    story.append(Paragraph("Summary Statistics", h2))
    summary_rows = [
        ["Word files discovered", str(stats.total_files)],
        ["Documents processed", str(stats.processed)],
        ["Succeeded", str(stats.succeeded)],
        ["Failed", str(stats.failed)],
        ["Skipped", str(stats.skipped)],
        ["Backup attempts", str(stats.backups_attempted)],
        ["Backup succeeded", str(stats.backups_ok)],
        ["Backup failed", str(stats.backups_failed)],
        ["Header post-table deletions", str(stats.header_post_table_deleted_total)],
        ["Footer post-table deletions", str(stats.footer_post_table_deleted_total)],
        ["Tables centered (count)", str(stats.tables_centered_total)],
        ["Text shapes formatted (count)", str(stats.text_shapes_formatted_total)],
        ["Format applications (count)", str(stats.format_apps_total)],
        ["Actions (total)", str(stats.total_actions)],
        ["PDF exports attempted", str(getattr(stats, "pdf_exports_attempted", 0))],
        ["PDF exports succeeded", str(getattr(stats, "pdf_exports_ok", 0))],
        ["PDF exports failed", str(getattr(stats, "pdf_exports_failed", 0))],
        ["DocID found (count)", str(getattr(stats, "docid_found_count", 0))],
        ["DocID used for naming (count)", str(getattr(stats, "docid_used_count", 0))],
        ["DocID missing (count)", str(getattr(stats, "docid_missing_count", 0))],
    ]
    summary_tbl = Table(summary_rows, colWidths=[3.0 * inch, 4.5 * inch])
    summary_tbl.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 9),
        ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.whitesmoke]),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(summary_tbl)
    story.append(Spacer(1, 12))

    story.append(Paragraph("Error Summary", h2))
    if error_rows:
        err_data = [["File", "Stage", "Error Summary"]] + list(error_rows)
        err_tbl = Table(err_data, colWidths=[2.1 * inch, 1.0 * inch, 4.4 * inch])
        err_tbl.setStyle(TableStyle([
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 9),
            ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(err_tbl)
    else:
        story.append(Paragraph("No errors recorded.", body))

    story.append(PageBreak())
    story.append(Paragraph("Detailed Event Log", h2))
    story.append(Paragraph("Logs are grouped by document. Timestamps appear only at document block start/end.", body))
    story.append(Spacer(1, 8))

    max_chars = 110
    for block in doc_blocks:
        wrapped = _wrap_lines(block, max_chars=max_chars)
        story.append(Preformatted("\n".join(wrapped), mono))
        story.append(Spacer(1, 10))

    from typing import cast, Any
    from reportlab.pdfgen.canvas import Canvas as _Canvas
    try:
        from reportlab.platypus.doctemplate import _CanvasMaker as _RLCanvasMaker  # type: ignore
    except Exception:
        _RLCanvasMaker = Any

    def _canvas_maker(filename, **kwargs) -> _Canvas:
        return NumberedCanvas(
            filename,
            run_id=run_id,
            left_margin=doc.leftMargin,
            right_margin=doc.rightMargin,
            **kwargs
        )

    doc.build(
        story,
        canvasmaker=cast(_RLCanvasMaker, _canvas_maker)
    )


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Header/Footer Formatter + Cleanup (Optional Backups) + PDF Report")
        self.geometry("1180x840")
        self.minsize(1050, 720)

        self._log_q: queue.Queue[str] = queue.Queue()
        self._worker_thread: threading.Thread | None = None
        self._stop_event = threading.Event()
        self._stats: RunStats | None = None

        self._status_lock = threading.Lock()
        self._current_file = ""
        self._current_step = ""
        self._current_step_since = time.time()
        self._word_pid: int | None = None
        self._word_killed_by_watchdog = False

        self._build_ui()
        self.after(100, self._drain_log_queue)

    def _set_status(self, file_: str, step: str):
        with self._status_lock:
            self._current_file = file_
            self._current_step = step
            self._current_step_since = time.time()

    def _set_word_pid(self, pid: int | None):
        with self._status_lock:
            self._word_pid = pid

    def _get_status_snapshot(self):
        with self._status_lock:
            return self._current_file, self._current_step, self._current_step_since, self._word_pid

    def _enqueue_lines(self, lines: list[str]):
        for line in lines:
            self._log_q.put(line)

    def _emit(self, line: str):
        self._log_q.put(line)

    def _drain_log_queue(self):
        try:
            while True:
                line = self._log_q.get_nowait()
                self.log.insert("end", line + "\n")
                self.log.see("end")
        except queue.Empty:
            pass
        self.after(100, self._drain_log_queue)

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="x", **pad)

        ttk.Label(frm, text="Input folder (Word docs):").grid(row=0, column=0, sticky="w")
        self.in_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.in_var, width=105).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm, text="Browse…", command=self._browse_in).grid(row=0, column=2)

        ttk.Label(frm, text="Output folder (run-stamped subfolder created):").grid(row=1, column=0, sticky="w")
        self.out_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.out_var, width=105).grid(row=1, column=1, sticky="we", padx=6)
        ttk.Button(frm, text="Browse…", command=self._browse_out).grid(row=1, column=2)

        ttk.Label(frm, text="Backup folder (optional; run-stamped subfolder created):").grid(row=2, column=0, sticky="w")
        self.backup_var = tk.StringVar()
        self.backup_entry = ttk.Entry(frm, textvariable=self.backup_var, width=105)
        self.backup_entry.grid(row=2, column=1, sticky="we", padx=6)
        self.backup_btn = ttk.Button(frm, text="Browse…", command=self._browse_backup)
        self.backup_btn.grid(row=2, column=2)

        self.backups_enabled_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frm,
            text="Enable backups (copy originals to Backup folder before processing)",
            variable=self.backups_enabled_var,
            command=self._toggle_backup_widgets
        ).grid(row=3, column=1, sticky="w", pady=(2, 0))

        opt = ttk.Frame(frm)
        opt.grid(row=4, column=1, sticky="w", pady=(6, 0))

        self.show_word_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="Show Word (debug)", variable=self.show_word_var).pack(side="left")

        ttk.Label(opt, text="Timeout (sec):").pack(side="left", padx=(18, 6))
        self.timeout_var = tk.StringVar(value="180")
        ttk.Entry(opt, textvariable=self.timeout_var, width=6).pack(side="left")

        self.process_shapes_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="Process Shapes/TextBoxes (optional)", variable=self.process_shapes_var).pack(side="left", padx=(18, 0))

        # NEW: per-doc PDF export
        self.export_pdf_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="Export PDF for each processed document", variable=self.export_pdf_var).pack(side="left", padx=(18, 0))

        # NEW: name outputs using Document ID from header
        self.name_by_docid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="Name outputs using Document ID (from header)", variable=self.name_by_docid_var).pack(side="left", padx=(18, 0))

        frm.columnconfigure(1, weight=1)

        ctr = ttk.Frame(self)
        ctr.pack(fill="x", **pad)

        self.run_label = ttk.Label(ctr, text="Run #: (not started)")
        self.run_label.pack(side="left")

        self.start_btn = ttk.Button(ctr, text="Start", command=self._start)
        self.start_btn.pack(side="right", padx=6)

        self.stop_btn = ttk.Button(ctr, text="Stop", command=self._stop, state="disabled")
        self.stop_btn.pack(side="right")

        pr = ttk.Frame(self)
        pr.pack(fill="x", **pad)

        self.prog = ttk.Progressbar(pr, mode="determinate")
        self.prog.pack(fill="x")

        self.stats_label = ttk.Label(pr, text="")
        self.stats_label.pack(anchor="w", pady=(6, 0))

        logf = ttk.Frame(self)
        logf.pack(fill="both", expand=True, **pad)

        ttk.Label(logf, text="Event log:").pack(anchor="w")
        self.log = tk.Text(logf, wrap="word", height=28)
        self.log.pack(fill="both", expand=True, pady=(6, 0))

        self._toggle_backup_widgets()

    def _toggle_backup_widgets(self):
        enabled = self.backups_enabled_var.get()
        state = "normal" if enabled else "disabled"
        self.backup_entry.configure(state=state)
        self.backup_btn.configure(state=state)

    def _browse_in(self):
        p = filedialog.askdirectory(title="Select input folder")
        if p:
            self.in_var.set(p)

    def _browse_out(self):
        p = filedialog.askdirectory(title="Select output folder")
        if p:
            self.out_var.set(p)

    def _browse_backup(self):
        p = filedialog.askdirectory(title="Select backup folder")
        if p:
            self.backup_var.set(p)

    def _update_stats_line(self):
        if not self._stats:
            return
        s = self._stats
        elapsed = time.time() - s.start_ts
        extra = ""
        if s.backups_enabled:
            extra = f" | Backups: {s.backups_ok}/{s.backups_attempted} OK ({s.backups_failed} fail)"
        if getattr(s, "pdf_exports_enabled", False):
            extra += f" | PDFs: {s.pdf_exports_ok}/{s.pdf_exports_attempted} OK ({s.pdf_exports_failed} fail)"
        if getattr(s, "docid_naming_enabled", False):
            extra += f" | DocID used: {s.docid_used_count}/{s.total_files}"
        self.stats_label.config(
            text=f"Total: {s.total_files} | Processed: {s.processed} | OK: {s.succeeded} | "
                 f"Failed: {s.failed} | Skipped: {s.skipped} | Actions: {s.total_actions} | "
                 f"Elapsed: {elapsed:.1f}s{extra}"
        )

    def _start(self):
        if win32 is None or pythoncom is None:
            messagebox.showerror("Missing dependency", "pywin32 is required on Windows with Microsoft Word installed.")
            return
        if SimpleDocTemplate is None:
            messagebox.showerror("Missing dependency", "reportlab is required: pip install reportlab")
            return

        in_dir = self.in_var.get().strip()
        out_dir = self.out_var.get().strip()
        backups_enabled = bool(self.backups_enabled_var.get())
        backup_root = self.backup_var.get().strip()

        export_pdf_each = bool(self.export_pdf_var.get())
        name_by_docid = bool(self.name_by_docid_var.get())

        if not in_dir or not os.path.isdir(in_dir):
            messagebox.showerror("Invalid input", "Please select a valid input folder.")
            return
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showerror("Invalid output", "Please select a valid output folder.")
            return
        if backups_enabled and (not backup_root or not os.path.isdir(backup_root)):
            messagebox.showerror("Invalid backup folder", "Backups are enabled. Please select a valid backup folder.")
            return

        try:
            timeout_s = int(self.timeout_var.get().strip())
            if timeout_s < 10:
                raise ValueError
        except Exception:
            messagebox.showerror("Invalid timeout", "Timeout (sec) must be an integer >= 10.")
            return

        if self._worker_thread and self._worker_thread.is_alive():
            messagebox.showwarning("Already running", "A run is already in progress.")
            return

        self._stop_event.clear()
        run_id = make_run_id()

        now_local = datetime.now().astimezone()
        tzname = now_local.tzname() or ""
        offset = now_local.strftime("%z")
        offset_fmt = f"UTC{offset[:3]}:{offset[3:]}" if offset else ""
        user = f"{os.environ.get('USERDOMAIN','')}\\{getpass.getuser()}" if os.environ.get("USERDOMAIN") else getpass.getuser()

        self._stats = RunStats(
            run_id=run_id,
            start_ts=time.time(),
            start_dt=now_local,
            backups_enabled=backups_enabled,
            pdf_exports_enabled=export_pdf_each,
            docid_naming_enabled=name_by_docid,
            user=user,
            computer=socket.gethostname(),
            os_info=f"{platform.system()} {platform.release()} ({platform.version()})",
            python_info=f"{platform.python_version()} ({platform.architecture()[0]})",
            tz_name=tzname,
            tz_offset=offset_fmt,
        )

        self.run_label.config(text=f"Run #: {run_id}")
        self.log.delete("1.0", "end")
        self.prog["value"] = 0
        self.stats_label.config(text="")
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")

        self._worker_thread = threading.Thread(
            target=self._worker,
            args=(
                in_dir, out_dir, run_id, backups_enabled, backup_root,
                timeout_s, bool(self.show_word_var.get()),
                bool(self.process_shapes_var.get()),
                export_pdf_each,
                name_by_docid
            ),
            daemon=True
        )
        self._worker_thread.start()

    def _stop(self):
        self._stop_event.set()
        self._enqueue_lines([ts_line(), "STOP REQUESTED", "Run will halt after current file.", ts_line()])

    def _finalize_ui(self):
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self._update_stats_line()

    def _copy_backup(self, abs_in: str, rel_path: str, backup_run_root: str) -> tuple[str, str]:
        if not self._stats or not self._stats.backups_enabled:
            return "DISABLED", "N/A"

        self._stats.backups_attempted += 1
        dest = os.path.join(backup_run_root, rel_path)
        try:
            ensure_dir(os.path.dirname(dest))
            shutil.copy2(abs_in, dest)
            self._stats.backups_ok += 1
            return "SUCCESS", dest
        except Exception:
            self._stats.backups_failed += 1
            return "FAIL", dest

    def _get_word_pid_from_app(self, word_app) -> int | None:
        if win32process is None:
            return None
        try:
            hwnd = word_app.Hwnd
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return int(pid)
        except Exception:
            return None

    def _kill_word_pid(self, pid: int) -> bool:
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True, text=True, check=False)
            return True
        except Exception:
            return False

    def _worker(
        self,
        in_dir: str,
        out_dir: str,
        run_id: str,
        backups_enabled: bool,
        backup_root: str,
        timeout_s: int,
        show_word: bool,
        enable_shapes: bool,
        export_pdf_each: bool,
        name_by_docid: bool
    ):
        assert self._stats is not None
        stats = self._stats

        run_root = os.path.join(out_dir, f"HF_Fix_{run_id}")
        out_docs_root = os.path.join(run_root, "docs")
        work_root = os.path.join(run_root, "working")
        ensure_dir(out_docs_root)
        ensure_dir(work_root)

        backup_run_root = "N/A"
        if backups_enabled:
            backup_run_root = os.path.join(backup_root, f"HF_Backups_{run_id}")
            ensure_dir(backup_run_root)

        files = list(iter_word_files(in_dir))
        stats.total_files = len(files)

        self.after(0, lambda: self.prog.configure(maximum=max(stats.total_files, 1)))
        self.after(0, self._update_stats_line)

        run_log_lines: list[str] = []
        doc_blocks: list[list[str]] = []
        error_rows: list[tuple[str, str, str]] = []

        run_start_lines = [
            ts_line(),
            "RUN START",
            f"RUN ID                      : {run_id}",
            f"INPUT ROOT                  : {os.path.abspath(in_dir)}",
            f"OUTPUT ROOT                 : {os.path.abspath(out_docs_root)}",
            f"WORKING ROOT                : {os.path.abspath(work_root)}",
            f"BACKUPS ENABLED             : {'YES' if backups_enabled else 'NO'}",
            f"BACKUP ROOT (RUN FOLDER)    : {os.path.abspath(backup_run_root) if backups_enabled else 'N/A'}",
            f"TIMEOUT (sec)               : {timeout_s}",
            f"SHOW WORD (debug)           : {'YES' if show_word else 'NO'}",
            f"PROCESS SHAPES              : {'YES' if enable_shapes else 'NO'}",
            f"EXPORT PDF EACH DOC         : {'YES' if export_pdf_each else 'NO'}",
            f"NAME OUTPUTS BY DOC ID      : {'YES' if name_by_docid else 'NO'}",
            ts_line(),
            "",
        ]
        self._enqueue_lines(run_start_lines)
        run_log_lines.extend(run_start_lines)

        if stats.total_files == 0:
            lines = [ts_line(), "No Word files found. Nothing to do.", ts_line()]
            self._enqueue_lines(lines)
            run_log_lines.extend(lines)
            stats.end_dt = datetime.now().astimezone()
            self.after(0, self._finalize_ui)
            return

        hb_stop = threading.Event()

        def heartbeat_and_watchdog():
            while not hb_stop.is_set():
                time.sleep(15)
                f, st, since, pid = self._get_status_snapshot()
                if f or st:
                    age = int(time.time() - since)
                    self._emit(f"[HEARTBEAT] still running | step={st} | file={f} | seconds_in_step={age}")
                    if pid and age >= timeout_s:
                        self._emit(f"[WATCHDOG] TIMEOUT exceeded ({timeout_s}s). Killing WINWORD pid={pid} ...")
                        ok = self._kill_word_pid(pid)
                        self._emit(f"[WATCHDOG] taskkill result: {'SUCCESS' if ok else 'FAIL'}")
                        with self._status_lock:
                            self._word_killed_by_watchdog = True

        threading.Thread(target=heartbeat_and_watchdog, daemon=True).start()

        pythoncom.CoInitialize()
        word = None

        def start_word() -> tuple[object | None, int | None]:
            try:
                w = win32.DispatchEx("Word.Application")
                try:
                    w.Visible = False
                except Exception:
                    pass
                w.DisplayAlerts = wdAlertsNone
                try:
                    w.Options.UpdateLinksAtOpen = False
                except Exception:
                    pass
                try:
                    w.Options.ConfirmConversions = False
                except Exception:
                    pass
                try:
                    w.Options.SaveNormalPrompt = False
                except Exception:
                    pass
                try:
                    w.AutomationSecurity = msoAutomationSecurityForceDisable
                except Exception:
                    pass
                pid = self._get_word_pid_from_app(w)
                return w, pid
            except Exception:
                return None, None

        def stop_word(w) -> None:
            try:
                w.Quit()
            except Exception:
                pass

        try:
            word, pid = start_word()
            self._set_word_pid(pid if pid else None)
            if word is None:
                raise RuntimeError("Could not start Word via COM automation.")
            try:
                stats.word_version = str(word.Version)
            except Exception:
                stats.word_version = "UNKNOWN"

            for idx, (abs_path, rel_path) in enumerate(files, start=1):
                if self._stop_event.is_set():
                    break

                file_label = os.path.basename(rel_path)
                abs_in = os.path.abspath(abs_path)

                # We'll compute abs_out after (optional) Document ID extraction.
                work_abs = os.path.abspath(os.path.join(work_root, rel_path))

                stats.processed = idx
                self.after(0, lambda v=idx: self.prog.configure(value=v))
                self.after(0, self._update_stats_line)

                with self._status_lock:
                    self._word_killed_by_watchdog = False

                self._set_status(file_label, "Working copy")
                self._emit(f"STEP: begin file {idx}/{stats.total_files} | {file_label}")
                self._emit("STEP: working copy (start)")
                try:
                    make_working_copy(abs_in, work_abs)
                    self._emit(f"STEP: working copy (done) | work_path={work_abs}")
                except Exception as e:
                    stats.failed += 1
                    err = f"FAIL: working copy error: {e}"
                    error_rows.append((file_label, "WorkingCopy", err[:250]))
                    self._emit(f"STEP: working copy (FAIL) | {e}")
                    continue

                self._set_status(file_label, "Backup copy")
                self._emit("STEP: backup copy (start)")
                backup_status, backup_dest = self._copy_backup(abs_in, rel_path, backup_run_root)
                self._emit(f"STEP: backup copy (done) | status={backup_status}")

                doc = None
                editability = {}
                details = {"sections": 0, "sections_data": []}
                actions = 0
                aggs = {"format_apps": 0, "header_post_table_deleted": 0, "footer_post_table_deleted": 0, "tables_centered": 0, "text_shapes_formatted": 0}

                doc_id_value: str | None = None
                output_docx_path = ""
                output_pdf_path = ""
                used_base_name = ""
                original_ext = os.path.splitext(file_label)[1] or ".docx"

                def substep(msg: str):
                    self._set_status(file_label, f"Process headers/footers | {msg}")
                    self._emit(f"SUBSTEP: {msg}")

                try:
                    self._set_status(file_label, "Word OPEN")
                    self._emit("STEP: Word OPEN (start)")
                    doc = word.Documents.Open(
                        work_abs,
                        ReadOnly=False,
                        AddToRecentFiles=False,
                        Visible=False,
                        ConfirmConversions=False
                    )
                    self._emit("STEP: Word OPEN (done)")
                    if self._word_killed_by_watchdog:
                        raise TimeoutError(f"TIMEOUT during Word OPEN (>{timeout_s}s)")

                    # Debug show Word if requested
                    if show_word:
                        try:
                            word.Visible = True
                            doc.Activate()
                        except Exception:
                            pass

                    self._set_status(file_label, "Editability")
                    self._emit("STEP: Editability (start)")
                    clear_readonly_attribute(work_abs)
                    editability = force_document_editable(doc)
                    self._emit(f"STEP: Editability (done) | doc_readonly_final={editability.get('doc_readonly_final','UNKNOWN')}")
                    if self._word_killed_by_watchdog:
                        raise TimeoutError(f"TIMEOUT during Editability (>{timeout_s}s)")

                    # NEW: Extract Document ID (optional) BEFORE SaveAs2
                    if name_by_docid:
                        self._set_status(file_label, "Extract Document ID")
                        self._emit("STEP: Extract Document ID (start)")
                        try:
                            doc_id_value = extract_document_id_from_headers(doc)
                            if doc_id_value:
                                stats.docid_found_count += 1
                                self._emit(f"STEP: Extract Document ID (done) | FOUND=YES | DOC_ID={doc_id_value}")
                            else:
                                stats.docid_missing_count += 1
                                self._emit("STEP: Extract Document ID (done) | FOUND=NO")
                        except Exception as e:
                            stats.docid_missing_count += 1
                            doc_id_value = None
                            self._emit(f"STEP: Extract Document ID (FAIL) | {e}")

                    self._set_status(file_label, "Process headers/footers")
                    self._emit("STEP: Process headers/footers (start)")
                    details, actions, aggs = process_headers_footers_only(
                        doc,
                        substep=substep,
                        enable_shapes=enable_shapes,
                        enable_table_centering=True,
                        enable_blank_cleanup=True,
                    )
                    details["editability"] = editability
                    self._emit("STEP: Process headers/footers (done)")
                    if self._word_killed_by_watchdog:
                        raise TimeoutError(f"TIMEOUT during processing (>{timeout_s}s)")

                    # NEW: Compute output paths (optionally rename by DocID)
                    self._set_status(file_label, "Compute output paths")
                    rel_dir = os.path.dirname(rel_path)
                    out_dir_for_file = os.path.abspath(os.path.join(out_docs_root, rel_dir))
                    ensure_dir(out_dir_for_file)

                    original_base = os.path.splitext(os.path.basename(rel_path))[0]
                    base_candidate = original_base
                    if name_by_docid and doc_id_value:
                        base_candidate = doc_id_value
                        stats.docid_used_count += 1

                    exts_to_check = [original_ext]
                    if export_pdf_each:
                        exts_to_check.append(".pdf")

                    used_base_name = pick_unique_base_name(out_dir_for_file, base_candidate, exts_to_check)
                    output_docx_path = os.path.join(out_dir_for_file, used_base_name + original_ext)
                    output_pdf_path = os.path.join(out_dir_for_file, used_base_name + ".pdf")

                    self._emit(
                        "STEP: Output naming | "
                        f"USE_DOCID={'YES' if (name_by_docid and doc_id_value) else 'NO'} | "
                        f"DOC_ID={doc_id_value or 'N/A'} | "
                        f"BASE={used_base_name} | "
                        f"DOCX_OUT={output_docx_path} | "
                        f"PDF_OUT={output_pdf_path if export_pdf_each else 'SKIPPED'}"
                    )

                    self._set_status(file_label, "Word SAVE")
                    self._emit("STEP: Word SAVE (start)")
                    clear_readonly_attribute(output_docx_path)
                    doc.SaveAs2(output_docx_path)
                    self._emit("STEP: Word SAVE (done)")

                    # NEW: optional per-document PDF export
                    if export_pdf_each:
                        stats.pdf_exports_attempted += 1
                        self._set_status(file_label, "Export PDF")
                        self._emit("STEP: Export PDF (start)")
                        try:
                            export_doc_to_pdf(doc, output_pdf_path)
                            stats.pdf_exports_ok += 1
                            self._emit(f"STEP: Export PDF (done) | status=SUCCESS | pdf_path={output_pdf_path}")
                        except Exception as e:
                            stats.pdf_exports_failed += 1
                            self._emit(f"STEP: Export PDF (FAIL) | {e}")

                    # Close document
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                    doc = None

                    # DOCX ZIP/XML post-processing DISABLED.
                    self._emit("STEP: DOCX postprocess (SKIPPED) | reason=com_only_to_avoid_word_repair_prompt")

                    stats.succeeded += 1
                    stats.total_actions += int(actions)
                    stats.format_apps_total += int(aggs.get("format_apps", 0))
                    stats.header_post_table_deleted_total += int(aggs.get("header_post_table_deleted", 0))
                    stats.footer_post_table_deleted_total += int(aggs.get("footer_post_table_deleted", 0))
                    stats.tables_centered_total += int(aggs.get("tables_centered", 0))
                    stats.text_shapes_formatted_total += int(aggs.get("text_shapes_formatted", 0))

                    block: list[str] = []
                    block.append(ts_line())
                    block.append("DOCUMENT START")
                    block.append(f"INDEX                        : {idx}/{stats.total_files}")
                    block.append(f"FILE (INPUT NAME)            : {file_label}")
                    block.append(f"INPUT ABS PATH               : {abs_in}")
                    block.append(f"WORKING ABS PATH             : {work_abs}")
                    block.append(f"OUTPUT DOCX ABS PATH         : {output_docx_path}")
                    block.append(f"OUTPUT PDF ABS PATH          : {output_pdf_path if export_pdf_each else 'SKIPPED'}")
                    block.append(f"DOC ID NAMING ENABLED        : {'YES' if name_by_docid else 'NO'}")
                    block.append(f"DOC ID FOUND                 : {'YES' if doc_id_value else 'NO'}")
                    block.append(f"DOC ID VALUE                 : {doc_id_value if doc_id_value else 'N/A'}")
                    block.append(f"OUTPUT BASE NAME             : {used_base_name if used_base_name else 'N/A'}")
                    block.append(f"EXPORT PDF ENABLED           : {'YES' if export_pdf_each else 'NO'}")
                    block.append(f"BACKUPS ENABLED              : {'YES' if backups_enabled else 'NO'}")
                    block.append(f"BACKUP COPY STATUS           : {backup_status}")
                    block.append(f"BACKUP DEST ABS PATH         : {backup_dest if backups_enabled else 'N/A'}")
                    block.append("")
                    block.append("EDITABILITY CHECK")
                    block.append(f"  DOC READONLY (INITIAL)     : {editability.get('doc_readonly_initial','UNKNOWN')}")
                    block.append(f"  UNPROTECT SUCCESS          : {editability.get('unprotect_success','NO')}")
                    block.append(f"  FINAL CLEARED SUCCESS      : {editability.get('final_cleared_success','NO')}")
                    block.append(f"  READONLY RECOMMENDED CLR   : {editability.get('readonly_recommended_cleared_success','NO')}")
                    block.append(f"  CHANGEFILEACCESS SUCCESS   : {editability.get('changefileaccess_success','NO')}")
                    block.append(f"  DOC READONLY (FINAL)       : {editability.get('doc_readonly_final','UNKNOWN')}")
                    block.append("")
                    block.append("OPEN/PROCESS STATUS          : SUCCESS")
                    block.append(f"SECTION COUNT                : {details.get('sections', 0)}")
                    block.append(f"PROCESS SHAPES               : {'YES' if enable_shapes else 'NO'}")
                    block.append("")

                    sec_count = int(details.get("sections", 0))
                    for sec in details.get("sections_data", []):
                        sidx = sec["section_index"]
                        block.append(f"SECTION {sidx}/{sec_count}")
                        for kind_name in ["PRIMARY", "FIRSTPAGE", "EVENPAGES"]:
                            hdr = sec["headers"].get(kind_name, {"present": "NO"})
                            block.append(f"  HEADER {kind_name}")
                            block.append(f"    PRESENT                  : {hdr.get('present','NO')}")
                            block.append(f"    RANGE EMPTY              : {hdr.get('range_empty','UNKNOWN')}")
                            block.append(f"    FORMAT NORMALIZATION     : APPLIED={hdr.get('format_applied','NO')} | TARGET=(Before=0, After=0, Line=Single, Align=Center)")
                            block.append(f"    TABLES FOUND             : {hdr.get('tables_found','NO')} | COUNT={hdr.get('table_count','UNKNOWN')} | ANCHOR={hdr.get('anchor','N/A')}")
                            block.append(f"    TABLE CENTERING          : APPLIED={hdr.get('tables_centering_applied','NO')} | CENTERED_COUNT={hdr.get('tables_centered_count',0)}")
                            block.append(f"    BLANK CLEANUP METHOD     : {hdr.get('blank_cleanup_method','UNKNOWN')}")
                            block.append(f"    POST-TABLE SPEC_MAX      : {hdr.get('spec_max','UNKNOWN')}")
                            block.append(f"    POST-TABLE BLANKS FOUND  : {hdr.get('post_table_blanks_found','N/A')} | COUNT={hdr.get('post_table_blank_count','N/A')}")
                            block.append(f"    POST-TABLE BLANKS ACTION : REQUIRED={hdr.get('post_table_deletions_required','N/A')} | DELETED={hdr.get('post_table_deleted',0)} | REMAINING={hdr.get('post_table_remaining','N/A')}")
                            block.append(f"    SHAPES FOUND             : {hdr.get('shapes_found','NO')} | COUNT={hdr.get('shapes_count',0)}")
                            block.append(f"    TEXT SHAPES FOUND        : {hdr.get('text_shapes_found','NO')} | COUNT={hdr.get('text_shapes_count',0)}")
                            block.append(f"    TEXT SHAPES FORMATTED    : {hdr.get('text_shapes_formatted',0)}")
                        for kind_name in ["PRIMARY", "FIRSTPAGE", "EVENPAGES"]:
                            ftr = sec["footers"].get(kind_name, {"present": "NO"})
                            block.append(f"  FOOTER {kind_name}")
                            block.append(f"    PRESENT                  : {ftr.get('present','NO')}")
                            block.append(f"    RANGE EMPTY              : {ftr.get('range_empty','UNKNOWN')}")
                            block.append(f"    FORMAT NORMALIZATION     : APPLIED={ftr.get('format_applied','NO')} | TARGET=(Before=0, After=0, Line=Single, Align=Center)")
                            block.append(f"    TABLES FOUND             : {ftr.get('tables_found','NO')} | COUNT={ftr.get('table_count','UNKNOWN')} | ANCHOR={ftr.get('anchor','N/A')}")
                            block.append(f"    TABLE CENTERING          : APPLIED={ftr.get('tables_centering_applied','NO')} | CENTERED_COUNT={ftr.get('tables_centered_count',0)}")
                            block.append(f"    BLANK CLEANUP METHOD     : {ftr.get('blank_cleanup_method','UNKNOWN')}")
                            block.append(f"    POST-TABLE SPEC_MAX      : {ftr.get('spec_max','UNKNOWN')}")
                            block.append(f"    POST-TABLE BLANKS FOUND  : {ftr.get('post_table_blanks_found','N/A')} | COUNT={ftr.get('post_table_blank_count','N/A')}")
                            block.append(f"    POST-TABLE BLANKS ACTION : REQUIRED={ftr.get('post_table_deletions_required','N/A')} | DELETED={ftr.get('post_table_deleted',0)} | REMAINING={ftr.get('post_table_remaining','N/A')}")
                            block.append(f"    SHAPES FOUND             : {ftr.get('shapes_found','NO')} | COUNT={ftr.get('shapes_count',0)}")
                            block.append(f"    TEXT SHAPES FOUND        : {ftr.get('text_shapes_found','NO')} | COUNT={ftr.get('text_shapes_count',0)}")
                            block.append(f"    TEXT SHAPES FORMATTED    : {ftr.get('text_shapes_formatted',0)}")
                        block.append("")

                    block.append("DOCUMENT SUMMARY")
                    block.append("STATUS                       : OK")
                    block.append(f"ACTIONS (TOTAL)              : {actions}")
                    block.append("DOCUMENT END")
                    block.append(ts_line())
                    block.append("")

                    self._enqueue_lines(block)
                    run_log_lines.extend(block)
                    doc_blocks.append(block)

                except TimeoutError as te:
                    stats.failed += 1
                    msg = f"TIMEOUT: {te}"
                    error_rows.append((file_label, "Timeout", msg[:250]))
                    self._enqueue_lines([f"ERROR: {msg}"])
                    try:
                        if doc is not None:
                            doc.Close(False)
                    except Exception:
                        pass
                    doc = None
                    try:
                        stop_word(word)
                    except Exception:
                        pass
                    word, pid = start_word()
                    self._set_word_pid(pid if pid else None)
                    if word is None:
                        raise RuntimeError("Word restart failed after timeout.")

                except Exception as e:
                    stats.failed += 1
                    err = f"FAIL: {e}"
                    error_rows.append((file_label, "Process", err[:250]))
                    self._enqueue_lines([f"ERROR: {err}", traceback.format_exc()])
                    try:
                        if doc is not None:
                            doc.Close(False)
                    except Exception:
                        pass
                    doc = None

                finally:
                    self.after(0, self._update_stats_line)
                    self._set_status(file_label, "Idle")

            ensure_dir(run_root)
            txt_path = os.path.join(run_root, f"run_log_{run_id}.txt")
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write("\n".join(run_log_lines))

            stats.end_dt = datetime.now().astimezone()
            pdf_path = os.path.join(run_root, f"run_report_{run_id}.pdf")
            generate_pdf_report(
                pdf_path=pdf_path,
                run_id=run_id,
                stats=stats,
                input_root=in_dir,
                output_root=out_docs_root,
                backups_enabled=backups_enabled,
                backup_root=backup_run_root,
                error_rows=error_rows,
                doc_blocks=doc_blocks,
            )
            self._enqueue_lines(["", f"PDF REPORT WRITTEN            : {pdf_path}", f"TEXT LOG WRITTEN              : {txt_path}", ""])

        except Exception as e:
            stats.end_dt = datetime.now().astimezone()
            fatal = [ts_line(), "FATAL ERROR", str(e), traceback.format_exc(), ts_line()]
            self._enqueue_lines(fatal)

        finally:
            hb_stop.set()
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

            elapsed = time.time() - stats.start_ts
            run_end_lines = [
                ts_line(),
                "RUN END",
                f"RUN ID                      : {run_id}",
                f"TOTAL FILES                 : {stats.total_files}",
                f"PROCESSED                   : {stats.processed}",
                f"SUCCEEDED                   : {stats.succeeded}",
                f"FAILED                      : {stats.failed}",
                f"SKIPPED                     : {stats.skipped}",
                f"ACTIONS (TOTAL)             : {stats.total_actions}",
                f"HEADER POST-TABLE DELETIONS : {stats.header_post_table_deleted_total}",
                f"FOOTER POST-TABLE DELETIONS : {stats.footer_post_table_deleted_total}",
                f"TABLES CENTERED (COUNT)     : {stats.tables_centered_total}",
                f"TEXT SHAPES FORMATTED       : {stats.text_shapes_formatted_total}",
                f"FORMAT APPS (COUNT)         : {stats.format_apps_total}",
                f"BACKUPS ENABLED             : {'YES' if stats.backups_enabled else 'NO'}",
                f"BACKUPS ATTEMPTED           : {stats.backups_attempted}",
                f"BACKUPS OK                  : {stats.backups_ok}",
                f"BACKUPS FAILED              : {stats.backups_failed}",
                f"EXPORT PDF EACH DOC         : {'YES' if export_pdf_each else 'NO'}",
                f"PDF EXPORTS ATTEMPTED       : {stats.pdf_exports_attempted}",
                f"PDF EXPORTS OK              : {stats.pdf_exports_ok}",
                f"PDF EXPORTS FAILED          : {stats.pdf_exports_failed}",
                f"NAME OUTPUTS BY DOC ID      : {'YES' if name_by_docid else 'NO'}",
                f"DOC ID FOUND (COUNT)        : {stats.docid_found_count}",
                f"DOC ID USED (COUNT)         : {stats.docid_used_count}",
                f"DOC ID MISSING (COUNT)      : {stats.docid_missing_count}",
                f"ELAPSED (s)                 : {elapsed:.2f}",
                ts_line(),
            ]
            self._enqueue_lines(run_end_lines)
            self.after(0, self._finalize_ui)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
