# header_footer_fixed_with_na_quickparts_restrict_gui_password.py
"""
Purpose:
    Batch-process Word documents to normalize header/footer formatting, clean footer/header
    spacing, optionally export PDFs, optionally rename outputs by Document ID, fill selected
    blank footer-table QuickParts/value cells with N/A, and optionally protect headers/footers
    while leaving the document body editable.

What this file does:
    - Provides a Tkinter GUI for selecting input/output/backup folders.
    - Creates a run-stamped output folder.
    - Opens Word documents through Microsoft Word COM automation.
    - Makes a working copy of each input document before editing.
    - Optionally backs up original files.
    - Clears read-only/final/read-only-recommended state where possible.
    - Normalizes header/footer paragraph spacing.
    - Centers header/footer tables as table objects.
    - Formats header table cell contents as Arial 11, left justified, except GxP cells,
      which are centered.
    - Formats footer table cell contents as Arial 9, center justified.
    - Deletes extra blank paragraphs after header/footer tables.
    - Deletes extra blank paragraphs after the Neurotech footer phrase.
    - Fills blank cells directly below these footer-table labels with "N/A":
        Batch ID
        Printed Time
        Printed Date
        Printed By
        Controlled Print #
    - Optionally formats text boxes/shapes in headers/footers.
    - Optionally extracts Document ID from header tables and uses it for output naming.
    - Optionally protects the output document so the body is editable by Everyone while
      headers/footers remain restricted, using the password entered in the GUI.
    - Periodically restarts Microsoft Word during large batches to prevent COM resource exhaustion.
    - Retries a failed Word open once after restarting Word.
    - Restarts Word after any document-level COM/process failure so one poisoned Word instance
      does not cascade into the remaining documents.
    - Generates a text run log and, when ReportLab is installed, a PDF run report.

Place in the larger scheme:
    This is a GMP document-preparation utility intended to apply consistent header/footer
    formatting, footer default-value handling, and controlled edit restrictions before
    downstream review, use, or distribution.

Why that matters:
    Header/footer consistency, controlled fields, source-document naming, and optional
    protection of controlled footer/header content support document traceability,
    auditability, and review control.
"""

from __future__ import annotations

import getpass
import os
import platform
import queue
import re
import shutil
import socket
import stat
import subprocess
import threading
import time
import traceback
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pythoncom
    import win32com.client as win32

    try:
        import win32process
    except Exception:
        win32process = None
except Exception:
    pythoncom = None
    win32 = None
    win32process = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import LETTER
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.pdfgen import canvas
    from reportlab.platypus import (
        PageBreak,
        Paragraph,
        Preformatted,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )
except Exception:
    colors = None
    LETTER = None
    inch = None
    ParagraphStyle = None
    getSampleStyleSheet = None
    canvas = None
    PageBreak = None
    Paragraph = None
    Preformatted = None
    SimpleDocTemplate = None
    Spacer = None
    Table = None
    TableStyle = None


# -----------------------------
# Word constants
# -----------------------------
wdHeaderFooterPrimary = 1
wdHeaderFooterFirstPage = 2
wdHeaderFooterEvenPages = 3

wdAlignParagraphLeft = 0
wdAlignParagraphCenter = 1
wdLineSpaceSingle = 0
wdAlertsNone = 0
wdAlignTableCenter = 1
wdAlignRowCenter = 1

msoAutomationSecurityForceDisable = 3

wdExportFormatPDF = 17

wdNoProtection = -1
wdAllowOnlyReading = 3
wdEditorEveryone = -1
wdMainTextStory = 1

HF_KINDS = [
    (wdHeaderFooterPrimary, "PRIMARY"),
    (wdHeaderFooterFirstPage, "FIRSTPAGE"),
    (wdHeaderFooterEvenPages, "EVENPAGES"),
]

DEFAULT_PROTECTION_PASSWORD = "colorado26"
CONFIDENTIAL_FOOTER_PHRASE = "Neurotech Pharmaceuticals"

_WS_RE = re.compile(r"\s+", re.UNICODE)
_FILENAME_BAD_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1F]+')
_FILENAME_DOTS_SPACES = re.compile(r"[\. ]+$")


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
    quickparts_na_filled_total: int = 0
    table_cells_formatted_total: int = 0

    restart_interval_docs: int = 100
    word_restarts_attempted: int = 0
    word_restarts_ok: int = 0
    word_restarts_failed: int = 0
    open_retries_attempted: int = 0
    open_retries_ok: int = 0
    open_retries_failed: int = 0
    detailed_pdf_report_enabled: bool = True

    protections_enabled: bool = False
    protection_password: str = ""
    protections_attempted: int = 0
    protections_ok: int = 0
    protections_failed: int = 0

    backups_enabled: bool = False
    backups_attempted: int = 0
    backups_ok: int = 0
    backups_failed: int = 0

    pdf_exports_enabled: bool = False
    pdf_exports_attempted: int = 0
    pdf_exports_ok: int = 0
    pdf_exports_failed: int = 0

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
# General helpers
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
        for filename in files:
            if not is_word_file(filename):
                continue
            abs_path = os.path.join(root, filename)
            rel_path = os.path.relpath(abs_path, input_root)
            yield abs_path, rel_path


def safe_strip_word_text(text: str) -> str:
    if text is None:
        return ""
    text = (
        str(text)
        .replace("\r", "")
        .replace("\x07", "")
        .replace("\xa0", " ")
        .replace("\u200b", "")
        .replace("\v", "")
        .replace("\f", "")
        .replace("\t", "")
    )
    return _WS_RE.sub("", text)


def is_blank_paragraph(paragraph) -> bool:
    try:
        return safe_strip_word_text(paragraph.Range.Text) == ""
    except Exception:
        return False


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


def sanitize_filename_component(name: str, *, max_len: int = 120) -> str:
    if not name:
        return ""

    s = str(name).strip()
    s = s.replace("\r", " ").replace("\n", " ").replace("\x07", " ")
    s = _WS_RE.sub(" ", s).strip()
    s = _FILENAME_BAD_CHARS.sub("_", s)
    s = _FILENAME_DOTS_SPACES.sub("", s)
    s = s.strip()

    if len(s) > max_len:
        s = s[:max_len].rstrip()

    return s


def pick_unique_base_name(out_dir: str, base: str, extensions: list[str]) -> str:
    base = sanitize_filename_component(base)
    if not base:
        base = "Document"

    def conflict(candidate: str) -> bool:
        for ext in extensions:
            if os.path.exists(os.path.join(out_dir, candidate + ext)):
                return True
        return False

    if not conflict(base):
        return base

    for idx in range(2, 10000):
        candidate = f"{base}_{idx}"
        if not conflict(candidate):
            return candidate

    return f"{base}_{int(time.time())}"


def _cell_text(cell) -> str:
    try:
        return str(cell.Range.Text or "")
    except Exception:
        return ""


def _cell_visible_text(cell) -> str:
    try:
        text = str(cell.Range.Text or "")
    except Exception:
        return ""

    text = text.replace("\r", "")
    text = text.replace("\x07", "")
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _normalize_key_text(text: str) -> str:
    text = (text or "").lower()
    text = text.replace("\r", "").replace("\x07", "")
    text = re.sub(r"\s+", " ", text).strip()
    return text


# -----------------------------
# Word formatting helpers
# -----------------------------
def apply_paragraph_format_to_range(rng) -> int:
    """
    General header/footer range normalization.

    This keeps the existing broad behavior. More specific table-cell rules are applied
    later so header cell contents become Arial 11 left, GxP cell centered, and footer
    cell contents become Arial 9 centered.
    """
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

        for idx in range(1, shapes.Count + 1):
            shp = shapes(idx)

            try:
                has_text = bool(shp.TextFrame.HasText)
            except Exception:
                has_text = False

            if not has_text:
                continue

            text_shapes += 1

            try:
                formatted += apply_paragraph_format_to_range(shp.TextFrame.TextRange)
            except Exception:
                continue

    except Exception:
        shapes_count = 0
        text_shapes = 0
        formatted = 0

    return shapes_count, text_shapes, formatted


def center_tables_in_range(rng) -> tuple[int, int]:
    """
    Center the table object/rows in the header/footer range.

    This does not determine the final alignment of text inside table cells. Header/footer
    cell-content alignment is enforced separately by format_header_table_cells() and
    format_footer_table_cells().
    """
    tables_found = 0
    centered = 0

    try:
        tables_found = int(rng.Tables.Count)

        for idx in range(1, tables_found + 1):
            tbl = rng.Tables(idx)

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


def _apply_font_and_alignment_to_cell(cell, *, font_name: str, font_size: int, alignment: int) -> bool:
    """
    Format only the contents of a table cell.

    This does not left-align the table object itself. It changes the font and paragraph
    alignment inside the cell range.
    """
    try:
        rng = cell.Range

        try:
            rng.Font.Name = font_name
        except Exception:
            pass

        try:
            rng.Font.NameAscii = font_name
        except Exception:
            pass

        try:
            rng.Font.NameOther = font_name
        except Exception:
            pass

        try:
            rng.Font.Size = font_size
        except Exception:
            pass

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
        pf.Alignment = alignment

        return True
    except Exception:
        return False


def format_header_table_cells(hf_range) -> int:
    """
    Header table cell-content requirement:
      - Arial 11
      - Left justified cell contents
      - Exception: the cell whose visible text is GxP is center justified

    This does not left-align the header table object itself.
    """
    formatted = 0

    try:
        table_count = int(hf_range.Tables.Count)
    except Exception:
        return 0

    for table_idx in range(1, table_count + 1):
        try:
            tbl = hf_range.Tables(table_idx)
            rows = int(tbl.Rows.Count)
            cols = int(tbl.Columns.Count)
        except Exception:
            continue

        for row_idx in range(1, rows + 1):
            for col_idx in range(1, cols + 1):
                try:
                    cell = tbl.Cell(row_idx, col_idx)
                    visible = _cell_visible_text(cell)
                except Exception:
                    continue

                alignment = wdAlignParagraphCenter if visible.strip().lower() == "gxp" else wdAlignParagraphLeft

                if _apply_font_and_alignment_to_cell(
                    cell,
                    font_name="Arial",
                    font_size=11,
                    alignment=alignment,
                ):
                    formatted += 1

    return formatted


def format_footer_table_cells(hf_range) -> int:
    """
    Footer table cell-content requirement:
      - Arial 9
      - Center justified cell contents
    """
    formatted = 0

    try:
        table_count = int(hf_range.Tables.Count)
    except Exception:
        return 0

    for table_idx in range(1, table_count + 1):
        try:
            tbl = hf_range.Tables(table_idx)
            rows = int(tbl.Rows.Count)
            cols = int(tbl.Columns.Count)
        except Exception:
            continue

        for row_idx in range(1, rows + 1):
            for col_idx in range(1, cols + 1):
                try:
                    cell = tbl.Cell(row_idx, col_idx)
                except Exception:
                    continue

                if _apply_font_and_alignment_to_cell(
                    cell,
                    font_name="Arial",
                    font_size=9,
                    alignment=wdAlignParagraphCenter,
                ):
                    formatted += 1

    return formatted


def range_is_empty(hf_range, table_count: int, text_shapes_count: int) -> bool:
    if table_count > 0 or text_shapes_count > 0:
        return False

    try:
        text = hf_range.Text or ""
    except Exception:
        return True

    return safe_strip_word_text(text) == ""


# -----------------------------
# Document editability and protection
# -----------------------------
def force_document_editable(doc, password_hint: str = "") -> dict:
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

    password_attempts = []
    if password_hint:
        password_attempts.append(password_hint)
    password_attempts.extend(["", DEFAULT_PROTECTION_PASSWORD])

    seen = set()
    password_attempts = [p for p in password_attempts if not (p in seen or seen.add(p))]

    for password in password_attempts:
        try:
            doc.Unprotect(password)
            result["unprotect_success"] = "YES"
            break
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
        doc.ChangeFileAccess(2)
        result["changefileaccess_success"] = "YES"
    except Exception:
        result["changefileaccess_success"] = "NO"

    try:
        result["doc_readonly_final"] = "YES" if bool(doc.ReadOnly) else "NO"
    except Exception:
        pass

    return result


def restrict_headers_footers_by_body_exception(
    doc,
    *,
    password: str,
) -> dict:
    result = {
        "attempted": "YES",
        "password_used": password,
        "unprotect_before_success": "NO",
        "body_exception_success": "NO",
        "protect_success": "NO",
        "protection_type_after": "UNKNOWN",
        "error": "",
    }

    if not password:
        result["error"] = "Password is blank."
        return result

    try:
        try:
            if int(doc.ProtectionType) != wdNoProtection:
                doc.Unprotect(password)
                result["unprotect_before_success"] = "YES"
            else:
                result["unprotect_before_success"] = "N/A"
        except Exception:
            result["unprotect_before_success"] = "NO"

        try:
            body_range = doc.StoryRanges(wdMainTextStory)

            try:
                body_range.Editors.DeleteAll()
            except Exception:
                pass

            body_range.Editors.Add(wdEditorEveryone)
            result["body_exception_success"] = "YES"
        except Exception as exc:
            result["body_exception_success"] = "NO"
            result["error"] = f"Body exception failed: {exc}"
            return result

        try:
            doc.Protect(
                Type=wdAllowOnlyReading,
                NoReset=True,
                Password=password,
                UseIRM=False,
                EnforceStyleLock=False,
            )
            result["protect_success"] = "YES"
        except TypeError:
            try:
                doc.Protect(wdAllowOnlyReading, True, password)
                result["protect_success"] = "YES"
            except Exception as exc:
                result["protect_success"] = "NO"
                result["error"] = f"Protect failed: {exc}"
        except Exception as exc:
            result["protect_success"] = "NO"
            result["error"] = f"Protect failed: {exc}"

        try:
            result["protection_type_after"] = str(int(doc.ProtectionType))
        except Exception:
            result["protection_type_after"] = "UNKNOWN"

        return result

    except Exception as exc:
        result["error"] = str(exc)
        return result


# -----------------------------
# Document ID helper
# -----------------------------
def extract_document_id_from_headers(doc) -> str | None:
    target = "document id"

    try:
        sec_count = int(doc.Sections.Count)
    except Exception:
        return None

    for section_idx in range(1, sec_count + 1):
        sec = doc.Sections(section_idx)

        for kind, _kind_name in HF_KINDS:
            try:
                hdr = sec.Headers(kind)
                rng = hdr.Range
                tbl_count = int(rng.Tables.Count)
            except Exception:
                continue

            for table_idx in range(1, tbl_count + 1):
                try:
                    tbl = rng.Tables(table_idx)
                    rows = int(tbl.Rows.Count)
                    cols = int(tbl.Columns.Count)
                except Exception:
                    continue

                for row_idx in range(1, rows + 1):
                    for col_idx in range(1, cols + 1):
                        try:
                            cell = tbl.Cell(row_idx, col_idx)
                            key_text = _normalize_key_text(_cell_text(cell))
                        except Exception:
                            continue

                        if target != key_text and target not in key_text:
                            continue

                        value = ""

                        if col_idx < cols:
                            try:
                                value = _cell_text(tbl.Cell(row_idx, col_idx + 1))
                            except Exception:
                                value = ""

                        if not sanitize_filename_component(safe_strip_word_text(value)):
                            try:
                                row = tbl.Rows(row_idx)
                                row_cell_count = int(row.Cells.Count)
                                this_start = int(cell.Range.Start)
                                best_idx = None

                                for idx in range(1, row_cell_count + 1):
                                    cc = row.Cells(idx)
                                    if int(cc.Range.Start) == this_start:
                                        best_idx = idx
                                        break

                                if best_idx is not None and best_idx < row_cell_count:
                                    value = _cell_text(row.Cells(best_idx + 1))
                            except Exception:
                                pass

                        candidate = value or ""
                        candidate = candidate.replace("\r", "").replace("\x07", "")
                        candidate = re.sub(r"\s+", " ", candidate).strip()
                        candidate = sanitize_filename_component(candidate)

                        if candidate:
                            return candidate

    return None


# -----------------------------
# Footer QuickParts N/A default helpers
# -----------------------------
_QUICKPART_NA_TARGETS = {
    "batch id",
    "printed time",
    "printed date",
    "printed by",
    "controlled print",
    "controlled print number",
}


def _normalize_quickpart_label(text: str) -> str:
    text = (text or "").lower()
    text = text.replace("\r", " ").replace("\x07", " ").replace("#", " number ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _quickpart_label_matches(cell_text: str) -> bool:
    label = _normalize_quickpart_label(cell_text)

    if not label:
        return False

    if label in _QUICKPART_NA_TARGETS:
        return True

    return any(label.startswith(target + " ") for target in _QUICKPART_NA_TARGETS)


def _cell_display_is_blank(cell) -> bool:
    try:
        return safe_strip_word_text(cell.Range.Text or "") == ""
    except Exception:
        return False


def _set_cell_text(cell, value: str) -> bool:
    try:
        cell.Range.Text = value
        return True
    except Exception:
        return False


def _cell_below(tbl, row_idx: int, col_idx: int):
    try:
        rows = int(tbl.Rows.Count)
        if row_idx < rows:
            return tbl.Cell(row_idx + 1, col_idx)
    except Exception:
        pass

    return None


def fill_footer_quickparts_defaults_in_range(hf_range, default_value: str = "N/A") -> int:
    """
    Fill blank footer-table value cells below selected QuickPart labels with N/A.

    Expected final footer layout:

        Controlled Print # | Printed Date | Printed Time | Printed By | Batch ID
        N/A                | N/A          | N/A          | N/A        | N/A

    Existing nonblank values are preserved.
    """
    filled = 0

    try:
        table_count = int(hf_range.Tables.Count)
    except Exception:
        return 0

    for table_idx in range(1, table_count + 1):
        try:
            tbl = hf_range.Tables(table_idx)
            rows = int(tbl.Rows.Count)
            cols = int(tbl.Columns.Count)
        except Exception:
            continue

        for row_idx in range(1, rows + 1):
            for col_idx in range(1, cols + 1):
                try:
                    label_cell = tbl.Cell(row_idx, col_idx)
                    label_text = _cell_text(label_cell)
                except Exception:
                    continue

                if not _quickpart_label_matches(label_text):
                    continue

                value_cell = _cell_below(tbl, row_idx, col_idx)

                if value_cell is None:
                    continue

                if _cell_display_is_blank(value_cell):
                    if _set_cell_text(value_cell, default_value):
                        filled += 1

    return filled


# -----------------------------
# Blank cleanup helpers
# -----------------------------
def delete_blank_paragraph_block(paras, start_idx: int, end_idx: int) -> int:
    if end_idx < start_idx:
        return 0

    try:
        p_start = paras(start_idx)
        p_end = paras(end_idx)

        delete_range = p_start.Range.Duplicate
        delete_range.End = p_end.Range.End
        delete_range.Delete()

        return end_idx - start_idx + 1
    except Exception:
        return 0


def _first_paragraph_index_at_or_after(paras, start_pos: int, max_scan: int = 500) -> int | None:
    try:
        count = int(paras.Count)
    except Exception:
        return None

    for idx in range(1, min(count, max_scan) + 1):
        try:
            if int(paras(idx).Range.Start) >= start_pos:
                return idx
        except Exception:
            continue

    return None


def enforce_max_blank_paragraphs_after_table_safe(hf_range, table, max_blanks: int) -> tuple[int, int]:
    try:
        container_start = int(table.Range.Start)
        container_end = int(hf_range.End)

        if container_start >= container_end:
            return 0, 0

        container_range = hf_range.Duplicate
        container_range.Start = container_start
        container_range.End = container_end
        paras = container_range.Paragraphs
        table_end = int(table.Range.End)

        first_idx = _first_paragraph_index_at_or_after(paras, table_end, max_scan=500)

        if first_idx is None:
            return 0, 0

        try:
            count = int(paras.Count)
        except Exception:
            count = 0

        blank_found = 0
        scan_limit = min(count, first_idx + 50 - 1)

        for idx in range(first_idx, scan_limit + 1):
            try:
                p = paras(idx)
            except Exception:
                break

            if is_blank_paragraph(p):
                blank_found += 1
            else:
                break

        if blank_found <= max_blanks:
            return blank_found, 0

        delete_start = first_idx + max_blanks
        delete_end = first_idx + blank_found - 1
        deleted = delete_blank_paragraph_block(paras, delete_start, delete_end)

        return blank_found, deleted

    except Exception:
        return 0, 0


def enforce_max_blank_paragraphs_after_phrase_safe(
    hf_range,
    phrase: str,
    max_blanks: int = 0,
) -> tuple[int, int]:
    try:
        phrase_norm = (phrase or "").strip().lower()
        if not phrase_norm:
            return 0, 0

        paras = hf_range.Paragraphs

        try:
            count = int(paras.Count)
        except Exception:
            return 0, 0

        target_idx = None

        for idx in range(1, min(count, 500) + 1):
            try:
                text = paras(idx).Range.Text or ""
            except Exception:
                continue

            if phrase_norm in text.lower():
                target_idx = idx
                break

        if target_idx is None:
            return 0, 0

        first_blank_idx = target_idx + 1
        blank_found = 0
        scan_limit = min(count, first_blank_idx + 50 - 1)

        for idx in range(first_blank_idx, scan_limit + 1):
            try:
                p = paras(idx)
            except Exception:
                break

            if is_blank_paragraph(p):
                blank_found += 1
            else:
                break

        if blank_found <= max_blanks:
            return blank_found, 0

        delete_start = first_blank_idx + max_blanks
        delete_end = first_blank_idx + blank_found - 1
        deleted = delete_blank_paragraph_block(paras, delete_start, delete_end)

        return blank_found, deleted

    except Exception:
        return 0, 0


# -----------------------------
# Header/footer processing
# -----------------------------
def process_headerfooter(
    hf,
    mode: str,
    *,
    substep: Optional[Callable[[str], None]] = None,
    enable_shapes: bool = False,
    enable_table_centering: bool = True,
    enable_blank_cleanup: bool = True,
) -> tuple[dict, int, dict]:
    actions = 0

    agg = {
        "format_apps": 0,
        "post_table_deleted": 0,
        "tables_centered": 0,
        "text_shapes_formatted": 0,
        "quickparts_na_filled": 0,
        "table_cells_formatted": 0,
    }

    res = {
        "present": "YES",
        "tables_found": "NO",
        "table_count": 0,
        "anchor": "N/A",
        "format_applied": "NO",
        "spec_max": 0,
        "tables_centering_applied": "NO",
        "tables_centered_count": 0,
        "table_cells_formatted": 0,
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
        "quickparts_na_filled": 0,
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
            table_count = int(hf.Range.Tables.Count)
        except Exception:
            table_count = 0

        res["table_count"] = table_count
        res["tables_found"] = "YES" if table_count > 0 else "NO"
        res["anchor"] = "LAST_TABLE" if table_count > 0 else "N/A"

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

            tables_found, tables_centered = center_tables_in_range(hf.Range)

            if tables_found > 0:
                res["tables_centering_applied"] = "YES"
                res["tables_centered_count"] = int(tables_centered)
                agg["tables_centered"] += int(tables_centered)
        else:
            res["tables_centering_applied"] = "SKIPPED"
            res["tables_centered_count"] = 0

        if mode == "header":
            if substep:
                substep("header_table_cell_font_alignment")

            formatted = format_header_table_cells(hf.Range)
            res["table_cells_formatted"] += int(formatted)
            agg["table_cells_formatted"] += int(formatted)

        if mode == "footer":
            if substep:
                substep("footer_table_cell_font_alignment_before_defaults")

            formatted_before = format_footer_table_cells(hf.Range)
            res["table_cells_formatted"] += int(formatted_before)
            agg["table_cells_formatted"] += int(formatted_before)

            if substep:
                substep("quickparts_na_defaults_below_labels")

            filled = fill_footer_quickparts_defaults_in_range(hf.Range, default_value="N/A")
            res["quickparts_na_filled"] = int(filled)
            agg["quickparts_na_filled"] += int(filled)
            actions += int(filled)

            if substep:
                substep("footer_table_cell_font_alignment_after_defaults")

            formatted_after = format_footer_table_cells(hf.Range)
            res["table_cells_formatted"] += int(formatted_after)
            agg["table_cells_formatted"] += int(formatted_after)

        if substep:
            substep("range_empty_check")

        text_shapes_for_empty = 0 if not enable_shapes else int(res.get("text_shapes_count", 0))
        res["range_empty"] = "YES" if range_is_empty(hf.Range, table_count, text_shapes_for_empty) else "NO"

        if enable_blank_cleanup and table_count > 0:
            if substep:
                substep("blank_cleanup")

            last_table = hf.Range.Tables(table_count)
            blank_count, deleted = enforce_max_blank_paragraphs_after_table_safe(
                hf.Range,
                last_table,
                max_blanks=res["spec_max"],
            )

            res["post_table_blanks_found"] = "YES" if blank_count > 0 else "NO"
            res["post_table_blank_count"] = blank_count
            res["post_table_deleted"] = deleted
            res["post_table_deletions_required"] = "YES" if blank_count > res["spec_max"] else "NO"
            res["post_table_remaining"] = max(blank_count - deleted, 0)

            actions += int(deleted)
            agg["post_table_deleted"] += int(deleted)

            if mode == "footer":
                _, deleted_after_phrase = enforce_max_blank_paragraphs_after_phrase_safe(
                    hf.Range,
                    CONFIDENTIAL_FOOTER_PHRASE,
                    max_blanks=0,
                )
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
        "quickparts_na_filled": 0,
        "table_cells_formatted": 0,
    }

    actions_total = 0

    sec_count = int(doc.Sections.Count)
    details["sections"] = sec_count

    zero_agg = {
        "format_apps": 0,
        "post_table_deleted": 0,
        "tables_centered": 0,
        "text_shapes_formatted": 0,
        "quickparts_na_filled": 0,
        "table_cells_formatted": 0,
    }

    for section_idx in range(1, sec_count + 1):
        sec = doc.Sections(section_idx)
        sec_block = {"section_index": section_idx, "headers": {}, "footers": {}}

        for kind, kind_name in HF_KINDS:
            try:
                hdr = sec.Headers(kind)

                def header_substep(step: str, s=section_idx, k=kind_name):
                    if substep:
                        substep(f"S{s} HEADER {k} | {step}")

                hdr_res, hdr_actions, hdr_agg = process_headerfooter(
                    hdr,
                    mode="header",
                    substep=header_substep,
                    enable_shapes=enable_shapes,
                    enable_table_centering=enable_table_centering,
                    enable_blank_cleanup=enable_blank_cleanup,
                )
            except Exception:
                hdr_res, hdr_actions, hdr_agg = {"present": "NO"}, 0, dict(zero_agg)

            sec_block["headers"][kind_name] = hdr_res

            actions_total += hdr_actions
            aggregates["format_apps"] += int(hdr_agg.get("format_apps", 0))
            aggregates["header_post_table_deleted"] += int(hdr_agg.get("post_table_deleted", 0))
            aggregates["tables_centered"] += int(hdr_agg.get("tables_centered", 0))
            aggregates["text_shapes_formatted"] += int(hdr_agg.get("text_shapes_formatted", 0))
            aggregates["quickparts_na_filled"] += int(hdr_agg.get("quickparts_na_filled", 0))
            aggregates["table_cells_formatted"] += int(hdr_agg.get("table_cells_formatted", 0))

            try:
                ftr = sec.Footers(kind)

                def footer_substep(step: str, s=section_idx, k=kind_name):
                    if substep:
                        substep(f"S{s} FOOTER {k} | {step}")

                ftr_res, ftr_actions, ftr_agg = process_headerfooter(
                    ftr,
                    mode="footer",
                    substep=footer_substep,
                    enable_shapes=enable_shapes,
                    enable_table_centering=enable_table_centering,
                    enable_blank_cleanup=enable_blank_cleanup,
                )
            except Exception:
                ftr_res, ftr_actions, ftr_agg = {"present": "NO"}, 0, dict(zero_agg)

            sec_block["footers"][kind_name] = ftr_res

            actions_total += ftr_actions
            aggregates["format_apps"] += int(ftr_agg.get("format_apps", 0))
            aggregates["footer_post_table_deleted"] += int(ftr_agg.get("post_table_deleted", 0))
            aggregates["tables_centered"] += int(ftr_agg.get("tables_centered", 0))
            aggregates["text_shapes_formatted"] += int(ftr_agg.get("text_shapes_formatted", 0))
            aggregates["quickparts_na_filled"] += int(ftr_agg.get("quickparts_na_filled", 0))
            aggregates["table_cells_formatted"] += int(ftr_agg.get("table_cells_formatted", 0))

        details["sections_data"].append(sec_block)

    return details, actions_total, aggregates


# -----------------------------
# Export helpers
# -----------------------------
def export_doc_to_pdf(doc, pdf_path: str) -> None:
    ensure_dir(os.path.dirname(pdf_path))
    doc.ExportAsFixedFormat(
        OutputFileName=pdf_path,
        ExportFormat=wdExportFormatPDF,
        OpenAfterExport=False,
    )


# -----------------------------
# Report generation
# -----------------------------
if canvas is not None:
    class NumberedCanvas(canvas.Canvas):
        def __init__(self, *args, run_id: str, left_margin: float, right_margin: float, **kwargs):
            super().__init__(*args, **kwargs)
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
            width, height = self._pagesize
            y = height - 0.50 * inch
            self.saveState()
            self.setFont("Helvetica", 9)
            self.drawString(self._left_margin, y, f"Run Report #: {self._run_id}")
            self.drawRightString(width - self._right_margin, y, f"Page {self.getPageNumber()} of {page_count}")
            self.restoreState()
else:
    NumberedCanvas = None


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
    wrapped = []

    for line in lines:
        wrapped.extend(_wrap_line(line, max_chars=max_chars))

    return wrapped


def generate_pdf_report(
    pdf_path: str,
    run_id: str,
    stats: RunStats,
    input_root: str,
    output_root: str,
    backups_enabled: bool,
    backup_root: str,
    error_rows: list[tuple[str, str, str]],
    doc_blocks: list[list[str]],
    include_detail: bool = True,
) -> None:
    if SimpleDocTemplate is None or canvas is None or NumberedCanvas is None:
        raise RuntimeError("ReportLab is required for PDF run-report generation. Install it with: pip install reportlab")

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

    story = []
    story.append(Paragraph("Header/Footer Formatting Run Report", title))
    story.append(Paragraph(f"Run ID: {run_id}", body))
    story.append(Spacer(1, 10))

    story.append(Paragraph("Run Metadata", h2))

    start_local = stats.start_dt
    end_local = stats.end_dt or datetime.now().astimezone()
    tz_text = f"{stats.tz_name} ({stats.tz_offset})" if stats.tz_name or stats.tz_offset else "UNKNOWN"
    elapsed = time.strftime("%H:%M:%S", time.gmtime(max(0, int(end_local.timestamp() - start_local.timestamp()))))

    meta_rows = [
        ["Run Start", start_local.strftime("%d-%b-%Y %H:%M:%S")],
        ["Run End", end_local.strftime("%d-%b-%Y %H:%M:%S")],
        ["Timezone", tz_text],
        ["Elapsed", elapsed],
        ["User", stats.user],
        ["Computer", stats.computer],
        ["OS", stats.os_info],
        ["Python", stats.python_info],
        ["MS Word", stats.word_version],
        ["Input Root", os.path.abspath(input_root)],
        ["Output Root", os.path.abspath(output_root)],
        ["Backups Enabled", "YES" if backups_enabled else "NO"],
        ["Backup Root", os.path.abspath(backup_root) if backups_enabled else "N/A"],
        ["PDF Export Enabled", "YES" if stats.pdf_exports_enabled else "NO"],
        ["DocID Naming Enabled", "YES" if stats.docid_naming_enabled else "NO"],
        ["Restrict Headers/Footers Enabled", "YES" if stats.protections_enabled else "NO"],
        ["Restrict Headers/Footers Password", stats.protection_password if stats.protections_enabled else "N/A"],
        ["Header Table Cell Format", "Arial 11; cell contents left except GxP centered"],
        ["Footer Table Cell Format", "Arial 9; cell contents centered"],
        ["Footer N/A Default Layout", "Values inserted in cells below matching labels"],
        ["Restart Word Every N Docs", str(stats.restart_interval_docs)],
        ["Open Retry After Word Restart", "YES"],
        ["Detailed PDF Report", "YES" if stats.detailed_pdf_report_enabled else "NO"],
    ]

    meta_tbl = Table(meta_rows, colWidths=[2.4 * inch, 5.1 * inch])
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
        ["Actions total", str(stats.total_actions)],
        ["Header post-table deletions", str(stats.header_post_table_deleted_total)],
        ["Footer post-table deletions", str(stats.footer_post_table_deleted_total)],
        ["Tables centered", str(stats.tables_centered_total)],
        ["Table cells formatted", str(stats.table_cells_formatted_total)],
        ["Word restarts attempted", str(stats.word_restarts_attempted)],
        ["Word restarts succeeded", str(stats.word_restarts_ok)],
        ["Word restarts failed", str(stats.word_restarts_failed)],
        ["Open retries attempted", str(stats.open_retries_attempted)],
        ["Open retries succeeded", str(stats.open_retries_ok)],
        ["Open retries failed", str(stats.open_retries_failed)],
        ["Text shapes formatted", str(stats.text_shapes_formatted_total)],
        ["Format applications", str(stats.format_apps_total)],
        ["QuickParts N/A defaults filled", str(stats.quickparts_na_filled_total)],
        ["Protections enabled", "YES" if stats.protections_enabled else "NO"],
        ["Protection password", stats.protection_password if stats.protections_enabled else "N/A"],
        ["Protections attempted", str(stats.protections_attempted)],
        ["Protections succeeded", str(stats.protections_ok)],
        ["Protections failed", str(stats.protections_failed)],
        ["Backup attempts", str(stats.backups_attempted)],
        ["Backup succeeded", str(stats.backups_ok)],
        ["Backup failed", str(stats.backups_failed)],
        ["PDF exports attempted", str(stats.pdf_exports_attempted)],
        ["PDF exports succeeded", str(stats.pdf_exports_ok)],
        ["PDF exports failed", str(stats.pdf_exports_failed)],
        ["DocID found", str(stats.docid_found_count)],
        ["DocID used for naming", str(stats.docid_used_count)],
        ["DocID missing", str(stats.docid_missing_count)],
    ]

    summary_tbl = Table(summary_rows, colWidths=[3.1 * inch, 4.4 * inch])
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
        err_rows = [["File", "Stage", "Error Summary"]] + error_rows
        err_tbl = Table(err_rows, colWidths=[2.2 * inch, 1.2 * inch, 4.1 * inch])
        err_tbl.setStyle(TableStyle([
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 9),
            ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ]))
        story.append(err_tbl)
    else:
        story.append(Paragraph("No errors recorded.", body))

    if include_detail:
        story.append(PageBreak())
        story.append(Paragraph("Detailed Event Log", h2))

        for block in doc_blocks:
            story.append(Preformatted("\n".join(_wrap_lines(block, 112)), mono))
            story.append(Spacer(1, 10))
    else:
        story.append(Spacer(1, 12))
        story.append(Paragraph("Detailed Event Log", h2))
        story.append(Paragraph("Detailed per-document report content was disabled for this run. The complete text log was still written to disk.", body))

    def canvas_maker(filename, **kwargs):
        return NumberedCanvas(
            filename,
            run_id=run_id,
            left_margin=doc.leftMargin,
            right_margin=doc.rightMargin,
            **kwargs,
        )

    doc.build(story, canvasmaker=canvas_maker)


# -----------------------------
# GUI application
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Header/Footer Formatter + N/A Defaults + Optional Header/Footer Restriction")
        self.geometry("1260x900")
        self.minsize(1120, 820)

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

    def _emit(self, line: str):
        self._log_q.put(line)

    def _enqueue_lines(self, lines: list[str]):
        for line in lines:
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
        ttk.Entry(frm, textvariable=self.in_var, width=110).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm, text="Browse…", command=self._browse_in).grid(row=0, column=2)

        ttk.Label(frm, text="Output folder (run-stamped subfolder created):").grid(row=1, column=0, sticky="w")
        self.out_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.out_var, width=110).grid(row=1, column=1, sticky="we", padx=6)
        ttk.Button(frm, text="Browse…", command=self._browse_out).grid(row=1, column=2)

        ttk.Label(frm, text="Backup folder (optional):").grid(row=2, column=0, sticky="w")
        self.backup_var = tk.StringVar()
        self.backup_entry = ttk.Entry(frm, textvariable=self.backup_var, width=110)
        self.backup_entry.grid(row=2, column=1, sticky="we", padx=6)
        self.backup_btn = ttk.Button(frm, text="Browse…", command=self._browse_backup)
        self.backup_btn.grid(row=2, column=2)

        self.backups_enabled_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frm,
            text="Enable backups",
            variable=self.backups_enabled_var,
            command=self._toggle_backup_widgets,
        ).grid(row=3, column=1, sticky="w", pady=(2, 0))

        frm.columnconfigure(1, weight=1)

        opt_frame = ttk.LabelFrame(self, text="Options")
        opt_frame.pack(fill="x", padx=10, pady=(4, 6))

        self.show_word_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame,
            text="Show Word (debug)",
            variable=self.show_word_var,
        ).grid(row=0, column=0, sticky="w", padx=10, pady=4)

        ttk.Label(opt_frame, text="Timeout (sec):").grid(row=0, column=1, sticky="e", padx=(20, 6), pady=4)
        self.timeout_var = tk.StringVar(value="180")
        ttk.Entry(opt_frame, textvariable=self.timeout_var, width=8).grid(row=0, column=2, sticky="w", padx=(0, 20), pady=4)

        self.process_shapes_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame,
            text="Process Shapes/TextBoxes",
            variable=self.process_shapes_var,
        ).grid(row=0, column=3, sticky="w", padx=10, pady=4)

        self.export_pdf_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame,
            text="Export PDF for each processed document",
            variable=self.export_pdf_var,
        ).grid(row=0, column=4, sticky="w", padx=10, pady=4)

        self.name_by_docid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame,
            text="Name outputs using Document ID",
            variable=self.name_by_docid_var,
        ).grid(row=1, column=0, sticky="w", padx=10, pady=4)

        self.protect_headers_footers_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame,
            text="Restrict Headers/Footers",
            variable=self.protect_headers_footers_var,
            command=self._toggle_password_widgets,
        ).grid(row=1, column=1, columnspan=2, sticky="w", padx=10, pady=4)

        ttk.Label(opt_frame, text="Restriction password:").grid(row=1, column=3, sticky="e", padx=(20, 6), pady=4)
        self.protection_password_var = tk.StringVar(value=DEFAULT_PROTECTION_PASSWORD)
        self.protection_password_entry = ttk.Entry(
            opt_frame,
            textvariable=self.protection_password_var,
            width=24,
            show="",
        )
        self.protection_password_entry.grid(row=1, column=4, sticky="w", padx=10, pady=4)

        ttk.Label(opt_frame, text="Restart Word every N docs:").grid(row=2, column=0, sticky="e", padx=(10, 6), pady=4)
        self.restart_interval_var = tk.StringVar(value="100")
        ttk.Entry(opt_frame, textvariable=self.restart_interval_var, width=8).grid(row=2, column=1, sticky="w", padx=10, pady=4)

        self.detailed_pdf_report_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opt_frame,
            text="Include detailed per-document PDF report",
            variable=self.detailed_pdf_report_var,
        ).grid(row=2, column=3, columnspan=2, sticky="w", padx=10, pady=4)

        for col in range(5):
            opt_frame.columnconfigure(col, weight=0)

        opt_frame.columnconfigure(5, weight=1)

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
        self._toggle_password_widgets()

    def _toggle_backup_widgets(self):
        enabled = self.backups_enabled_var.get()
        state = "normal" if enabled else "disabled"
        self.backup_entry.configure(state=state)
        self.backup_btn.configure(state=state)

    def _toggle_password_widgets(self):
        enabled = self.protect_headers_footers_var.get()
        state = "normal" if enabled else "disabled"
        self.protection_password_entry.configure(state=state)

    def _browse_in(self):
        path = filedialog.askdirectory(title="Select input folder")
        if path:
            self.in_var.set(path)

    def _browse_out(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.out_var.set(path)

    def _browse_backup(self):
        path = filedialog.askdirectory(title="Select backup folder")
        if path:
            self.backup_var.set(path)

    def _update_stats_line(self):
        if not self._stats:
            return

        stats = self._stats
        elapsed = time.time() - stats.start_ts

        extra = ""

        if stats.backups_enabled:
            extra += f" | Backups: {stats.backups_ok}/{stats.backups_attempted} OK ({stats.backups_failed} fail)"

        if stats.pdf_exports_enabled:
            extra += f" | PDFs: {stats.pdf_exports_ok}/{stats.pdf_exports_attempted} OK ({stats.pdf_exports_failed} fail)"

        if stats.docid_naming_enabled:
            extra += f" | DocID used: {stats.docid_used_count}/{stats.total_files}"

        if stats.protections_enabled:
            extra += f" | Protected: {stats.protections_ok}/{stats.protections_attempted}"

        self.stats_label.config(
            text=(
                f"Total: {stats.total_files} | Processed: {stats.processed} | "
                f"OK: {stats.succeeded} | Failed: {stats.failed} | Skipped: {stats.skipped} | "
                f"Actions: {stats.total_actions} | QuickParts N/A: {stats.quickparts_na_filled_total} | "
                f"Table cells formatted: {stats.table_cells_formatted_total} | "
                f"Elapsed: {elapsed:.1f}s{extra}"
            )
        )

    def _start(self):
        if win32 is None or pythoncom is None:
            messagebox.showerror("Missing dependency", "pywin32 is required on Windows with Microsoft Word installed.")
            return

        if SimpleDocTemplate is None:
            messagebox.showerror(
                "Missing dependency",
                "ReportLab is required for the PDF run report. Install it with: python -m pip install reportlab",
            )
            return

        in_dir = self.in_var.get().strip()
        out_dir = self.out_var.get().strip()
        backups_enabled = bool(self.backups_enabled_var.get())
        backup_root = self.backup_var.get().strip()

        export_pdf_each = bool(self.export_pdf_var.get())
        name_by_docid = bool(self.name_by_docid_var.get())
        protect_headers_footers = bool(self.protect_headers_footers_var.get())
        protection_password = self.protection_password_var.get().strip()
        detailed_pdf_report_enabled = bool(self.detailed_pdf_report_var.get())

        try:
            restart_interval_docs = int(self.restart_interval_var.get().strip())
            if restart_interval_docs < 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Invalid restart interval", "Restart Word every N docs must be an integer >= 0. Use 0 to disable periodic restarts.")
            return

        if not in_dir or not os.path.isdir(in_dir):
            messagebox.showerror("Invalid input", "Please select a valid input folder.")
            return

        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showerror("Invalid output", "Please select a valid output folder.")
            return

        if backups_enabled and (not backup_root or not os.path.isdir(backup_root)):
            messagebox.showerror("Invalid backup folder", "Backups are enabled. Please select a valid backup folder.")
            return

        if protect_headers_footers and not protection_password:
            messagebox.showerror(
                "Missing restriction password",
                "Restrict Headers/Footers is checked. Enter a nonblank password before starting.",
            )
            return

        try:
            timeout_s = int(self.timeout_var.get().strip())
            if timeout_s < 10:
                raise ValueError
        except Exception:
            messagebox.showerror("Invalid timeout", "Timeout must be an integer >= 10.")
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
        user = f"{os.environ.get('USERDOMAIN', '')}\\{getpass.getuser()}" if os.environ.get("USERDOMAIN") else getpass.getuser()

        self._stats = RunStats(
            run_id=run_id,
            start_ts=time.time(),
            start_dt=now_local,
            backups_enabled=backups_enabled,
            pdf_exports_enabled=export_pdf_each,
            docid_naming_enabled=name_by_docid,
            restart_interval_docs=restart_interval_docs,
            detailed_pdf_report_enabled=detailed_pdf_report_enabled,
            protections_enabled=protect_headers_footers,
            protection_password=protection_password if protect_headers_footers else "",
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
                in_dir,
                out_dir,
                run_id,
                backups_enabled,
                backup_root,
                timeout_s,
                bool(self.show_word_var.get()),
                bool(self.process_shapes_var.get()),
                export_pdf_each,
                name_by_docid,
                protect_headers_footers,
                protection_password,
                restart_interval_docs,
                detailed_pdf_report_enabled,
            ),
            daemon=True,
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
        name_by_docid: bool,
        protect_headers_footers: bool,
        protection_password: str,
        restart_interval_docs: int,
        detailed_pdf_report_enabled: bool,
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
            f"RESTRICT HEADERS/FOOTERS    : {'YES' if protect_headers_footers else 'NO'}",
            f"RESTRICTION PASSWORD        : {protection_password if protect_headers_footers else 'N/A'}",
            f"HEADER CELL FORMAT          : Arial 11; contents left except GxP centered",
            f"FOOTER CELL FORMAT          : Arial 9; contents centered",
            f"FOOTER N/A DEFAULT LAYOUT   : Values inserted below matching labels",
            f"RESTART WORD EVERY N DOCS   : {restart_interval_docs}",
            f"OPEN RETRY AFTER RESTART    : YES",
            f"DETAILED PDF REPORT         : {'YES' if detailed_pdf_report_enabled else 'NO'}",
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

        heartbeat_stop = threading.Event()

        def heartbeat_and_watchdog():
            while not heartbeat_stop.is_set():
                time.sleep(15)
                file_, step, since, pid = self._get_status_snapshot()

                if file_ or step:
                    age = int(time.time() - since)
                    self._emit(f"[HEARTBEAT] still running | step={step} | file={file_} | seconds_in_step={age}")

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

        def restart_word(reason: str) -> None:
            nonlocal word
            stats.word_restarts_attempted += 1
            self._emit(f"STEP: Restart Word (start) | reason={reason}")

            try:
                if word is not None:
                    stop_word(word)
            except Exception:
                pass

            word, new_pid = start_word()
            self._set_word_pid(new_pid if new_pid else None)

            if word is None:
                stats.word_restarts_failed += 1
                self._emit("STEP: Restart Word (FAIL)")
                raise RuntimeError("Word restart failed.")

            try:
                stats.word_version = str(word.Version)
            except Exception:
                pass

            stats.word_restarts_ok += 1
            self._emit(f"STEP: Restart Word (done) | pid={new_pid if new_pid else 'UNKNOWN'}")

        def open_document_with_retry(work_abs_path: str, file_label_for_log: str):
            nonlocal word

            try:
                return word.Documents.Open(
                    work_abs_path,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=False,
                    ConfirmConversions=False,
                )
            except Exception as first_exc:
                stats.open_retries_attempted += 1
                self._emit(f"STEP: Word OPEN retry triggered | file={file_label_for_log} | first_error={first_exc}")

                restart_word(f"retry open after Word OPEN failure for {file_label_for_log}")

                try:
                    retried_doc = word.Documents.Open(
                        work_abs_path,
                        ReadOnly=False,
                        AddToRecentFiles=False,
                        Visible=False,
                        ConfirmConversions=False,
                    )
                    stats.open_retries_ok += 1
                    self._emit(f"STEP: Word OPEN retry (done) | file={file_label_for_log}")
                    return retried_doc
                except Exception as retry_exc:
                    stats.open_retries_failed += 1
                    self._emit(f"STEP: Word OPEN retry (FAIL) | file={file_label_for_log} | retry_error={retry_exc}")
                    raise retry_exc from first_exc

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

                if restart_interval_docs > 0 and idx > 1 and (idx - 1) % restart_interval_docs == 0:
                    restart_word(f"periodic restart before file index {idx}")

                file_label = os.path.basename(rel_path)
                abs_in = os.path.abspath(abs_path)
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
                except Exception as exc:
                    stats.failed += 1
                    err = f"FAIL: working copy error: {exc}"
                    error_rows.append((file_label, "WorkingCopy", err[:250]))
                    self._emit(f"STEP: working copy (FAIL) | {exc}")
                    continue

                self._set_status(file_label, "Backup copy")
                self._emit("STEP: backup copy (start)")
                backup_status, backup_dest = self._copy_backup(abs_in, rel_path, backup_run_root)
                self._emit(f"STEP: backup copy (done) | status={backup_status}")

                doc = None
                editability = {}
                protection_result = {
                    "attempted": "NO",
                    "password_used": "N/A",
                    "unprotect_before_success": "N/A",
                    "body_exception_success": "N/A",
                    "protect_success": "N/A",
                    "protection_type_after": "N/A",
                    "error": "",
                }
                details = {"sections": 0, "sections_data": []}
                actions = 0
                aggs = {
                    "format_apps": 0,
                    "header_post_table_deleted": 0,
                    "footer_post_table_deleted": 0,
                    "tables_centered": 0,
                    "text_shapes_formatted": 0,
                    "quickparts_na_filled": 0,
                    "table_cells_formatted": 0,
                }

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

                    doc = open_document_with_retry(work_abs, file_label)

                    self._emit("STEP: Word OPEN (done)")

                    if self._word_killed_by_watchdog:
                        raise TimeoutError(f"TIMEOUT during Word OPEN (>{timeout_s}s)")

                    if show_word:
                        try:
                            word.Visible = True
                            doc.Activate()
                        except Exception:
                            pass

                    self._set_status(file_label, "Editability")
                    self._emit("STEP: Editability (start)")
                    clear_readonly_attribute(work_abs)
                    editability = force_document_editable(
                        doc,
                        password_hint=protection_password if protect_headers_footers else "",
                    )
                    self._emit(f"STEP: Editability (done) | doc_readonly_final={editability.get('doc_readonly_final', 'UNKNOWN')}")

                    if self._word_killed_by_watchdog:
                        raise TimeoutError(f"TIMEOUT during Editability (>{timeout_s}s)")

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

                        except Exception as exc:
                            stats.docid_missing_count += 1
                            doc_id_value = None
                            self._emit(f"STEP: Extract Document ID (FAIL) | {exc}")

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

                    if protect_headers_footers:
                        stats.protections_attempted += 1
                        self._set_status(file_label, "Restrict headers/footers")
                        self._emit("STEP: Restrict headers/footers (start)")

                        protection_result = restrict_headers_footers_by_body_exception(
                            doc,
                            password=protection_password,
                        )

                        details["protection"] = protection_result

                        self._emit(
                            "STEP: Restrict headers/footers (done) | "
                            f"PASSWORD={protection_result.get('password_used', 'N/A')} | "
                            f"BODY_EXCEPTION={protection_result.get('body_exception_success', 'NO')} | "
                            f"PROTECT={protection_result.get('protect_success', 'NO')} | "
                            f"PROTECTION_TYPE={protection_result.get('protection_type_after', 'UNKNOWN')} | "
                            f"ERROR={protection_result.get('error', '') or 'N/A'}"
                        )

                        if protection_result.get("protect_success") != "YES":
                            stats.protections_failed += 1
                            raise RuntimeError(
                                "Header/footer restriction failed: "
                                + str(protection_result.get("error", "unknown error"))
                            )

                        stats.protections_ok += 1

                        if self._word_killed_by_watchdog:
                            raise TimeoutError(f"TIMEOUT during header/footer restriction (>{timeout_s}s)")
                    else:
                        protection_result = {
                            "attempted": "NO",
                            "password_used": "N/A",
                            "unprotect_before_success": "N/A",
                            "body_exception_success": "N/A",
                            "protect_success": "N/A",
                            "protection_type_after": "N/A",
                            "error": "",
                        }

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

                    if export_pdf_each:
                        stats.pdf_exports_attempted += 1
                        self._set_status(file_label, "Export PDF")
                        self._emit("STEP: Export PDF (start)")

                        try:
                            export_doc_to_pdf(doc, output_pdf_path)
                            stats.pdf_exports_ok += 1
                            self._emit(f"STEP: Export PDF (done) | status=SUCCESS | pdf_path={output_pdf_path}")
                        except Exception as exc:
                            stats.pdf_exports_failed += 1
                            self._emit(f"STEP: Export PDF (FAIL) | {exc}")

                    try:
                        doc.Close(False)
                    except Exception:
                        pass

                    doc = None

                    stats.succeeded += 1
                    stats.total_actions += int(actions)
                    stats.format_apps_total += int(aggs.get("format_apps", 0))
                    stats.header_post_table_deleted_total += int(aggs.get("header_post_table_deleted", 0))
                    stats.footer_post_table_deleted_total += int(aggs.get("footer_post_table_deleted", 0))
                    stats.tables_centered_total += int(aggs.get("tables_centered", 0))
                    stats.text_shapes_formatted_total += int(aggs.get("text_shapes_formatted", 0))
                    stats.quickparts_na_filled_total += int(aggs.get("quickparts_na_filled", 0))
                    stats.table_cells_formatted_total += int(aggs.get("table_cells_formatted", 0))

                    block = []
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
                    block.append(f"  DOC READONLY (INITIAL)     : {editability.get('doc_readonly_initial', 'UNKNOWN')}")
                    block.append(f"  UNPROTECT SUCCESS          : {editability.get('unprotect_success', 'NO')}")
                    block.append(f"  FINAL CLEARED SUCCESS      : {editability.get('final_cleared_success', 'NO')}")
                    block.append(f"  READONLY RECOMMENDED CLR   : {editability.get('readonly_recommended_cleared_success', 'NO')}")
                    block.append(f"  CHANGEFILEACCESS SUCCESS   : {editability.get('changefileaccess_success', 'NO')}")
                    block.append(f"  DOC READONLY (FINAL)       : {editability.get('doc_readonly_final', 'UNKNOWN')}")
                    block.append("")
                    block.append("PROTECTION CHECK")
                    block.append(f"  ATTEMPTED                  : {protection_result.get('attempted', 'NO')}")
                    block.append(f"  PASSWORD USED              : {protection_result.get('password_used', 'N/A')}")
                    block.append(f"  UNPROTECT BEFORE SUCCESS   : {protection_result.get('unprotect_before_success', 'N/A')}")
                    block.append(f"  BODY EXCEPTION SUCCESS     : {protection_result.get('body_exception_success', 'N/A')}")
                    block.append(f"  PROTECT SUCCESS            : {protection_result.get('protect_success', 'N/A')}")
                    block.append(f"  PROTECTION TYPE AFTER      : {protection_result.get('protection_type_after', 'N/A')}")
                    block.append(f"  ERROR                      : {protection_result.get('error', '') or 'N/A'}")
                    block.append("")
                    block.append("OPEN/PROCESS STATUS          : SUCCESS")
                    block.append(f"SECTION COUNT                : {details.get('sections', 0)}")
                    block.append(f"PROCESS SHAPES               : {'YES' if enable_shapes else 'NO'}")
                    block.append("")

                    section_count = int(details.get("sections", 0))

                    for sec in details.get("sections_data", []):
                        section_idx = sec["section_index"]
                        block.append(f"SECTION {section_idx}/{section_count}")

                        for kind_name in ["PRIMARY", "FIRSTPAGE", "EVENPAGES"]:
                            hdr = sec["headers"].get(kind_name, {"present": "NO"})

                            block.append(f"  HEADER {kind_name}")
                            block.append(f"    PRESENT                  : {hdr.get('present', 'NO')}")
                            block.append(f"    RANGE EMPTY              : {hdr.get('range_empty', 'UNKNOWN')}")
                            block.append(f"    FORMAT NORMALIZATION     : APPLIED={hdr.get('format_applied', 'NO')} | TARGET=(Before=0, After=0, Line=Single)")
                            block.append(f"    TABLES FOUND             : {hdr.get('tables_found', 'NO')} | COUNT={hdr.get('table_count', 'UNKNOWN')} | ANCHOR={hdr.get('anchor', 'N/A')}")
                            block.append(f"    TABLE CENTERING          : APPLIED={hdr.get('tables_centering_applied', 'NO')} | CENTERED_COUNT={hdr.get('tables_centered_count', 0)}")
                            block.append(f"    TABLE CELL FORMAT        : Arial 11; contents left except GxP centered | CELLS_FORMATTED={hdr.get('table_cells_formatted', 0)}")
                            block.append(f"    BLANK CLEANUP METHOD     : {hdr.get('blank_cleanup_method', 'UNKNOWN')}")
                            block.append(f"    POST-TABLE SPEC_MAX      : {hdr.get('spec_max', 'UNKNOWN')}")
                            block.append(f"    POST-TABLE BLANKS FOUND  : {hdr.get('post_table_blanks_found', 'N/A')} | COUNT={hdr.get('post_table_blank_count', 'N/A')}")
                            block.append(f"    POST-TABLE BLANKS ACTION : REQUIRED={hdr.get('post_table_deletions_required', 'N/A')} | DELETED={hdr.get('post_table_deleted', 0)} | REMAINING={hdr.get('post_table_remaining', 'N/A')}")
                            block.append(f"    SHAPES FOUND             : {hdr.get('shapes_found', 'NO')} | COUNT={hdr.get('shapes_count', 0)}")
                            block.append(f"    TEXT SHAPES FOUND        : {hdr.get('text_shapes_found', 'NO')} | COUNT={hdr.get('text_shapes_count', 0)}")
                            block.append(f"    TEXT SHAPES FORMATTED    : {hdr.get('text_shapes_formatted', 0)}")
                            block.append(f"    QUICKPARTS N/A FILLED    : {hdr.get('quickparts_na_filled', 0)}")

                        for kind_name in ["PRIMARY", "FIRSTPAGE", "EVENPAGES"]:
                            ftr = sec["footers"].get(kind_name, {"present": "NO"})

                            block.append(f"  FOOTER {kind_name}")
                            block.append(f"    PRESENT                  : {ftr.get('present', 'NO')}")
                            block.append(f"    RANGE EMPTY              : {ftr.get('range_empty', 'UNKNOWN')}")
                            block.append(f"    FORMAT NORMALIZATION     : APPLIED={ftr.get('format_applied', 'NO')} | TARGET=(Before=0, After=0, Line=Single)")
                            block.append(f"    TABLES FOUND             : {ftr.get('tables_found', 'NO')} | COUNT={ftr.get('table_count', 'UNKNOWN')} | ANCHOR={ftr.get('anchor', 'N/A')}")
                            block.append(f"    TABLE CENTERING          : APPLIED={ftr.get('tables_centering_applied', 'NO')} | CENTERED_COUNT={ftr.get('tables_centered_count', 0)}")
                            block.append(f"    TABLE CELL FORMAT        : Arial 9; contents centered | CELLS_FORMATTED={ftr.get('table_cells_formatted', 0)}")
                            block.append(f"    BLANK CLEANUP METHOD     : {ftr.get('blank_cleanup_method', 'UNKNOWN')}")
                            block.append(f"    POST-TABLE SPEC_MAX      : {ftr.get('spec_max', 'UNKNOWN')}")
                            block.append(f"    POST-TABLE BLANKS FOUND  : {ftr.get('post_table_blanks_found', 'N/A')} | COUNT={ftr.get('post_table_blank_count', 'N/A')}")
                            block.append(f"    POST-TABLE BLANKS ACTION : REQUIRED={ftr.get('post_table_deletions_required', 'N/A')} | DELETED={ftr.get('post_table_deleted', 0)} | REMAINING={ftr.get('post_table_remaining', 'N/A')}")
                            block.append(f"    SHAPES FOUND             : {ftr.get('shapes_found', 'NO')} | COUNT={ftr.get('shapes_count', 0)}")
                            block.append(f"    TEXT SHAPES FOUND        : {ftr.get('text_shapes_found', 'NO')} | COUNT={ftr.get('text_shapes_count', 0)}")
                            block.append(f"    TEXT SHAPES FORMATTED    : {ftr.get('text_shapes_formatted', 0)}")
                            block.append(f"    QUICKPARTS N/A FILLED    : {ftr.get('quickparts_na_filled', 0)}")

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

                except TimeoutError as exc:
                    stats.failed += 1
                    msg = f"TIMEOUT: {exc}"
                    error_rows.append((file_label, "Timeout", msg[:250]))
                    self._enqueue_lines([f"ERROR: {msg}"])

                    try:
                        if doc is not None:
                            doc.Close(False)
                    except Exception:
                        pass

                    doc = None

                    restart_word(f"after timeout while processing {file_label}")

                except Exception as exc:
                    stats.failed += 1
                    err = f"FAIL: {exc}"
                    error_rows.append((file_label, "Process", err[:250]))
                    self._enqueue_lines([f"ERROR: {err}", traceback.format_exc()])

                    try:
                        if doc is not None:
                            doc.Close(False)
                    except Exception:
                        pass

                    doc = None

                    try:
                        restart_word(f"after document failure for {file_label}")
                    except Exception as restart_exc:
                        self._emit(f"ERROR: Word restart after failure also failed: {restart_exc}")

                finally:
                    self.after(0, self._update_stats_line)
                    self._set_status(file_label, "Idle")

            ensure_dir(run_root)

            txt_path = os.path.join(run_root, f"run_log_{run_id}.txt")
            with open(txt_path, "w", encoding="utf-8") as handle:
                handle.write("\n".join(run_log_lines))

            stats.end_dt = datetime.now().astimezone()

            pdf_path = os.path.join(run_root, f"run_report_{run_id}.pdf")

            try:
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
                    include_detail=detailed_pdf_report_enabled,
                )
                self._enqueue_lines(["", f"PDF REPORT WRITTEN            : {pdf_path}", f"TEXT LOG WRITTEN              : {txt_path}", ""])
            except Exception as exc:
                self._enqueue_lines(["", f"PDF REPORT FAILED             : {exc}", f"TEXT LOG WRITTEN              : {txt_path}", ""])

        except Exception as exc:
            stats.end_dt = datetime.now().astimezone()
            fatal = [ts_line(), "FATAL ERROR", str(exc), traceback.format_exc(), ts_line()]
            self._enqueue_lines(fatal)

        finally:
            heartbeat_stop.set()

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
                f"TABLE CELLS FORMATTED       : {stats.table_cells_formatted_total}",
                f"RESTART WORD EVERY N DOCS   : {restart_interval_docs}",
                f"WORD RESTARTS ATTEMPTED     : {stats.word_restarts_attempted}",
                f"WORD RESTARTS OK            : {stats.word_restarts_ok}",
                f"WORD RESTARTS FAILED        : {stats.word_restarts_failed}",
                f"OPEN RETRIES ATTEMPTED      : {stats.open_retries_attempted}",
                f"OPEN RETRIES OK             : {stats.open_retries_ok}",
                f"OPEN RETRIES FAILED         : {stats.open_retries_failed}",
                f"DETAILED PDF REPORT         : {'YES' if detailed_pdf_report_enabled else 'NO'}",
                f"TEXT SHAPES FORMATTED       : {stats.text_shapes_formatted_total}",
                f"FORMAT APPS (COUNT)         : {stats.format_apps_total}",
                f"QUICKPARTS N/A FILLED       : {stats.quickparts_na_filled_total}",
                f"RESTRICT HEADERS/FOOTERS    : {'YES' if protect_headers_footers else 'NO'}",
                f"RESTRICTION PASSWORD        : {protection_password if protect_headers_footers else 'N/A'}",
                f"PROTECTIONS ATTEMPTED       : {stats.protections_attempted}",
                f"PROTECTIONS OK              : {stats.protections_ok}",
                f"PROTECTIONS FAILED          : {stats.protections_failed}",
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
