# TW_document_version_history_generator.py
"""
Purpose:
    Generate a version-complete TrackWise DMS document version history workbook and,
    optionally, reviewer-ready PDF evidence packages.

What this file does:
    Reads raw TrackWise Salesforce export CSVs, using
    SPARTADMS__Document_Version__c.csv as the primary source for version history.
    It enriches those version rows with Corporate Document metadata, decoded user
    names, review/approval actor records, revision justification records,
    effective-version detail records, audit-history obsoletion events, and optional
    Neurotic DMS manifest evidence.

    If PDF packaging is enabled, it locates source PDFs in the raw DMS file tree,
    generates a one-page QA-facing cover page for each document version, and writes
    a new PDF package containing the generated cover page followed by the original
    TrackWise PDF rendition. Existing source PDFs are never modified.

Place in the larger scheme:
    This utility sits downstream of the TrackWise Salesforce/DMS export and the
    Neurotic DMS organization process. It converts raw export evidence into a
    reviewer-facing version register and optional evidence-package PDFs.

Why that matters:
    TrackWise DMS may preserve complete version history in
    SPARTADMS__Document_Version__c, while SPARTADMS__Corporate_Document__c may
    expose only a parent/current or reduced view. Version-level reconstruction is
    therefore required for document-control review, deviation support, and
    audit-ready migration evidence.
"""

from __future__ import annotations

import csv
import getpass
import hashlib
import importlib
import os
import platform
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
import math
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

APP_TITLE = "NEUROTIC DMS Document Version History Generator"
APP_VERSION = "4.5.7"
RUN_ID_PREFIX = "DMSHIST"
PACKAGE_RUN_ID_PREFIX = "DMSPKG"
ENCODING_FALLBACKS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
PDF_DEPENDENCY_MESSAGE = (
    "PDF packaging with a migration cover page requires reportlab and pypdf. "
    "Install with: python -m pip install reportlab pypdf. "
    "If you only want clean-named copies of the source PDFs, leave Generate clean version PDFs checked "
    "but uncheck Include PDF Cover Page."
)

EXPECTED_RAW_EXPORT_FILENAMES: Dict[str, str] = {
    "document_version_csv": "SPARTADMS__Document_Version__c.csv",
    "corporate_document_csv": "SPARTADMS__Corporate_Document__c.csv",
    "user_csv": "User.csv",
    "review_justification_csv": "SPARTADMS__Review_Justification__c.csv",
    "approvers_file": "SPARTADMS__Document_Approvers__c.csv",
    "reviewers_file": "SPARTADMS__Document_Reviewer__c.csv",
    "effective_version_details_csv": "SPARTADMS__Effective_Version_Details__c.csv",
    "audit_activity_type_csv": "CMPL123__Audit_Activity_Type__c.csv",
    "audit_history_csv": "CMPL123__Audit_History__c.csv",
}

DOC_TYPE_PREFIX_RULES: Sequence[Tuple[re.Pattern[str], str]] = (
    (re.compile(r"standard\s+operating\s+procedure|\(\s*SOP\s*\)|\bSOP\b", re.I), "SOP"),
    (re.compile(r"work\s+instruction|\(\s*WI\s*\)|\bWI\b", re.I), "WI"),
    (re.compile(r"form|\(\s*FRM\s*\)|\bFRM\b", re.I), "FRM"),
    (re.compile(r"validation|\(\s*VAL\s*\)|\bVAL\b", re.I), "VAL"),
    (re.compile(r"protocol|\(\s*PRO\s*\)|\bPRO\b", re.I), "PRO"),
    (re.compile(r"report|\(\s*RPT\s*\)|\bRPT\b", re.I), "RPT"),
    (re.compile(r"policy|\(\s*POL\s*\)|\bPOL\b", re.I), "POL"),
)

# Prefixes whose numeric sequence is part of the controlled document number, not merely
# a TrackWise source-lineage discriminator. CORP documents are explicitly numbered as
# CORP-000001.01, CORP-000002.01, etc. and must not be collapsed to CORP.01.
PRESERVE_FULL_SOURCE_NUMBER_PREFIXES = {"CORP"}

PHASE_WEIGHTS_NO_PDF: Dict[str, int] = {
    "VALIDATE_INPUTS": 3,
    "LOAD_INPUTS": 12,
    "BUILD_INDEXES": 10,
    "BUILD_SOURCE_VERSION_ROWS": 20,
    "MAP_JUSTIFICATIONS": 15,
    "MAP_ACTORS_AND_OBSOLETION": 10,
    "BUILD_WORKBOOK_TABLES": 10,
    "WRITE_WORKBOOK": 20,
}

PHASE_WEIGHTS_WITH_PDF: Dict[str, int] = {
    "VALIDATE_INPUTS": 3,
    "LOAD_INPUTS": 10,
    "BUILD_INDEXES": 8,
    "BUILD_SOURCE_VERSION_ROWS": 15,
    "MAP_JUSTIFICATIONS": 12,
    "MAP_ACTORS_AND_OBSOLETION": 8,
    "BUILD_WORKBOOK_TABLES": 9,
    "PDF_PACKAGING": 20,
    "WRITE_WORKBOOK": 15,
}

MAIN_COLUMNS: List[str] = [
    "Document Number",
    "Version",
    "Full Document Version",
    "Output PDF Filename",
    "Output PDF Path",
    "Document Title",
    "Document Type",
    "Department",
    "Document Status",
    "Version Status",
    "Lifecycle Status Normalized",
    "Approved Date",
    "Effective Date",
    "Obsolete Date",
    "Obsolete / Superseded Date Basis",
    "Canceled Date",
    "Canceled Date Basis",
    "Superseded By",
    "Justification ID",
    "Justification Content",
    "Justification Created Date",
    "Justification Created By",
    "Justification Owner",
    "Justification Last Modified Date",
    "Justification Last Modified By",
    "Justification Source Salesforce ID",
    "Justification Related Corporate Document ID",
    "Justification Raw Version Field",
    "Justification Raw Unformatted Version Field",
    "Justification Mapping Method",
    "Justification Confidence",
    "Justification Mapping Detail",
    "Reviewed Date",
    "Reviewed By",
    "Review Details",
    "Approved By",
    "Approval Details",
    "Obsoleted By",
    "Obsolete Details",
    "Canceled By",
    "Canceled Details",
    "Previous Display Version in Register",
    "Next Display Version in Register",
    "Created Date",
    "Created By",
    "Last Modified Date",
    "Last Modified By",
    "Related Quality Event",
    "Training Required",
    "Source Version Key(s)",
    "Raw TrackWise Document Number(s)",
    "Source Document Sequence Number(s)",
    "Raw Document Version ID(s)",
    "Raw Corporate Document ID(s)",
    "Source PDF Path(s)",
    "Source PDF SHA-256",
    "Selected Source PDF Path",
    "Selected Source PDF SHA-256",
    "Selected Source Score",
    "Selected Source Selection Reason",
    "Candidate Source PDF Count",
    "Distinct Candidate SHA-256 Count",
    "Candidate Source PDF Paths",
    "Candidate Source PDF SHA-256 Values",
    "Candidate Source Scores",
    "Source Conflict Status",
    "Manual Review Recommended",
    "Source Selection Rule",
    "Source Selection Basis",
    "Output PDF SHA-256",
    "PDF Packaging Result",
    "PDF Packaging Detail",
    "Source Notes",
]

SOURCE_SELECTION_TEXT_COLUMNS: List[str] = [
    "Selected Source PDF Path",
    "Selected Source PDF SHA-256",
    "Selected Source Score",
    "Selected Source Selection Reason",
    "Candidate Source PDF Count",
    "Distinct Candidate SHA-256 Count",
    "Candidate Source PDF Paths",
    "Candidate Source PDF SHA-256 Values",
    "Candidate Source Scores",
    "Source Conflict Status",
    "Manual Review Recommended",
    "Source Selection Rule",
    "Source Selection Basis",
    "Source PDF Path(s)",
    "Source PDF SHA-256",
    "Output PDF Path",
    "Output PDF Filename",
    "Output PDF SHA-256",
    "PDF Packaging Result",
    "PDF Packaging Detail",
]


TIMELINE_COLUMNS: List[str] = [
    "Timeline Sequence",
    "Document Number",
    "Version",
    "Full Document Version",
    "Timeline Event",
    "Event Date",
    "Status After Event",
    "Event Actor(s)",
    "Evidence Source",
    "Evidence Source Row / ID",
    "Justification ID",
    "Justification Content",
    "Justification Created By",
    "Justification Created Date",
    "Mapping Method",
    "Mapping Confidence",
    "Source Version Key(s)",
    "Detail",
]

EVENT_COLUMNS: List[str] = [
    "percent_complete",
    "phase_number",
    "phase_total",
    "phase_percent_complete",
    "timestamp",
    "level",
    "phase",
    "source_file",
    "source_row_number",
    "raw_document_version_id",
    "raw_corporate_document_id",
    "raw_trackwise_document_number",
    "source_document_sequence_number",
    "document_number",
    "version",
    "full_document_version",
    "justification_id",
    "action",
    "result",
    "mapping_method",
    "mapping_confidence",
    "detail",
]

FIXED_WIDTHS: Dict[str, Dict[str, int]] = {
    "Document Version Register": {
        "A": 18,
        "B": 10,
        "C": 22,
        "D": 34,
        "E": 44,
        "F": 28,
        "G": 24,
        "H": 16,
        "I": 16,
        "J": 20,
        "K": 20,
        "L": 20,
        "M": 16,
        "N": 65,
        "W": 35,
        "Y": 35,
        "AA": 30,
        "AO": 55,
    },
    "Document Version Timeline": {"A": 10, "B": 18, "D": 22, "E": 26, "F": 20, "H": 40, "K": 16, "L": 80, "R": 80},
    "Source Version Rows": {"A": 18, "B": 18, "C": 22, "D": 42, "E": 25, "F": 24, "G": 16},
    "Justification Mapping Diagnostics": {"A": 16, "B": 22, "C": 16, "D": 20, "E": 22, "F": 18, "G": 70},
    "Approver Detail": {"A": 20, "B": 20, "C": 28, "D": 28, "E": 18},
    "Reviewer Detail": {"A": 20, "B": 20, "C": 28, "D": 28, "E": 18},
    "Status Event Detail": {"A": 24, "B": 20, "C": 22, "D": 28, "E": 70},
    "Run Summary": {"A": 32, "B": 70},
    "Event Log": {"A": 10, "B": 10, "E": 22, "F": 10, "G": 24, "R": 26, "S": 14, "V": 80},
}


@dataclass
class EventRecord:
    percent_complete: int = 0
    phase_number: int = 0
    phase_total: int = 0
    phase_percent_complete: int = 0
    timestamp: str = ""
    level: str = "INFO"
    phase: str = ""
    source_file: str = ""
    source_row_number: str = ""
    raw_document_version_id: str = ""
    raw_corporate_document_id: str = ""
    raw_trackwise_document_number: str = ""
    source_document_sequence_number: str = ""
    document_number: str = ""
    version: str = ""
    full_document_version: str = ""
    justification_id: str = ""
    action: str = ""
    result: str = ""
    mapping_method: str = ""
    mapping_confidence: str = ""
    detail: str = ""


@dataclass
class ProgressTracker:
    phase_weights: Dict[str, int]
    current_phase: str = ""
    phase_number: int = 0
    phase_percent: int = 0

    def __post_init__(self) -> None:
        self.phase_names: List[str] = list(self.phase_weights.keys())
        self.phase_starts: Dict[str, int] = {}
        running_total = 0
        for phase_name, weight in self.phase_weights.items():
            self.phase_starts[phase_name] = running_total
            running_total += weight

    @property
    def phase_total(self) -> int:
        return len(self.phase_names)

    def set_phase(self, phase: str, phase_percent: int = 0) -> None:
        self.current_phase = phase
        self.phase_number = self.phase_names.index(phase) + 1 if phase in self.phase_names else 0
        self.phase_percent = max(0, min(100, int(phase_percent)))

    def update_phase_percent(self, completed: int, total: int) -> None:
        if total <= 0:
            self.phase_percent = 100
            return
        self.phase_percent = max(0, min(100, int(round(completed * 100 / total))))

    def overall_percent(self) -> int:
        if not self.current_phase:
            return 0
        start = self.phase_starts.get(self.current_phase, 0)
        weight = self.phase_weights.get(self.current_phase, 0)
        return max(0, min(100, int(round(start + (weight * self.phase_percent / 100)))))


@dataclass
class StructuredLogger:
    progress: ProgressTracker
    ui_queue: "queue.Queue[Dict[str, Any]]"
    event_rows: List[Dict[str, Any]] = field(default_factory=list)

    def emit(
        self,
        phase: Optional[str] = None,
        level: str = "INFO",
        action: str = "",
        result: str = "",
        detail: str = "",
        phase_percent: Optional[int] = None,
        **kwargs: Any,
    ) -> None:
        if phase:
            self.progress.set_phase(phase, self.progress.phase_percent if phase_percent is None else phase_percent)
        elif phase_percent is not None and self.progress.current_phase:
            self.progress.set_phase(self.progress.current_phase, phase_percent)

        record = EventRecord(
            percent_complete=self.progress.overall_percent(),
            phase_number=self.progress.phase_number,
            phase_total=self.progress.phase_total,
            phase_percent_complete=self.progress.phase_percent,
            timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            level=level,
            phase=self.progress.current_phase,
            action=action,
            result=result,
            detail=detail,
        )
        for key, value in kwargs.items():
            if hasattr(record, key):
                setattr(record, key, clean_scalar(value))
        row = asdict(record)
        self.event_rows.append(row)
        self.ui_queue.put(row)

    def progress_event(
        self,
        completed: int,
        total: int,
        action: str,
        result: str = "PROGRESS",
        detail: str = "",
        every: int = 500,
        force: bool = False,
        **kwargs: Any,
    ) -> None:
        self.progress.update_phase_percent(completed, total)
        if force or completed == total or completed == 0 or completed % every == 0:
            self.emit(action=action, result=result, detail=detail or f"{completed:,} / {total:,}", **kwargs)


@dataclass
class RunInputs:
    document_version_csv: Path
    corporate_document_csv: Path
    user_csv: Path
    output_run_folder: Path
    review_justification_csv: Optional[Path] = None
    approvers_file: Optional[Path] = None
    reviewers_file: Optional[Path] = None
    effective_version_details_csv: Optional[Path] = None
    audit_activity_type_csv: Optional[Path] = None
    audit_history_csv: Optional[Path] = None
    manifest_csv: Optional[Path] = None
    raw_dms_file_tree_folder: Optional[Path] = None
    generate_pdf_packages: bool = False
    include_pdf_cover_page: bool = False
    include_technical_evidence_sheet: bool = True
    include_event_log_sheet: bool = True
    fail_if_populated_optional_missing: bool = True


@dataclass
class RunOutputs:
    run_folder: Path
    workbook_path: Path
    event_log_csv_path: Path
    pdf_package_folder: Path
    package_log_folder: Path
    run_id: str
    package_run_id: str
    run_timestamp: str
    run_machine: str
    run_account: str


class CancelledRun(RuntimeError):
    """Raised when the user cancels an active run."""


class CancelToken:
    def __init__(self) -> None:
        self._event = threading.Event()

    def cancel(self) -> None:
        self._event.set()

    def is_cancelled(self) -> bool:
        return self._event.is_set()

    def check(self) -> None:
        if self.is_cancelled():
            raise CancelledRun("Run cancelled by user.")


def clean_scalar(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value)
    return "" if text.lower() in {"nan", "nat", "none"} else text


def first_present(row: pd.Series, columns: Sequence[str]) -> str:
    for column in columns:
        if column in row.index:
            value = clean_scalar(row[column]).strip()
            if value:
                return value
    return ""


def first_column(df: pd.DataFrame, columns: Sequence[str]) -> str:
    for column in columns:
        if column in df.columns:
            return column
    return ""


def as_timestamp(value: Any) -> Optional[pd.Timestamp]:
    text = clean_scalar(value).strip()
    if not text:
        return None
    normalized = text.replace("T", " ").replace("Z", "")
    normalized = re.sub(r"([+-]\d{2}:?\d{2})$", "", normalized).strip()
    compact = normalized[:19]
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        candidate = compact[:10] if fmt == "%Y-%m-%d" else compact
        try:
            return pd.Timestamp(datetime.strptime(candidate, fmt))
        except ValueError:
            pass
    parsed = pd.to_datetime(text, errors="coerce", utc=False)
    if pd.isna(parsed):
        return None
    if isinstance(parsed, pd.Timestamp):
        return parsed
    return pd.Timestamp(parsed)


def format_datetime(value: Any) -> str:
    parsed = as_timestamp(value)
    if parsed is None:
        return ""
    try:
        if parsed.tzinfo is not None:
            parsed = parsed.tz_convert(None)
    except (TypeError, AttributeError):
        pass
    return parsed.strftime("%Y-%m-%d %H:%M:%S")


def version_sort_key(version_text: Any) -> Tuple[int, str]:
    text = clean_scalar(version_text).strip()
    match = re.search(r"\d+", text)
    if match:
        return int(match.group(0)), text
    return 999999, text


def normalize_version(value: Any) -> str:
    text = clean_scalar(value).strip()
    if not text:
        return ""
    numeric = pd.to_numeric(text, errors="coerce")
    if not pd.isna(numeric):
        return f"{int(float(numeric)):02d}"
    match = re.search(r"\d+", text)
    if match:
        return f"{int(match.group(0)):02d}"
    return text


def normalize_version_int(value: Any) -> Optional[int]:
    text = clean_scalar(value).strip()
    if not text:
        return None
    numeric = pd.to_numeric(text, errors="coerce")
    if not pd.isna(numeric):
        return int(float(numeric))
    match = re.search(r"\d+", text)
    return int(match.group(0)) if match else None


def strip_source_sequence(raw_number: str) -> Tuple[str, str]:
    text = clean_scalar(raw_number).strip()
    match = re.match(r"^(.*?)-0*(\d+)$", text)
    if not match:
        return text, ""
    return match.group(1), str(int(match.group(2)))


def infer_prefix(document_type: str, document_number_field: str, raw_base: str) -> str:
    doc_num = clean_scalar(document_number_field).strip()
    match = re.search(r"-([A-Z]{2,5})$", doc_num)
    if match:
        return match.group(1).upper()
    doc_type = clean_scalar(document_type)
    for pattern, prefix in DOC_TYPE_PREFIX_RULES:
        if pattern.search(doc_type):
            return prefix
    alpha = re.match(r"^([A-Z]+)", raw_base.upper())
    if alpha:
        return alpha.group(1)
    return "DOC"


def build_display_document_number(raw_number: str, document_type: str, document_number_field: str) -> Tuple[str, str, str]:
    """Return the reviewer-facing controlled document number without artificial category stacking.

    TrackWise raw source folders can include source-lineage discriminators. The
    reviewer-facing document number must strip the trailing source sequence when the
    portion before that sequence already contains the controlled identity. For example,
    ``COT-0005-000001`` becomes ``COT-0005`` and ``PQP-V481-000001`` becomes
    ``PQP-V481``. The raw source identity remains preserved separately in source-key
    and source-path fields.

    CORP is the controlled exception: ``CORP-000001`` is the actual controlled document
    number, so it remains ``CORP-000001``.
    """
    raw_text = clean_scalar(raw_number).strip().replace("_", "-").upper()
    if not raw_text:
        return "", "", ""

    if not re.search(r"[A-Z]", raw_text):
        raw_base, source_sequence = strip_source_sequence(raw_text)
        raw_base = clean_scalar(raw_base).strip().replace("_", "-").upper()
        if not raw_base:
            return "", raw_base, source_sequence
        prefix = infer_prefix(document_type, document_number_field, raw_base)
        display = f"{prefix}-{raw_base}" if prefix else raw_base
    else:
        source_sequence = ""
        display = raw_text
        raw_base = raw_text

        preserve_match = re.match(r"^([A-Z]+)-0*(\d{6})$", raw_text)
        if preserve_match and preserve_match.group(1) in PRESERVE_FULL_SOURCE_NUMBER_PREFIXES:
            # CORP-000001, CORP-000002, etc. are true controlled numbers.
            display = raw_text
            raw_base = raw_text
            source_sequence = str(int(preserve_match.group(2)))
        else:
            suffix_match = re.match(r"^(?P<base>.+)-(?P<seq>\d{6})$", raw_text)
            if suffix_match:
                candidate_base = suffix_match.group("base")
                seq = suffix_match.group("seq")
                # Strip the final source sequence only when the base already carries
                # a specific controlled identity, not when it is just an alphabetic
                # family prefix such as RCD or MS.
                if re.search(r"\d", candidate_base):
                    display = candidate_base
                    raw_base = candidate_base
                    source_sequence = str(int(seq))
                else:
                    display = raw_text
                    raw_base = raw_text
                    source_sequence = ""

    display = re.sub(r"-{2,}", "-", display).strip("-").upper()

    # Defensive cleanup for stacked category prefixes. These patterns are not valid
    # reviewer-facing controlled-document identities for this migration output.
    display = re.sub(r"^FRM-(F\d+[A-Z]*)$", r"\1", display)
    display = re.sub(r"^MVP-(MVR-\d+)$", r"\1", display)
    display = re.sub(r"^MVR-(RPT-\d+)$", r"\1", display)
    display = re.sub(r"^(SOP|WI|TM|VAL)-\1-", r"\1-", display)
    return display, raw_base, source_sequence

def join_unique(values: Iterable[Any], separator: str = "; ") -> str:
    seen: Dict[str, None] = {}
    for value in values:
        text = clean_scalar(value).strip()
        if text and text not in seen:
            seen[text] = None
    return separator.join(seen.keys())



def summarize_detail_list(values: Sequence[str], limit: int = 10) -> str:
    unique_text = [value for value in dict.fromkeys(clean_scalar(v).strip() for v in values) if value]
    if len(unique_text) > limit:
        unique_text = unique_text[:limit] + [f"... {len(values) - limit} additional detail rows retained in supporting detail sheet."]
    return excel_cell("\n".join(unique_text), 6000)


def excel_cell(text: Any, max_chars: int = 32000) -> str:
    value = clean_scalar(text)
    if len(value) <= max_chars:
        return value
    return value[: max_chars - 120] + f"\n...[truncated for Excel cell limit; full source content remains in supporting diagnostics/source sheets, original length={len(value)}]"

def read_csv_with_fallback(path: Path) -> Tuple[pd.DataFrame, str]:
    last_error: Optional[BaseException] = None
    for encoding in ENCODING_FALLBACKS:
        try:
            df = pd.read_csv(path, dtype=str, encoding=encoding, low_memory=False, keep_default_na=False)
            return df, encoding
        except UnicodeDecodeError as exc:
            last_error = exc
        except pd.errors.ParserError as exc:
            last_error = exc
            try:
                df = pd.read_csv(path, dtype=str, encoding=encoding, engine="python", keep_default_na=False)
                return df, encoding
            except Exception as inner_exc:
                last_error = inner_exc
    raise RuntimeError(f"Unable to read CSV: {path}. Last error: {last_error}")


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(chunk_size)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def ensure_required_columns(df: pd.DataFrame, required: Sequence[str], source_name: str) -> None:
    missing = [column for column in required if column not in df.columns]
    if missing:
        raise RuntimeError(f"{source_name} is missing required columns: {', '.join(missing)}")



def ensure_object_columns(df: pd.DataFrame, columns: Sequence[str]) -> pd.DataFrame:
    """Ensure columns intended to hold paths/status strings can accept text values.

    Pandas can infer all-empty columns as float64/NaN. Later assigning a PDF path into
    those columns can raise ``Invalid value ... for dtype 'float64'``. The PDF packaging
    phase updates source-selection diagnostic columns row-by-row, so these columns must
    be object/string-capable before assignment begins.
    """
    for col in columns:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype("object")
        df[col] = df[col].where(pd.notna(df[col]), "")
    return df

def user_display_name(row: pd.Series) -> str:
    first = first_present(row, ["FirstName"])
    last = first_present(row, ["LastName"])
    full = " ".join(part for part in [first, last] if part).strip()
    if full:
        return full
    return first_present(row, ["Name", "Username", "Email", "Id"])


def build_user_map(user_df: pd.DataFrame) -> Dict[str, str]:
    user_map: Dict[str, str] = {}
    if "Id" not in user_df.columns:
        return user_map
    for row_tuple in user_df.itertuples(index=False):
        row = pd.Series(row_tuple._asdict())
        user_id = clean_scalar(row.get("Id", "")).strip()
        if user_id:
            user_map[user_id] = user_display_name(row)
    return user_map


def decode_user(user_id: Any, user_map: Dict[str, str]) -> str:
    text = clean_scalar(user_id).strip()
    if not text:
        return ""
    return user_map.get(text, text)


def format_event_line(record: Dict[str, Any]) -> str:
    percent = int(record.get("percent_complete") or 0)
    timestamp = clean_scalar(record.get("timestamp", ""))
    level = clean_scalar(record.get("level", "INFO"))
    phase = clean_scalar(record.get("phase", ""))
    full_doc = clean_scalar(record.get("full_document_version", ""))
    object_token = full_doc or clean_scalar(record.get("document_number", "")) or clean_scalar(record.get("raw_document_version_id", ""))
    justification = clean_scalar(record.get("justification_id", ""))
    result = clean_scalar(record.get("result", ""))
    method = clean_scalar(record.get("mapping_method", ""))
    confidence = clean_scalar(record.get("mapping_confidence", ""))
    mapping = " / ".join(part for part in [method, confidence] if part)
    detail = clean_scalar(record.get("detail", ""))
    middle = " | ".join(part for part in [object_token, justification, result, mapping] if part)
    if middle:
        return f"{percent:03d}% | {timestamp} | {level} | {phase} | {middle} | {detail}"
    return f"{percent:03d}% | {timestamp} | {level} | {phase} | {detail}"


def get_missing_pdf_dependencies(include_cover_page: bool) -> List[str]:
    """Return missing runtime dependencies for the selected PDF mode.

    Clean-named PDF copying does not require pypdf/PyPDF2/reportlab. A migration
    cover page requires ReportLab to create the cover and pypdf/PyPDF2 to merge
    the cover with the original TrackWise rendition. Several TrackWise renditions
    are AES-encrypted PDFs; pypdf requires the cryptography package to read those.
    """
    if not include_cover_page:
        return []
    missing: List[str] = []
    try:
        importlib.import_module("reportlab")
    except ImportError:
        missing.append("reportlab")
    try:
        importlib.import_module("pypdf")
    except ImportError:
        try:
            importlib.import_module("PyPDF2")
        except ImportError:
            missing.append("pypdf")
    try:
        importlib.import_module("cryptography")
    except ImportError:
        missing.append("cryptography")
    return missing


def validate_inputs(inputs: RunInputs, logger: StructuredLogger) -> None:
    logger.emit("VALIDATE_INPUTS", action="STARTED", result="STARTED", detail="Validating required and optional inputs.", phase_percent=0)
    required_paths = {
        "Document Version CSV": inputs.document_version_csv,
        "Corporate Document CSV": inputs.corporate_document_csv,
        "User CSV": inputs.user_csv,
        "Output Run Folder": inputs.output_run_folder,
    }
    for label, path in required_paths.items():
        if not path:
            raise RuntimeError(f"Required input missing: {label}")
        if label == "Output Run Folder":
            path.mkdir(parents=True, exist_ok=True)
        elif not path.exists():
            raise RuntimeError(f"Required input path does not exist: {label}: {path}")
        logger.emit(action="PATH_CHECK", result="OK", detail=f"{label}: {path}")

    optional_paths = {
        "Review Justification CSV": inputs.review_justification_csv,
        "Document Approvers CSV": inputs.approvers_file,
        "Document Reviewers CSV": inputs.reviewers_file,
        "Effective Version Details CSV": inputs.effective_version_details_csv,
        "Audit Activity Type CSV": inputs.audit_activity_type_csv,
        "Audit History CSV": inputs.audit_history_csv,
        "Neurotic DMS Manifest CSV": inputs.manifest_csv,
        "Raw DMS File Tree Folder": inputs.raw_dms_file_tree_folder,
    }
    for label, path in optional_paths.items():
        if path is None or not str(path).strip():
            continue
        if not path.exists() and inputs.fail_if_populated_optional_missing:
            raise RuntimeError(f"Optional path was populated but does not exist: {label}: {path}")
        if path.exists():
            logger.emit(action="OPTIONAL_PATH_CHECK", result="OK", detail=f"{label}: {path}")

    if inputs.generate_pdf_packages and (not inputs.raw_dms_file_tree_folder or not inputs.raw_dms_file_tree_folder.exists()):
        raise RuntimeError("Raw DMS File Tree Folder is required when PDF evidence packages are selected.")

    if inputs.generate_pdf_packages and inputs.include_pdf_cover_page:
        missing_pdf_deps = get_missing_pdf_dependencies(include_cover_page=True)
        if missing_pdf_deps:
            missing_text = ", ".join(missing_pdf_deps)
            logger.emit(
                action="PDF_DEPENDENCY_CHECK",
                result="FAILED",
                level="ERROR",
                detail=f"Missing PDF cover dependency/dependencies: {missing_text}. {PDF_DEPENDENCY_MESSAGE}",
            )
            raise RuntimeError(
                f"Missing PDF cover dependency/dependencies: {missing_text}. {PDF_DEPENDENCY_MESSAGE}"
            )
        logger.emit(
            action="PDF_DEPENDENCY_CHECK",
            result="OK",
            detail="PDF cover dependencies are available: reportlab and pypdf/PyPDF2 merge support.",
        )
    elif inputs.generate_pdf_packages:
        logger.emit(
            action="PDF_DEPENDENCY_CHECK",
            result="SKIPPED",
            detail="PDF cover page is unchecked; clean-named PDF copies do not require reportlab, pypdf, or PyPDF2.",
        )

    logger.emit(action="COMPLETE", result="OK", detail="Input validation complete.", phase_percent=100)


def create_run_outputs(base_output_folder: Path, generate_pdf_packages: bool) -> RunOutputs:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_id = f"{RUN_ID_PREFIX}-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
    package_run_id = f"{PACKAGE_RUN_ID_PREFIX}-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
    run_folder = base_output_folder / f"DMS_DOCUMENT_VERSION_HISTORY_RUN_{timestamp}"
    pdf_package_folder = run_folder / "01_Clean_Version_PDFs"
    register_folder = run_folder / "02_Register"
    package_log_folder = run_folder / "03_Package_Logs"
    register_folder.mkdir(parents=True, exist_ok=True)
    if generate_pdf_packages:
        pdf_package_folder.mkdir(parents=True, exist_ok=True)
        package_log_folder.mkdir(parents=True, exist_ok=True)
    workbook_path = register_folder / f"DMS_Document_Version_History_Register_{timestamp}.xlsx"
    event_log_csv_path = register_folder / f"DMS_Document_Version_History_Register_{timestamp}_{run_id}_event_log.csv"
    run_machine = platform.node() or os.environ.get("COMPUTERNAME", "") or os.environ.get("HOSTNAME", "")
    run_account = os.environ.get("USERDOMAIN", "")
    user_name = getpass.getuser() or os.environ.get("USERNAME", "") or os.environ.get("USER", "")
    run_account = f"{run_account}\\{user_name}" if run_account and user_name else user_name
    return RunOutputs(
        run_folder=run_folder,
        workbook_path=workbook_path,
        event_log_csv_path=event_log_csv_path,
        pdf_package_folder=pdf_package_folder,
        package_log_folder=package_log_folder,
        run_id=run_id,
        package_run_id=package_run_id,
        run_timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        run_machine=run_machine,
        run_account=run_account,
    )


def load_input_tables(inputs: RunInputs, logger: StructuredLogger, cancel: CancelToken) -> Tuple[Dict[str, pd.DataFrame], Dict[str, str]]:
    logger.emit("LOAD_INPUTS", action="STARTED", result="STARTED", detail="Loading raw TrackWise CSV exports.", phase_percent=0)
    tables: Dict[str, pd.DataFrame] = {}
    encodings: Dict[str, str] = {}
    load_plan: List[Tuple[str, Optional[Path], bool]] = [
        ("document_versions", inputs.document_version_csv, True),
        ("corporate_documents", inputs.corporate_document_csv, True),
        ("users", inputs.user_csv, True),
        ("review_justifications", inputs.review_justification_csv, False),
        ("approvers", inputs.approvers_file, False),
        ("reviewers", inputs.reviewers_file, False),
        ("effective_version_details", inputs.effective_version_details_csv, False),
        ("audit_activity_types", inputs.audit_activity_type_csv, False),
        ("audit_history", inputs.audit_history_csv, False),
        ("manifest", inputs.manifest_csv, False),
    ]
    total = len(load_plan)
    for position, (key, path, required) in enumerate(load_plan, start=1):
        cancel.check()
        if path is None or not str(path).strip() or not path.exists():
            tables[key] = pd.DataFrame()
            if required:
                raise RuntimeError(f"Required CSV path missing: {key}")
            logger.progress.update_phase_percent(position, total)
            logger.emit(action="LOAD_SKIPPED", result="SKIPPED", source_file=key, detail="Optional file not provided.")
            continue
        df, encoding = read_csv_with_fallback(path)
        tables[key] = df
        encodings[key] = encoding
        logger.progress.update_phase_percent(position, total)
        logger.emit(action="LOAD_CSV", result="OK", source_file=path.name, detail=f"{len(df):,} rows, {len(df.columns):,} columns, encoding={encoding}")

    ensure_required_columns(
        tables["document_versions"],
        ["Id", "SPARTADMS__Approved_Document_Number__c", "SPARTADMS__Related_Corporate_Document__c"],
        "SPARTADMS__Document_Version__c.csv",
    )
    ensure_required_columns(tables["corporate_documents"], ["Id"], "SPARTADMS__Corporate_Document__c.csv")
    ensure_required_columns(tables["users"], ["Id"], "User.csv")
    logger.emit(action="COMPLETE", result="OK", detail="All available CSV inputs loaded.", phase_percent=100)
    return tables, encodings


def build_indexes(tables: Dict[str, pd.DataFrame], logger: StructuredLogger, cancel: CancelToken) -> Dict[str, Any]:
    logger.emit("BUILD_INDEXES", action="STARTED", result="STARTED", detail="Building lookup indexes.", phase_percent=0)
    indexes: Dict[str, Any] = {}
    corporate_df = tables.get("corporate_documents", pd.DataFrame())
    user_df = tables.get("users", pd.DataFrame())
    indexes["users"] = build_user_map(user_df)
    logger.emit(action="USER_INDEX", result="OK", detail=f"{len(indexes['users']):,} Salesforce user IDs indexed.", phase_percent=25)

    corporate_by_id: Dict[str, pd.Series] = {}
    if "Id" in corporate_df.columns:
        for row_tuple in corporate_df.itertuples(index=False):
            cancel.check()
            row = pd.Series(row_tuple._asdict())
            corp_id = clean_scalar(row.get("Id", "")).strip()
            if corp_id:
                corporate_by_id[corp_id] = row
    indexes["corporate_by_id"] = corporate_by_id
    logger.emit(action="CORPORATE_INDEX", result="OK", detail=f"{len(corporate_by_id):,} corporate document IDs indexed.", phase_percent=60)

    audit_types_df = tables.get("audit_activity_types", pd.DataFrame())
    effective_to_obsolete_type_ids: set[str] = set()
    if not audit_types_df.empty and {"Id", "Name"}.issubset(set(audit_types_df.columns)):
        for row_tuple in audit_types_df.itertuples(index=False):
            row = pd.Series(row_tuple._asdict())
            if clean_scalar(row.get("Name", "")).strip().lower() == "effective to obsolete":
                effective_to_obsolete_type_ids.add(clean_scalar(row.get("Id", "")).strip())
    indexes["effective_to_obsolete_type_ids"] = effective_to_obsolete_type_ids
    logger.emit(action="AUDIT_ACTIVITY_INDEX", result="OK", detail=f"{len(effective_to_obsolete_type_ids):,} Effective to Obsolete activity type IDs indexed.", phase_percent=85)
    logger.emit(action="COMPLETE", result="OK", detail="Lookup indexes built.", phase_percent=100)
    return indexes


def build_source_version_rows(
    document_versions_df: pd.DataFrame,
    indexes: Dict[str, Any],
    logger: StructuredLogger,
    cancel: CancelToken,
) -> pd.DataFrame:
    logger.emit("BUILD_SOURCE_VERSION_ROWS", action="STARTED", result="STARTED", detail=f"{len(document_versions_df):,} raw document-version rows to evaluate.", phase_percent=0)
    corporate_by_id: Dict[str, pd.Series] = indexes.get("corporate_by_id", {})
    user_map: Dict[str, str] = indexes.get("users", {})
    rows: List[Dict[str, Any]] = []
    total = len(document_versions_df)
    for row_number, row_tuple in enumerate(document_versions_df.itertuples(index=False), start=2):
        cancel.check()
        raw = pd.Series(row_tuple._asdict())
        corporate_id = first_present(raw, ["SPARTADMS__Related_Corporate_Document__c", "SPARTADMS__Related_Document__c"])
        corporate_row = corporate_by_id.get(corporate_id)
        if corporate_row is None:
            corporate_row = pd.Series(dtype=object)
        raw_number = first_present(raw, ["SPARTADMS__Approved_Document_Number__c", "SPARTADMS__Document_Number_CD_Formula__c"])
        if not raw_number:
            raw_number = first_present(corporate_row, ["SPARTADMS__Approved_Document_Number__c", "SPARTADMS__Document_Number_CD_Formula__c"])
        document_type = first_present(raw, ["SPARTADMS__Document_Type__c"]) or first_present(corporate_row, ["SPARTADMS__Document_Type__c"])
        document_number_field = first_present(raw, ["SPARTADMS__Document_Number__c"]) or first_present(corporate_row, ["SPARTADMS__Document_Number__c"])
        display_document, raw_base, inferred_sequence = build_display_document_number(raw_number, document_type, document_number_field)
        explicit_sequence = first_present(raw, ["SPARTADMS__Document_Sequence_Number__c"]) or first_present(corporate_row, ["SPARTADMS__Document_Sequence_Number__c"])
        source_sequence = normalize_sequence(explicit_sequence or inferred_sequence)
        version = normalize_version(
            first_present(
                raw,
                [
                    "SPARTADMS__Document_Revision_At_Effective__c",
                    "SPARTADMS__Document_Revision_At_Approval__c",
                    "SPARTADMS__Document_Version__c",
                    "SPARTADMS__Document_Number_Major__c",
                ],
            )
        )
        source_version_key = f"{raw_number}.{version}" if raw_number and version else raw_number
        full_document_version = f"{display_document}.{version}" if display_document and version else display_document
        row = {
            "Source Row Number": row_number,
            "Raw Document Version ID": first_present(raw, ["Id"]),
            "Raw Corporate Document ID": corporate_id,
            "Raw TrackWise Document Number": raw_number,
            "Raw TrackWise Document Base": raw_base,
            "Source Document Sequence Number": source_sequence,
            "Source Version Key": source_version_key,
            "Document Number": display_document,
            "Version": version,
            "Full Document Version": full_document_version,
            "Document Title": first_present(raw, ["Name"]) or first_present(corporate_row, ["Name"]),
            "Document Type": document_type,
            "Department": first_present(raw, ["SPARTADMS__Document_Department__c"]) or first_present(corporate_row, ["SPARTADMS__Document_Department__c"]),
            "Document Status": first_present(raw, ["SPARTADMS__Document_Status__c"]) or first_present(corporate_row, ["SPARTADMS__Document_Status__c"]),
            "Version Status": first_present(raw, ["SPARTADMS__Version_Status__c", "SPARTADMS__Status__c"]),
            "Approved Date": format_datetime(first_present(raw, ["SPARTADMS__Approved_Date_Time__c"])),
            "Effective Date": format_datetime(first_present(raw, ["SPARTADMS__Effective_Datetime__c", "SPARTADMS__Effective_DT__c"])),
            "Created Date": format_datetime(first_present(raw, ["CreatedDate"])),
            "Created By": decode_user(first_present(raw, ["CreatedById"]), user_map),
            "Last Modified Date": format_datetime(first_present(raw, ["LastModifiedDate"])),
            "Last Modified By": decode_user(first_present(raw, ["LastModifiedById"]), user_map),
            "Related Quality Event": first_present(raw, ["SPARTADMS__Related_Quality_Event__c", "SPARTADMS__Quality_Event_Reference__c"]),
            "Training Required": first_present(raw, ["SPARTADMS__Requires_Training__c"]) or first_present(corporate_row, ["SPARTADMS__Requires_Training__c"]),
            "Latest File Name": first_present(raw, ["SPARTADMS__Latest_file_name__c", "SPARTADMS__File_Name__c"]),
            "Document Version Numeric": normalize_version_int(version),
        }
        rows.append(row)
        completed = row_number - 1
        logger.progress_event(completed, total, action="SOURCE_VERSION_ROWS", every=500, detail=f"{completed:,} / {total:,}")
    logger.emit(action="COMPLETE", result="OK", detail=f"{len(rows):,} source version rows built.", phase_percent=100)
    return pd.DataFrame(rows)


def normalize_sequence(value: Any) -> str:
    text = clean_scalar(value).strip()
    if not text:
        return ""
    numeric = pd.to_numeric(text, errors="coerce")
    if not pd.isna(numeric):
        return str(int(float(numeric)))
    return text


def extract_versions_from_text(text: str, document_number: str, raw_base: str) -> List[int]:
    content = clean_scalar(text)
    if not content:
        return []
    candidates: List[int] = []
    doc_tokens = {document_number.upper(), raw_base.upper(), document_number.replace("-", "").upper(), raw_base.replace("-", "").upper()}
    pattern = re.compile(r"([A-Z]{0,6}[-_]?\d{3,6}[A-Z]?)\s*[.]\s*(\d{1,3})", re.I)
    for match in pattern.finditer(content):
        token = match.group(1).replace("_", "-").upper()
        token_no_dash = token.replace("-", "")
        if token in doc_tokens or token_no_dash in doc_tokens:
            candidates.append(int(match.group(2)))
    if not candidates:
        plain_version = re.compile(r"\bversion\s*(\d{1,3})\b", re.I)
        for match in plain_version.finditer(content):
            candidates.append(int(match.group(1)))
    return candidates


def map_justifications(
    review_df: pd.DataFrame,
    source_rows_df: pd.DataFrame,
    user_map: Dict[str, str],
    logger: StructuredLogger,
    cancel: CancelToken,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    logger.emit("MAP_JUSTIFICATIONS", action="STARTED", result="STARTED", detail=f"{len(review_df):,} justification records to evaluate.", phase_percent=0)
    if review_df.empty:
        logger.emit(action="COMPLETE", result="OK", detail="No review justification CSV provided.", phase_percent=100)
        return pd.DataFrame(), source_rows_df.copy()

    source_rows = source_rows_df.copy()
    diagnostic_rows: List[Dict[str, Any]] = []
    candidate_by_corp: Dict[str, pd.DataFrame] = {corp_id: group.copy() for corp_id, group in source_rows.groupby("Raw Corporate Document ID", dropna=False)}
    mapped_by_full_doc: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    total = len(review_df)
    mapped_count = 0
    unmapped_count = 0

    for position, row_tuple in enumerate(review_df.itertuples(index=False), start=1):
        cancel.check()
        jrow = pd.Series(row_tuple._asdict())
        justification_id = first_present(jrow, ["Name", "Id"])
        corp_id = first_present(jrow, ["SPARTADMS__Related_Corporate_Document__c"])
        content = first_present(jrow, ["SPARTADMS__Justification__c"])
        explicit_version = normalize_version_int(first_present(jrow, ["SPARTADMS__Unformatted_Document_Version__c", "SPARTADMS__Document_Version__c"]))
        created_ts = as_timestamp(first_present(jrow, ["CreatedDate"]))
        candidates = candidate_by_corp.get(corp_id, pd.DataFrame())
        mapping_method = "UNMAPPED"
        confidence = "None"
        target_full_doc = ""
        target_version = ""
        target_doc_number = ""
        detail = "No same-corporate-document version candidate found."

        if not candidates.empty:
            direct_matches = pd.DataFrame()
            if explicit_version is not None:
                direct_matches = candidates[candidates["Document Version Numeric"] == explicit_version]
            if not direct_matches.empty:
                target = select_best_candidate(direct_matches, created_ts)
                mapping_method = "DIRECT_VERSION_FIELD"
                confidence = "High"
                detail = "Mapped by same Corporate Document ID and explicit version field."
            else:
                parsed_matches: List[pd.Series] = []
                for _, candidate in candidates.iterrows():
                    parsed_versions = extract_versions_from_text(
                        content,
                        clean_scalar(candidate.get("Document Number", "")),
                        clean_scalar(candidate.get("Raw TrackWise Document Base", "")),
                    )
                    candidate_version = normalize_version_int(candidate.get("Version", ""))
                    if candidate_version is not None and candidate_version in parsed_versions:
                        parsed_matches.append(candidate)
                if parsed_matches:
                    target = select_best_candidate(pd.DataFrame(parsed_matches), created_ts)
                    mapping_method = "TEXT_VERSION_PARSE"
                    confidence = "High"
                    detail = "Mapped by same Corporate Document ID and exact version reference in justification text."
                else:
                    target = select_best_candidate(candidates, created_ts)
                    if target is not None:
                        mapping_method = "DATE_NEAREST_CORPORATE_MATCH"
                        confidence = "Medium"
                        detail = "Mapped by same Corporate Document ID and nearest lifecycle date because no explicit version was available."
                    else:
                        target = None
            if target is not None:
                target_full_doc = clean_scalar(target.get("Full Document Version", ""))
                target_version = clean_scalar(target.get("Version", ""))
                target_doc_number = clean_scalar(target.get("Document Number", ""))
                mapped_by_full_doc[target_full_doc].append({
                    "Justification ID": justification_id,
                    "Justification Content": content,
                    "Justification Created Date": format_datetime(first_present(jrow, ["CreatedDate"])),
                    "Justification Created By": decode_user(first_present(jrow, ["CreatedById"]), user_map),
                    "Justification Owner": decode_user(first_present(jrow, ["OwnerId"]), user_map),
                    "Justification Last Modified Date": format_datetime(first_present(jrow, ["LastModifiedDate"])),
                    "Justification Last Modified By": decode_user(first_present(jrow, ["LastModifiedById"]), user_map),
                    "Justification Source Salesforce ID": first_present(jrow, ["Id"]),
                    "Justification Related Corporate Document ID": corp_id,
                    "Justification Raw Version Field": first_present(jrow, ["SPARTADMS__Document_Version__c"]),
                    "Justification Raw Unformatted Version Field": first_present(jrow, ["SPARTADMS__Unformatted_Document_Version__c"]),
                    "Justification Mapping Method": mapping_method,
                    "Justification Confidence": confidence,
                    "Justification Mapping Detail": detail,
                })
                mapped_count += 1
            else:
                unmapped_count += 1
        else:
            unmapped_count += 1

        diagnostic_rows.append({
            "Justification ID": justification_id,
            "Raw Justification ID": first_present(jrow, ["Id"]),
            "Related Corporate Document ID": corp_id,
            "Explicit Version Field": clean_scalar(explicit_version),
            "Mapped Full Document Version": target_full_doc,
            "Mapped Document Number": target_doc_number,
            "Mapped Version": target_version,
            "Mapping Method": mapping_method,
            "Mapping Confidence": confidence,
            "Justification Created Date": format_datetime(first_present(jrow, ["CreatedDate"])),
            "Justification Created By": decode_user(first_present(jrow, ["CreatedById"]), user_map),
            "Justification Owner": decode_user(first_present(jrow, ["OwnerId"]), user_map),
            "Justification Last Modified Date": format_datetime(first_present(jrow, ["LastModifiedDate"])),
            "Justification Last Modified By": decode_user(first_present(jrow, ["LastModifiedById"]), user_map),
            "Justification Content": content,
            "Diagnostic Detail": detail,
        })
        logger.progress.update_phase_percent(position, total)
        # Keep the live pop-out readable. Full per-justification detail is preserved in
        # the Justification Mapping Diagnostics sheet and event-log CSV; the live viewer
        # only shows periodic progress plus the SOP-1022 acceptance-test mappings.
        if mapping_method != "UNMAPPED" and target_full_doc.startswith("SOP-1022."):
            logger.emit(
                action="MAP_JUSTIFICATION",
                result="OK",
                detail="SOP-1022 acceptance-test mapping confirmed",
                full_document_version=target_full_doc,
                justification_id=justification_id,
                raw_corporate_document_id=corp_id,
                document_number=target_doc_number,
                version=target_version,
                mapping_method=mapping_method,
                mapping_confidence=confidence,
            )
        elif position % 500 == 0 or position == total:
            logger.emit(action="MAP_JUSTIFICATIONS", result="PROGRESS", detail=f"{position:,} / {total:,} evaluated; {mapped_count:,} mapped; {unmapped_count:,} unmapped")

    for column in [
        "Justification ID",
        "Justification Content",
        "Justification Created Date",
        "Justification Created By",
        "Justification Owner",
        "Justification Last Modified Date",
        "Justification Last Modified By",
        "Justification Source Salesforce ID",
        "Justification Related Corporate Document ID",
        "Justification Raw Version Field",
        "Justification Raw Unformatted Version Field",
        "Justification Mapping Method",
        "Justification Confidence",
        "Justification Mapping Detail",
    ]:
        source_rows[column] = ""

    for index, row in source_rows.iterrows():
        full_doc = clean_scalar(row.get("Full Document Version", ""))
        mapped_items = mapped_by_full_doc.get(full_doc, [])
        if not mapped_items:
            continue
        source_rows.at[index, "Justification ID"] = join_unique(item["Justification ID"] for item in mapped_items)
        source_rows.at[index, "Justification Content"] = "\n\n---\n\n".join(item["Justification Content"] for item in mapped_items if item["Justification Content"])
        source_rows.at[index, "Justification Created Date"] = join_unique(item["Justification Created Date"] for item in mapped_items)
        source_rows.at[index, "Justification Created By"] = join_unique(item["Justification Created By"] for item in mapped_items)
        source_rows.at[index, "Justification Owner"] = join_unique(item["Justification Owner"] for item in mapped_items)
        source_rows.at[index, "Justification Last Modified Date"] = join_unique(item["Justification Last Modified Date"] for item in mapped_items)
        source_rows.at[index, "Justification Last Modified By"] = join_unique(item["Justification Last Modified By"] for item in mapped_items)
        source_rows.at[index, "Justification Source Salesforce ID"] = join_unique(item.get("Justification Source Salesforce ID", "") for item in mapped_items)
        source_rows.at[index, "Justification Related Corporate Document ID"] = join_unique(item.get("Justification Related Corporate Document ID", "") for item in mapped_items)
        source_rows.at[index, "Justification Raw Version Field"] = join_unique(item.get("Justification Raw Version Field", "") for item in mapped_items)
        source_rows.at[index, "Justification Raw Unformatted Version Field"] = join_unique(item.get("Justification Raw Unformatted Version Field", "") for item in mapped_items)
        source_rows.at[index, "Justification Mapping Method"] = join_unique(item["Justification Mapping Method"] for item in mapped_items)
        source_rows.at[index, "Justification Confidence"] = join_unique(item["Justification Confidence"] for item in mapped_items)
        source_rows.at[index, "Justification Mapping Detail"] = join_unique(item.get("Justification Mapping Detail", "") for item in mapped_items)

    logger.emit(action="COMPLETE", result="OK", detail=f"{total:,} evaluated, {mapped_count:,} mapped, {unmapped_count:,} unmapped.", phase_percent=100)
    return pd.DataFrame(diagnostic_rows), source_rows


def select_best_candidate(candidates: pd.DataFrame, reference_date: Optional[pd.Timestamp]) -> Optional[pd.Series]:
    if candidates.empty:
        return None
    if reference_date is None:
        return candidates.iloc[0]
    best_label = candidates.index[0]
    best_delta: Optional[float] = None
    date_columns = ["Approved Date", "Effective Date", "Created Date", "Last Modified Date"]
    for label, row in candidates.iterrows():
        row_best: Optional[float] = None
        for column in date_columns:
            timestamp = as_timestamp(row.get(column, ""))
            if timestamp is None:
                continue
            delta = abs((timestamp - reference_date).total_seconds())
            if row_best is None or delta < row_best:
                row_best = delta
        if row_best is not None and (best_delta is None or row_best < best_delta):
            best_delta = row_best
            best_label = label
    return candidates.loc[best_label]


def build_actor_detail(
    actor_df: pd.DataFrame,
    source_rows_df: pd.DataFrame,
    user_map: Dict[str, str],
    actor_type: str,
) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]]]:
    if actor_df.empty:
        return pd.DataFrame(), {}
    corp_to_versions = source_rows_df.groupby("Raw Corporate Document ID", dropna=False)["Full Document Version"].apply(lambda values: join_unique(values)).to_dict()
    rows: List[Dict[str, Any]] = []
    names_by_corp: Dict[str, List[str]] = defaultdict(list)
    details_by_corp: Dict[str, List[str]] = defaultdict(list)
    dates_by_corp: Dict[str, List[str]] = defaultdict(list)
    if actor_type == "approver":
        name_cols = ["SPARTADMS__ApproverName__c", "SPARTADMS__Related_Box_User__c", "CreatedById"]
        completed_col = "SPARTADMS__Approval_Completed_Date__c"
        submitted_col = "SPARTADMS__Approval_Submitted_Date__c"
        cycle_col = "SPARTADMS__Approval_Cycle_Label__c"
        details_label = "Approval"
    else:
        name_cols = ["SPARTADMS__ReviewerName__c", "SPARTADMS__Related_Box_User__c", "CreatedById"]
        completed_col = "SPARTADMS__Review_Completed_Date__c"
        submitted_col = "SPARTADMS__Review_Submission_Date__c"
        cycle_col = "SPARTADMS__Review_Cycle_Label__c"
        details_label = "Review"

    columns = list(actor_df.columns)
    for offset, values in enumerate(actor_df.itertuples(index=False, name=None), start=2):
        row = dict(zip(columns, values))
        corp_id = first_present_dict(row, ["SPARTADMS__Related_Corporate_Document__c"])
        actor_name = first_present_dict(row, name_cols)
        actor_name = decode_user(actor_name, user_map)
        completed = format_datetime(first_present_dict(row, [completed_col]))
        submitted = format_datetime(first_present_dict(row, [submitted_col]))
        cycle = first_present_dict(row, [cycle_col])
        record_name = first_present_dict(row, ["Name", "Id"])
        detail = f"{details_label} record {record_name}; cycle={cycle}; submitted={submitted}; completed={completed}; mapping=CORPORATE_DOCUMENT_LEVEL / Medium"
        rows.append({
            "Source Row Number": offset,
            "Actor Type": actor_type.title(),
            "Actor Record ID": first_present_dict(row, ["Id"]),
            "Actor Record Name": first_present_dict(row, ["Name"]),
            "Related Corporate Document ID": corp_id,
            "Related Display Versions": corp_to_versions.get(corp_id, ""),
            "Actor Name": actor_name,
            "Cycle Label": cycle,
            "Submitted Date": submitted,
            "Completed Date": completed,
            "Mandatory": first_present_dict(row, ["SPARTADMS__Mandatory__c"]),
            "Final Approver": first_present_dict(row, ["SPARTADMS__Is_Final_Approver__c"]),
            "Mapping Method": "CORPORATE_DOCUMENT_LEVEL",
            "Mapping Confidence": "Medium",
            "Detail": detail,
        })
        if actor_name:
            names_by_corp[corp_id].append(actor_name)
        if detail:
            details_by_corp[corp_id].append(detail)
        if completed:
            dates_by_corp[corp_id].append(completed)

    summary_by_corp: Dict[str, Dict[str, str]] = {}
    for corp_id in set(names_by_corp) | set(details_by_corp) | set(dates_by_corp):
        summary_by_corp[corp_id] = {
            "names": join_unique(names_by_corp.get(corp_id, [])),
            "details": summarize_detail_list(details_by_corp.get(corp_id, []), limit=25),
            "date": join_unique(dates_by_corp.get(corp_id, [])),
        }
    return pd.DataFrame(rows), summary_by_corp


def first_present_dict(row: Dict[str, Any], columns: Sequence[str]) -> str:
    for column in columns:
        value = clean_scalar(row.get(column, "")).strip()
        if value:
            return value
    return ""


def build_obsoletion_detail(
    audit_df: pd.DataFrame,
    source_rows_df: pd.DataFrame,
    indexes: Dict[str, Any],
    user_map: Dict[str, str],
    logger: StructuredLogger,
    cancel: CancelToken,
) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]]]:
    effective_type_ids: set[str] = indexes.get("effective_to_obsolete_type_ids", set())
    status_rows: List[Dict[str, Any]] = []
    summary_by_corp: Dict[str, Dict[str, str]] = defaultdict(lambda: {"date": "", "by": "", "details": ""})
    if audit_df.empty or not effective_type_ids:
        return pd.DataFrame(), summary_by_corp
    total = len(audit_df)
    logger.emit(action="OBSOLETION_SCAN", result="STARTED", detail=f"Scanning {total:,} audit history rows.")
    obsolete_source_corps = set(
        source_rows_df.loc[source_rows_df["Version Status"].str.contains("obsolete", case=False, na=False), "Raw Corporate Document ID"].dropna().astype(str)
    )
    for row_number, row_tuple in enumerate(audit_df.itertuples(index=False), start=2):
        cancel.check()
        row = pd.Series(row_tuple._asdict())
        if first_present(row, ["CMPL123__Activity_Type__c"]) not in effective_type_ids:
            if row_number % 10000 == 0:
                logger.emit(action="OBSOLETION_SCAN", result="PROGRESS", detail=f"{row_number - 1:,} / {total:,} audit rows scanned")
            continue
        related_id = first_present(row, ["CMPL123__Related_Record_ID__c"])
        if obsolete_source_corps and related_id not in obsolete_source_corps:
            continue
        date_posted = format_datetime(first_present(row, ["CMPL123__Date_Posted__c", "CreatedDate"] ))
        responsible = decode_user(first_present(row, ["CMPL123__Responsible_User_ID__c", "CreatedById"]), user_map)
        summary = first_present(row, ["CMPL123__Summary__c"])
        comments = first_present(row, ["CMPL123__Comments__c"])
        detail = join_unique([summary, comments], separator=" | ")
        status_rows.append({
            "Audit History ID": first_present(row, ["Id"]),
            "Audit History Name": first_present(row, ["Name"]),
            "Related Corporate Document ID": related_id,
            "Activity Type": "Effective to Obsolete",
            "Date Posted": date_posted,
            "Responsible User": responsible,
            "E-Signature Applied": first_present(row, ["CMPL123__ESign_Applied__c"]),
            "Summary": summary,
            "Comments": comments,
            "Mapping Method": "CORPORATE_DOCUMENT_MATCH",
            "Mapping Confidence": "Medium",
        })
        current = summary_by_corp[related_id]
        current["date"] = join_unique([current["date"], date_posted])
        current["by"] = join_unique([current["by"], responsible])
        current["details"] = excel_cell(join_unique([current["details"], detail], separator="\n"), 6000)
    logger.emit(action="OBSOLETION_SCAN", result="OK", detail=f"{len(status_rows):,} Effective to Obsolete audit events matched.")
    return pd.DataFrame(status_rows), summary_by_corp


def normalize_lifecycle_status(document_status: str, version_status: str) -> str:
    text = f"{document_status} {version_status}".lower()
    if "cancel" in text:
        return "CANCELED"
    if "obsolete" in text or "retired" in text:
        return "OBSOLETE"
    if "effective" in text or "approved" in text:
        return "EFFECTIVE"
    if "draft" in text or "review" in text or "approval" in text:
        return "IN_WORKFLOW"
    return clean_scalar(version_status or document_status).upper()


def build_main_history_and_supporting_sheets(
    source_rows_df: pd.DataFrame,
    approver_detail_df: pd.DataFrame,
    reviewer_detail_df: pd.DataFrame,
    status_event_df: pd.DataFrame,
    approver_summary: Dict[str, Dict[str, str]],
    reviewer_summary: Dict[str, Dict[str, str]],
    obsoletion_summary: Dict[str, Dict[str, str]],
    logger: StructuredLogger,
    cancel: CancelToken,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    logger.emit("BUILD_WORKBOOK_TABLES", action="STARTED", result="STARTED", detail="Building QA-facing workbook tables.", phase_percent=0)
    rows: List[Dict[str, Any]] = []
    identity_rows: List[Dict[str, Any]] = []
    grouped = source_rows_df.groupby(["Document Number", "Version"], dropna=False, sort=True)
    total = len(grouped)
    for position, ((document_number, version), group) in enumerate(grouped, start=1):
        cancel.check()
        first = group.iloc[0]
        corp_ids = list(dict.fromkeys(clean_scalar(v) for v in group["Raw Corporate Document ID"].tolist() if clean_scalar(v)))
        full_document_version = clean_scalar(first.get("Full Document Version", ""))
        approver_names = join_unique(approver_summary.get(corp_id, {}).get("names", "") for corp_id in corp_ids)
        approver_details = excel_cell(join_unique(approver_summary.get(corp_id, {}).get("details", "") for corp_id in corp_ids), 8000)
        reviewer_names = join_unique(reviewer_summary.get(corp_id, {}).get("names", "") for corp_id in corp_ids)
        reviewer_details = excel_cell(join_unique(reviewer_summary.get(corp_id, {}).get("details", "") for corp_id in corp_ids), 8000)
        reviewer_dates = join_unique(reviewer_summary.get(corp_id, {}).get("date", "") for corp_id in corp_ids)
        obsolete_dates = join_unique(obsoletion_summary.get(corp_id, {}).get("date", "") for corp_id in corp_ids)
        obsoleted_by = join_unique(obsoletion_summary.get(corp_id, {}).get("by", "") for corp_id in corp_ids)
        obsolete_details = excel_cell(join_unique(obsoletion_summary.get(corp_id, {}).get("details", "") for corp_id in corp_ids), 8000)
        document_status = join_unique(group["Document Status"].tolist())
        version_status = join_unique(group["Version Status"].tolist())
        lifecycle_status = normalize_lifecycle_status(document_status, version_status)
        justification_content = excel_cell("\n\n---\n\n".join(v for v in group.get("Justification Content", pd.Series(dtype=str)).tolist() if clean_scalar(v)))
        source_note_parts = []
        if len(group) > 1:
            source_note_parts.append(f"Aggregated from {len(group)} raw source version rows.")
        if len(set(clean_scalar(v) for v in group["Raw TrackWise Document Number"].tolist())) > 1:
            source_note_parts.append("Multiple raw TrackWise document numbers contribute to this display version.")
        next_version = ""
        row = {
            "Document Number": document_number,
            "Version": version,
            "Full Document Version": full_document_version,
            "Output PDF Filename": safe_pdf_name(full_document_version) + ".pdf" if full_document_version else "",
            "Output PDF Path": "",
            "Document Title": join_unique(group["Document Title"].tolist()),
            "Document Type": join_unique(group["Document Type"].tolist()),
            "Department": join_unique(group["Department"].tolist()),
            "Document Status": document_status,
            "Version Status": version_status,
            "Lifecycle Status Normalized": lifecycle_status,
            "Approved Date": join_unique(group["Approved Date"].tolist()),
            "Effective Date": join_unique(group["Effective Date"].tolist()),
            "Obsolete Date": obsolete_dates,
            "Obsolete / Superseded Date Basis": "Direct audit history event" if obsolete_dates else "",
            "Canceled Date": "",
            "Canceled Date Basis": "",
            "Superseded By": next_version,
            "Justification ID": join_unique(group.get("Justification ID", pd.Series(dtype=str)).tolist()),
            "Justification Content": justification_content,
            "Justification Created Date": join_unique(group.get("Justification Created Date", pd.Series(dtype=str)).tolist()),
            "Justification Created By": join_unique(group.get("Justification Created By", pd.Series(dtype=str)).tolist()),
            "Justification Owner": join_unique(group.get("Justification Owner", pd.Series(dtype=str)).tolist()),
            "Justification Last Modified Date": join_unique(group.get("Justification Last Modified Date", pd.Series(dtype=str)).tolist()),
            "Justification Last Modified By": join_unique(group.get("Justification Last Modified By", pd.Series(dtype=str)).tolist()),
            "Justification Source Salesforce ID": join_unique(group.get("Justification Source Salesforce ID", pd.Series(dtype=str)).tolist()),
            "Justification Related Corporate Document ID": join_unique(group.get("Justification Related Corporate Document ID", pd.Series(dtype=str)).tolist()),
            "Justification Raw Version Field": join_unique(group.get("Justification Raw Version Field", pd.Series(dtype=str)).tolist()),
            "Justification Raw Unformatted Version Field": join_unique(group.get("Justification Raw Unformatted Version Field", pd.Series(dtype=str)).tolist()),
            "Justification Mapping Method": join_unique(group.get("Justification Mapping Method", pd.Series(dtype=str)).tolist()),
            "Justification Confidence": join_unique(group.get("Justification Confidence", pd.Series(dtype=str)).tolist()),
            "Justification Mapping Detail": join_unique(group.get("Justification Mapping Detail", pd.Series(dtype=str)).tolist()),
            "Reviewed Date": reviewer_dates,
            "Reviewed By": reviewer_names,
            "Review Details": reviewer_details,
            "Approved By": approver_names,
            "Approval Details": approver_details,
            "Obsoleted By": obsoleted_by,
            "Obsolete Details": obsolete_details,
            "Canceled By": "",
            "Canceled Details": "",
            "Previous Display Version in Register": "",
            "Next Display Version in Register": "",
            "Created Date": join_unique(group["Created Date"].tolist()),
            "Created By": join_unique(group["Created By"].tolist()),
            "Last Modified Date": join_unique(group["Last Modified Date"].tolist()),
            "Last Modified By": join_unique(group["Last Modified By"].tolist()),
            "Related Quality Event": join_unique(group["Related Quality Event"].tolist()),
            "Training Required": join_unique(group["Training Required"].tolist()),
            "Source Version Key(s)": join_unique(group["Source Version Key"].tolist()),
            "Raw TrackWise Document Number(s)": join_unique(group["Raw TrackWise Document Number"].tolist()),
            "Source Document Sequence Number(s)": join_unique(group["Source Document Sequence Number"].tolist()),
            "Raw Document Version ID(s)": join_unique(group["Raw Document Version ID"].tolist()),
            "Raw Corporate Document ID(s)": join_unique(group["Raw Corporate Document ID"].tolist()),
            "Source PDF Path(s)": "",
            "Source PDF SHA-256": "",
            "Output PDF SHA-256": "",
            "PDF Packaging Result": "Not run",
            "PDF Packaging Detail": "PDF generation not yet run or not selected.",
            "Source Notes": " ".join(source_note_parts),
        }
        rows.append(row)
        identity_rows.append({
            "Document Number": document_number,
            "Version": version,
            "Full Document Version": full_document_version,
            "Raw Source Row Count": len(group),
            "Source Version Key(s)": join_unique(group["Source Version Key"].tolist()),
            "Raw TrackWise Document Number(s)": join_unique(group["Raw TrackWise Document Number"].tolist()),
            "Source Document Sequence Number(s)": join_unique(group["Source Document Sequence Number"].tolist()),
            "Raw Corporate Document ID(s)": join_unique(group["Raw Corporate Document ID"].tolist()),
            "Collision / Diagnostic Note": "Multiple raw source rows for one display document/version." if len(group) > 1 else "",
        })
        logger.progress_event(position, total, action="BUILD_HISTORY_ROWS", every=500, detail=f"{position:,} / {total:,}")
    history_df = pd.DataFrame(rows, columns=MAIN_COLUMNS)
    if not history_df.empty:
        history_df["__version_sort"] = history_df["Version"].map(lambda value: version_sort_key(value)[0])
        history_df = history_df.sort_values(["Document Number", "__version_sort", "Version"]).drop(columns=["__version_sort"]).reset_index(drop=True)
    history_df = add_previous_next_versions(history_df)
    if not history_df.empty:
        history_df["Superseded By"] = history_df["Next Display Version in Register"].where(history_df["Lifecycle Status Normalized"].isin(["OBSOLETE", "CANCELED"]), "")
        effective_by_full_doc = {clean_scalar(row.get("Full Document Version", "")): clean_scalar(row.get("Effective Date", "")) for _, row in history_df.iterrows()}
        for idx, hist_row in history_df.iterrows():
            superseded_by = clean_scalar(hist_row.get("Superseded By", ""))
            obsolete_date = clean_scalar(hist_row.get("Obsolete Date", ""))
            if superseded_by and not obsolete_date:
                derived_date = effective_by_full_doc.get(superseded_by, "")
                if derived_date:
                    history_df.at[idx, "Obsolete Date"] = derived_date
                    history_df.at[idx, "Obsolete / Superseded Date Basis"] = "Derived from next effective version"
                else:
                    history_df.at[idx, "Obsolete / Superseded Date Basis"] = "No direct obsolete event or next effective version date available in export"
            elif obsolete_date:
                history_df.at[idx, "Obsolete / Superseded Date Basis"] = clean_scalar(hist_row.get("Obsolete / Superseded Date Basis", "")) or "Direct audit history event"
            elif clean_scalar(hist_row.get("Lifecycle Status Normalized", "")) == "EFFECTIVE":
                history_df.at[idx, "Obsolete / Superseded Date Basis"] = "Not applicable; current/final effective version"
            if not clean_scalar(hist_row.get("Canceled Date", "")):
                history_df.at[idx, "Canceled Date Basis"] = "No cancellation date available in source export"
    identity_diag_df = pd.DataFrame(identity_rows)
    technical_df = source_rows_df.copy()
    stats = {
        "source_version_rows": len(source_rows_df),
        "display_history_rows": len(history_df),
        "approver_detail_rows": len(approver_detail_df),
        "reviewer_detail_rows": len(reviewer_detail_df),
        "status_event_rows": len(status_event_df),
        "identity_diagnostic_rows": len(identity_diag_df),
        "technical_evidence_rows": len(technical_df),
    }
    logger.emit(action="COMPLETE", result="OK", detail=f"{len(history_df):,} display document-version rows built.", phase_percent=100)
    return history_df, approver_detail_df, reviewer_detail_df, status_event_df, identity_diag_df, technical_df, stats


def add_previous_next_versions(history_df: pd.DataFrame) -> pd.DataFrame:
    if history_df.empty:
        return history_df
    output = history_df.copy()
    for _, group in output.groupby("Document Number", sort=False):
        ordered = group.copy()
        ordered["__version_sort"] = ordered["Version"].map(lambda value: version_sort_key(value)[0])
        ordered = ordered.sort_values(["__version_sort", "Version"])
        full_docs = ordered["Full Document Version"].tolist()
        indexes = ordered.index.tolist()
        for position, index_value in enumerate(indexes):
            output.at[index_value, "Previous Display Version in Register"] = full_docs[position - 1] if position > 0 else ""
            output.at[index_value, "Next Display Version in Register"] = full_docs[position + 1] if position + 1 < len(full_docs) else ""
    return output



def first_date_from_cell(value: Any) -> str:
    text = clean_scalar(value)
    if not text:
        return ""
    for part in re.split(r";|\n", text):
        candidate = format_datetime(part.strip())
        if candidate:
            return candidate
    return text.strip()


def build_document_version_timeline(history_df: pd.DataFrame, source_rows_df: pd.DataFrame, logger: StructuredLogger) -> pd.DataFrame:
    """Build Sheet 2: a chronological event view of each document version."""
    rows: List[Dict[str, Any]] = []
    sequence = 1
    if history_df.empty:
        return pd.DataFrame(columns=TIMELINE_COLUMNS)
    for _, row in history_df.iterrows():
        full_doc = clean_scalar(row.get("Full Document Version", ""))
        doc_number = clean_scalar(row.get("Document Number", ""))
        version = clean_scalar(row.get("Version", ""))
        source_keys = clean_scalar(row.get("Source Version Key(s)", ""))
        j_id = clean_scalar(row.get("Justification ID", ""))
        j_content = clean_scalar(row.get("Justification Content", ""))
        j_created_by = clean_scalar(row.get("Justification Created By", ""))
        j_created_date = clean_scalar(row.get("Justification Created Date", ""))
        mapping_method = clean_scalar(row.get("Justification Mapping Method", ""))
        mapping_confidence = clean_scalar(row.get("Justification Confidence", ""))

        def add_event(event: str, date_value: Any, status_after: str, actors: str, source: str, evidence_id: str, detail: str) -> None:
            nonlocal sequence
            date_text = first_date_from_cell(date_value)
            if not date_text and event not in {"Justification Created"}:
                return
            rows.append({
                "Timeline Sequence": sequence,
                "Document Number": doc_number,
                "Version": version,
                "Full Document Version": full_doc,
                "Timeline Event": event,
                "Event Date": date_text,
                "Status After Event": status_after,
                "Event Actor(s)": actors,
                "Evidence Source": source,
                "Evidence Source Row / ID": evidence_id,
                "Justification ID": j_id,
                "Justification Content": j_content,
                "Justification Created By": j_created_by,
                "Justification Created Date": j_created_date,
                "Mapping Method": mapping_method,
                "Mapping Confidence": mapping_confidence,
                "Source Version Key(s)": source_keys,
                "Detail": detail,
            })
            sequence += 1

        if j_id or j_content:
            add_event(
                "Justification Created",
                j_created_date,
                "Revision Justified",
                j_created_by,
                "SPARTADMS__Review_Justification__c.csv",
                clean_scalar(row.get("Justification Source Salesforce ID", "")) or j_id,
                clean_scalar(row.get("Justification Mapping Detail", "")) or "J-record mapped to this document version.",
            )
        add_event("Version Created", row.get("Created Date", ""), "Created", row.get("Created By", ""), "SPARTADMS__Document_Version__c.csv", row.get("Raw Document Version ID(s)", ""), "Source document-version row created.")
        add_event("Reviewed", row.get("Reviewed Date", ""), "Reviewed", row.get("Reviewed By", ""), "Reviewer / workflow evidence", "", row.get("Review Details", ""))
        add_event("Approved", row.get("Approved Date", ""), "Approved", row.get("Approved By", ""), "Approver / workflow evidence", "", row.get("Approval Details", ""))
        add_event("Made Effective", row.get("Effective Date", ""), "Effective", row.get("Approved By", ""), "SPARTADMS__Document_Version__c.csv / Audit History", row.get("Raw Document Version ID(s)", ""), "Version became effective.")
        if clean_scalar(row.get("Obsolete Date", "")):
            basis = clean_scalar(row.get("Obsolete / Superseded Date Basis", ""))
            actor = row.get("Obsoleted By", "") if basis == "Direct audit history event" else "Derived from next effective version"
            source = "CMPL123__Audit_History__c.csv" if basis == "Direct audit history event" else "Document Version Timeline derivation"
            detail = row.get("Obsolete Details", "") if basis == "Direct audit history event" else f"Version was later superseded by {row.get('Superseded By', '')}."
            add_event("Effective to Obsolete" if basis == "Direct audit history event" else "Superseded by Next Version", row.get("Obsolete Date", ""), "Obsolete", actor, source, row.get("Superseded By", ""), detail)
        elif clean_scalar(row.get("Superseded By", "")):
            add_event("Superseded by Next Version", "", "Obsolete", "Derived from next effective version", "Document Version Timeline derivation", row.get("Superseded By", ""), f"Version was later superseded by {row.get('Superseded By', '')}; supersession date was not available.")
        if clean_scalar(row.get("Canceled Date", "")):
            add_event("Canceled", row.get("Canceled Date", ""), "Canceled", row.get("Canceled By", ""), "Audit History", "", row.get("Canceled Details", ""))
    timeline_df = pd.DataFrame(rows, columns=TIMELINE_COLUMNS)
    if not timeline_df.empty:
        timeline_df["__sort_date"] = timeline_df["Event Date"].map(lambda value: as_timestamp(value) or pd.Timestamp.max)
        timeline_df = timeline_df.sort_values(["Document Number", "Version", "__sort_date", "Timeline Sequence"], kind="mergesort").drop(columns=["__sort_date"]).reset_index(drop=True)
        timeline_df["Timeline Sequence"] = range(1, len(timeline_df) + 1)
    logger.emit(action="BUILD_TIMELINE", result="OK", detail=f"{len(timeline_df):,} timeline rows built.")
    return timeline_df

def make_source_field_map() -> pd.DataFrame:
    rows = [
        ("Document Version Register", "Document Number", "Derived from raw approved document number plus document type prefix; e.g., 1022-000003 + SOP -> SOP-1022."),
        ("Document Version Register", "Version", "SPARTADMS__Document_Revision_At_Effective__c, then approval revision, then document version fields."),
        ("Document Version Register", "Full Document Version", "Document Number + '.' + two-digit Version."),
        ("Document Version Register", "Source Version Key(s)", "Raw TrackWise approved document number plus raw version; preserves source discriminator."),
        ("Document Version Register", "Justification Content", "Exact SPARTADMS__Justification__c content; not cleaned or summarized."),
        ("Justification Mapping Diagnostics", "Mapping Method", "DIRECT_VERSION_FIELD, TEXT_VERSION_PARSE, DATE_NEAREST_CORPORATE_MATCH, or UNMAPPED."),
        ("Status Event Detail", "Obsoletion", "CMPL123__Audit_History__c rows whose activity type resolves to Effective to Obsolete."),
        ("Clean Version PDFs", "Source PDF", "raw_dms_root / raw_number / version_folder / renditions / *.pdf."),
    ]
    return pd.DataFrame(rows, columns=["Sheet", "Output Field", "Source Logic"])


def make_run_summary(inputs: RunInputs, outputs: RunOutputs, stats: Dict[str, Any], encodings: Dict[str, str]) -> pd.DataFrame:
    rows = [
        ("Application", f"{APP_TITLE} v{APP_VERSION}"),
        ("History Register Run ID", outputs.run_id),
        ("PDF Package Run ID", outputs.package_run_id if inputs.generate_pdf_packages else "Not generated"),
        ("Run Timestamp", outputs.run_timestamp),
        ("Run Machine", outputs.run_machine),
        ("Run Account", outputs.run_account),
        ("Run Folder", str(outputs.run_folder)),
        ("Workbook Path", str(outputs.workbook_path)),
        ("Event Log CSV", str(outputs.event_log_csv_path)),
        ("PDF Packaging Selected", str(inputs.generate_pdf_packages)),
        ("PDF Cover Page Selected", str(inputs.include_pdf_cover_page)),
        ("Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ]
    for key, value in stats.items():
        rows.append((key, value))
    for key, value in encodings.items():
        rows.append((f"Encoding: {key}", value))
    for field_name, file_name in EXPECTED_RAW_EXPORT_FILENAMES.items():
        rows.append((f"Expected File: {field_name}", file_name))
    return pd.DataFrame(rows, columns=["Metric", "Value"])



def sanitize_excel_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    output = df.copy()
    object_columns = output.select_dtypes(include=["object"]).columns
    for column in object_columns:
        output[column] = output[column].map(lambda value: excel_cell(value))
    return output

def write_workbook(
    outputs: RunOutputs,
    history_df: pd.DataFrame,
    timeline_df: pd.DataFrame,
    source_rows_df: pd.DataFrame,
    justification_diag_df: pd.DataFrame,
    approver_detail_df: pd.DataFrame,
    reviewer_detail_df: pd.DataFrame,
    status_event_df: pd.DataFrame,
    identity_diag_df: pd.DataFrame,
    source_field_map_df: pd.DataFrame,
    run_summary_df: pd.DataFrame,
    event_rows: List[Dict[str, Any]],
    technical_df: pd.DataFrame,
    include_event_log_sheet: bool,
    include_technical_evidence_sheet: bool,
    logger: StructuredLogger,
) -> None:
    logger.emit("WRITE_WORKBOOK", action="STARTED", result="STARTED", detail="Writing Excel workbook and event log CSV.", phase_percent=0)
    event_df = pd.DataFrame(event_rows, columns=EVENT_COLUMNS)
    event_df.to_csv(outputs.event_log_csv_path, index=False, encoding="utf-8-sig")
    logger.emit(action="WRITE_EVENT_LOG_CSV", result="OK", detail=str(outputs.event_log_csv_path), phase_percent=10)

    sheets: List[Tuple[str, pd.DataFrame]] = [
        ("Document Version Register", history_df),
        ("Document Version Timeline", timeline_df),
        ("Source Version Rows", source_rows_df),
        ("Justification Mapping Diagnostics", justification_diag_df),
        ("Approver Detail", approver_detail_df),
        ("Reviewer Detail", reviewer_detail_df),
        ("Status Event Detail", status_event_df),
        ("Source Identity Diagnostics", identity_diag_df),
        ("Source Field Map", source_field_map_df),
        ("Run Summary", run_summary_df),
    ]
    if include_event_log_sheet:
        sheets.append(("Event Log", event_df))
    if include_technical_evidence_sheet:
        sheets.append(("Technical Evidence", technical_df))

    with pd.ExcelWriter(outputs.workbook_path, engine="openpyxl") as writer:
        total = len(sheets)
        for position, (sheet_name, df) in enumerate(sheets, start=1):
            safe_df = sanitize_excel_df(df) if not df.empty else pd.DataFrame(columns=df.columns if len(df.columns) else ["No Data"])
            safe_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            logger.progress.update_phase_percent(10 + int(position * 80 / total), 100)
            logger.emit(action="WRITE_SHEET", result="OK", detail=f"{sheet_name}: {len(safe_df):,} rows")
    apply_light_workbook_formatting(outputs.workbook_path, include_technical_evidence_sheet)
    logger.emit(action="COMPLETE", result="OK", detail=str(outputs.workbook_path), phase_percent=100)


def apply_light_workbook_formatting(workbook_path: Path, include_technical_evidence_sheet: bool) -> None:
    openpyxl_module = importlib.import_module("openpyxl")
    styles_module = importlib.import_module("openpyxl.styles")
    workbook = openpyxl_module.load_workbook(workbook_path)
    header_fill = styles_module.PatternFill("solid", fgColor="1F4E78")
    header_font = styles_module.Font(color="FFFFFF", bold=True)
    wrap_alignment = styles_module.Alignment(wrap_text=True, vertical="top")
    for worksheet in workbook.worksheets:
        if worksheet.max_row >= 1:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = wrap_alignment
        for column_letter, width in FIXED_WIDTHS.get(worksheet.title, {}).items():
            worksheet.column_dimensions[column_letter].width = width
        if worksheet.title in {"Document Version Register", "Document Version Timeline", "Justification Mapping Diagnostics"}:
            for row in worksheet.iter_rows(min_row=2, max_row=min(worksheet.max_row, 5000)):
                for cell in row:
                    cell.alignment = wrap_alignment
        if worksheet.title == "Technical Evidence" and include_technical_evidence_sheet:
            worksheet.sheet_state = "hidden"
    workbook.save(workbook_path)




@dataclass
class PdfCandidate:
    path: Path
    sha256: str
    raw_document_number: str
    version_folder: str
    source_version_key: str
    is_rendition: bool
    is_lastuploaded: bool
    filename_contains_full_version: bool
    filename_contains_display_version: bool
    score: int
    score_reasons: List[str]
    selected: bool = False
    rank: int = 0


@dataclass
class PdfSelectionDecision:
    full_document_version: str
    selected_path: Path
    selected_sha256: str
    selected_score: int
    selected_reasons: List[str]
    candidate_count: int
    distinct_sha_count: int
    conflict_status: str
    manual_review_recommended: bool
    tie_on_top_score: bool
    selection_basis: str
    candidates: List[PdfCandidate]


def split_multi_value_cell(value: Any) -> List[str]:
    text = clean_scalar(value)
    if not text:
        return []
    return [part.strip() for part in re.split(r";|\n", text) if part.strip()]


def path_parts_lower(path: Path) -> List[str]:
    return [part.lower() for part in path.parts]


def candidate_raw_document_number(path: Path) -> str:
    parts = list(path.parts)
    lowered = [part.lower() for part in parts]
    if "renditions" in lowered:
        idx = lowered.index("renditions")
        if idx >= 2:
            return parts[idx - 2]
    if "lastuploaded" in lowered:
        idx = lowered.index("lastuploaded")
        if idx >= 1:
            return parts[idx - 1]
    for part in reversed(parts):
        if re.match(r"^[A-Za-z]+-?\d|^\d{3,6}-\d+", part):
            return part
    return ""


def candidate_version_folder(path: Path) -> str:
    parts = list(path.parts)
    lowered = [part.lower() for part in parts]
    if "renditions" in lowered:
        idx = lowered.index("renditions")
        if idx >= 1:
            return parts[idx - 1]
    return ""


def normalize_for_filename_match(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "", clean_scalar(value).upper())


def score_pdf_candidate(path: Path, sha: str, row: pd.Series) -> PdfCandidate:
    full_doc = clean_scalar(row.get("Full Document Version", "")).strip()
    display_doc = clean_scalar(row.get("Document Number", "")).strip()
    version = clean_scalar(row.get("Version", "")).strip()
    raw_numbers = [item.upper() for item in split_multi_value_cell(row.get("Raw TrackWise Document Number(s)", ""))]
    source_keys = [item.upper() for item in split_multi_value_cell(row.get("Source Version Key(s)", ""))]
    source_sequences = split_multi_value_cell(row.get("Source Document Sequence Number(s)", ""))

    path_text = str(path)
    path_text_upper = path_text.upper()
    file_stem_upper = path.stem.upper()
    parts_lower = path_parts_lower(path)
    raw_doc = candidate_raw_document_number(path).upper()
    version_folder = candidate_version_folder(path)
    source_version_key = f"{raw_doc}.{normalize_version(version_folder)}" if raw_doc and version_folder else ""
    is_rendition = "renditions" in parts_lower
    is_lastuploaded = "lastuploaded" in parts_lower
    filename_contains_full_version = normalize_for_filename_match(full_doc) in normalize_for_filename_match(path.name)
    filename_contains_display_version = normalize_for_filename_match(display_doc) in normalize_for_filename_match(path.name)

    score = 0
    reasons: List[str] = []

    # Direct ContentVersion/ContentDocument scoring can be added when the relevant CSV linkage
    # tables are available to this utility. Current evidence is therefore folder/name/source-key based.
    if raw_doc and raw_doc in raw_numbers:
        score += 500
        reasons.append("raw document number match")
    if source_version_key and source_version_key in source_keys:
        score += 400
        reasons.append("source version key match")
    expected_version_int = normalize_version_int(version)
    folder_version_int = normalize_version_int(version_folder)
    if expected_version_int is not None and folder_version_int == expected_version_int:
        score += 250
        reasons.append("version folder match")
    if filename_contains_full_version:
        score += 200
        reasons.append("filename contains full document version")
    if filename_contains_display_version:
        score += 150
        reasons.append("filename contains document number")
    if is_rendition:
        score += 100
        reasons.append("renditions folder")
    if is_lastuploaded:
        score += 50
        reasons.append("lastuploaded folder")
    if raw_doc and len(raw_numbers) == 1 and raw_doc == raw_numbers[0]:
        score += 75
        reasons.append("single raw source lineage")
    if source_sequences and version_folder and version_folder in source_sequences:
        score += 50
        reasons.append("source sequence/version folder agreement")

    if not reasons:
        reasons.append("deterministic fallback only")

    return PdfCandidate(
        path=path,
        sha256=sha,
        raw_document_number=raw_doc,
        version_folder=version_folder,
        source_version_key=source_version_key,
        is_rendition=is_rendition,
        is_lastuploaded=is_lastuploaded,
        filename_contains_full_version=filename_contains_full_version,
        filename_contains_display_version=filename_contains_display_version,
        score=score,
        score_reasons=reasons,
    )


def select_pdf_source_candidate(candidate_paths: Sequence[Path], row: pd.Series) -> PdfSelectionDecision:
    full_doc = clean_scalar(row.get("Full Document Version", ""))
    sha_by_path = {path: sha256_file(path) for path in candidate_paths}
    candidates = [score_pdf_candidate(path, sha, row) for path, sha in sha_by_path.items()]
    candidates.sort(key=lambda candidate: (-candidate.score, str(candidate.path).lower()))
    for idx, candidate in enumerate(candidates, start=1):
        candidate.rank = idx
    selected = candidates[0]
    selected.selected = True
    distinct_sha_count = len({candidate.sha256 for candidate in candidates})
    top_score = selected.score
    top_candidates = [candidate for candidate in candidates if candidate.score == top_score]
    tie_on_top_score = len(top_candidates) > 1

    if len(candidates) == 1:
        conflict_status = "Single source PDF"
        manual_review = False
        basis = "Single candidate source PDF selected."
    elif distinct_sha_count == 1:
        conflict_status = "Multiple identical candidate PDFs"
        manual_review = False
        basis = "Multiple candidate PDFs had identical SHA-256; highest-scoring candidate selected."
    elif tie_on_top_score:
        conflict_status = "Top candidate source PDFs tied"
        manual_review = True
        basis = f"{len(candidates)} non-identical candidate PDFs found; top candidates tied at score={top_score}. Deterministic path selected for package continuity."
    elif selected.score < 350:
        conflict_status = "Low-confidence source selection"
        manual_review = True
        basis = f"{len(candidates)} non-identical candidate PDFs found; selected candidate did not exceed source-evidence threshold."
    else:
        conflict_status = "Multiple non-identical candidate PDFs - resolved by scoring"
        manual_review = False
        basis = f"{len(candidates)} non-identical candidate PDFs found; selected highest-scoring candidate."

    return PdfSelectionDecision(
        full_document_version=full_doc,
        selected_path=selected.path,
        selected_sha256=selected.sha256,
        selected_score=selected.score,
        selected_reasons=selected.score_reasons,
        candidate_count=len(candidates),
        distinct_sha_count=distinct_sha_count,
        conflict_status=conflict_status,
        manual_review_recommended=manual_review,
        tie_on_top_score=tie_on_top_score,
        selection_basis=basis,
        candidates=candidates,
    )

def locate_source_pdfs(raw_dms_root: Path, raw_number: str, version: str) -> List[Path]:
    version_int = normalize_version_int(version)
    if version_int is None:
        return []
    candidate_dir = raw_dms_root / raw_number / str(version_int) / "renditions"
    if not candidate_dir.exists():
        return []
    return sorted(path for path in candidate_dir.glob("*.pdf") if path.is_file())


def make_cover_pdf(row: pd.Series, output_path: Path, outputs: RunOutputs) -> None:
    """Generate a QA-first multi-page migration cover section.

    Requirements implemented here:
    - Letter page, 1-inch left/right margins.
    - Header on every generated cover page with Run ID and Page X of Y.
    - Blank cover fields render as N/A.
    - Full justification content is never truncated in the PDF cover section.
    - QA reviewer lifecycle/justification fields appear before package/run/integrity metadata.
    """
    try:
        colors = importlib.import_module("reportlab.lib.colors")
        pagesizes = importlib.import_module("reportlab.lib.pagesizes")
        styles_module = importlib.import_module("reportlab.lib.styles")
        units = importlib.import_module("reportlab.lib.units")
        platypus = importlib.import_module("reportlab.platypus")
        canvas_module = importlib.import_module("reportlab.pdfgen.canvas")
    except ImportError as exc:
        raise RuntimeError(PDF_DEPENDENCY_MESSAGE) from exc

    letter = getattr(pagesizes, "letter")
    inch = getattr(units, "inch")
    SimpleDocTemplate = getattr(platypus, "SimpleDocTemplate")
    Paragraph = getattr(platypus, "Paragraph")
    Spacer = getattr(platypus, "Spacer")
    Table = getattr(platypus, "Table")
    TableStyle = getattr(platypus, "TableStyle")
    getSampleStyleSheet = getattr(styles_module, "getSampleStyleSheet")
    ParagraphStyle = getattr(styles_module, "ParagraphStyle")
    BaseCanvas = getattr(canvas_module, "Canvas")

    styles = getSampleStyleSheet()
    normal = ParagraphStyle("CoverNormal", parent=styles["Normal"], fontName="Helvetica", fontSize=8.5, leading=10.5)
    small = ParagraphStyle("CoverSmall", parent=styles["Normal"], fontName="Helvetica", fontSize=7.5, leading=9)
    heading = ParagraphStyle("CoverHeading", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=11, leading=13, spaceBefore=8, spaceAfter=4)
    title = ParagraphStyle("CoverTitle", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=15, leading=18, spaceAfter=6)
    subtitle = ParagraphStyle("CoverSubtitle", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, leading=11, spaceAfter=6)
    justification = ParagraphStyle("Justification", parent=styles["Normal"], fontName="Helvetica", fontSize=8.2, leading=10.2, leftIndent=6, rightIndent=6, spaceBefore=3, spaceAfter=6)

    def cover_value(value: Any) -> str:
        text = clean_scalar(value).strip()
        return text if text else "N/A"

    def esc(value: Any, force_na: bool = True) -> str:
        text = cover_value(value) if force_na else clean_scalar(value)
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")

    def para(value: Any, style: Any = normal, force_na: bool = True) -> Any:
        return Paragraph(esc(value, force_na=force_na), style)

    def section_table(rows: Sequence[Tuple[str, Any]]) -> Any:
        table_rows = [[para(label, small, force_na=False), para(value, normal)] for label, value in rows]
        tbl = Table(table_rows, colWidths=[1.85 * inch, 4.65 * inch], repeatRows=0)
        tbl.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BFBFBF")),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F2F2F2")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        return tbl

    def justification_content_box(value: Any) -> Any:
        tbl = Table([[para(value, justification)]], colWidths=[6.5 * inch])
        tbl.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#808080")),
            ("BACKGROUND", (0, 0), (-1, -1), colors.white),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        return tbl

    def lifecycle_summary_text() -> str:
        explicit_summary = clean_scalar(row.get("Lifecycle Summary", "")).strip()
        if explicit_summary:
            return explicit_summary
        superseded_by = clean_scalar(row.get("Superseded By", "")).strip()
        superseded_date = clean_scalar(row.get("Obsolete Date", "")).strip()
        status = clean_scalar(row.get("Lifecycle Status Normalized", "")).upper()
        if superseded_by and superseded_date:
            return f"This document version was approved, became effective, and was superseded by {superseded_by} on {superseded_date}."
        if superseded_by:
            return f"This document version was approved and became effective. The export indicates it was superseded by {superseded_by}; the supersession date was not available in the source fields used for this cover."
        if status == "EFFECTIVE":
            return "This document version is the current effective version in the reconstructed export."
        if status == "CANCELED":
            return "This document version was canceled; see lifecycle fields below."
        return "This document version is presented with the final lifecycle status recorded in the source export."

    source_version_keys = cover_value(row.get("Source Version Key(s)", ""))
    full_doc = cover_value(row.get("Full Document Version", ""))
    lifecycle_summary = lifecycle_summary_text()

    story: List[Any] = []
    story.append(para("NEUROTIC DMS Migration Cover Section", title, force_na=False))
    story.append(para("Document Version Evidence Package", subtitle, force_na=False))

    # QA/business reviewer sections come first.
    story.append(para("1. Document Version Identity", heading, force_na=False))
    story.append(section_table([
        ("Document Version", full_doc),
        ("Document Title", row.get("Document Title", "")),
        ("Document Type", row.get("Document Type", "")),
        ("Department", row.get("Department", "")),
        ("Current / Final Version Status", row.get("Lifecycle Status Normalized", "")),
        ("Lifecycle Summary", lifecycle_summary),
    ]))

    story.append(para("2. Document Lifecycle", heading, force_na=False))
    story.append(section_table([
        ("Created Date", row.get("Created Date", "")),
        ("Approved Date", row.get("Approved Date", "")),
        ("Effective Date", row.get("Effective Date", "")),
        ("Obsolete / Superseded Date", row.get("Obsolete Date", "")),
        ("Obsolete / Superseded Date Basis", row.get("Obsolete / Superseded Date Basis", "")),
        ("Canceled Date", row.get("Canceled Date", "")),
        ("Canceled Date Basis", row.get("Canceled Date Basis", "")),
        ("Superseded By", row.get("Superseded By", "")),
        ("Final Version Status", row.get("Version Status", "")),
    ]))

    story.append(para("3. People and Workflow", heading, force_na=False))
    people_rows: List[Tuple[str, Any]] = [
        ("Justification Created By", row.get("Justification Created By", "")),
        ("Justification Owner", row.get("Justification Owner", "")),
        ("Reviewed By", row.get("Reviewed By", "")),
        ("Approved By", row.get("Approved By", "")),
    ]
    if clean_scalar(row.get("Obsoleted By", "")):
        people_rows.append(("Direct Obsoletion Actor", row.get("Obsoleted By", "")))
    story.append(section_table(people_rows))

    story.append(para("4. Revision Justification", heading, force_na=False))
    story.append(section_table([
        ("Justification ID", row.get("Justification ID", "")),
        ("Justification", row.get("Justification Content", "")),
        ("Justification Created Date", row.get("Justification Created Date", "")),
        ("Justification Created By", row.get("Justification Created By", "")),
        ("Justification Owner", row.get("Justification Owner", "")),
        ("Justification Last Modified Date", row.get("Justification Last Modified Date", "")),
        ("Justification Last Modified By", row.get("Justification Last Modified By", "")),
    ]))

    # Source lineage is still reviewer-relevant, but after the lifecycle and justification.
    story.append(para("5. Source Reconstruction Summary", heading, force_na=False))
    story.append(section_table([
        ("Source Version Key(s)", source_version_keys),
        ("Raw TrackWise Document Number(s)", row.get("Raw TrackWise Document Number(s)", "")),
        ("Source Document Sequence Number(s)", row.get("Source Document Sequence Number(s)", "")),
        ("Selected Source PDF", row.get("Selected Source PDF Path", "") or row.get("Source PDF Path(s)", "")),
        ("Selected Source Score", row.get("Selected Source Score", "")),
        ("Selected Source Evidence", row.get("Selected Source Selection Reason", "")),
        ("Selected Source Rule", row.get("Source Selection Rule", "")),
        ("Candidate Source PDF Count", row.get("Candidate Source PDF Count", "")),
        ("Distinct Candidate SHA-256 Count", row.get("Distinct Candidate SHA-256 Count", "")),
        ("Source Conflict Status", row.get("Source Conflict Status", "")),
        ("Manual Review Recommended", row.get("Manual Review Recommended", "")),
        ("Candidate Details Location", "See PDF Package Manifest and Document Version Register."),
        ("Reconstruction Note", "This package was generated from raw TrackWise DMS export evidence. The original source PDF was not modified."),
    ]))

    # Technical/data-integrity sections are retained, but later.
    story.append(para("6. Package Integrity and Run Details", heading, force_na=False))
    story.append(section_table([
        ("Run ID", outputs.package_run_id),
        ("History Register Run ID", outputs.run_id),
        ("Run Timestamp", outputs.run_timestamp),
        ("Run Machine", outputs.run_machine),
        ("Run Account", outputs.run_account),
        ("Application", f"{APP_TITLE} v{APP_VERSION}"),
        ("Output PDF File", f"{safe_pdf_name(full_doc)}.pdf"),
        ("Output PDF Path", row.get("Output PDF Path", "")),
        ("Source PDF SHA-256", row.get("Source PDF SHA-256", "")),
        ("Final Package SHA-256", "Recorded in package manifest after finalization"),
        ("PDF Packaging Result", "Recorded in package manifest after finalization"),
        ("Package Manifest", str(outputs.package_log_folder / "pdf_package_manifest.csv")),
        ("Event Log", str(outputs.event_log_csv_path)),
    ]))

    story.append(para("7. Reader Note", heading, force_na=False))
    story.append(para("This generated cover section summarizes the TrackWise export metadata associated with this document version. The controlled document content that follows is the original TrackWise PDF rendition.", normal, force_na=False))

    class NumberedCoverCanvas(BaseCanvas):
        def __init__(self, *args: Any, **kwargs: Any) -> None:
            super().__init__(*args, **kwargs)
            self._saved_page_states: List[Dict[str, Any]] = []

        def showPage(self) -> None:  # type: ignore[override]
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self) -> None:  # type: ignore[override]
            page_count = len(self._saved_page_states)
            for page_number, state in enumerate(self._saved_page_states, start=1):
                self.__dict__.update(state)
                self._draw_header_footer(page_number, page_count)
                super().showPage()
            super().save()

        def _draw_header_footer(self, page_number: int, page_count: int) -> None:
            width, height = letter
            self.saveState()
            header = f"NEUROTIC DMS Migration Evidence Package | {full_doc} | Run ID: {outputs.package_run_id} | Page {page_number} of {page_count}"
            footer = "Generated migration cover section only. This section is not part of the original controlled document."
            self.setFont("Helvetica", 7.3)
            self.drawString(inch, height - 0.55 * inch, header[:150])
            self.line(inch, height - 0.62 * inch, width - inch, height - 0.62 * inch)
            self.line(inch, 0.62 * inch, width - inch, 0.62 * inch)
            self.drawString(inch, 0.45 * inch, footer[:145])
            self.restoreState()

    doc = SimpleDocTemplate(str(output_path), pagesize=letter, leftMargin=inch, rightMargin=inch, topMargin=0.85 * inch, bottomMargin=0.8 * inch)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.build(story, canvasmaker=NumberedCoverCanvas)

def merge_pdfs(cover_pdf: Path, source_pdf: Path, output_pdf: Path) -> None:
    try:
        pdf_module = importlib.import_module("pypdf")
    except ImportError as first_exc:
        try:
            pdf_module = importlib.import_module("PyPDF2")
        except ImportError as second_exc:
            raise RuntimeError(PDF_DEPENDENCY_MESSAGE) from second_exc
    reader_class = getattr(pdf_module, "PdfReader")
    writer_class = getattr(pdf_module, "PdfWriter")
    writer = writer_class()
    for input_pdf in [cover_pdf, source_pdf]:
        reader = reader_class(str(input_pdf))
        for page in reader.pages:
            writer.add_page(page)
    with output_pdf.open("wb") as handle:
        writer.write(handle)


def preflight_pdf_packaging(inputs: RunInputs, outputs: RunOutputs, logger: StructuredLogger) -> None:
    """Fail fast before packaging thousands of files if PDF-cover dependencies or styling are broken."""
    if not inputs.generate_pdf_packages or not inputs.include_pdf_cover_page:
        return
    try:
        # ReportLab cover generation depends on valid color/style syntax. This catches
        # errors like passing BFBFBF without a leading # before the real package loop.
        with tempfile.TemporaryDirectory() as tmpdir:
            test_cover = Path(tmpdir) / "cover_preflight.pdf"
            dummy = pd.Series({
                "Full Document Version": "PREFLIGHT-0001.01",
                "Document Title": "PDF cover preflight test",
                "Document Type": "PRE",
                "Department": "System",
                "Lifecycle Status Normalized": "PREFLIGHT",
                "Superseded By": "",
                "Justification Created Date": outputs.run_timestamp,
                "Approved Date": "",
                "Effective Date": "",
                "Obsolete Date": "",
                "Justification Created By": outputs.run_account,
                "Justification Owner": outputs.run_account,
                "Justification Last Modified By": outputs.run_account,
                "Justification Last Modified Date": outputs.run_timestamp,
                "Reviewed By": "",
                "Approved By": "",
                "Obsoleted By": "",
                "Justification ID": "PREFLIGHT",
                "Justification Mapping Method": "PREFLIGHT",
                "Justification Confidence": "High",
                "Justification Mapping Detail": "Preflight cover-generation test before package loop.",
                "Justification Content": "PDF cover preflight test. This text verifies wrapping and table styling without truncation.",
                "Source Version Key(s)": "PREFLIGHT-0001.01",
                "Raw TrackWise Document Number(s)": "PREFLIGHT-0001",
                "Source Document Sequence Number(s)": "1",
                "Raw Corporate Document ID(s)": "PREFLIGHT",
                "Source PDF Path(s)": "PREFLIGHT",
                "Output PDF Path": "PREFLIGHT",
                "Source PDF SHA-256": "PREFLIGHT",
                "Output PDF SHA-256": "",
                "PDF Packaging Result": "PREFLIGHT",
            })
            make_cover_pdf(dummy, test_cover, outputs)
        missing_pdf_deps = get_missing_pdf_dependencies(include_cover_page=True)
        if missing_pdf_deps:
            raise RuntimeError(
                f"Missing PDF cover dependency/dependencies: {', '.join(missing_pdf_deps)}. {PDF_DEPENDENCY_MESSAGE}"
            )
        logger.emit("PDF_PACKAGING", action="PREFLIGHT", result="OK", detail="PDF cover and merge dependencies passed preflight.", phase_percent=1)
    except Exception as exc:
        raise RuntimeError(f"PDF packaging preflight failed before package loop: {exc}") from exc


def safe_pdf_name(text: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9._-]+", "_", clean_scalar(text).strip())
    clean = re.sub(r"_+", "_", clean).strip("_")
    return clean or "document_version"


def generate_pdf_packages(
    inputs: RunInputs,
    outputs: RunOutputs,
    history_df: pd.DataFrame,
    logger: StructuredLogger,
    cancel: CancelToken,
) -> pd.DataFrame:
    updated_history = history_df.copy()
    updated_history = ensure_object_columns(updated_history, SOURCE_SELECTION_TEXT_COLUMNS)
    if not inputs.generate_pdf_packages:
        return updated_history
    logger.emit("PDF_PACKAGING", action="STARTED", result="STARTED", detail=f"{len(history_df):,} document-version packages to evaluate.", phase_percent=0)
    if inputs.raw_dms_file_tree_folder is None:
        raise RuntimeError("Raw DMS File Tree Folder is required when PDF packaging is enabled.")
    preflight_pdf_packaging(inputs, outputs, logger)
    manifest_rows: List[Dict[str, Any]] = []
    exception_rows: List[Dict[str, Any]] = []
    total = max(len(history_df), 1)

    for position, row_tuple in enumerate(history_df.itertuples(index=True), start=1):
        cancel.check()
        row_index = row_tuple.Index
        row = history_df.loc[row_index].copy()
        full_doc = clean_scalar(row.get("Full Document Version", ""))
        raw_numbers = [item.strip() for item in clean_scalar(row.get("Raw TrackWise Document Number(s)", "")).split(";") if item.strip()]
        version = clean_scalar(row.get("Version", ""))
        candidate_paths: List[Path] = []
        for raw_number in raw_numbers:
            candidate_paths.extend(locate_source_pdfs(inputs.raw_dms_file_tree_folder, raw_number, version))
        candidate_paths = list(dict.fromkeys(candidate_paths))
        logger.progress.update_phase_percent(position, total)

        if not candidate_paths:
            detail = "No source PDF found under raw DMS file tree."
            exception_rows.append({"Run ID": outputs.package_run_id, "Full Document Version": full_doc, "Result": "MISSING_SOURCE_PDF", "Detail": detail})
            updated_history.at[row_index, "PDF Packaging Result"] = "MISSING_SOURCE_PDF"
            updated_history.at[row_index, "PDF Packaging Detail"] = detail
            updated_history.at[row_index, "Manual Review Recommended"] = "Yes"
            if position <= 10 or position % 100 == 0:
                logger.emit(action="PDF_PACKAGE", result="MISSING_SOURCE_PDF", level="WARN", full_document_version=full_doc, detail=detail)
            continue

        decision = select_pdf_source_candidate(candidate_paths, row)
        selected_source_pdf = decision.selected_path
        selected_source_sha = decision.selected_sha256
        sorted_candidates = [candidate.path for candidate in decision.candidates]
        candidate_sha_values = join_unique(candidate.sha256 for candidate in decision.candidates)
        candidate_scores = join_unique(f"{candidate.rank}:{candidate.score}" for candidate in decision.candidates)
        source_selection_rule = decision.selection_basis
        selected_reason = "; ".join(decision.selected_reasons)
        source_conflict_status = decision.conflict_status
        manual_review = "Yes" if decision.manual_review_recommended else "No"

        if decision.candidate_count == 1:
            logger.emit(
                action="PDF_SOURCE_SELECTION",
                result="OK",
                full_document_version=full_doc,
                detail="selected source PDF by single candidate match",
            )
        elif decision.distinct_sha_count == 1:
            logger.emit(
                action="PDF_SOURCE_SELECTION",
                result="OK_DUPLICATE_SOURCE",
                full_document_version=full_doc,
                detail=f"{decision.candidate_count} candidate PDFs found with identical SHA-256; selected highest-scoring candidate. score={decision.selected_score}; reasons={selected_reason}. Duplicates recorded in manifest/register.",
            )
        elif decision.manual_review_recommended:
            logger.emit(
                action="PDF_SOURCE_SELECTION",
                result="REVIEW_REQUIRED",
                level="WARN",
                full_document_version=full_doc,
                detail=f"{decision.candidate_count} non-identical candidate PDFs found; selected score={decision.selected_score}; reasons={selected_reason}. {decision.selection_basis} Manual review recommended. All candidates recorded in manifest/register.",
            )
            exception_rows.append({
                "Run ID": outputs.package_run_id,
                "Full Document Version": full_doc,
                "Result": "REVIEW_REQUIRED_SOURCE_SELECTION",
                "Detail": decision.selection_basis,
                "Selected Source PDF Path": str(selected_source_pdf),
                "Selected Source Score": decision.selected_score,
                "Selected Source Selection Reason": selected_reason,
                "Candidate Source PDF Paths": join_unique(str(path) for path in sorted_candidates),
                "Candidate Source PDF SHA-256 Values": candidate_sha_values,
                "Candidate Source Scores": candidate_scores,
            })
        else:
            logger.emit(
                action="PDF_SOURCE_SELECTION",
                result="REVIEW_SOURCE_CONFLICT_RESOLVED",
                full_document_version=full_doc,
                detail=f"{decision.candidate_count} non-identical candidate PDFs found; selected highest-scoring source. score={decision.selected_score}; reasons={selected_reason}. All candidates recorded in manifest/register.",
            )

        output_stem = safe_pdf_name(full_doc)
        output_paths: List[str] = []
        source_paths: List[str] = []
        source_hashes: List[str] = []
        output_hashes: List[str] = []
        package_results: List[str] = []
        package_details: List[str] = []
        cover_pdf = Path(tempfile.gettempdir()) / f"{output_stem}_{outputs.package_run_id}_cover.pdf"
        output_pdf = outputs.pdf_package_folder / f"{output_stem}.pdf"
        detail = ""
        cover_generated = "No"
        merge_status = "Not applicable"
        package_sha = ""
        try:
            if inputs.include_pdf_cover_page:
                cover_row = row.copy()
                cover_row["Selected Source PDF Path"] = str(selected_source_pdf)
                cover_row["Selected Source PDF SHA-256"] = selected_source_sha
                cover_row["Selected Source Score"] = str(decision.selected_score)
                cover_row["Selected Source Selection Reason"] = selected_reason
                cover_row["Source PDF Path(s)"] = str(selected_source_pdf)
                cover_row["Candidate Source PDF Count"] = str(decision.candidate_count)
                cover_row["Distinct Candidate SHA-256 Count"] = str(decision.distinct_sha_count)
                cover_row["Candidate Source PDF Paths"] = join_unique(str(path) for path in sorted_candidates)
                cover_row["Candidate Source PDF SHA-256 Values"] = candidate_sha_values
                cover_row["Candidate Source Scores"] = candidate_scores
                cover_row["Source Selection Rule"] = source_selection_rule
                cover_row["Source Selection Basis"] = decision.selection_basis
                cover_row["Source Conflict Status"] = source_conflict_status
                cover_row["Manual Review Recommended"] = manual_review
                cover_row["Output PDF Path"] = str(output_pdf)
                cover_row["Source PDF SHA-256"] = selected_source_sha
                cover_row["PDF Packaging Result"] = "Recorded in package manifest after finalization"
                make_cover_pdf(cover_row, cover_pdf, outputs)
                merge_pdfs(cover_pdf, selected_source_pdf, output_pdf)
                detail = "Migration cover section merged with original TrackWise PDF rendition."
                cover_generated = "Yes"
                merge_status = "OK"
            else:
                output_pdf.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(selected_source_pdf, output_pdf)
                detail = "Clean-named copy of original TrackWise PDF rendition; no migration cover section included."
                cover_generated = "No"
                merge_status = "Not applicable"
            package_sha = sha256_file(output_pdf)
            source_paths.append(str(selected_source_pdf))
            output_paths.append(str(output_pdf))
            source_hashes.append(selected_source_sha)
            output_hashes.append(package_sha)
            package_results.append("OK")
            package_details.append(detail)
            if position <= 10 or position % 100 == 0 or position == total:
                logger.emit(action="PDF_PACKAGE", result="OK", full_document_version=full_doc, detail=str(output_pdf))
        except Exception as exc:
            package_results.append("FAIL")
            package_details.append(str(exc))
            exception_rows.append({"Run ID": outputs.package_run_id, "Full Document Version": full_doc, "Source PDF Path": str(selected_source_pdf), "Result": "FAIL", "Detail": str(exc)})
            logger.emit(action="PDF_PACKAGE", result="FAIL", level="ERROR", full_document_version=full_doc, detail=str(exc))
        finally:
            try:
                if cover_pdf.exists():
                    cover_pdf.unlink()
            except OSError:
                pass

        for candidate in decision.candidates:
            manifest_rows.append({
                "Run ID": outputs.package_run_id,
                "Document Number": clean_scalar(row.get("Document Number", "")),
                "Version": version,
                "Full Document Version": full_doc,
                "Output PDF Filename": output_pdf.name,
                "Output Package Path": str(output_pdf),
                "Cover Page Generated": cover_generated,
                "PDF Merge Status": merge_status,
                "Packaged PDF SHA-256": package_sha,
                "Result": join_unique(package_results) or "FAIL",
                "Candidate Rank": candidate.rank,
                "Candidate Selected": "Yes" if candidate.selected else "No",
                "Candidate Path": str(candidate.path),
                "Candidate SHA-256": candidate.sha256,
                "Candidate Score": candidate.score,
                "Candidate Score Reasons": "; ".join(candidate.score_reasons),
                "Candidate Raw Document Number": candidate.raw_document_number,
                "Candidate Version Folder": candidate.version_folder,
                "Candidate Is Rendition": str(candidate.is_rendition),
                "Candidate Is LastUploaded": str(candidate.is_lastuploaded),
                "Candidate Filename Contains Full Version": str(candidate.filename_contains_full_version),
                "Candidate Filename Contains Display Version": str(candidate.filename_contains_display_version),
                "Candidate Source Version Key": candidate.source_version_key,
                "Candidate Source Version Key Match": str(candidate.source_version_key.upper() in [key.upper() for key in split_multi_value_cell(row.get("Source Version Key(s)", ""))]),
                "Candidate Raw Document Number Match": str(candidate.raw_document_number.upper() in [raw.upper() for raw in raw_numbers]),
                "Selected Source PDF Path": str(selected_source_pdf),
                "Selected Source SHA-256": selected_source_sha,
                "Selected Source Score": decision.selected_score,
                "Selected Source Selection Reason": selected_reason,
                "Source Conflict Status": source_conflict_status,
                "Manual Review Recommended": manual_review,
                "Source Selection Rule": source_selection_rule,
                "Source Selection Basis": decision.selection_basis,
                "Candidate Source PDF Count": decision.candidate_count,
                "Distinct Candidate SHA-256 Count": decision.distinct_sha_count,
                "Detail": detail,
            })

        updated_history.at[row_index, "Selected Source PDF Path"] = str(selected_source_pdf)
        updated_history.at[row_index, "Selected Source PDF SHA-256"] = selected_source_sha
        updated_history.at[row_index, "Selected Source Score"] = str(decision.selected_score)
        updated_history.at[row_index, "Selected Source Selection Reason"] = selected_reason
        updated_history.at[row_index, "Candidate Source PDF Count"] = str(decision.candidate_count)
        updated_history.at[row_index, "Distinct Candidate SHA-256 Count"] = str(decision.distinct_sha_count)
        updated_history.at[row_index, "Candidate Source PDF Paths"] = join_unique(str(path) for path in sorted_candidates)
        updated_history.at[row_index, "Candidate Source PDF SHA-256 Values"] = candidate_sha_values
        updated_history.at[row_index, "Candidate Source Scores"] = candidate_scores
        updated_history.at[row_index, "Source Conflict Status"] = source_conflict_status
        updated_history.at[row_index, "Manual Review Recommended"] = manual_review
        updated_history.at[row_index, "Source Selection Basis"] = decision.selection_basis
        updated_history.at[row_index, "Source Selection Rule"] = source_selection_rule
        updated_history.at[row_index, "Source PDF Path(s)"] = join_unique(source_paths) or str(selected_source_pdf)
        updated_history.at[row_index, "Output PDF Path"] = join_unique(output_paths)
        updated_history.at[row_index, "Output PDF Filename"] = join_unique(Path(path).name for path in output_paths)
        updated_history.at[row_index, "Source PDF SHA-256"] = join_unique(source_hashes) or selected_source_sha
        updated_history.at[row_index, "Output PDF SHA-256"] = join_unique(output_hashes)
        updated_history.at[row_index, "PDF Packaging Result"] = join_unique(package_results)
        updated_history.at[row_index, "PDF Packaging Detail"] = excel_cell(join_unique(package_details, separator="\n"), 8000)
        if position % 100 == 0 or position == total:
            logger.emit(action="PDF_PACKAGING_PROGRESS", result="PROGRESS", detail=f"{position:,} / {total:,} document-version rows evaluated")

    manifest_df = pd.DataFrame(manifest_rows)
    exceptions_df = pd.DataFrame(exception_rows)
    manifest_path = outputs.package_log_folder / "pdf_package_manifest.csv"
    exceptions_path = outputs.package_log_folder / "pdf_package_exceptions.csv"
    manifest_df.to_csv(manifest_path, index=False, encoding="utf-8-sig")
    exceptions_df.to_csv(exceptions_path, index=False, encoding="utf-8-sig")
    logger.emit(action="COMPLETE", result="OK", detail=f"PDF package phase complete. Manifest: {manifest_path}", phase_percent=100)
    return updated_history

def run_document_history_generation(inputs: RunInputs, ui_queue: "queue.Queue[Dict[str, Any]]", cancel: CancelToken) -> RunOutputs:
    tracker = ProgressTracker(PHASE_WEIGHTS_WITH_PDF if inputs.generate_pdf_packages else PHASE_WEIGHTS_NO_PDF)
    logger = StructuredLogger(tracker, ui_queue)
    outputs = create_run_outputs(inputs.output_run_folder, inputs.generate_pdf_packages)
    try:
        validate_inputs(inputs, logger)
        tables, encodings = load_input_tables(inputs, logger, cancel)
        indexes = build_indexes(tables, logger, cancel)
        source_rows_df = build_source_version_rows(tables["document_versions"], indexes, logger, cancel)
        justification_diag_df, source_rows_df = map_justifications(tables.get("review_justifications", pd.DataFrame()), source_rows_df, indexes.get("users", {}), logger, cancel)
        logger.emit("MAP_ACTORS_AND_OBSOLETION", action="STARTED", result="STARTED", detail="Mapping reviewer, approver, and obsoletion evidence.", phase_percent=0)
        approver_detail_df, approver_summary = build_actor_detail(tables.get("approvers", pd.DataFrame()), source_rows_df, indexes.get("users", {}), "approver")
        logger.emit(action="APPROVER_MAP", result="OK", detail=f"{len(approver_detail_df):,} approver detail rows.", phase_percent=30)
        reviewer_detail_df, reviewer_summary = build_actor_detail(tables.get("reviewers", pd.DataFrame()), source_rows_df, indexes.get("users", {}), "reviewer")
        logger.emit(action="REVIEWER_MAP", result="OK", detail=f"{len(reviewer_detail_df):,} reviewer detail rows.", phase_percent=55)
        status_event_df, obsoletion_summary = build_obsoletion_detail(tables.get("audit_history", pd.DataFrame()), source_rows_df, indexes, indexes.get("users", {}), logger, cancel)
        logger.emit(action="COMPLETE", result="OK", detail="Actor and obsoletion mapping complete.", phase_percent=100)
        (
            history_df,
            approver_detail_df,
            reviewer_detail_df,
            status_event_df,
            identity_diag_df,
            technical_df,
            stats,
        ) = build_main_history_and_supporting_sheets(
            source_rows_df,
            approver_detail_df,
            reviewer_detail_df,
            status_event_df,
            approver_summary,
            reviewer_summary,
            obsoletion_summary,
            logger,
            cancel,
        )
        history_df = generate_pdf_packages(inputs, outputs, history_df, logger, cancel)
        timeline_df = build_document_version_timeline(history_df, source_rows_df, logger)
        source_field_map_df = make_source_field_map()
        run_summary_df = make_run_summary(inputs, outputs, stats, encodings)
        write_workbook(
            outputs,
            history_df,
            timeline_df,
            source_rows_df,
            justification_diag_df,
            approver_detail_df,
            reviewer_detail_df,
            status_event_df,
            identity_diag_df,
            source_field_map_df,
            run_summary_df,
            logger.event_rows,
            technical_df,
            inputs.include_event_log_sheet,
            inputs.include_technical_evidence_sheet,
            logger,
)
        logger.emit("WRITE_WORKBOOK", action="RUN_COMPLETE", result="OK", detail=str(outputs.run_folder), phase_percent=100)
        ui_queue.put({"_run_complete": True, "status": "Complete", "run_folder": str(outputs.run_folder), "workbook_path": str(outputs.workbook_path)})
        return outputs
    except CancelledRun as exc:
        logger.emit(level="WARNING", action="RUN_CANCELLED", result="CANCELLED", detail=str(exc))
        safe_write_event_log(outputs.event_log_csv_path, logger.event_rows)
        ui_queue.put({"_run_complete": True, "status": "Cancelled", "run_folder": str(outputs.run_folder), "error": str(exc)})
        raise
    except Exception as exc:
        logger.emit(level="ERROR", action="RUN_FAILED", result="FAIL", detail=f"{exc}\n{traceback.format_exc()}")
        safe_write_event_log(outputs.event_log_csv_path, logger.event_rows)
        ui_queue.put({"_run_complete": True, "status": "Failed", "run_folder": str(outputs.run_folder), "error": str(exc)})
        raise


def safe_write_event_log(path: Path, rows: List[Dict[str, Any]]) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(rows, columns=EVENT_COLUMNS).to_csv(path, index=False, encoding="utf-8-sig")
    except Exception:
        pass


class LiveEventLogWindow:
    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        self.window: Optional[tk.Toplevel] = None
        self.text: Optional[ScrolledText] = None
        self.autoscroll_var = tk.BooleanVar(value=True)

    def show(self) -> None:
        if self.window is not None and self.window.winfo_exists():
            self.window.deiconify()
            self.window.lift()
            return
        self.window = tk.Toplevel(self.master)
        self.window.title("NEUROTIC DMS Live Event Log")
        self.window.geometry("1250x650")
        top = ttk.Frame(self.window, padding=6)
        top.pack(fill=tk.X)
        ttk.Checkbutton(top, text="Auto-scroll", variable=self.autoscroll_var).pack(side=tk.LEFT)
        ttk.Button(top, text="Clear Display Only", command=self.clear).pack(side=tk.LEFT, padx=8)
        self.text = ScrolledText(self.window, wrap=tk.NONE, font=("Consolas", 9))
        self.text.pack(fill=tk.BOTH, expand=True)

    def append(self, line: str) -> None:
        self.show()
        if self.text is None:
            return
        self.text.insert(tk.END, line + "\n")
        if self.autoscroll_var.get():
            self.text.see(tk.END)

    def clear(self) -> None:
        if self.text is not None:
            self.text.delete("1.0", tk.END)


class DocumentVersionHistoryApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{APP_TITLE} v{APP_VERSION}")
        self.root.geometry("1180x780")
        self.ui_queue: "queue.Queue[Dict[str, Any]]" = queue.Queue()
        self.cancel_token = CancelToken()
        self.worker_thread: Optional[threading.Thread] = None
        self.latest_output_folder: Optional[Path] = None
        self.log_window = LiveEventLogWindow(root)
        self.path_vars: Dict[str, tk.StringVar] = {key: tk.StringVar() for key in [
            "raw_export_folder",
            "document_version_csv",
            "corporate_document_csv",
            "user_csv",
            "review_justification_csv",
            "approvers_file",
            "reviewers_file",
            "effective_version_details_csv",
            "audit_activity_type_csv",
            "audit_history_csv",
            "manifest_csv",
            "raw_dms_file_tree_folder",
            "output_run_folder",
        ]}
        self.generate_pdf_packages_var = tk.BooleanVar(value=False)
        self.include_pdf_cover_page_var = tk.BooleanVar(value=False)
        self.include_technical_evidence_var = tk.BooleanVar(value=True)
        self.include_event_log_var = tk.BooleanVar(value=True)
        self.fail_missing_optional_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Ready")
        self.percent_var = tk.StringVar(value="000%")
        self.phase_var = tk.StringVar(value="Phase: —")
        self.current_item_var = tk.StringVar(value="Current item: —")
        self._build_gui()
        self.root.after(100, self._poll_queue)

    def _build_gui(self) -> None:
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill=tk.BOTH, expand=True)
        title = ttk.Label(outer, text=f"{APP_TITLE} v{APP_VERSION}", font=("Segoe UI", 16, "bold"))
        title.pack(anchor=tk.W)

        autodetect = ttk.LabelFrame(outer, text="Auto-detect raw export folder", padding=8)
        autodetect.pack(fill=tk.X, pady=(10, 6))
        self._path_row(autodetect, "Top-level Raw TrackWise Export Folder", "raw_export_folder", folder=True, row=0)
        ttk.Button(autodetect, text="Auto-detect CSVs", command=self.auto_detect_csvs).grid(row=0, column=3, padx=8)

        required = ttk.LabelFrame(outer, text="Required raw TrackWise Salesforce inputs", padding=8)
        required.pack(fill=tk.X, pady=6)
        self._path_row(required, "Document Version CSV", "document_version_csv", row=0)
        self._path_row(required, "Corporate Document CSV", "corporate_document_csv", row=1)
        self._path_row(required, "User CSV", "user_csv", row=2)

        evidence = ttk.LabelFrame(outer, text="Revision/workflow/audit evidence inputs", padding=8)
        evidence.pack(fill=tk.X, pady=6)
        self._path_row(evidence, "Review Justification CSV", "review_justification_csv", row=0)
        self._path_row(evidence, "Document Approvers CSV", "approvers_file", row=1)
        self._path_row(evidence, "Document Reviewers CSV", "reviewers_file", row=2)
        self._path_row(evidence, "Effective Version Details CSV", "effective_version_details_csv", row=3)
        self._path_row(evidence, "Audit Activity Type CSV", "audit_activity_type_csv", row=4)
        self._path_row(evidence, "Audit History CSV", "audit_history_csv", row=5)

        optional = ttk.LabelFrame(outer, text="Optional evidence and PDF packaging inputs", padding=8)
        optional.pack(fill=tk.X, pady=6)
        self._path_row(optional, "Neurotic DMS Manifest CSV", "manifest_csv", row=0)
        self._path_row(optional, "Raw DMS File Tree Folder", "raw_dms_file_tree_folder", folder=True, row=1)
        self._path_row(optional, "Output Run Folder", "output_run_folder", folder=True, row=2)

        checks = ttk.Frame(outer)
        checks.pack(fill=tk.X, pady=8)
        ttk.Checkbutton(checks, text="Generate clean version PDFs", variable=self.generate_pdf_packages_var).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Checkbutton(checks, text="Include PDF Cover Page", variable=self.include_pdf_cover_page_var).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Checkbutton(checks, text="Include Technical Evidence sheet", variable=self.include_technical_evidence_var).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Checkbutton(checks, text="Include Event Log sheet", variable=self.include_event_log_var).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Checkbutton(checks, text="Fail if populated optional file path does not exist", variable=self.fail_missing_optional_var).pack(side=tk.LEFT, padx=(0, 16))

        controls = ttk.Frame(outer)
        controls.pack(fill=tk.X, pady=8)
        self.run_button = ttk.Button(controls, text="Run", command=self.start_run)
        self.run_button.pack(side=tk.LEFT, padx=(0, 8))
        self.cancel_button = ttk.Button(controls, text="Cancel", command=self.cancel_run, state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(controls, text="Open Live Event Log", command=self.log_window.show).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(controls, text="Open Output Folder", command=self.open_output_folder).pack(side=tk.LEFT, padx=(0, 8))

        status = ttk.LabelFrame(outer, text="Run Status", padding=8)
        status.pack(fill=tk.X, pady=8)
        ttk.Label(status, textvariable=self.percent_var, font=("Segoe UI", 20, "bold"), width=8).grid(row=0, column=0, rowspan=3, sticky=tk.W)
        ttk.Label(status, textvariable=self.status_var, font=("Segoe UI", 11, "bold")).grid(row=0, column=1, sticky=tk.W)
        ttk.Label(status, textvariable=self.phase_var).grid(row=1, column=1, sticky=tk.W)
        ttk.Label(status, textvariable=self.current_item_var).grid(row=2, column=1, sticky=tk.W)
        status.columnconfigure(1, weight=1)

    def _path_row(self, parent: tk.Widget, label: str, key: str, row: int, folder: bool = False) -> None:
        ttk.Label(parent, text=label, width=34).grid(row=row, column=0, sticky=tk.W, pady=2)
        entry = ttk.Entry(parent, textvariable=self.path_vars[key], width=112)
        entry.grid(row=row, column=1, sticky=tk.EW, pady=2, padx=4)
        command = (lambda k=key: self.browse_folder(k)) if folder else (lambda k=key: self.browse_file(k))
        ttk.Button(parent, text="Browse", command=command).grid(row=row, column=2, sticky=tk.E, pady=2)
        parent.columnconfigure(1, weight=1)

    def browse_file(self, key: str) -> None:
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.path_vars[key].set(path)

    def browse_folder(self, key: str) -> None:
        path = filedialog.askdirectory()
        if path:
            self.path_vars[key].set(path)

    def auto_detect_csvs(self) -> None:
        root_text = self.path_vars["raw_export_folder"].get().strip()
        if not root_text:
            messagebox.showerror("Missing Folder", "Select a top-level Raw TrackWise Export Folder first.")
            return
        root = Path(root_text)
        if not root.exists():
            messagebox.showerror("Folder Not Found", f"Folder does not exist: {root}")
            return
        expected_lower = {file_name.lower(): key for key, file_name in EXPECTED_RAW_EXPORT_FILENAMES.items()}
        found: Dict[str, Path] = {}
        for current_root, _, files in os.walk(root):
            for filename in files:
                key = expected_lower.get(filename.lower())
                if key and key not in found:
                    found[key] = Path(current_root) / filename
        for key, path in found.items():
            self.path_vars[key].set(str(path))
        if not self.path_vars["output_run_folder"].get().strip():
            self.path_vars["output_run_folder"].set(str(root / "NEUROTIC_DMS_OUTPUTS"))
        self.status_var.set(f"Auto-detected {len(found)} expected CSV file(s).")

    def start_run(self) -> None:
        if self.worker_thread is not None and self.worker_thread.is_alive():
            messagebox.showwarning("Run Active", "A run is already active.")
            return
        try:
            inputs = self._collect_inputs()
        except Exception as exc:
            messagebox.showerror("Invalid Inputs", str(exc))
            return
        self.cancel_token = CancelToken()
        self.status_var.set("Running")
        self.percent_var.set("000%")
        self.phase_var.set("Phase: starting")
        self.current_item_var.set("Current item: —")
        self.run_button.configure(state=tk.DISABLED)
        self.cancel_button.configure(state=tk.NORMAL)
        self.log_window.show()
        self.worker_thread = threading.Thread(target=lambda: self._worker(inputs), daemon=True)
        self.worker_thread.start()

    def _collect_inputs(self) -> RunInputs:
        def required_path(key: str, label: str) -> Path:
            text = self.path_vars[key].get().strip()
            if not text:
                raise RuntimeError(f"Missing required input: {label}")
            return Path(text)

        def optional_path(key: str) -> Optional[Path]:
            text = self.path_vars[key].get().strip()
            return Path(text) if text else None

        return RunInputs(
            document_version_csv=required_path("document_version_csv", "Document Version CSV"),
            corporate_document_csv=required_path("corporate_document_csv", "Corporate Document CSV"),
            user_csv=required_path("user_csv", "User CSV"),
            output_run_folder=required_path("output_run_folder", "Output Run Folder"),
            review_justification_csv=optional_path("review_justification_csv"),
            approvers_file=optional_path("approvers_file"),
            reviewers_file=optional_path("reviewers_file"),
            effective_version_details_csv=optional_path("effective_version_details_csv"),
            audit_activity_type_csv=optional_path("audit_activity_type_csv"),
            audit_history_csv=optional_path("audit_history_csv"),
            manifest_csv=optional_path("manifest_csv"),
            raw_dms_file_tree_folder=optional_path("raw_dms_file_tree_folder"),
            generate_pdf_packages=self.generate_pdf_packages_var.get(),
            include_pdf_cover_page=self.include_pdf_cover_page_var.get(),
            include_technical_evidence_sheet=self.include_technical_evidence_var.get(),
            include_event_log_sheet=self.include_event_log_var.get(),
            fail_if_populated_optional_missing=self.fail_missing_optional_var.get(),
        )

    def _worker(self, inputs: RunInputs) -> None:
        try:
            outputs = run_document_history_generation(inputs, self.ui_queue, self.cancel_token)
            self.latest_output_folder = outputs.run_folder
        except CancelledRun:
            pass
        except Exception:
            pass

    def cancel_run(self) -> None:
        self.cancel_token.cancel()
        self.status_var.set("Cancelling...")

    def _poll_queue(self) -> None:
        drained = 0
        while drained < 200:
            try:
                record = self.ui_queue.get_nowait()
            except queue.Empty:
                break
            drained += 1
            if record.get("_run_complete"):
                self._handle_run_complete(record)
                continue
            self._handle_event_record(record)
        self.root.after(75, self._poll_queue)

    def _handle_event_record(self, record: Dict[str, Any]) -> None:
        percent = int(record.get("percent_complete") or 0)
        self.percent_var.set(f"{percent:03d}%")
        status = clean_scalar(record.get("result", "")) or "Running"
        self.status_var.set(status)
        phase_no = int(record.get("phase_number") or 0)
        phase_total = int(record.get("phase_total") or 0)
        phase = clean_scalar(record.get("phase", ""))
        phase_pct = int(record.get("phase_percent_complete") or 0)
        self.phase_var.set(f"Phase: {phase_no} of {phase_total} - {phase} ({phase_pct:03d}%)")
        current = clean_scalar(record.get("full_document_version", "")) or clean_scalar(record.get("justification_id", "")) or clean_scalar(record.get("source_file", "")) or clean_scalar(record.get("action", ""))
        self.current_item_var.set(f"Current item: {current or '—'}")
        self.log_window.append(format_event_line(record))

    def _handle_run_complete(self, record: Dict[str, Any]) -> None:
        status = clean_scalar(record.get("status", "Complete"))
        self.status_var.set(status)
        self.cancel_button.configure(state=tk.DISABLED)
        self.run_button.configure(state=tk.NORMAL)
        run_folder = clean_scalar(record.get("run_folder", ""))
        if run_folder:
            self.latest_output_folder = Path(run_folder)
        if status == "Complete":
            self.percent_var.set("100%")
        error = clean_scalar(record.get("error", ""))
        if error:
            self.current_item_var.set(f"Current item: {error}")
            if status == "Failed":
                messagebox.showerror("Run Failed", error)
        else:
            self.current_item_var.set("Current item: —")

    def open_output_folder(self) -> None:
        path = self.latest_output_folder
        if path is None:
            text = self.path_vars["output_run_folder"].get().strip()
            path = Path(text) if text else None
        if path is None or not path.exists():
            messagebox.showwarning("No Output Folder", "No output folder is available yet.")
            return
        open_folder(path)


def open_folder(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


def main() -> None:
    root = tk.Tk()
    DocumentVersionHistoryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
