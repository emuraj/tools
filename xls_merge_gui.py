#!/usr/bin/env python3
# xls_merge_gui.py

"""
Excel Merge GUI (Tkinter) + CSV UTF-8 + PDF Summary Report

KEY BEHAVIOR (CURRENT)
- Merge input Excel files (.xls/.xlsx) into a normalized dataset
- Transformations:
  1) Document ID column from Training Number:
     - Numeric-start -> SOP-#### where #### is first segment (e.g., 2052-000002 -> SOP-2052)
     - Prefix-start -> remove only trailing instance suffix of 5+ digits
       (e.g., PRN-0015-000001 -> PRN-0015; TRN-016 stays TRN-016)
  2) Exclude CORP docs entirely (Training Number starts with CORP / CORP-)
  3) Split Completed Date/Time into Date + Time (keep original)
  4) Split Trainee Name into First + Last (keep original)

OUTPUTS
- XLSX and CSV are written as either:
  - single file each (if no splitting enabled), OR
  - multiple files (NO SHEET TABS) when splitting is enabled:
      * Split by year: YEAR_2020_PART_001.xlsx, YEAR_2021_PART_001.xlsx, YEAR_Unknown_PART_001.xlsx, ...
      * Max rows: PART_001.xlsx, PART_002.xlsx, ...
      * Additive: year split + max rows -> year-part chunking

- PDF report is Option A: key/value blocks with wrapping (no overflow tables).

DEPENDENCIES
- pandas, openpyxl, reportlab
- Optional for .xls:
  - xlrd OR Excel COM (win32com) OR LibreOffice soffice conversion
"""

import os
import sys
import re
import csv
import json
import time
import socket
import getpass
import shutil
import hashlib
import tempfile
import platform
import traceback
import subprocess
from dataclasses import dataclass
from datetime import datetime
from typing import List, Tuple, Dict, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

# PDF (ReportLab)
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas


# -----------------------------
# Configuration
# -----------------------------

CANONICAL_COLS = [
    "Training Number",
    "Training Name",
    "Version",
    "Trainee Name",
    "Completed Date/Time",
    "Training Status",
]

COLUMN_SYNONYMS = {
    "training number": "Training Number",
    "training no": "Training Number",
    "training #": "Training Number",
    "training name": "Training Name",
    "version": "Version",
    "trainee name": "Trainee Name",
    "completed date/time": "Completed Date/Time",
    "completed date": "Completed Date/Time",
    "completion date": "Completed Date/Time",
    "training status": "Training Status",
    "status": "Training Status",
}

DEFAULT_UTF8_BOM = False
DEFAULT_SPLIT_BY_YEAR = False
DEFAULT_MAX_ROWS_PER_PART = "4900"


# -----------------------------
# Utility helpers
# -----------------------------

def mmddyyyy_hhmmss() -> str:
    return datetime.now().strftime("%m%d%Y_%H%M%S")


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def ensure_dir(p: str):
    os.makedirs(p, exist_ok=True)


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def safe_sheet_title(s: str, max_len: int = 31) -> str:
    bad = set(r':\/?*[]')
    cleaned = "".join("_" if c in bad else c for c in s).strip()
    if not cleaned:
        cleaned = "Sheet"
    return cleaned[:max_len]


def safe_filename(s: str, max_len: int = 120) -> str:
    s = str(s).strip()
    s = re.sub(r"[^\w\-.]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    if not s:
        s = "part"
    return s[:max_len]


def has_module(name: str) -> bool:
    try:
        __import__(name)
        return True
    except Exception:
        return False


def normalize_colname(col: str) -> str:
    if col is None:
        return col
    key = str(col).strip().lower()
    key = re.sub(r"\s+", " ", key)
    return COLUMN_SYNONYMS.get(key, str(col).strip())


def convert_xls_to_xlsx(input_path: str, temp_dir: str, log_fn) -> str:
    base = os.path.splitext(os.path.basename(input_path))[0]
    out_path = os.path.join(temp_dir, base + ".xlsx")

    # Excel COM
    if sys.platform.startswith("win") and has_module("win32com.client"):
        try:
            log_fn(f"Converting via Excel COM: {input_path}")
            import win32com.client  # type: ignore
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False
            wb = excel.Workbooks.Open(os.path.abspath(input_path))
            wb.SaveAs(os.path.abspath(out_path), FileFormat=51)  # xlsx
            wb.Close(False)
            excel.Quit()
            if os.path.exists(out_path):
                return out_path
        except Exception as e:
            log_fn(f"Excel COM conversion failed: {e}")

    # LibreOffice
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            log_fn(f"Converting via LibreOffice: {input_path}")
            cmd = [soffice, "--headless", "--convert-to", "xlsx", "--outdir", temp_dir, input_path]
            res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if res.returncode != 0:
                raise RuntimeError(res.stderr.strip() or res.stdout.strip())
            if os.path.exists(out_path):
                return out_path
            candidates = [
                os.path.join(temp_dir, f) for f in os.listdir(temp_dir)
                if f.lower().endswith(".xlsx") and os.path.splitext(f)[0].startswith(base)
            ]
            if candidates:
                return max(candidates, key=os.path.getmtime)
        except Exception as e:
            log_fn(f"LibreOffice conversion failed: {e}")

    raise RuntimeError("Cannot read .xls without xlrd, Excel (win32com), or LibreOffice/soffice.")


def read_all_sheets_as_df(file_path: str, log_fn) -> List[Tuple[str, pd.DataFrame]]:
    ext = os.path.splitext(file_path)[1].lower()
    temp_dir = None
    path_to_read = file_path

    try:
        if ext == ".xls":
            if has_module("xlrd"):
                log_fn(f"Reading .xls with xlrd: {file_path}")
                xls = pd.ExcelFile(file_path, engine="xlrd")
                out = []
                for sh in xls.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sh, engine="xlrd")
                    out.append((sh, df))
                return out

            temp_dir = tempfile.mkdtemp(prefix="xls_convert_")
            path_to_read = convert_xls_to_xlsx(file_path, temp_dir, log_fn)

        log_fn(f"Reading workbook: {path_to_read}")
        xls = pd.ExcelFile(path_to_read, engine="openpyxl")
        out = []
        for sh in xls.sheet_names:
            df = pd.read_excel(path_to_read, sheet_name=sh, engine="openpyxl")
            out.append((sh, df))
        return out

    finally:
        if temp_dir and os.path.isdir(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


def coerce_to_canonical(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict]:
    diag = {
        "original_columns": [str(c) for c in df.columns],
        "normalized_columns": None,
        "missing_canonical_columns": [],
        "extra_columns": [],
    }

    new_cols = [normalize_colname(c) for c in df.columns]
    df = df.copy()
    df.columns = new_cols
    diag["normalized_columns"] = [str(c) for c in df.columns]

    for col in CANONICAL_COLS:
        if col not in df.columns:
            df[col] = pd.NA
            diag["missing_canonical_columns"].append(col)

    extras = [c for c in df.columns if c not in CANONICAL_COLS]
    diag["extra_columns"] = extras

    df = df[CANONICAL_COLS + extras]
    return df, diag


def count_blank_rows(df: pd.DataFrame, subset_cols: List[str]) -> int:
    work = df[subset_cols].copy()
    for c in subset_cols:
        work[c] = work[c].astype("string").str.strip()
    all_blank = work.isna() | (work == "")
    return int(all_blank.all(axis=1).sum())


def df_basic_stats(df: pd.DataFrame) -> Dict:
    stats = {}
    stats["rows"] = int(len(df))
    stats["cols"] = int(len(df.columns))
    stats["blank_rows_canonical"] = count_blank_rows(df, CANONICAL_COLS)
    try:
        stats["duplicate_rows_canonical"] = int(df.duplicated(subset=CANONICAL_COLS, keep="first").sum())
    except Exception:
        stats["duplicate_rows_canonical"] = None

    for col in ["Training Number", "Training Name", "Trainee Name", "Training Status"]:
        if col in df.columns:
            try:
                stats[f"unique_{col}"] = int(df[col].astype("string").nunique(dropna=True))
            except Exception:
                stats[f"unique_{col}"] = None
    return stats


def chunk_dataframe(df: pd.DataFrame, max_rows: int) -> List[pd.DataFrame]:
    if max_rows is None or max_rows <= 0:
        return [df]
    n = len(df)
    if n <= max_rows:
        return [df]
    return [df.iloc[i:i+max_rows].copy() for i in range(0, n, max_rows)]


def build_output_parts(
    df: pd.DataFrame,
    year_key: Optional[pd.Series],
    split_by_year: bool,
    max_rows: int
) -> List[Tuple[str, pd.DataFrame]]:
    """
    Returns list of (part_label, dataframe) where each label corresponds to one output file.
    Additive logic:
      - if split_by_year: group by year_key (else all in one group)
      - if max_rows>0: chunk each group to parts
    """
    parts: List[Tuple[str, pd.DataFrame]] = []
    chunking = bool(max_rows and max_rows > 0)

    if not split_by_year:
        if not chunking:
            return [("MERGED", df)]
        chunks = chunk_dataframe(df, max_rows)
        for i, ch in enumerate(chunks, start=1):
            parts.append((f"PART_{i:03d}", ch))
        return parts

    # split by year
    if year_key is None or len(year_key) != len(df):
        year_key = pd.Series(["Unknown"] * len(df), index=df.index)

    yseries = pd.Series(year_key).astype("string").fillna("Unknown")
    df2 = df.copy()
    df2["_YearKey"] = yseries.values

    keys = list(df2["_YearKey"].unique())

    def sort_key(k):
        if str(k) == "Unknown":
            return (1, 999999)
        try:
            return (0, int(str(k)))
        except Exception:
            return (0, 999998)

    for k in sorted(keys, key=sort_key):
        g = df2[df2["_YearKey"] == k].drop(columns=["_YearKey"])
        if len(g) == 0:
            continue
        base = f"YEAR_{k}"
        if not chunking:
            parts.append((base, g))
        else:
            chunks = chunk_dataframe(g, max_rows)
            for i, ch in enumerate(chunks, start=1):
                parts.append((f"{base}_PART_{i:03d}", ch))

    return parts


# -----------------------------
# Evidence structures
# -----------------------------

@dataclass
class FileEvidence:
    path: str
    size_bytes: int
    mtime_iso: str
    sha256: str


def file_evidence(path: str) -> FileEvidence:
    st = os.stat(path)
    return FileEvidence(
        path=os.path.abspath(path),
        size_bytes=int(st.st_size),
        mtime_iso=datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
        sha256=sha256_file(path),
    )


# -----------------------------
# PDF Report (Option A: key/value blocks)
# -----------------------------

class NumberedRunCanvas(Canvas):
    def __init__(self, *args, **kwargs):
        Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []
        self._run_name = ""
        self._header_y = 10.75 * inch
        self._footer_y = 0.6 * inch
        self._left_x = 0.75 * inch

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.setFont("Helvetica", 9)
            self.drawString(self._left_x, self._header_y, self._run_name)
            footer = f"{self._run_name}  Page {self.getPageNumber()} of {num_pages}"
            self.drawString(self._left_x, self._footer_y, footer)
            Canvas.showPage(self)
        Canvas.save(self)


def build_kv_block(title: str, kv: List[Tuple[str, str]], styles) -> KeepTogether:
    """
    Build a wrapped key/value block that will not overflow page width.
    """
    h = Paragraph(f"<b>{title}</b>", styles["h2"])
    rows = []
    for k, v in kv:
        rows.append([
            Paragraph(f"<b>{k}</b>", styles["kv_key"]),
            Paragraph(v, styles["kv_val"])
        ])
    tbl = Table(rows, colWidths=[1.45 * inch, 5.55 * inch])
    tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return KeepTogether([h, Spacer(1, 0.08 * inch), tbl, Spacer(1, 0.15 * inch)])


def build_pdf_report(
    pdf_path: str,
    run_name: str,
    run_meta: dict,
    system_meta: dict,
    input_evidence: List[FileEvidence],
    per_file_stats: List[dict],
    merged_stats: dict,
    outputs_evidence: List[FileEvidence],
    output_parts_manifest: List[dict],
    log_lines: List[str],
):
    base_styles = getSampleStyleSheet()

    styles = {
        "title": base_styles["Title"],
        "body": base_styles["BodyText"],
        "h2": base_styles["Heading2"],
        "mono": ParagraphStyle(
            "mono", parent=base_styles["BodyText"], fontName="Courier", fontSize=8, leading=9
        ),
        "kv_key": ParagraphStyle(
            "kv_key", parent=base_styles["BodyText"], fontSize=8, leading=9
        ),
        "kv_val": ParagraphStyle(
            "kv_val",
            parent=base_styles["BodyText"],
            fontSize=8,
            leading=9,
            wordWrap="CJK",          # breaks long unspaced strings (hashes/paths)
        ),
    }
    styles["kv_val"].splitLongWords = 1  # type: ignore

    try:
        pdfmetrics.registerFont(TTFont("Consolas", "consola.ttf"))
        styles["mono"].fontName = "Consolas"
    except Exception:
        pass

    story = []
    story.append(Paragraph(f"{run_name} — Merge Summary Report", styles["title"]))
    story.append(Spacer(1, 0.15 * inch))

    # Run meta block
    story.append(build_kv_block(
        "Run Metadata",
        [(k, str(v)) for k, v in run_meta.items()],
        styles
    ))

    # System meta block
    story.append(build_kv_block(
        "System Metadata",
        [(k, str(v)) for k, v in system_meta.items()],
        styles
    ))

    story.append(PageBreak())

    # Inputs evidence: one block per file
    story.append(Paragraph("Input Files (Evidence)", styles["h2"]))
    story.append(Spacer(1, 0.12 * inch))
    for ev in input_evidence:
        story.append(build_kv_block(
            os.path.basename(ev.path),
            [
                ("Path", ev.path),
                ("Size (bytes)", str(ev.size_bytes)),
                ("Modified", ev.mtime_iso),
                ("SHA-256", ev.sha256),
            ],
            styles
        ))

    story.append(PageBreak())

    # Per-file stats: block per (file, sheet)
    story.append(Paragraph("Per-File Data Stats", styles["h2"]))
    story.append(Spacer(1, 0.12 * inch))
    for s in per_file_stats:
        notes = []
        if s.get("missing_canonical_columns"):
            notes.append("Missing canonical: " + ", ".join(s["missing_canonical_columns"]))
        if s.get("extra_columns"):
            notes.append("Extra: " + ", ".join(s["extra_columns"]))
        story.append(build_kv_block(
            f"{s.get('file_base','')} — {s.get('sheet','')}",
            [
                ("Rows", str(s.get("rows", ""))),
                ("Cols", str(s.get("cols", ""))),
                ("Blank rows (canonical)", str(s.get("blank_rows_canonical", ""))),
                ("Dup rows (canonical)", str(s.get("duplicate_rows_canonical", ""))),
                ("Notes", "; ".join(notes) if notes else ""),
            ],
            styles
        ))

    story.append(PageBreak())

    # Merged stats
    story.append(build_kv_block(
        "Merged Output Stats",
        [(k, str(v)) for k, v in merged_stats.items()],
        styles
    ))

    # Output manifest (parts)
    story.append(Paragraph("Output Parts Manifest", styles["h2"]))
    story.append(Spacer(1, 0.12 * inch))
    for p in output_parts_manifest:
        story.append(build_kv_block(
            p.get("label", ""),
            [
                ("Rows", str(p.get("rows", ""))),
                ("Year", str(p.get("year", ""))),
                ("XLSX", str(p.get("xlsx_path", ""))),
                ("CSV", str(p.get("csv_path", ""))),
            ],
            styles
        ))

    story.append(PageBreak())

    # Outputs evidence
    story.append(Paragraph("Output Files (Evidence)", styles["h2"]))
    story.append(Spacer(1, 0.12 * inch))
    for ev in outputs_evidence:
        story.append(build_kv_block(
            os.path.basename(ev.path),
            [
                ("Path", ev.path),
                ("Size (bytes)", str(ev.size_bytes)),
                ("Modified", ev.mtime_iso),
                ("SHA-256", ev.sha256),
            ],
            styles
        ))

    story.append(PageBreak())

    # Log excerpt
    story.append(Paragraph("Run Log (Excerpt)", styles["h2"]))
    story.append(Spacer(1, 0.12 * inch))
    excerpt = log_lines[-500:] if len(log_lines) > 500 else log_lines
    log_text = "<br/>".join((l.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")) for l in excerpt)
    story.append(Paragraph(f"<font name='{styles['mono'].fontName}' size='8'>{log_text}</font>", styles["body"]))

    class RunCanvas(NumberedRunCanvas):
        def __init__(self, *args, **kwargs):
            NumberedRunCanvas.__init__(self, *args, **kwargs)
            self._run_name = run_name

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=1.0 * inch,
        bottomMargin=0.9 * inch,
        title=f"{run_name} Report",
        author=getpass.getuser(),
    )
    doc.build(story, canvasmaker=RunCanvas)


# -----------------------------
# GUI App
# -----------------------------

class MergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XLS Merge (XLSX + CSV UTF-8 + PDF Summary)")
        self.geometry("1250x860")

        self.selected_files: List[str] = []
        self.output_dir_var = tk.StringVar(value="")

        self.write_csv_var = tk.BooleanVar(value=True)
        self.write_xlsx_var = tk.BooleanVar(value=True)
        self.csv_bom_var = tk.BooleanVar(value=DEFAULT_UTF8_BOM)
        self.drop_exact_dups_var = tk.BooleanVar(value=False)

        self.split_by_year_var = tk.BooleanVar(value=DEFAULT_SPLIT_BY_YEAR)
        self.max_rows_var = tk.StringVar(value=DEFAULT_MAX_ROWS_PER_PART)

        self.log_lines: List[str] = []

        self._build_ui()

    def _build_ui(self):
        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)

        controls = ttk.Frame(outer)
        controls.pack(fill="x")

        ttk.Button(controls, text="Select Excel Files...", command=self.on_select_files).pack(side="left")
        ttk.Button(controls, text="Select Output Folder...", command=self.on_select_output).pack(side="left", padx=8)
        ttk.Entry(controls, textvariable=self.output_dir_var).pack(side="left", fill="x", expand=True)
        ttk.Button(controls, text="RUN MERGE", command=self.on_run_merge).pack(side="right")

        opts = ttk.Frame(outer)
        opts.pack(fill="x", pady=(8, 0))

        ttk.Checkbutton(opts, text="Write XLSX", variable=self.write_xlsx_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(opts, text="Write CSV (UTF-8)", variable=self.write_csv_var).grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Checkbutton(opts, text="CSV include BOM (utf-8-sig)", variable=self.csv_bom_var).grid(row=0, column=2, sticky="w", padx=(12, 0))
        ttk.Checkbutton(opts, text="Drop exact duplicates (canonical cols)", variable=self.drop_exact_dups_var).grid(row=0, column=3, sticky="w", padx=(12, 0))

        ttk.Checkbutton(
            opts,
            text="Split outputs by Year (from Completed Date/Time)",
            variable=self.split_by_year_var
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Label(opts, text="Max rows per output file (0=off):").grid(row=1, column=2, sticky="e", padx=(12, 6), pady=(6, 0))
        ttk.Entry(opts, textvariable=self.max_rows_var, width=10).grid(row=1, column=3, sticky="w", pady=(6, 0))

        for i in range(4):
            opts.grid_columnconfigure(i, weight=1)

        mid = ttk.PanedWindow(outer, orient=tk.HORIZONTAL)
        mid.pack(fill="both", expand=True, pady=(10, 0))

        left = ttk.Frame(mid, padding=(0, 0, 8, 0))
        right = ttk.Frame(mid)

        mid.add(left, weight=1)
        mid.add(right, weight=3)

        ttk.Label(left, text="Selected Files").pack(anchor="w")
        self.files_list = tk.Listbox(left, height=18)
        self.files_list.pack(fill="both", expand=True)

        ttk.Label(right, text="Log").pack(anchor="w")
        self.log_text = tk.Text(right, wrap="word", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True)

        self._log("Ready.")

    def _log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        self.log_lines.append(line)
        self.log_text.insert("end", line + "\n")
        self.log_text.see("end")
        self.update_idletasks()

    def on_select_files(self):
        paths = filedialog.askopenfilenames(
            title="Select Excel files",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if not paths:
            return
        self.selected_files = list(paths)
        self.files_list.delete(0, "end")
        for p in self.selected_files:
            self.files_list.insert("end", p)
        self._log(f"Selected {len(self.selected_files)} file(s).")

    def on_select_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if not d:
            return
        self.output_dir_var.set(d)
        self._log(f"Output folder set: {d}")

    def on_run_merge(self):
        if not self.selected_files:
            messagebox.showwarning("No files", "Select one or more Excel files first.")
            return
        out_root = self.output_dir_var.get().strip()
        if not out_root:
            messagebox.showwarning("No output folder", "Select an output folder first.")
            return
        if not (self.write_xlsx_var.get() or self.write_csv_var.get()):
            messagebox.showwarning("No outputs selected", "Select at least one output type (XLSX and/or CSV).")
            return

        # Parse max rows
        try:
            s = (self.max_rows_var.get() or "").strip()
            max_rows = int(s) if s else 0
            if max_rows < 0:
                raise ValueError
        except Exception:
            messagebox.showwarning("Invalid max rows", "Max rows must be a non-negative integer (e.g., 4900).")
            return

        split_by_year = bool(self.split_by_year_var.get())

        run_name = f"XLS_MERGE_{mmddyyyy_hhmmss()}"
        run_dir = os.path.join(out_root, run_name)
        ensure_dir(run_dir)

        self._log(f"RUN START: {run_name}")
        t0 = time.time()

        system_meta = {
            "user": getpass.getuser(),
            "hostname": socket.gethostname(),
            "platform": platform.platform(),
            "python_version": sys.version.replace("\n", " "),
            "cwd": os.getcwd(),
            "script_path": os.path.abspath(sys.argv[0]) if sys.argv and sys.argv[0] else "",
            "timestamp_start": now_iso(),
        }

        input_evidence: List[FileEvidence] = []
        for fp in self.selected_files:
            try:
                input_evidence.append(file_evidence(fp))
            except Exception as e:
                self._log(f"WARNING: cannot evidence file {fp}: {e}")

        per_file_stats: List[dict] = []
        merged_frames: List[pd.DataFrame] = []

        excluded_corp_rows = 0

        try:
            for fp in self.selected_files:
                base = os.path.basename(fp)
                self._log(f"Reading: {fp}")
                sheets = read_all_sheets_as_df(fp, log_fn=self._log)

                if not sheets:
                    self._log(f"WARNING: no sheets found in {fp}")
                    continue

                for sheet_name, df in sheets:
                    df_norm, diag = coerce_to_canonical(df)

                    stats = df_basic_stats(df_norm)
                    stats.update({
                        "file": fp,
                        "file_base": base,
                        "sheet": sheet_name,
                        "missing_canonical_columns": diag.get("missing_canonical_columns", []),
                        "extra_columns": diag.get("extra_columns", []),
                    })
                    per_file_stats.append(stats)

                    merged_frames.append(df_norm)

            if not merged_frames:
                raise RuntimeError("No data frames were loaded from the selected files.")

            merged = pd.concat(merged_frames, ignore_index=True)

            # Drop exact duplicates if requested (pre-transform)
            dropped_dups = 0
            if self.drop_exact_dups_var.get():
                before = len(merged)
                merged = merged.drop_duplicates(subset=CANONICAL_COLS, keep="first")
                dropped_dups = before - len(merged)
                self._log(f"Dropped duplicates (canonical): {dropped_dups}")

            # Exclude CORP
            tn_series = merged.get("Training Number", pd.Series([pd.NA] * len(merged))).astype("string")
            corp_mask = tn_series.str.upper().str.startswith("CORP", na=False)
            excluded_corp_rows = int(corp_mask.sum())
            if excluded_corp_rows:
                self._log(f"Excluding CORP rows: {excluded_corp_rows}")
                merged = merged.loc[~corp_mask].copy()

            # Track extras from pre-transform schema
            extra_cols = [c for c in merged.columns if c not in CANONICAL_COLS]

            # Document ID
            def _compute_document_id(val) -> str:
                if pd.isna(val):
                    return ""
                tn = str(val).strip()
                if tn == "":
                    return ""

                if re.match(r"^\d", tn):
                    first_seg = tn.split("-")[0].strip()
                    return f"SOP-{first_seg}"

                m = re.match(r"^(.*)-(\d+)$", tn)
                if m:
                    suffix = m.group(2)
                    if len(suffix) >= 5:
                        return m.group(1).strip()
                    return tn
                return tn

            merged["Document ID"] = merged["Training Number"].apply(_compute_document_id)

            # Date/Time split
            dt_raw = merged.get("Completed Date/Time", pd.Series([pd.NA] * len(merged)))
            dt_parsed = pd.to_datetime(dt_raw, errors="coerce", infer_datetime_format=True)

            def _fmt_date(x):
                if pd.isna(x):
                    return ""
                try:
                    return f"{x.month}/{x.day}/{x.year}"
                except Exception:
                    return ""

            def _fmt_time(x):
                if pd.isna(x):
                    return ""
                try:
                    t = x.strftime("%I:%M %p")
                    return t[1:] if t.startswith("0") else t
                except Exception:
                    return ""

            merged["Date"] = dt_parsed.apply(_fmt_date)
            merged["Time"] = dt_parsed.apply(_fmt_time)

            # Year key for splitting
            year_key = dt_parsed.dt.year.astype("Int64")
            year_key = year_key.apply(lambda y: str(int(y)) if pd.notna(y) else "Unknown")

            # Trainee name split
            name_raw = merged.get("Trainee Name", pd.Series([pd.NA] * len(merged))).astype("string")

            def _first_name(full: str) -> str:
                if full is None or pd.isna(full):
                    return ""
                s = str(full).strip()
                if not s:
                    return ""
                return s.split()[0]

            def _last_name(full: str) -> str:
                if full is None or pd.isna(full):
                    return ""
                s = str(full).strip()
                if not s:
                    return ""
                parts = s.split()
                return parts[-1] if len(parts) >= 2 else ""

            merged["Trainee First Name"] = name_raw.apply(_first_name)
            merged["Trainee Last Name"] = name_raw.apply(_last_name)

            # Reorder columns
            ordered = [
                "Document ID",
                "Training Number",
                "Training Name",
                "Version",
                "Trainee Name",
                "Trainee First Name",
                "Trainee Last Name",
                "Completed Date/Time",
                "Date",
                "Time",
                "Training Status",
            ]
            for c in extra_cols:
                if c not in ordered and c in merged.columns:
                    ordered.append(c)

            merged = merged[[c for c in ordered if c in merged.columns]]

            # Build output parts (files)
            parts = build_output_parts(
                df=merged,
                year_key=year_key,
                split_by_year=split_by_year,
                max_rows=max_rows
            )

            self._log(f"Output parts: {len(parts)} (split_by_year={split_by_year}, max_rows={max_rows})")

            # Stats
            merged_stats = df_basic_stats(merged)
            merged_stats["dropped_duplicates_canonical"] = dropped_dups
            merged_stats["excluded_CORP_rows"] = excluded_corp_rows
            merged_stats["split_by_year"] = split_by_year
            merged_stats["max_rows_per_part"] = max_rows
            merged_stats["output_part_count"] = len(parts)

            try:
                merged_stats["unique_Document ID"] = int(merged["Document ID"].astype("string").nunique(dropna=True))
            except Exception:
                merged_stats["unique_Document ID"] = None

            # Outputs list + manifest for PDF
            outputs: List[str] = []
            output_parts_manifest: List[dict] = []

            # Write XLSX/CSV as MULTIPLE FILES when parts>1 OR when split_by_year/chunking enabled
            # Always write part files if split_by_year True or max_rows>0. Otherwise single merged.
            # (This matches your “no tabs; individual sheets” requirement.)
            write_parts = (split_by_year or (max_rows and max_rows > 0))

            encoding = "utf-8-sig" if self.csv_bom_var.get() else "utf-8"

            if self.write_xlsx_var.get() or self.write_csv_var.get():
                if write_parts:
                    for label, df_part in parts:
                        base = safe_filename(f"{run_name}_{label}")
                        xlsx_path = os.path.join(run_dir, f"{base}.xlsx")
                        csv_path = os.path.join(run_dir, f"{base}.csv")

                        if self.write_xlsx_var.get():
                            self._log(f"Writing XLSX part: {xlsx_path} ({len(df_part)} rows)")
                            with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
                                df_part.to_excel(writer, sheet_name="DATA", index=False)
                            outputs.append(xlsx_path)

                        if self.write_csv_var.get():
                            self._log(f"Writing CSV part: {csv_path} ({len(df_part)} rows)")
                            df_part.to_csv(csv_path, index=False, encoding=encoding, quoting=csv.QUOTE_MINIMAL)
                            outputs.append(csv_path)

                        # For manifest, try to extract year from label if present
                        year_val = ""
                        m = re.search(r"YEAR_([^_]+)", label)
                        if m:
                            year_val = m.group(1)

                        output_parts_manifest.append({
                            "label": label,
                            "rows": int(len(df_part)),
                            "year": year_val,
                            "xlsx_path": xlsx_path if self.write_xlsx_var.get() else "",
                            "csv_path": csv_path if self.write_csv_var.get() else "",
                        })

                else:
                    # Single outputs
                    xlsx_path = os.path.join(run_dir, f"{run_name}_merged.xlsx")
                    csv_path = os.path.join(run_dir, f"{run_name}_merged.csv")

                    if self.write_xlsx_var.get():
                        self._log(f"Writing XLSX: {xlsx_path}")
                        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
                            merged.to_excel(writer, sheet_name="DATA", index=False)
                        outputs.append(xlsx_path)

                    if self.write_csv_var.get():
                        self._log(f"Writing CSV: {csv_path}")
                        merged.to_csv(csv_path, index=False, encoding=encoding, quoting=csv.QUOTE_MINIMAL)
                        outputs.append(csv_path)

                    output_parts_manifest.append({
                        "label": "MERGED",
                        "rows": int(len(merged)),
                        "year": "",
                        "xlsx_path": xlsx_path if self.write_xlsx_var.get() else "",
                        "csv_path": csv_path if self.write_csv_var.get() else "",
                    })

            # JSON meta
            pdf_path = os.path.join(run_dir, f"{run_name}_report.pdf")
            json_path = os.path.join(run_dir, f"{run_name}_run_meta.json")

            t1 = time.time()
            run_meta = {
                "run_name": run_name,
                "run_dir": os.path.abspath(run_dir),
                "timestamp_start": system_meta["timestamp_start"],
                "timestamp_end": now_iso(),
                "duration_seconds": round(t1 - t0, 3),
                "input_count": len(self.selected_files),
                "output_xlsx": bool(self.write_xlsx_var.get()),
                "output_csv_utf8": bool(self.write_csv_var.get()),
                "csv_bom_utf8_sig": bool(self.csv_bom_var.get()),
                "drop_exact_duplicates_canonical": bool(self.drop_exact_dups_var.get()),
                "split_by_year": split_by_year,
                "max_rows_per_part": max_rows,
                "excluded_CORP_rows": excluded_corp_rows,
                "output_part_files_mode": bool(write_parts),
                "output_part_count": int(len(output_parts_manifest)),
            }

            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "run_meta": run_meta,
                        "system_meta": system_meta,
                        "inputs": [ev.__dict__ for ev in input_evidence],
                        "per_file_stats": per_file_stats,
                        "merged_stats": merged_stats,
                        "output_parts_manifest": output_parts_manifest,
                    },
                    f,
                    indent=2
                )
            outputs.append(json_path)

            # Evidence for outputs (excluding PDF until created)
            outputs_evidence: List[FileEvidence] = []
            for op in outputs:
                try:
                    outputs_evidence.append(file_evidence(op))
                except Exception as e:
                    self._log(f"WARNING: cannot evidence output {op}: {e}")

            # Build PDF report (Option A)
            self._log(f"Building PDF report: {pdf_path}")
            build_pdf_report(
                pdf_path=pdf_path,
                run_name=run_name,
                run_meta=run_meta,
                system_meta=system_meta,
                input_evidence=input_evidence,
                per_file_stats=per_file_stats,
                merged_stats=merged_stats,
                outputs_evidence=outputs_evidence,
                output_parts_manifest=output_parts_manifest,
                log_lines=self.log_lines,
            )

            try:
                outputs_evidence.append(file_evidence(pdf_path))
            except Exception:
                pass

            self._log(f"RUN COMPLETE: {run_name}")
            self._log(f"Run folder: {run_dir}")

            # Show a concise completion dialog (avoid listing hundreds of parts)
            msg_lines = [
                f"Run folder:\n{run_dir}",
                "",
                f"Parts created: {len(output_parts_manifest)}",
                f"CORP rows excluded: {excluded_corp_rows}",
            ]
            messagebox.showinfo("Done", "\n".join(msg_lines))

        except Exception:
            self._log("ERROR:\n" + traceback.format_exc())
            messagebox.showerror("Run failed", "See log for details.")


if __name__ == "__main__":
    app = MergeApp()
    app.mainloop()
