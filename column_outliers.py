#!/usr/bin/env python3
# compare_spreadsheets_gui_with_pdf.py
#
# GUI to compare a selected column between two spreadsheets (Excel or CSV),
# print mismatch/outlier values, and generate a detailed PDF report with run metadata.
#
# Supports: .xlsx, .xlsm, .xls, .csv
#
# Requirements:
#   pip install pandas openpyxl reportlab

from __future__ import annotations

import os
import platform
import socket
import getpass
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import pandas as pd

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
)
from reportlab.pdfgen import canvas


SUPPORTED_EXTS = (".xlsx", ".xlsm", ".xls", ".csv")


def is_excel(path: str) -> bool:
    ext = os.path.splitext(path)[1].lower()
    return ext in (".xlsx", ".xlsm", ".xls")


def safe_basename(path: str) -> str:
    return os.path.basename(path) if path else ""


def stamp_ddmmyyyy_hhmmss() -> str:
    return datetime.now().strftime("%d%m%Y_%H%M%S")


def stamp_human() -> str:
    # Local computer time with timezone offset, e.g. 2025-12-23 14:31:05 -0300
    dt = datetime.now().astimezone()
    return dt.strftime("%Y-%m-%d %H:%M:%S %z")

def normalize_series(s: pd.Series, *, trim: bool, case_insensitive: bool) -> pd.Series:
    s2 = s.astype("string")
    if trim:
        s2 = s2.str.strip()
    if case_insensitive:
        s2 = s2.str.upper()
    s2 = s2.replace("", pd.NA)
    return s2


def get_excel_sheet_names(path: str) -> list[str]:
    xls = pd.ExcelFile(path)
    return list(xls.sheet_names)


def read_headers(path: str, sheet_name: str | None = None) -> list[str]:
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(path, nrows=0)
        return list(df.columns)

    if is_excel(path):
        df = pd.read_excel(path, sheet_name=sheet_name, nrows=0, engine="openpyxl")
        return list(df.columns)

    raise ValueError(f"Unsupported file type: {ext}")


def read_single_column(path: str, column_name: str, sheet_name: str | None = None) -> pd.Series:
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(path, usecols=[column_name], dtype="string")
        return df[column_name]

    if is_excel(path):
        df = pd.read_excel(path, sheet_name=sheet_name, usecols=[column_name], dtype="string", engine="openpyxl")
        return df[column_name]

    raise ValueError(f"Unsupported file type: {ext}")


def build_row_index_map(values: pd.Series) -> dict[str, list[int]]:
    """
    Map: normalized_value -> [excel_like_row_numbers]
    Assumes header is row 1, first data row is row 2 (index 0 -> row 2).
    """
    out: dict[str, list[int]] = {}
    for idx, v in values.items():
        if pd.isna(v):
            continue
        key = str(v)
        out.setdefault(key, []).append(int(idx) + 2)
    return out


def format_rows(rows: list[int], max_show: int = 30) -> str:
    if not rows:
        return "-"
    rows_sorted = sorted(rows)
    if len(rows_sorted) <= max_show:
        return ", ".join(map(str, rows_sorted))
    head = ", ".join(map(str, rows_sorted[:max_show]))
    return f"{head}, ... (+{len(rows_sorted) - max_show} more)"


# -----------------------------
# PDF: Page X of N + header/footer (FIXED)
# -----------------------------
class NumberedCanvas(canvas.Canvas):
    """
    Correct two-pass canvas for Platypus:
    - First pass: collect page states without writing pages to output.
    - Second pass: replay pages once, drawing header/footer with 'page X of N'.

    Header/footer text is the same string per your requirement:
      "<report_name>   page X of N"
    """
    def __init__(self, *args, report_name: str = "Report", **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states: list[dict] = []
        self.report_name = report_name

    def showPage(self):
        # Save state but do NOT write the page yet.
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_pages = len(self._saved_page_states)

        for page_num, state in enumerate(self._saved_page_states, start=1):
            self.__dict__.update(state)
            self._draw_header_footer(page_num, total_pages)
            canvas.Canvas.showPage(self)  # write page exactly once

        canvas.Canvas.save(self)

    def _draw_header_footer(self, page_num: int, total_pages: int):
        header_footer_text = f"{self.report_name}   page {page_num} of {total_pages}"

        page_w, page_h = LETTER
        left = 0.75 * inch
        right = 0.75 * inch

        self.setFont("Helvetica", 9)
        self.setFillColor(colors.black)

        # Place text safely away from rules (no overlap).
        # Header text
        self.drawString(left, page_h - 0.55 * inch, header_footer_text)
        # Footer text (moved slightly lower)
        self.drawString(left, 0.45 * inch, header_footer_text)

        # Subtle rules moved farther from text
        self.setLineWidth(0.4)
        self.setStrokeColor(colors.grey)
        # Header rule (higher than body, below header text)
        self.line(left, page_h - 0.78 * inch, page_w - right, page_h - 0.78 * inch)
        # Footer rule (above footer text and below body)
        self.line(left, 0.75 * inch, page_w - right, 0.75 * inch)


def get_identity_block() -> list[tuple[str, str]]:
    user = getpass.getuser()
    host = socket.gethostname()
    os_desc = platform.platform()
    py_desc = platform.python_version()
    cwd = os.getcwd()

    domain = os.environ.get("USERDOMAIN", "")
    computername = os.environ.get("COMPUTERNAME", "")

    items = [
        ("Username", user),
        ("Hostname", host),
        ("ComputerName", computername if computername else host),
        ("UserDomain", domain if domain else "(N/A)"),
        ("OS", os_desc),
        ("Python", py_desc),
        ("Working directory", cwd),
        ("Generated", stamp_human()),
    ]
    return items


def generate_pdf_report(
    pdf_path: str,
    report_name: str,
    run_id: str,
    *,
    file1_path: str,
    file2_path: str,
    sheet1: str | None,
    sheet2: str | None,
    col1: str,
    col2: str,
    trim_ws: bool,
    case_insensitive: bool,
    set1_count: int,
    set2_count: int,
    both_count: int,
    only1: list[str],
    only2: list[str],
    map1: dict[str, list[int]],
    map2: dict[str, list[int]],
) -> None:
    styles = getSampleStyleSheet()

    body = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=13,
        spaceAfter=6,
    )
    small = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
    )
    h1 = ParagraphStyle(
        "H1",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=18,
        spaceAfter=10,
    )
    h2 = ParagraphStyle(
        "H2",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        spaceBefore=10,
        spaceAfter=6,
    )

    # Slightly larger margins to guarantee clear separation from header/footer content.
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=LETTER,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=1.05 * inch,     # more room for header text + rule
        bottomMargin=1.05 * inch,  # more room for footer text + rule
        title=report_name,
    )

    story = []

    f1_name = safe_basename(file1_path)
    f2_name = safe_basename(file2_path)

    # Title
    story.append(Paragraph("Column Mismatch Analysis", h1))

    # Disclaimer / context (with explicit file names)
    disclaimer_text = (
        "This report compares values between two source datasets after optional normalization "
        "(trim whitespace and/or case-insensitive matching). "
        f"File 1 [{f1_name}] is compared against File 2 [{f2_name}]. "
        "Results identify values present in only one dataset, including occurrence counts and row references "
        "(row numbers are Excel-like: header row = 1; first data row = 2)."
    )
    story.append(Paragraph(disclaimer_text, body))
    story.append(Spacer(1, 8))

    # Run meta
    story.append(Paragraph("Run Metadata", h2))
    meta_lines = [
        f"Run ID: {run_id}",
        f"File 1 [{f1_name}]: {file1_path}",
        f"File 1 worksheet: {sheet1 if sheet1 else '(N/A)'}",
        f"File 1 column: {col1}",
        f"File 2 [{f2_name}]: {file2_path}",
        f"File 2 worksheet: {sheet2 if sheet2 else '(N/A)'}",
        f"File 2 column: {col2}",
        f"Normalization: trim_whitespace={trim_ws}, case_insensitive={case_insensitive}",
    ]
    for line in meta_lines:
        story.append(Paragraph(line, small))
    story.append(Spacer(1, 10))

    # Identity logs
    story.append(Paragraph("Identity Logs", h2))
    ident = get_identity_block()
    ident_table = Table(
        [["Field", "Value"]] + [[k, v] for k, v in ident],
        colWidths=[1.6 * inch, 5.6 * inch],
        repeatRows=1,
        hAlign="LEFT",
    )
    ident_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(ident_table)
    story.append(Spacer(1, 12))

    # Summary
    story.append(Paragraph("Summary", h2))
    summary_table = Table(
        [
            ["Metric", "Value"],
            ["Distinct values in File 1", str(set1_count)],
            ["Distinct values in File 2", str(set2_count)],
            ["Values in both", str(both_count)],
            ["Only in File 1", str(len(only1))],
            ["Only in File 2", str(len(only2))],
        ],
        colWidths=[3.0 * inch, 4.2 * inch],
        repeatRows=1,
        hAlign="LEFT",
    )
    summary_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(summary_table)
    story.append(Spacer(1, 12))

    # Detail tables
    def build_detail_table(title: str, values: list[str], which: int) -> list:
        """
        which=1 => only in file1 (use map1)
        which=2 => only in file2 (use map2)
        """
        story_bits = [Paragraph(title, h2)]

        if not values:
            story_bits.append(Paragraph("None.", body))
            story_bits.append(Spacer(1, 8))
            return story_bits

        rows = [["Value", "Occurrences", "Row numbers (Excel-like)"]]
        for v in values:
            rr = map1.get(v, []) if which == 1 else map2.get(v, [])
            rows.append(
                [
                    Paragraph(v, small),
                    Paragraph(str(len(rr)), small),
                    Paragraph(format_rows(rr), small),
                ]
            )

        t = Table(
            rows,
            colWidths=[1.6 * inch, 1.0 * inch, 4.6 * inch],
            repeatRows=1,
            hAlign="LEFT",
        )
        t.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                    ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )

        story_bits.append(t)
        story_bits.append(Spacer(1, 12))
        return story_bits

    # Force dedicated pages for each detail section
    story.append(PageBreak())
    story.extend(build_detail_table("Detail — Values ONLY in File 1 (missing from File 2)", only1, which=1))

    story.append(PageBreak())
    story.extend(build_detail_table("Detail — Values ONLY in File 2 (missing from File 1)", only2, which=2))

    # Build with numbered header/footer (fixed canvas)
    doc.build(
        story,
        canvasmaker=lambda *args, **kwargs: NumberedCanvas(*args, report_name=report_name, **kwargs),
    )


# -----------------------------
# GUI
# -----------------------------
class CompareGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Spreadsheet Column Mismatch Analyzer")
        self.geometry("1120x760")
        self.minsize(980, 660)

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_dir = tk.StringVar()

        self.sheet1 = tk.StringVar()
        self.sheet2 = tk.StringVar()

        self.col1 = tk.StringVar()
        self.col2 = tk.StringVar()

        self.trim_ws = tk.BooleanVar(value=True)
        self.case_insensitive = tk.BooleanVar(value=True)
        self.make_pdf = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)

        files_frame = ttk.LabelFrame(outer, text="Inputs", padding=10)
        files_frame.pack(fill="x")

        ttk.Label(files_frame, text="Spreadsheet 1:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(files_frame, textvariable=self.file1_path, width=92).grid(row=0, column=1, sticky="we", pady=4)
        ttk.Button(files_frame, text="Browse...", command=lambda: self._browse_file(1)).grid(
            row=0, column=2, padx=(8, 0), pady=4
        )

        ttk.Label(files_frame, text="Spreadsheet 2:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(files_frame, textvariable=self.file2_path, width=92).grid(row=1, column=1, sticky="we", pady=4)
        ttk.Button(files_frame, text="Browse...", command=lambda: self._browse_file(2)).grid(
            row=1, column=2, padx=(8, 0), pady=4
        )

        ttk.Label(files_frame, text="Output folder:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(files_frame, textvariable=self.output_dir, width=92).grid(row=2, column=1, sticky="we", pady=4)
        ttk.Button(files_frame, text="Pick folder...", command=self._pick_output_dir).grid(
            row=2, column=2, padx=(8, 0), pady=4
        )

        files_frame.columnconfigure(1, weight=1)

        select_frame = ttk.LabelFrame(outer, text="Selection", padding=10)
        select_frame.pack(fill="x", pady=(10, 0))

        ttk.Label(select_frame, text="Spreadsheet 1").grid(row=0, column=1, sticky="w")
        ttk.Label(select_frame, text="Spreadsheet 2").grid(row=0, column=2, sticky="w")

        ttk.Label(select_frame, text="Worksheet (Excel only):").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.sheet1_cb = ttk.Combobox(select_frame, textvariable=self.sheet1, state="readonly", width=45)
        self.sheet2_cb = ttk.Combobox(select_frame, textvariable=self.sheet2, state="readonly", width=45)
        self.sheet1_cb.grid(row=1, column=1, sticky="we", pady=4)
        self.sheet2_cb.grid(row=1, column=2, sticky="we", pady=4)
        self.sheet1_cb.bind("<<ComboboxSelected>>", lambda _e: self._load_columns(1))
        self.sheet2_cb.bind("<<ComboboxSelected>>", lambda _e: self._load_columns(2))

        ttk.Label(select_frame, text="Column to compare:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self.col1_cb = ttk.Combobox(select_frame, textvariable=self.col1, state="readonly", width=45)
        self.col2_cb = ttk.Combobox(select_frame, textvariable=self.col2, state="readonly", width=45)
        self.col1_cb.grid(row=2, column=1, sticky="we", pady=4)
        self.col2_cb.grid(row=2, column=2, sticky="we", pady=4)

        select_frame.columnconfigure(1, weight=1)
        select_frame.columnconfigure(2, weight=1)

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(10, 0))

        ttk.Checkbutton(actions, text="Trim whitespace", variable=self.trim_ws).pack(side="left")
        ttk.Checkbutton(actions, text="Case-insensitive match", variable=self.case_insensitive).pack(
            side="left", padx=(12, 0)
        )
        ttk.Checkbutton(actions, text="Generate PDF report", variable=self.make_pdf).pack(side="left", padx=(12, 0))

        self.analyze_btn = ttk.Button(actions, text="Analyze mismatches", command=self._on_analyze)
        self.analyze_btn.pack(side="right")
        ttk.Button(actions, text="Clear output", command=self._clear_output).pack(side="right", padx=(0, 10))

        out_frame = ttk.LabelFrame(outer, text="Output", padding=10)
        out_frame.pack(fill="both", expand=True, pady=(10, 0))
        self.output = ScrolledText(out_frame, wrap="word", height=22)
        self.output.pack(fill="both", expand=True)

        self._log("Ready. Select both spreadsheets, choose columns, then run Analyze mismatches.")

    def _log(self, msg: str) -> None:
        self.output.insert("end", msg + "\n")
        self.output.see("end")

    def _clear_output(self) -> None:
        self.output.delete("1.0", "end")

    def _pick_output_dir(self) -> None:
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self.output_dir.set(d)
            self._log(f"[{stamp_human()}] Output folder set to: {d}")

    def _browse_file(self, which: int) -> None:
        path = filedialog.askopenfilename(
            title=f"Select Spreadsheet {which}",
            filetypes=[
                ("Spreadsheets", "*.xlsx *.xlsm *.xls *.csv"),
                ("Excel", "*.xlsx *.xlsm *.xls"),
                ("CSV", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        if os.path.splitext(path)[1].lower() not in SUPPORTED_EXTS:
            messagebox.showerror("Unsupported file", f"Unsupported file type.\nSupported: {', '.join(SUPPORTED_EXTS)}")
            return

        if which == 1:
            self.file1_path.set(path)
        else:
            self.file2_path.set(path)

        try:
            self._load_sheets_and_columns(which)
        except Exception as e:
            messagebox.showerror("Load error", f"Could not read file metadata:\n{e}")

    def _load_sheets_and_columns(self, which: int) -> None:
        path = self.file1_path.get() if which == 1 else self.file2_path.get()
        if not path:
            return

        if is_excel(path):
            sheets = get_excel_sheet_names(path)
            if which == 1:
                self.sheet1_cb["values"] = sheets
                self.sheet1.set(sheets[0] if sheets else "")
            else:
                self.sheet2_cb["values"] = sheets
                self.sheet2.set(sheets[0] if sheets else "")
        else:
            if which == 1:
                self.sheet1_cb["values"] = ["(CSV)"]
                self.sheet1.set("(CSV)")
            else:
                self.sheet2_cb["values"] = ["(CSV)"]
                self.sheet2.set("(CSV)")

        self._load_columns(which)

    def _load_columns(self, which: int) -> None:
        path = self.file1_path.get() if which == 1 else self.file2_path.get()
        if not path:
            return

        sheet_name: str | None
        if is_excel(path):
            sheet_name = self.sheet1.get() if which == 1 else self.sheet2.get()
            sheet_name = sheet_name.strip() if sheet_name else None
        else:
            sheet_name = None

        cols = [str(c) for c in read_headers(path, sheet_name=sheet_name)]

        if which == 1:
            self.col1_cb["values"] = cols
            self.col1.set("Legacy Document #" if "Legacy Document #" in cols else (cols[0] if cols else ""))
        else:
            self.col2_cb["values"] = cols
            self.col2.set("Legacy Document #" if "Legacy Document #" in cols else (cols[0] if cols else ""))

        self._log(f"[{stamp_human()}] Loaded columns for Spreadsheet {which}: {safe_basename(path)}")

    def _on_analyze(self) -> None:
        p1 = self.file1_path.get().strip()
        p2 = self.file2_path.get().strip()
        out_dir = self.output_dir.get().strip()

        if not p1 or not os.path.exists(p1):
            messagebox.showerror("Missing input", "Please select Spreadsheet 1.")
            return
        if not p2 or not os.path.exists(p2):
            messagebox.showerror("Missing input", "Please select Spreadsheet 2.")
            return
        if self.make_pdf.get():
            if not out_dir:
                messagebox.showerror("Missing output folder", "Please select an output folder (for the PDF report).")
                return
            if not os.path.isdir(out_dir):
                messagebox.showerror("Invalid output folder", "Selected output folder does not exist.")
                return

        c1 = self.col1.get().strip()
        c2 = self.col2.get().strip()
        if not c1:
            messagebox.showerror("Missing selection", "Please select a column for Spreadsheet 1.")
            return
        if not c2:
            messagebox.showerror("Missing selection", "Please select a column for Spreadsheet 2.")
            return

        s1 = None if not is_excel(p1) else (self.sheet1.get().strip() or None)
        s2 = None if not is_excel(p2) else (self.sheet2.get().strip() or None)

        run_id = stamp_ddmmyyyy_hhmmss()
        report_name = f"Column_Mismatches_{run_id}_report"

        self.analyze_btn.configure(state="disabled")
        self._log("\n" + "=" * 120)
        self._log(f"[{stamp_human()}] Starting mismatch analysis. Run ID: {run_id}")
        self._log(f"File 1 [{safe_basename(p1)}]: {p1}")
        self._log(f"  Sheet: {s1 if s1 else '(N/A)'}")
        self._log(f"  Column: {c1}")
        self._log(f"File 2 [{safe_basename(p2)}]: {p2}")
        self._log(f"  Sheet: {s2 if s2 else '(N/A)'}")
        self._log(f"  Column: {c2}")
        self._log(f"Normalization: trim_whitespace={self.trim_ws.get()}, case_insensitive={self.case_insensitive.get()}")
        if self.make_pdf.get():
            self._log(f"PDF output folder: {out_dir}")
        self._log("-" * 120)

        t = threading.Thread(
            target=self._run_analysis,
            args=(report_name, run_id, p1, s1, c1, p2, s2, c2, out_dir),
            daemon=True,
        )
        t.start()

    def _run_analysis(
        self,
        report_name: str,
        run_id: str,
        p1: str,
        s1: str | None,
        c1: str,
        p2: str,
        s2: str | None,
        c2: str,
        out_dir: str,
    ) -> None:
        try:
            ser1_raw = read_single_column(p1, c1, sheet_name=s1)
            ser2_raw = read_single_column(p2, c2, sheet_name=s2)

            ser1 = normalize_series(ser1_raw, trim=self.trim_ws.get(), case_insensitive=self.case_insensitive.get())
            ser2 = normalize_series(ser2_raw, trim=self.trim_ws.get(), case_insensitive=self.case_insensitive.get())

            map1 = build_row_index_map(ser1)
            map2 = build_row_index_map(ser2)

            set1 = set(map1.keys())
            set2 = set(map2.keys())

            only1 = sorted(set1 - set2)
            only2 = sorted(set2 - set1)
            both = sorted(set1 & set2)

            lines: list[str] = []
            lines.append(f"[{stamp_human()}] Analysis complete. Run ID: {run_id}")
            lines.append("")
            lines.append("SUMMARY")
            lines.append(f"  Distinct values in File 1: {len(set1)}")
            lines.append(f"  Distinct values in File 2: {len(set2)}")
            lines.append(f"  In BOTH: {len(both)}")
            lines.append(f"  ONLY in File 1 (missing from File 2): {len(only1)}")
            lines.append(f"  ONLY in File 2 (missing from File 1): {len(only2)}")
            lines.append("")

            lines.append("DETAIL — Values ONLY in File 1")
            if only1:
                for v in only1:
                    rr = map1.get(v, [])
                    lines.append(f"  {v} | {len(rr)} | {format_rows(rr)}")
            else:
                lines.append("  None.")

            lines.append("")
            lines.append("DETAIL — Values ONLY in File 2")
            if only2:
                for v in only2:
                    rr = map2.get(v, [])
                    lines.append(f"  {v} | {len(rr)} | {format_rows(rr)}")
            else:
                lines.append("  None.")

            pdf_path = None
            if self.make_pdf.get():
                filename = f"{report_name}.pdf"
                pdf_path = os.path.join(out_dir, filename)

                generate_pdf_report(
                    pdf_path,
                    report_name,
                    run_id,
                    file1_path=p1,
                    file2_path=p2,
                    sheet1=s1,
                    sheet2=s2,
                    col1=c1,
                    col2=c2,
                    trim_ws=self.trim_ws.get(),
                    case_insensitive=self.case_insensitive.get(),
                    set1_count=len(set1),
                    set2_count=len(set2),
                    both_count=len(both),
                    only1=only1,
                    only2=only2,
                    map1=map1,
                    map2=map2,
                )

            if pdf_path:
                lines.append("")
                lines.append("PDF REPORT WRITTEN:")
                lines.append(f"  {pdf_path}")

            lines.append("")
            lines.append("=" * 120)

            self.after(0, lambda: self._log("\n".join(lines)))

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Analysis error", f"Mismatch analysis failed:\n{e}"))
        finally:
            self.after(0, lambda: self.analyze_btn.configure(state="normal"))


if __name__ == "__main__":
    # Windows DPI tweak (optional)
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = CompareGUI()
    app.mainloop()
