#!/usr/bin/env python3
from __future__ import annotations

import os
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from datetime import datetime

import pandas as pd


COL1_HEADER = "Last name, First name"
COL2_HEADER = "Employee"


def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalize_name(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s.casefold()


def find_sheet_with_header(filepath: str, header_name: str) -> str:
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    try:
        xl = pd.ExcelFile(filepath, engine="openpyxl")
    except Exception as e:
        raise RuntimeError(f"Could not open Excel file: {filepath}\n{e}")

    for sheet in xl.sheet_names:
        try:
            df = pd.read_excel(filepath, sheet_name=sheet, engine="openpyxl", nrows=5)
        except Exception:
            continue

        if header_name in df.columns:
            return sheet

    detail_lines = []
    for sheet in xl.sheet_names:
        try:
            df = pd.read_excel(filepath, sheet_name=sheet, engine="openpyxl", nrows=1)
            headers = [str(c) for c in df.columns.tolist()]
            detail_lines.append(f"  - {sheet}: {headers}")
        except Exception:
            detail_lines.append(f"  - {sheet}: (could not read headers)")

    details = "\n".join(detail_lines)
    raise KeyError(
        f"Column header {header_name!r} was not found in ANY sheet in:\n{filepath}\n\n"
        f"Headers seen per sheet (best effort):\n{details}"
    )


def load_name_set_autosheet(filepath: str, header_name: str) -> tuple[str, dict[str, str], list[str]]:
    sheet_used = find_sheet_with_header(filepath, header_name)

    df = pd.read_excel(filepath, sheet_name=sheet_used, engine="openpyxl")
    if header_name not in df.columns:
        raise KeyError(f"Expected header {header_name!r} not found in detected sheet {sheet_used!r}.")

    series = df[header_name]

    norm_to_original: dict[str, str] = {}
    duplicates: list[str] = []

    for v in series.tolist():
        if pd.isna(v):
            continue
        original = str(v).strip()
        if not original:
            continue
        norm = normalize_name(original)
        if not norm:
            continue
        if norm in norm_to_original:
            duplicates.append(original)
        else:
            norm_to_original[norm] = original

    return sheet_used, norm_to_original, duplicates


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Name List Comparator (Auto Sheet Detect)")
        self.geometry("980x650")
        self.minsize(860, 560)

        self.file1_var = tk.StringVar()
        self.file2_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        info = (
            "Compares two Excel files (sheet names do not matter).\n"
            f"  - File 1: finds a sheet containing header {COL1_HEADER!r}\n"
            f"  - File 2: finds a sheet containing header {COL2_HEADER!r}\n"
            "Then lists names present in one file but not the other (case-insensitive, whitespace-normalized)."
        )
        ttk.Label(root, text=info, justify="left").pack(fill="x", pady=(0, 10))

        file_frame = ttk.LabelFrame(root, text="Inputs", padding=10)
        file_frame.pack(fill="x")

        row1 = ttk.Frame(file_frame)
        row1.pack(fill="x", pady=4)
        ttk.Label(row1, text="Excel File 1:").pack(side="left")
        ttk.Entry(row1, textvariable=self.file1_var).pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(row1, text="Browse...", command=self.browse_file1).pack(side="left")

        row2 = ttk.Frame(file_frame)
        row2.pack(fill="x", pady=4)
        ttk.Label(row2, text="Excel File 2:").pack(side="left")
        ttk.Entry(row2, textvariable=self.file2_var).pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(row2, text="Browse...", command=self.browse_file2).pack(side="left")

        ctrl = ttk.Frame(root)
        ctrl.pack(fill="x", pady=(10, 8))

        self.run_btn = ttk.Button(ctrl, text="Run Compare", command=self.on_run)
        self.run_btn.pack(side="left")

        self.clear_btn = ttk.Button(ctrl, text="Clear Log", command=self.clear_log)
        self.clear_btn.pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(ctrl, textvariable=self.status_var).pack(side="right")

        log_frame = ttk.LabelFrame(root, text="Event Log", padding=10)
        log_frame.pack(fill="both", expand=True)

        self.log = ScrolledText(log_frame, height=18, wrap="word")
        self.log.pack(fill="both", expand=True)

        self.log.tag_configure("ok", foreground="green")
        self.log.tag_configure("warn", foreground="darkorange")
        self.log.tag_configure("err", foreground="red")
        self.log.tag_configure("hdr", font=("TkDefaultFont", 10, "bold"))

    def log_line(self, msg: str, tag: str | None = None) -> None:
        line = f"[{now_stamp()}] {msg}\n"
        self.log.insert("end", line, tag or ())
        self.log.see("end")
        try:
            print(line, end="")
        except Exception:
            pass

    def set_status(self, msg: str) -> None:
        self.status_var.set(msg)
        self.update_idletasks()

    def clear_log(self) -> None:
        self.log.delete("1.0", "end")
        self.set_status("Ready.")

    def browse_file1(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Excel File 1",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            self.file1_var.set(path)

    def browse_file2(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Excel File 2",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            self.file2_var.set(path)

    def on_run(self) -> None:
        f1 = self.file1_var.get().strip()
        f2 = self.file2_var.get().strip()

        if not f1 or not f2:
            messagebox.showerror("Missing input", "Please select both Excel files.")
            return

        self.run_btn.configure(state="disabled")
        self.clear_btn.configure(state="disabled")
        self.set_status("Running...")
        self.log_line("Starting comparison...", "hdr")
        self.log_line(f"File 1: {f1}")
        self.log_line(f"File 2: {f2}")

        threading.Thread(target=self._run_compare_thread, args=(f1, f2), daemon=True).start()

    def _run_compare_thread(self, f1: str, f2: str) -> None:
        try:
            self._compare(f1, f2)
            self.after(0, lambda: self.set_status("Done."))
        except Exception as e:
            self.after(0, lambda: self.set_status("Error."))
            self.after(0, lambda: self.log_line(f"ERROR: {e}", "err"))
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal"))
            self.after(0, lambda: self.clear_btn.configure(state="normal"))

    def _compare(self, f1: str, f2: str) -> None:
        base1 = os.path.basename(f1)
        base2 = os.path.basename(f2)

        self.after(0, lambda: self.log_line("Detecting sheet + loading names from File 1...", "ok"))
        sheet1, map1, dups1 = load_name_set_autosheet(f1, COL1_HEADER)
        self.after(0, lambda: self.log_line(f"File 1: using sheet {sheet1!r}", "ok"))
        self.after(0, lambda: self.log_line(f"File 1 loaded: {len(map1)} unique normalized names.", "ok"))
        if dups1:
            self.after(0, lambda: self.log_line(
                f"File 1 note: {len(dups1)} duplicates after normalization (kept first occurrence).",
                "warn"
            ))

        self.after(0, lambda: self.log_line("Detecting sheet + loading names from File 2...", "ok"))
        sheet2, map2, dups2 = load_name_set_autosheet(f2, COL2_HEADER)
        self.after(0, lambda: self.log_line(f"File 2: using sheet {sheet2!r}", "ok"))
        self.after(0, lambda: self.log_line(f"File 2 loaded: {len(map2)} unique normalized names.", "ok"))
        if dups2:
            self.after(0, lambda: self.log_line(
                f"File 2 note: {len(dups2)} duplicates after normalization (kept first occurrence).",
                "warn"
            ))

        set1 = set(map1.keys())
        set2 = set(map2.keys())

        only_in_1 = sorted(set1 - set2)
        only_in_2 = sorted(set2 - set1)
        in_both = len(set1 & set2)

        self.after(0, lambda: self.log_line("", None))
        self.after(0, lambda: self.log_line("RESULTS", "hdr"))
        self.after(0, lambda: self.log_line(f"Overlap (present in both): {in_both}", "ok"))
        self.after(0, lambda: self.log_line(
            f"Only in {base1}: {len(only_in_1)}", "warn" if only_in_1 else "ok"
        ))
        self.after(0, lambda: self.log_line(
            f"Only in {base2}: {len(only_in_2)}", "warn" if only_in_2 else "ok"
        ))

        self.after(0, lambda: self.log_line("", None))
        self.after(0, lambda: self.log_line(f"Detail — Names ONLY in {base1}", "hdr"))
        if only_in_1:
            for k in only_in_1:
                self.after(0, lambda v=map1[k]: self.log_line(f"  {v}"))
        else:
            self.after(0, lambda: self.log_line("  (none)", "ok"))

        self.after(0, lambda: self.log_line("", None))
        self.after(0, lambda: self.log_line(f"Detail — Names ONLY in {base2}", "hdr"))
        if only_in_2:
            for k in only_in_2:
                self.after(0, lambda v=map2[k]: self.log_line(f"  {v}"))
        else:
            self.after(0, lambda: self.log_line("  (none)", "ok"))

        self.after(0, lambda: self.log_line("", None))
        self.after(0, lambda: self.log_line("Comparison complete.", "ok"))


def main() -> int:
    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
