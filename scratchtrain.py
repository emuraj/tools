#!/usr/bin/env python3
# normalize_training_gui_min.py
# Zero-fuss GUI to normalize training assignments.
# - Input: CSV (always supported). Excel also supported if pandas is available.
# - Output: CSV always; Excel (.xlsx) if pandas+openpyxl present.
#
# Mapping:
#   User Name       <- Trainee Name
#   Training Plan   <- extract TP-#### from Training Plan Name
#   Document Number <- Library Item Number
#   Document Title  <- Library Item Name

import csv
import re
from pathlib import Path
from typing import Optional, List, Dict

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional imports
HAVE_PANDAS = False
HAVE_OPENPYXL = False
try:
    import pandas as pd  # type: ignore
    HAVE_PANDAS = True
    try:
        import openpyxl  # noqa: F401
        HAVE_OPENPYXL = True
    except Exception:
        HAVE_OPENPYXL = False
except Exception:
    HAVE_PANDAS = False
    HAVE_OPENPYXL = False

REQUIRED_MIN_COLS = [
    "trainee name",
    "training plan name",
    "library item number",
    "library item name",
]

def extract_tp(plan: Optional[str]) -> str:
    s = "" if plan is None else str(plan)
    m = re.search(r"\bTP-\d{4}\b", s, flags=re.IGNORECASE)
    return m.group(0).upper() if m else ""

def normalize_rows(rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    out = []
    for r in rows:
        # case-insensitive access
        lc = {k.lower().strip(): (v if v is not None else "") for k, v in r.items()}
        def g(name: str) -> str:
            return str(lc.get(name.lower(), "")).strip()

        out.append({
            "User Name":       g("Trainee Name"),
            "Training Plan":   extract_tp(g("Training Plan Name")),
            "Document Number": g("Library Item Number"),
            "Document Title":  g("Library Item Name"),
        })
    return out

def infer_has_required(headers: List[str]) -> bool:
    hset = {h.lower().strip() for h in headers}
    return all(c in hset for c in REQUIRED_MIN_COLS)

def read_csv_table(path: Path) -> List[Dict[str, str]]:
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        if not infer_has_required(headers):
            raise ValueError(
                "CSV is missing required columns. Need at least: "
                "Trainee Name, Training Plan Name, Library Item Number, Library Item Name"
            )
        return [dict(row) for row in reader]

def read_excel_table(path: Path, sheet_name: Optional[str]) -> List[Dict[str, str]]:
    if not HAVE_PANDAS:
        raise RuntimeError(
            "Excel reading requires pandas. Export your file to CSV or install pandas."
        )
    df = None
    if sheet_name:
        df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)
    else:
        # try to find a sheet with the needed headers
        xls = pd.ExcelFile(path)
        for s in xls.sheet_names:
            cand = pd.read_excel(path, sheet_name=s, dtype=str)
            cand.columns = [str(c).strip() for c in cand.columns]
            if infer_has_required(list(cand.columns)):
                df = cand
                break
        if df is None:
            # fallback: first sheet
            df = pd.read_excel(path, sheet_name=xls.sheet_names[0], dtype=str)

    df.columns = [str(c).strip() for c in df.columns]
    if not infer_has_required(list(df.columns)):
        raise ValueError(
            "Excel sheet is missing required columns. Need at least: "
            "Trainee Name, Training Plan Name, Library Item Number, Library Item Name"
        )
    # Convert to list-of-dicts
    return [{c: ("" if pd.isna(v) else str(v)) for c, v in row.items()} for _, row in df.iterrows()]

def write_csv(path: Path, rows: List[Dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["User Name","Training Plan","Document Number","Document Title"])
        writer.writeheader()
        writer.writerows(rows)

def write_excel(path: Path, rows: List[Dict[str, str]]) -> None:
    if not (HAVE_PANDAS and HAVE_OPENPYXL):
        # Fallback to CSV if Excel writer is not available
        csv_path = path.with_suffix(".csv")
        write_csv(csv_path, rows)
        raise RuntimeError(f"openpyxl/pandas not available to write .xlsx. Wrote CSV instead:\n{csv_path}")
    df = pd.DataFrame(rows, columns=["User Name","Training Plan","Document Number","Document Title"])
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(path, index=False)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Normalize Training Assignments")
        self.geometry("760x240")
        self.minsize(760, 240)

        self.input_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.sheet_name = tk.StringVar()  # optional (Excel)

        pad = {"padx": 10, "pady": 6}
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, **pad)

        ttk.Label(frm, text="File (CSV or Excel):").grid(row=0, column=0, sticky="e")
        ttk.Entry(frm, textvariable=self.input_path, width=70).grid(row=0, column=1, sticky="we")
        ttk.Button(frm, text="Browse…", command=self.browse_input).grid(row=0, column=2, sticky="w", padx=5)

        ttk.Label(frm, text="Output location (folder):").grid(row=1, column=0, sticky="e")
        ttk.Entry(frm, textvariable=self.output_dir, width=70).grid(row=1, column=1, sticky="we")
        ttk.Button(frm, text="Browse…", command=self.browse_output).grid(row=1, column=2, sticky="w", padx=5)

        ttk.Label(frm, text="Excel sheet (optional):").grid(row=2, column=0, sticky="e")
        ttk.Entry(frm, textvariable=self.sheet_name, width=30).grid(row=2, column=1, sticky="w")

        self.btn_start = ttk.Button(frm, text="Start", command=self.start)
        self.btn_start.grid(row=3, column=1, sticky="w", pady=(10, 0))

        self.status = tk.StringVar(value="Idle")
        ttk.Label(frm, textvariable=self.status).grid(row=4, column=1, sticky="w", pady=(8, 0))

        frm.columnconfigure(1, weight=1)

    def browse_input(self):
        path = filedialog.askopenfilename(
            title="Select CSV/Excel file",
            filetypes=[
                ("CSV", "*.csv"),
                ("Excel", "*.xlsx *.xlsm *.xls"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.input_path.set(path)

    def browse_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_dir.set(path)

    def start(self):
        in_path = Path(self.input_path.get().strip()) if self.input_path.get().strip() else None
        out_dir = Path(self.output_dir.get().strip()) if self.output_dir.get().strip() else None
        sheet = self.sheet_name.get().strip() or None

        if not in_path:
            messagebox.showerror("Missing file", "Please choose an input CSV/Excel file.")
            return
        if not in_path.exists():
            messagebox.showerror("File not found", f"Input not found:\n{in_path}")
            return
        if not out_dir:
            messagebox.showerror("Missing output location", "Please choose an output folder.")
            return
        if not out_dir.exists():
            messagebox.showerror("Folder not found", f"Output folder does not exist:\n{out_dir}")
            return

        self.btn_start.config(state="disabled")
        self.status.set("Running…")

        try:
            ext = in_path.suffix.lower()
            if ext == ".csv":
                raw_rows = read_csv_table(in_path)
            elif ext in (".xlsx", ".xlsm", ".xls"):
                raw_rows = read_excel_table(in_path, sheet)
            else:
                raise ValueError("Unsupported file type. Use CSV or Excel.")

            norm_rows = normalize_rows(raw_rows)

            # Try to write Excel first if possible; else CSV.
            out_xlsx = out_dir / f"{in_path.stem}_normalized.xlsx"
            try:
                write_excel(out_xlsx, norm_rows)
                self.status.set(f"Done → {out_xlsx}")
                messagebox.showinfo("Success", f"Saved Excel:\n{out_xlsx}")
            except Exception as e:
                # Fallback already wrote CSV
                self.status.set("Completed with fallback")
                messagebox.showwarning("Saved CSV", str(e))

        except Exception as e:
            self.status.set("Error")
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_start.config(state="normal")

if __name__ == "__main__":
    App().mainloop()
