import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# ----------------------------
# Parsing helpers
# ----------------------------
ANALYST_RE = re.compile(r'^([A-Za-z]+)(?=[_-])')
V_RE = re.compile(r'(?:^|[_-])V(\d+)(?=[_-]|$)', re.IGNORECASE)
D_RE = re.compile(r'(?:^|[_-])D(\d+)(?=[_-]|$)', re.IGNORECASE)

def parse_sample_id(s: str):
    """
    Returns: (analyst, v_num, d_num, type_token)

    Rules:
      - Analyst: letters before first '_' or '-'
      - V: digits after token V (e.g., V28 -> 28)
      - D: digits after token D (e.g., D25 -> 25)
      - Type normalized:
          * trailing token '2' -> 'BA2'
          * trailing token '1' -> 'BA1'
          * trailing token 'Bulk'/'bulk' -> 'Bulk'
          * trailing token starting with BA1/BA2 (e.g., BA2e5) -> BA1/BA2
          * otherwise blank
    """
    if pd.isna(s):
        return ("", "", "", "")

    text = str(s).strip()
    if not text:
        return ("", "", "", "")

    # Analyst
    m = ANALYST_RE.search(text)
    analyst = m.group(1).upper() if m else ""

    # V and D
    mv = V_RE.search(text)
    md = D_RE.search(text)
    v_num = mv.group(1) if mv else ""
    d_num = md.group(1) if md else ""

    # Type (trailing token), normalized
    parts = re.split(r'[_-]+', text)
    tail = parts[-1].strip() if parts else ""

    type_token = ""
    if tail == "2":
        type_token = "BA2"
    elif tail == "1":
        type_token = "BA1"
    elif tail.lower() == "bulk":
        type_token = "Bulk"
    else:
        # BA1 / BA2 / BA2e5 etc.
        mba = re.match(r'^(BA[12])', tail, flags=re.IGNORECASE)
        if mba:
            type_token = mba.group(1).upper()

    return (analyst, v_num, d_num, type_token)

# ----------------------------
# GUI app
# ----------------------------
class SampleIdSplitterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sample ID Splitter (Excel)")
        self.geometry("760x420")

        self.in_path = tk.StringVar()
        self.col_name = tk.StringVar(value="Sample ID")
        self.out_path = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm = tk.Frame(self)
        frm.pack(fill="x", **pad)

        tk.Label(frm, text="Input Excel (.xlsx):", width=18, anchor="w").grid(row=0, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.in_path).grid(row=0, column=1, sticky="we")
        tk.Button(frm, text="Browse...", command=self.browse_in).grid(row=0, column=2, padx=6)

        tk.Label(frm, text="Sample ID column:", width=18, anchor="w").grid(row=1, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.col_name).grid(row=1, column=1, sticky="we")

        tk.Label(frm, text="Output Excel (.xlsx):", width=18, anchor="w").grid(row=2, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.out_path).grid(row=2, column=1, sticky="we")
        tk.Button(frm, text="Save As...", command=self.browse_out).grid(row=2, column=2, padx=6)

        frm.grid_columnconfigure(1, weight=1)

        btns = tk.Frame(self)
        btns.pack(fill="x", **pad)
        tk.Button(btns, text="Run", width=12, command=self.run).pack(side="left")
        tk.Button(btns, text="Quit", width=12, command=self.destroy).pack(side="left", padx=8)

        tk.Label(self, text="Log:").pack(anchor="w", padx=10)
        self.log = tk.Text(self, height=14)
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def log_line(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.update_idletasks()

    def browse_in(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
        )
        if path:
            self.in_path.set(path)
            # default output next to input
            base, _ext = os.path.splitext(path)
            self.out_path.set(base + "_parsed.xlsx")

    def browse_out(self):
        path = filedialog.asksaveasfilename(
            title="Save output as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.out_path.set(path)

    def run(self):
        in_path = self.in_path.get().strip()
        out_path = self.out_path.get().strip()
        col = self.col_name.get().strip()

        if not in_path or not os.path.exists(in_path):
            messagebox.showerror("Error", "Please select a valid input Excel file.")
            return
        if not out_path:
            messagebox.showerror("Error", "Please choose an output file path.")
            return
        if not col:
            messagebox.showerror("Error", "Please specify the Sample ID column name.")
            return

        try:
            self.log_line(f"Reading: {in_path}")
            df = pd.read_excel(in_path, sheet_name=0)  # first sheet
            self.log_line(f"Rows: {len(df):,} | Columns: {len(df.columns)}")

            if col not in df.columns:
                messagebox.showerror(
                    "Error",
                    f"Column '{col}' not found.\n\nAvailable columns:\n" + "\n".join(map(str, df.columns))
                )
                return

            self.log_line(f"Parsing column: {col}")

            parsed = df[col].apply(parse_sample_id)
            df["Analyst"] = parsed.apply(lambda x: x[0])
            df["V"]       = parsed.apply(lambda x: x[1])
            df["D"]       = parsed.apply(lambda x: x[2])
            df["Type"]    = parsed.apply(lambda x: x[3])

            self.log_line("Writing output...")
            df.to_excel(out_path, index=False)

            self.log_line(f"Done. Saved: {out_path}")
            messagebox.showinfo("Success", f"Saved parsed file:\n{out_path}")

        except Exception as e:
            self.log_line(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = SampleIdSplitterApp()
    app.mainloop()