#!/usr/bin/env python3
"""
DMS Roles Filler (Standalone)

Small Tkinter utility to post-process a DMS migration file and populate the
'Role(s)' column using a Training/Roles CSV.

Inputs via GUI:
  1) DMS Migration file (CSV or XLSX)
     - Must contain 'Legacy Document #' and ideally 'Role(s)' columns.
       If 'Role(s)' is missing, it will be created.

  2) Training/Roles report CSV
     - Must contain 'Training Number' and 'Training Plan Name' columns.
     - Training Plan Name values look like 'TP-0001_ Data Integrity'.

  3) Output folder
     - The script will write a new file in this folder named:
         <migration_basename>_with_roles.<ext>

Mapping logic (catch-all):
  • Take Training Number, e.g.:
      'WI-0005-000001'
      'PRN-0013-000001'
      'POL-0001-000002'
      'QM-0001-000002'
      'A13-000002'
      'B22A-000001'
      '6000-000001'
      '1022-000003'
      'CORP-000001'        (exception: slug is meaningful)
  • Step 1: If it ends with '-00000n', strip that trailing slug for most codes:
      WI-0005-000001   → base 'WI-0005'
      PRN-0013-000001  → base 'PRN-0013'
      POL-0001-000002  → base 'POL-0001'
      QM-0001-000002   → base 'QM-0001'
      A13-000002       → base 'A13'
      B22A-000001      → base 'B22A'
      6000-000001      → base '6000'
      1022-000003      → base '1022'
      SOP-2157-000002  → base 'SOP-2157'

      EXCEPTION: for CORP, do NOT strip:
      CORP-000001      → base 'CORP-000001'

  • Step 2: Normalize the base:
      - If base matches 'SOP-\\d+'  → use as-is (uppercased).
      - If base is pure 3–5 digits → treat as SOP, prefix with 'SOP-'.
      - Otherwise                  → use base.upper() (WI-0005, PRN-0013, A13, B22A, CORP-000001, etc.)

  • Training Plan Name 'TP-0001_ Data Integrity' → 'TP-0001: Data Integrity'.

  • Aggregate unique plan names per normalized key, joined by '|':
      'TP-0001: Data Integrity|TP-0047: Manufacturing Unit 4A/B (Cell Culture: Thaw, Subculture & Harvest)'

  • In the migration file:
      - For each row, look up 'Legacy Document #' in the roles map.
      - If a match exists, set Role(s) to the aggregated string.
      - If not, leave Role(s) as-is (blank).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, BOTH, END


# ------------ Loader / writer helpers ---------------------------------------


def guess_migration_loader(path: Path):
    """Return a callable to load the migration file based on extension."""
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        return lambda p: pd.read_excel(p)
    elif ext in (".csv", ".txt"):
        return lambda p: pd.read_csv(p)
    else:
        raise ValueError(f"Unsupported migration file extension: {ext}")


def guess_migration_writer(path: Path):
    """Return a callable to write the migration file based on extension."""
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        def _write(df: pd.DataFrame, p: Path):
            df.to_excel(p, index=False)
        return _write
    elif ext in (".csv", ".txt"):
        def _write(df: pd.DataFrame, p: Path):
            df.to_csv(p, index=False)
        return _write
    else:
        raise ValueError(f"Unsupported output file extension: {ext}")


def read_training_csv(path: Path) -> pd.DataFrame:
    """
    Robust CSV reader for the Training/Roles file.

    Tries multiple encodings in order:
      utf-8 → utf-8-sig → cp1252 → latin1
    and falls back to utf-8-sig with replacement if needed.
    """
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin1"]
    last_err: Exception | None = None

    for enc in encodings:
        try:
            return pd.read_csv(path, encoding=enc)
        except UnicodeDecodeError as e:
            last_err = e
            continue

    # Final fallback: utf-8-sig with replacement
    try:
        with path.open("r", encoding="utf-8-sig", errors="replace") as fh:
            return pd.read_csv(fh)
    except Exception as e:
        raise last_err or e


# ------------ Training / Roles utilities ------------------------------------


def normalize_training_number(tnum: str) -> Optional[str]:
    """
    Normalize a Training Number into a key that matches 'Legacy Document #' in the migration file.

    Generic algorithm:
      1) If it ends with '-00000n', strip that trailing slug for most codes,
         BUT keep 'CORP-000001' as-is.
      2) Normalize the base:
         - If base matches 'SOP-\\d+'  → base.upper().
         - If base is pure 3–5 digits → 'SOP-' + base.
         - Otherwise                  → base.upper().
    """
    if not tnum:
        return None
    s = str(tnum).strip()
    if not s:
        return None

    # Step 1: strip trailing '-000000' for most codes, but honor CORP exception
    m = re.match(r"^(.*?)-(\d{6})$", s)
    if m:
        head = m.group(1).strip()
        # CORP exception: do not strip, keep full 'CORP-000001'
        if head.upper() == "CORP":
            base = s
        else:
            base = head
    else:
        base = s

    # Step 2: normalize the base
    # Case: SOP-#### (already SOP-prefixed)
    if re.fullmatch(r"SOP-\d+", base, re.IGNORECASE):
        return base.upper()

    # Case: pure numeric (3–5 digits) → SOP-####
    if re.fullmatch(r"\d{3,5}", base):
        return f"SOP-{base}"

    # Otherwise, just use the base itself (uppercased)
    return base.upper()


def transform_plan_name(raw_plan: str) -> str:
    """
    Transform 'TP-0001_ Data Integrity' → 'TP-0001: Data Integrity'.

    Only the FIRST underscore (with optional spaces around it) is replaced with ': '.
    """
    raw_plan = str(raw_plan or "").strip()
    if not raw_plan:
        return ""
    return re.sub(r"\s*_\s*", ": ", raw_plan, count=1)


def build_roles_map(training_df: pd.DataFrame) -> Dict[str, str]:
    """
    Given the Training/Roles dataframe, build a mapping:

        legacy_key (e.g. 'SOP-1018', 'A13', 'WI-0005', 'CORP-000001') →
            'TP-0001: Data Integrity|TP-0047: Manufacturing Unit 4A/B ...'

    Using:
      - 'Training Number' as the source key.
      - 'Training Plan Name' as the source text for role descriptions.
    """
    cols = {str(c).strip().lower(): c for c in training_df.columns}

    def _col(*names: str) -> Optional[str]:
        for n in names:
            n_low = n.lower()
            if n_low in cols:
                return cols[n_low]
        return None

    col_tnum = _col("training number")
    col_plan = _col("training plan name")

    if col_tnum is None or col_plan is None:
        raise ValueError(
            "Training CSV must contain 'Training Number' and 'Training Plan Name' columns "
            f"(found: {list(training_df.columns)})"
        )

    from collections import OrderedDict

    # Map: legacy_key → OrderedDict of roles (deduplicate while preserving order)
    tmp: Dict[str, OrderedDict[str, None]] = {}

    for _, row in training_df.iterrows():
        tnum = str(row.get(col_tnum, "")).strip()
        raw_plan = str(row.get(col_plan, "")).strip()
        if not tnum or not raw_plan:
            continue

        legacy_key = normalize_training_number(tnum)
        if not legacy_key:
            continue

        role = transform_plan_name(raw_plan)
        if not role:
            continue

        bucket = tmp.setdefault(legacy_key, OrderedDict())
        if role not in bucket:
            bucket[role] = None

    # Collapse to '|' separated strings
    roles_map: Dict[str, str] = {k: "|".join(od.keys()) for k, od in tmp.items()}
    return roles_map


def apply_roles_to_migration(
    mig_df: pd.DataFrame,
    roles_map: Dict[str, str],
) -> Tuple[pd.DataFrame, int]:
    """
    Update the migration dataframe with Roles, based on roles_map.

    roles_map: Legacy Document # (e.g. 'SOP-1018', 'A13', 'WI-0005', 'CORP-000001') → role string.

    Returns:
      (updated_dataframe, number_of_documents_with_roles_applied)
    """
    if "Legacy Document #" not in mig_df.columns:
        raise ValueError("Migration file is missing 'Legacy Document #' column.")
    if "Role(s)" not in mig_df.columns:
        # If Roles column is missing, create it
        mig_df["Role(s)"] = ""

    df = mig_df.copy()
    legacy_series = df["Legacy Document #"].astype(str).str.strip()

    roles_col: List[str] = []
    applied = 0

    for _, legacy in legacy_series.items():
        roles = roles_map.get(legacy, "")
        if roles:
            applied += 1
        roles_col.append(roles)

    df["Role(s)"] = roles_col
    return df, applied


# ------------ Tkinter GUI ----------------------------------------------------


class RolesFillerGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("DMS Migration Roles Filler")
        self.geometry("900x260")

        self._build_widgets()

    def _build_widgets(self) -> None:
        frm = ttk.LabelFrame(self, text="Inputs")
        frm.pack(fill="x", padx=10, pady=(10, 4))

        def row_file(lbl: str) -> ttk.Entry:
            rr = ttk.Frame(frm)
            rr.pack(fill="x", pady=3)
            ttk.Label(rr, text=lbl, width=30, anchor="w").pack(side="left")
            ent = ttk.Entry(rr, width=80)
            ent.pack(side="left", fill="x", expand=True)

            def browse():
                p = filedialog.askopenfilename(
                    filetypes=[
                        ("All supported", "*.csv *.xlsx *.xls *.txt"),
                        ("CSV files", "*.csv *.txt"),
                        ("Excel files", "*.xlsx *.xls *.xlsm"),
                        ("All files", "*.*"),
                    ]
                )
                if p:
                    ent.delete(0, "end")
                    ent.insert(0, p)
            ttk.Button(rr, text="Browse…", command=browse).pack(side="right")
            return ent

        def row_dir(lbl: str) -> ttk.Entry:
            rr = ttk.Frame(frm)
            rr.pack(fill="x", pady=3)
            ttk.Label(rr, text=lbl, width=30, anchor="w").pack(side="left")
            ent = ttk.Entry(rr, width=80)
            ent.pack(side="left", fill="x", expand=True)

            def browse():
                p = filedialog.askdirectory()
                if p:
                    ent.delete(0, "end")
                    ent.insert(0, p)
            ttk.Button(rr, text="Browse…", command=browse).pack(side="right")
            return ent

        self.ent_mig = row_file("DMS Migration file (CSV or XLSX):")
        self.ent_training = row_file("Training/Roles report CSV:")
        self.ent_out_dir = row_dir("Output folder:")

        # Control row
        ctrl = ttk.Frame(self)
        ctrl.pack(fill="x", padx=10, pady=(8, 6))
        self.btn_run = ttk.Button(ctrl, text="Run", command=self._run)
        self.btn_run.pack(side="left")
        ttk.Button(ctrl, text="Close", command=self.destroy).pack(side="right")

        # Status line
        self.var_status = tk.StringVar(value="Idle.")
        ttk.Label(self, textvariable=self.var_status, anchor="w").pack(fill="x", padx=10, pady=(0, 8))

    def _run(self) -> None:
        mig_path = Path(self.ent_mig.get().strip())
        training_path = Path(self.ent_training.get().strip())
        out_dir = Path(self.ent_out_dir.get().strip())

        # Basic validation
        if not mig_path.is_file():
            messagebox.showerror("Error", "Please select a valid DMS Migration file.")
            return
        if not training_path.is_file():
            messagebox.showerror("Error", "Please select a valid Training/Roles report CSV.")
            return
        if not out_dir.is_dir():
            messagebox.showerror("Error", f"Output folder does not exist:\n{out_dir}")
            return

        self.btn_run.config(state="disabled")
        self.var_status.set("Running …")
        self.update_idletasks()

        try:
            # Load migration
            mig_loader = guess_migration_loader(mig_path)
            mig_df = mig_loader(mig_path)

            # Load training CSV with robust encoding handling
            training_df = read_training_csv(training_path)

            # Build roles map
            roles_map = build_roles_map(training_df)

            # Apply to migration
            updated_df, applied = apply_roles_to_migration(mig_df, roles_map)

            # Construct output file name: <basename>_with_roles.<ext>
            out_name = f"{mig_path.stem}_with_roles{mig_path.suffix}"
            out_path = out_dir / out_name

            # Write output
            writer = guess_migration_writer(out_path)
            writer(updated_df, out_path)

            self.var_status.set(
                f"Done. Filled Roles for {applied} document(s). Output: {out_path}"
            )
            messagebox.showinfo(
                "Complete",
                f"Roles have been populated for {applied} document(s).\n\n"
                f"Output written to:\n{out_path}",
            )
        except Exception as exc:
            self.var_status.set("Error.")
            messagebox.showerror("Error", f"An error occurred:\n{exc}")
        finally:
            self.btn_run.config(state="normal")
            self.update_idletasks()


if __name__ == "__main__":
    app = RolesFillerGUI()
    app.mainloop()
