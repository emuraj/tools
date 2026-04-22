#!/usr/bin/env python3
"""
csv_peek_to_json_gui.py

Tkinter GUI + CLI-capable utility:
- Recursively scans a folder for *.csv
- Reads first N rows (default 20) with multiple encoding fallbacks
- Writes a single JSON summary file containing:
    * file metadata
    * encoding used (best-effort)
    * dialect (best-effort)
    * headers
    * first N rows (as dicts)
    * lightweight profiling for the sampled rows (null-ish counts, unique counts, examples)

Run:
  python csv_peek_to_json_gui.py

Optional CLI:
  python csv_peek_to_json_gui.py --folder "C:\\exports" --rows 20 --out "summary.json"
"""

from __future__ import annotations

import argparse
import csv
import json
import threading
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox

DEFAULT_ENCODINGS = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
NULL_LIKE = {"", "null", "none", "n/a", "na", "nan", "(null)", "(none)"}


def _iso_local(ts: float) -> str:
    return datetime.fromtimestamp(ts).isoformat(timespec="seconds")


def _try_read_csv_first_rows(
    csv_path: Path,
    n_rows: int,
    encodings: List[str],
) -> Tuple[Optional[str], Optional[str], List[str], List[Dict[str, str]], Optional[str]]:
    """
    Returns: (encoding_used, dialect_name, headers, first_rows, error)
    """
    last_err: Optional[str] = None
    for enc in encodings:
        try:
            with csv_path.open("r", encoding=enc, errors="replace", newline="") as f:
                sample = f.read(64_000)
                f.seek(0)

                try:
                    dialect = csv.Sniffer().sniff(sample)
                    dialect_name = type(dialect).__name__
                except Exception:
                    dialect = csv.excel
                    dialect_name = "excel"

                reader = csv.DictReader(f, dialect=dialect)
                headers = reader.fieldnames or []

                rows: List[Dict[str, str]] = []
                for i, row in enumerate(reader):
                    if i >= n_rows:
                        break
                    rows.append({k: (v if v is not None else "") for k, v in row.items()})

                return enc, dialect_name, headers, rows, None
        except Exception as e:
            last_err = f"{type(e).__name__}: {e}"
            continue

    return None, None, [], [], last_err


def _profile_first_rows(headers: List[str], rows: List[Dict[str, str]]) -> Dict[str, Any]:
    prof: Dict[str, Any] = {
        "sample_row_count": len(rows),
        "columns": {},
    }

    for h in headers:
        vals: List[str] = []
        nullish = 0
        for r in rows:
            v = (r.get(h, "") or "").strip()
            if v.lower() in NULL_LIKE:
                nullish += 1
            vals.append(v)

        uniq = set(vals)
        examples = [v for v in vals if v and v.lower() not in NULL_LIKE][:5]
        if not examples:
            examples = list(list(uniq)[:5])

        prof["columns"][h] = {
            "nullish_count_in_sample": nullish,
            "unique_count_in_sample": len(uniq),
            "example_values": examples[:5],
        }

    return prof


def scan_folder_to_json(
    folder: Path,
    n_rows: int = 20,
    encodings: Optional[List[str]] = None,
    logger=None,
    progress_cb=None,
    stop_flag=None,
) -> Dict[str, Any]:
    encodings = encodings or DEFAULT_ENCODINGS
    folder = folder.expanduser().resolve()

    csv_files = sorted(folder.rglob("*.csv"))
    total = len(csv_files)

    out: Dict[str, Any] = {
        "scanned_root": str(folder),
        "generated_at_local": datetime.now().isoformat(timespec="seconds"),
        "rows_per_file": n_rows,
        "encodings_tried": encodings,
        "totals": {
            "csv_files_found": total,
            "csv_files_summarized": 0,
            "csv_files_failed": 0,
        },
        "files": [],
    }

    def log(msg: str) -> None:
        if logger:
            logger(msg)

    log(f"Scanning: {folder}")
    log(f"CSV files found: {total}")

    for idx, p in enumerate(csv_files, start=1):
        if stop_flag and stop_flag.is_set():
            log("Stop requested. Exiting scan early.")
            break

        if progress_cb:
            progress_cb(idx - 1, total)

        meta = {
            "file_name": p.name,
            "relative_path": str(p.relative_to(folder)),
            "absolute_path": str(p),
            "size_bytes": p.stat().st_size,
            "modified_local": _iso_local(p.stat().st_mtime),
        }

        log(f"[{idx}/{total}] Reading {p.name}")

        enc_used, dialect_name, headers, first_rows, err = _try_read_csv_first_rows(
            p, n_rows, encodings
        )

        if err:
            out["totals"]["csv_files_failed"] += 1
            out["files"].append(
                {
                    **meta,
                    "status": "FAIL",
                    "error": err,
                    "encoding_used": enc_used,
                    "dialect": dialect_name,
                    "headers": headers,
                    "first_rows": first_rows,
                }
            )
            log(f"  FAIL: {err}")
        else:
            profile = _profile_first_rows(headers, first_rows)
            out["totals"]["csv_files_summarized"] += 1
            out["files"].append(
                {
                    **meta,
                    "status": "OK",
                    "encoding_used": enc_used,
                    "dialect": dialect_name,
                    "headers": headers,
                    "first_rows": first_rows,
                    "profile": profile,
                }
            )
            log(f"  OK: {len(headers)} columns, {len(first_rows)} sampled rows")

    if progress_cb:
        progress_cb(total, total)

    return out


def write_json(payload: Dict[str, Any], out_path: Path) -> None:
    out_path = out_path.expanduser().resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


# ──────────────────────────────────────────────────────────────────────
# Tkinter GUI
# ──────────────────────────────────────────────────────────────────────
@dataclass
class GuiState:
    stop_event: threading.Event


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("CSV Peek → JSON Summary")
        self.geometry("980x640")
        self.minsize(880, 560)

        self.state_obj = GuiState(stop_event=threading.Event())

        # Variables
        self.var_folder = tk.StringVar(value="")
        self.var_out = tk.StringVar(value=str(Path.cwd() / "csv_peek_summary.json"))
        self.var_rows = tk.IntVar(value=20)

        # Layout
        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        top = tk.Frame(self)
        top.pack(fill="x", **pad)

        # Folder row
        tk.Label(top, text="Input folder (scan recursively):").grid(row=0, column=0, sticky="w")
        ent_folder = tk.Entry(top, textvariable=self.var_folder)
        ent_folder.grid(row=1, column=0, sticky="ew")
        btn_browse = tk.Button(top, text="Browse…", command=self._browse_folder)
        btn_browse.grid(row=1, column=1, padx=(8, 0))

        # Output row
        tk.Label(top, text="Output JSON file:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ent_out = tk.Entry(top, textvariable=self.var_out)
        ent_out.grid(row=3, column=0, sticky="ew")
        btn_out = tk.Button(top, text="Save As…", command=self._browse_outfile)
        btn_out.grid(row=3, column=1, padx=(8, 0))

        # Rows + buttons
        mid = tk.Frame(self)
        mid.pack(fill="x", **pad)

        tk.Label(mid, text="Rows per CSV (sample):").grid(row=0, column=0, sticky="w")
        spn = tk.Spinbox(mid, from_=1, to=500, textvariable=self.var_rows, width=8)
        spn.grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.btn_run = tk.Button(mid, text="Run Scan", command=self._run_scan)
        self.btn_run.grid(row=0, column=2, padx=(16, 0))
        self.btn_stop = tk.Button(mid, text="Stop", command=self._stop, state="disabled")
        self.btn_stop.grid(row=0, column=3, padx=(8, 0))

        # Progress
        self.lbl_progress = tk.Label(mid, text="Progress: 0/0")
        self.lbl_progress.grid(row=0, column=4, padx=(16, 0), sticky="w")

        # Log area
        bot = tk.Frame(self)
        bot.pack(fill="both", expand=True, **pad)

        tk.Label(bot, text="Log:").pack(anchor="w")
        self.txt = tk.Text(bot, wrap="word", height=20)
        self.txt.pack(fill="both", expand=True)

        # Grid expansion
        top.grid_columnconfigure(0, weight=1)

    def _browse_folder(self) -> None:
        d = filedialog.askdirectory(title="Select folder containing CSV exports")
        if d:
            self.var_folder.set(d)

    def _browse_outfile(self) -> None:
        f = filedialog.asksaveasfilename(
            title="Save JSON summary as…",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if f:
            self.var_out.set(f)

    def _log(self, msg: str) -> None:
        self.txt.insert("end", msg.rstrip() + "\n")
        self.txt.see("end")
        self.update_idletasks()

    def _set_progress(self, cur: int, total: int) -> None:
        self.lbl_progress.config(text=f"Progress: {cur}/{total}")
        self.update_idletasks()

    def _stop(self) -> None:
        self.state_obj.stop_event.set()
        self._log("Stop signal sent…")

    def _run_scan(self) -> None:
        folder = (self.var_folder.get() or "").strip()
        out_path = (self.var_out.get() or "").strip()

        if not folder:
            messagebox.showerror("Missing input", "Please select an input folder.")
            return
        if not out_path:
            messagebox.showerror("Missing output", "Please select an output JSON file.")
            return

        in_dir = Path(folder)
        if not in_dir.exists() or not in_dir.is_dir():
            messagebox.showerror("Invalid input", "Input folder does not exist or is not a folder.")
            return

        # Reset stop
        self.state_obj.stop_event.clear()

        # Disable controls during run
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self._set_progress(0, 0)
        self._log("────────────────────────────────────────────────────")
        self._log(f"Run started: {datetime.now().isoformat(timespec='seconds')}")

        def worker() -> None:
            try:
                payload = scan_folder_to_json(
                    in_dir,
                    n_rows=int(self.var_rows.get()),
                    encodings=DEFAULT_ENCODINGS,
                    logger=self._log,
                    progress_cb=self._set_progress,
                    stop_flag=self.state_obj.stop_event,
                )
                write_json(payload, Path(out_path))
                self._log("────────────────────────────────────────────────────")
                self._log(f"✅ Wrote JSON summary: {Path(out_path).expanduser().resolve()}")
                self._log(
                    f"Totals: found={payload['totals']['csv_files_found']}, "
                    f"ok={payload['totals']['csv_files_summarized']}, "
                    f"fail={payload['totals']['csv_files_failed']}"
                )
                self._log("Done.")
            except Exception:
                self._log("────────────────────────────────────────────────────")
                self._log("❌ ERROR")
                self._log(traceback.format_exc())
                messagebox.showerror("Error", "Scan failed. See log for details.")
            finally:
                # Re-enable controls on UI thread
                self.after(0, self._finish_run)

        threading.Thread(target=worker, daemon=True).start()

    def _finish_run(self) -> None:
        self.btn_run.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.state_obj.stop_event.clear()


def main() -> None:
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("--folder", help="Folder to scan recursively for CSVs")
    ap.add_argument("--rows", type=int, default=20, help="Rows per CSV to sample (default 20)")
    ap.add_argument("--out", help="Output JSON path")
    args = ap.parse_args()

    # If CLI args provided, run headless; otherwise launch GUI.
    if args.folder and args.out:
        payload = scan_folder_to_json(Path(args.folder), n_rows=args.rows)
        write_json(payload, Path(args.out))
        print(f"Wrote: {Path(args.out).expanduser().resolve()}")
    else:
        App().mainloop()


if __name__ == "__main__":
    main()