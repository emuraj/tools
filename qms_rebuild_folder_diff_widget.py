#!/usr/bin/env python3
"""
qms_rebuild_folder_diff_widget.py

Tkinter widget to compare OLD vs NEW QMS rebuild output roots and generate:

1) Per-record-type CSV tables (CR / DEV / CAPA / INV / OOS / EC / etc.)
   - Columns:
       record_id
       num_new_items
       old_folder_contents
       new_folder_contents
       new_items_only
   - Includes unchanged records (num_new_items = 0)
   - Sorted by num_new_items DESC, then record_id ASC
   - One CSV per record type, e.g.:
       folder_diff_CR.csv
       folder_diff_DEV.csv
       folder_diff_CAPA.csv
       ...

2) One overall record-type summary CSV:
       record_type_summary.csv

3) Expanded "new items" CSVs where each new item gets its own row and record_id repeats:
   a) Per-record-type expanded files:
       folder_diff_<TYPE>_new_items_expanded.csv
       Columns:
         record_id, new_item_filename, new_item_ext, new_item_kind
   b) One cross-type expanded file:
       folder_diff_new_items_expanded_ALL.csv
       Columns:
         record_type, record_id, new_item_filename, new_item_ext, new_item_kind

Record folder detection:
- A folder is treated as a record folder if it contains:
    * <RID>.json   (strongest signal), or
    * <RID>_summary*.docx, or
    * files containing "<RID>_ATTACHMENT_" marker
- RID pattern: PREFIX-NUMBER, e.g. CR-00005, DEV-00102, CAPA-00435, CAR-00014, etc.

"New items" definition:
- Files present in NEW folder but not in OLD folder (by filename).
- Size-changes/removals are detected and logged but "num_new_items" counts only additions.

Outputs:
- CSVs written to an output folder you choose in the GUI.
- Optional JSON report for full fidelity inventories + diffs.

No external dependencies beyond Python stdlib + tkinter.
"""

from __future__ import annotations

import csv
import json
import os
import re
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

import tkinter as tk
from tkinter import filedialog, messagebox


# ---------------------------- Patterns / Constants ----------------------------

RECORD_ID_RE = re.compile(r"^(?P<prefix>[A-Z]+)-(?P<num>\d{1,})$", re.IGNORECASE)
JSON_NAME_RE = re.compile(r"^(?P<rid>[A-Z]+-\d+)\.json$", re.IGNORECASE)
SUMMARY_NAME_RE = re.compile(r"^(?P<rid>[A-Z]+-\d+)_summary.*$", re.IGNORECASE)

ATTACH_MARKER = "_ATTACHMENT_"

# Known record types (for grouping). If your environment has more, they will still appear.
# This is only used for display ordering, not detection.
TYPE_ORDER = ["CR", "DEV", "CAPA", "CAR", "INV", "OOS", "EC"]


# ---------------------------- Data Structures ----------------------------

@dataclass
class FileInfo:
    name: str
    rel_path: str
    size: int
    mtime: float


@dataclass
class RecordInventory:
    record_id: str
    record_type: str
    record_dir_abs: str
    record_dir_rel: str
    files: Dict[str, FileInfo]  # keyed by filename
    total_files: int
    attachment_files: int
    total_bytes: int


@dataclass
class RecordDeltaRow:
    record_id: str
    record_type: str
    num_new_items: int
    old_contents: List[str]
    new_contents: List[str]
    new_items_only: List[str]
    removed_items: List[str]
    changed_size: List[Tuple[str, int, int]]  # (name, old_size, new_size)


@dataclass
class TypeSummaryRow:
    record_type: str
    records_compared: int
    records_with_new_items: int
    pct_with_new_items: float
    total_new_items: int
    avg_new_items_per_record: float
    median_new_items_per_record: float
    max_new_items_single_record: int
    total_old_files: int
    total_new_files: int
    net_file_increase: int
    total_removed_items: int
    records_missing_in_old: int
    records_missing_in_new: int


# ---------------------------- Helpers ----------------------------

def _now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _safe_rel(p: Path, root: Path) -> str:
    try:
        return str(p.relative_to(root))
    except Exception:
        return str(p)


def _record_type_of(record_id: str) -> str:
    m = RECORD_ID_RE.match(record_id.strip())
    if not m:
        return "UNKNOWN"
    return m.group("prefix").upper()


def _find_record_id_in_dir(d: Path) -> Optional[str]:
    """
    Determine record id by scanning filenames in the directory.
    We intentionally derive RID from filenames (not folder names) to be robust.
    """
    try:
        for p in d.iterdir():
            if not p.is_file():
                continue

            m = JSON_NAME_RE.match(p.name)
            if m:
                return m.group("rid").upper()

            # summary file may be "CR-00005_summary.docx"
            m = SUMMARY_NAME_RE.match(p.name)
            if m:
                rid = m.group("rid").upper()
                if RECORD_ID_RE.match(rid):
                    return rid

            # attachments: "<RID>_ATTACHMENT_..."
            if ATTACH_MARKER in p.name:
                rid_guess = p.name.split(ATTACH_MARKER, 1)[0].upper()
                if RECORD_ID_RE.match(rid_guess):
                    return rid_guess
    except Exception:
        return None

    return None


def _build_inventory(root: Path, logger=None, stop_event: Optional[threading.Event] = None) -> Dict[str, RecordInventory]:
    """
    Walk root recursively and build inventory of record folders keyed by record_id.
    """
    root = root.expanduser().resolve()
    inv: Dict[str, RecordInventory] = {}

    def log(msg: str) -> None:
        if logger:
            logger(msg)

    log(f"Indexing: {root}")

    for dirpath, _, filenames in os.walk(root):
        if stop_event and stop_event.is_set():
            log("Stop requested; aborting inventory build.")
            break

        d = Path(dirpath)
        rid = _find_record_id_in_dir(d)
        if not rid:
            continue

        if rid in inv:
            # Keep first (deterministic) and log
            log(f"⚠️ Duplicate record_id detected: {rid} at {_safe_rel(d, root)} (keeping first at {inv[rid].record_dir_rel})")
            continue

        record_type = _record_type_of(rid)

        files: Dict[str, FileInfo] = {}
        total_bytes = 0
        attachment_files = 0

        for fn in filenames:
            fp = d / fn
            try:
                st = fp.stat()
            except Exception:
                continue

            if ATTACH_MARKER in fn:
                attachment_files += 1

            fi = FileInfo(
                name=fn,
                rel_path=_safe_rel(fp, root),
                size=st.st_size,
                mtime=st.st_mtime,
            )
            files[fn] = fi
            total_bytes += st.st_size

        inv[rid] = RecordInventory(
            record_id=rid,
            record_type=record_type,
            record_dir_abs=str(d),
            record_dir_rel=_safe_rel(d, root),
            files=files,
            total_files=len(files),
            attachment_files=attachment_files,
            total_bytes=total_bytes,
        )

    log(f"Inventory complete: {len(inv)} record folders found.")
    return inv


def _diff_records(old_inv: Dict[str, RecordInventory], new_inv: Dict[str, RecordInventory]) -> List[RecordDeltaRow]:
    """
    Compare inventories and produce per-record rows including:
    - old/new content lists
    - new_items_only (added)
    - removed_items
    - changed_size list
    """
    all_ids = sorted(set(old_inv.keys()) | set(new_inv.keys()))
    rows: List[RecordDeltaRow] = []

    for rid in all_ids:
        o = old_inv.get(rid)
        n = new_inv.get(rid)

        record_type = (n.record_type if n else (o.record_type if o else _record_type_of(rid)))

        o_names = sorted(o.files.keys()) if o else []
        n_names = sorted(n.files.keys()) if n else []

        o_set = set(o_names)
        n_set = set(n_names)

        added = sorted(n_set - o_set)
        removed = sorted(o_set - n_set)

        changed_size: List[Tuple[str, int, int]] = []
        for name in sorted(o_set & n_set):
            osz = o.files[name].size if o else -1
            nsz = n.files[name].size if n else -1
            if osz != nsz:
                changed_size.append((name, osz, nsz))

        rows.append(
            RecordDeltaRow(
                record_id=rid,
                record_type=record_type,
                num_new_items=len(added),
                old_contents=o_names,
                new_contents=n_names,
                new_items_only=added,
                removed_items=removed,
                changed_size=changed_size,
            )
        )

    return rows


def _median(nums: List[int]) -> float:
    if not nums:
        return 0.0
    s = sorted(nums)
    n = len(s)
    mid = n // 2
    if n % 2 == 1:
        return float(s[mid])
    return (s[mid - 1] + s[mid]) / 2.0


def _type_summary(
    record_rows: List[RecordDeltaRow],
    old_inv: Dict[str, RecordInventory],
    new_inv: Dict[str, RecordInventory],
) -> List[TypeSummaryRow]:
    """
    Build per-record-type stats.
    """
    by_type: Dict[str, List[RecordDeltaRow]] = {}
    for r in record_rows:
        by_type.setdefault(r.record_type, []).append(r)

    out: List[TypeSummaryRow] = []

    for rtype, rows in by_type.items():
        compared = len(rows)
        with_new = sum(1 for r in rows if r.num_new_items > 0)
        pct = (with_new / compared * 100.0) if compared else 0.0
        total_new = sum(r.num_new_items for r in rows)
        nums = [r.num_new_items for r in rows]
        avg = (total_new / compared) if compared else 0.0
        med = _median(nums)
        mx = max(nums) if nums else 0

        # totals of file counts (old/new)
        # Use inventory only if record exists in that side
        total_old_files = 0
        total_new_files = 0
        removed_total = 0
        missing_in_old = 0
        missing_in_new = 0

        for r in rows:
            if r.record_id in old_inv:
                total_old_files += old_inv[r.record_id].total_files
            else:
                missing_in_old += 1
            if r.record_id in new_inv:
                total_new_files += new_inv[r.record_id].total_files
            else:
                missing_in_new += 1

            removed_total += len(r.removed_items)

        net = total_new_files - total_old_files

        out.append(
            TypeSummaryRow(
                record_type=rtype,
                records_compared=compared,
                records_with_new_items=with_new,
                pct_with_new_items=pct,
                total_new_items=total_new,
                avg_new_items_per_record=avg,
                median_new_items_per_record=med,
                max_new_items_single_record=mx,
                total_old_files=total_old_files,
                total_new_files=total_new_files,
                net_file_increase=net,
                total_removed_items=removed_total,
                records_missing_in_old=missing_in_old,
                records_missing_in_new=missing_in_new,
            )
        )

    # deterministic ordering: known types first, then others
    def type_sort_key(t: str) -> Tuple[int, str]:
        t_up = t.upper()
        if t_up in TYPE_ORDER:
            return (TYPE_ORDER.index(t_up), t_up)
        return (999, t_up)

    out.sort(key=lambda r: type_sort_key(r.record_type))
    return out


def _group_by_type(rows: List[RecordDeltaRow]) -> Dict[str, List[RecordDeltaRow]]:
    out: Dict[str, List[RecordDeltaRow]] = {}
    for r in rows:
        out.setdefault(r.record_type, []).append(r)
    return out


def _write_per_type_csvs(out_dir: Path, grouped: Dict[str, List[RecordDeltaRow]]) -> List[Path]:
    """
    Original per-type CSVs (record-level, list-in-cell). UNCHANGED behavior.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    written: List[Path] = []

    # order types: TYPE_ORDER first, then alphabetical remainder
    types = list(grouped.keys())
    ordered = [t for t in TYPE_ORDER if t in grouped] + sorted([t for t in types if t not in TYPE_ORDER])

    for t in ordered:
        rows = grouped[t]
        # Sort: num_new_items desc, then record_id asc
        rows_sorted = sorted(rows, key=lambda r: (-r.num_new_items, r.record_id))

        path = out_dir / f"folder_diff_{t}.csv"
        with path.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["record_id", "num_new_items", "old_folder_contents", "new_folder_contents", "new_items_only"])
            for r in rows_sorted:
                w.writerow([
                    r.record_id,
                    r.num_new_items,
                    "; ".join(r.old_contents),
                    "; ".join(r.new_contents),
                    "; ".join(r.new_items_only) if r.new_items_only else "",
                ])

        written.append(path)

    return written


def _classify_item(name: str) -> str:
    n = name.lower()
    if n.endswith(".json"):
        return "json"
    if "_summary" in n and (n.endswith(".docx") or n.endswith(".pdf")):
        return "summary"
    if ATTACH_MARKER.lower() in n:
        return "attachment"
    return "other"


def _file_ext(name: str) -> str:
    return Path(name).suffix.lower().lstrip(".")


def _write_per_type_new_items_expanded_csvs(out_dir: Path, grouped: Dict[str, List[RecordDeltaRow]]) -> List[Path]:
    """
    Per-type expanded CSVs: one row per new item. record_id repeats.
    Only additions are emitted (rows with new_items_only).
    Always writes a file per type (header-only if no additions).
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    written: List[Path] = []

    # order types: TYPE_ORDER first, then alphabetical remainder
    types = list(grouped.keys())
    ordered = [t for t in TYPE_ORDER if t in grouped] + sorted([t for t in types if t not in TYPE_ORDER])

    for t in ordered:
        rows = grouped[t]

        expanded_rows: List[Tuple[str, str, str, str]] = []
        # (record_id, new_item_filename, ext, item_kind)

        for r in rows:
            if not r.new_items_only:
                continue
            for fn in r.new_items_only:
                expanded_rows.append((
                    r.record_id,
                    fn,
                    _file_ext(fn),
                    _classify_item(fn),
                ))

        path = out_dir / f"folder_diff_{t}_new_items_expanded.csv"
        with path.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["record_id", "new_item_filename", "new_item_ext", "new_item_kind"])
            for rid, fn, ext, kind in sorted(expanded_rows, key=lambda x: (x[0], x[1])):
                w.writerow([rid, fn, ext, kind])

        written.append(path)

    return written


def _write_all_new_items_expanded_csv(out_dir: Path, record_rows: List[RecordDeltaRow]) -> Path:
    """
    Single cross-type expanded CSV: one row per new item across all record types.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "folder_diff_new_items_expanded_ALL.csv"

    expanded: List[Tuple[str, str, str, str, str]] = []
    # (record_type, record_id, new_item_filename, ext, item_kind)

    for r in record_rows:
        if not r.new_items_only:
            continue
        for fn in r.new_items_only:
            expanded.append((
                r.record_type,
                r.record_id,
                fn,
                _file_ext(fn),
                _classify_item(fn),
            ))

    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["record_type", "record_id", "new_item_filename", "new_item_ext", "new_item_kind"])
        for rtype, rid, fn, ext, kind in sorted(expanded, key=lambda x: (x[0], x[1], x[2])):
            w.writerow([rtype, rid, fn, ext, kind])

    return path


def _write_type_summary_csv(out_dir: Path, summary_rows: List[TypeSummaryRow]) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "record_type_summary.csv"
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "record_type",
            "records_compared",
            "records_with_new_items",
            "pct_with_new_items",
            "total_new_items",
            "avg_new_items_per_record",
            "median_new_items_per_record",
            "max_new_items_single_record",
            "total_old_files",
            "total_new_files",
            "net_file_increase",
            "total_removed_items",
            "records_missing_in_old",
            "records_missing_in_new",
        ])
        for r in summary_rows:
            w.writerow([
                r.record_type,
                r.records_compared,
                r.records_with_new_items,
                f"{r.pct_with_new_items:.2f}",
                r.total_new_items,
                f"{r.avg_new_items_per_record:.4f}",
                f"{r.median_new_items_per_record:.4f}",
                r.max_new_items_single_record,
                r.total_old_files,
                r.total_new_files,
                r.net_file_increase,
                r.total_removed_items,
                r.records_missing_in_old,
                r.records_missing_in_new,
            ])
    return path


def _write_full_json(out_dir: Path, old_root: Path, new_root: Path,
                     old_inv: Dict[str, RecordInventory], new_inv: Dict[str, RecordInventory],
                     record_rows: List[RecordDeltaRow], summary_rows: List[TypeSummaryRow]) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "folder_diff_full.json"

    def inv_to_dict(inv: Dict[str, RecordInventory]) -> Dict[str, Any]:
        d: Dict[str, Any] = {}
        for rid, r in inv.items():
            d[rid] = {
                "record_type": r.record_type,
                "record_dir_rel": r.record_dir_rel,
                "record_dir_abs": r.record_dir_abs,
                "total_files": r.total_files,
                "attachment_files": r.attachment_files,
                "total_bytes": r.total_bytes,
                "files": {
                    fn: {
                        "rel_path": fi.rel_path,
                        "size": fi.size,
                        "mtime": fi.mtime,
                    } for fn, fi in r.files.items()
                }
            }
        return d

    payload = {
        "generated_at_local": datetime.now().isoformat(timespec="seconds"),
        "old_root": str(old_root.resolve()),
        "new_root": str(new_root.resolve()),
        "record_type_summary": [r.__dict__ for r in summary_rows],
        "record_rows": [
            {
                "record_id": r.record_id,
                "record_type": r.record_type,
                "num_new_items": r.num_new_items,
                "old_folder_contents": r.old_contents,
                "new_folder_contents": r.new_contents,
                "new_items_only": r.new_items_only,
                "removed_items": r.removed_items,
                "changed_size": [{"name": n, "old_size": osz, "new_size": nsz} for (n, osz, nsz) in r.changed_size],
            }
            for r in record_rows
        ],
        "inventories": {
            "old": inv_to_dict(old_inv),
            "new": inv_to_dict(new_inv),
        }
    }

    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


# ---------------------------- Tkinter Widget ----------------------------

class FolderDiffWidget(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("QMS Rebuild Folder Diff (OLD vs NEW)")
        self.geometry("1120x740")
        self.minsize(980, 620)

        self.var_old = tk.StringVar(value="")
        self.var_new = tk.StringVar(value="")
        self.var_out = tk.StringVar(value=str(Path.cwd() / f"QMS_folder_diff_reports_{_now_stamp()}"))
        self.var_write_json = tk.BooleanVar(value=True)

        self.stop_event = threading.Event()
        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        frm = tk.Frame(self)
        frm.pack(fill="x", **pad)

        tk.Label(frm, text="OLD rebuild root:").grid(row=0, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.var_old).grid(row=1, column=0, sticky="ew")
        tk.Button(frm, text="Browse…", command=self._browse_old).grid(row=1, column=1, padx=(8, 0))

        tk.Label(frm, text="NEW rebuild root:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        tk.Entry(frm, textvariable=self.var_new).grid(row=3, column=0, sticky="ew")
        tk.Button(frm, text="Browse…", command=self._browse_new).grid(row=3, column=1, padx=(8, 0))

        tk.Label(frm, text="Output report folder:").grid(row=4, column=0, sticky="w", pady=(10, 0))
        tk.Entry(frm, textvariable=self.var_out).grid(row=5, column=0, sticky="ew")
        tk.Button(frm, text="Browse…", command=self._browse_out).grid(row=5, column=1, padx=(8, 0))

        tk.Checkbutton(frm, text="Also write full JSON (folder_diff_full.json)", variable=self.var_write_json)\
            .grid(row=6, column=0, sticky="w", pady=(8, 0))

        frm.grid_columnconfigure(0, weight=1)

        btns = tk.Frame(self)
        btns.pack(fill="x", **pad)

        self.btn_run = tk.Button(btns, text="Run Diff + Write Reports", command=self._run)
        self.btn_run.pack(side="left")

        self.btn_stop = tk.Button(btns, text="Stop", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(8, 0))

        self.lbl_status = tk.Label(btns, text="Idle.")
        self.lbl_status.pack(side="right")

        logfrm = tk.Frame(self)
        logfrm.pack(fill="both", expand=True, **pad)
        tk.Label(logfrm, text="Log:").pack(anchor="w")
        self.txt = tk.Text(logfrm, wrap="word", height=22)
        self.txt.pack(fill="both", expand=True)

    def _log(self, msg: str) -> None:
        self.txt.insert("end", msg.rstrip() + "\n")
        self.txt.see("end")
        self.update_idletasks()

    def _set_status(self, msg: str) -> None:
        self.lbl_status.config(text=msg)
        self.update_idletasks()

    def _browse_old(self) -> None:
        d = filedialog.askdirectory(title="Select OLD rebuild root folder")
        if d:
            self.var_old.set(d)

    def _browse_new(self) -> None:
        d = filedialog.askdirectory(title="Select NEW rebuild root folder")
        if d:
            self.var_new.set(d)

    def _browse_out(self) -> None:
        d = filedialog.askdirectory(title="Select output folder for reports")
        if d:
            # keep current folder name under chosen directory
            base_name = Path(self.var_out.get()).name or f"QMS_folder_diff_reports_{_now_stamp()}"
            self.var_out.set(str(Path(d) / base_name))

    def _stop(self) -> None:
        self.stop_event.set()
        self._log("Stop signal sent...")

    def _run(self) -> None:
        old_root = (self.var_old.get() or "").strip()
        new_root = (self.var_new.get() or "").strip()
        out_dir = (self.var_out.get() or "").strip()

        if not old_root or not new_root:
            messagebox.showerror("Missing input", "Please select both OLD and NEW rebuild root folders.")
            return

        oldp = Path(old_root)
        newp = Path(new_root)
        if not oldp.is_dir() or not newp.is_dir():
            messagebox.showerror("Invalid input", "One or both root paths are not valid directories.")
            return

        outp = Path(out_dir)
        if not outp.name:
            messagebox.showerror("Invalid output", "Please set a valid output report folder.")
            return

        self.stop_event.clear()
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")

        self._log("────────────────────────────────────────────────────────────")
        self._log(f"Run started: {datetime.now().isoformat(timespec='seconds')}")
        self._log(f"OLD root: {oldp.resolve()}")
        self._log(f"NEW root: {newp.resolve()}")
        self._log(f"Report out: {outp.resolve()}")
        self._set_status("Indexing...")

        def worker() -> None:
            try:
                old_inv = _build_inventory(oldp, logger=self._log, stop_event=self.stop_event)
                if self.stop_event.is_set():
                    self._log("Stopped during OLD inventory.")
                    return

                new_inv = _build_inventory(newp, logger=self._log, stop_event=self.stop_event)
                if self.stop_event.is_set():
                    self._log("Stopped during NEW inventory.")
                    return

                self._set_status("Diffing...")
                record_rows = _diff_records(old_inv, new_inv)

                grouped = _group_by_type(record_rows)
                summary_rows = _type_summary(record_rows, old_inv, new_inv)

                self._set_status("Writing reports...")
                outp.mkdir(parents=True, exist_ok=True)

                per_type_paths = _write_per_type_csvs(outp, grouped)
                per_type_expanded_paths = _write_per_type_new_items_expanded_csvs(outp, grouped)

                summary_path = _write_type_summary_csv(outp, summary_rows)
                all_expanded_path = _write_all_new_items_expanded_csv(outp, record_rows)

                json_path = None
                if self.var_write_json.get():
                    json_path = _write_full_json(outp, oldp, newp, old_inv, new_inv, record_rows, summary_rows)

                # Logging summary
                total_records = len(record_rows)
                changed_records = sum(1 for r in record_rows if r.num_new_items > 0)
                total_new_items = sum(r.num_new_items for r in record_rows)

                self._log("────────────────────────────────────────────────────────────")
                self._log(f"Records compared (union): {total_records}")
                self._log(f"Records with new items:   {changed_records}")
                self._log(f"Total new items added:    {total_new_items}")
                self._log("")

                self._log("Wrote per-type CSVs:")
                for p in per_type_paths:
                    self._log(f"  - {p.name}")

                self._log("Wrote per-type expanded new-item CSVs:")
                for p in per_type_expanded_paths:
                    self._log(f"  - {p.name}")

                self._log(f"Wrote summary CSV: {summary_path.name}")
                self._log(f"Wrote ALL expanded new-items CSV: {all_expanded_path.name}")
                if json_path:
                    self._log(f"Wrote full JSON:   {json_path.name}")

                self._log("Done.")
                self._set_status("Done.")
            except Exception as e:
                self._log("────────────────────────────────────────────────────────────")
                self._log(f"❌ ERROR: {type(e).__name__}: {e}")
                self._log(_traceback_str())
                self._set_status("Error.")
                messagebox.showerror("Error", "Diff failed. See log for details.")
            finally:
                self.after(0, self._finish)

        threading.Thread(target=worker, daemon=True).start()

    def _finish(self) -> None:
        self.btn_run.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.stop_event.clear()

    def destroy(self) -> None:
        try:
            self.stop_event.set()
        except Exception:
            pass
        super().destroy()


def _traceback_str() -> str:
    import traceback
    return traceback.format_exc()


def main() -> None:
    app = FolderDiffWidget()
    app.mainloop()


if __name__ == "__main__":
    main()