#!/usr/bin/env python3
"""
decode_review_justifications_gui.py

SpartaDMS / TrackWise Export Utility
Decode SPARTADMS__Review_Justification__c.csv to a human-friendly CSV by resolving:
- Users (CreatedById) via User.csv
- Corporate Document identity via SPARTADMS__Corporate_Document__c.csv
- Version context MUST come from SPARTADMS__Document_Version__c.csv (no current-state fallback)

Version resolution order (deterministic):
  1) Direct join: Justification.SPARTADMS__Document_Version__c -> Document_Version.Id
  2) CorporateDocId as-of: Document_Version.SPARTADMS__Related_Corporate_Document__c == CorporateDocId
     choose latest version_event_dt <= justification_created_dt
  3) Approved doc number as-of: Corporate_Document.SPARTADMS__Approved_Document_Number__c OR Full_Document_Number__c
     match Document_Version.SPARTADMS__Approved_Document_Number__c, then choose as-of
  4) Optional parse from justification text (guarded and conservative)

If no version can be resolved:
- Version_Being_Revised and Version_File_Name are left blank
- the row is logged in the Event Log (no unresolved CSV is generated)

GUI:
- Select input folder containing the CSVs
- Select output folder
- Click GO
- Live event log
- Writes decoded CSV to output folder

Output columns (ordered per updated request):
Justification_Number
Justification_Text
Document_Full_Number
Document_Title
Document_Type
Version_Being_Revised
Version_File_Name
Justification_Created_by_name
Created_By_Email
Justification_Created_Date
Quality_Event_Reference
"""

import os
import re
import sys
import csv
import queue
import threading
from datetime import datetime
from typing import Dict, Tuple, Optional, List, Any

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pandas as pd
except Exception:
    pd = None


# -----------------------------
# Configuration
# -----------------------------

REQUIRED_FILES = {
    "justification": "SPARTADMS__Review_Justification__c.csv",
    "doc_version": "SPARTADMS__Document_Version__c.csv",
    "corp_doc": "SPARTADMS__Corporate_Document__c.csv",
    "user": "User.csv",
}

OUTPUT_COLUMNS = [
    "Justification_Number",
    "Justification_Text",
    "Document_Full_Number",
    "Document_Title",
    "Document_Type",
    "Version_Being_Revised",
    "Version_File_Name",
    "Justification_Created_by_name",
    "Created_By_Email",
    "Justification_Created_Date",
    "Quality_Event_Reference",
]

SF_ID_LIKE = re.compile(r"^[a-zA-Z0-9]{15,18}$")

# Parse doc + version from text patterns like "SOP-1061.02" or "RTM-0050.01"
DOC_VER_IN_TEXT = re.compile(r"\b([A-Z]{2,6}-\d{1,6})\.(\d{1,3})\b")


# -----------------------------
# Logging
# -----------------------------

def log_put(q: queue.Queue, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    q.put(f"[{ts}] {msg}")


# -----------------------------
# CSV helpers
# -----------------------------

def find_file(folder: str, filename: str) -> Optional[str]:
    p = os.path.join(folder, filename)
    if os.path.isfile(p):
        return p
    lower = filename.lower()
    for f in os.listdir(folder):
        if f.lower() == lower:
            return os.path.join(folder, f)
    return None


def safe_read_csv(path: str, logq: queue.Queue, expected_min_cols: int = 2) -> "pd.DataFrame":
    """
    Robust CSV reader:
    - reads all columns as str
    - tolerates embedded JSON blobs in fields (common in some SpartaDMS exports)
    - warns on bad lines instead of crashing
    """
    if pd is None:
        raise RuntimeError("pandas is required for this widget, but could not be imported.")

    log_put(logq, f"Loading CSV: {os.path.basename(path)}")

    last_err = None
    for engine in ("python", "c"):
        try:
            df = pd.read_csv(
                path,
                dtype=str,
                keep_default_na=False,
                na_filter=False,
                engine=engine,
                on_bad_lines="warn",
                quoting=csv.QUOTE_MINIMAL,
            )
            if df.shape[1] < expected_min_cols:
                raise ValueError(f"Parsed too few columns ({df.shape[1]}).")
            log_put(logq, f"Loaded {os.path.basename(path)}: {len(df):,} rows, {df.shape[1]:,} cols (engine={engine}).")
            return df
        except Exception as e:
            last_err = e
            log_put(logq, f"Read attempt failed (engine={engine}): {e}")

    raise RuntimeError(f"Failed to read CSV {path}: {last_err}")


def pick_first_present(row: Dict[str, str], candidates) -> str:
    for c in candidates:
        v = row.get(c, "")
        if v is not None and str(v).strip() != "":
            return str(v).strip()
    return ""


def is_id_like(val: str) -> bool:
    v = (val or "").strip()
    return bool(v) and bool(SF_ID_LIKE.match(v))


def parse_dt(s: str) -> Optional[datetime]:
    """
    Parse common Salesforce export datetime formats:
    - M/D/YYYY H:MM
    - M/D/YYYY H:MM:SS
    - ISO: 2020-07-23T12:38:48.000+0000 / ...Z
    Returns naive datetime.
    """
    s = (s or "").strip()
    if not s:
        return None

    for fmt in ("%m/%d/%Y %H:%M", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %I:%M:%S %p"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass

    try:
        s2 = re.sub(r"([+-]\d{4})$", "", s)
        s2 = s2[:-1] if s2.endswith("Z") else s2
        if "." in s2:
            s2 = s2.split(".", 1)[0]
        s2 = s2.replace("T", " ")
        return datetime.strptime(s2, "%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


# -----------------------------
# Version row helpers
# -----------------------------

def version_event_dt(v: Dict[str, str]) -> Optional[datetime]:
    """
    Timestamp used to order versions in time.
    Priority:
      Approved_Date_Time__c
      Effective_Datetime__c / Effective_DT__c
      LastModifiedDate / SystemModstamp
    """
    for c in ("SPARTADMS__Approved_Date_Time__c", "SPARTADMS__Effective_Datetime__c", "SPARTADMS__Effective_DT__c", "LastModifiedDate", "SystemModstamp"):
        d = parse_dt(v.get(c, ""))
        if d:
            return d
    return None


def make_version_label(v: Dict[str, str]) -> str:
    num = (v.get("SPARTADMS__Document_Version__c", "") or "").strip()
    fmt = (v.get("SPARTADMS__Document_Version_Format__c", "") or "").strip()
    if num and fmt:
        return f"{num} ({fmt})"
    return num


def make_version_file(v: Dict[str, str]) -> str:
    return pick_first_present(v, ["SPARTADMS__Latest_file_name__c", "SPARTADMS__File_Name__c"])


# -----------------------------
# Map builders
# -----------------------------

def build_user_map(user_df: "pd.DataFrame", logq: queue.Queue) -> Dict[str, Tuple[str, str]]:
    needed = ["Id", "FirstName", "LastName"]
    missing = [c for c in needed if c not in user_df.columns]
    if missing:
        raise RuntimeError(f"User.csv missing required columns: {missing}")

    has_email = "Email" in user_df.columns
    has_username = "Username" in user_df.columns

    m: Dict[str, Tuple[str, str]] = {}
    for _, r in user_df.iterrows():
        uid = str(r.get("Id", "")).strip()
        if not uid:
            continue
        first = str(r.get("FirstName", "")).strip()
        last = str(r.get("LastName", "")).strip()
        name = (" ".join([x for x in [first, last] if x])).strip() or uid
        email = ""
        if has_email:
            email = str(r.get("Email", "")).strip()
        if not email and has_username:
            email = str(r.get("Username", "")).strip()
        m[uid] = (name, email)

    log_put(logq, f"User map built: {len(m):,} users.")
    return m


def build_corp_doc_map(cd_df: "pd.DataFrame", logq: queue.Queue) -> Dict[str, Dict[str, str]]:
    if "Id" not in cd_df.columns:
        raise RuntimeError("SPARTADMS__Corporate_Document__c.csv missing 'Id' column.")

    m: Dict[str, Dict[str, str]] = {}
    for _, r in cd_df.iterrows():
        cid = str(r.get("Id", "")).strip()
        if not cid:
            continue
        v = {c: str(r.get(c, "")).strip() for c in cd_df.columns}
        m[cid] = {
            "Document_Full_Number": v.get("SPARTADMS__Full_Document_Number__c", ""),
            "Document_Title": v.get("Name", ""),
            "Document_Type": v.get("SPARTADMS__Document_Type__c", ""),
            "Quality_Event_Reference": v.get("SPARTADMS__Quality_Event_Reference__c", ""),
            # For matching strategy #3:
            "_cd_approved_doc_no": v.get("SPARTADMS__Approved_Document_Number__c", ""),
        }

    log_put(logq, f"Corporate Document map built: {len(m):,} corporate docs.")
    return m


def build_doc_version_indexes(ver_df: "pd.DataFrame", logq: queue.Queue) -> Tuple[
    Dict[str, Dict[str, str]],                 # by Id
    Dict[str, List[Dict[str, str]]],           # by corporate doc id
    Dict[str, List[Dict[str, str]]],           # by approved doc number
    Dict[Tuple[str, str], List[Dict[str, str]]]# by (doc_number, version_num)
]:
    if "Id" not in ver_df.columns:
        raise RuntimeError("SPARTADMS__Document_Version__c.csv missing 'Id' column.")

    by_id: Dict[str, Dict[str, str]] = {}
    by_corp: Dict[str, List[Dict[str, str]]] = {}
    by_approved: Dict[str, List[Dict[str, str]]] = {}
    by_doc_and_ver: Dict[Tuple[str, str], List[Dict[str, str]]] = {}

    cols = list(ver_df.columns)
    for _, r in ver_df.iterrows():
        v = {c: str(r.get(c, "")).strip() for c in cols}
        vid = v.get("Id", "")
        if not vid:
            continue
        by_id[vid] = v

        corp_id = (v.get("SPARTADMS__Related_Corporate_Document__c", "") or "").strip()
        if corp_id:
            by_corp.setdefault(corp_id, []).append(v)

        appr = (v.get("SPARTADMS__Approved_Document_Number__c", "") or "").strip()
        if appr:
            by_approved.setdefault(appr, []).append(v)

        docnum = (v.get("SPARTADMS__Document_Number__c", "") or "").strip()
        vernum = (v.get("SPARTADMS__Document_Version__c", "") or "").strip()
        if docnum and vernum:
            by_doc_and_ver.setdefault((docnum, vernum), []).append(v)

    def sort_key(vrow: Dict[str, str]) -> Any:
        d = version_event_dt(vrow)
        return (d is None, d)

    for dct in (by_corp, by_approved, by_doc_and_ver):
        for k in list(dct.keys()):
            dct[k].sort(key=sort_key)

    log_put(logq, f"Document Version indexes built: by_id={len(by_id):,}, by_corp={len(by_corp):,}, by_approved={len(by_approved):,}")
    return by_id, by_corp, by_approved, by_doc_and_ver


def resolve_user(user_map: Dict[str, Tuple[str, str]], user_id: str) -> Tuple[str, str]:
    user_id = (user_id or "").strip()
    if not user_id:
        return ("", "")
    return user_map.get(user_id, ("", ""))


def choose_asof(versions: List[Dict[str, str]], asof: Optional[datetime]) -> Optional[Dict[str, str]]:
    """
    Choose the latest version whose event_dt <= asof.
    If asof is None, choose the latest version with a non-null event_dt, else None.
    Assumes versions list is pre-sorted by event_dt ascending (None last).
    """
    if not versions:
        return None

    if asof is None:
        for v in reversed(versions):
            if version_event_dt(v) is not None:
                return v
        return None

    best = None
    for v in versions:
        d = version_event_dt(v)
        if d is None:
            continue
        if d <= asof:
            best = v
        else:
            break
    return best


def parse_docver_from_text(text: str) -> Optional[Tuple[str, str]]:
    """
    Extract (doc_code_like, version_num) from text e.g. "SOP-1061.02" -> ("SOP-1061", "02")
    """
    text = (text or "").strip()
    if not text:
        return None
    m = DOC_VER_IN_TEXT.search(text)
    if not m:
        return None
    doc = m.group(1)
    ver = m.group(2).lstrip("0") or "0"
    return doc, ver


# -----------------------------
# Decode core
# -----------------------------

def decode_justifications(input_folder: str, output_folder: str, logq: queue.Queue) -> str:
    if not os.path.isdir(output_folder):
        raise RuntimeError(f"Output folder does not exist: {output_folder}")

    # Locate files
    paths = {}
    log_put(logq, f"Scanning input folder: {input_folder}")
    for k, fn in REQUIRED_FILES.items():
        p = find_file(input_folder, fn)
        if not p:
            raise RuntimeError(f"Missing required file in input folder: {fn}")
        paths[k] = p
        log_put(logq, f"Found {fn}")

    # Load CSVs
    user_df = safe_read_csv(paths["user"], logq, expected_min_cols=5)
    ver_df = safe_read_csv(paths["doc_version"], logq, expected_min_cols=10)
    corp_df = safe_read_csv(paths["corp_doc"], logq, expected_min_cols=10)
    just_df = safe_read_csv(paths["justification"], logq, expected_min_cols=10)

    # Build maps/indexes
    user_map = build_user_map(user_df, logq)
    corp_map = build_corp_doc_map(corp_df, logq)
    ver_by_id, ver_by_corp, ver_by_approved, ver_by_doc_and_ver = build_doc_version_indexes(ver_df, logq)

    if "Name" not in just_df.columns:
        raise RuntimeError("Justification CSV missing 'Name' column.")

    justification_text_col = "SPARTADMS__Justification__c" if "SPARTADMS__Justification__c" in just_df.columns else None
    if justification_text_col:
        log_put(logq, f"Using justification narrative column: {justification_text_col}")
    else:
        log_put(logq, "NOTE: SPARTADMS__Justification__c not present; Justification_Text will be blank.")

    doc_ver_fk_col = "SPARTADMS__Document_Version__c" if "SPARTADMS__Document_Version__c" in just_df.columns else None
    if doc_ver_fk_col:
        log_put(logq, f"Using version FK column: {doc_ver_fk_col}")
    else:
        log_put(logq, "WARNING: No SPARTADMS__Document_Version__c column in justification CSV.")

    corp_fk_col = "SPARTADMS__Related_Corporate_Document__c" if "SPARTADMS__Related_Corporate_Document__c" in just_df.columns else None
    if corp_fk_col:
        log_put(logq, f"Using corporate FK column: {corp_fk_col}")
    else:
        log_put(logq, "WARNING: No SPARTADMS__Related_Corporate_Document__c column in justification CSV.")

    out_rows: List[Dict[str, str]] = []

    counts = {"direct_fk": 0, "asof_corp": 0, "asof_approved": 0, "text_match": 0, "unresolved": 0}
    unresolved_examples: List[str] = []

    log_put(logq, "Decoding justification rows...")

    for idx, r in just_df.iterrows():
        j = {c: str(r.get(c, "")).strip() for c in just_df.columns}

        jnum = j.get("Name", "").strip()
        if not jnum:
            continue

        jtext = j.get(justification_text_col, "").strip() if justification_text_col else ""
        j_created_str = j.get("CreatedDate", "").strip()
        j_created_dt = parse_dt(j_created_str)

        created_by_id = j.get("CreatedById", "").strip()
        created_by_name, created_by_email = resolve_user(user_map, created_by_id)

        # Corporate doc resolve
        corp_id = (j.get(corp_fk_col, "") if corp_fk_col else "").strip()
        cinfo = corp_map.get(
            corp_id,
            {"Document_Full_Number": "", "Document_Title": "", "Document_Type": "", "Quality_Event_Reference": "", "_cd_approved_doc_no": ""},
        )

        # Version resolve
        chosen_v: Optional[Dict[str, str]] = None

        # (1) direct FK
        if doc_ver_fk_col:
            vid = (j.get(doc_ver_fk_col, "") or "").strip()
            if vid and vid in ver_by_id:
                chosen_v = ver_by_id[vid]
                counts["direct_fk"] += 1

        # (2) corp-id as-of
        if chosen_v is None and corp_id:
            vlist = ver_by_corp.get(corp_id, [])
            cand = choose_asof(vlist, j_created_dt)
            if cand is not None:
                chosen_v = cand
                counts["asof_corp"] += 1

        # (3) approved doc number as-of
        if chosen_v is None:
            appr = (cinfo.get("_cd_approved_doc_no", "") or "").strip()
            full = (cinfo.get("Document_Full_Number", "") or "").strip()
            key = appr or full
            if key:
                vlist = ver_by_approved.get(key, [])
                cand = choose_asof(vlist, j_created_dt)
                if cand is not None:
                    chosen_v = cand
                    counts["asof_approved"] += 1

        # (4) optional text parse match (conservative)
        if chosen_v is None and jtext:
            parsed = parse_docver_from_text(jtext)
            if parsed:
                doc_code_like, vernum = parsed
                # Conservative: only attempt if doc_code_like appears in Document_Number__c
                candidates = []
                for (docnum, vn), rows in ver_by_doc_and_ver.items():
                    if vn == vernum and doc_code_like in docnum:
                        candidates.extend(rows)
                if candidates:
                    candidates.sort(key=lambda vr: (version_event_dt(vr) is None, version_event_dt(vr)))
                    cand = choose_asof(candidates, j_created_dt)
                    if cand is not None:
                        chosen_v = cand
                        counts["text_match"] += 1

        # Output version fields (only from Document_Version)
        if chosen_v is not None:
            v_being_revised = make_version_label(chosen_v)
            v_file = make_version_file(chosen_v)
        else:
            v_being_revised = ""
            v_file = ""
            counts["unresolved"] += 1
            if len(unresolved_examples) < 25:
                vid_raw = (j.get(doc_ver_fk_col, "") if doc_ver_fk_col else "").strip()
                unresolved_examples.append(
                    f"{jnum} | Created={j_created_str} | CorpId={corp_id or '(blank)'} | FullNo={cinfo.get('Document_Full_Number','')} | ApprovedNo={cinfo.get('_cd_approved_doc_no','')} | VerFK={vid_raw or '(blank)'}"
                )

        out_rows.append({
            "Justification_Number": jnum,
            "Justification_Text": jtext,
            "Document_Full_Number": cinfo.get("Document_Full_Number", ""),
            "Document_Title": cinfo.get("Document_Title", ""),
            "Document_Type": cinfo.get("Document_Type", ""),
            "Version_Being_Revised": v_being_revised,
            "Version_File_Name": v_file,
            "Justification_Created_by_name": created_by_name,
            "Created_By_Email": created_by_email,
            "Justification_Created_Date": j_created_str,
            "Quality_Event_Reference": cinfo.get("Quality_Event_Reference", ""),
        })

        if (idx + 1) % 2000 == 0:
            log_put(logq, f"Decoded {idx + 1:,} rows...")

    log_put(logq, f"Decoded rows: {len(out_rows):,}")
    log_put(
        logq,
        "Version match methods: "
        f"direct_fk={counts['direct_fk']:,}, "
        f"asof_corp={counts['asof_corp']:,}, "
        f"asof_approved={counts['asof_approved']:,}, "
        f"text_match={counts['text_match']:,}, "
        f"unresolved={counts['unresolved']:,}"
    )

    if counts["unresolved"] > 0:
        log_put(logq, "WARNING: Some justifications could not be matched to a specific Document_Version row. No fallback was used; version fields were left blank.")
        log_put(logq, "First unresolved examples (up to 25):")
        for ex in unresolved_examples:
            log_put(logq, f"  - {ex}")

    # Write output CSV
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"decoded_review_justifications_{ts}.csv"
    out_path = os.path.join(output_folder, out_name)

    log_put(logq, f"Writing output CSV: {out_path}")
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=OUTPUT_COLUMNS)
        w.writeheader()
        for row in out_rows:
            w.writerow({k: row.get(k, "") for k in OUTPUT_COLUMNS})

    log_put(logq, "Done.")
    return out_path


# -----------------------------
# GUI
# -----------------------------

class DecodeJustificationsGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SpartaDMS - Decode Review Justifications")
        self.geometry("1100x700")

        self.logq: queue.Queue = queue.Queue()
        self.worker_thread: Optional[threading.Thread] = None

        self._build_ui()
        self._poll_log_queue()

    def _build_ui(self):
        pad = 8

        container = ttk.Frame(self)
        container.pack(fill="x", padx=pad, pady=(pad, 0))

        # Input folder
        row1 = ttk.Frame(container)
        row1.pack(fill="x", pady=(0, 6))
        ttk.Label(row1, text="Input Folder (CSV export):", width=22).pack(side="left")
        self.input_folder_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.input_folder_var, width=90).pack(side="left", padx=(0, pad))
        ttk.Button(row1, text="Browse...", command=self._browse_input_folder).pack(side="left")

        # Output folder
        row2 = ttk.Frame(container)
        row2.pack(fill="x", pady=(0, 6))
        ttk.Label(row2, text="Output Folder:", width=22).pack(side="left")
        self.output_folder_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.output_folder_var, width=90).pack(side="left", padx=(0, pad))
        ttk.Button(row2, text="Browse...", command=self._browse_output_folder).pack(side="left")

        # Go row
        row3 = ttk.Frame(container)
        row3.pack(fill="x", pady=(0, 6))
        self.go_btn = ttk.Button(row3, text="GO", command=self._go)
        self.go_btn.pack(side="left")

        # Event log
        mid = ttk.Frame(self)
        mid.pack(fill="both", expand=True, padx=pad, pady=pad)

        log_frame = ttk.Labelframe(mid, text="Event Log")
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, wrap="word", height=28, state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True)

        yscroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        yscroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=yscroll.set)

        # Status bar
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=pad, pady=(0, pad))
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")

    def _browse_input_folder(self):
        d = filedialog.askdirectory(title="Select TrackWise/SpartaDMS Export Folder (Input)")
        if d:
            self.input_folder_var.set(d)
            # IMPORTANT: do NOT auto-fill output folder here (per request)

    def _browse_output_folder(self):
        d = filedialog.askdirectory(title="Select Output Folder")
        if d:
            self.output_folder_var.set(d)

    def _append_log(self, line: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _poll_log_queue(self):
        try:
            while True:
                line = self.logq.get_nowait()
                self._append_log(line)
        except queue.Empty:
            pass
        self.after(150, self._poll_log_queue)

    def _go(self):
        input_folder = self.input_folder_var.get().strip()
        output_folder = self.output_folder_var.get().strip()

        if not input_folder or not os.path.isdir(input_folder):
            messagebox.showerror("Error", "Please select a valid input folder.")
            return

        # If output folder blank, default to input folder at runtime (not on browse)
        if not output_folder:
            output_folder = input_folder
            self.output_folder_var.set(output_folder)

        if not os.path.isdir(output_folder):
            messagebox.showerror("Error", "Please select a valid output folder.")
            return

        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Busy", "A decode job is already running.")
            return

        self.go_btn.configure(state="disabled")
        self.status_var.set("Running...")
        log_put(self.logq, "=== Decode job started ===")

        def worker():
            try:
                out_path = decode_justifications(input_folder, output_folder, self.logq)
                log_put(self.logq, "=== Decode job completed successfully ===")
                self._on_worker_done(True, out_path, None)
            except Exception as e:
                log_put(self.logq, f"ERROR: {e}")
                log_put(self.logq, "=== Decode job failed ===")
                self._on_worker_done(False, None, str(e))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def _on_worker_done(self, success: bool, out_path: Optional[str], err: Optional[str]):
        def ui_update():
            self.go_btn.configure(state="normal")
            self.status_var.set("Ready." if success else "Failed.")
            if success and out_path:
                messagebox.showinfo("Complete", f"Decoded CSV written:\n{out_path}\n\nIf some version fields are blank, check the Event Log for unresolved version matching details.")
            elif not success and err:
                messagebox.showerror("Failed", err)
        self.after(0, ui_update)


if __name__ == "__main__":
    if pd is None:
        print("ERROR: pandas is required but not available.", file=sys.stderr)
        sys.exit(1)

    app = DecodeJustificationsGUI()
    app.mainloop()
