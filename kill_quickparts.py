#!/usr/bin/env python3
"""
quickparts_batch_cleaner_gui.py

Batch-clean Word documents in a folder:
- Removes "Quick Parts" implemented as DOCPROPERTY complex fields (BEGIN..END field group)
- Removes literal placeholder text "[Leave Blank]" in headers/footers (leaving structure intact for free-typing)
- Processes ALL .docx and .docm in the input folder
- If input is .docm, outputs a .docx (macro-free) by removing VBA parts and adjusting content-types + rels
- Outputs processed docs into a RUN FOLDER created under the selected Output folder
  (prevents accidental overwrite and provides a single place for run artifacts)

NEW:
- Detects and removes Word "Restrict Editing" / document protection if not password-enforced.
  If password-enforced (hash/salt present), it will NOT attempt to bypass; it logs a warning.

UI:
- Simple Tkinter GUI
- Visible event log (scrolling) under the fields, plus a saved run log file in the run folder

Dependencies:
  pip install lxml

Run:
  python quickparts_batch_cleaner_gui.py
"""

from __future__ import annotations

import os
import re
import sys
import queue
import zipfile
import threading
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import ctypes
from ctypes import wintypes

from lxml import etree as ET


# -----------------------------
# WordprocessingML constants
# -----------------------------
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": NS_W}

CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

DOCX_MAIN_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"

VBA_REL_TYPE_SUBSTR = "vbaProject"
VBA_PARTS = {
    "word/vbaProject.bin",
    "word/vbaData.xml",
}

SETTINGS_PART = "word/settings.xml"


def qn_w(tag: str) -> str:
    """Qualified name helper for w: tags."""
    prefix, local = tag.split(":")
    if prefix != "w":
        raise ValueError(f"Unsupported prefix: {prefix}")
    return f"{{{NS_W}}}{local}"


def now_tag() -> str:
    return datetime.now().strftime("%d%b%Y_%H%M%S").upper()


def safe_mkdir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


# -----------------------------
# Reliable folder picker (Windows)
# -----------------------------
def pick_folder(title: str, owner_hwnd: Optional[int] = None) -> Optional[str]:
    """
    Folder picker that behaves like a folder picker (not a file dialog).

    - On Windows: uses SHBrowseForFolderW (native folder picker). We explicitly CoInitialize/CoUninitialize
      and pass an owner window handle so the dialog actually appears/focuses.
    - Else: falls back to tkinter askdirectory.
    """
    if os.name != "nt":
        return filedialog.askdirectory(title=title)

    try:
        shell32 = ctypes.windll.shell32
        ole32 = ctypes.windll.ole32
        user32 = ctypes.windll.user32

        ole32.CoInitialize(None)

        BIF_RETURNONLYFSDIRS = 0x0001
        BIF_NEWDIALOGSTYLE = 0x0040
        MAX_PATH = 32768

        class BROWSEINFO(ctypes.Structure):
            _fields_ = [
                ("hwndOwner", wintypes.HWND),
                ("pidlRoot", ctypes.c_void_p),
                ("pszDisplayName", wintypes.LPWSTR),
                ("lpszTitle", wintypes.LPCWSTR),
                ("ulFlags", wintypes.UINT),
                ("lpfn", ctypes.c_void_p),
                ("lParam", wintypes.LPARAM),
                ("iImage", ctypes.c_int),
            ]

        hwnd = owner_hwnd if owner_hwnd else user32.GetForegroundWindow()

        display_buf = ctypes.create_unicode_buffer(MAX_PATH)

        bi = BROWSEINFO()
        bi.hwndOwner = wintypes.HWND(hwnd)
        bi.pidlRoot = None
        bi.pszDisplayName = display_buf
        bi.lpszTitle = title
        bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE
        bi.lpfn = None
        bi.lParam = 0
        bi.iImage = 0

        pidl = shell32.SHBrowseForFolderW(ctypes.byref(bi))
        if not pidl:
            return None

        path_buf = ctypes.create_unicode_buffer(MAX_PATH)
        ok = shell32.SHGetPathFromIDListW(pidl, path_buf)

        ole32.CoTaskMemFree(pidl)

        if ok:
            path = path_buf.value
            return path if path else None
        return None

    except Exception:
        return filedialog.askdirectory(title=title)

    finally:
        try:
            ctypes.windll.ole32.CoUninitialize()
        except Exception:
            pass


# -----------------------------
# Core DOCX/DOCM processing
# -----------------------------
def iter_relevant_word_xml_parts(zip_in: zipfile.ZipFile) -> Iterable[str]:
    """
    Yield XML parts that commonly contain WordprocessingML content where fields appear,
    plus settings.xml for protection logic.
    """
    for name in zip_in.namelist():
        if name == SETTINGS_PART:
            yield name
            continue

        if not name.startswith("word/"):
            continue
        if not name.endswith(".xml"):
            continue
        if (
            name == "word/document.xml"
            or name.startswith("word/header")
            or name.startswith("word/footer")
            or name == "word/footnotes.xml"
            or name == "word/endnotes.xml"
        ):
            yield name


def remove_docproperty_fields(root: ET._Element, keep_page_fields: bool = True) -> int:
    removed = 0

    fldCharType = qn_w("w:fldCharType")
    r_tag = qn_w("w:r")
    t_tag = qn_w("w:t")

    for p in root.findall(".//w:p", NS):
        runs = list(p.findall("./w:r", NS))
        i = 0
        while i < len(runs):
            r = runs[i]
            fc = r.find("./w:fldChar", NS)
            if fc is None or fc.get(fldCharType) != "begin":
                i += 1
                continue

            j = i + 1
            instr_parts = []
            seen_separate = False

            while j < len(runs):
                rj = runs[j]

                if not seen_separate:
                    it = rj.find("./w:instrText", NS)
                    if it is not None and it.text:
                        instr_parts.append(it.text)

                fcj = rj.find("./w:fldChar", NS)
                if fcj is not None:
                    ftype = fcj.get(fldCharType)
                    if ftype == "separate":
                        seen_separate = True
                    elif ftype == "end":
                        break
                j += 1

            if j >= len(runs):
                i += 1
                continue

            instr = "".join(instr_parts)
            instr_norm = re.sub(r"\s+", " ", instr).strip().upper()

            remove = "DOCPROPERTY" in instr_norm

            if keep_page_fields and (
                " PAGE " in instr_norm
                or " NUMPAGES" in instr_norm
                or instr_norm.startswith("PAGE")
                or instr_norm.startswith("NUMPAGES")
            ):
                remove = False

            if remove:
                for k in range(i, j + 1):
                    p.remove(runs[k])

                blank_r = ET.Element(r_tag)
                blank_t = ET.SubElement(blank_r, t_tag)
                blank_t.text = ""
                p.insert(min(i, len(p)), blank_r)

                removed += 1
                runs = list(p.findall("./w:r", NS))
                i = min(i + 1, len(runs))
            else:
                i += 1

    return removed


def remove_leave_blank_text(root: ET._Element) -> int:
    cleared = 0
    for t in root.findall(".//w:t", NS):
        if t.text is None:
            continue
        stripped = t.text.strip()
        if stripped in ("[Leave Blank]", "Leave Blank", "[LEAVE BLANK]", "LEAVE BLANK"):
            t.text = ""
            cleared += 1
        elif re.fullmatch(r"\[\s*Leave Blank\s*\]", stripped, flags=re.IGNORECASE):
            t.text = ""
            cleared += 1
    return cleared


def _looks_password_enforced(elem: ET._Element) -> bool:
    """
    Heuristic: if protection element carries hash/salt/crypto-related attributes,
    Word likely requires a password to disable via UI.
    """
    for k, v in elem.attrib.items():
        lk = k.lower()
        if "hash" in lk or "salt" in lk or "crypt" in lk or "algorithm" in lk:
            if v:
                return True
    return False


def remove_document_protection_if_safe(settings_xml: bytes) -> Tuple[bytes, bool, bool]:
    """
    Returns (new_bytes, protection_found, protection_removed)

    - If protection is present and NOT password-enforced (no hash/salt/crypto attrs): remove it.
    - If password-enforced: leave intact (do not bypass), but report found=True, removed=False.
    """
    root = ET.fromstring(settings_xml)

    prot_tags = [qn_w("w:documentProtection"), qn_w("w:writeProtection")]
    found = False
    removed = False

    for tag in prot_tags:
        for elem in list(root.findall(f".//{tag}")):
            found = True
            if _looks_password_enforced(elem):
                # Do not attempt to bypass password enforcement.
                continue
            # Not password-enforced -> remove
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)
                removed = True

    out = ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    return out, found, removed


def process_word_xml_part(xml_bytes: bytes, part_name: str) -> Tuple[bytes, int, int, bool, bool]:
    """
    Returns:
      (new_xml_bytes,
       docproperty_fields_removed,
       leave_blank_cleared,
       protection_found,
       protection_removed)
    """
    # Special-case settings.xml for protection removal
    if part_name == SETTINGS_PART:
        new_bytes, found, removed = remove_document_protection_if_safe(xml_bytes)
        return new_bytes, 0, 0, found, removed

    root = ET.fromstring(xml_bytes)

    removed_fields = remove_docproperty_fields(root, keep_page_fields=True)

    cleared_blanks = 0
    if part_name.startswith("word/header") or part_name.startswith("word/footer"):
        cleared_blanks = remove_leave_blank_text(root)

    out = ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    return out, removed_fields, cleared_blanks, False, False


def strip_macros_if_docm(file_map: dict[str, bytes], logger) -> None:
    for vba_part in list(VBA_PARTS):
        if vba_part in file_map:
            del file_map[vba_part]
            logger(f"  - Removed macro part: {vba_part}")

    ct_name = "[Content_Types].xml"
    if ct_name in file_map:
        try:
            ct_root = ET.fromstring(file_map[ct_name])
            ns = {"ct": CONTENT_TYPES_NS}

            removed_any = False
            for ov in list(ct_root.findall("ct:Override", ns)):
                part_name = ov.get("PartName", "")
                if part_name in ("/word/vbaProject.bin", "/word/vbaData.xml"):
                    ct_root.remove(ov)
                    removed_any = True
            if removed_any:
                logger("  - Updated [Content_Types].xml: removed VBA overrides")

            changed_main = False
            for ov in ct_root.findall("ct:Override", ns):
                if ov.get("PartName", "") == "/word/document.xml":
                    if ov.get("ContentType", "") != DOCX_MAIN_CT:
                        ov.set("ContentType", DOCX_MAIN_CT)
                        changed_main = True
            if changed_main:
                logger("  - Updated [Content_Types].xml: set document main type to DOCX")

            file_map[ct_name] = ET.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        except Exception:
            logger("  ! Warning: could not parse/modify [Content_Types].xml (left as-is)")

    rels_name = "word/_rels/document.xml.rels"
    if rels_name in file_map:
        try:
            rel_root = ET.fromstring(file_map[rels_name])
            rel_tag = f"{{{REL_NS}}}Relationship"
            removed_rels = 0
            for rel in list(rel_root.findall(rel_tag)):
                rel_type = rel.get("Type", "")
                if VBA_REL_TYPE_SUBSTR in rel_type:
                    rel_root.remove(rel)
                    removed_rels += 1
            if removed_rels:
                logger(f"  - Updated document.xml.rels: removed {removed_rels} VBA relationship(s)")
            file_map[rels_name] = ET.tostring(rel_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        except Exception:
            logger("  ! Warning: could not parse/modify document.xml.rels (left as-is)")


def clean_docx_or_docm_to_folder(in_path: Path, out_path: Path, logger) -> None:
    logger(f"Processing: {in_path.name}")

    with zipfile.ZipFile(in_path, "r") as zin:
        file_map: dict[str, bytes] = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        relevant_parts = set(iter_relevant_word_xml_parts(zin))

    total_removed_fields = 0
    total_cleared_blanks = 0
    protection_found_any = False
    protection_removed_any = False
    protection_password_enforced_suspected = False

    for part_name in list(file_map.keys()):
        if part_name in relevant_parts:
            try:
                new_xml, removed_fields, cleared_blanks, prot_found, prot_removed = process_word_xml_part(
                    file_map[part_name], part_name
                )
                file_map[part_name] = new_xml

                total_removed_fields += removed_fields
                total_cleared_blanks += cleared_blanks

                if prot_found:
                    protection_found_any = True
                    if prot_removed:
                        protection_removed_any = True
                    else:
                        # Found but not removed implies likely password-enforced (per our heuristic)
                        protection_password_enforced_suspected = True

            except Exception:
                logger(f"  ! Warning: failed to process {part_name} (left as-is)")

    logger(f"  - DOCPROPERTY fields removed: {total_removed_fields}")
    logger(f"  - Header/Footer '[Leave Blank]' cleared: {total_cleared_blanks}")

    if protection_found_any:
        if protection_removed_any:
            logger("  - Document protection detected in settings.xml: REMOVED (not password-enforced)")
        elif protection_password_enforced_suspected:
            logger("  ! Document protection detected: appears password-enforced; NOT removed (remove via Word with password)")
        else:
            logger("  - Document protection detected: not removed (unknown reason)")

    if in_path.suffix.lower() == ".docm":
        logger("  - Input is DOCM: converting to DOCX (macro-free)")
        strip_macros_if_docm(file_map, logger)

    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in file_map.items():
            if out_path.suffix.lower() == ".docx" and name in VBA_PARTS:
                continue
            zout.writestr(name, data)

    logger(f"  - Wrote: {out_path.name}")


def list_word_files(input_dir: Path, recursive: bool) -> list[Path]:
    exts = {".docx", ".docm"}
    if recursive:
        return sorted([p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in exts])
    return sorted([p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in exts])


# -----------------------------
# Tkinter GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DOCX Quick Parts + Leave Blank Cleaner (Batch)")
        self.resizable(True, True)
        self.minsize(820, 420)

        self.in_dir_var = tk.StringVar()
        self.out_dir_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=False)

        self._log_q: "queue.Queue[str]" = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._stop_flag = threading.Event()

        self._build_ui()
        self.after(100, self._drain_log_queue)

    def _build_ui(self):
        root = ttk.Frame(self, padding=10)
        root.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(4, weight=1)

        ttk.Label(root, text="Input Folder:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 6))
        ttk.Entry(root, textvariable=self.in_dir_var).grid(row=0, column=1, sticky="ew", pady=(0, 6))
        ttk.Button(root, text="Browse…", command=self._browse_input).grid(row=0, column=2, padx=(8, 0), pady=(0, 6))

        ttk.Label(root, text="Output Folder:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(0, 6))
        ttk.Entry(root, textvariable=self.out_dir_var).grid(row=1, column=1, sticky="ew", pady=(0, 6))
        ttk.Button(root, text="Browse…", command=self._browse_output).grid(row=1, column=2, padx=(8, 0), pady=(0, 6))

        opts = ttk.Frame(root)
        opts.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        opts.columnconfigure(0, weight=1)

        ttk.Checkbutton(opts, text="Include subfolders", variable=self.recursive_var).grid(row=0, column=0, sticky="w")

        btns = ttk.Frame(opts)
        btns.grid(row=0, column=1, sticky="e")
        self.run_btn = ttk.Button(btns, text="Run", command=self._run)
        self.run_btn.grid(row=0, column=0, padx=(0, 8))
        self.stop_btn = ttk.Button(btns, text="Stop", command=self._stop, state="disabled")
        self.stop_btn.grid(row=0, column=1)

        ttk.Separator(root).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 8))

        ttk.Label(root, text="Event Log:").grid(row=4, column=0, sticky="nw", padx=(0, 8))
        self.log_box = ScrolledText(root, height=14, wrap="word")
        self.log_box.grid(row=4, column=1, columnspan=2, sticky="nsew")
        self.log_box.configure(state="disabled")

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(root, textvariable=self.status_var).grid(row=5, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _browse_input(self):
        d = pick_folder("Select input folder", owner_hwnd=int(self.winfo_id()))
        if d:
            self.in_dir_var.set(d)
            self._log(f"Input folder set to: {d}")

    def _browse_output(self):
        d = pick_folder("Select output folder", owner_hwnd=int(self.winfo_id()))
        if d:
            self.out_dir_var.set(d)
            self._log(f"Output folder set to: {d}")

    def _log(self, msg: str):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._log_q.put(f"[{ts}] {msg}")

    def _drain_log_queue(self):
        try:
            while True:
                line = self._log_q.get_nowait()
                self.log_box.configure(state="normal")
                self.log_box.insert("end", line + "\n")
                self.log_box.see("end")
                self.log_box.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._drain_log_queue)

    def _set_running(self, running: bool):
        self.run_btn.configure(state="disabled" if running else "normal")
        self.stop_btn.configure(state="normal" if running else "disabled")

    def _stop(self):
        self._stop_flag.set()
        self._log("Stop requested; finishing current file then exiting.")
        self.status_var.set("Stop requested…")

    def _run(self):
        in_dir = Path(self.in_dir_var.get().strip())
        out_dir = Path(self.out_dir_var.get().strip())
        recursive = bool(self.recursive_var.get())

        if not in_dir.exists() or not in_dir.is_dir():
            messagebox.showerror("Invalid input", "Please select a valid input folder.")
            return
        if not out_dir.exists() or not out_dir.is_dir():
            messagebox.showerror("Invalid output", "Please select a valid output folder.")
            return

        self._stop_flag.clear()
        self._set_running(True)

        run_folder = out_dir / f"Processed_{now_tag()}"
        safe_mkdir(run_folder)
        log_path = run_folder / "run_log.txt"

        self._log("Run started.")
        self._log(f"Input folder: {in_dir}")
        self._log(f"Output folder: {run_folder}")
        self._log(f"Include subfolders: {recursive}")
        self.status_var.set("Running…")

        def worker():
            try:
                files = list_word_files(in_dir, recursive=recursive)
                if not files:
                    self._log("No .docx/.docm files found. Nothing to do.")
                    self.status_var.set("No files found.")
                    return

                self._log(f"Discovered {len(files)} Word file(s) to process.")

                with open(log_path, "a", encoding="utf-8") as lf:

                    def file_logger(s: str):
                        self._log(s)
                        lf.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {s}\n")
                        lf.flush()

                    success = 0
                    failed = 0

                    for idx, p in enumerate(files, start=1):
                        if self._stop_flag.is_set():
                            file_logger("Stop flag set; exiting before next file.")
                            break

                        file_logger(f"----- [{idx}/{len(files)}] {p.name} -----")

                        out_name = p.with_suffix(".docx").name if p.suffix.lower() == ".docm" else p.name
                        out_path = run_folder / out_name

                        try:
                            clean_docx_or_docm_to_folder(p, out_path, file_logger)
                            success += 1
                        except Exception as e:
                            failed += 1
                            file_logger(f"  ! ERROR processing {p.name}: {e}")

                    file_logger("Run complete.")
                    file_logger(f"Success: {success} | Failed: {failed}")
                    file_logger(f"Artifacts written to: {run_folder}")

                self.status_var.set("Complete.")
            finally:
                self._set_running(False)

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()


if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except KeyboardInterrupt:
        sys.exit(0)
