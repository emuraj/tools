#!/usr/bin/env python3
"""
Diagnose attachments that the finaliser would replace by a placeholder.

Usage
-----
Just edit the DEV_REBUILD_PATH below (or pass a folder on the command line)
and hit Run in your IDE.  No other dependencies than those the finaliser
already needs.

Output
------
* Human‑readable summary on stdout.
* CSV `attachment_diagnostics_<timestamp>.csv` next to this script.
"""

from __future__ import annotations

import csv
import shutil
import sys
from datetime import datetime
from pathlib import Path

# ── copy‑pasted helpers from finaliser (kept minimal & read‑only) ─────
def _sniff_magic(path: Path) -> str | None:
    try:
        with path.open("rb") as fh:
            sig = fh.read(8)
    except Exception:
        return None
    if sig.startswith(b"%PDF"):
        return "pdf"
    if sig.startswith(b"\x89PNG"):
        return "png"
    if sig[:3] == b"\xFF\xD8\xFF":
        return "jpeg"
    if sig.startswith(b"GIF8"):
        return "gif"
    if sig.startswith(b"II*\x00") or sig.startswith(b"MM\x00*"):
        return "tiff"
    return None

OFFICE_EXTS = {
    ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    ".odt", ".ods", ".odp", ".rtf", ".txt", ".csv", ".html",
}
IMG_EXTS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif"}

# ──────────────────────────────────────────────────────────────────────
DEV_REBUILD_PATH = r"C:\Users\e.muraj\OneDrive - Neurotech USA, Inc\TW_QMS\QMS_export_15\QMS_Rebuild_20250711_184755"  # ← put e.g. r"C:\Data\QMS_Rebuild_20250722_093010" here

root = Path(sys.argv[1] if len(sys.argv) > 1 else DEV_REBUILD_PATH)
if not root:
    print("❌  Specify rebuild folder via DEV_REBUILD_PATH or as CLI arg.")
    sys.exit(1)
root = Path(root).expanduser().resolve()
if not root.exists():
    print(f"❌  Folder not found: {root}")
    sys.exit(1)

records = [p for p in root.glob("*/*") if p.is_dir() and p.parent != root]
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
csv_path = Path(__file__).with_name(f"attachment_diagnostics_{timestamp}.csv")

problem_rows: list[list[str]] = []

for rec in records:
    for f in rec.iterdir():
        ext = f.suffix.lower()
        is_summary = "summary" in f.stem.lower()
        is_json    = ext == ".json"
        if is_summary or is_json:
            continue                                  # handled ok by finaliser

        ok = False
        if ext == ".pdf":
            ok = True
        elif not ext:                                # blob – sniff
            kind = _sniff_magic(f)
            ok = kind == "pdf" or kind in {"png", "jpeg", "gif", "tiff"}
        elif ext in IMG_EXTS:
            ok = True
        elif ext in OFFICE_EXTS:
            ok = shutil.which("soffice") is not None  # only ok if LO present

        if not ok:
            problem_rows.append([rec.name, f.name, ext or "(no ext)"])

# ── report ────────────────────────────────────────────────────────────
if not problem_rows:
    print("✅  No problematic attachments detected.")
else:
    print(f"⚠️  {len(problem_rows)} attachment(s) would be replaced by a placeholder:")
    for rec_id, fname, ext in problem_rows:
        print(f"   - {rec_id:12s}  {fname}  [{ext}]")

    with csv_path.open("w", newline="", encoding="utf-8") as fh:
        wr = csv.writer(fh)
        wr.writerow(["Record_ID", "File_Name", "Extension"])
        wr.writerows(problem_rows)
    print(f"\nDetails written to {csv_path}")
