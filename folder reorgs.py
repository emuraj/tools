#!/usr/bin/env python3
"""
Merge two folders into an output folder by base filename (case-insensitive by default).

Behavior:
- Copy ALL files from Folder A to Output (as-is).
- From Folder B, copy only files whose *base name* (filename without last extension) is NOT present in Folder A.
- Matching is by base name, case-insensitive by default (e.g., SOP-1094.02.docx vs SOP-1094.02.pdf).
- Preserves timestamps/basic metadata via shutil.copy2.
- Ignores common junk (Thumbs.db, .DS_Store, dotfiles, and "~$" temp files).

Edit the section labeled "USER SETTINGS" below with your Folder A/B/C paths and preferences.
"""

from __future__ import annotations

import os
import sys
import shutil
from pathlib import Path
from typing import Set, Tuple

# ============================
# ==== USER SETTINGS (EDIT) ===
# ============================
# Paste your paths between the quotes. Use raw strings (r"...") on Windows.
FOLDER_A   = r"C:\Users\e.muraj\Desktop\SEPTEMBER\SEP DMS Build\SEP DMS 1\Neurotic_DMS_ORG_20251020142658\EFFECTIVE_DOCUMENTS\Binaries"   # e.g., r"C:\Users\me\Docs\A"
FOLDER_B   = r"C:\Users\e.muraj\Desktop\SEPTEMBER\SEP DMS Build\SEP DMS 1\Neurotic_DMS_ORG_20251020142658\EFFECTIVE_DOCUMENTS\Renditions"   # e.g., r"C:\Users\me\Docs\B"
OUTPUT_DIR = r"C:\Users\e.muraj\Desktop\SEPTEMBER\SEP DMS Build\SEP DMS 1\EFFECTIVE_COMB\DOCS"    # e.g., r"C:\Users\me\Docs\Output"

# What to do if a file with the same full filename (including extension) already exists in OUTPUT_DIR:
#   "overwrite" -> replace it
#   "skip"      -> leave existing file, skip copying
#   "rename"    -> copy with " (1)", " (2)", ... appended before extension
ON_EXIST = "overwrite"

# Base-name matching mode:
# - If False: case-insensitive (recommended): "sop-1094.02" equals "SOP-1094.02"
# - If True:  case-sensitive
CASE_SENSITIVE = False

# If True, show the planned actions but do not actually copy files.
DRY_RUN = False

# If True, scan only the immediate files in each folder (not subfolders). This script is flat by default.
# (If you want recursive behavior, set this to True and we’ll copy only files, flattening directory structure in output.)
RECURSIVE = False
# ============================
# ==== END USER SETTINGS =====
# ============================

JUNK_NAMES = {"Thumbs.db", ".DS_Store"}
JUNK_PREFIXES = ("~$",)
HIDDEN_PREFIX = "."  # ignore dotfiles by default


def is_windows() -> bool:
    return os.name == "nt"


def normalize_path(p: Path) -> Path:
    """Add Windows long-path prefix if needed; otherwise return as-is."""
    if is_windows():
        s = str(p)
        if s.startswith("\\\\?\\") or s.startswith("\\\\"):
            return p
        if len(s) >= 240:
            return Path("\\\\?\\" + s)
    return p


def is_ignored(file_path: Path) -> bool:
    name = file_path.name
    if name in JUNK_NAMES:
        return True
    if name.startswith(HIDDEN_PREFIX):
        return True
    for pref in JUNK_PREFIXES:
        if name.startswith(pref):
            return True
    return False


def base_key(path: Path, case_insensitive: bool = True) -> str:
    """Return base name (filename without last extension)."""
    stem = path.stem
    return stem if case_insensitive else stem


def ensure_out_dir(out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)


def plan_copy(src: Path, dst_dir: Path, on_exist: str = "overwrite") -> Tuple[Path, str]:
    """
    Determine the destination path based on on_exist policy.
    Returns (dest_path, action) where action is one of: 'copy', 'skip', 'rename'.
    """
    dst_path = dst_dir / src.name
    dst_path = normalize_path(dst_path)

    if not dst_path.exists():
        return dst_path, "copy"

    if on_exist == "overwrite":
        return dst_path, "copy"
    elif on_exist == "skip":
        return dst_path, "skip"
    elif on_exist == "rename":
        stem = dst_path.stem
        suffix = dst_path.suffix
        i = 1
        while True:
            candidate = dst_dir / f"{stem} ({i}){suffix}"
            candidate = normalize_path(candidate)
            if not candidate.exists():
                return candidate, "rename"
            i += 1
    else:
        raise ValueError(f"Unknown on_exist policy: {on_exist}")


def copy_file(src: Path, dst: Path, dry_run: bool) -> None:
    if dry_run:
        print(f"[DRY-RUN] COPY: {src} -> {dst}")
        return
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    print(f"COPIED: {src} -> {dst}")


def gather_base_names(folder: Path, case_insensitive: bool = True) -> Set[str]:
    """Collect base-name keys from all non-ignored files in the folder (flat or recursive)."""
    keys: Set[str] = set()
    iterator = folder.rglob("*") if RECURSIVE else folder.iterdir()
    for p in iterator:
        if not p.is_file():
            continue
        if is_ignored(p):
            continue
        k = p.stem
        keys.add(k.lower() if case_insensitive else k)
    return keys


def iter_files(folder: Path):
    """Yield non-ignored files (flat or recursive)."""
    iterator = folder.rglob("*") if RECURSIVE else folder.iterdir()
    for p in iterator:
        if p.is_file() and not is_ignored(p):
            yield p


def main():
    folder_a = Path(FOLDER_A).resolve()
    folder_b = Path(FOLDER_B).resolve()
    out_dir = Path(OUTPUT_DIR).resolve()

    # Validate inputs
    for fld, label in [(folder_a, "Folder A"), (folder_b, "Folder B")]:
        if not fld.exists() or not fld.is_dir():
            print(f"ERROR: {label} does not exist or is not a directory: {fld}", file=sys.stderr)
            sys.exit(2)

    ensure_out_dir(out_dir)

    case_insensitive = not CASE_SENSITIVE

    # 1) Copy all files from A
    a_keys = gather_base_names(folder_a, case_insensitive=case_insensitive)

    copied_from_a = 0
    copied_from_b = 0
    skipped_b_due_to_a = 0
    skipped_due_to_exist = 0
    renamed_count = 0

    print(f"Scanning A: {folder_a}")
    for src in iter_files(folder_a):
        dst, action = plan_copy(src, out_dir, on_exist=ON_EXIST)
        if action == "skip":
            print(f"SKIP (exists): {src.name}")
            skipped_due_to_exist += 1
            continue
        if action == "rename":
            renamed_count += 1
        copy_file(src, dst, dry_run=DRY_RUN)
        copied_from_a += 1

    # 2) Copy files from B only if base not in A
    print(f"\nScanning B: {folder_b}")
    for src in iter_files(folder_b):
        key = src.stem
        key = key.lower() if case_insensitive else key
        if key in a_keys:
            print(f"SKIP (base in A): {src.name}")
            skipped_b_due_to_a += 1
            continue

        dst, action = plan_copy(src, out_dir, on_exist=ON_EXIST)
        if action == "skip":
            print(f"SKIP (exists): {src.name}")
            skipped_due_to_exist += 1
            continue
        if action == "rename":
            renamed_count += 1
        copy_file(src, dst, dry_run=DRY_RUN)
        copied_from_b += 1

    # Summary
    print("\n=== SUMMARY ===")
    print(f"Copied from A:          {copied_from_a}")
    print(f"Copied from B:          {copied_from_b}")
    print(f"Skipped (B in A):       {skipped_b_due_to_a}")
    print(f"Skipped (exists in C):  {skipped_due_to_exist} (due to ON_EXIST='{ON_EXIST}')")
    print(f"Renamed due to exist:   {renamed_count} (due to ON_EXIST='rename')")
    if DRY_RUN:
        print("\nNOTE: Dry-run mode; no files were actually copied.")


if __name__ == "__main__":
    main()
