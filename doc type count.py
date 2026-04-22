#!/usr/bin/env python3
"""
Two-stage merge & rename pipeline with folder-structure preservation.

Stage 1: Merge A → B → C into D (priority A > B > C), preserving subfolders
- "Document key" is derived from the filename stem; if COLLAPSE_REV_SUFFIX=True,
  a trailing ".<digits>" (e.g., ".02") is removed for *matching only*.
- The first folder to introduce a key claims it; later folders with the same key are skipped.
- Within the winning source, *all files* sharing that key are copied.
- The source file's relative subfolder path is preserved in D.
- Case-insensitive matching by default. Uses shutil.copy2. Ignores junk files.

Stage 2: Copy D → E, preserving subfolders, stripping trailing decimal from output filenames
- Every file in D is copied to E under the same relative subfolders.
- If the filename stem ends with ".<digits>", that suffix is removed in the *output filename*.
  Example: "SOP-1017.02.pdf" → "SOP-1017.pdf".
- Conflicts in E follow ON_EXIST_E.

Notes:
- Fully recursive; directory trees from sources are replicated in outputs.
- This script never deletes D or E.
"""

from __future__ import annotations

import os
import sys
import re
import shutil
from pathlib import Path
from typing import Dict, List, Set, Tuple

# ============================
# ==== USER SETTINGS (EDIT) ===
# ============================
# Paste your paths between the quotes. Use raw strings r"..." on Windows.
FOLDER_A = r"C:\Users\e.muraj\Desktop\SEPTEMBER\SEP DMS Build\SEP DMS 1\Neurotic_DMS_ORG_20251020142658\EFFECTIVE_DOCUMENTS\Binaries"   # Input 1 (highest priority)
FOLDER_B = r"C:\Users\e.muraj\Desktop\SEPTEMBER\SEP DMS Build\SEP DMS 1\Neurotic_DMS_ORG_20251020142658\EFFECTIVE_DOCUMENTS\Renditions"   # Input 2
FOLDER_C = r"C:\Users\e.muraj\Desktop\SEPTEMBER\SEP QMS Build\SEP QMS Build 19\QMS_FINAL_RECORDS_20251017_101801"   # Input 3 (lowest priority)
OUTPUT_D = r"C:\Users\e.muraj\Desktop\SEPTEMBER\SEP import builds\SEP import build 1\merged"   # Intermediate merged output (A+B+C), structure preserved
OUTPUT_E = r"C:\Users\e.muraj\Desktop\SEPTEMBER\SEP import builds\SEP import build 1\merged and clean"   # Final output (copy of D with decimals removed from names), structure preserved

# Conflict policy when writing into D (after A/B/C merge):
#   "overwrite" -> replace existing file in D
#   "skip"      -> keep existing file in D, skip copying
#   "rename"    -> copy to D using " (1)", " (2)", ... before extension
ON_EXIST_D = "overwrite"

# Conflict policy when writing into E (after D→E transform):
#   "overwrite" | "skip" | "rename" (same semantics)
ON_EXIST_E = "overwrite"

# Matching mode for A/B/C keys:
# - False: case-insensitive (recommended)
# - True:  case-sensitive
CASE_SENSITIVE = False

# Collapse trailing ".<digits>" from stem when forming keys for A/B/C:
#   "SOP-1017.02" -> "SOP-1017"
COLLAPSE_REV_SUFFIX = True

# Dry run: print actions only, perform no copies (applies to both stages).
DRY_RUN = False
# ============================
# ==== END USER SETTINGS =====
# ============================

JUNK_NAMES = {"Thumbs.db", ".DS_Store"}
JUNK_PREFIXES = ("~$",)
HIDDEN_PREFIX = "."   # ignore dotfiles by default

# Regex to strip a final ".digits" revision suffix from a stem.
REV_SUFFIX_RE = re.compile(r"^(?P<core>.*?)(?:\.(?P<rev>\d{1,3}))$")


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
    """Filter out temp/hidden/system junk."""
    name = file_path.name
    if name in JUNK_NAMES:
        return True
    if name.startswith(HIDDEN_PREFIX):
        return True
    for pref in JUNK_PREFIXES:
        if name.startswith(pref):
            return True
    return False


def collapse_revision_suffix(stem: str) -> str:
    """
    Remove a trailing .<digits> (1–3 digits) if present.
    'CAPA-00017.02' → 'CAPA-00017'
    """
    m = REV_SUFFIX_RE.match(stem)
    return m.group("core") if m else stem


def make_doc_key(stem: str, *, case_sensitive: bool, collapse_rev: bool) -> str:
    """
    Compute the document key from a filename stem (no extension):
    - Optionally strip trailing .<digits>.
    - Apply case normalization unless case-sensitive.
    """
    key = collapse_revision_suffix(stem) if collapse_rev else stem
    return key if case_sensitive else key.lower()


def iter_files_with_rel(root: Path):
    """
    Recursively yield (file_path, rel_parent) for each non-ignored file under root.
    rel_parent is the relative directory under root, or '.' for files at the root.
    """
    for p in root.rglob("*"):
        if p.is_file() and not is_ignored(p):
            rel = p.parent.relative_to(root)
            yield p, rel


def plan_copy_exact(dst_full_path: Path, on_exist: str) -> Tuple[Path, str]:
    """
    Decide destination path/action when the exact destination path (with dirs) is known.
    Returns (dest_path, action) where action ∈ {'copy', 'skip', 'rename'}.
    """
    dst_full_path = normalize_path(dst_full_path)
    if not dst_full_path.exists():
        return dst_full_path, "copy"

    if on_exist == "overwrite":
        return dst_full_path, "copy"
    elif on_exist == "skip":
        return dst_full_path, "skip"
    elif on_exist == "rename":
        stem, suffix = dst_full_path.stem, dst_full_path.suffix
        parent = dst_full_path.parent
        i = 1
        while True:
            candidate = normalize_path(parent / f"{stem} ({i}){suffix}")
            if not candidate.exists():
                return candidate, "rename"
            i += 1
    else:
        raise ValueError(f"Unknown on_exist policy: {on_exist}")


def copy_file(src: Path, dst: Path, *, dry_run: bool) -> None:
    if dry_run:
        print(f"[DRY-RUN] COPY: {src} -> {dst}")
        return
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    print(f"COPIED: {src} -> {dst}")


def gather_key_buckets(root: Path, *, case_sensitive: bool, collapse_rev: bool) -> Dict[str, List[Tuple[Path, Path]]]:
    """
    Build {doc_key: [(src_path, rel_parent), ...]} for all files under 'root'.
    This allows copying *all* files that share a key while preserving their relative subfolders.
    """
    buckets: Dict[str, List[Tuple[Path, Path]]] = {}
    for src, rel_parent in iter_files_with_rel(root):
        key = make_doc_key(src.stem, case_sensitive=case_sensitive, collapse_rev=collapse_rev)
        buckets.setdefault(key, []).append((src, rel_parent))
    return buckets


def stage1_merge_ABC_into_D(a_dir: Path, b_dir: Path, c_dir: Path, out_dir: Path) -> None:
    """
    Merge A→B→C into D by key, preserving relative subfolders.
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    case_sensitive = CASE_SENSITIVE
    collapse_rev = COLLAPSE_REV_SUFFIX

    sources = [("A", a_dir), ("B", b_dir), ("C", c_dir)]
    claimed_keys: Set[str] = set()

    per_source_copied = {"A": 0, "B": 0, "C": 0}
    per_source_skipped_keys = {"A": 0, "B": 0, "C": 0}
    skipped_due_to_exist = 0
    renamed_count = 0

    for label, folder in sources:
        print(f"\n[Stage 1] Scanning {label}: {folder}")
        key_to_items = gather_key_buckets(folder, case_sensitive=case_sensitive, collapse_rev=collapse_rev)

        for key, items in key_to_items.items():
            if key in claimed_keys:
                print(f"[Stage 1] SKIP KEY (claimed by earlier source): {key} ({len(items)} file(s) in {label})")
                per_source_skipped_keys[label] += 1
                continue

            claimed_keys.add(key)
            for src, rel_parent in items:
                dst_full = (out_dir / rel_parent / src.name)
                dst, action = plan_copy_exact(dst_full, on_exist=ON_EXIST_D)
                if action == "skip":
                    print(f"[Stage 1] SKIP (exists in D): {dst_full}")
                    skipped_due_to_exist += 1
                    continue
                if action == "rename":
                    renamed_count += 1
                copy_file(src, dst, dry_run=DRY_RUN)
                per_source_copied[label] += 1

    print("\n[Stage 1] === SUMMARY ===")
    print(f"Copied from A: {per_source_copied['A']}")
    print(f"Copied from B: {per_source_copied['B']}")
    print(f"Copied from C: {per_source_copied['C']}")
    print(f"Skipped keys in A (should be 0): {per_source_skipped_keys['A']}")
    print(f"Skipped keys in B (already in A): {per_source_skipped_keys['B']}")
    print(f"Skipped keys in C (already in A/B): {per_source_skipped_keys['C']}")
    print(f"Skipped files due to existing exact path in D (ON_EXIST_D='{ON_EXIST_D}'): {skipped_due_to_exist}")
    print(f"Renamed due to exist in D (ON_EXIST_D='rename'): {renamed_count}")


def stage2_copy_D_to_E_strip_decimals(d_dir: Path, e_dir: Path) -> None:
    """
    Copy every file from D to E, preserving subfolders, removing a trailing .<digits> from the *output filename*.
    """
    e_dir.mkdir(parents=True, exist_ok=True)

    copied = 0
    rename_collisions = 0
    skip_exists = 0

    print(f"\n[Stage 2] Transform-copy D → E (strip trailing decimals in filename), preserving folder tree:")
    print(f"[Stage 2] Scanning D: {d_dir}")

    for src, rel_parent in iter_files_with_rel(d_dir):
        new_stem = collapse_revision_suffix(src.stem)
        target_name = f"{new_stem}{src.suffix}"
        dst_full = e_dir / rel_parent / target_name

        dst, action = plan_copy_exact(dst_full, on_exist=ON_EXIST_E)
        if action == "skip":
            print(f"[Stage 2] SKIP (exists in E): {dst_full}")
            skip_exists += 1
            continue
        if action == "rename":
            rename_collisions += 1
        copy_file(src, dst, dry_run=DRY_RUN)
        copied += 1

    print("\n[Stage 2] === SUMMARY ===")
    print(f"Copied into E: {copied}")
    print(f"Filename collisions in E resolved by ON_EXIST_E='rename': {rename_collisions}")
    print(f"Skipped in E due to ON_EXIST_E='skip': {skip_exists}")


def main():
    a_dir = Path(FOLDER_A).resolve()
    b_dir = Path(FOLDER_B).resolve()
    c_dir = Path(FOLDER_C).resolve()
    d_dir = Path(OUTPUT_D).resolve()
    e_dir = Path(OUTPUT_E).resolve()

    # Validate inputs (A/B/C must exist; D/E will be created if missing)
    for fld, label in [(a_dir, "Folder A"), (b_dir, "Folder B"), (c_dir, "Folder C")]:
        if not fld.exists() or not fld.is_dir():
            print(f"ERROR: {label} does not exist or is not a directory: {fld}", file=sys.stderr)
            sys.exit(2)

    # Stage 1: Merge A/B/C → D (structure preserved)
    stage1_merge_ABC_into_D(a_dir, b_dir, c_dir, d_dir)

    # Stage 2: Copy D → E with decimal stripped in filenames (structure preserved)
    stage2_copy_D_to_E_strip_decimals(d_dir, e_dir)


if __name__ == "__main__":
    main()
