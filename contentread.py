#!/usr/bin/env python3
# ────────────────────────────────────────────────────────────────
# export_selfcheck_plus.py
# ---------------------------------------------------------------
# One-stop sanity checker for a TrackWise / Salesforce bulk export.
#
#   • Accepts only the *highest* folder (the one that holds all
#     WE_… sub-folders) – no other arguments required.
#   • Finds ContentVersion.csv, validates headers.
#   • Counts every CSV file by package prefix.
#   • Detects every ContentVersion directory, computes BLOB_ROOT,
#     counts blobs, size, duplicates, missing, orphans.
#   • Writes an Id→blob mapping file (id_to_blob.csv).
#   • Prints:
#       – Three-line summary block (BLOB_ROOT, rule, headers)
#       – Extended stats table
#       – First 20 rows of ContentVersion.csv
#       – First 20 rows of the mapping file
#
# Standard-library only; run with Python 3.8+.
# ────────────────────────────────────────────────────────────────

from __future__ import annotations
import csv
import os
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Iterable

# >>>>>>>>>>>>>>>>>>>>>>>>>> EDIT THIS ONLY <<<<<<<<<<<<<<<<<<<<<<<<<
ROOT_DIR = Path(
    r"C:\Users\e.muraj\Downloads\Trackwise_Exports\Trackwise_Exports_12APR\Unzipped"
)
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

PRIMARY_COLS = {"Id", "Title"}
ALT_EXT_COLS = ("FileExtension", "FileType")

# ───────────────────────── helpers ───────────────────────────────


def find_contentversion_csv(root: Path) -> Path:
    try:
        return next(root.rglob("ContentVersion.csv"))
    except StopIteration:
        sys.exit("❌  ContentVersion.csv not found – check ROOT_DIR.")


def csv_headers(path: Path) -> list[str]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        return next(csv.reader(f))


def iter_csv_rows(path: Path, field: str):
    with path.open(newline="", encoding="utf-8-sig") as f:
        rdr = csv.DictReader(f)
        for row in rdr:
            yield row[field]


def package_breakdown(root: Path) -> Counter:
    cnt = Counter()
    for csv_path in root.rglob("*.csv"):
        name = csv_path.name
        if "__" in name:
            prefix = name.split("__", 1)[0] + "__"
        else:
            prefix = "(core / other)"
        cnt[prefix] += 1
    return cnt


def locate_blob_folders(root: Path) -> set[Path]:
    return {p for p in root.rglob("ContentVersion") if p.is_dir()}


def gather_blobs(blob_dirs: set[Path]) -> set[Path]:
    blobs: set[Path] = set()
    for d in blob_dirs:
        blobs.update(d.iterdir())
    return blobs


def common_root(paths: Iterable[Path]) -> Path:
    return Path(os.path.commonpath([str(p) for p in paths]))


# ───────────────────────── main ──────────────────────────────────
def main() -> None:
    if not ROOT_DIR.exists():
        sys.exit(f"❌  ROOT_DIR does not exist: {ROOT_DIR}")

    print(f"🔍  Scanning export under: {ROOT_DIR}\n")

    # 1. CSV manifest -------------------------------------------------
    csv_manifest = package_breakdown(ROOT_DIR)
    total_csv = sum(csv_manifest.values())
    print("CSV manifest:")
    for pkg, n in csv_manifest.most_common():
        print(f"   {pkg:<25} {n:>4}")
    print(f"   {'TOTAL':<25} {total_csv:>4}\n")

    # 2. ContentVersion checks ---------------------------------------
    cv_csv = find_contentversion_csv(ROOT_DIR)
    print(f"✓ Found ContentVersion.csv → {cv_csv}")

    headers = csv_headers(cv_csv)
    missing = PRIMARY_COLS - set(headers)
    if missing:
        sys.exit(
            f"❌  ContentVersion.csv missing column(s): {', '.join(sorted(missing))}"
        )
    ext_col = next((c for c in ALT_EXT_COLS if c in headers), None)
    if not ext_col:
        need = " or ".join(ALT_EXT_COLS)
        sys.exit(f"❌  Need {need} column in ContentVersion.csv header.")
    print(f"✓ Required columns ok: Id, Title, {ext_col}")

    ids = list(iter_csv_rows(cv_csv, "Id"))
    print(f"✓ ContentVersion rows: {len(ids):,}")

    # 3. Blob discovery ----------------------------------------------
    blob_dirs = locate_blob_folders(ROOT_DIR)
    if not blob_dirs:
        sys.exit("❌  No ContentVersion directories found – cannot locate blobs.")

    blob_root = common_root(blob_dirs)
    all_blobs = gather_blobs(blob_dirs)
    total_size = sum(p.stat().st_size for p in all_blobs) / (1024**3)
    print(
        f"✓ Blobs: {len(all_blobs):,} files across {len(blob_dirs)} folders "
        f"({total_size:.1f} GB)"
    )

    # 4. Duplicate / missing / orphan analysis -----------------------
    duplicates: defaultdict[str, list[Path]] = defaultdict(list)
    missing_ids = []
    blob_index = defaultdict(list)
    for p in all_blobs:
        blob_index[p.stem].append(p)

    for id_ in ids:
        hits = blob_index.get(id_)
        if not hits:
            missing_ids.append(id_)
        elif len(hits) > 1:
            duplicates[id_] = hits

    orphan_blobs = [p for stem, paths in blob_index.items() if stem not in ids]

    # 5. Id → blob mapping file --------------------------------------
    mapping_path = ROOT_DIR / "id_to_blob.csv"
    with mapping_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Id", "BlobPath"])
        for id_ in ids:
            hits = blob_index.get(id_)
            if hits:
                # first hit; swap for sorted(hits)[-1] to choose "newest"
                w.writerow([id_, hits[0]])

    # 6. Three-line summary ------------------------------------------
    print("\n" + "=" * 64)
    print("SUMMARY (copy these three lines)".center(64))
    print("=" * 64)
    rule_msg = (
        "prefix match is safe."
        if not duplicates
        else f"{len(duplicates)} duplicate-prefix Ids – first-match or newest-timestamp."
    )
    print(f"BLOB_ROOT          = {blob_root}")
    print(f"Blob-matching rule = {rule_msg}")
    print(f"Headers            = Id, Title, {ext_col}")
    print("-" * 64)

    # 7. Extended stats ----------------------------------------------
    print(f"Total CSV files         : {total_csv}")
    print(f"Total blobs on disk     : {len(all_blobs)}")
    print(f"Total ContentVersion Ids: {len(ids)}")
    print(f"Missing blobs           : {len(missing_ids)}")
    print(f"Orphan blobs            : {len(orphan_blobs)}")
    print(f"Duplicate-prefix Ids    : {len(duplicates)}")

    # sample duplicates
    if duplicates:
        some = list(duplicates.items())[:5]
        print("\nDuplicate examples:")
        for stem, paths in some:
            print(f"  {stem} → {len(paths)} files")

    # 8. Preview first 20 rows of ContentVersion.csv -----------------
    print("\nPreview: first 20 rows of ContentVersion.csv\n")
    with cv_csv.open(newline="", encoding="utf-8-sig") as f:
        rdr = csv.reader(f)
        header = next(rdr)
        print(", ".join(header))
        for i, row in enumerate(rdr, 1):
            print(", ".join(row))
            if i == 20:
                break

    # 9. Preview first 20 rows of the mapping file -------------------
    print(f"\n✓ Mapping file written → {mapping_path}  ({len(ids)} rows)")
    print("\nFirst 20 rows of id_to_blob.csv\n")
    with mapping_path.open(newline="", encoding="utf-8") as f:
        rdr = csv.reader(f)
        header = next(rdr)
        print(", ".join(header))
        for i, row in enumerate(rdr, 1):
            print(", ".join(row))
            if i == 20:
                break

    print("\nDone.")


if __name__ == "__main__":
    main()
