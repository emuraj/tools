# neurotic_export_sampler.py
"""
Purpose
-------
Scan a parent TrackWise export location and extract a representative evidence pack
from the specific source files and folder structures that Neurotic actually uses.

What this script does
---------------------
- Accepts a single parent export folder where raw TrackWise exports exist.
- Accepts an output folder, defaulting to a DEV location.
- Locates QMS- and DMS-relevant CSV files referenced by Neurotic.
- Pulls header rows and a small number of representative data rows from each.
- Summarizes DMS folder/version/binaries/renditions structure.
- Summarizes QMS export-batch / ContentVersion payload structure.
- Writes:
    * evidence_manifest.json
    * evidence_report.txt
    * sampled CSV extracts

Place in the larger scheme
--------------------------
This is not part of Neurotic runtime processing. It is an analysis utility to
collect example evidence from raw TrackWise exports for documentation, QA review,
and technical reporting.

Run examples
------------
python neurotic_export_sampler.py --parent_root "C:\\Users\\e.muraj\\Desktop\\Production\\1_TrackWise Raw Export Repository"

python neurotic_export_sampler.py ^
  --parent_root "C:\\Users\\e.muraj\\Desktop\\Production\\1_TrackWise Raw Export Repository" ^
  --output_root "C:\\Users\\e.muraj\\Desktop\\DEV\\NEUROTIC_EVIDENCE" ^
  --sample_rows 8
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Iterable, Optional


# ============================================================
# USER-EDITABLE FIELDS
# ============================================================

# Parent folder containing raw TrackWise export roots.
# You can leave this blank and pass --parent_root at runtime.
PARENT_ROOT = r"C:\Users\e.muraj\Desktop\Production\1_TrackWise Raw Export Repository"

# Default output folder for the evidence pack.
DEFAULT_OUTPUT_ROOT = r"C:\Users\e.muraj\Desktop\DEV\NEUROTIC_EVIDENCE"

# Number of sample data rows to capture per CSV.
DEFAULT_SAMPLE_ROWS = 5

# Number of DMS document folders to summarize.
DEFAULT_MAX_DMS_FOLDERS = 30

# Number of QMS export batches to summarize.
DEFAULT_MAX_QMS_BATCHES = 15

# Maximum preferred columns to keep in sampled CSV extracts.
DEFAULT_MAX_SAMPLE_COLUMNS = 12


# ============================================================
# Config
# ============================================================

CSV_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")

# QMS files clearly used by Neurotic
QMS_TARGET_FILES = [
    "ContentVersion.csv",
    "ContentDocumentLink.csv",
    "Attachment.csv",
    "User.csv",
    "CMPL123QMS__CAPA__c.csv",
    "CMPL123QMS__Deviation__c.csv",
    "CMPL123QMS__Change_Control__c.csv",
    "CMPL123QMS__Investigation__c.csv",
    "CMPL123QMS__Effectiveness_Check__c.csv",
    "CMPL123QMS__OOS__c.csv",
    "CMPL123__Action__c.csv",
    "CMPL123__Audit_Trail_Entry__c.csv",
    "CMPL123__WF_History__c.csv",
    "CMPL123QMS__Extension_Request__c.csv",
    "CMPL123QMS__Change_Assessment__c.csv",
]

# DMS files clearly used by Neurotic
DMS_TARGET_FILES = [
    "SPARTADMS__Corporate_Document__c.csv",
    "ContentVersion.csv",
    "ContentDocumentLink.csv",
    "User.csv",
    "SPARTADMS__Document_Version__c.csv",
    "SPARTADMS__Effective_Version_Details__c.csv",
    "SPARTADMS__Rendition_Status__c.csv",
    "SPARTADMS__Document_Approvers__c.csv",
    "SPARTADMS__Document_Reviewer__c.csv",
    "SPARTADMS__Workflow_History__c.csv",
    "SPARTADMS__Related_Document__c.csv",
]

# Preferred columns to include first in the sampled outputs
PREFERRED_COLUMNS = [
    "Id",
    "Name",
    "OwnerId",
    "CreatedById",
    "LastModifiedById",
    "ParentId",
    "LinkedEntityId",
    "ContentDocumentId",
    "FirstPublishLocationId",
    "Title",
    "FileType",
    "PathOnClient",
    "SPARTADMS__Approved_Document_Number__c",
    "SPARTADMS__Full_Document_Number__c",
    "SPARTADMS__Document_Number_Major__c",
    "SPARTADMS__Document_Revision_At_Effective__c",
    "SPARTADMS__Document_Status__c",
    "SPARTADMS__Document_Type__c",
    "SPARTADMS__Alphanumeric_Sequence__c",
]


# ============================================================
# Models
# ============================================================

@dataclass
class CsvSample:
    name: str
    path: str
    encoding: str
    row_count_estimate: int
    headers: list[str]
    sampled_columns: list[str]
    sampled_rows: list[dict[str, str]] = field(default_factory=list)


@dataclass
class DmsFolderExample:
    folder_name: str
    folder_path: str
    version_dirs: list[str]
    has_lastreviewed: bool
    versions_with_binaries: list[str]
    versions_with_renditions: list[str]


@dataclass
class QmsBatchExample:
    batch_name: str
    batch_path: str
    has_contentversion_dir: bool
    contentversion_file_count: int
    first_contentversion_files: list[str]


@dataclass
class EvidencePack:
    parent_root: str
    output_root: str
    qms_csv_samples: list[CsvSample] = field(default_factory=list)
    dms_csv_samples: list[CsvSample] = field(default_factory=list)
    dms_folder_examples: list[DmsFolderExample] = field(default_factory=list)
    qms_batch_examples: list[QmsBatchExample] = field(default_factory=list)
    discovered_qms_roots: list[str] = field(default_factory=list)
    discovered_dms_roots: list[str] = field(default_factory=list)


# ============================================================
# CSV helpers
# ============================================================

def read_csv_rows(path: Path) -> tuple[list[str], list[list[str]], str]:
    last_exc: Optional[Exception] = None
    for enc in CSV_ENCODINGS:
        try:
            with path.open("r", newline="", encoding=enc, errors="strict") as f:
                rdr = csv.reader(f)
                rows = list(rdr)
            if not rows:
                return [], [], enc
            return rows[0], rows[1:], enc
        except Exception as exc:
            last_exc = exc
            continue
    raise RuntimeError(f"Could not decode CSV: {path} :: {last_exc}")


def choose_columns(headers: list[str], max_cols: int = DEFAULT_MAX_SAMPLE_COLUMNS) -> list[str]:
    normalized = {h.strip(): h for h in headers}
    chosen: list[str] = []
    for col in PREFERRED_COLUMNS:
        if col in normalized and col not in chosen:
            chosen.append(col)
    for h in headers:
        if h not in chosen:
            chosen.append(h)
        if len(chosen) >= max_cols:
            break
    return chosen[:max_cols]


def sample_csv(path: Path, sample_rows: int) -> CsvSample:
    headers, body, enc = read_csv_rows(path)
    sampled_cols = choose_columns(headers)
    index = {h: i for i, h in enumerate(headers)}
    sampled = []
    for row in body[:sample_rows]:
        rec = {}
        for col in sampled_cols:
            idx = index.get(col)
            rec[col] = row[idx] if idx is not None and idx < len(row) else ""
        sampled.append(rec)
    return CsvSample(
        name=path.name,
        path=str(path),
        encoding=enc,
        row_count_estimate=len(body),
        headers=headers,
        sampled_columns=sampled_cols,
        sampled_rows=sampled,
    )


# ============================================================
# Discovery helpers
# ============================================================

def find_files(root: Path, filenames: Iterable[str]) -> dict[str, list[Path]]:
    wanted = {name.lower() for name in filenames}
    found: dict[str, list[Path]] = defaultdict(list)
    for p in root.rglob("*"):
        if p.is_file() and p.name.lower() in wanted:
            found[p.name].append(p)
    return dict(found)


def discover_qms_roots(parent: Path) -> list[Path]:
    roots: list[Path] = []
    for p in parent.rglob("*"):
        if not p.is_dir():
            continue
        names = {child.name for child in p.iterdir() if child.is_file()}
        if (
            "CMPL123QMS__CAPA__c.csv" in names
            and "CMPL123QMS__Deviation__c.csv" in names
            and "ContentVersion.csv" in names
        ):
            roots.append(p)
    return sorted(set(roots))


def discover_dms_roots(parent: Path) -> list[Path]:
    roots: list[Path] = []
    for p in parent.rglob("*"):
        if not p.is_dir():
            continue
        child_dirs = [c for c in p.iterdir() if c.is_dir()]
        if not child_dirs:
            continue
        match_count = 0
        for c in child_dirs[:200]:
            if re.fullmatch(r"[A-Za-z0-9]+-\d{6}", c.name) or re.fullmatch(r"\d+", c.name):
                match_count += 1
        if match_count >= 10:
            roots.append(p)
    return sorted(set(roots))


# ============================================================
# DMS structure sampling
# ============================================================

def collect_dms_folder_examples(
    dms_root: Path, max_folders: int = DEFAULT_MAX_DMS_FOLDERS
) -> list[DmsFolderExample]:
    examples: list[DmsFolderExample] = []
    candidates = [p for p in dms_root.iterdir() if p.is_dir()]
    candidates = sorted(candidates, key=lambda p: p.name.lower())

    for folder in candidates[:max_folders]:
        version_dirs = [p for p in folder.iterdir() if p.is_dir()]
        versions = sorted([p.name for p in version_dirs if p.name.isdigit()])
        has_lastreviewed = any(p.name.lower() == "lastreviewed" for p in version_dirs)

        versions_with_binaries: list[str] = []
        versions_with_renditions: list[str] = []

        for vdir in version_dirs:
            if not vdir.is_dir() or not vdir.name.isdigit():
                continue
            if (vdir / "binaries").is_dir():
                versions_with_binaries.append(vdir.name)
            if (vdir / "renditions").is_dir():
                versions_with_renditions.append(vdir.name)

        if versions or has_lastreviewed:
            examples.append(
                DmsFolderExample(
                    folder_name=folder.name,
                    folder_path=str(folder),
                    version_dirs=versions,
                    has_lastreviewed=has_lastreviewed,
                    versions_with_binaries=versions_with_binaries,
                    versions_with_renditions=versions_with_renditions,
                )
            )

    return examples


# ============================================================
# QMS batch / payload sampling
# ============================================================

def collect_qms_batch_examples(
    qms_parent: Path, max_batches: int = DEFAULT_MAX_QMS_BATCHES, max_files_per_batch: int = 12
) -> list[QmsBatchExample]:
    batches: list[QmsBatchExample] = []
    candidates = [p for p in qms_parent.iterdir() if p.is_dir() and p.name.startswith("WE_")]
    candidates = sorted(candidates, key=lambda p: p.name.lower())

    for batch in candidates[:max_batches]:
        cv_dir = batch / "ContentVersion"
        files = []
        count = 0
        if cv_dir.is_dir():
            files = sorted([p.name for p in cv_dir.iterdir() if p.is_file()])[:max_files_per_batch]
            count = len([p for p in cv_dir.iterdir() if p.is_file()])

        batches.append(
            QmsBatchExample(
                batch_name=batch.name,
                batch_path=str(batch),
                has_contentversion_dir=cv_dir.is_dir(),
                contentversion_file_count=count,
                first_contentversion_files=files,
            )
        )

    return batches


# ============================================================
# Writers
# ============================================================

def write_csv_sample(sample: CsvSample, out_dir: Path) -> None:
    out_path = out_dir / f"sample__{sample.name}"
    with out_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=sample.sampled_columns)
        w.writeheader()
        for row in sample.sampled_rows:
            w.writerow(row)


def write_text_report(pack: EvidencePack, out_dir: Path) -> None:
    report = out_dir / "evidence_report.txt"
    with report.open("w", encoding="utf-8") as f:
        f.write("Neurotic Evidence Sampler Report\n")
        f.write("=" * 80 + "\n\n")
        f.write(f"Parent root: {pack.parent_root}\n")
        f.write(f"Output root: {pack.output_root}\n\n")

        f.write("Discovered QMS roots\n")
        f.write("-" * 80 + "\n")
        for p in pack.discovered_qms_roots:
            f.write(f"{p}\n")
        f.write("\n")

        f.write("Discovered DMS roots\n")
        f.write("-" * 80 + "\n")
        for p in pack.discovered_dms_roots:
            f.write(f"{p}\n")
        f.write("\n")

        f.write("DMS folder examples\n")
        f.write("-" * 80 + "\n")
        for ex in pack.dms_folder_examples:
            f.write(f"Folder: {ex.folder_name}\n")
            f.write(f"  Path: {ex.folder_path}\n")
            f.write(f"  Versions: {', '.join(ex.version_dirs) if ex.version_dirs else '(none)'}\n")
            f.write(f"  Has lastreviewed: {ex.has_lastreviewed}\n")
            f.write(
                f"  Versions with binaries: "
                f"{', '.join(ex.versions_with_binaries) if ex.versions_with_binaries else '(none)'}\n"
            )
            f.write(
                f"  Versions with renditions: "
                f"{', '.join(ex.versions_with_renditions) if ex.versions_with_renditions else '(none)'}\n\n"
            )

        f.write("QMS batch examples\n")
        f.write("-" * 80 + "\n")
        for ex in pack.qms_batch_examples:
            f.write(f"Batch: {ex.batch_name}\n")
            f.write(f"  Path: {ex.batch_path}\n")
            f.write(f"  Has ContentVersion dir: {ex.has_contentversion_dir}\n")
            f.write(f"  ContentVersion file count: {ex.contentversion_file_count}\n")
            if ex.first_contentversion_files:
                f.write("  First ContentVersion files:\n")
                for name in ex.first_contentversion_files:
                    f.write(f"    - {name}\n")
            f.write("\n")

        f.write("QMS CSV samples\n")
        f.write("-" * 80 + "\n")
        for s in pack.qms_csv_samples:
            f.write(f"{s.name} :: {s.path}\n")
            f.write(f"  Encoding: {s.encoding}\n")
            f.write(f"  Row count estimate: {s.row_count_estimate}\n")
            f.write(f"  Sampled columns: {', '.join(s.sampled_columns)}\n\n")

        f.write("DMS CSV samples\n")
        f.write("-" * 80 + "\n")
        for s in pack.dms_csv_samples:
            f.write(f"{s.name} :: {s.path}\n")
            f.write(f"  Encoding: {s.encoding}\n")
            f.write(f"  Row count estimate: {s.row_count_estimate}\n")
            f.write(f"  Sampled columns: {', '.join(s.sampled_columns)}\n\n")


# ============================================================
# Core execution
# ============================================================

def build_evidence_pack(
    parent_root: Path,
    output_root: Path,
    sample_rows: int = DEFAULT_SAMPLE_ROWS,
    max_dms_folders: int = DEFAULT_MAX_DMS_FOLDERS,
    max_qms_batches: int = DEFAULT_MAX_QMS_BATCHES,
) -> EvidencePack:
    pack = EvidencePack(
        parent_root=str(parent_root),
        output_root=str(output_root),
    )

    qms_roots = discover_qms_roots(parent_root)
    dms_roots = discover_dms_roots(parent_root)
    pack.discovered_qms_roots = [str(p) for p in qms_roots]
    pack.discovered_dms_roots = [str(p) for p in dms_roots]

    # QMS sampling
    qms_seen: set[Path] = set()
    for root in qms_roots:
        found = find_files(root, QMS_TARGET_FILES)
        for _, paths in found.items():
            for p in paths:
                if p in qms_seen:
                    continue
                qms_seen.add(p)
                try:
                    pack.qms_csv_samples.append(sample_csv(p, sample_rows))
                except Exception as exc:
                    print(f"[WARN] Could not sample QMS CSV {p}: {exc}")

        batch_parent = root.parent if root.name.startswith("WE_") else root
        pack.qms_batch_examples.extend(
            collect_qms_batch_examples(batch_parent, max_batches=max_qms_batches)
        )
        break

    # DMS sampling
    dms_seen: set[Path] = set()
    for root in dms_roots:
        found = find_files(root, DMS_TARGET_FILES)
        for _, paths in found.items():
            for p in paths:
                if p in dms_seen:
                    continue
                dms_seen.add(p)
                try:
                    pack.dms_csv_samples.append(sample_csv(p, sample_rows))
                except Exception as exc:
                    print(f"[WARN] Could not sample DMS CSV {p}: {exc}")

        pack.dms_folder_examples.extend(
            collect_dms_folder_examples(root, max_folders=max_dms_folders)
        )
        break

    pack.qms_csv_samples.sort(key=lambda s: s.name.lower())
    pack.dms_csv_samples.sort(key=lambda s: s.name.lower())

    return pack


def write_evidence_pack(pack: EvidencePack, output_root: Path) -> None:
    output_root.mkdir(parents=True, exist_ok=True)

    for s in pack.qms_csv_samples:
        write_csv_sample(s, output_root)
    for s in pack.dms_csv_samples:
        write_csv_sample(s, output_root)

    write_text_report(pack, output_root)

    manifest_path = output_root / "evidence_manifest.json"
    with manifest_path.open("w", encoding="utf-8") as f:
        json.dump(asdict(pack), f, indent=2)

    print(f"Evidence pack written to: {output_root}")
    print(f"Manifest: {manifest_path}")


# ============================================================
# CLI
# ============================================================

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sample TrackWise exports for Neurotic documentation."
    )
    parser.add_argument(
        "--parent_root",
        default=PARENT_ROOT,
        help="Parent folder containing raw TrackWise export roots.",
    )
    parser.add_argument(
        "--output_root",
        default=DEFAULT_OUTPUT_ROOT,
        help="Output folder for the evidence pack.",
    )
    parser.add_argument(
        "--sample_rows",
        type=int,
        default=DEFAULT_SAMPLE_ROWS,
        help="Number of sample rows per CSV.",
    )
    parser.add_argument(
        "--max_dms_folders",
        type=int,
        default=DEFAULT_MAX_DMS_FOLDERS,
        help="Number of DMS document folders to summarize.",
    )
    parser.add_argument(
        "--max_qms_batches",
        type=int,
        default=DEFAULT_MAX_QMS_BATCHES,
        help="Number of QMS export batches to summarize.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if not args.parent_root:
        raise SystemExit(
            "No parent_root provided. Set PARENT_ROOT at the top of the script "
            "or pass --parent_root on the command line."
        )

    parent_root = Path(args.parent_root).expanduser().resolve()
    output_root = Path(args.output_root).expanduser().resolve()

    if not parent_root.exists():
        raise SystemExit(f"Parent root does not exist: {parent_root}")

    pack = build_evidence_pack(
        parent_root=parent_root,
        output_root=output_root,
        sample_rows=args.sample_rows,
        max_dms_folders=args.max_dms_folders,
        max_qms_batches=args.max_qms_batches,
    )
    write_evidence_pack(pack, output_root)


if __name__ == "__main__":
    main()