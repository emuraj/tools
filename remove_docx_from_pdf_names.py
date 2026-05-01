# remove_source_extensions_from_pdf_names.py
"""
Rename PDF files that mistakenly include a source-file extension before '.pdf'.

Examples:
    "report.docx.pdf" -> "report.pdf"
    "report.doc.pdf"  -> "report.pdf"
    "report.xlsx.pdf" -> "report.pdf"

What this file does:
- Recursively scans a root folder and all subfolders
- Renames PDF files ending with known non-PDF source extensions before '.pdf'
- Leaves every other file untouched

Place in the larger scheme:
- Use this as a one-off cleanup utility for bad filenames
"""

from pathlib import Path


# ===== EDIT THIS =====
ROOT_FOLDER = r"D:\laptop software\Production\1_TrackWise Raw Export Repository\TrackWise Raw DMS Export"
DRY_RUN = False   # True = preview only, False = actually rename files
# =====================


SOURCE_EXTENSIONS = {
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".ppt",
    ".pptx",
    ".rtf",
    ".txt",
    ".csv",
    ".odt",
    ".ods",
    ".odp",
}


def remove_source_extensions_from_pdf_names(root_folder: str, dry_run: bool = True) -> None:
    root = Path(root_folder)

    if not root.exists():
        raise FileNotFoundError(f"Folder does not exist: {root}")
    if not root.is_dir():
        raise NotADirectoryError(f"Not a folder: {root}")

    renamed_count = 0

    for path in root.rglob("*"):
        if not path.is_file():
            continue

        if path.suffix.lower() != ".pdf":
            continue

        pdf_stem = Path(path.stem)
        source_extension = pdf_stem.suffix.lower()

        if source_extension not in SOURCE_EXTENSIONS:
            continue

        new_name = pdf_stem.stem + ".pdf"
        new_path = path.with_name(new_name)

        if new_path.exists():
            print(f"SKIP (target exists): {path} -> {new_path}")
            continue

        print(f"RENAME: {path} -> {new_path}")

        if not dry_run:
            path.rename(new_path)

        renamed_count += 1

    mode = "DRY RUN" if dry_run else "DONE"
    print(f"\n{mode}: {renamed_count} file(s) matched.")


if __name__ == "__main__":
    remove_source_extensions_from_pdf_names(ROOT_FOLDER, dry_run=DRY_RUN)