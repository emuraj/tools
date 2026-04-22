# remove_docx_from_pdf_names.py
"""
Rename PDF files that mistakenly include '.docx' before '.pdf'.

Example:
    "report.docx.pdf" -> "report.pdf"

What this file does:
- Recursively scans a root folder and all subfolders
- Renames only files ending in '.docx.pdf'
- Leaves every other file untouched

Place in the larger scheme:
- Use this as a one-off cleanup utility for bad filenames
"""

from pathlib import Path


# ===== EDIT THIS =====
ROOT_FOLDER = r"C:\Users\e.muraj\Desktop\Production\1_TrackWise Raw Export Repository\TrackWise Raw DMS Export"
DRY_RUN = False   # True = preview only, False = actually rename files
# =====================


def remove_docx_from_pdf_names(root_folder: str, dry_run: bool = True) -> None:
    root = Path(root_folder)

    if not root.exists():
        raise FileNotFoundError(f"Folder does not exist: {root}")
    if not root.is_dir():
        raise NotADirectoryError(f"Not a folder: {root}")

    renamed_count = 0

    for path in root.rglob("*"):
        if not path.is_file():
            continue

        if path.name.endswith(".docx.pdf"):
            new_name = path.name[:-9] + ".pdf"
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
    remove_docx_from_pdf_names(ROOT_FOLDER, dry_run=DRY_RUN)