# document_source_file_discrepancy_checker.py
"""
Purpose:
    Check whether Word document source files listed in a CSV are present in a designated folder.

What this file does:
    This script reads a CSV file, extracts Column A as "Document ID" and Column F as
    "Document Source File", then compares the listed Word document filenames against
    the actual .docx and .doc files found in a target folder.

    It outputs a discrepancy catalog CSV containing only records where the CSV lists
    a .docx or .doc source file that is missing from the folder.

    The output catalog is sorted naturally by Document ID, meaning IDs such as
    MBR-1, MBR-2, MBR-3, and MBR-10 are sorted numerically after the prefix instead
    of alphabetically.

Place in the larger scheme:
    This utility supports document inventory reconciliation by identifying source
    Word documents referenced in a controlled catalog that are not present in the
    expected storage folder.
"""

from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path


# =============================================================================
# USER-EDITABLE PATH FIELDS
# Replace these paths before running the script.
# =============================================================================

INPUT_CSV_PATH = Path(
    r"C:\Users\e.muraj\OneDrive - Neurotech USA, Inc\LOCKED\Intellect_20260430_193817507_100000102000000080.csv"
)

SOURCE_FOLDER_PATH = Path(
    r"C:\Users\e.muraj\OneDrive - Neurotech USA, Inc\LOCKED\original"
)

OUTPUT_FOLDER_PATH = Path(
    r"C:\Users\e.muraj\OneDrive - Neurotech USA, Inc\LOCKED"
)


# =============================================================================
# OUTPUT FILE SETTINGS
# The script will deposit the output CSV inside OUTPUT_FOLDER_PATH.
# =============================================================================

OUTPUT_CSV_FILENAME_PREFIX = "missing_word_documents"


# =============================================================================
# CSV COLUMN SETTINGS
# Column A = index 0
# Column F = index 5
# =============================================================================

DOCUMENT_ID_COLUMN_INDEX = 0
DOCUMENT_SOURCE_FILE_COLUMN_INDEX = 5


# =============================================================================
# FILE EXTENSION SETTINGS
# Only these document types are checked.
# PDFs and all other file types are ignored.
# =============================================================================

VALID_WORD_EXTENSIONS = {".docx", ".doc"}


def normalize_filename(filename: str) -> str:
    """
    Normalize filenames for comparison.

    This trims whitespace, removes any folder path portion, and compares filenames
    case-insensitively.
    """
    return Path(filename.strip()).name.lower()


def is_word_document(filename: str) -> bool:
    """
    Return True only for .docx or .doc filenames.
    """
    suffix = Path(filename.strip()).suffix.lower()
    return suffix in VALID_WORD_EXTENSIONS


def build_output_csv_path(output_folder_path: Path) -> Path:
    """
    Build the final output CSV path inside the designated output folder.

    A timestamp is included so prior output files are not overwritten.
    """
    if output_folder_path.exists() and not output_folder_path.is_dir():
        raise NotADirectoryError(
            f"OUTPUT_FOLDER_PATH exists but is not a folder: {output_folder_path}"
        )

    output_folder_path.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"{OUTPUT_CSV_FILENAME_PREFIX}_{timestamp}.csv"

    return output_folder_path / output_filename


def get_word_files_in_folder(folder_path: Path) -> set[str]:
    """
    Return a normalized set of .docx and .doc filenames found in the folder.

    This checks files directly inside the folder, not subfolders.
    """
    if not folder_path.exists():
        raise FileNotFoundError(f"Folder does not exist: {folder_path}")

    if not folder_path.is_dir():
        raise NotADirectoryError(f"Path is not a folder: {folder_path}")

    word_files = set()

    for item in folder_path.iterdir():
        if item.is_file() and item.suffix.lower() in VALID_WORD_EXTENSIONS:
            word_files.add(normalize_filename(item.name))

    return word_files


def find_missing_documents(csv_path: Path, folder_path: Path) -> list[dict[str, str]]:
    """
    Compare CSV-listed Word documents against files present in the folder.

    Returns a list of discrepancy records with:
        - Document ID
        - Document Source File
    """
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file does not exist: {csv_path}")

    if not csv_path.is_file():
        raise IsADirectoryError(f"CSV path points to a folder, not a file: {csv_path}")

    available_word_files = get_word_files_in_folder(folder_path)
    missing_documents: list[dict[str, str]] = []

    with csv_path.open(mode="r", encoding="utf-8-sig", newline="") as csv_file:
        reader = csv.reader(csv_file)

        header = next(reader, None)
        if header is None:
            raise ValueError("CSV file is empty.")

        required_columns = max(
            DOCUMENT_ID_COLUMN_INDEX,
            DOCUMENT_SOURCE_FILE_COLUMN_INDEX,
        )

        if len(header) <= required_columns:
            raise ValueError(
                "CSV does not contain enough columns. "
                "Expected at least Column A and Column F."
            )

        for row_number, row in enumerate(reader, start=2):
            if len(row) <= required_columns:
                continue

            document_id = row[DOCUMENT_ID_COLUMN_INDEX].strip()
            document_source_file = row[DOCUMENT_SOURCE_FILE_COLUMN_INDEX].strip()

            if not document_source_file:
                continue

            if not is_word_document(document_source_file):
                continue

            normalized_source_file = normalize_filename(document_source_file)

            if normalized_source_file not in available_word_files:
                missing_documents.append(
                    {
                        "Document ID": document_id,
                        "Document Source File": document_source_file,
                    }
                )

    return missing_documents


def document_id_sort_key(record: dict[str, str]) -> tuple[str, int, str]:
    """
    Sort Document IDs by letter prefix first, then numeric value.

    Examples:
        GDL-2, GDL-3, GDL-4
        MBR-1, MBR-2, MBR-3, MBR-10
        MS-4, MS-6, MS-16
    """
    document_id = record["Document ID"].strip()

    if "-" not in document_id:
        return document_id.upper(), -1, document_id.upper()

    prefix, number_part = document_id.rsplit("-", 1)

    try:
        number = int(number_part)
    except ValueError:
        number = -1

    return prefix.upper(), number, document_id.upper()


def write_discrepancy_catalog(
    missing_documents: list[dict[str, str]],
    output_csv_path: Path,
) -> None:
    """
    Write the missing-document discrepancy catalog to a CSV.

    Output is sorted naturally by Document ID:
        prefix alphabetically, then numeric suffix numerically.
    """
    sorted_missing_documents = sorted(
        missing_documents,
        key=document_id_sort_key,
    )

    try:
        with output_csv_path.open(
            mode="w",
            encoding="utf-8-sig",
            newline="",
        ) as output_file:
            fieldnames = ["Document ID", "Document Source File"]
            writer = csv.DictWriter(output_file, fieldnames=fieldnames)

            writer.writeheader()
            writer.writerows(sorted_missing_documents)

    except PermissionError as error:
        raise PermissionError(
            "Could not write the output CSV. Check that the output folder is writable "
            "and that a CSV with the same name is not open in Excel. "
            f"Attempted output path: {output_csv_path}"
        ) from error


def main() -> None:
    output_csv_path = build_output_csv_path(OUTPUT_FOLDER_PATH)

    missing_documents = find_missing_documents(
        csv_path=INPUT_CSV_PATH,
        folder_path=SOURCE_FOLDER_PATH,
    )

    write_discrepancy_catalog(
        missing_documents=missing_documents,
        output_csv_path=output_csv_path,
    )

    print("Discrepancy check complete.")
    print(f"Missing Word documents found: {len(missing_documents)}")
    print(f"Output written to: {output_csv_path}")


if __name__ == "__main__":
    main()