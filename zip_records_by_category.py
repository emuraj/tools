# zip_records_by_category.py
"""
Purpose:
Create separate ZIP files for selected record files in a folder tree, grouped by filename prefix.

What this file does:
- Recursively reads files from an input folder and all subfolders.
- Lets you choose which categories to process up front.
- Skips explicitly designated files.
- Creates ZIP files per selected category:
  - CAPA
  - CR
  - DEV
  - EC
  - INV
- Can either:
  - create one ZIP per selected category, or
  - split ZIPs into capped chunks, such as DEV1_files.zip, DEV2_files.zip.
- Writes the ZIP files to an output folder.
- Preserves each file's relative subfolder path inside the ZIP.
- Prints progress to the IDE/terminal.
- Writes an event log to the same output folder as the ZIP files.
- Checks available output drive space before creating ZIP files.

Place in the larger scheme:
This is a packaging utility for splitting exported records into selected category-specific archives
with traceable progress logging and optional ZIP-size chunking.
"""

import shutil
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED


# ============================================================
# INPUT FIELDS - EDIT THESE TWO VALUES
# ============================================================

INPUT_FOLDER = r"C:\Users\e.muraj\Desktop\Production-2026a\Production-2026a\QMS_Rebuild_20260226_084225"
OUTPUT_FOLDER = r"D:\PRO2026zips\additionals"


# ============================================================
# CATEGORY SELECTION - SET EACH ONE TO "Yes" OR "No"
# ============================================================

CAPA = "No"
CR = "No"
EC = "Yes"
DEV = "Yes"
INV = "Yes"


# ============================================================
# ZIP SIZE CONTROL
# ============================================================
# Set SPLIT_ZIPS_BY_SIZE to "No" to create one ZIP per selected category.
# Set SPLIT_ZIPS_BY_SIZE to "Yes" to split each category into capped chunks.

SPLIT_ZIPS_BY_SIZE = "No"

MAX_ZIP_SIZE_MB = 25
PLANNING_ZIP_SIZE_MB = 24

MAX_ZIP_SIZE_BYTES = MAX_ZIP_SIZE_MB * 1024 * 1024
PLANNING_ZIP_SIZE_BYTES = PLANNING_ZIP_SIZE_MB * 1024 * 1024


# ============================================================
# LOGGING OPTIONS
# ============================================================

LOG_SKIPPED_FILES = "Yes"
LOG_IGNORED_FILES = "No"
LOG_UNSELECTED_CATEGORY_FILES = "No"
LOG_FILES_TOO_LARGE_FOR_ZIP = "Yes"

# Print/log every N zipped files.
# Use 1 for every file. Use 10, 25, 50, or 100 for less noisy output.

PROGRESS_EVERY_N_FILES = 1


# ============================================================
# CONFIGURATION
# ============================================================

CATEGORY_CHOICES = {
    "CAPA": CAPA,
    "CR": CR,
    "EC": EC,
    "DEV": DEV,
    "INV": INV,
}

SKIPPED_FILES = {
    "_attachment_summary_20260226_084026.csv",
    "_final_sha256_20260226_084026.csv",
    "OOS-00001.pdf",
    "OOS-00002.pdf",
    "_sha_mismatches_20260226_084026.pdf",
    "_sha_mismatches_20260226_084026.csv",
    "_sha_register_20260226_084026.pdf",
}

LOG_FILE_NAME = "zip_event_log.txt"


def is_yes(value: str) -> bool:
    """
    Return True if the configured value means yes.
    """
    return value.strip().lower() == "yes"


def get_selected_categories() -> list[str]:
    """
    Return only categories marked Yes.
    """
    return [
        category
        for category, choice in CATEGORY_CHOICES.items()
        if is_yes(choice)
    ]


def get_category(filename: str, selected_categories: list[str]) -> str | None:
    """
    Return the matching selected category if the filename starts with one of the selected prefixes.
    Otherwise return None.
    """
    for category in selected_categories:
        if filename.startswith(category):
            return category

    return None


def get_known_category(filename: str) -> str | None:
    """
    Return the matching known category even if that category was not selected.
    Otherwise return None.
    """
    for category in CATEGORY_CHOICES:
        if filename.startswith(category):
            return category

    return None


def write_log(log_file, message: str) -> None:
    """
    Print a message to the IDE/terminal and write it to the event log.
    """
    print(message)
    log_file.write(message + "\n")


def format_percent(completed: int, total: int) -> str:
    """
    Return a formatted completion percentage.
    """
    if total == 0:
        return "100.0%"

    return f"{(completed / total) * 100:.1f}%"


def format_bytes(size_bytes: int) -> str:
    """
    Format bytes as a readable size.
    """
    size = float(size_bytes)

    for unit in ["bytes", "KB", "MB", "GB", "TB"]:
        if size < 1024 or unit == "TB":
            return f"{size:.2f} {unit}"
        size /= 1024

    return f"{size_bytes} bytes"


def should_report_progress(completed_files: int, total_files: int) -> bool:
    """
    Return True when progress should be printed/logged for the current file.
    """
    if completed_files == 1:
        return True

    if completed_files == total_files:
        return True

    return completed_files % PROGRESS_EVERY_N_FILES == 0


def check_output_space(output_path: Path, total_input_bytes: int, log_file) -> None:
    """
    Check whether the output drive appears to have enough free space.

    This is conservative because ZIP compression may reduce size, but PDFs and images
    often do not compress much. The script requires free space at least equal to the
    selected input file size.
    """
    disk_usage = shutil.disk_usage(output_path)
    free_bytes = disk_usage.free
    required_bytes = total_input_bytes

    write_log(log_file, "OUTPUT SPACE CHECK")
    write_log(log_file, f"Estimated selected input size: {format_bytes(total_input_bytes)}")
    write_log(log_file, f"Available output drive space: {format_bytes(free_bytes)}")

    if free_bytes < required_bytes:
        raise OSError(
            "Not enough free space on the output drive. "
            f"Estimated need: {format_bytes(required_bytes)}. "
            f"Available: {format_bytes(free_bytes)}. "
            "Choose a different output drive or run fewer categories."
        )

    write_log(log_file, "Output drive has enough free space based on estimated input size.")
    write_log(log_file, "")


def scan_files(
    input_path: Path,
    selected_categories: list[str],
    log_file,
) -> tuple[dict[str, list[Path]], list[Path], list[Path], list[Path]]:
    """
    Recursively scan the input folder and classify files into:
    - files to zip by selected category
    - skipped files
    - unselected category files
    - ignored unmatched files
    """
    files_by_category: dict[str, list[Path]] = {
        category: [] for category in selected_categories
    }

    skipped_files: list[Path] = []
    unselected_category_files: list[Path] = []
    ignored_files: list[Path] = []

    write_log(log_file, "Scanning input folder and subfolders...")

    for file_path in sorted(input_path.rglob("*")):
        if not file_path.is_file():
            continue

        relative_path = file_path.relative_to(input_path)
        filename = file_path.name

        if filename in SKIPPED_FILES:
            skipped_files.append(file_path)

            if is_yes(LOG_SKIPPED_FILES):
                write_log(log_file, f"[SKIPPED] {relative_path}")

            continue

        selected_category = get_category(filename, selected_categories)

        if selected_category is not None:
            files_by_category[selected_category].append(file_path)
            continue

        known_category = get_known_category(filename)

        if known_category is not None:
            unselected_category_files.append(file_path)

            if is_yes(LOG_UNSELECTED_CATEGORY_FILES):
                write_log(
                    log_file,
                    f"[NOT SELECTED] {relative_path} - category {known_category} is set to No",
                )

            continue

        ignored_files.append(file_path)

        if is_yes(LOG_IGNORED_FILES):
            write_log(
                log_file,
                f"[IGNORED] {relative_path} - filename does not start with a configured category",
            )

    return files_by_category, skipped_files, unselected_category_files, ignored_files


def remove_files_too_large_for_single_zip(
    input_path: Path,
    files_by_category: dict[str, list[Path]],
    log_file,
) -> tuple[dict[str, list[Path]], list[Path]]:
    """
    Remove files that are too large to fit safely into one capped ZIP.

    This only runs when SPLIT_ZIPS_BY_SIZE is set to Yes.
    """
    cleaned_files_by_category: dict[str, list[Path]] = {
        category: [] for category in files_by_category
    }

    files_too_large: list[Path] = []

    write_log(log_file, "")
    write_log(log_file, "Checking for files too large for capped ZIP chunks...")

    for category, files in files_by_category.items():
        for file_path in files:
            file_size = file_path.stat().st_size

            if file_size > PLANNING_ZIP_SIZE_BYTES:
                files_too_large.append(file_path)

                if is_yes(LOG_FILES_TOO_LARGE_FOR_ZIP):
                    relative_path = file_path.relative_to(input_path)
                    write_log(
                        log_file,
                        f"[TOO LARGE] {relative_path} is {format_bytes(file_size)}; "
                        f"larger than planning cap of {PLANNING_ZIP_SIZE_MB} MB. "
                        "Skipped because it may exceed the final ZIP limit by itself.",
                    )

                continue

            cleaned_files_by_category[category].append(file_path)

    write_log(log_file, f"Files too large for capped ZIP chunks: {len(files_too_large)}")
    write_log(log_file, "")

    return cleaned_files_by_category, files_too_large


def plan_zip_chunks(
    files_by_category: dict[str, list[Path]],
    log_file,
) -> dict[str, list[list[Path]]]:
    """
    Create ZIP chunk plans per category.

    If SPLIT_ZIPS_BY_SIZE is No:
    - Each selected category gets one ZIP chunk.

    If SPLIT_ZIPS_BY_SIZE is Yes:
    - Each category is split into capped chunks.
    """
    chunk_plan_by_category: dict[str, list[list[Path]]] = {}

    if not is_yes(SPLIT_ZIPS_BY_SIZE):
        write_log(log_file, "ZIP chunking is OFF. Creating one ZIP per selected category.")

        for category, files in files_by_category.items():
            if files:
                chunk_plan_by_category[category] = [files]
                write_log(
                    log_file,
                    f"{category}: planned 1 ZIP with {len(files)} files.",
                )
            else:
                chunk_plan_by_category[category] = []
                write_log(log_file, f"{category}: no ZIP planned.")

        write_log(log_file, "")
        return chunk_plan_by_category

    write_log(log_file, "ZIP chunking is ON.")
    write_log(
        log_file,
        f"Planning ZIP chunks using {PLANNING_ZIP_SIZE_MB} MB planning cap.",
    )

    for category, files in files_by_category.items():
        chunks: list[list[Path]] = []
        current_chunk: list[Path] = []
        current_chunk_size = 0

        for file_path in files:
            file_size = file_path.stat().st_size

            if not current_chunk:
                current_chunk.append(file_path)
                current_chunk_size = file_size
                continue

            would_exceed_planning_limit = (
                current_chunk_size + file_size > PLANNING_ZIP_SIZE_BYTES
            )

            if would_exceed_planning_limit:
                chunks.append(current_chunk)
                current_chunk = [file_path]
                current_chunk_size = file_size
            else:
                current_chunk.append(file_path)
                current_chunk_size += file_size

        if current_chunk:
            chunks.append(current_chunk)

        chunk_plan_by_category[category] = chunks

        if chunks:
            write_log(
                log_file,
                f"{category}: planned {len(chunks)} ZIP chunk(s).",
            )

            for index, chunk in enumerate(chunks, start=1):
                chunk_source_size = sum(file_path.stat().st_size for file_path in chunk)

                write_log(
                    log_file,
                    f"  {category}{index}_files.zip: "
                    f"{len(chunk)} files, source size {format_bytes(chunk_source_size)}",
                )
        else:
            write_log(log_file, f"{category}: no ZIP chunks planned.")

    write_log(log_file, "")

    return chunk_plan_by_category


def get_zip_filename(category: str, chunk_index: int, total_chunks: int) -> str:
    """
    Return the ZIP filename.

    If chunking is off:
    - DEV_files.zip

    If chunking is on:
    - DEV1_files.zip
    - DEV2_files.zip
    """
    if not is_yes(SPLIT_ZIPS_BY_SIZE):
        return f"{category}_files.zip"

    return f"{category}{chunk_index}_files.zip"


def create_zip_chunks(
    input_path: Path,
    output_path: Path,
    chunk_plan_by_category: dict[str, list[list[Path]]],
    total_files_to_zip: int,
    log_file,
) -> dict[str, int]:
    """
    Create ZIP files from the chunk plan.

    Returns a dictionary:
    - ZIP filename -> number of files included
    """
    completed_files = 0
    created_zips: dict[str, int] = {}

    for category, chunks in chunk_plan_by_category.items():
        if not chunks:
            write_log(log_file, f"No files found for selected category: {category}")
            continue

        total_chunks = len(chunks)

        for chunk_index, files in enumerate(chunks, start=1):
            zip_name = get_zip_filename(category, chunk_index, total_chunks)
            zip_path = output_path / zip_name

            write_log(log_file, f"Creating ZIP: {zip_path}")

            with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zip_file:
                for file_path in files:
                    relative_path = file_path.relative_to(input_path)
                    zip_file.write(file_path, arcname=relative_path)

                    completed_files += 1

                    if should_report_progress(completed_files, total_files_to_zip):
                        percent_complete = format_percent(completed_files, total_files_to_zip)

                        write_log(
                            log_file,
                            f"[{percent_complete}] Zipped {relative_path} into {zip_path.name}",
                        )

            actual_zip_size = zip_path.stat().st_size
            created_zips[zip_path.name] = len(files)

            write_log(
                log_file,
                f"Completed ZIP: {zip_path.name} "
                f"({len(files)} files, actual size {format_bytes(actual_zip_size)})",
            )

            if is_yes(SPLIT_ZIPS_BY_SIZE) and actual_zip_size > MAX_ZIP_SIZE_BYTES:
                write_log(
                    log_file,
                    f"[WARNING] {zip_path.name} is larger than {MAX_ZIP_SIZE_MB} MB. "
                    "Review this file manually.",
                )

            write_log(log_file, "")

    return created_zips


def zip_files_by_category(input_folder: str, output_folder: str) -> None:
    """
    Create ZIP files per selected category from files in the input folder and its subfolders.
    Also writes an event log to the output folder.
    """
    input_path = Path(input_folder)
    output_path = Path(output_folder)

    if not input_path.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_path}")

    if not input_path.is_dir():
        raise NotADirectoryError(f"Input path is not a folder: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)

    selected_categories = get_selected_categories()
    log_path = output_path / LOG_FILE_NAME

    started_at = datetime.now()

    with open(log_path, "w", encoding="utf-8") as log_file:
        write_log(log_file, "ZIP EVENT LOG")
        write_log(log_file, f"Input folder: {input_path}")
        write_log(log_file, f"Output folder: {output_path}")
        write_log(log_file, f"Started: {started_at:%Y-%m-%d %H:%M:%S}")
        write_log(log_file, f"Split ZIPs by size: {SPLIT_ZIPS_BY_SIZE}")

        if is_yes(SPLIT_ZIPS_BY_SIZE):
            write_log(log_file, f"Maximum final ZIP size: {MAX_ZIP_SIZE_MB} MB")
            write_log(log_file, f"Planning ZIP size cap: {PLANNING_ZIP_SIZE_MB} MB")

        write_log(log_file, "")

        write_log(log_file, "CATEGORY SELECTION")
        for category, choice in CATEGORY_CHOICES.items():
            write_log(log_file, f"{category}: {choice}")

        write_log(log_file, "")

        if not selected_categories:
            write_log(log_file, "No categories were selected. Set at least one category to Yes.")
            return

        files_by_category, skipped_files, unselected_category_files, ignored_files = scan_files(
            input_path=input_path,
            selected_categories=selected_categories,
            log_file=log_file,
        )

        total_eligible_files_before_size_check = sum(
            len(files) for files in files_by_category.values()
        )

        total_scanned_files = (
            total_eligible_files_before_size_check
            + len(skipped_files)
            + len(unselected_category_files)
            + len(ignored_files)
        )

        total_input_bytes = sum(
            file_path.stat().st_size
            for files in files_by_category.values()
            for file_path in files
        )

        write_log(log_file, "")
        write_log(log_file, "SCAN COMPLETE")
        write_log(log_file, f"Total files scanned: {total_scanned_files}")
        write_log(log_file, f"Eligible files before ZIP-size check: {total_eligible_files_before_size_check}")
        write_log(log_file, f"Skipped files: {len(skipped_files)}")
        write_log(log_file, f"Unselected category files: {len(unselected_category_files)}")
        write_log(log_file, f"Ignored unmatched files: {len(ignored_files)}")
        write_log(log_file, "")

        check_output_space(output_path, total_input_bytes, log_file)

        files_too_large: list[Path] = []

        if is_yes(SPLIT_ZIPS_BY_SIZE):
            files_by_category, files_too_large = remove_files_too_large_for_single_zip(
                input_path=input_path,
                files_by_category=files_by_category,
                log_file=log_file,
            )
        else:
            write_log(log_file, "ZIP-size cap is OFF. No files will be excluded due to size.")
            write_log(log_file, "")

        total_files_to_zip = sum(len(files) for files in files_by_category.values())

        chunk_plan_by_category = plan_zip_chunks(
            files_by_category=files_by_category,
            log_file=log_file,
        )

        created_zips = create_zip_chunks(
            input_path=input_path,
            output_path=output_path,
            chunk_plan_by_category=chunk_plan_by_category,
            total_files_to_zip=total_files_to_zip,
            log_file=log_file,
        )

        finished_at = datetime.now()
        elapsed_time = finished_at - started_at

        write_log(log_file, "SUMMARY")
        write_log(log_file, f"Total files scanned: {total_scanned_files}")
        write_log(log_file, f"Eligible files before ZIP-size check: {total_eligible_files_before_size_check}")
        write_log(log_file, f"Total files zipped: {total_files_to_zip}")
        write_log(log_file, f"Total files skipped by skip list: {len(skipped_files)}")
        write_log(log_file, f"Total files too large for ZIP cap: {len(files_too_large)}")
        write_log(log_file, f"Total unselected category files: {len(unselected_category_files)}")
        write_log(log_file, f"Total files ignored: {len(ignored_files)}")
        write_log(log_file, "")

        if created_zips:
            write_log(log_file, "ZIP FILES CREATED")
            for zip_name, file_count in created_zips.items():
                zip_path = output_path / zip_name
                zip_size = zip_path.stat().st_size

                write_log(
                    log_file,
                    f"{zip_name}: {file_count} files, {format_bytes(zip_size)}",
                )
        else:
            write_log(log_file, "No ZIP files were created.")

        write_log(log_file, "")
        write_log(log_file, f"Finished: {finished_at:%Y-%m-%d %H:%M:%S}")
        write_log(log_file, f"Elapsed time: {elapsed_time}")

    print(f"\nEvent log written to: {log_path}")


if __name__ == "__main__":
    zip_files_by_category(INPUT_FOLDER, OUTPUT_FOLDER)