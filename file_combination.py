"""
Combine files from DEV, EC, and INV record folders into one output folder.

What this script does:
1. Looks inside a main source folder that contains folders named DEV, EC, and INV.
2. Scans each record subfolder inside those folders (examples: DEV-00002, EC-00055, INV-00068).
3. Copies all files from those record folders into one single folder named "combined files".
4. Preserves all original files by automatically renaming duplicates if the same filename appears more than once.

Example source structure:
MAIN_SOURCE_FOLDER/
    DEV/
        DEV-00002/
            file1.pdf
            file2.docx
    EC/
        EC-00055/
            image1.png
    INV/
        INV-00068/
            notes.txt

Output:
OUTPUT_PARENT_FOLDER/
    combined files/
        file1.pdf
        file2.docx
        image1.png
        notes.txt
"""

from pathlib import Path
import shutil

# =========================
# MAIN LOCATIONS TO EDIT
# =========================

MAIN_SOURCE_FOLDER = r"C:\Users\e.muraj\OneDrive - Neurotech USA, Inc\Production-2026\3_TrackWise QMS Rebuild\QMS_Rebuild_20260227_143354"
OUTPUT_PARENT_FOLDER = r"C:\Users\e.muraj\Desktop\combined files prod26"

# =========================
# SCRIPT LOGIC
# =========================

TOP_LEVEL_FOLDERS = ["DEV", "EC", "INV"]
COMBINED_FOLDER_NAME = "combined files"


def get_unique_destination(dest_folder: Path, filename: str) -> Path:
    """
    If filename already exists in destination, create a unique version:
    example.pdf -> example_1.pdf -> example_2.pdf, etc.
    """
    candidate = dest_folder / filename
    if not candidate.exists():
        return candidate

    stem = candidate.stem
    suffix = candidate.suffix
    counter = 1

    while True:
        new_name = f"{stem}_{counter}{suffix}"
        new_candidate = dest_folder / new_name
        if not new_candidate.exists():
            return new_candidate
        counter += 1


def main():
    source_root = Path(MAIN_SOURCE_FOLDER)
    output_root = Path(OUTPUT_PARENT_FOLDER)
    combined_folder = output_root / COMBINED_FOLDER_NAME

    if not source_root.exists():
        print(f"ERROR: Source folder does not exist:\n{source_root}")
        return

    combined_folder.mkdir(parents=True, exist_ok=True)

    copied_count = 0
    skipped_count = 0

    for top_folder_name in TOP_LEVEL_FOLDERS:
        top_folder = source_root / top_folder_name

        if not top_folder.exists() or not top_folder.is_dir():
            print(f"Skipping missing folder: {top_folder}")
            continue

        for record_folder in top_folder.iterdir():
            if not record_folder.is_dir():
                continue

            for item in record_folder.rglob("*"):
                if item.is_file():
                    destination = get_unique_destination(combined_folder, item.name)
                    shutil.copy2(item, destination)
                    copied_count += 1

    print("Done.")
    print(f"Source root   : {source_root}")
    print(f"Output folder : {combined_folder}")
    print(f"Files copied  : {copied_count}")
    print(f"Files skipped : {skipped_count}")


if __name__ == "__main__":
    main()