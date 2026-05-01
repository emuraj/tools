# flag_dash_spacing.py
"""
Purpose:
    Identify rows in a CSV where Column A contains a dash with spaces before or after it.

What this file does:
    Reads an input CSV file, always writes the first row as the header, checks the first
    column of each remaining row, and writes each flagged full row to a new output CSV file.

Place in the larger scheme:
    This can be used as a data-quality check before importing, validating, or reconciling
    record identifiers.
"""

import csv
import re
from pathlib import Path


# =============================================================================
# CSV FILE PATHS — EDIT THESE VALUES
# =============================================================================

INPUT_CSV_PATH = Path(
    r"C:\Users\e.muraj\Downloads\Intellect_20260427_074117635_100000102000000080.csv"
)

OUTPUT_CSV_PATH = Path(r"C:\Users\e.muraj\Downloads")

OUTPUT_FILE_NAME = "flagged_dash_spacing_rows.csv"

# =============================================================================
# END CSV FILE PATHS
# =============================================================================


DASH_SPACING_PATTERN = re.compile(r"\s-\s|\s-|-\s")


def flag_rows_with_dash_spacing(input_csv: Path, output_csv: Path) -> int:
    flagged_count = 0

    output_csv.parent.mkdir(parents=True, exist_ok=True)

    with input_csv.open("r", newline="", encoding="utf-8-sig") as infile, \
         output_csv.open("w", newline="", encoding="utf-8") as outfile:

        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        header = next(reader, None)

        if header is not None:
            writer.writerow(header)

        for row in reader:
            if not row:
                continue

            column_a_value = row[0]

            if DASH_SPACING_PATTERN.search(column_a_value):
                writer.writerow(row)
                flagged_count += 1

    return flagged_count


def main() -> None:
    if not INPUT_CSV_PATH.exists():
        raise FileNotFoundError(f"Input CSV not found: {INPUT_CSV_PATH}")

    if INPUT_CSV_PATH.is_dir():
        raise IsADirectoryError(f"Input path is a folder, not a CSV file: {INPUT_CSV_PATH}")

    OUTPUT_CSV_PATH.mkdir(parents=True, exist_ok=True)

    if not OUTPUT_CSV_PATH.is_dir():
        raise NotADirectoryError(f"Output path is not a folder: {OUTPUT_CSV_PATH}")

    output_csv_file = OUTPUT_CSV_PATH / OUTPUT_FILE_NAME

    flagged_count = flag_rows_with_dash_spacing(
        input_csv=INPUT_CSV_PATH,
        output_csv=output_csv_file,
    )

    print(f"Done. Flagged {flagged_count} row(s).")
    print(f"Output written to: {output_csv_file}")


if __name__ == "__main__":
    main()