# augment_training_history.py
"""
Purpose:
Standardize training history workbook document identifiers and create an optional combined CSV.

What this file does:
The script scans a folder for Excel workbooks, finds the Training Number or Document Number column,
renames it to Document Number, inserts Document Type before it, inserts Version and Original TW File name
after it, and optionally combines all processed rows into one CSV file.

Place in the larger scheme:
This prepares historical training records for cleanup, reconciliation, migration, analysis, and GMP-report
use while preserving the original TrainingWise identifier for traceability.
"""

from pathlib import Path
import csv
import re
import shutil
import win32com.client as win32


FOLDER_PATH = r"C:\Users\e.muraj\Downloads\TW training"
COMBINE = "Yes"

OUTPUT_FOLDER_NAME = "_augmented_output"
COMBINED_CSV_NAME = "combined_training_history.csv"

HEADER_SEARCH_ROWS = 20
ROW_PRINT_INTERVAL = 100

DOCUMENT_NUMBER_HEADERS = {"training number", "document number"}

VERSION_SPLIT_EXCLUDED_PREFIXES = {
    "CORP",
    "CRS",
    "TRN",
    "EXT",
    "ISO",
    "PROF",
    "Z",
}

PREFIX_TO_DOCUMENT_TYPE = {
    "SOP": "SOP (Standard Operating Procedure)",
    "CORP": "CORP (Corporate Document)",
    "PRN": "PRN (Presentation Powerpoint)",
    "B": "Master Batch Record",
    "POL": "POL (Policy)",
    "CRS": "CRS (CRS)",
    "SSP": "SSP (Stability Study Protocol)",
    "TMCP": "TMCP (Test Material Certification Protocol)",
    "D": "Master Batch Record",
    "QM": "QM (Quality Manual)",
    "WI": "WI (Work Instruction)",
    "C": "Master Batch Record",
    "PLN": "PLN (Plan)",
    "A": "Master Batch Record",
    "TM": "TM (Test Method or Standard Operating Procedure)",
    "Z": "Master Batch Record",
    "TRN": "TRN (Training Document)",
    "PRO": "PRO (Protocol)",
    "MVP": "MVP (Method Validation Protocol)",
    "OJT": "OJT (On the Job Training)",
    "ISO": "ISO (International Standard)",
    "PROF": "PROF (Proficiency Training)",
    "EXT": "EXT (External Document)",
}


def normalize_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def normalize_combine_setting(value: str) -> bool:
    return str(value).strip().lower() in {"yes", "y", "true", "1"}


def percent(numerator: int, denominator: int) -> float:
    if denominator <= 0:
        return 100.0
    return round((numerator / denominator) * 100, 2)


def print_progress(message: str, same_line: bool = False) -> None:
    end = "\r" if same_line else "\n"
    print(message, end=end, flush=True)


def clear_progress_line() -> None:
    print(" " * 160, end="\r", flush=True)


def safely_set_excel_property(excel, property_name: str, value: object, warning_label: str) -> None:
    try:
        setattr(excel, property_name, value)
    except Exception as error:
        print(f"Warning: could not {warning_label}. Continuing anyway: {error}", flush=True)


def get_prefix(original_document_number: object) -> str:
    value = normalize_text(original_document_number)

    if not value:
        return ""

    if value[0].isdigit():
        return "SOP"

    match = re.match(r"^[A-Za-z]+", value)

    if not match:
        return ""

    return match.group(0).upper()


def get_document_type(original_document_number: object) -> str:
    prefix = get_prefix(original_document_number)
    return PREFIX_TO_DOCUMENT_TYPE.get(prefix, "UNKNOWN")


def should_split_version(original_document_number: object) -> bool:
    value = normalize_text(original_document_number)
    prefix = get_prefix(value)

    if not value:
        return False

    if prefix in VERSION_SPLIT_EXCLUDED_PREFIXES:
        return False

    if value.startswith("SOP-") and not re.fullmatch(r"SOP-\d{4,}-\d{6}", value):
        return False

    if prefix == "PRO" and not re.fullmatch(r"PRO-\d{4,}-\d{6}", value):
        return False

    return bool(re.search(r"-\d{6}$", value))


def split_document_number_and_version(original_document_number: object) -> tuple[str, str]:
    original_value = normalize_text(original_document_number)

    if not original_value:
        return "", ""

    if not should_split_version(original_value):
        return original_value, ""

    base_value, version_text = original_value.rsplit("-", 1)
    version = str(int(version_text))

    if original_value[0].isdigit():
        return f"SOP-{base_value}", version

    return base_value, version


def find_header_cell(worksheet) -> tuple[int | None, int | None, str | None]:
    used_range = worksheet.UsedRange
    max_row = min(HEADER_SEARCH_ROWS, used_range.Rows.Count)
    max_col = used_range.Columns.Count

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            value = worksheet.Cells(row, col).Value
            normalized = normalize_text(value).lower()

            if normalized in DOCUMENT_NUMBER_HEADERS:
                return row, col, normalized

    return None, None, None


def header_value(worksheet, row: int, col: int) -> str:
    if col < 1:
        return ""

    return normalize_text(worksheet.Cells(row, col).Value).lower()


def ensure_document_type_column(worksheet, header_row: int, document_col: int) -> tuple[int, int]:
    if header_value(worksheet, header_row, document_col - 1) == "document type":
        worksheet.Cells(header_row, document_col - 1).Value = "Document Type"
        return document_col - 1, document_col

    worksheet.Columns(document_col).Insert()
    worksheet.Cells(header_row, document_col).Value = "Document Type"

    return document_col, document_col + 1


def ensure_right_side_columns(worksheet, header_row: int, document_col: int) -> tuple[int, int]:
    next_header = header_value(worksheet, header_row, document_col + 1)
    second_next_header = header_value(worksheet, header_row, document_col + 2)

    if next_header == "version" and second_next_header == "original tw file name":
        worksheet.Cells(header_row, document_col + 1).Value = "Version"
        worksheet.Cells(header_row, document_col + 2).Value = "Original TW File name"
        return document_col + 1, document_col + 2

    if next_header != "version":
        worksheet.Columns(document_col + 1).Insert()
        worksheet.Cells(header_row, document_col + 1).Value = "Version"

    if header_value(worksheet, header_row, document_col + 2) != "original tw file name":
        worksheet.Columns(document_col + 2).Insert()
        worksheet.Cells(header_row, document_col + 2).Value = "Original TW File name"

    worksheet.Cells(header_row, document_col + 1).Value = "Version"
    worksheet.Cells(header_row, document_col + 2).Value = "Original TW File name"

    return document_col + 1, document_col + 2


def get_last_used_col(worksheet, header_row: int) -> int:
    return worksheet.Cells(header_row, worksheet.Columns.Count).End(-4159).Column


def get_last_used_row(worksheet, document_col: int) -> int:
    return worksheet.Cells(worksheet.Rows.Count, document_col).End(-4162).Row


def get_headers(worksheet, header_row: int, last_col: int) -> list[str]:
    headers = []

    for col in range(1, last_col + 1):
        header = normalize_text(worksheet.Cells(header_row, col).Value)

        if not header:
            header = f"Column {col}"

        headers.append(header)

    return headers


def process_worksheet(
    worksheet,
    workbook_name: str,
    file_index: int,
    file_count: int,
    completed_rows_before_sheet: int,
    total_rows_all_files: int,
    combine_rows: list[dict[str, str]],
    combined_headers: list[str],
    combine_enabled: bool,
) -> tuple[bool, int]:
    print(f"    Checking sheet: {worksheet.Name}", flush=True)

    header_row, document_col, header_name = find_header_cell(worksheet)

    if header_row is None or document_col is None:
        print("    No Training Number or Document Number column found on this sheet.", flush=True)
        return False, 0

    print(f"    Found '{header_name}' on row {header_row}, column {document_col}.", flush=True)

    worksheet.Cells(header_row, document_col).Value = "Document Number"

    document_type_col, document_col = ensure_document_type_column(
        worksheet,
        header_row,
        document_col,
    )

    version_col, original_col = ensure_right_side_columns(
        worksheet,
        header_row,
        document_col,
    )

    last_row = get_last_used_row(worksheet, document_col)
    total_rows = max(0, last_row - header_row)

    print(f"    Last data row: {last_row}", flush=True)
    print(f"    Rows to process: {total_rows:,}", flush=True)

    if total_rows == 0:
        print("    No data rows to process.", flush=True)
        return True, 0

    print("    Reading document numbers into memory...", flush=True)

    source_range = worksheet.Range(
        worksheet.Cells(header_row + 1, document_col),
        worksheet.Cells(last_row, document_col),
    )

    raw_values = source_range.Value

    if total_rows == 1:
        raw_values = ((raw_values,),)

    document_type_values = []
    document_number_values = []
    version_values = []
    original_values = []

    print("    Transforming rows in Python...", flush=True)

    for index, row_tuple in enumerate(raw_values, start=1):
        original_value = normalize_text(row_tuple[0])
        document_type = get_document_type(original_value)
        cleaned_document_number, version = split_document_number_and_version(original_value)

        document_type_values.append((document_type,))
        document_number_values.append((cleaned_document_number,))
        version_values.append((version,))
        original_values.append((original_value,))

        if index % ROW_PRINT_INTERVAL == 0 or index == total_rows:
            overall_completed = completed_rows_before_sheet + index
            sheet_pct = percent(index, total_rows)
            overall_pct = percent(overall_completed, total_rows_all_files)

            print_progress(
                f"        [{file_index}/{file_count}] {worksheet.Name}: "
                f"{index:,} / {total_rows:,} rows | "
                f"{sheet_pct:.2f}% sheet | "
                f"{overall_pct:.2f}% overall",
                same_line=True,
            )

    clear_progress_line()
    print(
        f"        [{file_index}/{file_count}] {worksheet.Name}: "
        f"{total_rows:,} / {total_rows:,} rows | "
        f"100.00% sheet | "
        f"{percent(completed_rows_before_sheet + total_rows, total_rows_all_files):.2f}% overall",
        flush=True,
    )

    print("    Writing transformed columns back to Excel in bulk...", flush=True)

    worksheet.Range(
        worksheet.Cells(header_row + 1, document_type_col),
        worksheet.Cells(last_row, document_type_col),
    ).Value = tuple(document_type_values)

    worksheet.Range(
        worksheet.Cells(header_row + 1, document_col),
        worksheet.Cells(last_row, document_col),
    ).Value = tuple(document_number_values)

    worksheet.Range(
        worksheet.Cells(header_row + 1, version_col),
        worksheet.Cells(last_row, version_col),
    ).Value = tuple(version_values)

    worksheet.Range(
        worksheet.Cells(header_row + 1, original_col),
        worksheet.Cells(last_row, original_col),
    ).Value = tuple(original_values)

    print("    Bulk write complete.", flush=True)

    print("    Autofitting updated columns...", flush=True)

    worksheet.Columns(document_type_col).AutoFit()
    worksheet.Columns(document_col).AutoFit()
    worksheet.Columns(version_col).AutoFit()
    worksheet.Columns(original_col).AutoFit()

    if combine_enabled:
        print("    Adding rows from this sheet to combined CSV data...", flush=True)

        last_col = get_last_used_col(worksheet, header_row)
        sheet_headers = get_headers(worksheet, header_row, last_col)

        for standard_header in ["Source Workbook", "Source Worksheet"]:
            if standard_header not in combined_headers:
                combined_headers.append(standard_header)

        for header in sheet_headers:
            if header not in combined_headers:
                combined_headers.append(header)

        print("    Reading full sheet data for CSV in bulk...", flush=True)

        full_data_range = worksheet.Range(
            worksheet.Cells(header_row + 1, 1),
            worksheet.Cells(last_row, last_col),
        )

        full_data = full_data_range.Value

        if total_rows == 1:
            full_data = (full_data,)

        for index, row_values_tuple in enumerate(full_data, start=1):
            row_values = [normalize_text(value) for value in row_values_tuple]

            if not any(row_values):
                continue

            row_dict = {
                "Source Workbook": workbook_name,
                "Source Worksheet": worksheet.Name,
            }

            for header, value in zip(sheet_headers, row_values):
                row_dict[header] = value

            combine_rows.append(row_dict)

            if index % ROW_PRINT_INTERVAL == 0 or index == total_rows:
                print_progress(
                    f"        [{file_index}/{file_count}] {worksheet.Name}: "
                    f"collected {index:,} / {total_rows:,} CSV rows | "
                    f"{percent(index, total_rows):.2f}% sheet CSV collection",
                    same_line=True,
                )

        clear_progress_line()
        print(f"    Combined CSV rows collected so far: {len(combine_rows):,}", flush=True)

    print(f"    Finished sheet: {worksheet.Name}", flush=True)

    return True, total_rows


def count_processable_rows_in_workbook(excel, workbook_path: Path, file_index: int, file_count: int) -> int:
    print(f"[{file_index}/{file_count}] Pre-scan opening workbook for row count: {workbook_path.name}", flush=True)

    workbook = excel.Workbooks.Open(str(workbook_path))
    workbook_row_count = 0

    try:
        sheet_count = workbook.Worksheets.Count

        for sheet_index in range(1, sheet_count + 1):
            worksheet = workbook.Worksheets(sheet_index)
            header_row, document_col, _ = find_header_cell(worksheet)

            if header_row is None or document_col is None:
                continue

            last_row = get_last_used_row(worksheet, document_col)
            workbook_row_count += max(0, last_row - header_row)

    finally:
        workbook.Close(SaveChanges=False)

    print(f"[{file_index}/{file_count}] Pre-scan rows found: {workbook_row_count:,}", flush=True)

    return workbook_row_count


def pre_scan_total_rows(excel, excel_files: list[Path]) -> int:
    print("", flush=True)
    print("=" * 80, flush=True)
    print("Pre-scanning files to calculate total row count for percent-complete ticker.", flush=True)

    total_rows = 0
    file_count = len(excel_files)

    for file_index, source_path in enumerate(excel_files, start=1):
        total_rows += count_processable_rows_in_workbook(
            excel,
            source_path,
            file_index,
            file_count,
        )

        print(
            f"Pre-scan progress: {file_index:,} / {file_count:,} files | "
            f"{percent(file_index, file_count):.2f}% files | "
            f"{total_rows:,} rows found so far",
            flush=True,
        )

    print(f"Pre-scan complete. Total processable rows: {total_rows:,}", flush=True)

    return total_rows


def process_workbook(
    excel,
    workbook_path: Path,
    file_index: int,
    file_count: int,
    completed_rows_before_file: int,
    total_rows_all_files: int,
    combine_rows: list[dict[str, str]],
    combined_headers: list[str],
    combine_enabled: bool,
) -> tuple[int, int]:
    print(f"[{file_index}/{file_count}] Opening workbook: {workbook_path.name}", flush=True)

    workbook = excel.Workbooks.Open(str(workbook_path))
    changed_sheet_count = 0
    rows_processed_in_workbook = 0

    try:
        sheet_count = workbook.Worksheets.Count
        print(f"[{file_index}/{file_count}] Workbook has {sheet_count} sheet(s).", flush=True)

        for sheet_index in range(1, sheet_count + 1):
            worksheet = workbook.Worksheets(sheet_index)

            print(
                f"[{file_index}/{file_count}] Processing sheet {sheet_index}/{sheet_count}: {worksheet.Name}",
                flush=True,
            )

            changed, rows_processed_in_sheet = process_worksheet(
                worksheet,
                workbook_path.name,
                file_index,
                file_count,
                completed_rows_before_file + rows_processed_in_workbook,
                total_rows_all_files,
                combine_rows,
                combined_headers,
                combine_enabled,
            )

            if changed:
                changed_sheet_count += 1

            rows_processed_in_workbook += rows_processed_in_sheet

        print(f"[{file_index}/{file_count}] Saving workbook: {workbook_path.name}", flush=True)
        workbook.Save()
        print(f"[{file_index}/{file_count}] Saved workbook: {workbook_path.name}", flush=True)

    finally:
        workbook.Close(SaveChanges=False)
        print(f"[{file_index}/{file_count}] Closed workbook: {workbook_path.name}", flush=True)

    return changed_sheet_count, rows_processed_in_workbook


def write_combined_csv(output_folder: Path, combined_headers: list[str], combine_rows: list[dict[str, str]]) -> None:
    csv_path = output_folder / COMBINED_CSV_NAME

    print("", flush=True)
    print("=" * 80, flush=True)
    print(f"Writing combined CSV: {csv_path}", flush=True)
    print(f"Combined row count: {len(combine_rows):,}", flush=True)

    with csv_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=combined_headers, extrasaction="ignore")
        writer.writeheader()

        for index, row in enumerate(combine_rows, start=1):
            writer.writerow(row)

            if index % 5000 == 0 or index == len(combine_rows):
                print_progress(
                    f"    Wrote {index:,} / {len(combine_rows):,} CSV rows | "
                    f"{percent(index, len(combine_rows)):.2f}% CSV complete",
                    same_line=True,
                )

    clear_progress_line()
    print(f"Combined CSV created: {csv_path}", flush=True)


def main() -> None:
    source_folder = Path(FOLDER_PATH)
    output_folder = source_folder / OUTPUT_FOLDER_NAME
    combine_enabled = normalize_combine_setting(COMBINE)

    output_folder.mkdir(exist_ok=True)

    excel_files = [
        path for path in source_folder.iterdir()
        if path.suffix.lower() in {".xls", ".xlsx", ".xlsm"}
        and not path.name.startswith("~$")
        and path.parent != output_folder
    ]

    file_count = len(excel_files)
    combine_rows = []
    combined_headers = []

    print(f"Source folder: {source_folder}", flush=True)
    print(f"Output folder: {output_folder}", flush=True)
    print(f"Combine: {'Yes' if combine_enabled else 'No'}", flush=True)
    print(f"Found {file_count} Excel file(s).", flush=True)

    if file_count == 0:
        print("No Excel files found. Nothing to process.", flush=True)
        return

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    safely_set_excel_property(
        excel=excel,
        property_name="ScreenUpdating",
        value=False,
        warning_label="disable ScreenUpdating",
    )

    safely_set_excel_property(
        excel=excel,
        property_name="EnableEvents",
        value=False,
        warning_label="disable EnableEvents",
    )

    safely_set_excel_property(
        excel=excel,
        property_name="Calculation",
        value=-4135,  # xlCalculationManual
        warning_label="set Excel calculation to manual",
    )

    completed_rows = 0

    try:
        total_rows_all_files = pre_scan_total_rows(excel, excel_files)

        print("", flush=True)
        print("=" * 80, flush=True)
        print("Beginning workbook updates.", flush=True)

        for file_index, source_path in enumerate(excel_files, start=1):
            output_path = output_folder / source_path.name

            print("", flush=True)
            print("=" * 80, flush=True)
            print(
                f"[{file_index}/{file_count}] Starting file | "
                f"{percent(file_index - 1, file_count):.2f}% files complete | "
                f"{percent(completed_rows, total_rows_all_files):.2f}% rows complete",
                flush=True,
            )
            print(f"[{file_index}/{file_count}] Copying: {source_path.name}", flush=True)

            shutil.copy2(source_path, output_path)

            changed_sheet_count, rows_processed_in_workbook = process_workbook(
                excel,
                output_path,
                file_index,
                file_count,
                completed_rows,
                total_rows_all_files,
                combine_rows,
                combined_headers,
                combine_enabled,
            )

            completed_rows += rows_processed_in_workbook

            if changed_sheet_count:
                print(
                    f"[{file_index}/{file_count}] DONE. Updated {changed_sheet_count} sheet(s): {output_path}",
                    flush=True,
                )
            else:
                print(
                    f"[{file_index}/{file_count}] DONE. No matching Document Number or Training Number column found: {output_path}",
                    flush=True,
                )

            print(
                f"[{file_index}/{file_count}] Progress checkpoint | "
                f"{percent(file_index, file_count):.2f}% files complete | "
                f"{completed_rows:,} / {total_rows_all_files:,} rows | "
                f"{percent(completed_rows, total_rows_all_files):.2f}% rows complete",
                flush=True,
            )

    finally:
        safely_set_excel_property(
            excel=excel,
            property_name="Calculation",
            value=-4105,  # xlCalculationAutomatic
            warning_label="restore Excel calculation mode",
        )

        safely_set_excel_property(
            excel=excel,
            property_name="EnableEvents",
            value=True,
            warning_label="restore EnableEvents",
        )

        safely_set_excel_property(
            excel=excel,
            property_name="ScreenUpdating",
            value=True,
            warning_label="restore ScreenUpdating",
        )

        excel.Quit()
        print("", flush=True)
        print("Excel closed.", flush=True)

    if combine_enabled:
        write_combined_csv(output_folder, combined_headers, combine_rows)

    print("", flush=True)
    print("=" * 80, flush=True)
    print("Processing complete.", flush=True)
    print(f"Files processed: {file_count:,}", flush=True)
    print(f"Rows processed: {completed_rows:,}", flush=True)
    print(f"Output folder: {output_folder}", flush=True)

    if combine_enabled:
        print(f"Combined CSV: {output_folder / COMBINED_CSV_NAME}", flush=True)


if __name__ == "__main__":
    main()