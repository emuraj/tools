# csv_snippet_sampler.py

import os
import csv
from tqdm import tqdm

SOURCE_DIR = r"C:\Users\e.muraj\Downloads\Trackwise_Exports\Trackwise_Exports_12APR\Unzipped\WE_00D6g000005ilBVEAY_1"
OUTPUT_FILE = r"C:\Users\e.muraj\OneDrive - Neurotech USA, Inc\TW_Export\csv_snippet_catalog.csv"
SAMPLE_ROWS = 2  # How many rows per CSV do we want?

# We'll also reuse the fallback logic you have in parse_csv
def try_parse_csv(file_path, max_rows=SAMPLE_ROWS, encodings=("utf-8-sig", "cp1252", "latin-1")):
    """
    Attempts multiple encodings, returns up to 'max_rows' from the CSV as a list of dicts.
    """
    for enc in encodings:
        try:
            with open(file_path, encoding=enc, errors="replace", newline="") as f:
                reader = csv.DictReader(f)
                result = []
                row_count = 0
                for row in reader:
                    result.append(row)
                    row_count += 1
                    if row_count >= max_rows:
                        break
                return result
        except:
            continue
    return []

def main():
    # We'll store all data in a list-of-rows approach, then write out at the end.
    # Our final columns: ["Source CSV", "Row #", "Header 1", "Header 2", ...]
    # But we can't know all possible headers in advance,
    # so we may do a "largest union" approach or just write out as we go in some flexible manner.

    combined_rows = []
    all_headers = set()

    csv_files = [f for f in os.listdir(SOURCE_DIR) if f.lower().endswith(".csv")]
    print(f"Found {len(csv_files)} CSV files in {SOURCE_DIR}...")

    # We'll gather partial data from each.
    for csv_file in tqdm(csv_files, desc="Sampling CSVs", unit="csv"):
        full_path = os.path.join(SOURCE_DIR, csv_file)
        snippet = try_parse_csv(full_path, max_rows=SAMPLE_ROWS)
        # snippet is a list of up to 2 row-dicts
        for i, row_dict in enumerate(snippet, start=1):
            # track all headers
            all_headers.update(row_dict.keys())
            # We'll store each snippet row with a "Source CSV" and "RowInSource" prefix
            row_dict["_SourceCSV"] = csv_file
            row_dict["_RowInSource"] = i
            combined_rows.append(row_dict)

    # Now write out the combined data.
    # all_headers might not be in the same order for each row, but let's do a sorted approach.
    # And let's put _SourceCSV, _RowInSource at the front.
    final_headers = ["_SourceCSV", "_RowInSource"] + sorted(all_headers - {"_SourceCSV", "_RowInSource"})

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as out_f:
        writer = csv.DictWriter(out_f, fieldnames=final_headers)
        writer.writeheader()
        for row_dict in combined_rows:
            writer.writerow(row_dict)

    print(f"✅ Wrote snippet of each CSV (max {SAMPLE_ROWS} rows) to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
