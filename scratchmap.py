import os
import csv

# Paths (update if needed)
csv_path = r"C:\Users\e.muraj\Downloads\Trackwise_Exports\Trackwise_Exports_12APR\Unzipped\WE_00D6g000005ilBVEAY_1\ContentDocumentLink.csv"
search_root = r"C:\Users\e.muraj\Downloads\Trackwise_Exports\Trackwise_Exports_12APR\Unzipped"


def get_all_filenames_and_dirs(root_folder):
    filenames = set()
    checked_dirs = []
    for dirpath, _, files in os.walk(root_folder):
        checked_dirs.append(dirpath)
        for file in files:
            filenames.add(file)
            name_no_ext = os.path.splitext(file)[0]
            filenames.add(name_no_ext)
    return filenames, checked_dirs


def get_csv_ids(csv_path):
    ids = set()
    with open(csv_path, newline='', encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            for key in ['Id', 'LinkedEntityId', 'ContentDocumentId']:
                if row.get(key):
                    ids.add(row[key].strip())
    return ids


if __name__ == "__main__":
    print("Scanning for blob file matches in exported TrackWise DMS folders...\n")
    blob_filenames, checked_dirs = get_all_filenames_and_dirs(search_root)
    csv_ids = get_csv_ids(csv_path)

    found = []
    not_found = []
    for id_val in csv_ids:
        if id_val in blob_filenames:
            print(f"Found as filename/blob: {id_val}")
            found.append(id_val)
        else:
            not_found.append(id_val)

    print("\nFolders checked during scan:")
    for folder in checked_dirs:
        print(f"  - {folder}")

    print("\nSummary:")
    if found:
        print("The following values from CSV were found as blob filenames (with or without extension):")
        for v in found:
            print(f"  - {v}")
    else:
        print("No CSV values were found as blob filenames.")

    print(f"\nTotal CSV IDs checked: {len(csv_ids)}")
    print(f"Total matches found: {len(found)}")
    print(f"Total not found: {len(not_found)}")
