# count_non_csv_files.py

import os

def count_non_csv_files_by_folder(directory):
    total_count = 0

    print("\n📌 Purpose: This script scans the specified directory and all its subdirectories,")
    print("   reporting the number of non-CSV files found in each folder.\n")
    print(f"📁 Starting directory scan:\n  {directory}")
    print("────────────────────────────────────────────")

    for root, _, files in os.walk(directory):
        non_csv_files = [file for file in files if not file.lower().endswith('.csv')]
        count = len(non_csv_files)
        if count > 0:
            print(f"\n📂 Folder: {root}")
            print(f"  → Non-CSV files in this folder: {count}")
        total_count += count

    print("\n────────────────────────────────────────────")
    print(f"\n📊 Total non-CSV files across all folders: {total_count}\n")

if __name__ == "__main__":
    folder_path = r"C:\Users\e.muraj\Downloads\Trackwise_Exports\Trackwise_Exports_12APR\Unzipped"
    if os.path.isdir(folder_path):
        count_non_csv_files_by_folder(folder_path)
    else:
        print("\n❌ Invalid directory path.\n")
