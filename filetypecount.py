import os
from collections import Counter

def count_filetypes(folder_path):
    # Extensions we care about
    extensions = {".json", ".docx", ".ppt", ".xlsx"}
    counts = Counter()

    for root, _, files in os.walk(folder_path):
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in extensions:
                counts[ext] += 1

    # Print results
    for ext in extensions:
        print(f"{ext} = {counts.get(ext, 0)}")

if __name__ == "__main__":
    folder = r"C:\path\to\your\folder"  # <-- change this to your folder path
    count_filetypes(folder)
