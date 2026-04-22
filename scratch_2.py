#!/usr/bin/env python3
# map_trays_path_explicit.py — prints "filename = Tray X" for files inside subfolders named "Tray X"
# 1) >>>> PASTE YOUR FOLDER PATH BETWEEN THE QUOTES BELOW <<<<
#    Examples:
#      Windows: r"C:\Users\you\Desktop\TempTales"
#      macOS/Linux: "/Users/you/Desktop/TempTales"
ROOT_PATH = r"C:\Users\e.muraj\Desktop\temptales\all_2"

# Optional: file extension filter (set to "" to include all files)
FILE_EXT = ".pdf"   # examples: ".pdf", ".csv", "" for all files

import re
from pathlib import Path
import sys

def main():
    if not ROOT_PATH or "<<<" in ROOT_PATH:
        print("Please set ROOT_PATH at the top of the script to your parent folder path.")
        sys.exit(1)

    parent = Path(ROOT_PATH).expanduser().resolve()
    if not parent.exists() or not parent.is_dir():
        print(f"Path not found or not a directory: {parent}")
        sys.exit(1)

    tray_re = re.compile(r"^Tray\s+\d+$", re.IGNORECASE)

    # Find tray subfolders like "Tray 8", sort by tray number
    trays = sorted(
        (p for p in parent.iterdir() if p.is_dir() and tray_re.match(p.name)),
        key=lambda p: int(re.search(r"\d+", p.name).group())
    )

    if not trays:
        print(f"No tray folders found in {parent} (looking for folders named like 'Tray 8').")
        sys.exit(0)

    for tray in trays:
        files = []
        if FILE_EXT:
            files = sorted(tray.glob(f"*{FILE_EXT}"))
        else:
            files = sorted(f for f in tray.iterdir() if f.is_file())

        for f in files:
            # Print just the filename on the left, and the tray folder name on the right
            print(f"{f.name} = {tray.name}")

if __name__ == "__main__":
    main()
