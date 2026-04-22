# print_folder_structure.py

import os

def print_folder_structure(root_folder, indent=""):
    try:
        entries = sorted(os.listdir(root_folder))
    except Exception as e:
        print(f"{indent}[Error accessing folder: {e}]")
        return

    for entry in entries:
        full_path = os.path.join(root_folder, entry)
        if os.path.isdir(full_path):
            print(f"{indent}📁 {entry}/")
            print_folder_structure(full_path, indent + "    ")
        else:
            print(f"{indent}📄 {entry}")

if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog

    # Use a file picker to select the root folder
    tk.Tk().withdraw()
    folder_path = filedialog.askdirectory(title="Select Root Folder to Print Structure")

    if folder_path:
        print(f"Folder Structure for: {folder_path}")
        print_folder_structure(folder_path)
    else:
        print("No folder selected.")
