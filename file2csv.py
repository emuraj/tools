import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def build_dms_export(root_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    output_excel = os.path.join(output_folder, "DMS_export.xlsx")

    writer = pd.ExcelWriter(output_excel, engine="openpyxl")

    for parent in os.listdir(root_folder):
        parent_path = os.path.join(root_folder, parent)
        if not os.path.isdir(parent_path):
            continue

        binaries_path = os.path.join(parent_path, "Binaries")
        renditions_path = os.path.join(parent_path, "Renditions")

        binaries, renditions = [], []

        if os.path.isdir(binaries_path):
            binaries = [f for f in os.listdir(binaries_path) if os.path.isfile(os.path.join(binaries_path, f))]
        if os.path.isdir(renditions_path):
            renditions = [f for f in os.listdir(renditions_path) if os.path.isfile(os.path.join(renditions_path, f))]

        max_len = max(len(binaries), len(renditions))
        df = pd.DataFrame({
            "Binaries": binaries + [""] * (max_len - len(binaries)),
            "Renditions": renditions + [""] * (max_len - len(renditions)),
        })

        df.to_excel(writer, sheet_name=parent[:31], index=False)

    writer.close()
    return output_excel


# ---------- GUI ----------
class DMSExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DMS Export Tool")
        self.root.geometry("650x250")  # Larger window

        self.root_folder = tk.StringVar()
        self.output_folder = tk.StringVar()

        # Root folder
        tk.Label(root, text="Root Folder:", anchor="w").pack(fill="x", padx=10, pady=(10, 2))
        root_frame = tk.Frame(root)
        root_frame.pack(fill="x", padx=10)
        self.root_entry = tk.Entry(root_frame, textvariable=self.root_folder, width=70, bg="white")
        self.root_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(root_frame, text="Browse", command=self.select_root).pack(side="right")

        # Output folder
        tk.Label(root, text="Output Folder:", anchor="w").pack(fill="x", padx=10, pady=(15, 2))
        output_frame = tk.Frame(root)
        output_frame.pack(fill="x", padx=10)
        self.output_entry = tk.Entry(output_frame, textvariable=self.output_folder, width=70, bg="white")
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(output_frame, text="Browse", command=self.select_output).pack(side="right")

        # Run button
        tk.Button(root, text="Run Export", command=self.run_export, height=2, width=20, bg="lightblue").pack(pady=20)

    def select_root(self):
        folder = filedialog.askdirectory(title="Select Root Folder")
        if folder:
            self.root_folder.set(folder)

    def select_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder.set(folder)

    def run_export(self):
        if not self.root_folder.get() or not self.output_folder.get():
            messagebox.showerror("Error", "Please select both root and output folders.")
            return

        try:
            output_file = build_dms_export(self.root_folder.get(), self.output_folder.get())
            messagebox.showinfo("Success", f"Export completed!\nFile saved at:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DMSExportApp(root)
    root.mainloop()
