#!/usr/bin/env python3
"""
docx_to_pdf_gui.py

Quick GUI to convert a single Word document (.docx/.doc) to PDF using MS Word.
Windows + Microsoft Word required.

Install:
  pip install pywin32
"""

from __future__ import annotations

import os
import queue
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from pathlib import Path

import win32com.client

WD_EXPORT_FORMAT_PDF = 17  # Word constant


def convert_word_to_pdf(input_path: str, output_path: str, progress_cb) -> None:
    """
    progress_cb(message: str, percent: int) -> None
    """
    in_path = Path(input_path).resolve()
    if not in_path.exists():
        raise FileNotFoundError(f"Input not found: {in_path}")

    out_path = Path(output_path).resolve()
    if out_path.suffix.lower() != ".pdf":
        out_path = out_path.with_suffix(".pdf")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    progress_cb("Validating paths…", 5)

    word = None
    doc = None
    try:
        progress_cb("Launching Microsoft Word…", 15)
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        progress_cb("Opening document…", 35)
        # Word COM can take a moment; this call is the “real” work.
        doc = word.Documents.Open(str(in_path), ReadOnly=True)

        progress_cb("Exporting to PDF…", 70)
        doc.ExportAsFixedFormat(
            OutputFileName=str(out_path),
            ExportFormat=WD_EXPORT_FORMAT_PDF,
            OpenAfterExport=False,
            OptimizeFor=0,      # 0=Print, 1=OnScreen
            CreateBookmarks=1,  # 0=None, 1=Headings, 2=Bookmarks
        )

        progress_cb("Finalizing…", 95)

    finally:
        # Close cleanly even if export errors
        try:
            if doc is not None:
                progress_cb("Closing document…", 97)
                doc.Close(False)
        finally:
            if word is not None:
                progress_cb("Closing Word…", 99)
                word.Quit()

    progress_cb(f"Complete. PDF written: {out_path}", 100)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Word → PDF Converter")
        self.geometry("820x420")
        self.minsize(760, 380)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()

        self.percent_var = tk.IntVar(value=0)
        self.status_var = tk.StringVar(value="Select input and output, then Convert.")

        self._ui_queue: queue.Queue[tuple[str, int] | tuple[str, str]] = queue.Queue()
        # queue messages:
        #   ("progress", message, percent) encoded as ("p", message, percent_str)
        #   ("done", output_pdf)
        #   ("error", error_text)

        self._build_ui()
        self.after(100, self._poll_ui_queue)

    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 8}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True)

        # Input
        ttk.Label(frm, text="Input Word file (.docx/.doc):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.input_var).grid(row=1, column=0, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.browse_input).grid(row=1, column=1, sticky="ew", padx=6, pady=8)

        # Output
        ttk.Label(frm, text="Output PDF path:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.output_var).grid(row=3, column=0, sticky="ew", **pad)
        ttk.Button(frm, text="Save As…", command=self.browse_output).grid(row=3, column=1, sticky="ew", padx=6, pady=8)

        # Progress row
        prog_row = ttk.Frame(frm)
        prog_row.grid(row=4, column=0, columnspan=2, sticky="ew", padx=12, pady=(8, 4))
        prog_row.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            prog_row,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            variable=self.percent_var,
        )
        self.progress.grid(row=0, column=0, sticky="ew")

        self.percent_label = ttk.Label(prog_row, text="0%")
        self.percent_label.grid(row=0, column=1, sticky="e", padx=(10, 0))

        # Actions
        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=(6, 2))
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)

        self.convert_btn = ttk.Button(btns, text="Convert", command=self.start_convert)
        self.convert_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        ttk.Button(btns, text="Open Output Folder", command=self.open_output_folder).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )

        # Status
        ttk.Separator(frm).grid(row=6, column=0, columnspan=2, sticky="ew", padx=12, pady=8)
        ttk.Label(frm, textvariable=self.status_var).grid(
            row=7, column=0, columnspan=2, sticky="w", padx=12, pady=(0, 6)
        )

        # Live log
        ttk.Label(frm, text="Live progress log:").grid(row=8, column=0, sticky="w", padx=12, pady=(6, 2))
        self.log_box = ScrolledText(frm, height=10, wrap="word")
        self.log_box.grid(row=9, column=0, columnspan=2, sticky="nsew", padx=12, pady=(0, 12))
        self.log_box.configure(state="disabled")

        frm.columnconfigure(0, weight=1)
        frm.columnconfigure(1, weight=0)
        frm.rowconfigure(9, weight=1)

    def _append_log(self, line: str) -> None:
        self.log_box.configure(state="normal")
        ts = time.strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{ts}] {line}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_progress(self, message: str, percent: int) -> None:
        percent = max(0, min(100, int(percent)))
        self.percent_var.set(percent)
        self.percent_label.config(text=f"{percent}%")
        self.status_var.set(message)
        self._append_log(message)

    def _poll_ui_queue(self) -> None:
        try:
            while True:
                item = self._ui_queue.get_nowait()
                # item formats:
                # ("p", message, "percent")
                # ("done", out_pdf)
                # ("error", error_text)
                kind = item[0]
                if kind == "p":
                    _, msg, pct_str = item
                    self._set_progress(msg, int(pct_str))
                elif kind == "done":
                    _, out_pdf = item
                    self.convert_btn.config(state="normal")
                    self._set_progress(f"Done. PDF written: {out_pdf}", 100)
                    messagebox.showinfo("Success", f"PDF created:\n{out_pdf}")
                elif kind == "error":
                    _, err_text = item
                    self.convert_btn.config(state="normal")
                    self.status_var.set("Failed.")
                    self._append_log(f"ERROR: {err_text}")
                    messagebox.showerror("Conversion failed", err_text)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll_ui_queue)

    def browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Word document",
            filetypes=[("Word Documents", "*.docx *.doc"), ("All files", "*.*")],
        )
        if not path:
            return
        self.input_var.set(path)

        suggested = str(Path(path).with_suffix(".pdf"))
        if not self.output_var.get():
            self.output_var.set(suggested)

    def browse_output(self) -> None:
        initial = self.output_var.get().strip()
        initial_dir = str(Path(initial).parent) if initial else None
        initial_file = Path(initial).name if initial else None

        path = filedialog.asksaveasfilename(
            title="Choose output PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialdir=initial_dir,
            initialfile=initial_file,
        )
        if not path:
            return
        self.output_var.set(path)

    def start_convert(self) -> None:
        in_path = self.input_var.get().strip()
        out_path = self.output_var.get().strip()

        if not in_path:
            messagebox.showerror("Missing input", "Please select an input Word document.")
            return
        if not out_path:
            messagebox.showerror("Missing output", "Please choose an output PDF path.")
            return

        # Reset UI
        self.convert_btn.config(state="disabled")
        self.percent_var.set(0)
        self.percent_label.config(text="0%")
        self.status_var.set("Starting…")
        self._append_log("Starting conversion…")

        def progress_cb(message: str, percent: int) -> None:
            # marshal progress back to UI thread
            self._ui_queue.put(("p", message, str(percent)))

        def worker() -> None:
            try:
                convert_word_to_pdf(in_path, out_path, progress_cb)
                out_pdf = str(Path(out_path).with_suffix(".pdf").resolve())
                self._ui_queue.put(("done", out_pdf))
            except Exception as e:
                self._ui_queue.put(("error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def open_output_folder(self) -> None:
        out_path = self.output_var.get().strip()
        if not out_path:
            messagebox.showinfo("No output selected", "Choose an output PDF path first.")
            return
        folder = str(Path(out_path).resolve().parent)
        try:
            os.startfile(folder)  # Windows only
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{e}")


if __name__ == "__main__":
    App().mainloop()
