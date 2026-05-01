# run_intellect_batch_upload.py
"""
Purpose:
Run the IntellectBatchFileUpload executable from Python.

What this file does:
Starts the local Windows executable and waits for it to finish.

Place in the larger scheme:
Use this as a PyCharm-run wrapper for launching the Intellect batch upload tool.
"""

import subprocess
from pathlib import Path


EXE_PATH = Path(r"C:\Users\e.muraj\Downloads\IntellectBatchFileUpload.exe")


def main() -> None:
    if not EXE_PATH.exists():
        raise FileNotFoundError(f"Executable not found: {EXE_PATH}")

    result = subprocess.run(
        [str(EXE_PATH)],
        check=False,
    )

    print(f"Process exited with code: {result.returncode}")


if __name__ == "__main__":
    main()