import subprocess
from pathlib import Path

exe_path = Path(r"C:\Users\e.muraj\Downloads\IntellectBatchFileUpload.exe")

try:
    subprocess.run([str(exe_path)], check=True)
except FileNotFoundError:
    print(f"Executable not found: {exe_path}")
except PermissionError as e:
    print(f"Permission denied: {e}")
except subprocess.CalledProcessError as e:
    print(f"Program exited with error code: {e.returncode}")