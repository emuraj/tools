# print_trackwise_export_contents.py
"""
Purpose:
    Print a recursive, human-readable inventory of the TrackWise DMS and QMS
    export folders.

What this file does:
    - Walks the DMS export root and QMS export root
    - Prints folder and file contents in a tree-like layout
    - Includes file sizes
    - Writes separate text reports for DMS and QMS
    - Lets you limit depth if the trees are too large

Place in the larger scheme:
    This is a support utility for preparing concrete examples for the
    Neurotic GMP technical report. It does not modify source data.
"""

from __future__ import annotations

from pathlib import Path
from datetime import datetime


# ============================================================================
# EDIT THESE PATHS
# ============================================================================
DMS_EXPORT_ROOT = r"C:\Users\e.muraj\Desktop\Production\1_TrackWise Raw Export Repository\TrackWise Raw DMS Export"
QMS_EXPORT_ROOT = r"C:\Users\e.muraj\Desktop\Production\1_TrackWise Raw Export Repository\TrackWise Raw QMS Export\Unzipped"

# Optional: set to None for full recursion, or 2 / 3 / 4 for shallower output
MAX_DEPTH: int | None = 3

# Output folder for the text reports
OUTPUT_DIR = r"C:\Users\e.muraj\Desktop\DEV\NEUROTIC\example_exports"
# ============================================================================


def human_size(num_bytes: int) -> str:
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(num_bytes)
    for unit in units:
        if size < 1024.0 or unit == units[-1]:
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{num_bytes} B"


def safe_iterdir(path: Path) -> list[Path]:
    try:
        return sorted(path.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower()))
    except Exception:
        return []


def build_tree_lines(root: Path, max_depth: int | None = None) -> list[str]:
    lines: list[str] = []
    if not root.exists():
        return [f"[MISSING] {root}"]

    lines.append(f"ROOT: {root}")
    lines.append("")

    def walk(current: Path, prefix: str, depth: int) -> None:
        if max_depth is not None and depth > max_depth:
            return

        children = safe_iterdir(current)
        for idx, child in enumerate(children):
            is_last = idx == len(children) - 1
            branch = "└── " if is_last else "├── "

            if child.is_dir():
                lines.append(f"{prefix}{branch}[DIR]  {child.name}")
                extension = "    " if is_last else "│   "
                walk(child, prefix + extension, depth + 1)
            else:
                try:
                    size = human_size(child.stat().st_size)
                except Exception:
                    size = "size-unavailable"
                lines.append(f"{prefix}{branch}[FILE] {child.name}  ({size})")

    walk(root, prefix="", depth=1)
    return lines


def write_report(root_path: str, label: str, output_dir: Path, max_depth: int | None) -> Path:
    root = Path(root_path).expanduser()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir.mkdir(parents=True, exist_ok=True)
    out_file = output_dir / f"{label}_export_contents_{timestamp}.txt"

    lines = build_tree_lines(root, max_depth=max_depth)
    out_file.write_text("\n".join(lines), encoding="utf-8")

    return out_file


def main() -> None:
    output_dir = Path(OUTPUT_DIR).expanduser()

    dms_report = write_report(
        root_path=DMS_EXPORT_ROOT,
        label="DMS",
        output_dir=output_dir,
        max_depth=MAX_DEPTH,
    )
    qms_report = write_report(
        root_path=QMS_EXPORT_ROOT,
        label="QMS",
        output_dir=output_dir,
        max_depth=MAX_DEPTH,
    )

    print(f"DMS report written to: {dms_report}")
    print(f"QMS report written to: {qms_report}")
    print("")
    print("Done.")


if __name__ == "__main__":
    main()