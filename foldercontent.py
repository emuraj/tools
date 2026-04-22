#!/usr/bin/env python3

from pathlib import Path
import sys
from typing import Union, List

# List of default root directories (each as a separate raw string)
DEFAULT_ROOTS = [
    r"D:\laptop software\Production\1_TrackWise Raw Export Repository\TrackWise Raw QMS Export\Unzipped\WE_00D6g000005ilBVEAY_1"
]

def print_tree(path: Union[str, Path], indent: str = "") -> None:
    """
    Recursively print folders and files in a tree-like format.
    """
    p = Path(path)
    marker = "/" if p.is_dir() else ""
    print(f"{indent}{p.name}{marker}")

    if p.is_dir():
        children = sorted(p.iterdir(), key=lambda c: (c.is_file(), c.name.lower()))
        for child in children:
            print_tree(child, indent + "    ")

def main() -> None:
    # Use command-line paths if provided; otherwise, use all defaults
    roots: List[str] = sys.argv[1:] if len(sys.argv) > 1 else DEFAULT_ROOTS

    for root_str in roots:
        root = Path(root_str)
        if not root.exists():
            sys.stderr.write(f"Error: Path does not exist → {root}\n")
            continue

        header = f"Folder structure for: {root}"
        print(header)
        print("=" * len(header))
        print_tree(root)
        print()  # Blank line between directory trees

if __name__ == "__main__":
    main()
