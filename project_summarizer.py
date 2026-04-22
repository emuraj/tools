# project_summarizer.py
"""
Generates a Markdown tree, a CSV manifest with SHA-256, and a JSON summary
for the current project. Run from your project root:
    python project_summarizer.py

Outputs:
  - tree.md
  - manifest.csv
  - summary.json
"""

from __future__ import annotations

import csv
import hashlib
import json
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Counter as TCounter, Iterable, List, Optional, Tuple, Dict

import re
from collections import Counter

# ---------------------- Configuration ----------------------

IGNORE_DIRS = {
    ".git", ".idea", ".venv", "venv", "__pycache__", "build", "dist",
    ".mypy_cache", ".pytest_cache", ".ruff_cache", ".eggs", ".tox"
}
IGNORE_SUFFIXES = {".pyc", ".pyo", ".pyd", ".log", ".tmp", ".ds_store"}

# None means "hash everything". Set to an int to skip hashing files larger than N MB.
MAX_HASH_MB: Optional[int] = None

ROOT = Path.cwd().resolve()

# ---------------------- Helpers ----------------------------

def is_ignored_dir(name: str) -> bool:
    return name in IGNORE_DIRS


def should_skip_file(name: str) -> bool:
    lower = name.lower()
    return any(lower.endswith(suf) for suf in IGNORE_SUFFIXES)


def utc_iso(ts: float) -> str:
    return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def sha256_file(path: Path) -> str:
    if MAX_HASH_MB is not None:
        try:
            if path.stat().st_size > MAX_HASH_MB * 1024 * 1024:
                return f"SKIPPED-LARGER-THAN-{MAX_HASH_MB}MB"
        except OSError:
            return "ERROR"

    h = hashlib.sha256()
    try:
        with path.open("rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return "ERROR"


def rel(p: Path) -> str:
    try:
        return str(p.relative_to(ROOT)).replace("\\", "/")
    except ValueError:
        # Not under ROOT (unlikely when running from ROOT)
        return str(p).replace("\\", "/")

# ---------------------- Collection -------------------------

@dataclass
class FileRow:
    rel_path: str
    type: str
    size_bytes: int
    modified_utc: str
    sha256: str


def collect_tree() -> Tuple[List[Tuple[int, str, bool]], List[FileRow]]:
    rows: List[Tuple[int, str, bool]] = []
    file_rows: List[FileRow] = []

    for cur, dirs, files in os.walk(ROOT):
        # filter ignore dirs in place
        dirs[:] = [d for d in dirs if not is_ignored_dir(d)]

        cur_path = Path(cur)
        try:
            level = len(rel(cur_path).split("/")) if rel(cur_path) != "." else 0
        except Exception:
            level = 0

        name = cur_path.name if level > 0 else ROOT.name
        rows.append((level, name or ".", True))

        for fn in sorted(files, key=str.casefold):
            if should_skip_file(fn):
                continue
            full = cur_path / fn
            rows.append((level + 1, fn, False))
            try:
                st = full.stat()
                file_rows.append(FileRow(
                    rel_path=rel(full),
                    type="FILE",
                    size_bytes=int(st.st_size),
                    modified_utc=utc_iso(st.st_mtime),
                    sha256=sha256_file(full),
                ))
            except OSError:
                file_rows.append(FileRow(
                    rel_path=rel(full),
                    type="FILE",
                    size_bytes=0,
                    modified_utc="",
                    sha256="ERROR",
                ))
    return rows, file_rows

# ---------------------- Writers ----------------------------

def write_tree_md(rows: Iterable[Tuple[int, str, bool]]) -> None:
    out_lines: List[str] = ["# Project Tree", ""]
    for level, name, is_dir in rows:
        indent = "  " * level
        bullet = "📁" if is_dir else "📄"
        out_lines.append(f"{indent}- {bullet} {name}")
    Path(
        "../tree.md").write_text("\n".join(out_lines) + "\n", encoding="utf-8")


def write_manifest_csv(file_rows: Iterable[FileRow]) -> None:
    with Path("../manifest.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["rel_path", "type", "size_bytes", "modified_utc", "sha256"])
        for r in file_rows:
            w.writerow([r.rel_path, r.type, r.size_bytes, r.modified_utc, r.sha256])

# ---------------------- Simple analysis --------------------

# Use raw strings and avoid redundant escapes.
PY_IMPORT_RE = re.compile(
    r'^\s*(?:from\s+([a-zA-Z0-9_\.]+)\s+import|import\s+([a-zA-Z0-9_\.]+))',
    re.M
)
MAIN_GUARD_RE = re.compile(r'if\s+__name__\s*==\s*[\'"]__main__[\'"]\s*:', re.M)
TK_RE = re.compile(r'\b(tkinter|tk)\b|\bfrom\s+tkinter\s+import\b')

def analyze_python_sources(file_rows: Iterable[FileRow]) -> Tuple[Counter, List[str], List[str]]:
    imports: Counter = Counter()
    entry_points: List[str] = []
    tkinter_users: List[str] = []

    for r in file_rows:
        if not r.rel_path.lower().endswith(".py"):
            continue
        path = ROOT / r.rel_path
        try:
            src = path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue

        for m in PY_IMPORT_RE.finditer(src):
            mod = m.group(1) or m.group(2)
            if mod:
                root_mod = mod.split(".")[0]
                imports[root_mod] += 1

        if MAIN_GUARD_RE.search(src):
            entry_points.append(r.rel_path)

        if TK_RE.search(src):
            tkinter_users.append(r.rel_path)

    return imports, entry_points, tkinter_users


def write_summary_json(
    file_rows: Iterable[FileRow],
    imports: Counter,
    entry_points: List[str],
    tkinter_users: List[str]
) -> None:
    by_ext: Counter = Counter(Path(r.rel_path).suffix.lower() for r in file_rows)
    total_size = sum(int(r.size_bytes) for r in file_rows if isinstance(r.size_bytes, int))

    summary: Dict[str, object] = {
        "generated_utc": utc_iso(datetime.now(tz=timezone.utc).timestamp()),
        "root": ROOT.name,
        "file_counts_by_extension": dict(by_ext.most_common()),
        "approx_total_size_bytes": total_size,
        "top_imports": imports.most_common(30),
        "likely_entry_points": entry_points,
        "tkinter_files": tkinter_users,
        "ignored_dirs": sorted(list(IGNORE_DIRS)),
        "ignored_suffixes": sorted(list(IGNORE_SUFFIXES)),
        "hash_limit_mb": MAX_HASH_MB,
    }
    Path("../summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")

# ---------------------- Main -------------------------------

def main() -> None:
    rows, file_rows = collect_tree()
    write_tree_md(rows)
    write_manifest_csv(file_rows)
    imports, entry_points, tkinter_users = analyze_python_sources(file_rows)
    write_summary_json(file_rows, imports, entry_points, tkinter_users)

    # Console summary
    print("\nCreated:")
    print("  - tree.md")
    print("  - manifest.csv")
    print("  - summary.json\n")

    by_ext: Counter = Counter(Path(r.rel_path).suffix.lower() for r in file_rows)
    print("Quick summary:")
    print("  Top extensions:", dict(by_ext.most_common(8)))
    print("  Likely entry points:", entry_points)
    print("  Tkinter files:", tkinter_users)
    print("  Top imports:", [m for m, _ in imports.most_common(10)])


if __name__ == "__main__":
    main()
