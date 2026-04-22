# pdf_to_md_chunks.py
"""
Purpose:
Extract text from a PDF and write it to Markdown and JSON files with visible progress reporting.

What this file does:
- Opens a PDF from disk
- Reports startup details and major phase transitions
- Extracts text page by page with periodic progress updates
- Writes one combined Markdown file
- Writes one page-by-page JSON file
- Splits the Markdown into smaller upload-friendly chunk files

Place in the larger scheme:
Useful for converting large evidence/log PDFs into GPT-friendly text artifacts while
showing enough runtime status to tell whether the job is progressing or stalled.

Why that matters:
Large PDFs can appear hung for long stretches. This script prints progress, rate, and ETA
so you can see where time is being spent.
"""

from __future__ import annotations

import json
import re
import time
from datetime import timedelta
from pathlib import Path
from typing import List, TypedDict

try:
    from pypdf import PdfReader
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Missing dependency: pypdf\n"
        "Install it with:\n"
        "python -m pip install pypdf"
    ) from exc


class ExtractedPage(TypedDict):
    page_number: int
    text: str


SOURCE_PDF = Path(
    r"C:\Users\e.muraj\Desktop\Production\3_TrackWise QMS Rebuild\NEUROTIC_QMS_RUN_20251215_183029_EVENT_LOG.pdf"
)
DEST_FOLDER = Path(r"C:\Users\e.muraj\Desktop\DEV\NEUROTIC_EVIDENCE")

# Change to 500 if you want only the first 500 pages processed.
PAGE_LIMIT: int | None = None

# Chunk files will target this max size.
MAX_CHUNK_BYTES = 2_000_000

# Print a progress line every N pages.
PROGRESS_EVERY_PAGES = 25

# Also print an explicit "Starting page X..." line every N pages.
STARTING_PAGE_PRINT_EVERY = 25


def log(message: str) -> None:
    """Print immediately so PyCharm terminal shows live progress."""
    print(message, flush=True)


def format_seconds(seconds: float) -> str:
    """Format elapsed or remaining seconds as H:MM:SS."""
    return str(timedelta(seconds=max(0, int(seconds))))


def format_bytes(num_bytes: int) -> str:
    """Format bytes in a readable MB string."""
    return f"{num_bytes / (1024 * 1024):.1f} MB"


def normalize_text(text: str) -> str:
    """
    Clean extracted text to reduce noise and keep output size reasonable.
    """
    if not text:
        return ""

    text = text.replace("\x00", "")
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def extract_pages(pdf_path: Path, page_limit: int | None = None) -> List[ExtractedPage]:
    """
    Extract text from each page with detailed progress reporting.
    """
    source_size = pdf_path.stat().st_size

    log("Opening PDF file...")
    open_start = time.time()
    reader = PdfReader(str(pdf_path))
    log(f"PDF opened in {format_seconds(time.time() - open_start)}")

    log("Counting pages...")
    count_start = time.time()
    total_pages = len(reader.pages)
    log(f"Total pages detected: {total_pages}")
    log(f"Page counting took: {format_seconds(time.time() - count_start)}")

    pages_to_process = min(page_limit, total_pages) if page_limit is not None else total_pages

    log("")
    log("Startup details")
    log(f"Source: {pdf_path}")
    log(f"File size: {format_bytes(source_size)}")
    log(f"Page limit: {page_limit if page_limit is not None else 'all pages'}")
    log(f"Pages to process: {pages_to_process}")
    log(f"Output folder: {DEST_FOLDER}")
    log(f"Chunk target size: {MAX_CHUNK_BYTES:,} bytes")
    log(f"Progress interval: every {PROGRESS_EVERY_PAGES} pages")
    log("Starting extraction...")

    extraction_start = time.time()
    extracted: List[ExtractedPage] = []
    empty_pages = 0
    total_chars = 0

    for page_index in range(pages_to_process):
        page_number = page_index + 1

        if page_number == 1 or page_number % STARTING_PAGE_PRINT_EVERY == 0:
            log(f"[Extraction] Starting page {page_number}...")

        page_start = time.time()
        raw_text = reader.pages[page_index].extract_text() or ""
        cleaned_text = normalize_text(raw_text)
        page_seconds = time.time() - page_start

        if not cleaned_text:
            empty_pages += 1

        total_chars += len(cleaned_text)

        extracted.append(
            {
                "page_number": page_number,
                "text": cleaned_text,
            }
        )

        should_print_progress = (
            page_number == 1
            or page_number == pages_to_process
            or page_number % PROGRESS_EVERY_PAGES == 0
        )

        if should_print_progress:
            elapsed = time.time() - extraction_start
            rate = page_number / elapsed if elapsed > 0 else 0.0
            remaining_pages = pages_to_process - page_number
            eta_seconds = remaining_pages / rate if rate > 0 else 0.0
            percent = (page_number / pages_to_process) * 100 if pages_to_process > 0 else 100.0

            text_status = "empty" if not cleaned_text else f"{len(cleaned_text):,} chars"

            log(
                "[Extraction] "
                f"Page {page_number}/{pages_to_process} ({percent:.1f}%) | "
                f"Elapsed: {format_seconds(elapsed)} | "
                f"Rate: {rate:.2f} pages/sec | "
                f"ETA: {format_seconds(eta_seconds)} | "
                f"This page: {text_status} | "
                f"Page time: {page_seconds:.2f}s"
            )

    total_elapsed = time.time() - extraction_start
    avg_chars = (total_chars / pages_to_process) if pages_to_process > 0 else 0

    log("")
    log("Extraction complete.")
    log(f"Pages extracted: {pages_to_process}")
    log(f"Empty pages: {empty_pages}")
    log(f"Total characters: {total_chars:,}")
    log(f"Average chars/page: {avg_chars:,.1f}")
    log(f"Extraction runtime: {format_seconds(total_elapsed)}")

    return extracted


def build_full_markdown(pages: List[ExtractedPage], source_name: str) -> str:
    """
    Build one Markdown document containing all extracted pages.
    """
    log("")
    log("Building combined markdown...")
    build_start = time.time()

    parts: List[str] = [f"# Extracted text from {source_name}\n"]

    for page in pages:
        page_number = page["page_number"]
        text = page["text"].strip()
        if not text:
            text = "[No extractable text on this page]"

        parts.append(f"## Page {page_number}\n\n{text}\n")

    full_md = "\n".join(parts).strip() + "\n"

    elapsed = time.time() - build_start
    log(f"Combined markdown size: {len(full_md.encode('utf-8')):,} bytes")
    log(f"Markdown build runtime: {format_seconds(elapsed)}")

    return full_md


def split_markdown_by_size(full_md: str, max_chunk_bytes: int) -> List[str]:
    """
    Split the markdown into upload-friendly chunks, preferring page boundaries.
    """
    log("")
    log("Starting chunking...")
    chunk_start = time.time()

    sections: List[str] = re.split(r"(?=^## Page \d+\s*$)", full_md, flags=re.MULTILINE)

    if not sections:
        log("Chunking fallback: no sections found; returning single chunk.")
        return [full_md]

    chunks: List[str] = []
    current: str = ""

    for section in sections:
        if not section.strip():
            continue

        candidate = section if not current else current + "\n" + section

        if len(candidate.encode("utf-8")) <= max_chunk_bytes:
            current = candidate
            continue

        if current:
            chunks.append(current.strip() + "\n")
            log(
                f"[Chunking] Finalized chunk {len(chunks)} | "
                f"Size: {len(chunks[-1].encode('utf-8')):,} bytes"
            )
            current = section
            continue

        paragraphs = section.split("\n\n")
        temp: str = ""

        for para in paragraphs:
            candidate_para = para if not temp else temp + "\n\n" + para
            if len(candidate_para.encode("utf-8")) <= max_chunk_bytes:
                temp = candidate_para
            else:
                if temp:
                    chunks.append(temp.strip() + "\n")
                    log(
                        f"[Chunking] Finalized chunk {len(chunks)} | "
                        f"Size: {len(chunks[-1].encode('utf-8')):,} bytes"
                    )
                temp = para

        current = temp

    if current:
        chunks.append(current.strip() + "\n")
        log(
            f"[Chunking] Finalized chunk {len(chunks)} | "
            f"Size: {len(chunks[-1].encode('utf-8')):,} bytes"
        )

    elapsed = time.time() - chunk_start
    log(f"Chunk count: {len(chunks)}")
    log(f"Chunking runtime: {format_seconds(elapsed)}")

    return chunks


def write_json_pages(pages: List[ExtractedPage], out_path: Path) -> None:
    """
    Write page-by-page JSON output.
    """
    log("")
    log("Writing JSON...")

    write_start = time.time()
    payload = {
        "source_pdf": str(SOURCE_PDF),
        "page_count": len(pages),
        "pages": pages,
    }

    out_path.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    elapsed = time.time() - write_start
    log(f"Done: {out_path}")
    log(f"JSON size: {out_path.stat().st_size:,} bytes")
    log(f"JSON write runtime: {format_seconds(elapsed)}")


def write_full_markdown(full_md: str, out_path: Path) -> None:
    """
    Write the combined Markdown file.
    """
    log("")
    log("Writing full markdown...")

    write_start = time.time()
    out_path.write_text(full_md, encoding="utf-8")
    elapsed = time.time() - write_start

    log(f"Done: {out_path}")
    log(f"Markdown size: {out_path.stat().st_size:,} bytes")
    log(f"Markdown write runtime: {format_seconds(elapsed)}")


def write_chunk_files(chunks: List[str], chunk_dir: Path, base_name: str) -> None:
    """
    Write each markdown chunk to its own file.
    """
    log("")
    log("Writing chunk files...")

    chunk_dir.mkdir(parents=True, exist_ok=True)
    write_start = time.time()

    largest_size = 0

    for idx, chunk in enumerate(chunks, start=1):
        chunk_path = chunk_dir / f"{base_name}_part_{idx:03d}.md"
        chunk_path.write_text(chunk, encoding="utf-8")

        chunk_size = chunk_path.stat().st_size
        largest_size = max(largest_size, chunk_size)

        log(
            f"[Write] Chunk {idx}/{len(chunks)} saved: {chunk_path.name} | "
            f"Size: {chunk_size:,} bytes"
        )

    elapsed = time.time() - write_start
    log(f"Chunk folder: {chunk_dir}")
    log(f"Largest chunk: {largest_size:,} bytes")
    log(f"Chunk file write runtime: {format_seconds(elapsed)}")


def main() -> None:
    """
    Main program flow.
    """
    overall_start = time.time()

    if not SOURCE_PDF.exists():
        raise FileNotFoundError(f"Source PDF not found: {SOURCE_PDF}")

    DEST_FOLDER.mkdir(parents=True, exist_ok=True)

    base_name = SOURCE_PDF.stem

    pages = extract_pages(SOURCE_PDF, PAGE_LIMIT)

    json_path = DEST_FOLDER / f"{base_name}_extracted_pages.json"
    full_md_path = DEST_FOLDER / f"{base_name}_extracted_full.md"
    chunk_dir = DEST_FOLDER / f"{base_name}_md_chunks"

    write_json_pages(pages, json_path)

    full_md = build_full_markdown(pages, SOURCE_PDF.name)
    write_full_markdown(full_md, full_md_path)

    chunks = split_markdown_by_size(full_md, MAX_CHUNK_BYTES)
    write_chunk_files(chunks, chunk_dir, base_name)

    total_runtime = time.time() - overall_start

    log("")
    log("Done.")
    log(f"Total runtime: {format_seconds(total_runtime)}")
    log(f"JSON: {json_path}")
    log(f"Full Markdown: {full_md_path}")
    log(f"Chunk folder: {chunk_dir}")
    log(f"Chunk count: {len(chunks)}")


if __name__ == "__main__":
    main()