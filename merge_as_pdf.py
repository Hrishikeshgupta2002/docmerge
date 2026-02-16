#!/usr/bin/env python3
"""
Convert all DOCX files to PDF and merge into a single PDF.

Bypasses DOCX merge entirely - each file is converted individually by LibreOffice,
then PDFs are merged with pypdf. Preserves full formatting and images.

Requirements: LibreOffice (libreoffice --headless)

Run from docmerge directory:
  python merge_as_pdf.py
  python merge_as_pdf.py --input new --output merged_output.pdf

Or with venv:
  docmerge_env/bin/python merge_as_pdf.py --input new

API: POST /merge-pdf/ (same logic, accepts uploaded DOCX files)
"""
import argparse
import logging
import os
import shutil
import sys
import tempfile
from pathlib import Path

# Ensure we can import from parent
_script_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(_script_dir))
os.chdir(_script_dir)

from main import convert_docx_to_pdf, merge_pdfs

logger = logging.getLogger(__name__)


def merge_docx_to_pdf(docx_paths: list, output_path: str) -> None:
    """
    Convert DOCX files to PDF via LibreOffice and merge into one PDF.
    Shared logic for CLI and API. Preserves full formatting.
    """
    if len(docx_paths) < 2:
        raise ValueError("Need at least 2 DOCX files to merge")
    temp_dir = tempfile.mkdtemp(prefix="docmerge_pdf_")
    pdf_files = []
    try:
        for idx, docx_path in enumerate(docx_paths, start=1):
            logger.info(f"Converting {idx}/{len(docx_paths)}: {os.path.basename(docx_path)}")
            pdf_path = convert_docx_to_pdf(docx_path, temp_dir)
            pdf_files.append(pdf_path)
        merge_pdfs(pdf_files, output_path)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(
        description="Convert DOCX files to PDF and merge into one PDF"
    )
    parser.add_argument(
        "--input",
        "-i",
        default="new",
        help="Directory containing DOCX files (default: new)",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="merged_output.pdf",
        help="Output PDF path (default: merged_output.pdf)",
    )
    args = parser.parse_args()

    input_dir = Path(args.input)
    output_path = Path(args.output)

    if not input_dir.is_dir():
        print(f"Error: {input_dir} not found")
        sys.exit(1)

    docx_files = sorted(
        [str(p) for p in input_dir.glob("*.docx")],
        key=lambda x: os.path.basename(x),
    )

    if len(docx_files) < 2:
        print(f"Error: Need at least 2 DOCX files. Found {len(docx_files)} in {input_dir}")
        sys.exit(1)

    print(f"Found {len(docx_files)} DOCX files in {input_dir}")
    for f in docx_files[:5]:
        print(f"  - {os.path.basename(f)}")
    if len(docx_files) > 5:
        print(f"  ... and {len(docx_files) - 5} more")

    try:
        print("\nConverting DOCX to PDF (LibreOffice)...")
        for idx, docx_path in enumerate(docx_files, start=1):
            print(f"  ({idx}/{len(docx_files)}) {os.path.basename(docx_path)}")
        print(f"\nMerging PDFs...")
        merge_docx_to_pdf(docx_files, str(output_path))
        print(f"Saved: {output_path}")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
