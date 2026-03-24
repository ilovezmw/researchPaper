"""
Export an existing .docx to PDF using Microsoft Word (Windows + pywin32).

Usage:
  python src/word_to_pdf.py "In Process/my-topic/draft.docx"
  python src/word_to_pdf.py "In Process/my-topic/draft.docx" -o Published/my-topic.pdf
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from src.generate_publication import export_pdf_word_win32  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert .docx to .pdf via Word")
    parser.add_argument("docx", type=Path, help="Path to the Word document")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output PDF path (default: same folder/name as .docx with .pdf)",
    )
    args = parser.parse_args()

    docx = args.docx.resolve() if not args.docx.is_absolute() else args.docx
    if not docx.is_file():
        print(f"File not found: {docx}", file=sys.stderr)
        return 2
    if docx.suffix.lower() != ".docx":
        print("Expected a .docx file.", file=sys.stderr)
        return 2

    out = args.output
    if out is None:
        pdf_path = docx.with_suffix(".pdf")
    else:
        pdf_path = out.resolve() if not out.is_absolute() else out
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

    export_pdf_word_win32(docx, pdf_path)
    print(f"Wrote {pdf_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
