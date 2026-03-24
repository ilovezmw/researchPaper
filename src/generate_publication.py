"""
Build a research note .docx from structured YAML (matches the two-column + AI disclosure layout).

Usage:
  python src/generate_publication.py --input "In Process/my-topic"
  python src/generate_publication.py --input "In Process/my-topic" --pdf
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import yaml
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Inches, Pt

# Allow running as script from repo root
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.docx_layout import (  # noqa: E402
    configure_page,
    set_section_columns,
    set_table_rows_cant_split,
    style_abstract_label,
    style_body,
    style_heading,
    style_meta_line,
    style_reference,
    style_table_caption,
    style_table_note,
    style_title,
)

DEFAULT_AI_DISCLOSURE = (
    "Portions of this research note were developed with the assistance of AI-based analytical and editorial tools to "
    "support research synthesis, language refinement, and structural organization. All analytical frameworks, "
    "interpretations, and investment perspectives presented in this report are independently developed and remain "
    "the sole responsibility of the author."
)


def _slug(s: str) -> str:
    s = re.sub(r"[^\w\s-]", "", s, flags=re.UNICODE)
    s = re.sub(r"[-\s]+", "_", s.strip()).strip("_")
    return s or "research_note"


def _add_paragraph(document: Document, text: str, style_fn) -> None:
    p = document.add_paragraph()
    r = p.add_run(text)
    style_fn(p)


def _add_heading(document: Document, text: str) -> None:
    p = document.add_paragraph()
    p.add_run(text)
    style_heading(p)


def _append_continuous_section(document: Document, num_cols: int) -> None:
    document.add_section(WD_SECTION.CONTINUOUS)
    s = document.sections[-1]
    configure_page(s)
    set_section_columns(s, num_cols)


def _apply_table_grid_fonts(table, font_pt: int = 8) -> None:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(font_pt)


def _add_data_table(
    document: Document,
    headers: list[str],
    rows: list[list[str]],
    col_width_in: float | None,
    *,
    font_pt: int = 8,
) -> None:
    ncols = len(headers)
    table = document.add_table(rows=1 + len(rows), cols=ncols)
    table.style = "Table Grid"
    set_table_rows_cant_split(table)
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = str(h)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            table.cell(ri + 1, ci).text = str(val)
    _apply_table_grid_fonts(table, font_pt=font_pt)
    if col_width_in is not None and ncols > 0:
        w = Inches(col_width_in)
        for row in table.rows:
            for cell in row.cells:
                cell.width = w


def _column_break_paragraph(document: Document) -> None:
    p = document.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run()
    r.add_break(WD_BREAK.COLUMN)


def _emit_table_blocks(
    document: Document,
    tables: list[dict],
    *,
    default_layout: str,
) -> None:
    """
    full_page: insert a single-column section so the table is not split across columns (table uses full text width).
    narrow: keep inside the current two-column section; use column break + cantSplit to reduce splits.
    """
    i = 0
    while i < len(tables):
        tbl = tables[i]
        layout = (tbl.get("layout") or default_layout or "full_page").strip().lower()
        if layout == "narrow":
            if tbl.get("column_break_before", True):
                _column_break_paragraph(document)
            cap = tbl.get("caption")
            if cap:
                p = document.add_paragraph()
                p.add_run(str(cap))
                style_table_caption(p)
                p.paragraph_format.keep_with_next = True
            headers = tbl.get("headers") or []
            rows = tbl.get("rows") or []
            _add_data_table(
                document,
                headers,
                rows,
                tbl.get("col_width_in"),
                font_pt=8,
            )
            note = tbl.get("note")
            if note:
                p = document.add_paragraph()
                p.add_run(str(note))
                style_table_note(p)
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            i += 1
            continue

        # Batch consecutive full-page tables into one single-column section
        group: list[dict] = [tbl]
        j = i + 1
        while j < len(tables):
            nxt = tables[j]
            nxt_layout = (nxt.get("layout") or default_layout or "full_page").strip().lower()
            if nxt_layout == "narrow":
                break
            group.append(nxt)
            j += 1

        _append_continuous_section(document, 1)
        for t in group:
            cap = t.get("caption")
            if cap:
                p = document.add_paragraph()
                p.add_run(str(cap))
                style_table_caption(p)
                p.paragraph_format.keep_with_next = True
            headers = t.get("headers") or []
            rows = t.get("rows") or []
            _add_data_table(
                document,
                headers,
                rows,
                t.get("col_width_in"),
                font_pt=8,
            )
            note = t.get("note")
            if note:
                p = document.add_paragraph()
                p.add_run(str(note))
                style_table_note(p)
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _append_continuous_section(document, 2)
        i = j


def build_document(data: dict) -> Document:
    document = Document()

    # --- Section 0: title block + abstract (single column) ---
    s0 = document.sections[0]
    configure_page(s0)
    set_section_columns(s0, 1)

    title = data.get("title", "Untitled")
    title_lines = title if isinstance(title, list) else [title]
    for line in title_lines:
        p = document.add_paragraph()
        p.add_run(str(line))
        style_title(p)

    meta = [
        ("author", True),
        ("affiliation", False),
        ("location", False),
        ("date", False),
    ]
    for key, bold in meta:
        val = data.get(key)
        if not val:
            continue
        p = document.add_paragraph()
        p.add_run(str(val))
        style_meta_line(p, bold=bold)

    abstract = (data.get("abstract") or "").strip()
    if abstract:
        p = document.add_paragraph()
        p.add_run("Abstract")
        style_abstract_label(p)
        for block in abstract.split("\n\n"):
            block = block.strip()
            if not block:
                continue
            p2 = document.add_paragraph()
            p2.add_run(block)
            style_body(p2)
            p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # --- Body: two columns; tables default to a single-column "island" to avoid splitting across columns ---
    _append_continuous_section(document, 2)
    default_table_layout = (data.get("default_table_layout") or "full_page").strip().lower()

    for sec in data.get("sections") or []:
        heading = sec.get("heading")
        if heading:
            _add_heading(document, str(heading))
        for para in sec.get("paragraphs") or []:
            p = document.add_paragraph()
            p.add_run(str(para))
            style_body(p)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        tables = sec.get("tables") or []
        if tables:
            _emit_table_blocks(document, tables, default_layout=default_table_layout)

    # --- Final section: AI disclosure + references (single column) ---
    document.add_section(WD_SECTION.CONTINUOUS)
    s_ai = document.sections[-1]
    configure_page(s_ai)
    set_section_columns(s_ai, 1)

    p = document.add_paragraph()
    p.add_run("AI Assistance Disclosure")
    style_heading(p)

    disclosure = (data.get("ai_disclosure") or DEFAULT_AI_DISCLOSURE).strip()
    p2 = document.add_paragraph()
    p2.add_run(disclosure)
    style_body(p2)
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    refs = data.get("references") or []
    if refs:
        pr = document.add_paragraph()
        pr.add_run("References")
        style_heading(pr)
        for ref in refs:
            rp = document.add_paragraph()
            rp.add_run(str(ref))
            style_reference(rp)

    return document


def load_yaml(path: Path) -> dict:
    text = path.read_text(encoding="utf-8")
    return yaml.safe_load(text) or {}


def export_pdf_word_win32(docx_path: Path, pdf_path: Path) -> None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError("PDF export on Windows requires pywin32. Install: pip install pywin32") from e

    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        # 17 = wdFormatPDF
        doc.SaveAs(str(pdf_path.resolve()), FileFormat=17)
        doc.Close(False)
        word.Quit()
    finally:
        pythoncom.CoUninitialize()


def mark_done(topic_dir: Path) -> None:
    marker = topic_dir / "Done.txt"
    marker.write_text(
        "Marked complete by generate_publication.py — you can remove this file if you prefer another convention.\n",
        encoding="utf-8",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate research note DOCX (and optional PDF) from manuscript.yaml")
    parser.add_argument("--input", "-i", required=True, help="Path to topic folder under In Process")
    parser.add_argument("--pdf", action="store_true", help="Also export PDF via Microsoft Word (Windows)")
    parser.add_argument("--no-done", action="store_true", help="Do not write Done.txt in the topic folder")
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    topic_dir = (root / args.input).resolve() if not Path(args.input).is_absolute() else Path(args.input).resolve()
    manuscript = topic_dir / "manuscript.yaml"
    if not manuscript.is_file():
        print(f"Missing manuscript.yaml in {topic_dir}", file=sys.stderr)
        return 2

    data = load_yaml(manuscript)
    doc = build_document(data)

    published = root / "published"
    published.mkdir(parents=True, exist_ok=True)

    title_slug = _slug(str(data.get("title") if not isinstance(data.get("title"), list) else data["title"][0]))
    base = f"{title_slug}_{data.get('date', '').replace(' ', '_')}"
    docx_name = f"{base}.docx"
    docx_path = published / docx_name

    doc.save(str(docx_path))
    print(f"Wrote {docx_path}")

    if args.pdf:
        pdf_path = published / f"{base}.pdf"
        export_pdf_word_win32(docx_path, pdf_path)
        print(f"Wrote {pdf_path}")

    if not args.no_done:
        mark_done(topic_dir)
        print(f"Marked done: {topic_dir / 'Done.txt'}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
