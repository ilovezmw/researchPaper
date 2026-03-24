"""
Layout helpers: single-column / two-column sections for research notes.
"""
from __future__ import annotations

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


def set_table_rows_cant_split(table) -> None:
    """Reduce row splits; helps with awkward breaks (esp. pagination)."""
    for row in table.rows:
        tr = row._tr
        tr_pr = tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))


def _ensure_cols(sect_pr: OxmlElement) -> OxmlElement:
    cols = sect_pr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sect_pr.append(cols)
    return cols


def set_section_columns(section, num: int, space_pt: int = 12) -> None:
    cols = _ensure_cols(section._sectPr)
    cols.set(qn("w:num"), str(num))
    if num > 1:
        cols.set(qn("w:space"), str(space_pt))


def configure_page(section, margin_in: float = 1.0) -> None:
    section.top_margin = Inches(margin_in)
    section.bottom_margin = Inches(margin_in)
    section.left_margin = Inches(margin_in)
    section.right_margin = Inches(margin_in)
    section.orientation = WD_ORIENT.PORTRAIT


def style_title(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(6)
    for run in paragraph.runs:
        run.bold = True
        run.font.size = Pt(16)
        run.font.name = "Times New Roman"


def style_meta_line(paragraph, bold: bool = False, size: int = 11) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    for run in paragraph.runs:
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = "Times New Roman"


def style_heading(paragraph) -> None:
    paragraph.paragraph_format.space_before = Pt(10)
    paragraph.paragraph_format.space_after = Pt(4)
    for run in paragraph.runs:
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = "Times New Roman"


def style_body(paragraph) -> None:
    paragraph.paragraph_format.space_after = Pt(4)
    paragraph.paragraph_format.line_spacing = 1.15
    for run in paragraph.runs:
        run.font.size = Pt(10)
        run.font.name = "Times New Roman"


def style_abstract_label(paragraph) -> None:
    paragraph.paragraph_format.space_before = Pt(12)
    paragraph.paragraph_format.space_after = Pt(6)
    for run in paragraph.runs:
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = "Times New Roman"


def style_table_caption(paragraph) -> None:
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(4)
    for run in paragraph.runs:
        run.bold = True
        run.italic = True
        run.font.size = Pt(9)
        run.font.name = "Times New Roman"


def style_table_note(paragraph) -> None:
    paragraph.paragraph_format.space_after = Pt(6)
    for run in paragraph.runs:
        run.font.size = Pt(9)
        run.font.name = "Times New Roman"


def style_reference(paragraph) -> None:
    paragraph.paragraph_format.space_after = Pt(2)
    paragraph.paragraph_format.line_spacing = 1.0
    for run in paragraph.runs:
        run.font.size = Pt(9)
        run.font.name = "Times New Roman"
