"""
Layout helpers: single-column / two-column sections for research notes.
Uses Times New Roman, US Letter, and OOXML tweaks for consistent academic-style output.
"""
from __future__ import annotations

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


def apply_document_defaults(document: Document) -> None:
    """Base font and spacing so ad-hoc paragraphs inherit something sensible."""
    normal = document.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(11)
    n_pf = normal.paragraph_format
    n_pf.line_spacing = 1.15
    n_pf.space_after = Pt(0)
    n_pf.space_before = Pt(0)


def set_table_rows_cant_split(table) -> None:
    """Reduce row splits; helps with awkward breaks (esp. pagination)."""
    for row in table.rows:
        tr = row._tr
        tr_pr = tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))


def polish_research_table(table, *, header_fill: str = "D9D9D9", cell_margin_twips: int = 100) -> None:
    """
    Full-width table (pct), header row repeat, grey header band, padding, header text bold.
    """
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    for tbl_w in tbl_pr.findall(qn("w:tblW")):
        tbl_pr.remove(tbl_w)
    tbl_w = OxmlElement("w:tblW")
    tbl_w.set(qn("w:type"), "pct")
    tbl_w.set(qn("w:w"), "5000")  # 100% of container (Word pct = 50ths of a percent)
    tbl_pr.append(tbl_w)

    if not table.rows:
        return

    tr0 = table.rows[0]._tr
    tr_pr0 = tr0.get_or_add_trPr()
    if tr_pr0.find(qn("w:tblHeader")) is None:
        tr_pr0.append(OxmlElement("w:tblHeader"))

    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            if ri == 0:
                for shd in tc_pr.findall(qn("w:shd")):
                    tc_pr.remove(shd)
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), header_fill)
                tc_pr.append(shd)

            if tc_pr.find(qn("w:tcMar")) is None:
                tc_mar = OxmlElement("w:tcMar")
                for edge in ("top", "left", "bottom", "right"):
                    m = OxmlElement(f"w:{edge}")
                    m.set(qn("w:w"), str(cell_margin_twips))
                    m.set(qn("w:type"), "dxa")
                    tc_mar.append(m)
                tc_pr.append(tc_mar)

            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                if ri == 0:
                    for r in p.runs:
                        r.bold = True
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _ensure_cols(sect_pr: OxmlElement) -> OxmlElement:
    cols = sect_pr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sect_pr.append(cols)
    return cols


def set_section_columns(section, num: int, *, gap_twips: int | None = None) -> None:
    """
    gap_twips: space between columns in dxa/twips (not points). ~480 ≈ 0.33\" gutter.
    """
    cols = _ensure_cols(section._sectPr)
    cols.set(qn("w:num"), str(num))
    if num > 1:
        g = 480 if gap_twips is None else gap_twips
        cols.set(qn("w:space"), str(g))
        cols.set(qn("w:equalWidth"), "1")
    else:
        cols.set(qn("w:space"), "0")
        ew = qn("w:equalWidth")
        if ew in cols.attrib:
            del cols.attrib[ew]


def configure_page(section, margin_in: float = 1.0) -> None:
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(margin_in)
    section.bottom_margin = Inches(margin_in)
    section.left_margin = Inches(margin_in)
    section.right_margin = Inches(margin_in)
    section.orientation = WD_ORIENT.PORTRAIT


def style_title(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(10)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.widow_control = True
    for run in paragraph.runs:
        run.bold = True
        run.font.size = Pt(15)
        run.font.name = "Times New Roman"


def style_title_subline(paragraph) -> None:
    """Slightly smaller second line when title is split across runs/lines."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(8)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.widow_control = True
    for run in paragraph.runs:
        run.bold = True
        run.font.size = Pt(13)
        run.font.name = "Times New Roman"


def style_meta_line(paragraph, bold: bool = False, size: int = 11) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(2)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    for run in paragraph.runs:
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = "Times New Roman"


def style_heading(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(12)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.08
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.widow_control = True
    for run in paragraph.runs:
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = "Times New Roman"


def style_body(paragraph) -> None:
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.widow_control = True
    for run in paragraph.runs:
        run.font.size = Pt(11)
        run.font.name = "Times New Roman"


def style_abstract_body(paragraph) -> None:
    """Abstract text often one size smaller than body."""
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.widow_control = True
    for run in paragraph.runs:
        run.font.size = Pt(10)
        run.font.name = "Times New Roman"


def style_abstract_label(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(14)
    paragraph.paragraph_format.space_after = Pt(8)
    paragraph.paragraph_format.line_spacing = 1.0
    for run in paragraph.runs:
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = "Times New Roman"


def style_table_caption(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(10)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.08
    paragraph.paragraph_format.keep_with_next = True
    for run in paragraph.runs:
        run.bold = True
        run.italic = True
        run.font.size = Pt(9)
        run.font.name = "Times New Roman"


def style_table_note(paragraph) -> None:
    paragraph.paragraph_format.space_after = Pt(10)
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.line_spacing = 1.08
    for run in paragraph.runs:
        run.font.size = Pt(9)
        run.font.name = "Times New Roman"


def style_reference(paragraph) -> None:
    """Hanging indent: first line outdented, body indented (common for numbered refs)."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = paragraph.paragraph_format
    hang = Inches(0.32)
    pf.left_indent = hang
    pf.first_line_indent = -hang
    pf.space_after = Pt(4)
    pf.line_spacing = 1.08
    pf.widow_control = True
    for run in paragraph.runs:
        run.font.size = Pt(10)
        run.font.name = "Times New Roman"


def style_section_divider_heading(paragraph) -> None:
    """AI disclosure / references block titles in final single-column section."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(18)
    paragraph.paragraph_format.space_after = Pt(8)
    paragraph.paragraph_format.line_spacing = 1.08
    paragraph.paragraph_format.keep_with_next = True
    for run in paragraph.runs:
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = "Times New Roman"
