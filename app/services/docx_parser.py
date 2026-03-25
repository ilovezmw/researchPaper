"""
从 DOCX 中按文档顺序解析段落与表格，供格式化与章节推断使用。
表格以 OOXML 元素形式保留，便于复制到新文档。
"""
from __future__ import annotations

from copy import deepcopy
from typing import Literal

from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_body_blocks(doc: DocumentType) -> list[tuple[Literal["p", "t"], object]]:
    """
    按 body 子元素顺序返回列表项：
    ("p", Paragraph) 或 ("t", Table)
    """
    out: list[tuple[Literal["p", "t"], object]] = []
    for child in doc.element.body:
        if child.tag == qn("w:p"):
            out.append(("p", Paragraph(child, doc)))
        elif child.tag == qn("w:tbl"):
            out.append(("t", Table(child, doc)))
    return out


def paragraph_text(p: Paragraph) -> str:
    """合并段落内全部 run 的文本。"""
    return (p.text or "").strip()


def append_table_copy(target_doc: DocumentType, table: Table) -> None:
    """将表格 OOXML 深拷贝到目标文档末尾（保留合并单元格等）。"""
    target_doc.element.body.append(deepcopy(table._tbl))


def blocks_to_plain_lines(
    blocks: list[tuple[Literal["p", "t"], object]],
) -> list[str]:
    """仅段落块转为文本行列表，用于章节关键词匹配。"""
    lines: list[str] = []
    for kind, obj in blocks:
        if kind != "p":
            continue
        t = paragraph_text(obj)  # type: ignore[arg-type]
        if t:
            lines.append(t)
    return lines
