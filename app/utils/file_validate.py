"""
上传 DOCX 校验：扩展名、大小、ZIP/OOXML 结构。
"""
from __future__ import annotations

import logging
import zipfile
from pathlib import Path

from app.config import MAX_UPLOAD_BYTES

logger = logging.getLogger(__name__)


def validate_docx_upload(filename: str, size: int) -> None:
    """拒绝非 docx 或超大文件。"""
    if not filename.lower().endswith(".docx"):
        raise ValueError("仅支持 .docx 文件")
    if size > MAX_UPLOAD_BYTES:
        raise ValueError(f"文件过大，最大允许 {MAX_UPLOAD_BYTES // (1024 * 1024)} MB")


def validate_docx_on_disk(path: Path) -> None:
    """确认磁盘上的文件为有效 DOCX（ZIP 且含 word/ 部件）。"""
    if not path.is_file():
        raise ValueError("文件不存在")
    if not zipfile.is_zipfile(path):
        raise ValueError("不是有效的 DOCX（应为 ZIP 容器）")
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        if not any(n.startswith("word/") for n in names):
            raise ValueError("DOCX 结构异常：缺少 word/ 部件")
    logger.debug("DOCX 校验通过: %s", path.name)
