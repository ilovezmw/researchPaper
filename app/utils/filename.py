"""
安全文件名：防止路径穿越与奇怪字符。
"""
from __future__ import annotations

import re
from pathlib import Path


def safe_original_filename(name: str, suffix: str = ".docx") -> str:
    """保留 basename，替换危险字符，限制长度。"""
    base = Path(name).name
    base = re.sub(r"[^\w.\-\s\u4e00-\u9fff]", "_", base, flags=re.UNICODE)
    base = base.strip() or "document"
    if not base.lower().endswith(suffix.lower()):
        base = f"{base}{suffix}"
    return base[:200]


def safe_storage_basename(prefix: str, original: str) -> str:
    """生成存储用文件名：prefix_uuid_safe.docx 由调用方拼 UUID。"""
    return f"{prefix}_{safe_original_filename(original)}"
