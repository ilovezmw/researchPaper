"""
应用配置：从环境变量加载，所有路径限制在项目根目录下。
部署时通过 .env 设置 HOST/PORT，避免与现有站点冲突。
"""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

# 项目根目录 = app/ 的上一级（包含 data/、.env）
BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")


def _env_str(key: str, default: str) -> str:
    v = os.getenv(key)
    return v if v is not None and v != "" else default


def _env_int(key: str, default: int) -> int:
    raw = os.getenv(key)
    if raw is None or raw == "":
        return default
    try:
        return int(raw)
    except ValueError:
        return default


# --- 网络（默认仅本机 + 高端口，降低误暴露风险；需要外网访问时设 HOST=0.0.0.0）---
HOST = _env_str("HOST", "0.0.0.0")
PORT = _env_int("PORT", 8765)

# Session 签名密钥（生产环境必须替换为随机长字符串）
SECRET_KEY = _env_str("SECRET_KEY", "change-me-in-production-use-openssl-rand-hex-32")

# SQLite 数据库文件始终放在项目 data 目录内
DATABASE_PATH = BASE_DIR / _env_str("DATABASE_PATH", "data/app.db")

# 上传与生成文件目录（均在项目内，不指向系统其他路径）
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
FORMATTED_DIR = DATA_DIR / "formatted"
FINAL_UPLOAD_DIR = DATA_DIR / "final_uploads"
LOG_DIR = DATA_DIR / "logs"
REFERENCE_SAMPLES_DIR = DATA_DIR / "reference_samples"

# 单文件大小上限（字节），默认 15MB
MAX_UPLOAD_BYTES = _env_int("MAX_UPLOAD_BYTES", 15 * 1024 * 1024)

# 参考样式 DOCX 文件名（放在 reference_samples/ 下）
REFERENCE_STYLE_DOCX = _env_str(
    "REFERENCE_STYLE_DOCX",
    "Research Paper Generation Request.docx",
)

# GitHub 发布（令牌仅服务端使用，禁止写入日志或模板）
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN") or ""
GITHUB_REPO = _env_str("GITHUB_REPO", "ilovezmw/researchPaper")
GITHUB_PUBLISH_PREFIX = _env_str("GITHUB_PUBLISH_PREFIX", "published/")

# 格式化时注入作者行（源稿标题下无姓名/单位时仍显示；不参与 skip 计数）
# 含空格请用双引号，例如 AUTHOR_LINE_2="Independent Researcher, Singapore"
AUTHOR_LINE_1 = _env_str("AUTHOR_LINE_1", "")
AUTHOR_LINE_2 = _env_str("AUTHOR_LINE_2", "")
AUTHOR_LINE_3 = _env_str("AUTHOR_LINE_3", "")

# 确保目录存在（启动时创建，不影响其他项目）
for _d in (
    DATA_DIR,
    UPLOAD_DIR,
    FORMATTED_DIR,
    FINAL_UPLOAD_DIR,
    LOG_DIR,
    REFERENCE_SAMPLES_DIR,
):
    _d.mkdir(parents=True, exist_ok=True)
