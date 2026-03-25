"""
使用 PyGithub 将最终 DOCX 提交到指定仓库路径。
GITHUB_TOKEN 仅从环境变量读取，禁止记录到日志。
"""
from __future__ import annotations

import logging
import re
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

from github import Github
from github.GithubException import GithubException

from app.config import GITHUB_PUBLISH_PREFIX, GITHUB_REPO, GITHUB_TOKEN

logger = logging.getLogger(__name__)


def _build_raw_content_url(repo_full: str, branch: str, remote_path: str) -> str:
    """
    生成可直接访问文件内容的直链（浏览器下载 / 外部应用打开）。
    blob 页面 https://github.com/.../blob/... 适合浏览仓库，不适合直接当文件 URL。
    """
    parts = repo_full.strip().split("/")
    if len(parts) != 2:
        return ""
    owner, repo = parts[0], parts[1]
    path_segments = remote_path.strip("/").split("/")
    encoded = "/".join(quote(seg, safe="") for seg in path_segments)
    return f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{encoded}"


def _safe_github_filename(original: str) -> str:
    """仅保留安全字符，避免路径注入。"""
    base = Path(original).name
    base = re.sub(r"[^\w.\-]", "_", base, flags=re.UNICODE)
    if not base.lower().endswith(".docx"):
        base = f"{base}.docx"
    return base[:180]


def upload_docx_to_github(local_path: Path, original_filename: str) -> tuple[str, str, str]:
    """
    上传文件到 published/YYYYMMDD_HHMMSS_<name>.docx。
    返回 (github_path, html_url, raw_url)。
    失败时抛出 GithubException 或 ValueError。
    """
    if not GITHUB_TOKEN:
        raise ValueError("未配置 GITHUB_TOKEN，无法发布到 GitHub")

    safe_name = _safe_github_filename(original_filename)
    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    remote_path = f"{GITHUB_PUBLISH_PREFIX.rstrip('/')}/{ts}_{safe_name}"

    g = Github(GITHUB_TOKEN)
    repo = g.get_repo(GITHUB_REPO)
    branch = repo.default_branch
    content_bytes = local_path.read_bytes()
    message = f"Publish paper: {ts}_{safe_name}"

    try:
        try:
            existing = repo.get_contents(remote_path, ref=branch)
            if isinstance(existing, list):
                raise GithubException(500, {}, "Unexpected directory content", None)
            res = repo.update_file(
                remote_path,
                message,
                content_bytes,
                existing.sha,
                branch=branch,
            )
            content_file = res["content"]
        except GithubException as e:
            if e.status == 404:
                res = repo.create_file(
                    remote_path,
                    message,
                    content_bytes,
                    branch=branch,
                )
                content_file = res["content"]
            else:
                raise
    except GithubException:
        logger.exception("GitHub API 错误（未记录 token）")
        raise

    html_url = content_file.html_url
    # API 的 download_url 有时为空或非 raw 域名；统一用 raw.githubusercontent.com 直链
    raw_url = _build_raw_content_url(GITHUB_REPO, branch, remote_path)
    if not raw_url:
        raw_url = getattr(content_file, "download_url", None) or html_url

    logger.info("已发布到 GitHub: %s", remote_path)
    return remote_path, html_url, raw_url
