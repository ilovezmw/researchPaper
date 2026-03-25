"""
仪表盘：格式化 DOCX、下载、发布到 GitHub。
所有文件路径均在项目 data/ 下，下载前校验归属用户。
"""
from __future__ import annotations

import logging
import uuid
from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, Request, UploadFile, status
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from sqlalchemy.orm import Session

from app.config import (
    FINAL_UPLOAD_DIR,
    FORMATTED_DIR,
    MAX_UPLOAD_BYTES,
    REFERENCE_SAMPLES_DIR,
    REFERENCE_STYLE_DOCX,
    UPLOAD_DIR,
)
from app.database import get_db
from app.models.file_history import FileHistory
from app.services.formatter import format_docx_to_path
from app.services.github_uploader import upload_docx_to_github
from app.services.style_profile import load_profile_from_reference
from app.templating import templates
from app.utils.file_validate import validate_docx_on_disk, validate_docx_upload
from app.utils.filename import safe_original_filename

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/dashboard", tags=["dashboard"])


def _require_user_id(request: Request) -> int:
    uid = request.session.get("user_id")
    if not uid:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="未登录")
    return int(uid)


@router.get("", response_class=HTMLResponse, response_model=None)
async def dashboard_page(request: Request) -> HTMLResponse | RedirectResponse:
    if not request.session.get("user_id"):
        return RedirectResponse("/login", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(request, "dashboard.html", {})


@router.post("/format")
async def format_docx(
    request: Request,
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
) -> JSONResponse:
    """上传原始 DOCX，生成格式化版本并登记 file_history。"""
    _require_user_id(request)
    user_id = int(request.session["user_id"])

    if not file.filename:
        return JSONResponse({"ok": False, "error": "未选择文件"}, status_code=400)

    safe_name = safe_original_filename(file.filename)
    content = await file.read()
    try:
        validate_docx_upload(file.filename, len(content))
    except ValueError as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=400)

    if len(content) > MAX_UPLOAD_BYTES:
        return JSONResponse({"ok": False, "error": "文件过大"}, status_code=400)

    uid = uuid.uuid4().hex
    local_in = UPLOAD_DIR / f"{uid}_{safe_name}"

    local_in.write_bytes(content)

    try:
        validate_docx_on_disk(local_in)
    except ValueError as e:
        local_in.unlink(missing_ok=True)
        return JSONResponse({"ok": False, "error": str(e)}, status_code=400)

    ref_path = REFERENCE_SAMPLES_DIR / REFERENCE_STYLE_DOCX
    profile = load_profile_from_reference(ref_path)

    row = FileHistory(
        user_id=user_id,
        original_filename=safe_name,
        status="processing",
    )
    db.add(row)
    db.commit()
    db.refresh(row)

    out_name = f"{row.id}_{uid}_formatted.docx"
    out_path = FORMATTED_DIR / out_name

    try:
        format_docx_to_path(local_in, out_path, profile)
        row.formatted_filename = out_name
        row.status = "formatted"
        row.error_message = None
    except Exception as e:
        logger.exception("格式化失败")
        row.status = "error_format"
        row.error_message = str(e)[:2000]
        db.commit()
        local_in.unlink(missing_ok=True)
        return JSONResponse({"ok": False, "error": "格式化失败，请检查文档内容是否有效"}, status_code=500)

    db.commit()

    return JSONResponse(
        {
            "ok": True,
            "file_id": row.id,
            "download_url": f"/dashboard/download/{row.id}",
            "message": "格式化完成，可下载",
        }
    )


@router.get("/download/{file_id}")
async def download_formatted(
    request: Request,
    file_id: int,
    db: Session = Depends(get_db),
) -> FileResponse:
    """仅允许下载当前用户名下记录的格式化文件。"""
    user_id = _require_user_id(request)
    row = db.get(FileHistory, file_id)
    if not row or row.user_id != user_id:
        raise HTTPException(status_code=404, detail="文件不存在")

    if not row.formatted_filename:
        raise HTTPException(status_code=404, detail="暂无可下载的格式化文件")

    name = Path(row.formatted_filename).name
    path = (FORMATTED_DIR / name).resolve()
    if not path.is_file() or path.parent != FORMATTED_DIR.resolve():
        raise HTTPException(status_code=404, detail="文件已丢失")

    return FileResponse(
        path,
        filename=f"formatted_{name}",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@router.post("/publish")
async def publish_final(
    request: Request,
    file: UploadFile = File(...),
    format_id: int | None = Form(None),
    db: Session = Depends(get_db),
) -> JSONResponse:
    """上传最终审阅版 DOCX 并推送到 GitHub published/。"""
    _require_user_id(request)
    user_id = int(request.session["user_id"])

    if not file.filename:
        return JSONResponse({"ok": False, "error": "未选择文件"}, status_code=400)

    safe_name = safe_original_filename(file.filename)
    content = await file.read()
    try:
        validate_docx_upload(file.filename, len(content))
    except ValueError as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=400)

    if len(content) > MAX_UPLOAD_BYTES:
        return JSONResponse({"ok": False, "error": "文件过大"}, status_code=400)

    uid = uuid.uuid4().hex
    local_path = FINAL_UPLOAD_DIR / f"{uid}_{safe_name}"
    local_path.write_bytes(content)

    try:
        validate_docx_on_disk(local_path)
    except ValueError as e:
        local_path.unlink(missing_ok=True)
        return JSONResponse({"ok": False, "error": str(e)}, status_code=400)

    row: FileHistory | None = None
    if format_id is not None:
        row = db.get(FileHistory, format_id)
        if not row or row.user_id != user_id:
            local_path.unlink(missing_ok=True)
            return JSONResponse({"ok": False, "error": "无效的 format_id"}, status_code=400)

    if row is None:
        row = FileHistory(
            user_id=user_id,
            original_filename=safe_name,
            status="processing",
        )
        db.add(row)
        db.commit()
        db.refresh(row)

    final_name = f"{row.id}_{uid}_final.docx"
    final_path = FINAL_UPLOAD_DIR / final_name
    local_path.rename(final_path)

    try:
        gh_path, html_url, raw_url = upload_docx_to_github(final_path, safe_name)
        row.final_uploaded_filename = final_name
        row.github_path = gh_path
        row.github_url = html_url
        row.github_raw_url = raw_url
        row.status = "published"
        row.error_message = None
        db.commit()
    except Exception as e:
        logger.exception("GitHub 发布失败")
        row.status = "error_publish"
        row.error_message = str(e)[:2000]
        db.commit()
        return JSONResponse(
            {"ok": False, "error": f"发布失败: {e!s}"},
            status_code=502,
        )

    return JSONResponse(
        {
            "ok": True,
            # 主链接：直链，便于复制后在浏览器直接下载或用 Word 打开
            "github_url": row.github_raw_url or row.github_url,
            "github_blob_url": row.github_url,
            "github_raw_url": row.github_raw_url,
            "github_path": row.github_path,
            "message": "已发布到 GitHub",
        }
    )
