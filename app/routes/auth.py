"""
登录 / 登出：基于服务端 Session，不在响应中返回密码或哈希。
"""
from __future__ import annotations

import logging

from fastapi import APIRouter, Depends, Form, Request, status
from fastapi.responses import HTMLResponse, RedirectResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.services.auth_service import get_user_by_username, verify_password
from app.templating import templates

logger = logging.getLogger(__name__)

router = APIRouter(tags=["auth"])


@router.get("/login", response_class=HTMLResponse, response_model=None)
async def login_page(request: Request) -> HTMLResponse:
    if request.session.get("user_id"):
        return RedirectResponse("/dashboard", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(request, "login.html", {})


@router.post("/login", response_model=None)
async def login_submit(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    db: Session = Depends(get_db),
) -> RedirectResponse:
    user = get_user_by_username(db, username.strip())
    if not user or not verify_password(password, user.password_hash):
        logger.warning("登录失败: 用户名或密码错误")
        return templates.TemplateResponse(
            request,
            "login.html",
            {"error": "用户名或密码错误"},
            status_code=status.HTTP_200_OK,
        )
    request.session["user_id"] = user.id
    request.session["username"] = user.username
    logger.info("用户登录成功: %s", user.username)
    return RedirectResponse("/dashboard", status_code=status.HTTP_302_FOUND)


@router.post("/logout", response_model=None)
async def logout(request: Request) -> RedirectResponse:
    request.session.clear()
    return RedirectResponse("/login", status_code=status.HTTP_302_FOUND)
