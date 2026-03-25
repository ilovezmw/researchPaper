"""
FastAPI 入口：挂载 Session、静态资源、路由；启动时初始化 SQLite。
监听地址与端口由环境变量 HOST/PORT 控制，避免与服务器上其他站点冲突。
"""
from __future__ import annotations

import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware

from app.config import BASE_DIR, LOG_DIR, SECRET_KEY
from app.database import init_db
from app.routes import auth as auth_routes
from app.routes import dashboard as dashboard_routes

logger = logging.getLogger(__name__)


def _setup_logging() -> None:
    log_path = LOG_DIR / "app.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


@asynccontextmanager
async def lifespan(app: FastAPI):
    _setup_logging()
    init_db()
    logger.info("数据库已初始化")
    yield


app = FastAPI(
    title="Research Paper Portal",
    lifespan=lifespan,
)

# Session：cookie 内签名，服务端不存 session 文件
app.add_middleware(
    SessionMiddleware,
    secret_key=SECRET_KEY,
    same_site="lax",
    https_only=False,
)

app.mount(
    "/static",
    StaticFiles(directory=str(BASE_DIR / "app" / "static")),
    name="static",
)

app.include_router(auth_routes.router)
app.include_router(dashboard_routes.router)


@app.get("/", response_model=None)
async def root(request: Request) -> RedirectResponse:
    if request.session.get("user_id"):
        return RedirectResponse("/dashboard")
    return RedirectResponse("/login")


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}
