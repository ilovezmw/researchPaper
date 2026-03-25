"""集中挂载 Jinja2，避免 routes 与 main 循环依赖。"""
from __future__ import annotations

from fastapi.templating import Jinja2Templates

from app.config import BASE_DIR

templates = Jinja2Templates(directory=str(BASE_DIR / "app" / "templates"))
