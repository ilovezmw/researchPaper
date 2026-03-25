#!/usr/bin/env python3
"""
初始化默认管理员：用户名 admin / 密码 admin123（请在生产环境立即修改）。
需在项目根目录执行：python scripts/seed_admin.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from sqlalchemy import select

from app.database import SessionLocal, init_db
from app.models.user import User
from app.services.auth_service import hash_password


def main() -> None:
    init_db()
    db = SessionLocal()
    try:
        existing = db.execute(select(User).where(User.username == "admin")).scalar_one_or_none()
        if existing:
            print("用户 admin 已存在，跳过创建。")
            return
        u = User(username="admin", password_hash=hash_password("admin123"))
        db.add(u)
        db.commit()
        print("已创建默认用户 admin（请尽快修改密码）。")
    finally:
        db.close()


if __name__ == "__main__":
    main()
