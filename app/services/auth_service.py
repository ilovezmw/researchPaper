"""
登录校验：密码仅存储 bcrypt 哈希，不在日志或响应中泄露。
"""
from __future__ import annotations

import bcrypt
from sqlalchemy import select
from sqlalchemy.orm import Session

from app.models.user import User


def hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt(rounds=12)).decode("utf-8")


def verify_password(plain: str, password_hash: str) -> bool:
    try:
        return bcrypt.checkpw(plain.encode("utf-8"), password_hash.encode("utf-8"))
    except ValueError:
        return False


def get_user_by_username(db: Session, username: str) -> User | None:
    return db.execute(select(User).where(User.username == username)).scalar_one_or_none()
