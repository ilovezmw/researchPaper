"""
SQLite + SQLAlchemy：单文件数据库位于项目 data/ 下。
"""
from __future__ import annotations

from collections.abc import Generator

from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

from app.config import DATABASE_PATH

# check_same_thread=False 允许多线程下由 FastAPI 使用同一引擎（会话仍按请求创建）
DATABASE_URL = f"sqlite:///{DATABASE_PATH.as_posix()}"
engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False},
    echo=False,
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


class Base(DeclarativeBase):
    pass


def get_db() -> Generator[Session, None, None]:
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db() -> None:
    """创建表结构（启动时调用）。"""
    from app.models import FileHistory, User  # noqa: F401

    Base.metadata.create_all(bind=engine)
