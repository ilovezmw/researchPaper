from __future__ import annotations

from datetime import datetime

from sqlalchemy import DateTime, ForeignKey, Integer, String, Text, func
from sqlalchemy.orm import Mapped, mapped_column

from app.database import Base


class FileHistory(Base):
    """格式化与 GitHub 发布记录，用于下载鉴权与审计。"""

    __tablename__ = "file_history"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_id: Mapped[int] = mapped_column(Integer, ForeignKey("users.id"), nullable=False, index=True)

    original_filename: Mapped[str] = mapped_column(String(512), nullable=False)
    formatted_filename: Mapped[str | None] = mapped_column(String(512), nullable=True)
    final_uploaded_filename: Mapped[str | None] = mapped_column(String(512), nullable=True)

    github_path: Mapped[str | None] = mapped_column(String(1024), nullable=True)
    github_url: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    github_raw_url: Mapped[str | None] = mapped_column(String(2048), nullable=True)

    status: Mapped[str] = mapped_column(String(64), nullable=False, index=True)
    error_message: Mapped[str | None] = mapped_column(Text, nullable=True)

    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )
