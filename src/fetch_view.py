from __future__ import annotations

from typing import Any

from src.config import Settings, get_settings
from src.kingdee_client import KingdeeClient


def fetch_view(form_id: str, para: dict[str, Any], *, settings: Settings | None = None) -> Any:
    """通过官方 SDK 执行 View，结果仅在内存中返回。"""
    s = settings or get_settings()
    with KingdeeClient(s) as client:
        return client.view(form_id, para)
