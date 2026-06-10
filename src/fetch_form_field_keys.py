from __future__ import annotations

from typing import Any

from src.config import Settings, get_settings
from src.kingdee_client import KingdeeClient


def fetch_form_field_keys(form_id: str, *, settings: Settings | None = None) -> dict[str, Any]:
    """通过 FormId 读取元数据，并返回可用于查询的字段清单。"""
    s = settings or get_settings()
    with KingdeeClient(s) as client:
        return client.get_form_field_keys(form_id)
