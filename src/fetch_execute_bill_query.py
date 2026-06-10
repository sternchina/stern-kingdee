from __future__ import annotations

from typing import Any

from src.config import Settings, get_settings
from src.kingdee_client import KingdeeClient


def fetch_execute_bill_query(query: dict[str, Any], *, settings: Settings | None = None) -> Any:
    """通过官方 SDK 执行 ExecuteBillQuery，结果仅在内存中返回。"""
    s = settings or get_settings()
    with KingdeeClient(s) as client:
        return client.execute_bill_query(query)
