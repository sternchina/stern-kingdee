from __future__ import annotations

import configparser
import importlib
import json
import sys
from pathlib import Path
from typing import Any

from src.config import PROJECT_ROOT, Settings, get_settings


_REQUIRED_CONFIG_KEYS = (
    "x-kdapi-acctid",
    "x-kdapi-username",
    "x-kdapi-appid",
    "x-kdapi-appsec",
    "x-kdapi-serverurl",
)


def _mask(value: str) -> str:
    text = value.strip()
    if not text:
        return "(empty)"
    if len(text) <= 6:
        return "*" * len(text)
    return f"{text[:3]}***{text[-3:]}"


def _normalize_response(response: Any) -> Any:
    if isinstance(response, str):
        stripped = response.strip()
        if stripped.startswith("{") or stripped.startswith("["):
            try:
                return json.loads(stripped)
            except json.JSONDecodeError:
                return response
    return response


def _pick_locale_value(values: Any, locale_key: int = 2052) -> str:
    if not isinstance(values, list) or not values:
        return ""
    for item in values:
        if isinstance(item, dict) and item.get("Key") == locale_key:
            return str(item.get("Value", ""))
    first = values[0]
    if isinstance(first, dict):
        return str(first.get("Value", ""))
    return ""


def _build_form_field_summary(raw: dict[str, Any]) -> dict[str, Any]:
    result = raw.get("Result", {})
    status = result.get("ResponseStatus", {})
    if not status.get("IsSuccess", False):
        errors = status.get("Errors", [])
        messages = [str(item.get("Message", "")) for item in errors if isinstance(item, dict)]
        raise RuntimeError("QueryBusinessInfo failed: " + "; ".join(msg for msg in messages if msg))

    data = result.get("NeedReturnData", {})
    entries_summary: list[dict[str, Any]] = []
    all_field_keys: list[str] = []
    for entry in data.get("Entrys", []):
        fields_summary: list[dict[str, Any]] = []
        field_keys: list[str] = []
        for field in entry.get("Fields", []):
            key = str(field.get("Key", "")).strip()
            if not key:
                continue
            field_keys.append(key)
            all_field_keys.append(key)
            fields_summary.append(
                {
                    "key": key,
                    "name": _pick_locale_value(field.get("Name")),
                    "db_field": str(field.get("FieldName", "")),
                    "property_name": str(field.get("PropertyName", "")),
                    "must_input": bool(field.get("MustInput")),
                    "is_view_visible": bool(field.get("IsViewVisible")),
                    "is_edit_visible": bool(field.get("IsEditVisible")),
                    "entity_key": str(field.get("EntityKey", "")),
                    "table_name": str(field.get("TableName", "")),
                }
            )

        entries_summary.append(
            {
                "entry_key": str(entry.get("Key", "")),
                "entry_name": _pick_locale_value(entry.get("Name")),
                "table_name": str(entry.get("TableName", "")),
                "parent_key": entry.get("ParentKey"),
                "field_count": len(fields_summary),
                "field_keys": field_keys,
                "field_keys_csv": ",".join(field_keys),
                "fields": fields_summary,
            }
        )

    unique_field_keys = list(dict.fromkeys(all_field_keys))
    return {
        "form_id": str(data.get("Id", "")),
        "form_name": _pick_locale_value(data.get("Name")),
        "entry_count": len(entries_summary),
        "field_count": len(unique_field_keys),
        "field_keys": unique_field_keys,
        "field_keys_csv": ",".join(unique_field_keys),
        "entries": entries_summary,
    }


class KingdeeClient:
    """基于官方 k3cloud_webapi_sdk 的只读封装。"""

    def __init__(self, settings: Settings | None = None) -> None:
        self._s = settings or get_settings()
        self._sdk: Any | None = None

    def close(self) -> None:
        self._sdk = None

    def __enter__(self) -> KingdeeClient:
        return self

    def __exit__(self, *args: object) -> None:
        self.close()

    def _discover_sdk_wheel(self) -> Path | None:
        configured = self._s.resolved_kingdee_sdk_wheel_path
        if configured is not None and configured.is_file():
            return configured

        vendor_dir = PROJECT_ROOT / ".vendor"
        candidates = sorted(vendor_dir.glob("kingdee*.whl"), reverse=True)
        return candidates[0] if candidates else None

    def _load_sdk_class(self) -> type[Any]:
        try:
            module = importlib.import_module("k3cloud_webapi_sdk.main")
        except ModuleNotFoundError:
            wheel_path = self._discover_sdk_wheel()
            if wheel_path is not None:
                wheel_str = str(wheel_path)
                if wheel_str not in sys.path:
                    sys.path.insert(0, wheel_str)
                module = importlib.import_module("k3cloud_webapi_sdk.main")
            else:
                raise RuntimeError(
                    "Official Kingdee SDK is not installed. Place the wheel under .vendor/ "
                    "or set KINGDEE_SDK_WHEEL_PATH, for example: "
                    "C:\\Users\\KiCo\\Downloads\\SDK_Python3.0_V8.2.0\\python_sdk_v8.2.0\\"
                    "kingdee.cdp.webapi.sdk-8.2.0-py3-none-any.whl"
                )

        sdk_class = getattr(module, "K3CloudApiSdk", None)
        if sdk_class is None:
            raise RuntimeError("k3cloud_webapi_sdk.main does not expose K3CloudApiSdk")
        return sdk_class

    def _read_config(self) -> tuple[Path, configparser.SectionProxy]:
        config_path = self._s.resolved_kingdee_config_path
        if not config_path.is_file():
            raise FileNotFoundError(f"Kingdee SDK config file not found: {config_path}")

        parser = configparser.ConfigParser()
        parser.read(config_path, encoding="utf-8")
        if not parser.has_section(self._s.kingdee_config_node):
            raise ValueError(
                f"Config node [{self._s.kingdee_config_node}] was not found in {config_path}"
            )
        return config_path, parser[self._s.kingdee_config_node]

    def describe_runtime(self) -> dict[str, Any]:
        config_path, section = self._read_config()
        missing = [key for key in _REQUIRED_CONFIG_KEYS if not section.get(key, "").strip()]
        sdk_available = True
        sdk_error = ""
        sdk_class_name = ""
        sdk_wheel = self._discover_sdk_wheel()
        raw_server_url = section.get("X-KDApi-ServerUrl", "").strip()
        request_timeout = section.getint("X-KDApi-RequestTimeout", fallback=120)
        try:
            sdk_class = self._load_sdk_class()
            sdk_class_name = f"{sdk_class.__module__}.{sdk_class.__name__}"
        except RuntimeError as exc:
            sdk_available = False
            sdk_error = str(exc)
        return {
            "sdk_available": sdk_available,
            "sdk_class": sdk_class_name,
            "sdk_error": sdk_error,
            "sdk_wheel_path": str(sdk_wheel) if sdk_wheel is not None else "",
            "config_path": str(config_path),
            "config_node": self._s.kingdee_config_node,
            "missing_keys": missing,
            "login_mode": "sdk-managed",
            "server_url": _mask(raw_server_url),
            "raw_server_url": raw_server_url,
            "request_timeout": request_timeout,
            "acct_id": _mask(section.get("X-KDApi-AcctID", "")),
            "username": section.get("X-KDApi-UserName", "").strip() or "(empty)",
            "app_id": _mask(section.get("X-KDApi-AppID", "")),
        }

    def _ensure_sdk(self) -> Any:
        if self._sdk is not None:
            return self._sdk

        runtime = self.describe_runtime()
        if not runtime["sdk_available"]:
            raise RuntimeError(runtime["sdk_error"])
        if runtime["missing_keys"]:
            missing = ", ".join(runtime["missing_keys"])
            raise ValueError(f"Kingdee SDK config is missing required keys: {missing}")

        sdk_class = self._load_sdk_class()
        server_url = runtime["raw_server_url"]
        request_timeout = runtime["request_timeout"]
        sdk = sdk_class(server_url, timeout=request_timeout)
        sdk.Init(
            config_path=str(self._s.resolved_kingdee_config_path),
            config_node=self._s.kingdee_config_node,
        )
        self._sdk = sdk
        return sdk

    def execute_bill_query(self, query: dict[str, Any]) -> Any:
        sdk = self._ensure_sdk()
        return _normalize_response(sdk.ExecuteBillQuery(query))

    def view(self, form_id: str, para: dict[str, Any]) -> Any:
        name = form_id.strip()
        if not name:
            raise ValueError("form_id must not be empty")
        sdk = self._ensure_sdk()
        return _normalize_response(sdk.View(name, para))

    def query_business_info(self, form_id: str) -> dict[str, Any]:
        name = form_id.strip()
        if not name:
            raise ValueError("form_id must not be empty")
        sdk = self._ensure_sdk()
        raw = _normalize_response(sdk.QueryBusinessInfo({"FormId": name}))
        if not isinstance(raw, dict):
            raise RuntimeError("QueryBusinessInfo returned an unexpected payload")
        return raw

    def get_form_field_keys(self, form_id: str) -> dict[str, Any]:
        return _build_form_field_summary(self.query_business_info(form_id))
