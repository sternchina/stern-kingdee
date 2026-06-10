from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


PROJECT_ROOT = Path(__file__).resolve().parent.parent


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    kingdee_sdk_wheel_path: Path | None = Field(default=None, validation_alias="KINGDEE_SDK_WHEEL_PATH")
    kingdee_config_path: Path = Field(default=Path("conf.ini"), validation_alias="KINGDEE_CONFIG_PATH")
    kingdee_config_node: str = Field(default="config", validation_alias="KINGDEE_CONFIG_NODE")

    kingdee_query_json: str | None = Field(default=None, validation_alias="KINGDEE_QUERY_JSON")
    kingdee_view_form_id: str | None = Field(default=None, validation_alias="KINGDEE_VIEW_FORM_ID")
    kingdee_view_para_json: str | None = Field(default=None, validation_alias="KINGDEE_VIEW_PARA_JSON")

    @field_validator("kingdee_config_node")
    @classmethod
    def _validate_config_node(cls, v: str) -> str:
        s = v.strip()
        if not s:
            raise ValueError("KINGDEE_CONFIG_NODE must not be empty")
        return s

    @property
    def resolved_kingdee_config_path(self) -> Path:
        path = self.kingdee_config_path
        return path if path.is_absolute() else (PROJECT_ROOT / path).resolve()

    @property
    def resolved_kingdee_sdk_wheel_path(self) -> Path | None:
        path = self.kingdee_sdk_wheel_path
        if path is None:
            return None
        return path if path.is_absolute() else (PROJECT_ROOT / path).resolve()


@lru_cache
def get_settings() -> Settings:
    return Settings()
