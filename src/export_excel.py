from __future__ import annotations

from pathlib import Path

import pandas as pd


def export_dataframe(df: pd.DataFrame, out_path: Path, *, sheet_name: str = "Sheet1") -> None:
    """将 DataFrame 写入 xlsx（目录不存在则创建）。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out_path, sheet_name=sheet_name, index=False, engine="openpyxl")


def export_sample_demo(out_dir: Path) -> Path:
    """生成示例 Excel，验证环境与 openpyxl。"""
    df = pd.DataFrame(
        {
            "示例列A": [1, 2, 3],
            "示例列B": ["金蝶", "云星空", "只读"],
        }
    )
    path = out_dir / "demo_export.xlsx"
    export_dataframe(df, path)
    return path
