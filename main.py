"""
CLI：检查金蝶 SDK 配置、执行只读查询、导出演示 Excel。

使用前复制 .env.example 为 .env、conf.ini.example 为 conf.ini，并安装官方 SDK。
详见 docs/kingdee-python-integration.md
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# 自项目根运行：python main.py
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def main() -> int:
    parser = argparse.ArgumentParser(description="金蝶云星空只读集成 CLI")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_check = sub.add_parser(
        "check",
        aliases=["login"],
        help="检查官方 SDK 与 conf.ini 配置；业务调用时由 SDK 隐式登录",
    )
    p_check.add_argument("--dry-run", action="store_true", help="仅打印配置检查结果")

    p_query = sub.add_parser("query", help="通过 SDK 执行一次 ExecuteBillQuery（需 KINGDEE_QUERY_JSON）")
    p_query.add_argument("--dry-run", action="store_true", help="只解析 JSON，不请求")

    p_view = sub.add_parser("view", help="通过 SDK 执行一次 View（需 KINGDEE_VIEW_FORM_ID 与 KINGDEE_VIEW_PARA_JSON）")
    p_view.add_argument("--dry-run", action="store_true", help="只解析参数，不请求")

    p_fields = sub.add_parser("fields", help="通过 FormId 查询字段元数据，并返回可用 FieldKeys")
    p_fields.add_argument("form_id", help="业务对象标识，例如 QM_QualityStd")

    p_demo = sub.add_parser("export-demo", help="写入 data/out/demo_export.xlsx（不连金蝶）")

    args = parser.parse_args()

    try:
        if args.cmd == "export-demo":
            from src.export_excel import export_sample_demo

            out = export_sample_demo(ROOT / "data" / "out")
            print(f"Wrote: {out}")
            return 0

        from src.config import get_settings
        from src.kingdee_client import KingdeeClient

        settings = get_settings()

        if args.cmd in ("check", "login"):
            with KingdeeClient(settings) as client:
                data = client.describe_runtime()
            data.pop("raw_server_url", None)
            if args.dry_run:
                data["dry_run"] = True
            print(json.dumps(data, ensure_ascii=False, indent=2))
            return 0 if data["sdk_available"] and not data["missing_keys"] else 1

        if args.cmd == "query":
            if not settings.kingdee_query_json:
                print("Set KINGDEE_QUERY_JSON in .env (ExecuteBillQuery JSON).", file=sys.stderr)
                return 1
            query = json.loads(settings.kingdee_query_json)
            if args.dry_run:
                print(json.dumps(query, ensure_ascii=False, indent=2))
                return 0
            with KingdeeClient(settings) as client:
                data = client.execute_bill_query(query)
            print(json.dumps(_redact_response(data), ensure_ascii=False, indent=2))
            return 0

        if args.cmd == "view":
            form_id = (settings.kingdee_view_form_id or "").strip()
            para_json = settings.kingdee_view_para_json
            if not form_id or not para_json:
                print("Set both KINGDEE_VIEW_FORM_ID and KINGDEE_VIEW_PARA_JSON in .env.", file=sys.stderr)
                return 1
            para = json.loads(para_json)
            if args.dry_run:
                print(json.dumps({"FormId": form_id, "Parameters": para}, ensure_ascii=False, indent=2))
                return 0
            with KingdeeClient(settings) as client:
                data = client.view(form_id, para)
            print(json.dumps(_redact_response(data), ensure_ascii=False, indent=2))
            return 0

        if args.cmd == "fields":
            with KingdeeClient(settings) as client:
                data = client.get_form_field_keys(args.form_id)
            print(json.dumps(_redact_response(data), ensure_ascii=False, indent=2))
            return 0
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except json.JSONDecodeError as exc:
        print(f"Invalid JSON in .env settings: {exc}", file=sys.stderr)
        return 1

    return 1


def _redact_response(data: object) -> object:
    """避免把过长业务数据刷屏时可再收缩；当前仅原样返回结构。"""
    if isinstance(data, dict) and len(json.dumps(data, ensure_ascii=False)) > 200_000:
        return {"_note": "Response too large; persist or parse selectively in code.", "keys": list(data.keys())}
    return data


if __name__ == "__main__":
    raise SystemExit(main())
