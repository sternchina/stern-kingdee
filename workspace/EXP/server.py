"""
BOM_API 数据同步工具 - Flask 后端服务

提供 Web 界面和 API，用于启动主程序并追踪进度。

用法：
    python workspace/BOM_API/server.py
    浏览器访问 http://localhost:5000
"""
from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory

# ── 路径 ───────────────────────────────────────────────────────────────────────
WORKSPACE = Path(__file__).resolve().parent
PROJECT_ROOT = WORKSPACE.parent.parent
SYNC_SCRIPT = WORKSPACE / "sync_and_generate.py"
OUTPUT_DIR = WORKSPACE / "output"
LOG_DIR = WORKSPACE / "log"
DEFAULT_NETWORK_OUTPUT = r"\\192.168.1.181\12_SupplyChain\07_MPS\MPS_Upgrade"

# 文件名 → 网络子目录映射
FILE_NET_MAP = {
    "生产领料单列表.xlsx": "K3_Data_PickingList",
    "K3_Pickinglist.xlsx": "K3_Data_PickingList",
    "生产入库单列表.xlsx": "K3_Data_Production_Receipts_MTD",
    "K3_Production_Receipts.xlsx": "K3_Data_Production_Receipts_MTD",
    "生产订单列表.xlsx": "K3_Data_ProductionOrders",
    "K3_ManufacturingOrders.xlsx": "K3_Data_ProductionOrders",
    "计划订单列表.xlsx": "K3_Data_ProductionRequests",
    "K3_PlannedOrders.xlsx": "K3_Data_ProductionRequests",
    "采购订单列表.xlsx": "K3_Data_PurchaseOrders",
    "K3_PurchaseOrders.xlsx": "K3_Data_PurchaseOrders",
    "预测单列表.xlsx": "K3_Data_SalesForecast",
    "K3_Forecast.xlsx": "K3_Data_SalesForecast",
}

# 专用网络同步路径
BOM_NETWORK_PATH = r"\\192.168.1.181\99_Share\04_IT\KingdeeExport\K3_Data_BOM"
# Material Master 专用网络同步路径
MAT_NETWORK_PATH = r"\\192.168.1.181\99_Share\04_IT\KingdeeExport\K3_MaterialMaster"
# 预测单专用网络同步路径
FORECAST_NETWORK_PATH = r"\\192.168.1.181\99_Share\04_IT\KingdeeExport\K3_Data_SalesForecast"
FORECAST_NAMES = {"预测单列表.xlsx", "K3_Forecast.xlsx"}

# 即时库存专用网络同步路径（按语言区分）
# 中文版 → 财务部下载目录 | 英文版 → IT部导出目录
INV_NETWORK_PATHS = {
    "CN": r"\\192.168.1.181\05_Finance\Download Report from Kingdee\SICHI",
    "EN": r"\\192.168.1.181\99_Share\04_IT\KingdeeExport\K3_Data_Inventory",
}
PROGRESS_FILE = WORKSPACE / ".progress.json"

# 当前运行时使用的网络输出路径
_current_network_output: str = DEFAULT_NETWORK_OUTPUT

app = Flask(__name__, static_folder=None)

# ── 全局状态 ───────────────────────────────────────────────────────────────────
_current_process: subprocess.Popen | None = None
_process_lock = threading.Lock()

# ── 表单定义 ───────────────────────────────────────────────────────────────────
FORM_DEFS = [
    {"id": "BD_MATERIAL", "name": "物料基础资料", "desc": "物料主数据"},
    {"id": "ENG_BOM", "name": "物料清单", "desc": "BOM 清单"},
    {"id": "PLN_FORECAST", "name": "预测单", "desc": "计划预测"},
    {"id": "PLN_PLANORDER", "name": "计划订单", "desc": "计划订单"},
    {"id": "PRD_MO", "name": "生产订单", "desc": "生产制造订单"},
    {"id": "PUR_PurchaseOrder", "name": "采购订单", "desc": "采购订单"},
    {"id": "PRD_INSTOCK", "name": "生产入库单", "desc": "生产入库"},
    {"id": "STK_Inventory", "name": "即时库存", "desc": "库存查询"},
    {"id": "PRD_PickMtrl", "name": "生产领料单", "desc": "生产领料（需单独选择）"},
]

LANG_DEFS = [
    {"id": "CN", "name": "中文版", "desc": "Chinese"},
    {"id": "EN", "name": "英文版", "desc": "English"},
]


# ── 进度文件操作 ───────────────────────────────────────────────────────────────
def _init_progress(total_steps: int, forms: list[str], langs: list[str]) -> None:
    data = {
        "running": True,
        "progress": 0,
        "step": "正在初始化...",
        "total_steps": total_steps,
        "completed_steps": 0,
        "forms": forms,
        "langs": langs,
        "start_time": datetime.now().isoformat(),
        "end_time": None,
        "error": None,
    }
    PROGRESS_FILE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def _update_progress(completed_steps: int, step: str, progress: int) -> None:
    if not PROGRESS_FILE.exists():
        return
    data = json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))
    data["completed_steps"] = completed_steps
    data["step"] = step
    data["progress"] = progress
    PROGRESS_FILE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def _finish_progress(error: str | None = None) -> None:
    if not PROGRESS_FILE.exists():
        return
    data = json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))
    data["running"] = False
    data["progress"] = 100 if error is None else data["progress"]
    data["step"] = "完成!" if error is None else f"出错: {error}"
    data["end_time"] = datetime.now().isoformat()
    data["error"] = error
    PROGRESS_FILE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def _read_progress() -> dict:
    if not PROGRESS_FILE.exists():
        return {
            "running": False,
            "progress": 0,
            "step": "",
            "total_steps": 0,
            "completed_steps": 0,
        }
    return json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))


def _sync_to_network(network_root: str) -> dict | None:
    """将本地 output 目录中的 xlsx 文件同步到网络共享目录，每个文件对应独立子目录。"""
    if not OUTPUT_DIR.exists():
        print(f"[同步] 本地 output 目录不存在，跳过同步")
        return None
    if not network_root:
        print(f"[同步] 未指定网络输出路径，跳过同步")
        return None

    xlsx_files = list(OUTPUT_DIR.glob("*.xlsx"))
    print(f"[同步] 本地 output 共 {len(xlsx_files)} 个 xlsx 文件，目标根: {network_root}")

    if not xlsx_files:
        print(f"[同步] 无 xlsx 文件需要复制")
        return {"copied": 0, "total": 0, "log": "无文件需复制"}

    all_log_lines = []

    # 分离即时库存文件（使用独立语言路径）和 BOM 文件
    inv_cn_name = "即时库存.xlsx"
    inv_en_name = "K3_Inventory.xlsx"
    inv_files = [f for f in xlsx_files if f.name in (inv_cn_name, inv_en_name)]
    bom_names = {"物料清单列表.xlsx", "K3_BOM.xlsx"}
    mat_names = {"物料列表.xlsx", "K3_MaterialMaster.xlsx"}
    bom_files = [f for f in xlsx_files if f.name in bom_names]
    mat_files = [f for f in xlsx_files if f.name in mat_names]
    forecast_files = [f for f in xlsx_files if f.name in FORECAST_NAMES]
    other_files = [f for f in xlsx_files if f not in inv_files and f not in bom_files and f not in mat_files and f not in forecast_files]

    # 网络认证参数
    net_user = r"STERN\shenjj"
    net_pass = r"zH820303."

    # 生成临时批处理脚本：映射驱动器 → 创建子目录 → 复制 → 断开驱动器
    import tempfile
    bat_lines = [
        'cd /d C:\\',
        f'net use * /delete /yes 2>nul',
        f'timeout /t 1 /nobreak >nul',
    ]

    # ── BOM 专用路径（映射到 W:）──
    if bom_files:
        bom_parts = Path(BOM_NETWORK_PATH).parts
        bom_share = bom_parts[0].rstrip("\\") if bom_parts else BOM_NETWORK_PATH
        bom_sub = str(Path(BOM_NETWORK_PATH)).replace(bom_share, "").strip("\\")
        bat_lines.append(f'net use W: {bom_share} {net_pass} /user:{net_user} /persistent:no')
        for f in bom_files:
            print(f"[同步] BOM ({f.name}) → {BOM_NETWORK_PATH}")
            src = str(f).replace("/", "\\")
            if bom_sub:
                dst_dir = f"W:\\{bom_sub}"
            else:
                dst_dir = f"W:"
            bat_lines.append(f'if not exist "{dst_dir}" mkdir "{dst_dir}"')
            dst = f"{dst_dir}\\{f.name}"
            bat_lines.append(f'copy /Y "{src}" "{dst}"')
        bat_lines.append(f'net use W: /delete /yes')

    # ── Material Master 专用路径（映射到 V:）──
    if mat_files:
        mat_parts = Path(MAT_NETWORK_PATH).parts
        mat_share = mat_parts[0].rstrip("\\") if mat_parts else MAT_NETWORK_PATH
        mat_sub = str(Path(MAT_NETWORK_PATH)).replace(mat_share, "").strip("\\")
        bat_lines.append(f'net use V: {mat_share} {net_pass} /user:{net_user} /persistent:no')
        for f in mat_files:
            print(f"[同步] Material ({f.name}) → {MAT_NETWORK_PATH}")
            src = str(f).replace("/", "\\")
            if mat_sub:
                dst_dir = f"V:\\{mat_sub}"
            else:
                dst_dir = f"V:"
            bat_lines.append(f'if not exist "{dst_dir}" mkdir "{dst_dir}"')
            dst = f"{dst_dir}\\{f.name}"
            bat_lines.append(f'copy /Y "{src}" "{dst}"')
        bat_lines.append(f'net use V: /delete /yes')

    # ── 预测单专用路径（映射到 U:）──
    if forecast_files:
        forecast_parts = Path(FORECAST_NETWORK_PATH).parts
        forecast_share = forecast_parts[0].rstrip("\\") if forecast_parts else FORECAST_NETWORK_PATH
        forecast_sub = str(Path(FORECAST_NETWORK_PATH)).replace(forecast_share, "").strip("\\")
        bat_lines.append(f'net use U: {forecast_share} {net_pass} /user:{net_user} /persistent:no')
        for f in forecast_files:
            print(f"[同步] 预测单 ({f.name}) → {FORECAST_NETWORK_PATH}")
            src = str(f).replace("/", "\\")
            if forecast_sub:
                dst_dir = f"U:\\{forecast_sub}"
            else:
                dst_dir = f"U:"
            bat_lines.append(f'if not exist "{dst_dir}" mkdir "{dst_dir}"')
            dst = f"{dst_dir}\\{f.name}"
            bat_lines.append(f'copy /Y "{src}" "{dst}"')
        bat_lines.append(f'net use U: /delete /yes')

    # ── 主共享：其他文件（映射到 Z:）──
    if other_files:
        parts = Path(network_root).parts
        share_root = parts[0].rstrip("\\") if parts else network_root
        root_sub_path = str(Path(network_root)).replace(share_root, "").strip("\\")
        bat_lines.append(f'net use Z: {share_root} {net_pass} /user:{net_user} /persistent:no')
        for f in other_files:
            sub_dir = FILE_NET_MAP.get(f.name)
            if not sub_dir:
                print(f"[同步] 跳过未映射文件: {f.name}")
                continue
            src = str(f).replace("/", "\\")
            if root_sub_path:
                dst_dir = f"Z:\\{root_sub_path}\\{sub_dir}"
            else:
                dst_dir = f"Z:\\{sub_dir}"
            bat_lines.append(f'if not exist "{dst_dir}" mkdir "{dst_dir}"')
            dst = f"{dst_dir}\\{f.name}"
            bat_lines.append(f'copy /Y "{src}" "{dst}"')
        bat_lines.append(f'net use Z: /delete /yes')

    # ── 即时库存专用路径（按语言映射到独立驱动器）──
    inv_drives = ["Y:", "X:"]  # Y→CN, X→EN
    for f in inv_files:
        if f.name == inv_cn_name:
            inv_path = INV_NETWORK_PATHS["CN"]
            drive = inv_drives[0]
        else:
            inv_path = INV_NETWORK_PATHS["EN"]
            drive = inv_drives[1]

        inv_parts = Path(inv_path).parts
        inv_share = inv_parts[0].rstrip("\\") if inv_parts else inv_path
        inv_sub = str(Path(inv_path)).replace(inv_share, "").strip("\\")

        print(f"[同步] 即时库存 ({'CN' if f.name == inv_cn_name else 'EN'}) → {inv_path}")
        bat_lines.append(f'net use {drive} {inv_share} {net_pass} /user:{net_user} /persistent:no')
        src = str(f).replace("/", "\\")
        if inv_sub:
            dst_dir = f"{drive}\\{inv_sub}"
        else:
            dst_dir = f"{drive}"
        bat_lines.append(f'if not exist "{dst_dir}" mkdir "{dst_dir}"')
        dst = f"{dst_dir}\\{f.name}"
        bat_lines.append(f'copy /Y "{src}" "{dst}"')
        bat_lines.append(f'net use {drive} /delete /yes')

    bat_lines.append('exit /b 0')

    bat_content = "\r\n".join(bat_lines)
    bat_file = Path(tempfile.gettempdir()) / "bom_sync.bat"
    bat_file.write_text(bat_content, encoding="gbk")

    print(f"[同步] 执行批处理同步")
    try:
        result = subprocess.run(
            ["cmd.exe", "/c", str(bat_file)],
            cwd="C:\\",
            capture_output=True, text=True, timeout=180,
            encoding="gbk", errors="replace",
        )
        output = (result.stdout or "") + (result.stderr or "")
        print(f"[同步] 批处理返回码: {result.returncode}")
        if output.strip():
            for line in output.strip().splitlines():
                line = line.strip()
                if line:
                    print(f"[同步] {line}")
                    all_log_lines.append(line)
    except Exception as e:
        msg = f"批处理执行失败: {e}"
        print(f"[同步] {msg}")
        all_log_lines.append(msg)
        return {"copied": 0, "total": len(xlsx_files), "log": "\n".join(all_log_lines)}
    finally:
        try:
            bat_file.unlink()
        except OSError:
            pass

    # 统计复制结果：检查目标目录中的文件
    copied = 0
    for f in xlsx_files:
        if f.name in (inv_cn_name, inv_en_name):
            # 即时库存：检查独立路径
            if f.name == inv_cn_name:
                dest_dir = Path(INV_NETWORK_PATHS["CN"])
            else:
                dest_dir = Path(INV_NETWORK_PATHS["EN"])
            dest_file = dest_dir / f.name
        elif f.name in bom_names:
            # BOM：检查独立路径
            dest_dir = Path(BOM_NETWORK_PATH)
            dest_file = dest_dir / f.name
        elif f.name in mat_names:
            # Material Master：检查独立路径
            dest_dir = Path(MAT_NETWORK_PATH)
            dest_file = dest_dir / f.name
        elif f.name in FORECAST_NAMES:
            # 预测单：检查独立路径
            dest_dir = Path(FORECAST_NETWORK_PATH)
            dest_file = dest_dir / f.name
        else:
            sub_dir = FILE_NET_MAP.get(f.name)
            if sub_dir:
                dest_file = Path(network_root) / sub_dir / f.name
            else:
                dest_file = Path(network_root) / f.name
        if dest_file.exists():
            copied += 1

    print(f"[同步] 完成，成功复制 {copied}/{len(xlsx_files)} 个文件")
    return {"copied": copied, "total": len(xlsx_files), "log": "\n".join(all_log_lines)}


# ── 日志写入 ───────────────────────────────────────────────────────────────────
def _write_run_log(
    forms: list[str],
    langs: list[str],
    network_output: str,
    start_time: str,
    end_time: str,
    success: bool,
    error_msg: str | None,
    step_results: list[dict],
    sync_detail: dict | None = None,
) -> Path:
    """将本次运行记录写入日志文件，返回日志文件路径。"""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = LOG_DIR / f"run_{ts}.log"

    lines = []
    lines.append("=" * 60)
    lines.append("BOM 数据同步运行日志")
    lines.append("=" * 60)
    lines.append("")
    lines.append(f"开始时间:  {start_time}")
    lines.append(f"结束时间:  {end_time}")
    lines.append(f"运行状态:  {'成功' if success else '失败'}")
    lines.append("")
    lines.append("-" * 40)
    lines.append("运行参数")
    lines.append("-" * 40)
    lines.append(f"表单类型:  {', '.join(forms)}")
    lines.append(f"语言:      {', '.join(langs)}")
    lines.append(f"本地输出:  {OUTPUT_DIR}")
    lines.append(f"网络输出:  {network_output}")
    lines.append("")
    lines.append("-" * 40)
    lines.append("执行步骤")
    lines.append("-" * 40)

    for idx, r in enumerate(step_results, 1):
        status = "OK" if r.get("ok") else "FAIL"
        lines.append(
            f"[{idx}] {r.get('form','')} ({r.get('lang','')}) - "
            f"{status}"
        )
        if not r.get("ok") and r.get("error"):
            lines.append(f"     错误: {r['error']}")

    lines.append("")
    lines.append("-" * 40)
    lines.append("网络同步结果")
    lines.append("-" * 40)

    if sync_detail:
        copied = sync_detail.get("copied", 0)
        total = sync_detail.get("total", 0)
        rc = sync_detail.get("returncode", "N/A")
        lines.append(f"目标路径:   {network_output}")
        lines.append(f"复制结果:   {copied}/{total} 个文件")
        lines.append(f"robocopy 返回码: {rc}")

        sync_log = sync_detail.get("log", "")
        if sync_log:
            lines.append("")
            lines.append("  --- robocopy 详细日志 ---")
            for line in sync_log.splitlines():
                lines.append(f"  {line}")
            lines.append("  --- robocopy 日志结束 ---")
    else:
        lines.append("未执行同步")

    if error_msg:
        lines.append("")
        lines.append("-" * 40)
        lines.append("错误信息")
        lines.append("-" * 40)
        lines.append(error_msg)

    lines.append("")
    lines.append("=" * 60)

    log_file.write_text("\n".join(lines), encoding="utf-8")
    print(f"[日志] 已写入: {log_file}")
    return log_file


# ── 运行主程序 ─────────────────────────────────────────────────────────────────
def _run_sync(forms: list[str], langs: list[str], network_output: str) -> None:
    global _current_process

    total_steps = len(forms) * len(langs)
    _init_progress(total_steps, forms, langs)
    start_time = datetime.now().isoformat()
    step_results: list[dict] = []
    success = True
    error_msg: str | None = None
    sync_detail: dict | None = None

    try:
        for i, form in enumerate(forms):
            for j, lang in enumerate(langs):
                step_idx = i * len(langs) + j
                form_name = next(
                    (f["name"] for f in FORM_DEFS if f["id"] == form), form
                )
                lang_name = next(
                    (l["name"] for l in LANG_DEFS if l["id"] == lang), lang
                )
                step_text = f"正在生成 {form_name}（{lang_name}）..."
                progress = int((step_idx / total_steps) * 100)
                _update_progress(step_idx, step_text, progress)

                cmd = [
                    sys.executable,
                    str(SYNC_SCRIPT),
                    "--form", form,
                    "--lang", lang,
                    "--netaddr", "",
                ]

                env = os.environ.copy()
                env["PYTHONIOENCODING"] = "utf-8"

                proc = subprocess.Popen(
                    cmd,
                    cwd=str(PROJECT_ROOT),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    env=env,
                )
                with _process_lock:
                    _current_process = proc

                stdout, stderr = proc.communicate()

                if proc.returncode != 0:
                    err_msg = stderr.decode("utf-8", errors="replace").strip()
                    if not err_msg:
                        err_msg = stdout.decode("utf-8", errors="replace").strip()
                    step_results.append({"form": form_name, "lang": lang_name, "ok": False, "error": err_msg[:200]})
                    success = False
                    error_msg = f"{form}({lang}) 失败: {err_msg[:200]}"
                    _finish_progress(error=error_msg)
                    end_time = datetime.now().isoformat()
                    _write_run_log(forms, langs, network_output, start_time, end_time, False, error_msg, step_results)
                    return

                step_results.append({"form": form_name, "lang": lang_name, "ok": True})

        _update_progress(total_steps, "正在同步到网络共享目录...", 98)
        sync_detail = _sync_to_network(network_output)
        _update_progress(total_steps, "完成!", 100)
        _finish_progress()

    except Exception as exc:
        success = False
        error_msg = str(exc)
        _finish_progress(error=error_msg)
    finally:
        with _process_lock:
            _current_process = None
        # 写入日志
        end_time = datetime.now().isoformat()
        _write_run_log(
            forms, langs, network_output, start_time, end_time,
            success, error_msg, step_results, sync_detail,
        )


# ── 路由 ───────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory(str(WORKSPACE), "index.html")


@app.route("/api/forms", methods=["GET"])
def get_forms():
    return jsonify({"forms": FORM_DEFS, "langs": LANG_DEFS})


@app.route("/api/run", methods=["POST"])
def run():
    global _current_process

    # 检查是否已有任务在运行
    progress = _read_progress()
    if progress.get("running"):
        return jsonify({"success": False, "error": "已有任务正在运行"}), 409

    body = request.get_json(silent=True) or {}
    forms = body.get("forms", [])
    langs = body.get("langs", [])
    network_output = body.get("network_output", DEFAULT_NETWORK_OUTPUT).strip()

    if not forms:
        return jsonify({"success": False, "error": "请至少选择一个表单类型"}), 400
    if not langs:
        return jsonify({"success": False, "error": "请至少选择一种语言"}), 400

    # 验证表单类型
    valid_form_ids = {f["id"] for f in FORM_DEFS}
    for f in forms:
        if f not in valid_form_ids:
            return jsonify({"success": False, "error": f"无效的表单类型: {f}"}), 400

    # 验证语言
    valid_lang_ids = {l["id"] for l in LANG_DEFS}
    for l in langs:
        if l not in valid_lang_ids:
            return jsonify({"success": False, "error": f"无效的语言: {l}"}), 400

    # 保存当前网络输出路径
    global _current_network_output
    _current_network_output = network_output

    # 启动后台线程运行主程序
    thread = threading.Thread(
        target=_run_sync, args=(forms, langs, _current_network_output), daemon=True
    )
    thread.start()

    return jsonify({"success": True, "task_id": datetime.now().strftime("%Y%m%d_%H%M%S")})


@app.route("/api/progress", methods=["GET"])
def get_progress():
    return jsonify(_read_progress())


@app.route("/api/output", methods=["GET"])
def get_output():
    network_root = _current_network_output
    display_dir = network_root or "（未设置网络输出路径）"

    files = []
    inv_cn_name = "即时库存.xlsx"
    inv_en_name = "K3_Inventory.xlsx"
    if OUTPUT_DIR.exists():
        for f in sorted(OUTPUT_DIR.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
            sub_dir = ""
            if f.name in (inv_cn_name, inv_en_name) and network_root:
                # 即时库存使用独立路径
                if f.name == inv_cn_name:
                    net_path = str(Path(INV_NETWORK_PATHS["CN"]) / f.name)
                else:
                    net_path = str(Path(INV_NETWORK_PATHS["EN"]) / f.name)
            elif f.name in bom_names and BOM_NETWORK_PATH:
                # BOM 使用独立路径
                net_path = str(Path(BOM_NETWORK_PATH) / f.name)
            elif f.name in mat_names and MAT_NETWORK_PATH:
                # Material Master 使用独立路径
                net_path = str(Path(MAT_NETWORK_PATH) / f.name)
            elif f.name in FORECAST_NAMES and FORECAST_NETWORK_PATH:
                # 预测单使用独立路径
                net_path = str(Path(FORECAST_NETWORK_PATH) / f.name)
            else:
                sub_dir = FILE_NET_MAP.get(f.name, "")
                if network_root and sub_dir:
                    net_path = str(Path(network_root) / sub_dir / f.name)
                elif network_root:
                    net_path = str(Path(network_root) / f.name)
                else:
                    net_path = ""
            files.append({
                "name": f.name,
                "path": net_path,
                "size": f.stat().st_size,
                "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat(),
                "sub_dir": sub_dir,
            })

    return jsonify({"output_dir": display_dir, "files": files})


def _clear_network(network_path: str | None) -> list[str]:
    """清空所有网络路径中的 xlsx 文件，返回错误列表。"""
    import tempfile

    errors: list[str] = []
    net_user = r"STERN\\shenjj"
    net_pass = r"zH820303."
    bat_lines = [
        'net use * /delete /yes 2>nul',
        'timeout /t 1 /nobreak >nul',
    ]

    # 主共享路径（仅当 network_path 非空）
    if network_path:
        parts = Path(network_path).parts
        share_root = parts[0].rstrip("\\") if parts else network_path
        root_sub_path = str(Path(network_path)).replace(share_root, "").strip("\\")
        sub_dirs = set(FILE_NET_MAP.values())
        bat_lines.append(f'net use Z: {share_root} {net_pass} /user:{net_user} /persistent:no')
        for sd in sub_dirs:
            if root_sub_path:
                pattern = f"{root_sub_path}\\{sd}\\*.xlsx"
            else:
                pattern = f"{sd}\\*.xlsx"
            bat_lines.append(f'del /Q /F "Z:\\{pattern}"')
        bat_lines.append(f'net use Z: /delete /yes')

    # 即时库存专用网络路径（始终清空，独立于主 network_path）
    inv_cleared: set[str] = set()
    for inv_path in INV_NETWORK_PATHS.values():
        if inv_path in inv_cleared:
            continue
        inv_cleared.add(inv_path)
        inv_parts = Path(inv_path).parts
        inv_share = inv_parts[0].rstrip("\\") if inv_parts else inv_path
        inv_sub = str(Path(inv_path)).replace(inv_share, "").strip("\\")
        bat_lines.append(f'net use Y: {inv_share} {net_pass} /user:{net_user} /persistent:no')
        if inv_sub:
            pattern = f"{inv_sub}\\*.xlsx"
        else:
            pattern = "*.xlsx"
        bat_lines.append(f'del /Q /F "Y:\\{pattern}"')
        bat_lines.append(f'net use Y: /delete /yes')

    # BOM 专用网络路径（始终清空）
    bom_parts = Path(BOM_NETWORK_PATH).parts
    bom_share = bom_parts[0].rstrip("\\") if bom_parts else BOM_NETWORK_PATH
    bom_sub = str(Path(BOM_NETWORK_PATH)).replace(bom_share, "").strip("\\")
    bat_lines.append(f'net use W: {bom_share} {net_pass} /user:{net_user} /persistent:no')
    if bom_sub:
        pattern = f"{bom_sub}\\*.xlsx"
    else:
        pattern = "*.xlsx"
    bat_lines.append(f'del /Q /F "W:\\{pattern}"')
    bat_lines.append(f'net use W: /delete /yes')

    # Material Master 专用网络路径（始终清空）
    mat_parts = Path(MAT_NETWORK_PATH).parts
    mat_share = mat_parts[0].rstrip("\\") if mat_parts else MAT_NETWORK_PATH
    mat_sub = str(Path(MAT_NETWORK_PATH)).replace(mat_share, "").strip("\\")
    bat_lines.append(f'net use U: {mat_share} {net_pass} /user:{net_user} /persistent:no')
    if mat_sub:
        pattern = f"{mat_sub}\\*.xlsx"
    else:
        pattern = "*.xlsx"
    bat_lines.append(f'del /Q /F "U:\\{pattern}"')
    bat_lines.append(f'net use U: /delete /yes')

    # 预测单专用网络路径（始终清空）
    forecast_parts = Path(FORECAST_NETWORK_PATH).parts
    forecast_share = forecast_parts[0].rstrip("\\") if forecast_parts else FORECAST_NETWORK_PATH
    forecast_sub = str(Path(FORECAST_NETWORK_PATH)).replace(forecast_share, "").strip("\\")
    bat_lines.append(f'net use T: {forecast_share} {net_pass} /user:{net_user} /persistent:no')
    if forecast_sub:
        pattern = f"{forecast_sub}\\*.xlsx"
    else:
        pattern = "*.xlsx"
    bat_lines.append(f'del /Q /F "T:\\{pattern}"')
    bat_lines.append(f'net use T: /delete /yes')

    bat_lines.append('exit /b 0')

    bat_content = "\r\n".join(bat_lines)
    bat_file = Path(tempfile.gettempdir()) / "bom_clear.bat"
    bat_file.write_text(bat_content, encoding="gbk")

    try:
        result = subprocess.run(
            ["cmd.exe", "/c", str(bat_file)],
            cwd="C:\\",
            capture_output=True, text=True, timeout=30,
            encoding="gbk", errors="replace",
        )
        output = (result.stdout or "") + (result.stderr or "")
        logging.info(f"[清空] 返回码: {result.returncode}, 输出: {output}")
        if output.strip():
            for line in output.strip().splitlines():
                line = line.strip()
                if line:
                    print(f"[清空] {line}")

        # 统计各网络路径残留文件（递归检查子目录）
        def _count_remaining_xlsx(dir_path: str) -> int:
            p = Path(dir_path)
            if not p.exists():
                return -1
            return len(list(p.glob("**/*.xlsx")))

        for label, d in [
            ("主共享", network_path),
            ("即时库存(CN)", INV_NETWORK_PATHS["CN"]),
            ("即时库存(EN)", INV_NETWORK_PATHS["EN"]),
            ("BOM", BOM_NETWORK_PATH),
            ("Material Master", MAT_NETWORK_PATH),
            ("预测单", FORECAST_NETWORK_PATH),
        ]:
            if not d:
                continue
            remaining = _count_remaining_xlsx(d)
            if remaining < 0:
                print(f"[清空] 网络目录不存在: {d}")
            elif remaining > 0:
                msg = f"网络目录 {label} 清空后仍有 {remaining} 个文件: {d}"
                errors.append(msg)
                print(f"[清空] {msg}")
            else:
                print(f"[清空] 网络目录 {label} 清空完成，无残留")

    except Exception as e:
        error_msg = f"网络输出清空失败: {e}"
        errors.append(error_msg)
        print(f"[清空] {error_msg}")
        logging.error(f"[清空] {error_msg}", exc_info=True)
    finally:
        try:
            bat_file.unlink()
        except OSError:
            pass

    return errors


@app.route("/api/clear-output", methods=["POST"])
def clear_output():
    body = request.get_json(silent=True) or {}
    network_path = body.get("network_output", _current_network_output)
    if network_path:
        network_path = network_path.strip()

    deleted = 0
    errors: list[str] = []

    # 清空本地 output
    if OUTPUT_DIR.exists():
        for f in OUTPUT_DIR.iterdir():
            if f.is_file() and f.suffix == ".xlsx":
                try:
                    f.unlink()
                    deleted += 1
                except OSError as e:
                    error_msg = f"无法删除本地文件 {f.name}: {e}"
                    errors.append(error_msg)
                    print(f"[清空] {error_msg}")
                    logging.error("[清空] %s", error_msg)

    # 清空进度文件
    if PROGRESS_FILE.exists():
        try:
            PROGRESS_FILE.unlink()
            print("[清空] 进度文件已重置")
        except OSError as e:
            print(f"[清空] 无法删除进度文件: {e}")

    # 清空网络输出
    errors.extend(_clear_network(network_path if network_path else None))

    return jsonify({"success": len(errors) == 0, "deleted": deleted, "errors": errors})


@app.route("/api/open-output", methods=["POST"])
def open_output():
    target = Path(_current_network_output) if _current_network_output else OUTPUT_DIR
    if not target.exists():
        target = OUTPUT_DIR
    if not target.exists():
        target.mkdir(parents=True, exist_ok=True)
    try:
        # 使用 start 命令替代 explorer，更兼容服务器环境
        subprocess.Popen(["cmd.exe", "/c", "start", str(target)], shell=True)
        return jsonify({"success": True, "path": str(target)})
    except Exception as exc:
        # 如果无法打开，至少返回路径供用户手动访问
        return jsonify({"success": False, "error": str(exc), "path": str(target)})


@app.route("/api/stop", methods=["POST"])
def stop():
    global _current_process
    with _process_lock:
        if _current_process is not None:
            _current_process.terminate()
            _current_process = None
    _finish_progress(error="用户手动停止")
    return jsonify({"success": True})


@app.route("/api/open-log", methods=["POST"])
def open_log():
    if not LOG_DIR.exists():
        LOG_DIR.mkdir(parents=True, exist_ok=True)
    try:
        # 使用 start 命令替代 explorer，更兼容服务器环境
        subprocess.Popen(["cmd.exe", "/c", "start", str(LOG_DIR)], shell=True)
        return jsonify({"success": True, "path": str(LOG_DIR)})
    except Exception as exc:
        # 如果无法打开，至少返回路径供用户手动访问
        return jsonify({"success": False, "error": str(exc), "path": str(LOG_DIR)})


@app.route("/api/clear-log", methods=["POST"])
def clear_log():
    if not LOG_DIR.exists():
        return jsonify({"success": True, "deleted": 0})
    deleted = 0
    for f in LOG_DIR.iterdir():
        if f.is_file() and f.suffix == ".log":
            try:
                f.unlink()
                deleted += 1
            except OSError:
                pass
    return jsonify({"success": True, "deleted": deleted})


@app.route("/api/reset", methods=["POST"])
def reset():
    """统一重置：清空本地 output、网络输出、日志、进度。"""
    body = request.get_json(silent=True) or {}
    network_path = body.get("network_output", _current_network_output)
    if network_path:
        network_path = network_path.strip()

    local_deleted = 0
    log_deleted = 0
    errors: list[str] = []

    # 清空本地 output
    if OUTPUT_DIR.exists():
        for f in OUTPUT_DIR.iterdir():
            if f.is_file() and f.suffix == ".xlsx":
                try:
                    f.unlink()
                    local_deleted += 1
                except OSError as e:
                    errors.append(f"无法删除本地文件 {f.name}: {e}")
                    logging.error("[重置] 无法删除本地文件 %s: %s", f.name, e)

    # 清空日志
    if LOG_DIR.exists():
        for f in LOG_DIR.iterdir():
            if f.is_file() and f.suffix == ".log":
                try:
                    f.unlink()
                    log_deleted += 1
                except OSError:
                    pass

    # 清空进度文件
    if PROGRESS_FILE.exists():
        try:
            PROGRESS_FILE.unlink()
        except OSError as e:
            errors.append(f"无法删除进度文件: {e}")
            logging.error("[重置] 无法删除进度文件: %s", e)

    # 清空网络输出
    errors.extend(_clear_network(network_path if network_path else None))

    print(f"[重置] 本地文件删除 {local_deleted}, 日志删除 {log_deleted}, 错误 {len(errors)}")
    return jsonify({
        "success": len(errors) == 0,
        "local_deleted": local_deleted,
        "log_deleted": log_deleted,
        "errors": errors,
    })


if __name__ == "__main__":
    # 配置文件日志
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    debug_log = LOG_DIR / "server_debug.log"
    logging.basicConfig(
        filename=str(debug_log),
        level=logging.DEBUG,
        format="%(asctime)s %(levelname)s %(message)s",
        encoding="utf-8",
    )
    logging.info("BOM_API 服务器启动")
    print(f"BOM_API 数据同步工具 Web 界面")
    print(f"访问 http://localhost:5000")
    print(f"调试日志: {debug_log}")
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
