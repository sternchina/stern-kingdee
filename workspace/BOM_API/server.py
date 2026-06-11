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
DEFAULT_NETWORK_OUTPUT = r"\\192.168.1.181\99_Share\04_IT\KingdeeExport"
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
    {"id": "PRD_PickMtrl", "name": "生产领料单", "desc": "生产领料"},
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


def _sync_to_network(network_path: str) -> dict | None:
    """将本地 output 目录中的 xlsx 文件同步到网络共享目录。"""
    if not OUTPUT_DIR.exists():
        print(f"[同步] 本地 output 目录不存在，跳过同步")
        return None
    if not network_path:
        print(f"[同步] 未指定网络输出路径，跳过同步")
        return None

    xlsx_files = list(OUTPUT_DIR.glob("*.xlsx"))
    print(f"[同步] 本地 output 共 {len(xlsx_files)} 个 xlsx 文件，目标: {network_path}")

    if not xlsx_files:
        print(f"[同步] 无 xlsx 文件需要复制")
        return {"copied": 0, "total": 0, "log": "无文件需复制"}

    all_log_lines = []

    # 网络认证参数
    parts = Path(network_path).parts
    share_root = parts[0].rstrip("\\") if parts else network_path
    net_user = r"STERN\shenjj"
    net_pass = r"zH820303."
    # 计算共享根目录下的子路径
    sub_path = str(Path(network_path)).replace(share_root, "").strip("\\")

    # 生成临时批处理脚本：映射驱动器 → 复制 → 断开驱动器
    import tempfile
    bat_lines = [
        # 先断开所有与该服务器的连接（解决错误 1219）
        f'net use * /delete /yes 2>nul',
        f'timeout /t 1 /nobreak >nul',
        f'net use Z: {share_root} {net_pass} /user:{net_user} /persistent:no',
    ]
    for f in xlsx_files:
        src = str(f).replace("/", "\\")
        if sub_path:
            dst = f"Z:\\{sub_path}\\{f.name}"
        else:
            dst = f"Z:\\{f.name}"
        bat_lines.append(f'copy /Y "{src}" "{dst}"')
    bat_lines.append(f'net use Z: /delete /yes')
    bat_lines.append('exit /b 0')

    bat_content = "\r\n".join(bat_lines)
    bat_file = Path(tempfile.gettempdir()) / "bom_sync.bat"
    bat_file.write_text(bat_content, encoding="gbk")

    print(f"[同步] 执行批处理同步")
    try:
        result = subprocess.run(
            ["cmd", "/c", str(bat_file)],
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
    target_dir = Path(network_path)
    copied = 0
    for f in xlsx_files:
        dest_file = target_dir / f.name
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
    _current_network_output = network_output or DEFAULT_NETWORK_OUTPUT

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
    network_path = _current_network_output
    display_dir = network_path
    scan_dir = Path(network_path) if Path(network_path).exists() else OUTPUT_DIR

    files = []
    if scan_dir.exists():
        for f in sorted(scan_dir.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
            files.append({
                "name": f.name,
                "path": str(Path(network_path) / f.name),
                "size": f.stat().st_size,
                "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat(),
            })

    return jsonify({"output_dir": display_dir, "files": files})


@app.route("/api/clear-output", methods=["POST"])
def clear_output():
    body = request.get_json(silent=True) or {}
    network_path = body.get("network_output", _current_network_output)
    if network_path:
        network_path = network_path.strip()

    deleted = 0
    errors = []

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

    # 清空网络输出（批处理：映射驱动器 + del + 断开驱动器）
    if network_path:
        try:
            import tempfile
            parts = Path(network_path).parts
            share_root = parts[0].rstrip("\\") if parts else network_path
            net_user = r"STERN\shenjj"
            net_pass = r"zH820303."
            # 计算共享根目录下的子路径
            sub_path = str(Path(network_path)).replace(share_root, "").strip("\\")
            pattern = sub_path + "\\*.xlsx" if sub_path else "*.xlsx"
            bat_lines = [
                # 先断开所有与该服务器的连接（解决错误 1219）
                f'net use * /delete /yes 2>nul',
                f'timeout /t 1 /nobreak >nul',
                f'net use Z: {share_root} {net_pass} /user:{net_user} /persistent:no',
                f'del /Q /F "Z:\\{pattern}"',
                f'net use Z: /delete /yes',
                'exit /b 0',
            ]
            bat_content = "\r\n".join(bat_lines)
            bat_file = Path(tempfile.gettempdir()) / "bom_clear.bat"
            bat_file.write_text(bat_content, encoding="gbk")
            try:
                result = subprocess.run(
                    ["cmd", "/c", str(bat_file)],
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
                # 统计删除数量
                net_dir = Path(network_path)
                if net_dir.exists():
                    remaining = len(list(net_dir.glob("*.xlsx")))
                    print(f"[清空] 网络输出清空完成，剩余 {remaining} 个文件")
                    if remaining > 0:
                        errors.append(f"网络目录清空后仍有 {remaining} 个文件")
                else:
                    errors.append(f"网络目录不存在或无法访问: {network_path}")
                    print(f"[清空] 网络目录不存在: {network_path}")
            finally:
                try:
                    bat_file.unlink()
                except OSError:
                    pass
        except Exception as e:
            error_msg = f"网络输出清空失败: {e}"
            errors.append(error_msg)
            print(f"[清空] {error_msg}")
            logging.error(f"[清空] {error_msg}", exc_info=True)

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
        subprocess.Popen(["cmd", "/c", "start", str(target)], shell=True)
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
        subprocess.Popen(["cmd", "/c", "start", str(LOG_DIR)], shell=True)
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
