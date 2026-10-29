"""
Kingdee Master Data 数据同步与导入 Excel 生成工具

用法（在项目根目录运行）：
    python workspace/BOM_API/sync_and_generate.py
    python workspace/BOM_API/sync_and_generate.py --skip-fetch   # 跳过拉取，用本地 CSV
    python workspace/BOM_API/sync_and_generate.py --form PRD_INSTOCK --lang EN   # 导出指定表单的英文版
    python workspace/BOM_API/sync_and_generate.py --netaddr ""   # 不同步到网络

参数说明：
    --form: 指定导出的表单类型，可选值见下方"支持的表单类型"，默认 all
    --lang: 指定输出语言，EN 为英文版（默认），CN 为中文版，ALL 为中英文都输出
    --skip-fetch: 跳过从金蝶拉取数据，使用本地已有的 CSV 文件
    --netaddr: 网络输出根路径（SMB共享），默认 \\\\192.168.1.181\\12_SupplyChain\\07_MPS\\MPS_Upgrade，留空则不同步

流程：
    1. 从金蝶拉取 BD_MATERIAL、ENG_BOM、PLN_FORECAST、PLN_PLANORDER、PRD_MO、PUR_PurchaseOrder、PRD_INSTOCK、STK_Inventory、PRD_PickMtrl 数据，保存到 data/ 目录
    2. 根据 template/ 下的金蝶模板生成导入 Excel
    3. 输出到 output/ 目录

输出文件：
    - 物料列表.xlsx / K3_MaterialMaster.xlsx
    - 物料清单列表.xlsx / K3_BOM.xlsx
    - 预测单列表.xlsx / K3_Forecast.xlsx
    - 计划订单列表.xlsx / K3_PlannedOrders.xlsx
    - 生产订单列表.xlsx / K3_ManufacturingOrders.xlsx
    - 采购订单列表.xlsx / K3_PurchaseOrders.xlsx
    - 生产入库单列表.xlsx / K3_Production_Receipts.xlsx
    - 即时库存.xlsx / K3_Inventory.xlsx
    - 生产领料单列表.xlsx / K3_18015_PickingList.xlsx

支持的表单类型（通过 --form 参数指定）：
    - all: 导出所有表单
    - BD_MATERIAL: 物料基础资料
    - ENG_BOM: 物料清单
    - PLN_FORECAST: 预测单
    - PLN_PLANORDER: 计划订单
    - PRD_MO: 生产订单
    - PUR_PurchaseOrder: 采购订单
    - PRD_INSTOCK: 生产入库单
    - STK_Inventory: 即时库存
    - PRD_PickMtrl: 生产领料单

"""
from __future__ import annotations

import argparse
import logging
import sys

# 避免编码问题导致输出被抑制
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd

# ── 路径 ───────────────────────────────────────────────────────────────────────
WORKSPACE    = Path(__file__).resolve().parent          # workspace/BOM_API
PROJECT_ROOT = WORKSPACE.parent.parent                  # 项目根目录
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.config import get_settings
from src.kingdee_client import KingdeeClient


TEMPLATE_DIR = WORKSPACE / "template"
DATA_DIR     = WORKSPACE / "data"
OUTPUT_DIR   = WORKSPACE / "output"
MAP_DIR      = WORKSPACE / "map"
LOG_DIR      = WORKSPACE / "log"

# ── 日志配置 ──────────────────────────────────────────────────────────────────
LOG_DIR.mkdir(exist_ok=True)
_log_file = LOG_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logger = logging.getLogger("sync_and_generate")
logger.setLevel(logging.DEBUG)

# 文件handler：记录所有级别
_fh = logging.FileHandler(_log_file, encoding="utf-8")
_fh.setLevel(logging.DEBUG)
_fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(_fh)

# 控制台handler：只记录WARNING以上
_ch = logging.StreamHandler(sys.stderr)
_ch.setLevel(logging.WARNING)
_ch.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
logger.addHandler(_ch)

ORG_ID   = "100"
ORG_NAME = "斯顿食品原料（苏州）有限公司"
TODAY    = date.today().isoformat()
NOW      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
RUN_TS   = datetime.now().strftime("%Y%m%d_%H%M%S")
EXPIRE   = "9999-12-31"

FORM_ID = "ENG_BOM"
MATERIAL_FORM_ID = "BD_MATERIAL"
PLN_FORECAST_FORM_ID = "PLN_FORECAST"
PLANORDER_FORM_ID = "PLN_PLANORDER"
MO_FORM_ID = "PRD_MO"
PO_FORM_ID = "PUR_PurchaseOrder"
INSTOCK_FORM_ID = "PRD_INSTOCK"
INV_FORM_ID = "STK_Inventory"
PICKMTRL_FORM_ID = "PRD_PickMtrl"

BILL_TYPE_MAPPING = {
    "683e2791298b4e79974ef3d2adc187ff": "WLQD01_SYS",
}

BILL_TYPE_NAME_MAPPING = {
    "WLQD01_SYS": "物料清单",
}

FUNITID_MAPPING = {
    "10095": "kg",
    "109968": "ppm",
}

FGROUP_MAPPING = {
    "116995": "01",
    "117003": "0206",
    "116998": "0201",
    "116999": "0202",
    "117000": "0203",
    "116997": "03",
    "117002": "0205",
    "117001": "0204",
    "116996": "02",
    "1": "01",
    "2": "02",
    "3": "03",
    "201": "0201",
    "202": "0202",
    "203": "0203",
    "204": "0204",
    "205": "0205",
    "206": "0206",
}

FGROUP_NAME_MAPPING = {
    "01": "生产BOM",
    "02": "研发BOM",
    "03": "SC注册BOM",
    "0201": "MC/DB",
    "0202": "DB",
    "0203": "SE",
    "0204": "HY",
    "0205": "PL",
    "0206": "SV",
}

FBOMCATEGORY_MAPPING = {
    "1": "标准BOM",
}

FBOMUSE_MAPPING = {
    "99": "通用",
}

FORGANIZATION_MAPPING = {
    "1": "斯顿食品原料（苏州）有限公司",
}

FAUXPROPID_MAPPING = {
    "0": "",
}

FMDLID_MAPPING = {
    "0": "",
}


def _resolve_local_csv(form_id: str) -> Path | None:
    """本地缓存：优先取最新的 ``{form_id}_*.csv``（按修改时间），否则回退 ``{form_id}.csv``。"""
    stamped = list(DATA_DIR.glob(f"{form_id}_*.csv"))
    if stamped:
        return max(stamped, key=lambda p: p.stat().st_mtime)
    legacy = DATA_DIR / f"{form_id}.csv"
    return legacy if legacy.exists() else None


def _convert_id_to_code(df: pd.DataFrame, client: KingdeeClient = None, material_df: pd.DataFrame = None) -> None:
    """将 DataFrame 中的各种 ID 字段转换为编码。"""
    mappings = [
        ("FBILLTYPE", BILL_TYPE_MAPPING, "BOS_BILLTYPE"),
        ("FUNITID", FUNITID_MAPPING, "BD_UNIT"),
        ("FGroup", FGROUP_MAPPING, "ENG_BOMGROUP"),
        ("FBOMCATEGORY", FBOMCATEGORY_MAPPING, "ENG_BOMCATEGORY"),
        ("FBOMUSE", FBOMUSE_MAPPING, "ENG_BOMUSE"),
    ]
    
    for field_name, mapping, form_id in mappings:
        if field_name not in df.columns:
            continue
        
        if client:
            unique_ids = df[field_name].dropna().unique()
            for id_val in unique_ids:
                id_str = str(int(id_val)) if isinstance(id_val, (int, float)) and id_val == int(id_val) else str(id_val).strip()
                if not id_str or id_str in mapping:
                    continue
                
                try:
                    result = client.view(form_id, {'Id': id_str, 'Number': '', 'CreateOrgId': 0})
                    code = result.get('Result', {}).get('Result', {}).get('Number', '')
                    if code:
                        mapping[id_str] = code
                        print(f"  [API查询] {field_name}: {id_str} -> {code}")
                except Exception as e:
                    logger.error("查询%s编码映射失败: %s", field_name, e, exc_info=True)
                    pass
        
        df[field_name] = df[field_name].apply(
            lambda x: mapping.get(str(int(x)) if isinstance(x, (int, float)) and x == int(x) else str(x).strip(), str(x).strip())
        ).astype(str)
        print(f"  [转换] {field_name} ID 已转换为编码")
    
    if "FGroup" in df.columns and FGROUP_NAME_MAPPING:
        df["FGroup"] = df["FGroup"].apply(
            lambda x: FGROUP_NAME_MAPPING.get(str(x).strip(), str(x).strip())
        ).astype(str)
        print(f"  [转换] FGroup 编码已转换为名称")
    
    if "FMATERIALID" in df.columns:
        print("  [物料] 正在转换物料编码...")
        material_mapping = {}
        
        use_api_query = False
        if material_df is not None and not material_df.empty:
            print("    [来源] 尝试使用本地 BD_MATERIAL 数据")
            # 检查是否有FId字段（物料内码）
            has_fid = "FId" in material_df.columns or "Id" in material_df.columns
            has_fnumber = "FNumber" in material_df.columns or "Number" in material_df.columns
            
            if has_fid and has_fnumber:
                id_col = "FId" if "FId" in material_df.columns else "Id"
                for _, row in material_df.iterrows():
                    f_id = str(int(row[id_col])) if isinstance(row[id_col], (int, float)) and row[id_col] == int(row[id_col]) else str(row[id_col]).strip()
                    f_number = str(row["FNumber"]).strip() if pd.notna(row["FNumber"]) else ""
                    if f_id and f_number:
                        material_mapping[f_id] = f_number
                print(f"      找到 {len(material_mapping)} 个物料映射")
            else:
                print("      BD_MATERIAL 数据中无物料内码字段，使用API查询")
                use_api_query = True
        else:
            use_api_query = True
        
        if use_api_query and client:
            print("    [来源] 使用API查询物料编码")
            material_ids = df["FMATERIALID"].dropna().unique()
            for id_val in material_ids:
                id_str = str(int(id_val)) if isinstance(id_val, (int, float)) and id_val == int(id_val) else str(id_val).strip()
                if not id_str or id_str in material_mapping:
                    continue
                
                try:
                    result = client.view("BD_MATERIAL", {"Id": id_str, "Number": "", "CreateOrgId": 0})
                    code = result.get("Result", {}).get("Result", {}).get("Number", "")
                    if code:
                        material_mapping[id_str] = code
                except Exception as e:
                    logger.error("查询物料编码映射失败: %s", e, exc_info=True)
                    pass
            df["FMATERIALID"] = df["FMATERIALID"].apply(
                lambda x: material_mapping.get(str(int(x)) if isinstance(x, (int, float)) and x == int(x) else str(x).strip(), str(x).strip())
            ).astype(str)
            print(f"  [转换] FMATERIALID 已转换为物料编码，共 {len(material_mapping)} 个")
        else:
            print("  [警告] 物料编码转换失败")


def _convert_id_to_code_with_mapping(df: pd.DataFrame, client: KingdeeClient = None, material_code_mapping: dict[str, str] = None, material_name_mapping: dict[str, str] = None) -> None:
    """将 DataFrame 中的各种 ID 字段转换为编码（使用预构建的物料映射字典）。"""
    mappings = [
        ("FBILLTYPE", BILL_TYPE_MAPPING, "BOS_BILLTYPE"),
        ("FUNITID", FUNITID_MAPPING, "BD_UNIT"),
        ("FGroup", FGROUP_MAPPING, "ENG_BOMGROUP"),
        ("FBOMCATEGORY", FBOMCATEGORY_MAPPING, "ENG_BOMCATEGORY"),
        ("FBOMUSE", FBOMUSE_MAPPING, "ENG_BOMUSE"),
    ]
    
    for field_name, mapping, form_id in mappings:
        if field_name not in df.columns:
            continue
        
        if client:
            unique_ids = df[field_name].dropna().unique()
            for id_val in unique_ids:
                id_str = str(int(id_val)) if isinstance(id_val, (int, float)) and id_val == int(id_val) else str(id_val).strip()
                if not id_str or id_str in mapping:
                    continue
                
                try:
                    result = client.view(form_id, {'Id': id_str, 'Number': '', 'CreateOrgId': 0})
                    code = result.get('Result', {}).get('Result', {}).get('Number', '')
                    if code:
                        mapping[id_str] = code
                        print(f"  [API查询] {field_name}: {id_str} -> {code}")
                except Exception as e:
                    logger.error("查询%s编码映射失败: %s", field_name, e, exc_info=True)
                    pass
        
        df[field_name] = df[field_name].apply(
            lambda x: mapping.get(str(int(x)) if isinstance(x, (int, float)) and x == int(x) else str(x).strip(), str(x).strip())
        ).astype(str)
        print(f"  [转换] {field_name} ID 已转换为编码")
    
    if "FGroup" in df.columns and FGROUP_NAME_MAPPING:
        df["FGroup"] = df["FGroup"].apply(
            lambda x: FGROUP_NAME_MAPPING.get(str(x).strip(), str(x).strip())
        ).astype(str)
        print(f"  [转换] FGroup 编码已转换为名称")
    
    if "FMATERIALID" in df.columns:
        print("  [物料] 正在转换物料编码...")
        
        if material_code_mapping and len(material_code_mapping) > 0:
            print(f"    [来源] 使用预构建的物料映射，共 {len(material_code_mapping)} 个")
            
            # 先保存原始ID，用于获取物料名称
            df["FMATERIALID_ORIGINAL"] = df["FMATERIALID"].apply(
                lambda x: str(int(x)) if isinstance(x, (int, float)) and x == int(x) else str(x).strip()
            )
            
            # 转换物料编码
            df["FMATERIALID"] = df["FMATERIALID_ORIGINAL"].apply(
                lambda x: material_code_mapping.get(x, x)
            ).astype(str)
            print(f"  [转换] FMATERIALID 已转换为物料编码")
            
            # 如果有物料名称映射，添加物料名称列（使用原始ID查找）
            if material_name_mapping and len(material_name_mapping) > 0:
                df["FMATERIALNAME"] = df["FMATERIALID_ORIGINAL"].apply(
                    lambda x: material_name_mapping.get(x, "")
                ).astype(str)
                print(f"  [转换] 添加物料名称列 FMATERIALNAME")
            
            # 删除临时列
            df.drop("FMATERIALID_ORIGINAL", axis=1, inplace=True)
        elif client:
            print("    [来源] 使用API查询物料编码")
            api_material_mapping = {}
            api_material_name_mapping = {}
            material_ids = df["FMATERIALID"].dropna().unique()
            for id_val in material_ids:
                id_str = str(int(id_val)) if isinstance(id_val, (int, float)) and id_val == int(id_val) else str(id_val).strip()
                if not id_str or id_str in api_material_mapping:
                    continue
                
                try:
                    result = client.view("BD_MATERIAL", {"Id": id_str, "Number": "", "CreateOrgId": 0})
                    inner_result = result.get("Result", {}).get("Result", {})
                    if inner_result:
                        code = inner_result.get("Number", "")
                        name = inner_result.get("Name", "")
                        
                        # 处理名称（可能是列表格式）
                        if isinstance(name, list) and len(name) > 0:
                            name = name[0].get("Value", "") if isinstance(name[0], dict) else str(name[0])
                        
                        if code:
                            api_material_mapping[id_str] = code
                            if name:
                                api_material_name_mapping[id_str] = name
                except Exception as e:
                    logger.error("API查询物料编码和名称映射失败: %s", e, exc_info=True)
                    pass
                df["FMATERIALID"] = df["FMATERIALID"].apply(
                    lambda x: api_material_mapping.get(str(int(x)) if isinstance(x, (int, float)) and x == int(x) else str(x).strip(), str(x).strip())
                ).astype(str)
                print(f"  [转换] FMATERIALID 已转换为物料编码，共 {len(api_material_mapping)} 个")
                
                if api_material_name_mapping:
                    df["FMATERIALNAME"] = df["FMATERIALID"].apply(
                        lambda x: api_material_name_mapping.get(str(int(x)) if isinstance(x, (int, float)) and x == int(x) else str(x).strip(), "")
                    ).astype(str)
                    print(f"  [转换] 添加物料名称列 FMATERIALNAME")
            else:
                print("  [警告] 物料编码转换失败")
        else:
            print("  [警告] 无物料映射和客户端，无法转换物料编码")


# ══════════════════════════════════════════════════════════════════════════════
# 1. 从金蝶拉取数据
# ══════════════════════════════════════════════════════════════════════════════

def _main_entry(entries: list[dict]) -> dict | None:
    """取主表实体（含 FBillHead 的 entry），没有则取第一个。"""
    for e in entries:
        if "FBillHead" in e.get("entry_key", ""):
            return e
    return entries[0] if entries else None


def _first_entry(entries: list[dict]) -> dict | None:
    """取第一个 entry（与测试脚本一致）。"""
    return entries[0] if entries else None


def _print_progress(current: int, total: int, prefix: str = "", suffix: str = "") -> None:
    """显示进度条。"""
    if total == 0:
        return
    
    percent = (current / total) * 100
    bar_length = 30
    filled_length = int(bar_length * current // total)
    
    bar = "█" * filled_length + "░" * (bar_length - filled_length)
    print(f"\r{prefix} [{bar}] {percent:.1f}% ({current}/{total}) {suffix}", end="", flush=True)


def _clean_string(s):
    """清理字符串中的非法字符，使其可以用于Excel工作表。datetime 对象直接返回。"""
    if isinstance(s, datetime):
        return s
    if not s:
        return ""
    
    # 确保是字符串类型
    s = str(s)
    
    # openpyxl不允许的字符范围（控制字符）
    # 允许的控制字符: \t (0x09), \n (0x0A), \r (0x0D)
    result = []
    for char in s:
        char_code = ord(char)
        # 移除控制字符（除了制表符、换行、回车）
        if (char_code < 32 and char not in '\t\n\r') or char_code == 127:
            continue
        result.append(char)
    
    return ''.join(result)


def _save_workbook(workbook, output_path) -> None:
    """保存工作簿，文件被占用时跳过并提示。"""
    try:
        workbook.save(output_path)
    except PermissionError:
        logger.warning("文件被占用，保存失败: %s", output_path.name)
        print(f"  [警告] 文件被占用，保存失败: {output_path.name}（请关闭 Excel 后重试）")


def _format_date(s: str):
    """将金蝶日期格式 '2023-12-01T00:00:00' 转换为 datetime 对象，Excel 按单元格格式自动显示。"""
    if not s:
        return s
    try:
        if "T" in s:
            # 金蝶格式: 2023-12-01T00:00:00 或 2023-12-01T00:00:00.123
            date_part, time_part = s.split("T", 1)
            time_part = time_part.split(".")[0]
            return datetime.strptime(f"{date_part} {time_part}", "%Y-%m-%d %H:%M:%S")
        # 尝试其他常见格式
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        return s
    except Exception:
        return s


def _apply_date_formats(ws, header_row: list, date_col_names: set, field_mapping: dict):
    """为日期列设置 Excel 单元格格式。根据模板列名判断是否含时间部分。"""
    for tpl_idx, data_col in field_mapping.items():
        if data_col not in date_col_names:
            continue
        header = header_row[tpl_idx] if tpl_idx < len(header_row) else ""
        # 含"时间"/"Time"的列用带时间的格式，其他用纯日期格式
        if any(kw in str(header) for kw in ("时间", "Time", "开始日期", "结束日期", "Start Date", "End Date",
                                              "开工", "完工", "建议", "确认", "需求")):
            fmt = "yyyy/M/d HH:mm:ss"
        else:
            fmt = "yyyy/M/d"
        # 从第2行开始（第1行是表头），设置该列所有数据行的格式
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=tpl_idx + 1)
            if isinstance(cell.value, datetime):
                cell.number_format = fmt


def _load_material_mapping() -> tuple[dict[str, str], dict[str, str]]:
    """从map文件夹加载物料映射。"""
    MAP_DIR.mkdir(exist_ok=True)
    code_map_path = MAP_DIR / "material_code_map.json"
    name_map_path = MAP_DIR / "material_name_map.json"
    
    material_code_mapping = {}
    material_name_mapping = {}
    
    if code_map_path.exists():
        try:
            import json
            with open(code_map_path, 'r', encoding='utf-8') as f:
                material_code_mapping = json.load(f)
            print(f"    [缓存] 已加载 {len(material_code_mapping)} 个物料编码映射")
        except Exception as e:
            logger.error("加载物料编码映射失败: %s", e, exc_info=True)
            print(f"    [警告] 加载物料编码映射失败: {e}")
    
    if name_map_path.exists():
        try:
            import json
            with open(name_map_path, 'r', encoding='utf-8') as f:
                material_name_mapping = json.load(f)
            print(f"    [缓存] 已加载 {len(material_name_mapping)} 个物料名称映射")
        except Exception as e:
            logger.error("加载物料名称映射失败: %s", e, exc_info=True)
            print(f"    [警告] 加载物料名称映射失败: {e}")
    
    return material_code_mapping, material_name_mapping


def _save_material_mapping(material_code_mapping: dict[str, str], material_name_mapping: dict[str, str]) -> None:
    """保存物料映射到map文件夹。"""
    MAP_DIR.mkdir(exist_ok=True)
    code_map_path = MAP_DIR / "material_code_map.json"
    name_map_path = MAP_DIR / "material_name_map.json"
    
    try:
        import json
        with open(code_map_path, 'w', encoding='utf-8') as f:
            json.dump(material_code_mapping, f, ensure_ascii=False, indent=2)
        print(f"    [缓存] 已保存 {len(material_code_mapping)} 个物料编码映射")
    except Exception as e:
        logger.error("保存物料编码映射失败: %s", e, exc_info=True)
        print(f"    [警告] 保存物料编码映射失败: {e}")
    
    try:
        import json
        with open(name_map_path, 'w', encoding='utf-8') as f:
            json.dump(material_name_mapping, f, ensure_ascii=False, indent=2)
        print(f"    [缓存] 已保存 {len(material_name_mapping)} 个物料名称映射")
    except Exception as e:
        logger.error("保存物料名称映射失败: %s", e, exc_info=True)
        print(f"    [警告] 保存物料名称映射失败: {e}")


def _load_generic_mapping(map_name: str) -> dict[str, str]:
    """从map文件夹加载通用映射。"""
    MAP_DIR.mkdir(exist_ok=True)
    map_path = MAP_DIR / f"{map_name}.json"
    mapping = {}
    if map_path.exists():
        try:
            import json
            with open(map_path, 'r', encoding='utf-8') as f:
                mapping = json.load(f)
            print(f"    [缓存] 已加载 {len(mapping)} 个 {map_name} 映射")
        except Exception as e:
            logger.error("加载 %s 映射失败: %s", map_name, e, exc_info=True)
            print(f"    [警告] 加载 {map_name} 映射失败: {e}")
    return mapping


def _save_generic_mapping(map_name: str, mapping: dict[str, str]) -> None:
    """保存通用映射到map文件夹。"""
    MAP_DIR.mkdir(exist_ok=True)
    map_path = MAP_DIR / f"{map_name}.json"
    try:
        import json
        with open(map_path, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
        print(f"    [缓存] 已保存 {len(mapping)} 个 {map_name} 映射")
    except Exception as e:
        logger.error("保存 %s 映射失败: %s", map_name, e, exc_info=True)
        print(f"    [警告] 保存 {map_name} 映射失败: {e}")


def _is_guid(value: str) -> bool:
    if not value:
        return False
    s = value.strip().replace("-", "")
    return len(s) == 32 and all(c in "0123456789abcdefABCDEF" for c in s)


def fetch_generic_mapping(client: KingdeeClient, form_id: str, id_field: str, number_field: str, map_name: str, required_ids: list[str] = None) -> dict[str, str]:
    mapping = _load_generic_mapping(map_name)
    
    if required_ids:
        missing_ids = [mid for mid in required_ids if mid not in mapping]
        if not missing_ids:
            print(f"    [缓存] 所有 {len(required_ids)} 个ID都已在 {map_name} 缓存中")
            return mapping
        print(f"    [缓存] 缓存中有 {len(mapping)} 个映射，还需要查询 {len(missing_ids)} 个")
    else:
        missing_ids = None
    
    initial_count = len(mapping)
    
    has_guid_ids = required_ids and any(_is_guid(mid) for mid in required_ids)
    
    if has_guid_ids and missing_ids:
        for mid in missing_ids:
            if mid in mapping:
                continue
            try:
                filter_str = f"{id_field} like '%{mid}%'"
                rows = client.execute_bill_query({
                    "FormId": form_id,
                    "FieldKeys": f"{id_field},{number_field}",
                    "FilterString": filter_str,
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": 10,
                })
                if rows:
                    for row in rows:
                        if len(row) >= 2 and row[0] and row[1]:
                            f_id = str(row[0]).strip()
                            f_number = str(row[1]).strip()
                            if f_id and f_number:
                                mapping[f_id] = f_number
            except Exception as e:
                logger.error("查询 %s ID=%s 失败: %s", map_name, mid, e, exc_info=True)
                print(f"    [警告] 查询 {map_name} ID={mid} 失败: {e}")
    else:
        field_keys_csv = f"{id_field},{number_field}"
        
        start_row, limit = 0, 2000
        while True:
            rows = client.execute_bill_query({
                "FormId": form_id,
                "FieldKeys": field_keys_csv,
                "FilterString": "",
                "OrderString": "",
                "TopRowCount": 0,
                "StartRow": start_row,
                "Limit": limit,
            })
            
            if not rows:
                break
            
            for row in rows:
                if len(row) >= 2 and row[0] and str(row[0]).strip() != "0":
                    raw_id = str(row[0]).strip()
                    if _is_guid(raw_id):
                        f_id = raw_id
                    else:
                        try:
                            f_id = str(int(float(raw_id)))
                        except (ValueError, TypeError, OverflowError):
                            f_id = raw_id
                    
                    if missing_ids and f_id in mapping:
                        continue
                    
                    f_number = str(row[1]).strip() if row[1] else ""
                    if f_id and f_number:
                        mapping[f_id] = f_number
            
            if len(rows) < limit:
                break
            start_row += limit
    
    new_count = len(mapping) - initial_count
    if new_count > 0:
        _save_generic_mapping(map_name, mapping)
        print(f"    [映射] {map_name}: 本次新增 {new_count} 个映射")
    else:
        print(f"    [映射] {map_name}: 无新增映射")
    
    return mapping


def fetch_material_mapping(client: KingdeeClient, required_material_ids: list[str] = None) -> tuple[dict[str, str], dict[str, str]]:
    """
    获取物料内码到编码和名称的映射
    先从缓存加载，查不到再通过API查询，查询结果自动保存到缓存供后续使用
    
    Args:
        client: KingdeeClient实例
        required_material_ids: 需要查询的物料ID列表（可选），如果提供则只查询这些ID
    
    Returns:
        tuple: (material_code_mapping, material_name_mapping)
            material_code_mapping: 物料内码到编码的映射
            material_name_mapping: 物料内码到名称的映射
    """
    print("  [物料] 正在获取物料ID映射...")
    
    # 步骤0: 先从缓存加载映射
    material_code_mapping, material_name_mapping = _load_material_mapping()
    
    # 如果有指定需要的物料ID，检查缓存是否已经包含所有需要的ID
    if required_material_ids:
        missing_ids = [mid for mid in required_material_ids if mid not in material_code_mapping]
        if not missing_ids:
            print(f"    [缓存] 所有 {len(required_material_ids)} 个物料ID都已在缓存中")
            return material_code_mapping, material_name_mapping
        print(f"    [缓存] 缓存中有 {len(material_code_mapping)} 个映射，还需要查询 {len(missing_ids)} 个")
    else:
        missing_ids = None
    
    initial_count = len(material_code_mapping)
    
    # 步骤1: 先尝试使用ExecuteBillQuery批量获取
    try:
        field_info = client.get_form_field_keys("BD_MATERIAL")
        entry = _main_entry(field_info.get("entries", []))
        field_keys = entry["field_keys"] if entry else field_info["field_keys"]
        
        # 查找FId、FNumber、FName字段
        id_field = None
        number_field = None
        name_field = None
        for key in field_keys:
            if key.endswith(".FId") or key.endswith(".FID"):
                id_field = key
            elif key.endswith(".FNumber") or key.endswith(".FNUMBER"):
                number_field = key
            elif key.endswith(".FName") or key.endswith(".FNAME"):
                name_field = key
        
        # 如果找不到完整字段名，尝试使用别名
        if not id_field:
            id_field = "FId"
        if not number_field:
            number_field = "FNumber"
        if not name_field:
            name_field = "FName"
        
        # 构建查询字段（包含ID、编码、名称）
        if name_field:
            field_keys_csv = f"{id_field},{number_field},{name_field}"
        else:
            field_keys_csv = f"{id_field},{number_field}"
        
        print(f"    [步骤1] 使用 ExecuteBillQuery 查询")
        print(f"    查询字段: {field_keys_csv}")
        
        # 分页查询
        start_row, limit = 0, 2000
        
        while True:
            rows = client.execute_bill_query({
                "FormId": "BD_MATERIAL",
                "FieldKeys": field_keys_csv,
                "FilterString": "",
                "OrderString": "",
                "TopRowCount": 0,
                "StartRow": start_row,
                "Limit": limit,
            })
            
            if not rows:
                break
            
            for row in rows:
                if len(row) >= 2:
                    # 检查第一个字段是否看起来像ID（数字且不为0）
                    if row[0] and str(row[0]).strip() != "0":
                        try:
                            f_id = str(int(row[0]))
                        except:
                            f_id = str(row[0]).strip()
                        
                        # 如果只需要查询缺失的ID，跳过已有的
                        if missing_ids and f_id in material_code_mapping:
                            continue
                        
                        f_number = str(row[1]).strip() if row[1] else ""
                        
                        # 尝试获取名称（第三个字段）
                        f_name = ""
                        if len(row) >= 3 and row[2]:
                            # 名称可能是列表格式 {"Key": xxx, "Value": "名称"}
                            try:
                                import json
                                if isinstance(row[2], str) and row[2].startswith("["):
                                    name_list = json.loads(row[2])
                                    if name_list and isinstance(name_list, list) and len(name_list) > 0:
                                        f_name = name_list[0].get("Value", "")
                                else:
                                    f_name = str(row[2]).strip()
                            except:
                                f_name = str(row[2]).strip()
                        
                        if f_id and f_number:
                            material_code_mapping[f_id] = f_number
                            if f_name:
                                material_name_mapping[f_id] = f_name
            
            if len(rows) < limit:
                break
            start_row += limit
        
        print(f"    ExecuteBillQuery 获取到 {len(material_code_mapping)} 个物料映射")
        
    except Exception as e:
        logger.error("ExecuteBillQuery获取物料映射失败: %s", e, exc_info=True)
        print(f"    ExecuteBillQuery 获取失败: {e}")
    
    # 步骤2: 如果ExecuteBillQuery没有获取到足够的映射，使用client.view逐个查询
    if len(material_code_mapping) == 0:
        print("    [步骤2] ExecuteBillQuery未获取到物料内码，切换到API逐个查询")
        
        if required_material_ids:
            ids_to_query = required_material_ids
        else:
            print("    正在获取所有物料ID列表...")
            ids_to_query = []
            try:
                field_info = client.get_form_field_keys("BD_MATERIAL")
                entry = _main_entry(field_info.get("entries", []))
                field_keys = entry["field_keys"] if entry else field_info["field_keys"]
                
                rows = client.execute_bill_query({
                    "FormId": "BD_MATERIAL",
                    "FieldKeys": field_keys[0],
                    "FilterString": "",
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": 2000,
                })
                
                print("    警告: 无法从ExecuteBillQuery获取物料ID列表")
                
            except Exception as e:
                logger.error("获取物料ID列表失败: %s", e, exc_info=True)
                print(f"    获取物料ID列表失败: {e}")
        
        # 如果没有指定需要的物料ID，使用一些默认测试ID
        if not ids_to_query:
            ids_to_query = ['106175', '106170', '106176', '107923', '107390']
        
        # 使用client.view逐个查询
        total_count = len(ids_to_query)
        success_count = 0
        print(f"    准备查询 {total_count} 个物料...")
        
        for idx, id_str in enumerate(ids_to_query, 1):
            if id_str in material_code_mapping:
                _print_progress(idx, total_count, prefix="    查询进度", suffix=f"已完成 {success_count} 个")
                continue
            
            try:
                result = client.view("BD_MATERIAL", {"Id": id_str, "Number": "", "CreateOrgId": 0})
                inner_result = result.get("Result", {}).get("Result", {})
                if inner_result:
                    code = inner_result.get("Number", "")
                    name = inner_result.get("Name", "")
                    
                    # 处理名称（可能是列表格式）
                    if isinstance(name, list) and len(name) > 0:
                        name = name[0].get("Value", "") if isinstance(name[0], dict) else str(name[0])
                    
                    if code:
                        material_code_mapping[id_str] = code
                        success_count += 1
                        if name:
                            material_name_mapping[id_str] = name
            except Exception as e:
                logger.error("API逐个查询物料编码和名称失败: %s", e, exc_info=True)
                pass
            _print_progress(idx, total_count, prefix="    查询进度", suffix=f"已完成 {success_count} 个")
        
        print()  # 换行
        print(f"    API查询获取到 {len(material_code_mapping)} 个物料映射")

    # 步骤2.5: 如果仍有未解析的ID，用view API逐个查询
    if required_material_ids:
        still_missing = [mid for mid in required_material_ids if mid not in material_code_mapping or not material_code_mapping[mid]]
        if still_missing:
            print(f"    [步骤2.5] 逐个查询 {len(still_missing)} 个仍未解析的物料ID")
            import time as _time
            for mid in still_missing:
                try:
                    result = client.view("BD_MATERIAL", {"Id": int(mid), "Number": "", "CreateOrgId": 0})
                    inner = result.get("Result", {}).get("Result", {})
                    code = inner.get("Number", "")
                    if code:
                        material_code_mapping[mid] = code
                    name_raw = inner.get("Name", "")
                    if isinstance(name_raw, list) and len(name_raw) > 0:
                        fname = name_raw[0].get("Value", "") if isinstance(name_raw[0], dict) else str(name_raw[0])
                    elif name_raw:
                        fname = str(name_raw).strip()
                    else:
                        fname = ""
                    if fname:
                        material_name_mapping[mid] = fname
                except Exception:
                    pass
                _time.sleep(0.1)
            resolved_count = sum(1 for mid in still_missing if mid in material_code_mapping and material_code_mapping[mid])
            print(f"    逐个查询后解析了 {resolved_count}/{len(still_missing)} 个")
    
    # 步骤3: 保存缓存（包括新查询到的数据）
    new_count = len(material_code_mapping) - initial_count
    if new_count > 0:
        print(f"    [缓存] 本次新增 {new_count} 个物料映射，已保存到map文件夹")
    _save_material_mapping(material_code_mapping, material_name_mapping)
    
    return material_code_mapping, material_name_mapping


def fetch_form_data(client: KingdeeClient, form_id: str, specific_fields: list[str] = None, filter_string: str = "") -> pd.DataFrame:
    """
    1. QueryBusinessInfo → 取主表 FieldKeys
    2. ExecuteBillQuery 分页拉全量数据
    如果指定了specific_fields，则只查询这些字段
    如果指定了filter_string，则按条件过滤
    """
    print(f"  [字段] {form_id} ...", end=" ", flush=True)
    
    if specific_fields:
        field_keys = specific_fields
        print(f"指定字段: {len(field_keys)} 个")
    else:
        field_info = client.get_form_field_keys(form_id)
        field_keys = field_info.get("field_keys", [])
        print(f"{len(field_keys)} 个字段")

    field_keys_csv = ",".join(field_keys)
    all_rows: list[Any] = []
    start_row, limit = 0, 2000

    while True:
        rows = client.execute_bill_query({
            "FormId":      form_id,
            "FieldKeys":   field_keys_csv,
            "FilterString": filter_string,
            "OrderString":  "",
            "TopRowCount":  0,
            "StartRow":     start_row,
            "Limit":        limit,
            "SubSystemId":  "",
        })
        if not rows:
            break
        all_rows.extend(rows)
        if len(rows) < limit:
            break
        start_row += limit

    print(f"  [数据] {form_id}: 共 {len(all_rows)} 行")
    if all_rows:
        return pd.DataFrame(all_rows, columns=field_keys)
    return pd.DataFrame(columns=field_keys)


def _clear_data_dir_csv_cache() -> None:
    """清理 data/ 目录下的 CSV 缓存文件（保留最近2个）。"""
    csv_files = sorted(DATA_DIR.glob("*.csv"), key=lambda p: p.stat().st_mtime)
    for f in csv_files[:-2]:
        f.unlink()


def fetch_all_from_kingdee(skip_fetch: bool = False, form: str = "all") -> None:
    """从金蝶拉取表单数据。"""
    if skip_fetch:
        print("[跳过] 从金蝶拉取数据")
        return

    print("[开始] 从金蝶拉取数据")
    DATA_DIR.mkdir(exist_ok=True)
    
    # 清空data文件夹中的旧CSV文件
    for old_file in DATA_DIR.glob("*.csv"):
        old_file.unlink()
    print("  [清理] data/ 目录已清空")
    
    client = KingdeeClient()
    
    # 步骤1: 导出 BD_MATERIAL（物料基础资料）
    if form == "all" or form == "BD_MATERIAL":
        print(f"\n[步骤1] 导出 {MATERIAL_FORM_ID}（物料基础资料）")
        # 获取字段信息
        field_info = client.get_form_field_keys(MATERIAL_FORM_ID)
        # BD_MATERIAL 数据分布在多个 entry 中，合并主表和需要的子表字段
        # Entry 0: FBillHead(主表), Entry 4-9: SubHeadEntity(计划/采购/质检等), Entry 13: SubHeadEntity6
        entries = field_info.get("entries", [])
        needed_entry_keys = {0, 4, 5, 6, 7, 8, 9, 13}
        field_keys = []
        for i, entry in enumerate(entries):
            if i in needed_entry_keys:
                field_keys.extend(entry.get("field_keys", []))
        
        # 分页查询所有数据
        all_rows = []
        start_row, limit = 0, 2000
        total_fetched = 0
        
        print("  [数据] 正在分页查询...")
        while True:
            rows = client.execute_bill_query({
                "FormId": MATERIAL_FORM_ID,
                "FieldKeys": ",".join(field_keys),
                "FilterString": "",
                "OrderString": "",
                "TopRowCount": 0,
                "StartRow": start_row,
                "Limit": limit,
            })
            
            if not rows:
                break
            
            all_rows.extend(rows)
            total_fetched += len(rows)
            print(f"    已获取 {total_fetched} 行...")
            
            if len(rows) < limit:
                break
            start_row += limit
        
        material_df = pd.DataFrame(all_rows, columns=field_keys)
        material_csv_path = DATA_DIR / f"{MATERIAL_FORM_ID}_{RUN_TS}.csv"
        material_df.to_csv(material_csv_path, index=False, encoding="utf-8-sig", quoting=1)
        print(f"  [保存] {material_csv_path.name}")
        print(f"  [数据] 共 {len(material_df)} 行")
    
    # 步骤2: 导出 ENG_BOM（物料清单）
    if form == "all" or form == "ENG_BOM":
        print(f"\n[步骤2] 导出 {FORM_ID}（物料清单）")
        df = fetch_form_data(client, FORM_ID)
        
        # 获取物料ID映射（需要BD_MATERIAL数据）
        print(f"\n[步骤3] 获取物料ID映射")
        required_material_ids = []
        if "FMATERIALID" in df.columns:
            required_material_ids = df["FMATERIALID"].dropna().unique().tolist()
            # 转换为字符串
            required_material_ids = [str(int(id_val)) if isinstance(id_val, (int, float)) and id_val == int(id_val) else str(id_val).strip() 
                                    for id_val in required_material_ids if id_val and str(id_val).strip() != "0"]
            print(f"    需要查询 {len(required_material_ids)} 个物料ID")
        
        material_code_mapping, material_name_mapping = fetch_material_mapping(client, required_material_ids)
        
        # 转换ID为编码和名称
        _convert_id_to_code_with_mapping(df, client, material_code_mapping, material_name_mapping)
        
        csv_path = DATA_DIR / f"{FORM_ID}_{RUN_TS}.csv"
        df.to_csv(csv_path, index=False, encoding="utf-8-sig", quoting=1)
        print(f"  [保存] {csv_path.name}")
    
    # 步骤3: 导出 PLN_FORECAST（预测单）
    if form == "all" or form == "PLN_FORECAST":
        print(f"\n[步骤3] 导出 {PLN_FORECAST_FORM_ID}（预测单）")
        try:
            forecast_df = fetch_form_data(client, PLN_FORECAST_FORM_ID, filter_string="FCloseStatus = 'B'")
            forecast_csv_path = DATA_DIR / f"{PLN_FORECAST_FORM_ID}_{RUN_TS}.csv"
            forecast_df.to_csv(forecast_csv_path, index=False, encoding="utf-8-sig", quoting=1)
            print(f"  [保存] {forecast_csv_path.name}")
            print(f"  [数据] 共 {len(forecast_df)} 行")
        except Exception as e:
            logger.error("PLN_FORECAST 数据拉取失败: %s", e, exc_info=True)
            print(f"  [错误] PLN_FORECAST 数据拉取失败: {e}")
            import traceback
            traceback.print_exc()

    # 步骤4: 导出 PLN_PLANORDER（计划订单）
    if form == "all" or form == "PLN_PLANORDER":
        print(f"\n[步骤4] 导出 {PLANORDER_FORM_ID}（计划订单）")
        try:
            field_info = client.get_form_field_keys(PLANORDER_FORM_ID)
            entry = _first_entry(field_info.get("entries", []))
            main_fields = entry["field_keys"] if entry else field_info["field_keys"]
            planorder_df = fetch_form_data(client, PLANORDER_FORM_ID, specific_fields=main_fields)
            planorder_csv_path = DATA_DIR / f"{PLANORDER_FORM_ID}_{RUN_TS}.csv"
            planorder_df.to_csv(planorder_csv_path, index=False, encoding="utf-8-sig", quoting=1)
            print(f"  [保存] {planorder_csv_path.name}")
            print(f"  [数据] 共 {len(planorder_df)} 行")
        except Exception as e:
            logger.error("PLN_PLANORDER 数据拉取失败: %s", e, exc_info=True)
            print(f"  [错误] PLN_PLANORDER 数据拉取失败: {e}")
            import traceback
            traceback.print_exc()

    # 步骤5: 导出 PRD_MO（生产订单）
    if form == "all" or form == "PRD_MO":
        print(f"\n[步骤5] 导出 {MO_FORM_ID}（生产订单）")
        try:
            field_info = client.get_form_field_keys(MO_FORM_ID)
            entries = field_info.get("entries", [])
            # PRD_MO 关键字段分布在 Entry 0（单据头）和 Entry 1（明细），需要合并
            header_fields = entries[0]["field_keys"] if len(entries) > 0 else []
            detail_fields = entries[1]["field_keys"] if len(entries) > 1 else []
            # 合并字段，去重（明细字段优先，因为包含 FMaterialId, FQty, FStatus 等）
            seen = set()
            mo_fields = []
            for f in header_fields + detail_fields:
                if f not in seen:
                    seen.add(f)
                    mo_fields.append(f)
            print(f"  [字段] 合并单据头({len(header_fields)})+明细({len(detail_fields)}) = {len(mo_fields)} 个字段")
            mo_df = fetch_form_data(client, MO_FORM_ID, specific_fields=mo_fields)
            mo_csv_path = DATA_DIR / f"{MO_FORM_ID}_{RUN_TS}.csv"
            mo_df.to_csv(mo_csv_path, index=False, encoding="utf-8-sig", quoting=1)
            print(f"  [保存] {mo_csv_path.name}")
            print(f"  [数据] 共 {len(mo_df)} 行")
        except Exception as e:
            logger.error("PRD_MO 数据拉取失败: %s", e, exc_info=True)
            print(f"  [错误] PRD_MO 数据拉取失败: {e}")
            import traceback
            traceback.print_exc()

    # 步骤6: 导出 PUR_PurchaseOrder（采购订单）
    if form == "all" or form == "PUR_PurchaseOrder":
        print(f"\n[步骤6] 导出 {PO_FORM_ID}（采购订单）")
        try:
            field_info = client.get_form_field_keys(PO_FORM_ID)
            entries = field_info.get("entries", [])
            # PUR_PurchaseOrder 关键字段分布在 Entry 0（基本信息）和 Entry 3（明细信息），需要合并
            header_fields = entries[0]["field_keys"] if len(entries) > 0 else []
            detail_fields = entries[3]["field_keys"] if len(entries) > 3 else []
            # 合并字段，去重
            seen = set()
            po_fields = []
            for f in header_fields + detail_fields:
                if f not in seen:
                    seen.add(f)
                    po_fields.append(f)
            print(f"  [字段] 合并基本信息({len(header_fields)})+明细信息({len(detail_fields)}) = {len(po_fields)} 个字段")
            po_df = fetch_form_data(client, PO_FORM_ID, specific_fields=po_fields, filter_string="FCloseStatus = 'A'")
            po_csv_path = DATA_DIR / f"{PO_FORM_ID}_{RUN_TS}.csv"
            po_df.to_csv(po_csv_path, index=False, encoding="utf-8-sig", quoting=1)
            print(f"  [保存] {po_csv_path.name}")
            print(f"  [数据] 共 {len(po_df)} 行")
        except Exception as e:
            logger.error("PUR_PurchaseOrder 数据拉取失败: %s", e, exc_info=True)
            print(f"  [错误] PUR_PurchaseOrder 数据拉取失败: {e}")
            import traceback
            traceback.print_exc()

    # 步骤7: 导出 PRD_INSTOCK（生产入库单）
    if form == "all" or form == "PRD_INSTOCK":
        print(f"\n[步骤7] 导出 {INSTOCK_FORM_ID}（生产入库单）")
        try:
            field_info = client.get_form_field_keys(INSTOCK_FORM_ID)
            entries = field_info.get("entries", [])
            # PRD_INSTOCK 关键字段分布在多个分录，合并所有字段
            all_fields = []
            seen = set()
            for entry in entries:
                for f in entry.get("field_keys", []):
                    if f not in seen:
                        seen.add(f)
                        all_fields.append(f)
            print(f"  [字段] 合并 {len(entries)} 个分录 = {len(all_fields)} 个字段")
            instock_df = fetch_form_data(client, INSTOCK_FORM_ID, specific_fields=all_fields)
            instock_csv_path = DATA_DIR / f"{INSTOCK_FORM_ID}_{RUN_TS}.csv"
            instock_df.to_csv(instock_csv_path, index=False, encoding="utf-8-sig", quoting=1)
            print(f"  [保存] {instock_csv_path.name}")
            print(f"  [数据] 共 {len(instock_df)} 行")
        except Exception as e:
            logger.error("PRD_INSTOCK 数据拉取失败: %s", e, exc_info=True)
            print(f"  [错误] PRD_INSTOCK 数据拉取失败: {e}")
            import traceback
            traceback.print_exc()

    # 步骤8: 导出 STK_Inventory（即时库存）
    if form == "all" or form == "STK_Inventory":
        print(f"\n[步骤8] 导出 {INV_FORM_ID}（即时库存）")
        # STK_Inventory 不需要预拉取CSV，导出函数直接查询API
        print("  [跳过] STK_Inventory 导出时直接查询API，无需预拉取CSV")

    # 步骤9: 导出 PRD_PickMtrl（生产领料单）— 默认不拉取，需显式指定
    if form == "PRD_PickMtrl":
        print(f"\n[步骤9] 导出 {PICKMTRL_FORM_ID}（生产领料单）")
        try:
            field_info = client.get_form_field_keys(PICKMTRL_FORM_ID)
            entries = field_info.get("entries", [])
            header_fields = entries[0]["field_keys"] if len(entries) > 0 else []
            detail_fields = entries[1]["field_keys"] if len(entries) > 1 else []
            seen = set()
            pickmtrl_fields = []
            for f in header_fields + detail_fields:
                if f not in seen:
                    seen.add(f)
                    pickmtrl_fields.append(f)
            print(f"  [字段] 合并单据头({len(header_fields)})+明细({len(detail_fields)}) = {len(pickmtrl_fields)} 个字段")
            pickmtrl_df = fetch_form_data(client, PICKMTRL_FORM_ID, specific_fields=pickmtrl_fields)
            pickmtrl_csv_path = DATA_DIR / f"{PICKMTRL_FORM_ID}_{RUN_TS}.csv"
            pickmtrl_df.to_csv(pickmtrl_csv_path, index=False, encoding="utf-8-sig", quoting=1)
            print(f"  [保存] {pickmtrl_csv_path.name}")
            print(f"  [数据] 共 {len(pickmtrl_df)} 行")
        except Exception as e:
            logger.error("PRD_PickMtrl 数据拉取失败: %s", e, exc_info=True)
            print(f"  [错误] PRD_PickMtrl 数据拉取失败: {e}")
            import traceback
            traceback.print_exc()

    print("[完成] 从金蝶拉取数据")


# ══════════════════════════════════════════════════════════════════════════════
# 2. 生成导入 Excel
# ══════════════════════════════════════════════════════════════════════════════

def _match_field(tpl_field: str, data_cols: list[str]) -> int | None:
    """匹配模板字段到数据字段，返回数据列索引。"""
    tpl_field = str(tpl_field).strip()
    if not tpl_field:
        return None
    
    if tpl_field in data_cols:
        return data_cols.index(tpl_field)
    
    # 大小写不敏感匹配
    tpl_field_upper = tpl_field.upper()
    for idx, col in enumerate(data_cols):
        if col.upper() == tpl_field_upper:
            return idx
    
    if "(" in tpl_field:
        clean = tpl_field.split("(")[0].strip()
        if clean in data_cols:
            return data_cols.index(clean)
    
    if "#" in tpl_field:
        base_field = tpl_field.split("#")[0].strip()
        if base_field in data_cols:
            return data_cols.index(base_field)
    
    special_mappings = [
        ("BOM类别", "FBOMCATEGORY"),
        ("BOM用途", "FBOMUSE"),
        ("单据类型", "FBILLTYPE"),
        ("父项物料编码", "FMATERIALID"),
        ("BOM分组", "FGroup"),
        ("创建组织", "FCreateOrgId"),
        ("使用组织", "FUseOrgId"),
        ("包装规格", "FParentAuxPropId"),
        ("产品模型", "FMDLID"),
    ]
    for keyword, data_field in special_mappings:
        if keyword in tpl_field and data_field in data_cols:
            return data_cols.index(data_field)
    
    forecast_mappings = [
        ("FBILLTYPE", "FBillTypeID"),
        ("FBILLTYPEID", "FBillTypeID"),
        ("FBILLNO", "FBillNo"),
        ("FDATE", "FDate"),
        ("FORGID", "FForeOrgId"),
        ("FCREATEORGID", "FCreatorId"),
        ("FMATERIALID", "FMaterialID"),
        ("FPREDICTDATE", "FStartDate"),
        ("FPREDICTQTY", "FQty"),
        ("FUNITID", "FUnitID"),
        ("FBASEUNITID", "FBaseUnitID"),
        ("FBASEQTY", "FBaseQty"),
        ("FSALEORGID", "FSupplyOrgId"),
        ("FSTOCKORGID", "FStockOrgId"),
        ("FRECEIVEORGID", "FStockOrgId"),
        ("FSUPPLIERID", "FSupplyOrgId"),
        ("FCUSTID", "FCustID"),
        ("FREMARK", "FRemark"),
        ("FBILLSTATUS", "FDocumentStatus"),
        ("FDELETEFLAG", "FCancelStatus"),
    ]
    for tpl_key, data_field in forecast_mappings:
        if tpl_key in tpl_field_upper:
            for idx, col in enumerate(data_cols):
                if col.upper() == data_field.upper():
                    return idx
    
    return None


def generate_material_excel(lang: str = "CN") -> None:
    import time as _time
    lang_label = "中文" if lang == "CN" else "英文"
    print(f"[开始] 生成物料基础资料 Excel ({lang_label})")
    OUTPUT_DIR.mkdir(exist_ok=True)

    template_suffix = "中文模板.xlsx" if lang == "CN" else "英文模板.xlsx"
    template_files = list(TEMPLATE_DIR.glob(f"{MATERIAL_FORM_ID}*{template_suffix}"))
    if not template_files:
        print(f"  [错误] 找不到物料基础资料{lang_label}模板文件")
        return

    # ── 辅助函数：递归搜索View API返回结果中的F_PALW_Assistant5品牌名称 ──
    def _find_brand_name(obj, target_guid, result_map):
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k == "F_PALW_Assistant5" and isinstance(v, dict):
                    aid = v.get("Id", "")
                    if aid == target_guid:
                        fdata_values = v.get("FDataValue", [])
                        for item in fdata_values:
                            if item.get("Key") == 1033:  # English
                                result_map[target_guid] = item.get("Value", "")
                                return True
                        result_map[target_guid] = v.get("FNumber", "")
                        return True
                elif isinstance(v, (dict, list)):
                    if _find_brand_name(v, target_guid, result_map):
                        return True
        elif isinstance(obj, list):
            for item in obj:
                if _find_brand_name(item, target_guid, result_map):
                    return True
        return False

    template_path = template_files[0]
    print(f"  [模板] {template_path.name}")

    wb = openpyxl.load_workbook(template_path)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    h0 = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            h0 = [str(c).strip() if c else "" for c in row]
            break

    print(f"  [表头] {len(h0)} 列")

    data_path = _resolve_local_csv(MATERIAL_FORM_ID)
    if not data_path:
        print("  [错误] 找不到物料基础资料数据文件")
        return

    print(f"  [数据] {data_path.name}")
    df = pd.read_csv(data_path, encoding="utf-8-sig", dtype=str)
    data_cols = df.columns.tolist()
    print(f"  [数据] {len(df)} 行, {len(data_cols)} 列")

    # ── 去重：同一FNumber可能有多行（不同使用组织），FForbidStatus=B(已禁用)优先 ──
    if "FNumber" in data_cols and "FForbidStatus" in data_cols:
        before_dedup = len(df)
        # 按FForbidStatus排序，B(已禁用)排在前面，然后取每个FNumber的第一行
        df["_sort_key"] = df["FForbidStatus"].apply(lambda x: 0 if str(x).strip() == "B" else 1)
        df = df.sort_values("_sort_key").drop_duplicates(subset=["FNumber"], keep="first").drop(columns=["_sort_key"])
        df = df.reset_index(drop=True)
        if len(df) < before_dedup:
            print(f"  [去重] FNumber去重: {before_dedup} → {len(df)} 行 (保留FForbidStatus=B的行)")
        # 过滤掉已禁用的物料（只保留Disabling Status=No）
        before_filter = len(df)
        df = df[df["FForbidStatus"].apply(lambda x: str(x).strip() != "B")].reset_index(drop=True)
        if len(df) < before_filter:
            print(f"  [过滤] 移除已禁用物料: {before_filter} → {len(df)} 行")

    material_field_map = {
        "使用组织": "FUseOrgId", "Use Org.": "FUseOrgId",
        "编码": "FNumber", "Code": "FNumber",
        "名称": "FName", "Name": "FName",
        "规格型号": "FSpecification", "Specification": "FSpecification",
        "助记码": "FMnemonicCode", "Mnemocode": "FMnemonicCode",
        "数据状态": "FDocumentStatus", "Data Status": "FDocumentStatus",
        "禁用状态": "FForbidStatus", "Disabling Status": "FForbidStatus",
        "物料属性": "FErpClsID", "Material Property": "FErpClsID",
        "存货类别": "FCategoryID", "Inventory Category": "FCategoryID",
        "基本单位": "FBaseUnitId", "BUoM": "FBaseUnitId",
        "启用保质期管理": "FIsKFPeriod", "Enable guarantee period management": "FIsKFPeriod",
        "旧物料编码": "FOldNumber", "Old Material Code": "FOldNumber",
        "保质期": "FExpPeriod", "Guarantee Period": "FExpPeriod",
        "最小库存": "FMinStock", "Min. Stock": "FMinStock",
        "安全库存": "FSafeStock", "Safety Stock": "FSafeStock",
        "最大库存": "FMaxStock", "Max. Inventory": "FMaxStock",
        "采购单位": "FPurchaseUnitId", "Purchase Unit": "FPurchaseUnitId",
        "采购组织": "FPurchaseOrgId", "Purchase Org.": "FPurchaseOrgId",
        "计划策略": "FPlanningStrategy", "Planning Strategy": "FPlanningStrategy",
        "制造策略": "FMfgPolicyId", "Manufacturing Policy": "FMfgPolicyId",
        "订货策略": "FOrderPolicy", "Order Policy": "FOrderPolicy",
        "最小订货量": "FMinPOQty", "Min. Order Qty": "FMinPOQty",
        "固定/经济批量": "FEOQ", "FOQ/EOQ": "FEOQ",
        "预计入库允许部分提前": "FAllowPartAhead", "Partial Advance Allowed for Scheduled Receipt": "FAllowPartAhead",
        "时间紧迫系数": "FTimeFactorId", "Time Urgent Coefficient": "FTimeFactorId",
        "预计入库允许部分延后": "FAllowPartDelay", "Partial Delay Allowed for Scheduled Receipt": "FAllowPartDelay",
        "数量负荷系数": "FQtyFactorId", "Qty Load Coefficient": "FQtyFactorId",
        "标准人员准备工时": "FStdLaborPrePareTime", "Standard Labor Setup Time": "FStdLaborPrePareTime",
        "标准人员实作工时": "FStdLaborProcessTime", "Standard Actual Labor Hours": "FStdLaborProcessTime",
        "标准机器准备工时": "FStdMachinePrepareTime", "Standard Machine Setup Time": "FStdMachinePrepareTime",
        "标准机器实作工时": "FStdMachineProcessTime", "Standard Actual Machine Hours": "FStdMachineProcessTime",
        "辅助属性": "F_WTRM_Combo", "Secondary Property": "F_WTRM_Combo",
        "来料检验": "FCheckIncoming", "Incoming Material Inspection": "FCheckIncoming",
        "库存检验": "FCheckStock", "Inventory Inspection": "FCheckStock",
        "产品检验": "FCheckProduct", "Product Inspection": "FCheckProduct",
        "抽样方案": "FIncSampSchemeId", "Sampling Scheme": "FIncSampSchemeId",
        "质检方案": "FIncQcSchemeId", "Inspection Scheme": "FIncQcSchemeId",
        "启用库存周期复检": "FEnableCyclistQCSTK", "Enable inventory cycle reinspection": "FEnableCyclistQCSTK",
        "复检周期": "FStockCycle", "Reinspection Cycle": "FStockCycle",
        "启用库存周期复检提醒": "FEnableCyclistQCSTKEW", "Enable reinspection reminder of inventory cycle": "FEnableCyclistQCSTKEW",
        "提醒提前期": "FEWLeadDay", "Alert Lead Time": "FEWLeadDay",
        "质检组": "FInspectGroupId", "Inspection Team": "FInspectGroupId",
        "固定损耗": "FFIXLOSS", "Fixed Loss": "FFIXLOSS",
        "已使用": "FRefStatus", "Used": "FRefStatus",
        "排产识别码": "F_WTRM_Text211111", "Sequence": "F_WTRM_Text211111",
        "密度": "F_WTRM_Text31", "Density": "F_WTRM_Text31",
        "品牌": "F_PALW_Assistant5", "Brand": "F_PALW_Assistant5",
        "文件要求": "", "Documents need": "",
        "Blending Key": "", "BlendingKey": "",
        "订货间隔期": "FOrderIntervalTime", "Order Interval": "FOrderIntervalTime",
        "Packaging Key": "F_WTRM_Text3111", "Packaging Key": "F_WTRM_Text3111",
        "最大订货量": "FMaxPOQty", "Max. Order Qty": "FMaxPOQty",
        "高乳化剂含量产品": "F_PALW_CheckBox_apv1", "High Eumlsifier": "F_PALW_CheckBox_apv1",
        "Premix Product（预拌粉）": "F_PALW_CheckBox_tzk1", "Premix Product": "F_PALW_CheckBox_tzk1",
        "KG/carton": "F_PALW_Decimal_qtr2",
        "饱和脂肪每100g1": "F_baohe100", "Text": "F_baohe100",
        "饱和脂肪NRV%1": "F_baoheNRV",
        "糖每100g1": "F_tang100",
        "糖NRV%1": "F_tangNRV",
        "满缸量": "F_PALW_Decimal_qtr6", "full IBC quantity": "F_PALW_Decimal_qtr6",
    }

    field_mapping = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if not tpl_field:
            continue
        if tpl_field in material_field_map:
            data_field = material_field_map[tpl_field]
            if data_field in data_cols:
                field_mapping[tpl_idx] = data_field

    print(f"  [映射] 找到 {len(field_mapping)} 个字段映射")

    client = KingdeeClient()

    def _collect_ids(col_name, id_type="numeric"):
        if col_name not in data_cols:
            return []
        ids = []
        for val in df[col_name].dropna().unique():
            id_str = str(val).strip()
            if not id_str or id_str in ("0", "0.0"):
                continue
            if id_type == "numeric":
                if not _is_guid(id_str):
                    try:
                        num = int(float(id_str))
                        if num > 0:
                            ids.append(str(num))
                    except (ValueError, TypeError, OverflowError):
                        pass
            else:
                ids.append(id_str)
        return list(set(ids))

    org_ids = _collect_ids("FUseOrgId")
    org_code_map = {}
    org_name_map = {}
    if org_ids:
        print(f"  [解析] 需要解析 {len(org_ids)} 个组织ID")
        org_code_map = fetch_generic_mapping(client, "ORG_Organizations", "FOrgId", "FNumber", "org_code_map", org_ids)
        org_name_map = _fetch_name_map(client, "ORG_Organizations", list(set(org_code_map.values())), "org_name_map")

    # 解析基本单位ID
    base_unit_ids = _collect_ids("FBaseUnitId")
    base_unit_code_map = {}
    if base_unit_ids:
        print(f"  [解析] 需要解析 {len(base_unit_ids)} 个基本单位ID")
        base_unit_code_map = fetch_generic_mapping(client, "BD_UNIT", "FUnitId", "FNumber", "unit_code_map", base_unit_ids)

    # 解析采购单位ID
    pur_unit_ids = _collect_ids("FPurchaseUnitId")
    pur_unit_code_map = {}
    if pur_unit_ids:
        print(f"  [解析] 需要解析 {len(pur_unit_ids)} 个采购单位ID")
        pur_unit_code_map = fetch_generic_mapping(client, "BD_UNIT", "FUnitId", "FNumber", "unit_code_map", pur_unit_ids)

    # 单位编码 → 名称映射（用于将L→Liter等）
    all_unit_codes = list(set(list(base_unit_code_map.values()) + list(pur_unit_code_map.values())))
    unit_name_map = {}
    if all_unit_codes:
        unit_name_map = _fetch_name_map(client, "BD_UNIT", all_unit_codes, "unit_name_map")

    # 解析采购组织ID
    pur_org_ids = _collect_ids("FPurchaseOrgId")
    pur_org_code_map = {}
    if pur_org_ids:
        print(f"  [解析] 需要解析 {len(pur_org_ids)} 个采购组织ID")
        pur_org_code_map = fetch_generic_mapping(client, "ORG_Organizations", "FOrgId", "FNumber", "org_code_map", pur_org_ids)

    # 解析存货类别ID → 名称（而非编码）
    category_ids = _collect_ids("FCategoryID")
    category_name_map = {}
    if category_ids:
        print(f"  [解析] 需要解析 {len(category_ids)} 个存货类别ID")
        category_code_map = fetch_generic_mapping(client, "BD_MATERIALCATEGORY", "FCategoryId", "FNumber", "category_code_map", category_ids)
        category_name_map = _fetch_name_map(client, "BD_MATERIALCATEGORY", list(set(category_code_map.values())), "category_name_map")

    # 解析品牌ID (F_PALW_Assistant5 → 内部品牌MC/DB/SV等)
    brand_ids = _collect_ids("F_PALW_Assistant5", id_type="guid")
    brand_name_map = {}
    if brand_ids:
        print(f"  [解析] 需要解析 {len(brand_ids)} 个品牌ID (F_PALW_Assistant5)")
        # 已知的GUID→品牌名称映射
        _known_brand_map = {
            "64dee90e2beea9": "DB", "64dee90e2beea12": "MC", "64dee90e2beea18": "SV",
            "64dee90e2beea16": "SE", "64dee90e2beea11": "HY", "64dee90e2beea14": "PL",
            "64dee90e2beea13": "OA", "64dee90e2beea15": "SC",
        }
        # 对已知GUID直接映射，未知GUID通过View API获取
        unknown_guids = []
        for gid in brand_ids:
            if gid in _known_brand_map:
                brand_name_map[gid] = _known_brand_map[gid]
            else:
                unknown_guids.append(gid)
        if unknown_guids:
            print(f"  [解析] {len(unknown_guids)} 个未知品牌GUID，通过View API获取...")
            for gid in unknown_guids:
                try:
                    # 找到一个使用此GUID的物料
                    sample_rows = client.execute_bill_query({
                        "FormId": "BD_MATERIAL",
                        "FieldKeys": "FMATERIALID",
                        "FilterString": f"F_PALW_Assistant5 = '{gid}'",
                        "OrderString": "", "TopRowCount": 0, "StartRow": 0, "Limit": 1,
                    })
                    if sample_rows:
                        result = client.view("BD_MATERIAL", {"Id": sample_rows[0][0]})
                        if isinstance(result, str):
                            result = json.loads(result)
                        _find_brand_name(result, gid, brand_name_map)
                except Exception as e:
                    logger.error("获取品牌GUID %s 名称失败: %s", gid, e, exc_info=True)
                    print(f"  [警告] 获取品牌GUID {gid} 名称失败: {e}")
                _time.sleep(0.2)
        print(f"  [解析] 品牌映射完成: {brand_name_map}")

    # 解析制造策略ID
    mfg_policy_ids = _collect_ids("FMfgPolicyId")
    mfg_policy_name_map = {}
    if mfg_policy_ids:
        print(f"  [解析] 需要解析 {len(mfg_policy_ids)} 个制造策略ID")
        mfg_policy_code_map = fetch_generic_mapping(client, "PLN_MANUFACTUREPOLICY", "FId", "FNumber", "mfg_policy_code_map", mfg_policy_ids)
        mfg_policy_name_map = _fetch_name_map(client, "PLN_MANUFACTUREPOLICY", list(set(mfg_policy_code_map.values())), "mfg_policy_name_map")

    # 解析质检方案ID
    qc_scheme_ids = _collect_ids("FIncQcSchemeId")
    qc_scheme_name_map = {}
    if qc_scheme_ids:
        print(f"  [解析] 需要解析 {len(qc_scheme_ids)} 个质检方案ID")
        qc_scheme_code_map = fetch_generic_mapping(client, "QM_QCScheme", "FId", "FNumber", "qc_scheme_code_map", qc_scheme_ids)
        qc_scheme_name_map = _fetch_name_map(client, "QM_QCScheme", list(set(qc_scheme_code_map.values())), "qc_scheme_name_map")

    if lang == "CN":
        status_map = {
            "FDocumentStatus": {"A": "已审核", "B": "暂存", "C": "已审核", "D": "重新审核"},
            "FForbidStatus": {"A": "否", "B": "是"},
            "FErpClsID": {"1": "外购", "2": "自制", "3": "委外", "4": "特征", "5": "虚拟", "6": "服务", "7": "一次性", "9": "配置", "10": "资产", "11": "费用", "12": "模型", "13": "产品系列"},
            "FPlanningStrategy": {"0": "MPS", "1": "MRP", "2": "无"},
            "FOrderPolicy": {"0": "LFL(批对批)", "1": "POQ(期间订货)", "2": "FOQ(固定批量)", "3": "ROP(再订货点)"},
            "FRefStatus": {"0": "否", "1": "是"},
        }
        bool_map = {"True": "是", "False": "否", "true": "是", "false": "否", "1": "是", "0": "否"}
        # 辅助属性中文映射：所有非空值都输出RM Specification（对标模板）
        aux_prop_map_cn = {"贸易": "RM specification", "生产": "RM specification", "无特殊": "RM specification", "共用": "RM specification", "出口": "RM specification", "仅供出口用原料": "RM specification"}
    else:
        status_map = {
            "FDocumentStatus": {"A": "Approved", "B": "Draft", "C": "Approved", "D": "Re-audit"},
            "FForbidStatus": {"A": "No", "B": "Yes"},
            "FErpClsID": {"1": "Purchase", "2": "Self-made", "3": "Subcontracted", "4": "Feature", "5": "Virtual", "6": "Service", "7": "One-time", "9": "Configured", "10": "Asset", "11": "Expense", "12": "Model", "13": "Product Family"},
            "FPlanningStrategy": {"0": "MPS", "1": "MRP", "2": "None"},
            "FOrderPolicy": {"0": "LFL", "1": "POQ", "2": "FOQ", "3": "ROP"},
            "FRefStatus": {"0": "No", "1": "Yes"},
        }
        bool_map = {"True": "Yes", "False": "No", "true": "Yes", "false": "No", "1": "Yes", "0": "No"}
        # 辅助属性英文映射：所有非空值都输出RM Specification（对标模板）
        aux_prop_map_en = {"贸易": "RM Specification", "生产": "RM Specification", "无特殊": "RM Specification", "共用": "RM Specification", "出口": "RM Specification", "仅供出口用原料": "RM Specification"}
        # 制造策略英文映射（FNumber→英文名）
        mfg_policy_en_map = {"ZZCL001_SYS": "MTS10(Include inventory)"}
        # 存货类别英文映射
        category_en_map = {"原材料": "Raw material", "自制半成品": "Semi-finished product", "产成品": "Finished product", "辅料": "Auxiliary material", "资产": "Asset", "服务": "Service", "委外半成品": "Subcontracted semi-finished product"}
        # 单位中文名→英文名映射
        unit_en_map = {"千克": "kg", "个": "Pcs", "ppm": "ppm", "Pcs": "Pcs", "套": "Set", "台": "Unit", "升": "Liter", "L": "Liter", "ge": "ge", "tao": "Set"}

    id_to_code_map = {
        "FUseOrgId": org_code_map,
        "FBaseUnitId": base_unit_code_map,
        "FPurchaseUnitId": pur_unit_code_map,
        "FPurchaseOrgId": pur_org_code_map,
        "FCategoryID": category_code_map if 'category_code_map' in dir() else {},
        "F_PALW_Assistant5": brand_name_map,
        "FMfgPolicyId": mfg_policy_code_map if 'mfg_policy_code_map' in dir() else {},
        "FIncQcSchemeId": qc_scheme_code_map if 'qc_scheme_code_map' in dir() else {},
    }

    special_cols = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if tpl_field in ("使用组织", "Use Org."):
            special_cols[tpl_idx] = "org_name"
        elif tpl_field in ("基本单位", "BUoM"):
            special_cols[tpl_idx] = "unit_name"
        elif tpl_field in ("采购单位", "Purchase Unit"):
            special_cols[tpl_idx] = "unit_name"
        elif tpl_field in ("制造策略", "Manufacturing Policy"):
            special_cols[tpl_idx] = "mfg_policy_name"
        elif tpl_field in ("质检方案", "Inspection Scheme"):
            special_cols[tpl_idx] = "qc_scheme_name"
        elif tpl_field in ("存货类别", "Inventory Category"):
            special_cols[tpl_idx] = "category_name"
        elif tpl_field in ("辅助属性", "Secondary Property"):
            special_cols[tpl_idx] = "aux_prop"
        elif tpl_field in ("高乳化剂含量产品", "Premix Product（预拌粉）", "High Eumlsifier", "Premix Product",
                           "启用保质期管理", "Enable guarantee period management",
                           "来料检验", "Incoming Material Inspection",
                           "库存检验", "Inventory Inspection",
                           "产品检验", "Product Inspection",
                           "启用库存周期复检", "Enable inventory cycle reinspection",
                           "启用库存周期复检提醒", "Enable reinspection reminder of inventory cycle",
                           "预计入库允许部分提前", "Partial Advance Allowed for Scheduled Receipt",
                           "预计入库允许部分延后", "Partial Delay Allowed for Scheduled Receipt"):
            special_cols[tpl_idx] = "bool"

    if lang == "CN":
        output_name = "物料列表.xlsx"
        sheet_title = "Sheet1"
    else:
        output_name = "K3_MaterialMaster.xlsx"
        sheet_title = "Sheet1"
    output_path = OUTPUT_DIR / output_name
    for old_file in OUTPUT_DIR.glob(f"{output_name.split('.')[0]}*.xlsx"):
        try:
            old_file.unlink()
        except PermissionError:
            print(f"  [警告] 文件被占用，跳过: {old_file.name}")

    print(f"  [输出] {output_path.name}")

    workbook = openpyxl.Workbook()
    ws_out = workbook.active
    ws_out.title = sheet_title

    ws_out.append(h0)

    for _, row in df.iterrows():
        output_row = [""] * len(h0)

        for tpl_idx, data_col in field_mapping.items():
            val = row[data_col]
            if pd.isna(val) or str(val).strip() == "":
                # 辅助属性列即使为空也需要输出RM Specification（对标模板）
                if tpl_idx in special_cols and special_cols[tpl_idx] == "aux_prop":
                    str_val = "RM Specification" if lang != "CN" else "RM specification"
                    output_row[tpl_idx] = _clean_string(str_val)
                continue

            str_val = str(val).strip()

            if str_val in ("0", "0.0") and data_col in ("FUseOrgId", "FBaseUnitId", "FPurchaseUnitId", "FPurchaseOrgId", "FCategoryID", "F_PALW_Assistant5", "FBomId", "FInspectGroupId", "FIncSampSchemeId", "FIncQcSchemeId", "FMfgPolicyId"):
                continue

            if data_col in id_to_code_map and id_to_code_map[data_col]:
                str_val = _resolve_id(str_val, id_to_code_map[data_col])

            if data_col in status_map:
                str_val = status_map[data_col].get(str_val, str_val)

            if tpl_idx in special_cols:
                sc = special_cols[tpl_idx]
                if sc == "org_name":
                    str_val = org_name_map.get(str_val, str_val)
                elif sc == "unit_name":
                    # 单位编码 → 名称（如 L→Liter, kg→千克），英文版再映射为英文名
                    str_val = unit_name_map.get(str_val, str_val)
                    if lang != "CN":
                        str_val = unit_en_map.get(str_val, str_val)
                elif sc == "mfg_policy_name":
                    # str_val已经是FNumber编码（经过id_to_code_map转换）
                    if lang != "CN":
                        # 英文版：先查mfg_policy_en_map(FNumber→英文)，回退到中文名
                        en_name = mfg_policy_en_map.get(str_val)
                        if en_name:
                            str_val = en_name
                        else:
                            str_val = mfg_policy_name_map.get(str_val, str_val)
                    else:
                        str_val = mfg_policy_name_map.get(str_val, str_val)
                elif sc == "qc_scheme_name":
                    str_val = qc_scheme_name_map.get(str_val, str_val)
                elif sc == "category_name":
                    # FCategoryID → category_code_map → category_name_map → category_en_map
                    cn_name = category_name_map.get(str_val, str_val)
                    if lang != "CN":
                        str_val = category_en_map.get(cn_name, cn_name)
                    else:
                        str_val = cn_name
                elif sc == "aux_prop":
                    # F_WTRM_Combo → RM Specification（对标模板：所有物料都输出RM Specification）
                    if str_val and str_val.strip():
                        mapped = aux_prop_map_en.get(str_val, "RM Specification") if lang != "CN" else aux_prop_map_cn.get(str_val, "RM specification")
                        str_val = mapped
                    else:
                        str_val = "RM Specification" if lang != "CN" else "RM specification"
                elif sc == "bool":
                    str_val = bool_map.get(str_val, str_val)

            output_row[tpl_idx] = _clean_string(str_val)

        ws_out.append(output_row)

    _save_workbook(workbook, output_path)
    print(f"  [完成] 文件大小: {output_path.stat().st_size / 1024 / 1024:.2f} MB")
    print(f"[完成] 生成物料基础资料 Excel ({lang_label})")


def generate_import_excel(lang: str = "CN") -> None:
    lang_label = "中文" if lang == "CN" else "英文"
    print(f"[开始] 生成物料清单 Excel ({lang_label})")
    OUTPUT_DIR.mkdir(exist_ok=True)

    template_suffix = "中文模板.xlsx" if lang == "CN" else "英文模板.xlsx"
    template_files = list(TEMPLATE_DIR.glob(f"{FORM_ID}*{template_suffix}"))
    if not template_files:
        print(f"  [错误] 找不到物料清单{lang_label}模板文件")
        return

    template_path = template_files[0]
    print(f"  [模板] {template_path.name}")

    wb = openpyxl.load_workbook(template_path)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    h0 = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            h0 = [str(c).strip() if c else "" for c in row]
            break

    print(f"  [表头] {len(h0)} 列")

    data_path = _resolve_local_csv(FORM_ID)
    if not data_path:
        print("  [错误] 找不到物料清单数据文件")
        return

    print(f"  [数据] {data_path.name}")
    df = pd.read_csv(data_path, encoding="utf-8-sig", dtype=str)
    data_cols = df.columns.tolist()
    print(f"  [数据] {len(df)} 行, {len(data_cols)} 列")

    bom_field_map = {
        "使用组织": "FUseOrgId", "Use Org.": "FUseOrgId",
        "BOM版本": "FNumber", "BOM Version": "FNumber",
        "BOM简称": "FName", "BOM Short Name": "FName",
        "BOM分类": "FBOMCATEGORY", "BOM Category": "FBOMCATEGORY",
        "BOM用途": "FBOMUSE", "BOM Purpose": "FBOMUSE",
        "父项物料编码": "FMATERIALID", "Parent Material Code": "FMATERIALID",
        "物料名称": "FMATERIALNAME", "Material Name": "FMATERIALNAME",
        "规格型号": "FMATERIALID", "Specification": "FMATERIALID",
        "父项物料单位": "FUNITID", "Parent Material Unit": "FUNITID",
        "数据状态": "FDocumentStatus", "Data Status": "FDocumentStatus",
        "禁用状态": "FForbidStatus", "Disabling Status": "FForbidStatus",
        "项次": "FEntryRowId", "Item": "FEntryRowId",
        "子项物料编码": "FMATERIALIDCHILD", "Subitem Code": "FMATERIALIDCHILD",
        "子项物料名称": "FMATERIALIDCHILD", "Subitem Name": "FMATERIALIDCHILD",
        "子项规格型号": "FMATERIALIDCHILD", "Subitem Specification": "FMATERIALIDCHILD",
        "供应类型": "FISSUETYPE", "Supply Type": "FISSUETYPE",
        "子项单位": "FCHILDUNITID", "Subitem UoM": "FCHILDUNITID",
        "供应组织": "FSUPPLYORG", "Supply Org.": "FSUPPLYORG",
        "用量:分子": "FNUMERATOR", "Consumption: Numerator": "FNUMERATOR",
        "用量:分母": "FDENOMINATOR", "Consumption: Denominator": "FDENOMINATOR",
    }

    field_mapping = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if not tpl_field:
            continue
        if tpl_field in bom_field_map:
            data_field = bom_field_map[tpl_field]
            if data_field in data_cols:
                field_mapping[tpl_idx] = data_field

    print(f"  [映射] 找到 {len(field_mapping)} 个字段映射")

    client = KingdeeClient()

    def _collect_ids(col_name, id_type="numeric"):
        if col_name not in data_cols:
            return []
        ids = []
        for val in df[col_name].dropna().unique():
            id_str = str(val).strip()
            if not id_str or id_str in ("0", "0.0"):
                continue
            if id_type == "numeric":
                if not _is_guid(id_str):
                    try:
                        num = int(float(id_str))
                        if num > 0:
                            ids.append(str(num))
                    except (ValueError, TypeError, OverflowError):
                        pass
            else:
                ids.append(id_str)
        return list(set(ids))

    org_ids = _collect_ids("FUseOrgId")
    supply_org_ids = _collect_ids("FSUPPLYORG")
    all_org_ids = list(set(org_ids + supply_org_ids))

    org_code_map = {}
    org_name_map = {}
    if all_org_ids:
        print(f"  [解析] 需要解析 {len(all_org_ids)} 个组织ID")
        org_code_map = fetch_generic_mapping(client, "ORG_Organizations", "FOrgId", "FNumber", "org_code_map", all_org_ids)
        org_name_map = _fetch_name_map(client, "ORG_Organizations", list(set(org_code_map.values())), "org_name_map")

    child_material_ids = _collect_ids("FMATERIALIDCHILD")
    material_code_map = {}
    material_code_to_name = {}
    material_spec_map = {}
    if child_material_ids:
        print(f"  [解析] 需要解析 {len(child_material_ids)} 个子项物料ID")
        material_code_map, material_name_map = fetch_material_mapping(client, child_material_ids)
        for mid, code in material_code_map.items():
            name = material_name_map.get(mid, "")
            if code and name:
                material_code_to_name[code] = name

        # 补充查询缺少名称的子项物料编码
        missing_name_codes = [c for c in set(material_code_map.values()) if c not in material_code_to_name]
        if missing_name_codes:
            print(f"    需要解析 {len(missing_name_codes)} 个子项物料编码的名称")
            batch_size = 50
            for i in range(0, len(missing_name_codes), batch_size):
                batch = missing_name_codes[i:i + batch_size]
                filter_parts = [f"FNumber = '{code}'" for code in batch]
                filter_str = " or ".join(filter_parts)
                try:
                    rows = client.execute_bill_query({
                        "FormId": "BD_MATERIAL",
                        "FieldKeys": "FNumber,FName",
                        "FilterString": filter_str,
                        "OrderString": "",
                        "TopRowCount": 0,
                        "StartRow": 0,
                        "Limit": len(batch) + 10,
                    })
                    if rows:
                        for r in rows:
                            fnumber = str(r[0]).strip() if r[0] else ""
                            fname_raw = r[1] if r[1] else ""
                            fname = ""
                            if fname_raw:
                                try:
                                    import json as _json
                                    if isinstance(fname_raw, str) and fname_raw.startswith("["):
                                        name_list = _json.loads(fname_raw)
                                        if name_list and isinstance(name_list, list) and len(name_list) > 0:
                                            fname = name_list[0].get("Value", str(name_list[0])) if isinstance(name_list[0], dict) else str(name_list[0])
                                    else:
                                        fname = str(fname_raw).strip()
                                except Exception:
                                    fname = str(fname_raw).strip()
                            if fnumber and fname:
                                material_code_to_name[fnumber] = fname
                except Exception:
                    pass

            # 对于批量查询仍未获取到名称的编码，用view API逐个查询
            still_missing = [c for c in missing_name_codes if c not in material_code_to_name]
            if still_missing:
                print(f"    需要逐个查询 {len(still_missing)} 个仍未获取名称的物料编码")
                # 先找到编码对应的ID
                code_to_id = {v: k for k, v in material_code_map.items()}
                for code in still_missing:
                    mid = code_to_id.get(code, "")
                    if mid:
                        try:
                            result = client.view("BD_MATERIAL", {"Id": int(mid), "Number": "", "CreateOrgId": 0})
                            inner = result.get("Result", {}).get("Result", {})
                            name_raw = inner.get("Name", "")
                            if isinstance(name_raw, list) and len(name_raw) > 0:
                                fname = name_raw[0].get("Value", "") if isinstance(name_raw[0], dict) else str(name_raw[0])
                            elif name_raw:
                                fname = str(name_raw).strip()
                            else:
                                fname = ""
                            if fname:
                                material_code_to_name[code] = fname
                        except Exception:
                            pass
                        import time as _time
                        _time.sleep(0.1)

            # 将新获取的名称保存到material_name_map缓存
            new_name_count = 0
            for mid, code in material_code_map.items():
                name = material_code_to_name.get(code, "")
                if name and mid not in material_name_map:
                    material_name_map[mid] = name
                    new_name_count += 1
            if new_name_count > 0:
                _save_material_mapping(material_code_map, material_name_map)
                print(f"    已保存 {new_name_count} 个新名称到缓存")

        material_spec_map = _load_generic_mapping("material_spec_map")
        missing_spec_codes = [c for c in set(material_code_map.values()) if c not in material_spec_map]
        if missing_spec_codes:
            print(f"    需要解析 {len(missing_spec_codes)} 个子项物料编码的规格型号")
            _fetch_batch_spec(client, missing_spec_codes, material_spec_map)
            _save_generic_mapping("material_spec_map", material_spec_map)

    parent_material_codes = []
    parent_material_ids = []  # 6位物料内码，需要映射为编码
    if "FMATERIALID" in data_cols:
        for val in df["FMATERIALID"].dropna().unique():
            s = str(val).strip()
            if s and s != "0" and s != "0.0":
                # 区分物料编码(>6位)和物料内码(<=6位)
                if len(s) <= 6:
                    parent_material_ids.append(s)
                else:
                    parent_material_codes.append(s)
    parent_material_codes = list(set(parent_material_codes))
    parent_material_ids = list(set(parent_material_ids))

    # 解析6位物料内码为编码
    parent_id_to_code_map = {}
    if parent_material_ids:
        print(f"  [解析] 需要解析 {len(parent_material_ids)} 个父项物料内码为编码")
        parent_id_to_code_map, _ = fetch_material_mapping(client, parent_material_ids)
        resolved = sum(1 for v in parent_id_to_code_map.values() if v)
        print(f"    已解析 {resolved}/{len(parent_material_ids)} 个")

        # 用view API逐个查询未解析的父项物料ID
        unresolved_parent_ids = [pid for pid in parent_material_ids if pid not in parent_id_to_code_map or not parent_id_to_code_map[pid]]
        if unresolved_parent_ids:
            print(f"    需要逐个查询 {len(unresolved_parent_ids)} 个未解析的父项物料ID")
            for pid in unresolved_parent_ids:
                try:
                    result = client.view("BD_MATERIAL", {"Id": int(pid), "Number": "", "CreateOrgId": 0})
                    inner = result.get("Result", {}).get("Result", {})
                    code = inner.get("Number", "")
                    if code:
                        parent_id_to_code_map[pid] = code
                except Exception:
                    pass
                import time as _time
                _time.sleep(0.1)
            resolved2 = sum(1 for v in parent_id_to_code_map.values() if v)
            print(f"    逐个查询后已解析 {resolved2}/{len(parent_material_ids)} 个")

    parent_spec_map = _load_generic_mapping("material_spec_map")
    missing_pspec = [c for c in parent_material_codes if c not in parent_spec_map]
    if missing_pspec:
        print(f"  [解析] 需要解析 {len(missing_pspec)} 个父项物料编码的规格型号")
        _fetch_batch_spec(client, missing_pspec, parent_spec_map)
        _save_generic_mapping("material_spec_map", parent_spec_map)

    parent_unit_codes = []
    if "FUNITID" in data_cols:
        for val in df["FUNITID"].dropna().unique():
            s = str(val).strip()
            if s and s != "0" and s != "0.0":
                parent_unit_codes.append(s)
    parent_unit_codes = list(set(parent_unit_codes))

    unit_name_map = _fetch_name_map(client, "BD_UNIT", parent_unit_codes, "unit_name_map")

    unit_code_map = {}
    child_unit_ids = _collect_ids("FCHILDUNITID")
    if child_unit_ids:
        print(f"  [解析] 需要解析 {len(child_unit_ids)} 个子项单位ID")
        child_unit_code_map = fetch_generic_mapping(client, "BD_UNIT", "FUnitID", "FNumber", "unit_code_map", child_unit_ids)
        child_unit_codes = list(set(child_unit_code_map.values()))
        child_unit_name_map = _fetch_name_map(client, "BD_UNIT", child_unit_codes, "unit_name_map")
        unit_name_map.update(child_unit_name_map)
        unit_code_map.update(child_unit_code_map)

    if lang == "CN":
        status_map = {
            "FDocumentStatus": {"A": "已审核", "B": "暂存", "C": "已审核", "D": "重新审核"},
            "FForbidStatus": {"A": "否", "B": "是"},
            "FISSUETYPE": {"1": "倒冲", "2": "领料", "3": "发料"},
        }
    else:
        status_map = {
            "FDocumentStatus": {"A": "Approved", "B": "Draft", "C": "Approved", "D": "Re-audit"},
            "FForbidStatus": {"A": "No", "B": "Yes"},
            "FISSUETYPE": {"1": "Backflush", "2": "Picking", "3": "Issuing"},
        }

    id_to_code_map = {
        "FUseOrgId": org_code_map,
        "FSUPPLYORG": org_code_map,
        "FMATERIALIDCHILD": material_code_map,
        "FCHILDUNITID": unit_code_map,
        "FMATERIALID": parent_id_to_code_map,
    }

    zero_to_empty_cols = {"FSUPPLYORG", "FMATERIALIDCHILD"}

    special_cols = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if tpl_field in ("使用组织", "Use Org."):
            special_cols[tpl_idx] = "org_name"
        elif tpl_field in ("供应组织", "Supply Org."):
            special_cols[tpl_idx] = "org_name"
        elif tpl_field in ("规格型号", "Specification"):
            special_cols[tpl_idx] = "parent_spec"
        elif tpl_field in ("父项物料单位", "Parent Material Unit"):
            special_cols[tpl_idx] = "unit_name"
        elif tpl_field in ("子项物料名称", "Subitem Name"):
            special_cols[tpl_idx] = "child_name"
        elif tpl_field in ("子项规格型号", "Subitem Specification"):
            special_cols[tpl_idx] = "child_spec"
        elif tpl_field in ("子项单位", "Subitem UoM"):
            special_cols[tpl_idx] = "unit_name"

    if lang == "CN":
        output_name = "物料清单列表.xlsx"
        sheet_title = "Sheet1"
    else:
        output_name = "K3_BOM.xlsx"
        sheet_title = "Sheet1"
    output_path = OUTPUT_DIR / output_name
    for old_file in OUTPUT_DIR.glob(f"{output_name.split('.')[0]}*.xlsx"):
        try:
            old_file.unlink()
        except PermissionError:
            print(f"  [警告] 文件被占用，跳过: {old_file.name}")

    print(f"  [输出] {output_path.name}")

    workbook = openpyxl.Workbook()
    ws_out = workbook.active
    ws_out.title = sheet_title

    if lang == "CN" and "S/N" not in h0:
        h0 = ["序号"] + h0
        field_mapping = {k + 1: v for k, v in field_mapping.items()}
        special_cols = {k + 1: v for k, v in special_cols.items()}

    ws_out.append(h0)

    sn_counter = 0
    for _, row in df.iterrows():
        output_row = [""] * len(h0)
        sn_counter += 1

        for tpl_idx, data_col in field_mapping.items():
            val = row[data_col]
            if pd.isna(val) or str(val).strip() == "":
                continue

            str_val = str(val).strip()

            if str_val in ("0", "0.0") and data_col in zero_to_empty_cols:
                continue

            if data_col in id_to_code_map and id_to_code_map[data_col]:
                str_val = _resolve_id(str_val, id_to_code_map[data_col])

            if data_col in status_map:
                str_val = status_map[data_col].get(str_val, str_val)

            if tpl_idx in special_cols:
                sc = special_cols[tpl_idx]
                if sc == "org_name":
                    str_val = org_name_map.get(str_val, str_val)
                elif sc == "parent_spec":
                    spec = parent_spec_map.get(str_val, "")
                    if spec and spec != str_val:
                        str_val = spec
                    else:
                        str_val = ""
                elif sc == "unit_name":
                    if lang == "CN":
                        str_val = unit_name_map.get(str_val, str_val)
                elif sc == "child_name":
                    name = material_code_to_name.get(str_val, "")
                    if name and name != str_val:
                        str_val = name
                    else:
                        str_val = ""
                elif sc == "child_spec":
                    spec = material_spec_map.get(str_val, "")
                    if spec and spec != str_val:
                        str_val = spec
                    else:
                        str_val = ""

            output_row[tpl_idx] = _clean_string(str_val)

        if "S/N" in h0:
            sn_idx = h0.index("S/N")
            output_row[sn_idx] = str(sn_counter)
        elif "序号" in h0:
            sn_idx = h0.index("序号")
            output_row[sn_idx] = str(sn_counter)

        ws_out.append(output_row)

    _save_workbook(workbook, output_path)
    print(f"  [完成] 文件大小: {output_path.stat().st_size / 1024 / 1024:.2f} MB")
    print(f"[完成] 生成物料清单 Excel ({lang_label})")


def _resolve_id(val: str, cmap: dict[str, str]) -> str:
    if not val or not cmap:
        return val
    if val in cmap:
        return cmap[val]
    if _is_guid(val):
        return cmap.get(val, val)
    try:
        num = int(float(val))
        if num > 0:
            return cmap.get(str(num), val)
    except (ValueError, TypeError, OverflowError):
        pass
    return val


def _fetch_batch_spec(client: KingdeeClient, codes: list[str], spec_map: dict[str, str]) -> None:
    batch_size = 50
    for i in range(0, len(codes), batch_size):
        batch = codes[i:i + batch_size]
        filter_parts = [f"FNumber = '{code}'" for code in batch]
        filter_str = " or ".join(filter_parts)
        try:
            rows = client.execute_bill_query({
                "FormId": "BD_MATERIAL",
                "FieldKeys": "FNumber,FSpecification",
                "FilterString": filter_str,
                "OrderString": "",
                "TopRowCount": 0,
                "StartRow": 0,
                "Limit": len(batch) + 10,
            })
            if rows:
                for row in rows:
                    fnumber = str(row[0]).strip() if row[0] else ""
                    fspec = str(row[1]).strip() if row[1] else ""
                    if fnumber:
                        if not fspec or fspec == " " or fspec == fnumber:
                            spec_map[fnumber] = ""
                        else:
                            spec_map[fnumber] = fspec
        except Exception:
            pass


def _fetch_name_map(client: KingdeeClient, form_id: str, codes: list[str], map_name: str) -> dict[str, str]:
    name_map = _load_generic_mapping(map_name)
    missing = [c for c in codes if c not in name_map]
    if not missing:
        return name_map
    print(f"    需要解析 {len(missing)} 个{map_name}的中文名")
    for code in missing:
        try:
            rows = client.execute_bill_query({
                "FormId": form_id,
                "FieldKeys": "FNumber,FName",
                "FilterString": f"FNumber = '{code}'",
                "OrderString": "",
                "TopRowCount": 0,
                "StartRow": 0,
                "Limit": 10,
            })
            if rows:
                for row in rows:
                    fnumber = str(row[0]).strip() if row[0] else ""
                    fname = str(row[1]).strip() if row[1] else ""
                    if fnumber and fname:
                        name_map[fnumber] = fname
        except Exception:
            pass
    _save_generic_mapping(map_name, name_map)
    return name_map


def generate_forecast_excel(lang: str = "CN") -> None:
    lang_label = "中文" if lang == "CN" else "英文"
    print(f"[开始] 生成预测单 Excel ({lang_label})")
    OUTPUT_DIR.mkdir(exist_ok=True)

    template_suffix = "中文模板.xlsx" if lang == "CN" else "英文模板.xlsx"
    template_files = list(TEMPLATE_DIR.glob(f"{PLN_FORECAST_FORM_ID}*{template_suffix}"))
    if not template_files:
        print(f"  [错误] 找不到预测单{lang_label}模板文件")
        return

    template_path = template_files[0]
    print(f"  [模板] {template_path.name}")

    wb = openpyxl.load_workbook(template_path)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    h0 = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            h0 = [str(c).strip() if c else "" for c in row]
            break

    print(f"  [表头] {len(h0)} 列")

    data_path = _resolve_local_csv(PLN_FORECAST_FORM_ID)
    if not data_path:
        print("  [错误] 找不到预测单数据文件")
        return

    print(f"  [数据] {data_path.name}")
    df = pd.read_csv(data_path, encoding="utf-8-sig", dtype=str)
    data_cols = df.columns.tolist()
    print(f"  [数据] {len(df)} 行, {len(data_cols)} 列")

    forecast_field_map = {
        "单据编号": "FBillNo", "Doc No.": "FBillNo",
        "单据类型": "FBillTypeID", "Doc Type": "FBillTypeID",
        "日期": "FDate", "Date": "FDate",
        "单据状态": "FDocumentStatus", "Doc Status": "FDocumentStatus",
        "预测组织": "FForeOrgId", "Forecast Org.": "FForeOrgId",
        "客户名称": "FCustID", "Customer Name": "FCustID",
        "物料编码": "FMaterialID", "Material Code": "FMaterialID",
        "物料名称": "FMaterialID", "Material Name": "FMaterialID",
        "规格型号": "FMaterialID", "Specification": "FMaterialID",
        "单位": "FUnitID", "UoM": "FUnitID",
        "数量": "FQty", "Qty": "FQty",
        "预测开始日期": "FStartDate", "Forecast Start Date": "FStartDate",
        "预测结束日期": "FEndDate", "Forecast End Date": "FEndDate",
        "计划跟踪号": "FMtoNo", "Plan Tracking No.": "FMtoNo",
        "备注": "FRemark", "Remarks": "FRemark",
        "执行数量": "FWriteOffQty", "Execution Qty": "FWriteOffQty",
        "创建人": "FCreatorId", "Created by": "FCreatorId",
        "关联数量": "FOrderQty", "Associated Qty": "FOrderQty",
        "创建日期": "FCreateDate", "Created on": "FCreateDate",
        "关联生产委外数量": "FSelPrdQty", "Associated Production Subcontract Qty": "FSelPrdQty",
        "审核人": "FApproverId", "Approved by": "FApproverId",
        "关闭状态": "FCloseStatus", "Closing Status": "FCloseStatus",
        "审核日期": "FApproveDate", "Approved on": "FApproveDate",
        "预留类型": "FReserveType", "Reservation Type": "FReserveType",
        "需求优先级": "FPriority", "Requirement Priority": "FPriority",
        "国家": "FCountry", "Country": "FCountry",
        "预测人员": "FForecastPersonId", "Seller": "FForecastPersonId",
        "辅助属性": "FAuxPropId", "Secondary Property": "FAuxPropId",
        "均化类型": "FAVERATYPE", "Equalization Type": "FAVERATYPE",
    }

    field_mapping = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if not tpl_field:
            continue
        if tpl_field in forecast_field_map:
            data_field = forecast_field_map[tpl_field]
            if data_field in data_cols:
                field_mapping[tpl_idx] = data_field

    print(f"  [映射] 找到 {len(field_mapping)} 个字段映射")

    client = KingdeeClient()

    def _collect_ids(col_name, id_type="numeric"):
        if col_name not in data_cols:
            return []
        ids = []
        for val in df[col_name].dropna().unique():
            id_str = str(val).strip()
            if not id_str or id_str in ("0", "0.0"):
                continue
            if id_type == "guid":
                if _is_guid(id_str):
                    ids.append(id_str)
            elif id_type == "numeric":
                if not _is_guid(id_str):
                    try:
                        num = int(float(id_str))
                        if num > 0:
                            ids.append(str(num))
                    except (ValueError, TypeError, OverflowError):
                        pass
            elif id_type == "mixed":
                has_chinese = any('\u4e00' <= c <= '\u9fff' for c in id_str)
                if not has_chinese:
                    ids.append(id_str)
            else:
                ids.append(id_str)
        return list(set(ids))

    print(f"  [解析] 开始解析预测单ID字段...")

    material_ids = _collect_ids("FMaterialID")
    unit_ids = _collect_ids("FUnitID")
    base_unit_ids = _collect_ids("FBaseUnitID")
    org_ids = _collect_ids("FForeOrgId")
    stock_org_ids = _collect_ids("FStockOrgId")
    supply_org_ids = _collect_ids("FSupplyOrgId")
    cust_ids = _collect_ids("FCustID")
    bill_type_ids = _collect_ids("FBillTypeID", "guid")
    creator_ids = _collect_ids("FCreatorId")
    approver_ids = _collect_ids("FApproverId")
    forecast_person_ids = _collect_ids("FForecastPersonId")

    all_org_ids = list(set(org_ids + stock_org_ids + supply_org_ids))
    all_unit_ids = list(set(unit_ids + base_unit_ids))
    all_user_ids = list(set(creator_ids + approver_ids + forecast_person_ids))

    if material_ids:
        print(f"    需要解析 {len(material_ids)} 个物料ID")
        material_code_map, material_name_map = fetch_material_mapping(client, material_ids)
    else:
        material_code_map, material_name_map = _load_material_mapping()

    material_code_to_name = {}
    for mid, code in material_code_map.items():
        name = material_name_map.get(mid, "")
        if code and name:
            material_code_to_name[code] = name

    missing_name_codes = [c for c in set(material_code_map.values()) if c not in material_code_to_name]
    if missing_name_codes:
        print(f"    需要解析 {len(missing_name_codes)} 个物料编码的中文名")
        batch_size = 50
        for i in range(0, len(missing_name_codes), batch_size):
            batch = missing_name_codes[i:i + batch_size]
            filter_parts = [f"FNumber = '{code}'" for code in batch]
            filter_str = " or ".join(filter_parts)
            try:
                rows = client.execute_bill_query({
                    "FormId": "BD_MATERIAL",
                    "FieldKeys": "FNumber,FName",
                    "FilterString": filter_str,
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": len(batch) + 10,
                })
                if rows:
                    for row in rows:
                        fnumber = str(row[0]).strip() if row[0] else ""
                        fname_raw = row[1] if row[1] else ""
                        fname = ""
                        if fname_raw:
                            try:
                                import json as _json
                                if isinstance(fname_raw, str) and fname_raw.startswith("["):
                                    name_list = _json.loads(fname_raw)
                                    if name_list and isinstance(name_list, list) and len(name_list) > 0:
                                        fname = name_list[0].get("Value", str(name_list[0])) if isinstance(name_list[0], dict) else str(name_list[0])
                                else:
                                    fname = str(fname_raw).strip()
                            except Exception:
                                fname = str(fname_raw).strip()
                        if fnumber and fname:
                            material_code_to_name[fnumber] = fname
            except Exception:
                pass

    material_spec_map = _load_generic_mapping("material_spec_map")
    missing_spec_codes = [c for c in set(material_code_map.values()) if c not in material_spec_map]
    if missing_spec_codes:
        print(f"    需要解析 {len(missing_spec_codes)} 个物料编码的规格型号")
        _fetch_batch_spec(client, missing_spec_codes, material_spec_map)
        _save_generic_mapping("material_spec_map", material_spec_map)

    if all_unit_ids:
        print(f"    需要解析 {len(all_unit_ids)} 个单位ID")
        unit_code_map = fetch_generic_mapping(client, "BD_UNIT", "FUnitID", "FNumber", "unit_code_map", all_unit_ids)
    else:
        unit_code_map = _load_generic_mapping("unit_code_map")
    unit_name_map = _fetch_name_map(client, "BD_UNIT", list(set(unit_code_map.values())), "unit_name_map")

    if all_org_ids:
        print(f"    需要解析 {len(all_org_ids)} 个组织ID")
        org_code_map = fetch_generic_mapping(client, "ORG_Organizations", "FOrgId", "FNumber", "org_code_map", all_org_ids)
    else:
        org_code_map = _load_generic_mapping("org_code_map")
    org_name_map = _fetch_name_map(client, "ORG_Organizations", list(set(org_code_map.values())), "org_name_map")

    if cust_ids:
        print(f"    需要解析 {len(cust_ids)} 个客户ID")
        cust_code_map = fetch_generic_mapping(client, "BD_Customer", "FCUSTID", "FNumber", "cust_code_map", cust_ids)
    else:
        cust_code_map = _load_generic_mapping("cust_code_map")
    cust_name_map = _fetch_name_map(client, "BD_Customer", list(set(cust_code_map.values())), "cust_name_map")

    if bill_type_ids:
        print(f"    需要解析 {len(bill_type_ids)} 个单据类型ID(GUID)")
        bill_type_code_map = fetch_generic_mapping(client, "BOS_BillType", "FBILLTYPEID", "FNumber", "bill_type_map", bill_type_ids)
        guid_name_map = fetch_generic_mapping(client, "BOS_BillType", "FBILLTYPEID", "FName", "bill_type_name_map_guid", bill_type_ids)
        bill_type_name_map = {}
        for gid, name in guid_name_map.items():
            code = bill_type_code_map.get(gid, "")
            if code and name:
                bill_type_name_map[code] = name
            if gid and name:
                bill_type_name_map[gid] = name
    else:
        bill_type_code_map = _load_generic_mapping("bill_type_map")
        bill_type_name_map = {}
    bill_type_name_map.update(_load_generic_mapping("bill_type_name_map"))

    bill_type_codes = []
    if "FBillTypeID" in data_cols:
        for val in df["FBillTypeID"].dropna().unique():
            s = str(val).strip()
            if s and s != "0" and not _is_guid(s):
                bill_type_codes.append(s)
    bill_type_codes = list(set(bill_type_codes))
    missing_bt = [c for c in bill_type_codes if c not in bill_type_name_map]
    if missing_bt:
        print(f"    需要解析 {len(missing_bt)} 个单据类型编码的中文名")
        for code in missing_bt:
            try:
                rows = client.execute_bill_query({
                    "FormId": "BOS_BillType",
                    "FieldKeys": "FNumber,FName",
                    "FilterString": f"FNumber = '{code}'",
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": 10,
                })
                if rows:
                    for row in rows:
                        fnumber = str(row[0]).strip() if row[0] else ""
                        fname = str(row[1]).strip() if row[1] else ""
                        if fnumber and fname:
                            bill_type_name_map[fnumber] = fname
            except Exception:
                pass
        _save_generic_mapping("bill_type_name_map", bill_type_name_map)

    user_name_map = _load_generic_mapping("user_name_map")
    missing_users = [uid for uid in all_user_ids if uid not in user_name_map]
    if missing_users:
        print(f"    需要解析 {len(missing_users)} 个用户ID的中文名")
        for uid in missing_users:
            try:
                result = client.view("SEC_User", {"Id": uid, "Number": "", "CreateOrgId": 0})
                inner = result.get("Result", {}).get("Result", {})
                if inner:
                    user_name = inner.get("Name", "")
                    if isinstance(user_name, list) and len(user_name) > 0:
                        user_name = user_name[0].get("Value", "") if isinstance(user_name[0], dict) else str(user_name[0])
                    if not user_name:
                        user_name = inner.get("UserAccount", "")
                    if user_name:
                        user_name_map[uid] = str(user_name).strip()
            except Exception:
                pass
        _save_generic_mapping("user_name_map", user_name_map)

    if lang == "CN":
        status_map = {
            "FDocumentStatus": {"A": "已审核", "B": "暂存", "C": "已审核", "D": "重新审核"},
            "FCancelStatus": {"A": "正常", "B": "已取消"},
            "FCloseStatus": {"A": "已关闭", "B": "未关闭"},
            "FReserveType": {"1": "弱预留", "2": "强预留"},
            "FAVERATYPE": {"N": "不均化", "W": "周均化", "M": "月均化"},
        }
    else:
        status_map = {
            "FDocumentStatus": {"A": "Approved", "B": "Draft", "C": "Approved", "D": "Re-audit"},
            "FCancelStatus": {"A": "Normal", "B": "Cancelled"},
            "FCloseStatus": {"A": "Closed", "B": "Unclosed"},
            "FReserveType": {"1": "Weak Reservation", "2": "Strong Reservation"},
            "FAVERATYPE": {"N": "Unequalization", "W": "Weekly Equalization", "M": "Monthly Equalization"},
        }

    id_to_code_map = {
        "FBillTypeID": bill_type_code_map,
        "FMaterialID": material_code_map,
        "FUnitID": unit_code_map,
        "FBaseUnitID": unit_code_map,
        "FForeOrgId": org_code_map,
        "FStockOrgId": org_code_map,
        "FSupplyOrgId": org_code_map,
        "FCustID": cust_code_map,
        "FCreatorId": user_name_map,
        "FApproverId": user_name_map,
        "FForecastPersonId": user_name_map,
    }

    zero_to_empty_cols = {"FCustID", "FSupplyOrgId", "FAuxPropId", "FBomID", "FParentMtrlId"}

    special_cols = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if tpl_field in ("单据类型", "Doc Type"):
            special_cols[tpl_idx] = "bill_type_name"
        elif tpl_field in ("预测组织", "创建组织", "销售组织", "库存组织", "收货组织", "Forecast Org."):
            special_cols[tpl_idx] = "org_name"
        elif tpl_field in ("物料名称", "Material Name"):
            special_cols[tpl_idx] = "material_name"
        elif tpl_field in ("规格型号", "Specification"):
            special_cols[tpl_idx] = "material_spec"
        elif tpl_field in ("单位", "UoM"):
            special_cols[tpl_idx] = "unit_name"
        elif tpl_field in ("客户名称", "Customer Name"):
            special_cols[tpl_idx] = "cust_name"

    date_cols = {"FDate", "FStartDate", "FEndDate", "FCreateDate", "FApproveDate"}

    if lang == "CN":
        output_name = "预测单列表.xlsx"
        sheet_title = "Sheet1"
    else:
        output_name = "K3_Forecast.xlsx"
        sheet_title = "Sheet1"
    output_path = OUTPUT_DIR / output_name
    for old_file in OUTPUT_DIR.glob(f"{output_name.split('.')[0]}*.xlsx"):
        try:
            old_file.unlink()
        except PermissionError:
            print(f"  [警告] 文件被占用，跳过: {old_file.name}")

    print(f"  [输出] {output_path.name}")

    workbook = openpyxl.Workbook()
    ws_out = workbook.active
    ws_out.title = sheet_title

    ws_out.append(h0)

    for _, row in df.iterrows():
        output_row = [""] * len(h0)

        for tpl_idx, data_col in field_mapping.items():
            val = row[data_col]
            if pd.isna(val) or str(val).strip() == "":
                continue

            str_val = str(val).strip()

            if str_val in ("0", "0.0") and data_col in zero_to_empty_cols:
                continue

            if data_col in id_to_code_map and id_to_code_map[data_col]:
                str_val = _resolve_id(str_val, id_to_code_map[data_col])

            if data_col in status_map:
                str_val = status_map[data_col].get(str_val, str_val)

            if tpl_idx in special_cols:
                sc = special_cols[tpl_idx]
                if sc == "bill_type_name":
                    str_val = bill_type_name_map.get(str_val, str_val)
                elif sc == "org_name":
                    str_val = org_name_map.get(str_val, str_val)
                elif sc == "material_name":
                    name = material_code_to_name.get(str_val, "")
                    if not name or name == str_val:
                        name = material_name_map.get(str_val, "")
                    if name and name != str_val:
                        str_val = name
                    else:
                        str_val = ""
                elif sc == "material_spec":
                    spec = material_spec_map.get(str_val, "")
                    if spec and spec != str_val:
                        str_val = spec
                    else:
                        str_val = ""
                elif sc == "unit_name":
                    if lang == "CN":
                        str_val = unit_name_map.get(str_val, str_val)
                elif sc == "cust_name":
                    str_val = cust_name_map.get(str_val, str_val)

            if data_col in date_cols and str_val:
                str_val = _format_date(str_val)

            output_row[tpl_idx] = _clean_string(str_val)

        ws_out.append(output_row)

    _apply_date_formats(ws_out, h0, date_cols, field_mapping)
    _save_workbook(workbook, output_path)
    print(f"  [完成] 文件大小: {output_path.stat().st_size / 1024 / 1024:.2f} MB")
    print(f"[完成] 生成预测单 Excel ({lang_label})")


def generate_planorder_excel(lang: str = "CN") -> None:
    import time as _time
    lang_label = "中文" if lang == "CN" else "英文"
    print(f"[开始] 生成计划订单 Excel ({lang_label})")
    OUTPUT_DIR.mkdir(exist_ok=True)

    template_suffix = "中文模板.xlsx" if lang == "CN" else "英文模板.xlsx"
    template_files = list(TEMPLATE_DIR.glob(f"{PLANORDER_FORM_ID}*{template_suffix}"))
    if not template_files:
        print(f"  [错误] 找不到计划订单{lang_label}模板文件")
        return

    template_path = template_files[0]
    print(f"  [模板] {template_path.name}")

    wb = openpyxl.load_workbook(template_path)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    h0 = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            h0 = [str(c).strip() if c else "" for c in row]
            break

    print(f"  [表头] {len(h0)} 列")

    data_path = _resolve_local_csv(PLANORDER_FORM_ID)
    if not data_path:
        print("  [错误] 找不到计划订单数据文件")
        return

    print(f"  [数据] {data_path.name}")
    df = pd.read_csv(data_path, encoding="utf-8-sig", dtype=str)
    print(f"  [数据] {len(df)} 行")
    data_cols = df.columns.tolist()
    print(f"  [数据] {len(df)} 行, {len(data_cols)} 列")

    # 只导出业务状态为"供应确认"（FReleaseStatus = 2）的计划订单
    if "FReleaseStatus" in data_cols:
        before_count = len(df)
        df = df[df["FReleaseStatus"] == "2"].copy()
        print(f"  [过滤] 业务状态=供应确认: {before_count} → {len(df)} 行")
    else:
        print("  [警告] 数据中缺少 FReleaseStatus 字段，无法过滤")

    date_cols = {"FPlanStartDate", "FFirmStartDate", "FPlanFinishDate", "FFirmFinishDate", "FDemandDate"}

    planorder_field_map = {
        "业务状态": "FReleaseStatus", "Business Status": "FReleaseStatus",
        "单据类型": "FBillTypeID", "Doc Type": "FBillTypeID",
        "单据编号": "FBillNo", "Doc No.": "FBillNo",
        "单据状态": "FDocumentStatus", "Doc Status": "FDocumentStatus",
        "投放类型": "FReleaseType", "Release Type": "FReleaseType",
        "投放单据类型": "FReleaseType", "Released Doc Type": "FReleaseType",
        "物料编码": "FMaterialId", "Material Code": "FMaterialId",
        "物料名称": "FMaterialId", "Material Name": "FMaterialId",
        "规格型号": "FMaterialId", "Specification": "FMaterialId",
        "辅助属性": "FAuxPropId", "Secondary Property": "FAuxPropId",
        "单位": "FUnitId", "UoM": "FUnitId",
        "确认订单量": "FFirmQty", "Confirmed Order Qty": "FFirmQty",
        "数据来源": "FDataSource", "Data Source": "FDataSource",
        "BOM版本": "FBomId", "BOM Version": "FBomId",
        "运算编号": "FComputerNo", "Operation No.": "FComputerNo",
        "采购/生产组织": "FSupplyOrgId", "Purchase/Production Org.": "FSupplyOrgId",
        "建议采购/生产日期": "FPlanStartDate", "Sugg. Pur./Prod. Dt": "FPlanStartDate",
        "确认采购/生产日期": "FFirmStartDate", "Cfm Pur./Prod. Dt.": "FFirmStartDate",
        "建议到货/完工日期": "FPlanFinishDate", "Sugg. Arrival/Finish Dt": "FPlanFinishDate",
        "确认到货/完工日期": "FFirmFinishDate", "Cfm Arrival/Finish Dt": "FFirmFinishDate",
        "需求组织": "FDemandOrgId", "Requirement Org.": "FDemandOrgId",
        "入库组织": "FInStockOrgId", "Receipt Org.": "FInStockOrgId",
        "净需求": "FDemandQty", "Net Requirement": "FDemandQty",
        "生产车间": "FPrdDeptId", "Production Workshop": "FPrdDeptId",
        "采购申请部门": "FPurDeptId", "Purchase Request Dept": "FPurDeptId",
        "计划订单量": "FOrderQty", "Planned Order Qty": "FOrderQty",
        "需求日期": "FDemandDate", "Requirement Date": "FDemandDate",
        "建议订单量": "FSugQty", "Proposed Order Qty": "FSugQty",
        "需求单据编号": "FSaleOrderNo", "Requirement Doc No.": "FSaleOrderNo",
        "需求单据行号": "FSaleOrderEntrySeq", "Row No. of Requirement Doc": "FSaleOrderEntrySeq",
        "需求来源": "FDemandType", "Requirement Source": "FDemandType",
        "物料属性": "FMaterialId", "Material property": "FMaterialId",
        "订货策略": "F_PALW_Text_qtr", "Order Policy": "F_PALW_Text_qtr",
        "排产识别码汇总": "FMaterialId", "Master Data Property": "FMaterialId",
        "国家": "F_PALW_Assistant_83g", "Country": "F_PALW_Assistant_83g",
        "需求客户": "F_PALW_Base_qtr", "Customer": "F_PALW_Base_qtr",
        "客户星级": "F_PALW_Base_qtr", "Customer Segment": "F_PALW_Base_qtr",
    }

    field_mapping = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if not tpl_field:
            continue
        if tpl_field in planorder_field_map:
            data_field = planorder_field_map[tpl_field]
            if data_field in data_cols:
                field_mapping[tpl_idx] = data_field

    print(f"  [映射] 找到 {len(field_mapping)} 个字段映射")

    client = KingdeeClient()

    def _collect_ids(col_name, id_type="numeric"):
        if col_name not in data_cols:
            return []
        ids = []
        for val in df[col_name].dropna().unique():
            id_str = str(val).strip()
            if not id_str or id_str in ("0", "0.0"):
                continue
            if id_type == "guid":
                if _is_guid(id_str):
                    ids.append(id_str)
            elif id_type == "numeric":
                if not _is_guid(id_str):
                    try:
                        num = int(float(id_str))
                        if num > 0:
                            ids.append(str(num))
                    except (ValueError, TypeError, OverflowError):
                        pass
            elif id_type == "mixed":
                has_chinese = any('\u4e00' <= c <= '\u9fff' for c in id_str)
                if not has_chinese:
                    ids.append(id_str)
            else:
                ids.append(id_str)
        return list(set(ids))

    print(f"  [解析] 开始解析计划订单ID字段...")

    material_ids = _collect_ids("FMaterialId")
    unit_ids = _collect_ids("FUnitId")
    base_unit_ids = _collect_ids("FBaseUnitId")
    org_ids = _collect_ids("FDemandOrgId")
    supply_org_ids = _collect_ids("FSupplyOrgId")
    instock_org_ids = _collect_ids("FInStockOrgId")
    bill_type_ids = _collect_ids("FBillTypeID", "guid")
    bom_ids = _collect_ids("FBomId")
    prd_dept_ids = _collect_ids("FPrdDeptId")
    pur_dept_ids = _collect_ids("FPurDeptId")
    aux_prop_ids = _collect_ids("FAuxPropId")
    country_ids = _collect_ids("F_PALW_Base")
    cust_base_ids = _collect_ids("F_PALW_Base_qtr")
    schedule_ids = _collect_ids("F_PALW_Assistant_83g", "mixed")

    all_org_ids = list(set(org_ids + supply_org_ids + instock_org_ids))
    all_unit_ids = list(set(unit_ids + base_unit_ids))

    if material_ids:
        print(f"    需要解析 {len(material_ids)} 个物料ID")
        material_code_map, material_name_map = fetch_material_mapping(client, material_ids)
    else:
        material_code_map, material_name_map = _load_material_mapping()

    material_code_to_name = {}
    for mid, code in material_code_map.items():
        name = material_name_map.get(mid, "")
        if code and name:
            material_code_to_name[code] = name

    missing_name_codes = [c for c in set(material_code_map.values()) if c not in material_code_to_name]
    if missing_name_codes:
        print(f"    需要解析 {len(missing_name_codes)} 个物料编码的中文名")
        batch_size = 200
        for i in range(0, len(missing_name_codes), batch_size):
            batch = missing_name_codes[i:i + batch_size]
            filter_parts = [f"FNumber = '{code}'" for code in batch]
            filter_str = " or ".join(filter_parts)
            try:
                rows = client.execute_bill_query({
                    "FormId": "BD_MATERIAL",
                    "FieldKeys": "FNumber,FName",
                    "FilterString": filter_str,
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": len(batch) + 10,
                })
                if rows:
                    for row in rows:
                        fnumber = str(row[0]).strip() if row[0] else ""
                        fname_raw = row[1] if row[1] else ""
                        fname = ""
                        if fname_raw:
                            try:
                                import json as _json
                                if isinstance(fname_raw, str) and fname_raw.startswith("["):
                                    name_list = _json.loads(fname_raw)
                                    if name_list and isinstance(name_list, list) and len(name_list) > 0:
                                        fname = name_list[0].get("Value", str(name_list[0])) if isinstance(name_list[0], dict) else str(name_list[0])
                                else:
                                    fname = str(fname_raw).strip()
                            except Exception:
                                fname = str(fname_raw).strip()
                        if fnumber and fname:
                            material_code_to_name[fnumber] = fname
            except Exception:
                pass

    material_spec_map = _load_generic_mapping("material_spec_map")
    missing_spec_codes = [c for c in set(material_code_map.values()) if c not in material_spec_map]
    if missing_spec_codes:
        print(f"    需要解析 {len(missing_spec_codes)} 个物料编码的规格型号")
        _fetch_batch_spec(client, missing_spec_codes, material_spec_map)
        _save_generic_mapping("material_spec_map", material_spec_map)

    # 物料属性映射（Purchase/Self-made）：从BD_MATERIAL的FIsPurchase/FIsProduce获取
    # 只查询当前数据中实际出现的物料编码
    _planorder_material_codes = set()
    if "FMaterialId" in data_cols:
        for mid in df["FMaterialId"].dropna().unique():
            s = str(mid).strip()
            if s and s not in ("0", "0.0"):
                code = material_code_map.get(s, "")
                if code:
                    _planorder_material_codes.add(code)
    material_attr_map = _load_generic_mapping("material_attr_map")
    missing_attr_codes = [c for c in _planorder_material_codes if c and c not in material_attr_map]
    if missing_attr_codes:
        print(f"    需要解析 {len(missing_attr_codes)} 个物料编码的物料属性")
        batch_size = 50
        for i in range(0, len(missing_attr_codes), batch_size):
            batch = missing_attr_codes[i:i + batch_size]
            filter_parts = [f"FNumber = '{code}'" for code in batch]
            filter_str = " or ".join(filter_parts)
            try:
                rows = client.execute_bill_query({
                    "FormId": "BD_MATERIAL",
                    "FieldKeys": "FNumber,FIsPurchase,FIsProduce",
                    "FilterString": filter_str,
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": len(batch) + 10,
                })
                if rows:
                    for row in rows:
                        code = str(row[0]).strip() if row[0] else ""
                        is_purchase = row[1] if len(row) > 1 else False
                        is_produce = row[2] if len(row) > 2 else False
                        if code:
                            if is_produce:
                                material_attr_map[code] = "Self-made"
                            elif is_purchase:
                                material_attr_map[code] = "Purchase"
                            else:
                                material_attr_map[code] = ""
            except Exception:
                pass
            _time.sleep(0.2)
        _save_generic_mapping("material_attr_map", material_attr_map)
    print(f"    物料属性映射: {len(material_attr_map)} 个")

    # 物料排产代码映射：从BOM TreeEntity的F_PALW_Text_83g获取（如V/D/M等），汇总显示
    material_schedule_code_map = _load_generic_mapping("material_schedule_code_map")
    # 获取当前数据中涉及的BOM ID
    bom_ids_in_data = set()
    if "FBomId" in data_cols:
        for bid in df["FBomId"].dropna().unique():
            s = str(bid).strip()
            if s and s not in ("0", "0.0"):
                bom_ids_in_data.add(s)
    missing_sched_boms = [b for b in bom_ids_in_data if b not in material_schedule_code_map]
    if missing_sched_boms:
        print(f"    需要解析 {len(missing_sched_boms)} 个BOM的排产代码")
        for bom_id in missing_sched_boms:
            try:
                bid_int = int(float(bom_id))
                result = client.view("ENG_BOM", {"Id": bid_int, "Number": "", "CreateOrgId": 0})
                inner = result.get("Result", {}).get("Result", {})
                tree = inner.get("TreeEntity", [])
                # 收集所有子项的F_PALW_Text_83g值，去重后用逗号连接
                sched_vals = set()
                for entry in tree:
                    txt83g = entry.get("F_PALW_Text_83g", "")
                    if txt83g and str(txt83g).strip():
                        sched_vals.add(str(txt83g).strip())
                if sched_vals:
                    # 按字母排序后用逗号连接
                    material_schedule_code_map[bom_id] = ",".join(sorted(sched_vals))
                else:
                    material_schedule_code_map[bom_id] = ""
            except Exception:
                material_schedule_code_map[bom_id] = ""
            _time.sleep(0.2)
        _save_generic_mapping("material_schedule_code_map", material_schedule_code_map)
    print(f"    物料排产代码映射: {len(material_schedule_code_map)} 个")

    if all_unit_ids:
        print(f"    需要解析 {len(all_unit_ids)} 个单位ID")
        unit_code_map = fetch_generic_mapping(client, "BD_UNIT", "FUnitID", "FNumber", "unit_code_map", all_unit_ids)
    else:
        unit_code_map = _load_generic_mapping("unit_code_map")
    unit_name_map = _fetch_name_map(client, "BD_UNIT", list(set(unit_code_map.values())), "unit_name_map")

    if all_org_ids:
        print(f"    需要解析 {len(all_org_ids)} 个组织ID")
        org_code_map = fetch_generic_mapping(client, "ORG_Organizations", "FOrgId", "FNumber", "org_code_map", all_org_ids)
    else:
        org_code_map = _load_generic_mapping("org_code_map")
    org_name_map = _fetch_name_map(client, "ORG_Organizations", list(set(org_code_map.values())), "org_name_map")

    if bill_type_ids:
        print(f"    需要解析 {len(bill_type_ids)} 个单据类型ID(GUID)")
        bill_type_code_map = fetch_generic_mapping(client, "BOS_BillType", "FBILLTYPEID", "FNumber", "bill_type_map", bill_type_ids)
        guid_name_map = fetch_generic_mapping(client, "BOS_BillType", "FBILLTYPEID", "FName", "bill_type_name_map_guid", bill_type_ids)
        bill_type_name_map = {}
        for gid, name in guid_name_map.items():
            code = bill_type_code_map.get(gid, "")
            if code and name:
                bill_type_name_map[code] = name
            if gid and name:
                bill_type_name_map[gid] = name
    else:
        bill_type_code_map = _load_generic_mapping("bill_type_map")
        bill_type_name_map = {}
    bill_type_name_map.update(_load_generic_mapping("bill_type_name_map"))

    bill_type_codes = []
    if "FBillTypeID" in data_cols:
        for val in df["FBillTypeID"].dropna().unique():
            s = str(val).strip()
            if s and s != "0" and not _is_guid(s):
                bill_type_codes.append(s)
    bill_type_codes = list(set(bill_type_codes))
    missing_bt = [c for c in bill_type_codes if c not in bill_type_name_map]
    if missing_bt:
        print(f"    需要解析 {len(missing_bt)} 个单据类型编码的中文名")
        for code in missing_bt:
            try:
                rows = client.execute_bill_query({
                    "FormId": "BOS_BillType",
                    "FieldKeys": "FNumber,FName",
                    "FilterString": f"FNumber = '{code}'",
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": 10,
                })
                if rows:
                    for row in rows:
                        fnumber = str(row[0]).strip() if row[0] else ""
                        fname = str(row[1]).strip() if row[1] else ""
                        if fnumber and fname:
                            bill_type_name_map[fnumber] = fname
            except Exception:
                pass
        _save_generic_mapping("bill_type_name_map", bill_type_name_map)

    bom_code_map = {}
    if bom_ids:
        print(f"    需要解析 {len(bom_ids)} 个BOM ID")
        bom_code_map = fetch_generic_mapping(client, "ENG_BOM", "FId", "FNumber", "bom_code_map", bom_ids)

    dept_code_map = {}
    dept_name_map = {}
    all_dept_ids = list(set(prd_dept_ids + pur_dept_ids))
    if all_dept_ids:
        print(f"    需要解析 {len(all_dept_ids)} 个部门ID")
        dept_code_map = fetch_generic_mapping(client, "BD_Department", "FDeptID", "FNumber", "dept_code_map", all_dept_ids)
        dept_name_map = _fetch_name_map(client, "BD_Department", list(set(dept_code_map.values())), "dept_name_map")

    # F_PALW_Base字段标签为"国家"，但实际存储的是员工ID（BD_Empinfo）
    country_code_map_local = {}
    if country_ids:
        country_code_map_local = fetch_generic_mapping(client, "BD_Empinfo", "FID", "FNumber", "emp_code_map", country_ids)

    cust_base_code_map_local = {}
    if cust_base_ids:
        cust_base_code_map_local = fetch_generic_mapping(client, "BD_Customer", "FCUSTID", "FNumber", "cust_code_map", cust_base_ids)

    country_name_map = {}
    if country_ids:
        print(f"    需要解析 {len(country_ids)} 个国家(员工)ID")
        emp_codes = list(set(country_code_map_local.values()))
        if emp_codes:
            country_name_map = _load_generic_mapping("emp_name_map")
            missing_emp_names = [c for c in emp_codes if c not in country_name_map]
            if missing_emp_names:
                print(f"    需要解析 {len(missing_emp_names)} 个员工编码的中文名")
                batch_size = 200
                for i in range(0, len(missing_emp_names), batch_size):
                    batch = missing_emp_names[i:i + batch_size]
                    filter_parts = [f"FNumber = '{code}'" for code in batch]
                    filter_str = " or ".join(filter_parts)
                    try:
                        rows = client.execute_bill_query({
                            "FormId": "BD_Empinfo",
                            "FieldKeys": "FNumber,FName",
                            "FilterString": filter_str,
                            "OrderString": "",
                            "TopRowCount": 0,
                            "StartRow": 0,
                            "Limit": len(batch) + 10,
                        })
                        if rows:
                            for r in rows:
                                fnumber = str(r[0]).strip() if r[0] else ""
                                fname_raw = r[1] if r[1] else ""
                                fname = ""
                                if fname_raw:
                                    try:
                                        import json as _json
                                        if isinstance(fname_raw, str) and fname_raw.startswith("["):
                                            name_list = _json.loads(fname_raw)
                                            if name_list and isinstance(name_list, list) and len(name_list) > 0:
                                                fname = name_list[0].get("Value", str(name_list[0])) if isinstance(name_list[0], dict) else str(name_list[0])
                                        else:
                                            fname = str(fname_raw).strip()
                                    except Exception:
                                        fname = str(fname_raw).strip()
                                if fnumber and fname:
                                    country_name_map[fnumber] = fname
                    except Exception:
                        pass
                _save_generic_mapping("emp_name_map", country_name_map)

    cust_base_name_map = {}
    if cust_base_ids:
        print(f"    需要解析 {len(cust_base_ids)} 个需求客户ID")
        cust_base_codes = list(set(cust_base_code_map_local.values()))
        cust_base_name_map = _load_generic_mapping("cust_name_map")
        missing_cust_names = [c for c in cust_base_codes if c not in cust_base_name_map]
        if missing_cust_names:
            print(f"    需要解析 {len(missing_cust_names)} 个客户编码的中文名")
            batch_size = 200
            for i in range(0, len(missing_cust_names), batch_size):
                batch = missing_cust_names[i:i + batch_size]
                filter_parts = [f"FNumber = '{code}'" for code in batch]
                filter_str = " or ".join(filter_parts)
                try:
                    rows = client.execute_bill_query({
                        "FormId": "BD_Customer",
                        "FieldKeys": "FNumber,FName",
                        "FilterString": filter_str,
                        "OrderString": "",
                        "TopRowCount": 0,
                        "StartRow": 0,
                        "Limit": len(batch) + 10,
                    })
                    if rows:
                        for r in rows:
                            fnumber = str(r[0]).strip() if r[0] else ""
                            fname_raw = r[1] if r[1] else ""
                            fname = ""
                            if fname_raw:
                                try:
                                    import json as _json
                                    if isinstance(fname_raw, str) and fname_raw.startswith("["):
                                        name_list = _json.loads(fname_raw)
                                        if name_list and isinstance(name_list, list) and len(name_list) > 0:
                                            fname = name_list[0].get("Value", str(name_list[0])) if isinstance(name_list[0], dict) else str(name_list[0])
                                    else:
                                        fname = str(fname_raw).strip()
                                except Exception:
                                    fname = str(fname_raw).strip()
                            if fnumber and fname:
                                cust_base_name_map[fnumber] = fname
                except Exception:
                    pass
            _save_generic_mapping("cust_name_map", cust_base_name_map)

    # 客户星级映射：从BD_Customer的F_PALW_Assistant_83g获取Star/Contributor/Questionmark
    # 直接用Number映射：01=Star, 02=Contributor, 04=Questionmark
    _cust_seg_number_map = {"01": "Star", "02": "Contributor", "04": "Questionmark"}
    customer_segment_map = {}
    if cust_base_ids:
        customer_segment_map = _load_generic_mapping("customer_segment_map")
        # 获取客户编码列表
        cust_codes_for_seg = list(set(cust_base_code_map_local.values()))
        missing_seg_codes = [c for c in cust_codes_for_seg if c and c not in customer_segment_map]
        if missing_seg_codes:
            print(f"    需要解析 {len(missing_seg_codes)} 个客户编码的客户星级")
            batch_size = 50
            for i in range(0, len(missing_seg_codes), batch_size):
                batch = missing_seg_codes[i:i + batch_size]
                filter_parts = [f"FNumber = '{code}'" for code in batch]
                filter_str = " or ".join(filter_parts)
                try:
                    rows = client.execute_bill_query({
                        "FormId": "BD_Customer",
                        "FieldKeys": "FNumber,F_PALW_Assistant_83g.FNumber",
                        "FilterString": filter_str,
                        "OrderString": "",
                        "TopRowCount": 0,
                        "StartRow": 0,
                        "Limit": len(batch) + 10,
                    })
                    if rows:
                        for r in rows:
                            fnumber = str(r[0]).strip() if r[0] else ""
                            seg_number = str(r[1]).strip() if len(r) > 1 and r[1] else ""
                            if fnumber and seg_number:
                                seg_val = _cust_seg_number_map.get(seg_number, "")
                                if seg_val:
                                    customer_segment_map[fnumber] = seg_val
                except Exception:
                    pass
                _time.sleep(0.2)
            _save_generic_mapping("customer_segment_map", customer_segment_map)
        print(f"    客户星级映射: {len(customer_segment_map)} 个")

    schedule_name_map = {}
    if schedule_ids:
        print(f"    需要解析 {len(schedule_ids)} 个排产识别码ID")
        schedule_name_map = _load_generic_mapping("schedule_name_map")
        missing_schedule = [sid for sid in schedule_ids if sid not in schedule_name_map]
        if missing_schedule:
            print(f"    需要查询 {len(missing_schedule)} 个排产识别码名称")
            for sid in missing_schedule:
                try:
                    result = client.view("BOS_ASSISTANTDATA_Detail", {"Id": sid, "Number": "", "CreateOrgId": 0})
                    inner = result.get("Result", {}).get("Result", {})
                    if inner:
                        fnumber = inner.get("Number", "") or inner.get("FNumber", "")
                        fdata_value = inner.get("FDataValue", "")
                        name = ""
                        if isinstance(fdata_value, list) and len(fdata_value) > 0:
                            for item in fdata_value:
                                if isinstance(item, dict) and item.get("Key") == 2052:
                                    name = item.get("Value", "")
                                    break
                            if not name and fdata_value:
                                first = fdata_value[0]
                                name = first.get("Value", "") if isinstance(first, dict) else str(first)
                        elif isinstance(fdata_value, str):
                            name = fdata_value
                        if not name:
                            name = fnumber
                        if name:
                            schedule_name_map[sid] = name
                except Exception:
                    pass
            _save_generic_mapping("schedule_name_map", schedule_name_map)

    # 国家简写映射：从PLN_PLANORDER的F_PALW_Assistant_83g获取Number（如RU/AR/BD）
    country_abbr_map = {}
    if schedule_ids:
        country_abbr_map = _load_generic_mapping("country_abbr_map")
        missing_abbr = [sid for sid in schedule_ids if sid not in country_abbr_map]
        if missing_abbr:
            print(f"    需要查询 {len(missing_abbr)} 个国家简写")
            for sid in missing_abbr:
                try:
                    sid_str = str(sid).strip()
                    try:
                        sid_int = int(float(sid_str))
                        view_params = {"Id": sid_int, "Number": "", "CreateOrgId": 0}
                    except (ValueError, OverflowError):
                        view_params = {"Id": sid_str, "Number": "", "CreateOrgId": 0}
                    result = client.view("BOS_ASSISTANTDATA_Detail", view_params)
                    inner = result.get("Result", {}).get("Result", {})
                    if inner:
                        fnumber = inner.get("Number", "") or inner.get("FNumber", "")
                        if fnumber:
                            country_abbr_map[sid] = fnumber
                except Exception:
                    pass
            _save_generic_mapping("country_abbr_map", country_abbr_map)

    if lang == "CN":
        status_map = {
            "FDocumentStatus": {"A": "已审核", "B": "暂存", "C": "已审核", "D": "重新审核"},
            "FReleaseType": {"1": "生产订单类", "2": "采购申请类", "3": "委外订单类"},
            "FReleaseStatus": {"0": "计划", "1": "需求确认", "2": "供应确认", "3": "手工投放关闭", "4": "手工关闭", "5": "自动投放关闭", "6": "合并关闭", "7": "拆分关闭", "8": "自动关闭"},
            "FDataSource": {"1": "手工录入", "2": "MRP运算", "3": "MPS运算"},
            "FDemandType": {"1": "预测", "2": "销售订单", "7": "其他"},
        }
    else:
        status_map = {
            "FDocumentStatus": {"A": "Approved", "B": "Draft", "C": "Approved", "D": "Re-audit"},
            "FReleaseType": {"1": "MO Type", "2": "Purchase Request Type", "3": "Subcontract Order Type"},
            "FReleaseStatus": {"0": "Plan", "1": "Requirement Confirmation", "2": "Supply Confirmation", "3": "Manually Release Closed", "4": "Manually Closed", "5": "Auto Release Closed", "6": "Merged and Closed", "7": "Split and Closed", "8": "Auto Closed"},
            "FDataSource": {"1": "Manually Input", "2": "MRP Calculation", "3": "MPS Calculation"},
            "FDemandType": {"1": "Forecast", "2": "Sales Order", "7": "Other", "": "Forecast"},
        }

    id_to_code_map = {
        "FBillTypeID": bill_type_code_map,
        "FMaterialId": material_code_map,
        "FUnitId": unit_code_map,
        "FBaseUnitId": unit_code_map,
        "FDemandOrgId": org_code_map,
        "FSupplyOrgId": org_code_map,
        "FInStockOrgId": org_code_map,
        "FBomId": bom_code_map,
        "FPrdDeptId": dept_code_map,
        "FPurDeptId": dept_code_map,
        "F_PALW_Base": country_code_map_local,
        "F_PALW_Base_qtr": cust_base_code_map_local,
        "F_PALW_Assistant_83g": country_abbr_map if country_abbr_map else {},
    }

    zero_to_empty_cols = {"FBomId", "FAuxPropId", "FPrdDeptId", "FPurDeptId", "FSupplierId", "FPlanerId", "FBuyerId", "F_PALW_Base", "F_PALW_Base_qtr", "F_PALW_Decimal_uky"}

    special_cols = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if tpl_field in ("单据类型", "Doc Type"):
            special_cols[tpl_idx] = "bill_type_name"
        elif tpl_field in ("采购/生产组织", "需求组织", "入库组织", "Purchase/Production Org.", "Requirement Org.", "Receipt Org."):
            special_cols[tpl_idx] = "org_name"
        elif tpl_field in ("投放单据类型", "Released Doc Type"):
            special_cols[tpl_idx] = "release_type_name"
        elif tpl_field in ("物料名称", "Material Name"):
            special_cols[tpl_idx] = "material_name"
        elif tpl_field in ("规格型号", "Specification"):
            special_cols[tpl_idx] = "material_spec"
        elif tpl_field in ("单位", "UoM"):
            special_cols[tpl_idx] = "unit_name"
        elif tpl_field in ("物料属性", "Material property"):
            special_cols[tpl_idx] = "material_attr"
        elif tpl_field in ("生产车间", "Production Workshop"):
            special_cols[tpl_idx] = "dept_name"
        elif tpl_field in ("采购申请部门", "Purchase Request Dept"):
            special_cols[tpl_idx] = "dept_name"
        elif tpl_field in ("国家", "Country"):
            special_cols[tpl_idx] = "country_name"
        elif tpl_field in ("需求客户", "Customer"):
            special_cols[tpl_idx] = "cust_base_name"
        elif tpl_field in ("排产识别码汇总", "Master Data Property"):
            special_cols[tpl_idx] = "schedule_name"
        elif tpl_field in ("订货策略", "Order Policy"):
            special_cols[tpl_idx] = "order_policy_lfl"
        elif tpl_field in ("客户星级", "Customer Segment"):
            special_cols[tpl_idx] = "cust_segment"

    if lang == "CN":
        output_name = "计划订单列表.xlsx"
        sheet_title = "Sheet1"
    else:
        output_name = "K3_PlannedOrders.xlsx"
        sheet_title = "Sheet1"
    output_path = OUTPUT_DIR / output_name
    for old_file in OUTPUT_DIR.glob(f"{output_name.split('.')[0]}*.xlsx"):
        try:
            old_file.unlink()
        except PermissionError:
            print(f"  [警告] 文件被占用，跳过: {old_file.name}")

    print(f"  [输出] {output_path.name}")

    workbook = openpyxl.Workbook()
    ws_out = workbook.active
    ws_out.title = sheet_title

    ws_out.append(h0)

    for _, row in df.iterrows():
        output_row = [""] * len(h0)

        for tpl_idx, data_col in field_mapping.items():
            val = row[data_col]
            is_empty = pd.isna(val) or str(val).strip() == ""
            is_zero = not is_empty and str(val).strip() in ("0", "0.0") and data_col in zero_to_empty_cols

            # 对于某些special_cols字段，优先处理（如固定值LFL、映射查找等）
            sc = special_cols.get(tpl_idx, "")
            if sc == "order_policy_lfl":
                output_row[tpl_idx] = "LFL"
                continue
            if sc == "material_attr":
                if not is_empty and not is_zero:
                    str_val = str(val).strip()
                    if data_col in id_to_code_map and id_to_code_map[data_col]:
                        str_val = _resolve_id(str_val, id_to_code_map[data_col])
                    attr_val = material_attr_map.get(str_val, "")
                    output_row[tpl_idx] = _clean_string(attr_val)
                continue
            if sc == "schedule_name":
                # 排产代码从BOM获取，key是BOM ID
                bom_id_raw = row.get("FBomId", "")
                if pd.isna(bom_id_raw):
                    bom_id_val = ""
                else:
                    bom_id_val = str(bom_id_raw).strip()
                if bom_id_val and bom_id_val not in ("0", "0.0", "nan"):
                    sched_code = material_schedule_code_map.get(bom_id_val, "")
                    output_row[tpl_idx] = _clean_string(sched_code)
                continue
            if sc == "cust_segment":
                if not is_empty and not is_zero:
                    str_val = str(val).strip()
                    if data_col in id_to_code_map and id_to_code_map[data_col]:
                        str_val = _resolve_id(str_val, id_to_code_map[data_col])
                    seg_val = customer_segment_map.get(str_val, "")
                    output_row[tpl_idx] = _clean_string(seg_val)
                continue

            if is_empty:
                # FDemandType空值需要映射为Forecast
                if data_col == "FDemandType":
                    output_row[tpl_idx] = "Forecast" if lang != "CN" else "预测"
                continue

            if is_zero:
                continue

            str_val = str(val).strip()

            if data_col in id_to_code_map and id_to_code_map[data_col]:
                str_val = _resolve_id(str_val, id_to_code_map[data_col])

            if data_col in status_map:
                str_val = status_map[data_col].get(str_val, str_val)

            if data_col in date_cols and str_val:
                str_val = _format_date(str_val)

            if tpl_idx in special_cols:
                sc = special_cols[tpl_idx]
                if sc == "bill_type_name":
                    str_val = bill_type_name_map.get(str_val, str_val)
                elif sc == "org_name":
                    str_val = org_name_map.get(str_val, str_val)
                elif sc == "release_type_name":
                    if lang == "CN":
                        release_names = {"生产订单类": "汇报入库-普通生产", "采购申请类": "采购申请-标准采购", "委外订单类": "委外订单-委外生产"}
                    else:
                        release_names = {"MO Type": "Report Receipt - Common Production", "Purchase Request Type": "Purchase Request - Standard Purchase", "Subcontract Order Type": "Subcontract Order - Subcontract Production"}
                    str_val = release_names.get(str_val, str_val)
                elif sc == "material_name":
                    name = material_code_to_name.get(str_val, "")
                    if not name or name == str_val:
                        name = material_name_map.get(str_val, "")
                    if name and name != str_val:
                        str_val = name
                    else:
                        str_val = ""
                elif sc == "material_spec":
                    spec = material_spec_map.get(str_val, "")
                    if spec and spec != str_val:
                        str_val = spec
                    else:
                        str_val = ""
                elif sc == "unit_name":
                    if lang == "CN":
                        str_val = unit_name_map.get(str_val, str_val)
                elif sc == "dept_name":
                    str_val = dept_name_map.get(str_val, str_val)
                elif sc == "country_name":
                    # _resolve_id已经把F_PALW_Assistant_83g的十六进制ID转成国家简写(如RU/BD)
                    # str_val此时已经是国家简写，无需再查country_abbr_map
                    pass
                elif sc == "cust_base_name":
                    name = cust_base_name_map.get(str_val, "")
                    str_val = name if name else str_val

            output_row[tpl_idx] = _clean_string(str_val)

        ws_out.append(output_row)

    _apply_date_formats(ws_out, h0, date_cols, field_mapping)
    _save_workbook(workbook, output_path)
    print(f"  [完成] 文件大小: {output_path.stat().st_size / 1024 / 1024:.2f} MB")
    print(f"[完成] 生成计划订单 Excel ({lang_label})")


def generate_mo_excel(lang: str = "CN") -> None:
    lang_label = "中文" if lang == "CN" else "英文"
    print(f"[开始] 生成生产订单 Excel ({lang_label})")
    OUTPUT_DIR.mkdir(exist_ok=True)

    template_suffix = "中文模板.xlsx" if lang == "CN" else "英文模板.xlsx"
    template_files = list(TEMPLATE_DIR.glob(f"{MO_FORM_ID}*{template_suffix}"))
    if not template_files:
        print(f"  [错误] 找不到生产订单{lang_label}模板文件")
        return

    template_path = template_files[0]
    print(f"  [模板] {template_path.name}")

    wb = openpyxl.load_workbook(template_path)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    h0 = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            h0 = [str(c).strip() if c else "" for c in row]
            break

    print(f"  [表头] {len(h0)} 列")

    data_path = _resolve_local_csv(MO_FORM_ID)
    if not data_path:
        print("  [错误] 找不到生产订单数据文件")
        return

    print(f"  [数据] {data_path.name}")
    df = pd.read_csv(data_path, encoding="utf-8-sig", dtype=str)
    data_cols = df.columns.tolist()
    print(f"  [数据] {len(df)} 行, {len(data_cols)} 列")

    date_cols = {"FDate", "FPlanStartDate", "FPlanFinishDate", "FStartDate", "FFinishDate", "FCreateDate", "FApproveDate"}

    mo_field_map = {
        "单据编号": "FBillNo", "Doc No.": "FBillNo",
        "单据类型": "FBillType", "Doc Type": "FBillType",
        "单据日期": "FDate", "Doc Date": "FDate",
        "单据状态": "FDocumentStatus", "Doc Status": "FDocumentStatus",
        "产品类型": "FProductType", "Product Type": "FProductType",
        "物料编码": "FMaterialId", "Material Code": "FMaterialId",
        "物料名称": "FMaterialId", "Material Name": "FMaterialId",
        "规格型号": "FMaterialId", "Specification": "FMaterialId",
        "生产车间": "FWorkShopID", "Production Workshop": "FWorkShopID",
        "单位": "FUnitId", "UoM": "FUnitId",
        "数量": "FQty", "Qty": "FQty",
        "业务状态": "FStatus", "Business Status": "FStatus",
        "计划开工时间": "FPlanStartDate", "Plan Start Time": "FPlanStartDate",
        "计划完工时间": "FPlanFinishDate", "Plan Finish Time": "FPlanFinishDate",
        "BOM版本": "FBomId", "BOM Version": "FBomId",
        "开工日期": "FStartDate", "Start Date": "FStartDate",
        "完工日期": "FFinishDate", "Finish Date": "FFinishDate",
        "生成方式": "FCreateType", "Generation Method": "FCreateType",
        "源单类型": "FSrcBillType", "Source Doc Type": "FSrcBillType",
        "源单编号": "FSrcBillNo", "Source Doc No.": "FSrcBillNo",
        "辅助属性": "FAuxPropId", "Secondary Property": "FAuxPropId",
        "领料状态": "FPickMtrlStatus", "Material Picking Status": "FPickMtrlStatus",
        "成品批号": "FLot", "Lot No": "FLot",
    }

    field_mapping = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if not tpl_field:
            continue
        if tpl_field in mo_field_map:
            data_field = mo_field_map[tpl_field]
            if data_field in data_cols:
                field_mapping[tpl_idx] = data_field

    print(f"  [映射] 找到 {len(field_mapping)} 个字段映射")

    client = KingdeeClient()

    def _collect_ids(col_name, id_type="numeric"):
        if col_name not in data_cols:
            return []
        ids = []
        for val in df[col_name].dropna().unique():
            id_str = str(val).strip()
            if not id_str or id_str in ("0", "0.0"):
                continue
            if id_type == "guid":
                if _is_guid(id_str):
                    ids.append(id_str)
            elif id_type == "numeric":
                if not _is_guid(id_str):
                    try:
                        num = int(float(id_str))
                        if num > 0:
                            ids.append(str(num))
                    except (ValueError, TypeError, OverflowError):
                        pass
            else:
                ids.append(id_str)
        return list(set(ids))

    print(f"  [解析] 开始解析生产订单ID字段...")

    material_ids = _collect_ids("FMaterialId")
    unit_ids = _collect_ids("FUnitId")
    org_ids = _collect_ids("FPrdOrgId")
    workshop_ids = _collect_ids("FWorkShopID")
    bom_ids = _collect_ids("FBomId")
    bill_type_ids = _collect_ids("FBillType", "guid")
    aux_prop_ids = _collect_ids("FAuxPropId")
    lot_ids = _collect_ids("FLot")

    # 物料映射
    if material_ids:
        print(f"    需要解析 {len(material_ids)} 个物料ID")
        material_code_map, material_name_map = fetch_material_mapping(client, material_ids)
    else:
        material_code_map, material_name_map = _load_material_mapping()

    material_code_to_name = {}
    for mid, code in material_code_map.items():
        name = material_name_map.get(mid, "")
        if code and name:
            material_code_to_name[code] = name

    missing_name_codes = [c for c in set(material_code_map.values()) if c not in material_code_to_name]
    if missing_name_codes:
        print(f"    需要解析 {len(missing_name_codes)} 个物料编码的中文名")
        batch_size = 200
        for i in range(0, len(missing_name_codes), batch_size):
            batch = missing_name_codes[i:i + batch_size]
            filter_parts = [f"FNumber = '{code}'" for code in batch]
            filter_str = " or ".join(filter_parts)
            try:
                rows = client.execute_bill_query({
                    "FormId": "BD_MATERIAL",
                    "FieldKeys": "FNumber,FName",
                    "FilterString": filter_str,
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": len(batch) + 10,
                })
                if rows:
                    for row in rows:
                        fnumber = str(row[0]).strip() if row[0] else ""
                        fname_raw = row[1] if row[1] else ""
                        fname = ""
                        if fname_raw:
                            try:
                                import json as _json
                                if isinstance(fname_raw, str) and fname_raw.startswith("["):
                                    name_list = _json.loads(fname_raw)
                                    if name_list and isinstance(name_list, list) and len(name_list) > 0:
                                        fname = name_list[0].get("Value", str(name_list[0])) if isinstance(name_list[0], dict) else str(name_list[0])
                                else:
                                    fname = str(fname_raw).strip()
                            except Exception:
                                fname = str(fname_raw).strip()
                        if fnumber and fname:
                            material_code_to_name[fnumber] = fname
            except Exception:
                pass

    material_spec_map = _load_generic_mapping("material_spec_map")
    missing_spec_codes = [c for c in set(material_code_map.values()) if c not in material_spec_map]
    if missing_spec_codes:
        print(f"    需要解析 {len(missing_spec_codes)} 个物料编码的规格型号")
        _fetch_batch_spec(client, missing_spec_codes, material_spec_map)
        _save_generic_mapping("material_spec_map", material_spec_map)

    # 单位映射
    if unit_ids:
        print(f"    需要解析 {len(unit_ids)} 个单位ID")
        unit_code_map = fetch_generic_mapping(client, "BD_UNIT", "FUnitID", "FNumber", "unit_code_map", unit_ids)
    else:
        unit_code_map = _load_generic_mapping("unit_code_map")
    unit_name_map = _fetch_name_map(client, "BD_UNIT", list(set(unit_code_map.values())), "unit_name_map")

    # 组织映射
    if org_ids:
        print(f"    需要解析 {len(org_ids)} 个组织ID")
        org_code_map = fetch_generic_mapping(client, "ORG_Organizations", "FOrgId", "FNumber", "org_code_map", org_ids)
    else:
        org_code_map = _load_generic_mapping("org_code_map")
    org_name_map = _fetch_name_map(client, "ORG_Organizations", list(set(org_code_map.values())), "org_name_map")

    # 生产车间映射
    dept_code_map = {}
    dept_name_map = {}
    if workshop_ids:
        print(f"    需要解析 {len(workshop_ids)} 个生产车间ID")
        dept_code_map = fetch_generic_mapping(client, "BD_Department", "FDeptID", "FNumber", "dept_code_map", workshop_ids)
        dept_name_map = _fetch_name_map(client, "BD_Department", list(set(dept_code_map.values())), "dept_name_map")

    # BOM映射
    bom_code_map = {}
    if bom_ids:
        print(f"    需要解析 {len(bom_ids)} 个BOM ID")
        bom_code_map = fetch_generic_mapping(client, "ENG_BOM", "FId", "FNumber", "bom_code_map", bom_ids)

    # 批号映射
    lot_name_map = {}
    if lot_ids:
        print(f"    需要解析 {len(lot_ids)} 个批号ID")
        lot_name_map = _load_generic_mapping("inv_lot_name_map")
        # 兼容旧格式：如果值是dict则转为字符串
        for k, v in list(lot_name_map.items()):
            if isinstance(v, dict):
                lot_name_map[k] = v.get("cn", str(v))
        missing_lots = [lid for lid in lot_ids if lid not in lot_name_map]
        if missing_lots:
            print(f"    需要查询 {len(missing_lots)} 个批号名称")
            batch_size = 200
            for i in range(0, len(missing_lots), batch_size):
                batch = missing_lots[i:i + batch_size]
                filter_parts = [f"FLot = {lid}" for lid in batch]
                filter_str = " or ".join(filter_parts)
                try:
                    rows = client.execute_bill_query({
                        "FormId": "STK_Inventory",
                        "FieldKeys": "FLot,FLot.FNumber",
                        "FilterString": filter_str,
                        "OrderString": "",
                        "TopRowCount": 0,
                        "StartRow": 0,
                        "Limit": len(batch) + 10,
                    })
                    if rows:
                        for row in rows:
                            lot_id = str(row[0]).strip() if row[0] else ""
                            lot_number = str(row[1]).strip() if row[1] else ""
                            if lot_id and lot_number and lot_number not in ("None", "0"):
                                lot_name_map[lot_id] = lot_number
                except Exception:
                    pass
            _save_generic_mapping("inv_lot_name_map", lot_name_map)

    # 单据类型映射
    bill_type_code_map = {}
    bill_type_name_map = {}
    if bill_type_ids:
        print(f"    需要解析 {len(bill_type_ids)} 个单据类型ID(GUID)")
        bill_type_code_map = fetch_generic_mapping(client, "BOS_BillType", "FBILLTYPEID", "FNumber", "bill_type_map", bill_type_ids)
        guid_name_map = fetch_generic_mapping(client, "BOS_BillType", "FBILLTYPEID", "FName", "bill_type_name_map_guid", bill_type_ids)
        for gid, name in guid_name_map.items():
            code = bill_type_code_map.get(gid, "")
            if code and name:
                bill_type_name_map[code] = name
            if gid and name:
                bill_type_name_map[gid] = name
    else:
        bill_type_code_map = _load_generic_mapping("bill_type_map")
    bill_type_name_map.update(_load_generic_mapping("bill_type_name_map"))

    # 状态映射
    if lang == "CN":
        status_map = {
            "FDocumentStatus": {"A": "已审核", "B": "暂存", "C": "已审核", "D": "重新审核"},
            "FProductType": {"1": "主产品", "2": "副产品", "3": "联产品"},
            "FStatus": {"1": "计划", "2": "计划", "3": "下达", "4": "开工", "5": "完工", "6": "关闭", "7": "结算"},
            "FCreateType": {"1": "手工录入", "2": "MRP生成", "3": "MPS生成", "4": "计划订单投放", "5": "委外订单投放", "6": "拆分生成", "7": "销售订单下推"},
            "FPickMtrlStatus": {"0": "未领料", "1": "部分领料", "2": "全部领料", "3": "超额领料", "4": "待确认"},
        }
    else:
        status_map = {
            "FDocumentStatus": {"A": "Approved", "B": "Draft", "C": "Approved", "D": "Re-audit"},
            "FProductType": {"1": "Main Product", "2": "By-product", "3": "Co-product"},
            "FStatus": {"1": "Planned", "2": "Planned", "3": "Released", "4": "Started", "5": "Finished", "6": "Closed", "7": "Settled"},
            "FCreateType": {"1": "Manual Input", "2": "Generate MRP", "3": "Generate MPS", "4": "Planned Order Release", "5": "Subcontract Order Release", "6": "Split Generated", "7": "Sales Order Push"},
            "FPickMtrlStatus": {"0": "Not Picked", "1": "Partially Picked", "2": "Fully Picked", "3": "Over Picked", "4": "Pending Confirmation"},
        }

    id_to_code_map = {
        "FBillType": bill_type_code_map,
        "FMaterialId": material_code_map,
        "FUnitId": unit_code_map,
        "FBaseUnitId": unit_code_map,
        "FPrdOrgId": org_code_map,
        "FWorkShopID": dept_code_map,
        "FBomId": bom_code_map,
        "FLot": lot_name_map,
    }

    zero_to_empty_cols = {"FBomId", "FAuxPropId", "FWorkShopID", "FSrcBillNo", "FLot"}

    special_cols = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if tpl_field in ("单据类型", "Doc Type"):
            special_cols[tpl_idx] = "bill_type_name"
        elif tpl_field in ("物料名称", "Material Name"):
            special_cols[tpl_idx] = "material_name"
        elif tpl_field in ("规格型号", "Specification"):
            special_cols[tpl_idx] = "material_spec"
        elif tpl_field in ("生产车间", "Production Workshop"):
            special_cols[tpl_idx] = "dept_name"
        elif tpl_field in ("单位", "UoM"):
            special_cols[tpl_idx] = "unit_name"
        elif tpl_field in ("生产组织", "Production Org."):
            special_cols[tpl_idx] = "org_name"

    if lang == "CN":
        output_name = "生产订单列表.xlsx"
        sheet_title = "Sheet1"
    else:
        output_name = "K3_ManufacturingOrders.xlsx"
        sheet_title = "Sheet1"
    output_path = OUTPUT_DIR / output_name
    for old_file in OUTPUT_DIR.glob(f"{output_name.split('.')[0]}*.xlsx"):
        try:
            old_file.unlink()
        except PermissionError:
            print(f"  [警告] 文件被占用，跳过: {old_file.name}")

    print(f"  [输出] {output_path.name}")

    workbook = openpyxl.Workbook()
    ws_out = workbook.active
    ws_out.title = sheet_title

    ws_out.append(h0)

    for _, row in df.iterrows():
        output_row = [""] * len(h0)

        for tpl_idx, data_col in field_mapping.items():
            val = row[data_col]
            if pd.isna(val) or str(val).strip() == "":
                continue

            str_val = str(val).strip()

            if str_val in ("0", "0.0") and data_col in zero_to_empty_cols:
                continue

            if data_col in id_to_code_map and id_to_code_map[data_col]:
                str_val = _resolve_id(str_val, id_to_code_map[data_col])

            if data_col in status_map:
                str_val = status_map[data_col].get(str_val, str_val)

            if data_col in date_cols and str_val:
                str_val = _format_date(str_val)

            if tpl_idx in special_cols:
                sc = special_cols[tpl_idx]
                if sc == "bill_type_name":
                    str_val = bill_type_name_map.get(str_val, str_val)
                elif sc == "material_name":
                    name = material_code_to_name.get(str_val, "")
                    if not name or name == str_val:
                        name = material_name_map.get(str_val, "")
                    if name and name != str_val:
                        str_val = name
                    else:
                        str_val = ""
                elif sc == "material_spec":
                    spec = material_spec_map.get(str_val, "")
                    if spec and spec != str_val:
                        str_val = spec
                    else:
                        str_val = ""
                elif sc == "dept_name":
                    str_val = dept_name_map.get(str_val, str_val)
                elif sc == "unit_name":
                    if lang == "CN":
                        str_val = unit_name_map.get(str_val, str_val)
                elif sc == "org_name":
                    str_val = org_name_map.get(str_val, str_val)

            output_row[tpl_idx] = _clean_string(str_val)

        ws_out.append(output_row)

    _apply_date_formats(ws_out, h0, date_cols, field_mapping)
    _save_workbook(workbook, output_path)
    print(f"  [完成] 文件大小: {output_path.stat().st_size / 1024 / 1024:.2f} MB")
    print(f"[完成] 生成生产订单 Excel ({lang_label})")


def generate_purchaseorder_excel(lang: str = "CN") -> None:
    lang_label = "中文" if lang == "CN" else "英文"
    print(f"[开始] 生成采购订单 Excel ({lang_label})")
    OUTPUT_DIR.mkdir(exist_ok=True)

    template_suffix = "中文模板.xlsx" if lang == "CN" else "英文模板.xlsx"
    template_files = list(TEMPLATE_DIR.glob(f"{PO_FORM_ID}*{template_suffix}"))
    if not template_files:
        print(f"  [错误] 找不到采购订单{lang_label}模板文件")
        return

    template_path = template_files[0]
    print(f"  [模板] {template_path.name}")

    wb = openpyxl.load_workbook(template_path)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    h0 = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            h0 = [str(c).strip() if c else "" for c in row]
            break

    print(f"  [表头] {len(h0)} 列")

    data_path = _resolve_local_csv(PO_FORM_ID)
    if not data_path:
        print("  [错误] 找不到采购订单数据文件")
        return

    print(f"  [数据] {data_path.name}")
    df = pd.read_csv(data_path, encoding="utf-8-sig", dtype=str)
    data_cols = df.columns.tolist()
    print(f"  [数据] {len(df)} 行, {len(data_cols)} 列")

    date_cols = {"FDate", "FDeliveryDate", "FCreateDate", "FApproveDate", "FModifyDate", "FCloseDate", "FChangeDate", "FConfirmDate"}

    po_field_map = {
        "单据编号": "FBillNo", "Doc No.": "FBillNo",
        "采购日期": "FDate", "Purchase Date": "FDate",
        "供应商": "FSupplierId", "Supplier": "FSupplierId",
        "单据状态": "FDocumentStatus", "Doc Status": "FDocumentStatus",
        "采购组织": "FPurchaseOrgId", "Purchase Org.": "FPurchaseOrgId",
        "采购部门": "FPurchaseDeptId", "Purchase Dept": "FPurchaseDeptId",
        "结算币别": "FSettleCurrId", "Settlement Currency": "FSettleCurrId",
        "付款条件": "FPayConditionId", "Payment Terms": "FPayConditionId",
        "税额": "FBillTaxAmount", "Tax Amount": "FBillTaxAmount",
        "价税合计": "FBillAllAmount", "Tax-inclusive Price": "FBillAllAmount",
        "本位币": "FLocalCurrId", "Functional Currency": "FLocalCurrId",
        "创建人": "FCreatorId", "Created by": "FCreatorId",
        "关闭状态": "FCloseStatus", "Closing Status": "FCloseStatus",
        "物料编码": "FMaterialId", "Material Code": "FMaterialId",
        "物料名称": "FMaterialId", "Material Name": "FMaterialId",
        "辅助属性": "FAuxPropId", "Secondary Property": "FAuxPropId",
        "辅助属性.包装规格.名称": "FAuxPropId", "Secondary Property.RM Specification.Name": "FAuxPropId",
        "采购单位": "FUnitId", "Purchase Unit": "FUnitId",
        "采购数量": "FQty", "Purchase Qty": "FQty",
        "交货日期": "FDeliveryDate", "Delivery Date": "FDeliveryDate",
        "含税单价": "FTaxNetPrice", "Tax-inclusive U/P": "FTaxNetPrice",
        "业务关闭": "FMRPCloseStatus", "Close Business": "FMRPCloseStatus",
        "累计收料数量": "FReceiveQty", "Accumulated Receipt Qty": "FReceiveQty",
        "累计退料数量": "FSTOCKRETQTY", "Accumulated Return Qty": "FSTOCKRETQTY",
        "委外领料套数": "FWWPickMtlQty", "Subcontract Picking Sets": "FWWPickMtlQty",
        "是否赠品": "FGiveAway", "Gift": "FGiveAway",
        "累计入库数量": "FStockInQty", "Accumulated Warehouse Receipt Qty": "FStockInQty",
    }

    field_mapping = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if not tpl_field:
            continue
        if tpl_field in po_field_map:
            data_field = po_field_map[tpl_field]
            if data_field in data_cols:
                field_mapping[tpl_idx] = data_field

    print(f"  [映射] 找到 {len(field_mapping)} 个字段映射")

    client = KingdeeClient()

    def _collect_ids(col_name, id_type="numeric"):
        if col_name not in data_cols:
            return []
        ids = []
        for val in df[col_name].dropna().unique():
            id_str = str(val).strip()
            if not id_str or id_str in ("0", "0.0"):
                continue
            if id_type == "guid":
                if _is_guid(id_str):
                    ids.append(id_str)
            elif id_type == "numeric":
                if not _is_guid(id_str):
                    try:
                        num = int(float(id_str))
                        if num > 0:
                            ids.append(str(num))
                    except (ValueError, TypeError, OverflowError):
                        pass
            else:
                ids.append(id_str)
        return list(set(ids))

    print(f"  [解析] 开始解析采购订单ID字段...")

    material_ids = _collect_ids("FMaterialId")
    unit_ids = _collect_ids("FUnitId")
    supplier_ids = _collect_ids("FSupplierId")
    org_ids = _collect_ids("FPurchaseOrgId")
    dept_ids = _collect_ids("FPurchaseDeptId")
    settle_curr_ids = _collect_ids("FSettleCurrId")
    local_curr_ids = _collect_ids("FLocalCurrId")
    pay_condition_ids = _collect_ids("FPayConditionId")
    bill_type_ids = _collect_ids("FBillTypeID", "guid")
    aux_prop_ids = _collect_ids("FAuxPropId")
    creator_ids = _collect_ids("FCreatorId", "mixed")

    # 物料映射
    if material_ids:
        print(f"    需要解析 {len(material_ids)} 个物料ID")
        material_code_map, material_name_map = fetch_material_mapping(client, material_ids)
    else:
        material_code_map, material_name_map = _load_material_mapping()

    material_code_to_name = {}
    for mid, code in material_code_map.items():
        name = material_name_map.get(mid, "")
        if code and name:
            material_code_to_name[code] = name

    missing_name_codes = [c for c in set(material_code_map.values()) if c not in material_code_to_name]
    if missing_name_codes:
        print(f"    需要解析 {len(missing_name_codes)} 个物料编码的中文名")
        batch_size = 200
        for i in range(0, len(missing_name_codes), batch_size):
            batch = missing_name_codes[i:i + batch_size]
            filter_parts = [f"FNumber = '{code}'" for code in batch]
            filter_str = " or ".join(filter_parts)
            try:
                rows = client.execute_bill_query({
                    "FormId": "BD_MATERIAL",
                    "FieldKeys": "FNumber,FName",
                    "FilterString": filter_str,
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": len(batch) + 10,
                })
                if rows:
                    for row in rows:
                        fnumber = str(row[0]).strip() if row[0] else ""
                        fname_raw = row[1] if row[1] else ""
                        fname = ""
                        if fname_raw:
                            try:
                                import json as _json
                                if isinstance(fname_raw, str) and fname_raw.startswith("["):
                                    name_list = _json.loads(fname_raw)
                                    if name_list and isinstance(name_list, list) and len(name_list) > 0:
                                        fname = name_list[0].get("Value", str(name_list[0])) if isinstance(name_list[0], dict) else str(name_list[0])
                                else:
                                    fname = str(fname_raw).strip()
                            except Exception:
                                fname = str(fname_raw).strip()
                        if fnumber and fname:
                            material_code_to_name[fnumber] = fname
            except Exception:
                pass

    material_spec_map = _load_generic_mapping("material_spec_map")
    missing_spec_codes = [c for c in set(material_code_map.values()) if c not in material_spec_map]
    if missing_spec_codes:
        print(f"    需要解析 {len(missing_spec_codes)} 个物料编码的规格型号")
        _fetch_batch_spec(client, missing_spec_codes, material_spec_map)
        _save_generic_mapping("material_spec_map", material_spec_map)

    # 单位映射
    if unit_ids:
        print(f"    需要解析 {len(unit_ids)} 个单位ID")
        unit_code_map = fetch_generic_mapping(client, "BD_UNIT", "FUnitID", "FNumber", "unit_code_map", unit_ids)
    else:
        unit_code_map = _load_generic_mapping("unit_code_map")
    unit_name_map = _fetch_name_map(client, "BD_UNIT", list(set(unit_code_map.values())), "unit_name_map")

    # 供应商映射
    supplier_code_map = {}
    supplier_name_map = {}
    if supplier_ids:
        print(f"    需要解析 {len(supplier_ids)} 个供应商ID")
        supplier_code_map = fetch_generic_mapping(client, "BD_Supplier", "FSUPPLIERID", "FNumber", "supplier_code_map", supplier_ids)
        supplier_name_map = _fetch_name_map(client, "BD_Supplier", list(set(supplier_code_map.values())), "supplier_name_map")

    # 组织映射
    if org_ids:
        print(f"    需要解析 {len(org_ids)} 个组织ID")
        org_code_map = fetch_generic_mapping(client, "ORG_Organizations", "FOrgId", "FNumber", "org_code_map", org_ids)
    else:
        org_code_map = _load_generic_mapping("org_code_map")
    org_name_map = _fetch_name_map(client, "ORG_Organizations", list(set(org_code_map.values())), "org_name_map")

    # 采购部门映射
    dept_code_map = {}
    dept_name_map = {}
    if dept_ids:
        print(f"    需要解析 {len(dept_ids)} 个采购部门ID")
        dept_code_map = fetch_generic_mapping(client, "BD_Department", "FDeptID", "FNumber", "dept_code_map", dept_ids)
        dept_name_map = _fetch_name_map(client, "BD_Department", list(set(dept_code_map.values())), "dept_name_map")

    # 币别映射
    currency_code_map = {}
    currency_name_map = {}
    all_curr_ids = list(set(settle_curr_ids + local_curr_ids))
    if all_curr_ids:
        print(f"    需要解析 {len(all_curr_ids)} 个币别ID")
        currency_code_map = fetch_generic_mapping(client, "BD_Currency", "FCURRENCYID", "FNumber", "currency_code_map", all_curr_ids)
        currency_name_map = _fetch_name_map(client, "BD_Currency", list(set(currency_code_map.values())), "currency_name_map")

    # 付款条件映射
    pay_condition_code_map = {}
    pay_condition_name_map = {}
    if pay_condition_ids:
        print(f"    需要解析 {len(pay_condition_ids)} 个付款条件ID")
        pay_condition_code_map = fetch_generic_mapping(client, "BD_PayCondition", "FID", "FNumber", "pay_condition_code_map", pay_condition_ids)
        pay_condition_name_map = _fetch_name_map(client, "BD_PayCondition", list(set(pay_condition_code_map.values())), "pay_condition_name_map")

    # 单据类型映射
    bill_type_code_map = {}
    bill_type_name_map = {}
    if bill_type_ids:
        print(f"    需要解析 {len(bill_type_ids)} 个单据类型ID(GUID)")
        bill_type_code_map = fetch_generic_mapping(client, "BOS_BillType", "FBILLTYPEID", "FNumber", "bill_type_map", bill_type_ids)
        guid_name_map = fetch_generic_mapping(client, "BOS_BillType", "FBILLTYPEID", "FName", "bill_type_name_map_guid", bill_type_ids)
        for gid, name in guid_name_map.items():
            code = bill_type_code_map.get(gid, "")
            if code and name:
                bill_type_name_map[code] = name
            if gid and name:
                bill_type_name_map[gid] = name
    else:
        bill_type_code_map = _load_generic_mapping("bill_type_map")
    bill_type_name_map.update(_load_generic_mapping("bill_type_name_map"))

    # 创建人映射
    creator_name_map = {}
    if creator_ids:
        print(f"    需要解析 {len(creator_ids)} 个创建人ID")
        creator_name_map = _load_generic_mapping("user_name_map")
        missing_users = [uid for uid in creator_ids if uid not in creator_name_map]
        if missing_users:
            print(f"    需要解析 {len(missing_users)} 个用户ID的中文名")
            for uid in missing_users:
                try:
                    result = client.view("SEC_User", {"Id": uid, "Number": "", "CreateOrgId": 0})
                    inner = result.get("Result", {}).get("Result", {})
                    if inner:
                        user_name = inner.get("Name", "")
                        if isinstance(user_name, list) and len(user_name) > 0:
                            user_name = user_name[0].get("Value", "") if isinstance(user_name[0], dict) else str(user_name[0])
                        if not user_name:
                            user_name = inner.get("UserAccount", "")
                        if user_name:
                            creator_name_map[uid] = str(user_name).strip()
                except Exception:
                    pass
            _save_generic_mapping("user_name_map", creator_name_map)

    # 辅助属性映射（包装规格名称）
    # FAuxPropId 是采购订单明细中的辅助属性分录ID
    # 需要通过 PO view API 获取 AuxPropId.F100001.FDataValue（包装规格名称）
    # 缓存格式：{aux_prop_id: {"cn": "中文包装规格", "en": "English spec"}}
    # 辅助属性类别固定为"包装规格"/"RM Specification"
    aux_prop_name_map = {}
    if aux_prop_ids:
        print(f"    需要解析 {len(aux_prop_ids)} 个辅助属性ID的包装规格名称")
        # 先加载缓存
        aux_prop_name_map = _load_generic_mapping("po_aux_prop_name_map")
        # 兼容旧格式缓存：如果值是字符串则转为dict
        for k, v in list(aux_prop_name_map.items()):
            if isinstance(v, str):
                aux_prop_name_map[k] = {"cn": v, "en": v}
        missing_aux_ids = [aid for aid in aux_prop_ids if aid not in aux_prop_name_map]
        if missing_aux_ids:
            # 获取包含这些 FAuxPropId 的单据编号
            aux_bill_map = {}
            for _, row in df.iterrows():
                aid = str(row.get("FAuxPropId", "")).strip()
                if aid in missing_aux_ids and aid not in aux_bill_map:
                    bill_no = str(row.get("FBillNo", "")).strip()
                    if bill_no:
                        aux_bill_map[aid] = bill_no
            # 通过 view API 查询 PO 单据获取辅助属性信息
            for aid, bill_no in aux_bill_map.items():
                try:
                    result = client.view("PUR_PurchaseOrder", {"Id": 0, "Number": bill_no, "CreateOrgId": 0})
                    inner = result.get("Result", {}).get("Result", {})
                    entries = inner.get("POOrderEntry", [])
                    for entry in entries:
                        aux_prop = entry.get("AuxPropId", {})
                        if aux_prop:
                            entry_aux_id = str(aux_prop.get("Id", ""))
                            if entry_aux_id and entry_aux_id not in aux_prop_name_map:
                                # 解析包装规格名称（F100001）
                                cn_spec = ""
                                en_spec = ""
                                f100001 = aux_prop.get("F100001", {})
                                if f100001:
                                    fdata_values = f100001.get("FDataValue", [])
                                    for dv in fdata_values:
                                        if isinstance(dv, dict):
                                            if dv.get("Key") == 2052:
                                                cn_spec = dv.get("Value", "")
                                            elif dv.get("Key") == 1033:
                                                en_spec = dv.get("Value", "")
                                # 清理尾部多余的空格和点号（金蝶数据格式问题）
                                cn_spec = cn_spec.strip().rstrip('.').strip()
                                en_spec = en_spec.strip().rstrip('.').strip()
                                aux_prop_name_map[entry_aux_id] = {
                                    "cn": cn_spec,
                                    "en": en_spec if en_spec else cn_spec,
                                }
                except Exception:
                    pass
            _save_generic_mapping("po_aux_prop_name_map", aux_prop_name_map)

    # 状态映射
    if lang == "CN":
        status_map = {
            "FDocumentStatus": {"A": "已审核", "B": "暂存", "C": "已审核", "D": "重新审核"},
            "FCloseStatus": {"A": "未关闭", "B": "已关闭"},
            "FMRPCloseStatus": {"0": "未关闭", "1": "已关闭"},
            "FGiveAway": {"0": "否", "1": "是"},
            "FConfirmStatus": {"A": "未确认", "B": "已确认", "C": "已拒绝"},
        }
    else:
        status_map = {
            "FDocumentStatus": {"A": "Approved", "B": "Draft", "C": "Approved", "D": "Re-audit"},
            "FCloseStatus": {"A": "Unclosed", "B": "Closed"},
            "FMRPCloseStatus": {"0": "Unclosed", "1": "Closed"},
            "FGiveAway": {"0": "No", "1": "Yes"},
            "FConfirmStatus": {"A": "Unconfirmed", "B": "Confirmed", "C": "Rejected"},
        }

    id_to_code_map = {
        "FBillTypeID": bill_type_code_map,
        "FMaterialId": material_code_map,
        "FUnitId": unit_code_map,
        "FSupplierId": supplier_code_map,
        "FPurchaseOrgId": org_code_map,
        "FPurchaseDeptId": dept_code_map,
        "FSettleCurrId": currency_code_map,
        "FLocalCurrId": currency_code_map,
        "FPayConditionId": pay_condition_code_map,
    }

    zero_to_empty_cols = {"FAuxPropId", "FSupplierId", "FPurchaseDeptId", "FPayConditionId"}

    special_cols = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if tpl_field in ("供应商", "Supplier"):
            special_cols[tpl_idx] = "supplier_name"
        elif tpl_field in ("物料名称", "Material Name"):
            special_cols[tpl_idx] = "material_name"
        elif tpl_field in ("辅助属性.包装规格.名称", "Secondary Property.RM Specification.Name"):
            special_cols[tpl_idx] = "aux_prop_name"
        elif tpl_field in ("辅助属性", "Secondary Property"):
            special_cols[tpl_idx] = "aux_prop_category"
        elif tpl_field in ("采购组织", "Purchase Org."):
            special_cols[tpl_idx] = "org_name"
        elif tpl_field in ("采购部门", "Purchase Dept"):
            special_cols[tpl_idx] = "dept_name"
        elif tpl_field in ("采购单位", "Purchase Unit"):
            special_cols[tpl_idx] = "unit_name"
        elif tpl_field in ("结算币别", "Settlement Currency"):
            special_cols[tpl_idx] = "currency_name"
        elif tpl_field in ("本位币", "Functional Currency"):
            special_cols[tpl_idx] = "currency_name"
        elif tpl_field in ("付款条件", "Payment Terms"):
            special_cols[tpl_idx] = "pay_condition_name"
        elif tpl_field in ("创建人", "Created by"):
            special_cols[tpl_idx] = "creator_name"

    if lang == "CN":
        output_name = "采购订单列表.xlsx"
        sheet_title = "Sheet1"
    else:
        output_name = "K3_PurchaseOrders.xlsx"
        sheet_title = "Sheet1"
    output_path = OUTPUT_DIR / output_name
    for old_file in OUTPUT_DIR.glob(f"{output_name.split('.')[0]}*.xlsx"):
        try:
            old_file.unlink()
        except PermissionError:
            print(f"  [警告] 文件被占用，跳过: {old_file.name}")

    print(f"  [输出] {output_path.name}")

    workbook = openpyxl.Workbook()
    ws_out = workbook.active
    ws_out.title = sheet_title

    ws_out.append(h0)

    for _, row in df.iterrows():
        output_row = [""] * len(h0)

        for tpl_idx, data_col in field_mapping.items():
            val = row[data_col]
            if pd.isna(val) or str(val).strip() == "":
                continue

            str_val = str(val).strip()

            if str_val in ("0", "0.0") and data_col in zero_to_empty_cols:
                continue

            if data_col in id_to_code_map and id_to_code_map[data_col]:
                str_val = _resolve_id(str_val, id_to_code_map[data_col])

            if data_col in status_map:
                str_val = status_map[data_col].get(str_val, str_val)

            if data_col in date_cols and str_val:
                str_val = _format_date(str_val)

            if tpl_idx in special_cols:
                sc = special_cols[tpl_idx]
                if sc == "supplier_name":
                    str_val = supplier_name_map.get(str_val, str_val)
                elif sc == "material_name":
                    name = material_code_to_name.get(str_val, "")
                    if not name or name == str_val:
                        name = material_name_map.get(str_val, "")
                    if name and name != str_val:
                        str_val = name
                    else:
                        str_val = ""
                elif sc == "aux_prop_name":
                    info = aux_prop_name_map.get(str_val, {})
                    if isinstance(info, dict):
                        name = info.get("cn" if lang == "CN" else "en", "")
                    else:
                        name = str(info) if info else ""
                    str_val = name if name else ""
                elif sc == "aux_prop_category":
                    # 辅助属性类别固定为"包装规格"/"RM Specification"
                    if str_val and str_val != "0" and str_val != "0.0":
                        str_val = "包装规格" if lang == "CN" else "RM Specification"
                    else:
                        str_val = ""
                elif sc == "org_name":
                    str_val = org_name_map.get(str_val, str_val)
                elif sc == "dept_name":
                    str_val = dept_name_map.get(str_val, str_val)
                elif sc == "unit_name":
                    if lang == "CN":
                        str_val = unit_name_map.get(str_val, str_val)
                elif sc == "currency_name":
                    str_val = currency_name_map.get(str_val, str_val)
                elif sc == "pay_condition_name":
                    str_val = pay_condition_name_map.get(str_val, str_val)
                elif sc == "creator_name":
                    str_val = creator_name_map.get(str_val, str_val)

            output_row[tpl_idx] = _clean_string(str_val)

        ws_out.append(output_row)

    _apply_date_formats(ws_out, h0, date_cols, field_mapping)
    _save_workbook(workbook, output_path)
    print(f"  [完成] 文件大小: {output_path.stat().st_size / 1024 / 1024:.2f} MB")
    print(f"[完成] 生成采购订单 Excel ({lang_label})")


def generate_instock_excel(lang: str = "CN") -> None:
    lang_label = "中文" if lang == "CN" else "英文"
    print(f"[开始] 生成生产入库单 Excel ({lang_label})")
    OUTPUT_DIR.mkdir(exist_ok=True)

    template_suffix = "中文模板.xlsx" if lang == "CN" else "英文模板.xlsx"
    template_files = list(TEMPLATE_DIR.glob(f"{INSTOCK_FORM_ID}*{template_suffix}"))
    if not template_files:
        print(f"  [错误] 找不到生产入库单{lang_label}模板文件")
        return

    template_path = template_files[0]
    print(f"  [模板] {template_path.name}")

    wb = openpyxl.load_workbook(template_path)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    h0 = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            h0 = [str(c).strip() if c else "" for c in row]
            break

    print(f"  [表头] {len(h0)} 列")

    data_path = _resolve_local_csv(INSTOCK_FORM_ID)
    if not data_path:
        print("  [错误] 找不到生产入库单数据文件")
        return

    print(f"  [数据] {data_path.name}")
    df = pd.read_csv(data_path, encoding="utf-8-sig", dtype=str)
    data_cols = df.columns.tolist()
    print(f"  [数据] {len(df)} 行, {len(data_cols)} 列")

    date_cols = {"FDate", "FProduceDate", "FExpiryDate"}

    # 新模板23列字段映射
    instock_field_map = {
        "单据编号": "FBillNo", "Doc No.": "FBillNo",
        "日期": "FDate", "Date": "FDate",
        "单据状态": "FDocumentStatus", "Doc Status": "FDocumentStatus",
        "入库组织": "FStockOrgId", "Receipt Org.": "FStockOrgId",
        "生产组织": "FPrdOrgId", "Production Org.": "FPrdOrgId",
        "总成本": "FAmount", "Amount": "FAmount",
        "成本价": "FPrice", "U/P": "FPrice",
        "物料编码": "FMaterialId", "Material Code": "FMaterialId",
        "物料名称": "FMaterialId", "Material Name": "FMaterialId",
        "规格型号": "FMaterialId", "Specification": "FMaterialId",
        "产品类型": "FProductType", "Product Type": "FProductType",
        "单位": "FUnitID", "UoM": "FUnitID",
        "应收数量": "FMustQty", "Qty Receivable": "FMustQty",
        "实收数量": "FRealQty", "Qty Received": "FRealQty",
        "仓库": "FStockId", "Warehouse": "FStockId",
        "批号": "FLot", "Lot No.": "FLot",
        "生产车间": "FWorkShopId1", "Production Workshop": "FWorkShopId1",
        "班组": "FShiftGroupId", "Team": "FShiftGroupId",
        "辅助属性": "FAuxpropId", "Secondary Property": "FAuxpropId",
        "生产日期": "FProduceDate", "Production Date": "FProduceDate",
        "有效期至": "FExpiryDate", "Valid till": "FExpiryDate",
        "生产订单编号": "FMoBillNo", "MO No.": "FMoBillNo",
        "生产订单行号": "FMoEntrySeq", "MO Row No.": "FMoEntrySeq",
    }

    field_mapping = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if not tpl_field:
            continue
        if tpl_field in instock_field_map:
            data_field = instock_field_map[tpl_field]
            if data_field in data_cols:
                field_mapping[tpl_idx] = data_field

    print(f"  [映射] 找到 {len(field_mapping)} 个字段映射")

    client = KingdeeClient()

    def _collect_ids(col_name, id_type="numeric"):
        if col_name not in data_cols:
            return []
        ids = []
        for val in df[col_name].dropna().unique():
            id_str = str(val).strip()
            if not id_str or id_str in ("0", "0.0"):
                continue
            if id_type == "guid":
                if _is_guid(id_str):
                    ids.append(id_str)
            elif id_type == "numeric":
                if not _is_guid(id_str):
                    try:
                        num = int(float(id_str))
                        if num > 0:
                            ids.append(str(num))
                    except (ValueError, TypeError, OverflowError):
                        pass
            else:
                ids.append(id_str)
        return list(set(ids))

    print(f"  [解析] 开始解析生产入库单ID字段...")

    material_ids = _collect_ids("FMaterialId")
    unit_ids = _collect_ids("FUnitID")
    stock_org_ids = _collect_ids("FStockOrgId")
    prd_org_ids = _collect_ids("FPrdOrgId")
    stock_ids = _collect_ids("FStockId")
    lot_ids = _collect_ids("FLot")
    workshop_ids = _collect_ids("FWorkShopId1")
    shift_group_ids = _collect_ids("FShiftGroupId")
    aux_prop_ids = _collect_ids("FAuxpropId")

    # 合并所有组织ID（入库组织 + 生产组织）
    all_org_ids = list(set(stock_org_ids + prd_org_ids))

    # 物料映射
    if material_ids:
        print(f"    需要解析 {len(material_ids)} 个物料ID")
        material_code_map, material_name_map = fetch_material_mapping(client, material_ids)
    else:
        material_code_map, material_name_map = _load_material_mapping()

    material_code_to_name = {}
    for mid, code in material_code_map.items():
        name = material_name_map.get(mid, "")
        if code and name:
            material_code_to_name[code] = name

    missing_name_codes = [c for c in set(material_code_map.values()) if c not in material_code_to_name]
    if missing_name_codes:
        print(f"    需要解析 {len(missing_name_codes)} 个物料编码的中文名")
        batch_size = 200
        for i in range(0, len(missing_name_codes), batch_size):
            batch = missing_name_codes[i:i + batch_size]
            filter_parts = [f"FNumber = '{code}'" for code in batch]
            filter_str = " or ".join(filter_parts)
            try:
                rows = client.execute_bill_query({
                    "FormId": "BD_MATERIAL",
                    "FieldKeys": "FNumber,FName",
                    "FilterString": filter_str,
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": len(batch) + 10,
                })
                if rows:
                    for row in rows:
                        fnumber = str(row[0]).strip() if row[0] else ""
                        fname_raw = row[1] if row[1] else ""
                        fname = ""
                        if fname_raw:
                            try:
                                import json as _json
                                if isinstance(fname_raw, str) and fname_raw.startswith("["):
                                    name_list = _json.loads(fname_raw)
                                    if name_list and isinstance(name_list, list) and len(name_list) > 0:
                                        fname = name_list[0].get("Value", str(name_list[0])) if isinstance(name_list[0], dict) else str(name_list[0])
                                else:
                                    fname = str(fname_raw).strip()
                            except Exception:
                                fname = str(fname_raw).strip()
                        if fnumber and fname:
                            material_code_to_name[fnumber] = fname
            except Exception:
                pass

    material_spec_map = _load_generic_mapping("material_spec_map")
    missing_spec_codes = [c for c in set(material_code_map.values()) if c not in material_spec_map]
    if missing_spec_codes:
        print(f"    需要解析 {len(missing_spec_codes)} 个物料编码的规格型号")
        _fetch_batch_spec(client, missing_spec_codes, material_spec_map)
        _save_generic_mapping("material_spec_map", material_spec_map)

    # 单位映射
    if unit_ids:
        print(f"    需要解析 {len(unit_ids)} 个单位ID")
        unit_code_map = fetch_generic_mapping(client, "BD_UNIT", "FUnitID", "FNumber", "unit_code_map", unit_ids)
    else:
        unit_code_map = _load_generic_mapping("unit_code_map")
    unit_name_map = _fetch_name_map(client, "BD_UNIT", list(set(unit_code_map.values())), "unit_name_map")

    # 组织映射（入库组织 + 生产组织共用）
    if all_org_ids:
        print(f"    需要解析 {len(all_org_ids)} 个组织ID（入库+生产）")
        org_code_map = fetch_generic_mapping(client, "ORG_Organizations", "FOrgId", "FNumber", "org_code_map", all_org_ids)
    else:
        org_code_map = _load_generic_mapping("org_code_map")
    org_name_map = _fetch_name_map(client, "ORG_Organizations", list(set(org_code_map.values())), "org_name_map")

    # 仓库映射
    stock_code_map = {}
    stock_name_map = {}
    if stock_ids:
        print(f"    需要解析 {len(stock_ids)} 个仓库ID")
        stock_code_map = fetch_generic_mapping(client, "BD_STOCK", "FStockID", "FNumber", "stock_code_map", stock_ids)
        stock_name_map = _fetch_name_map(client, "BD_STOCK", list(set(stock_code_map.values())), "stock_name_map")

    # 批号映射
    lot_name_map = {}
    if lot_ids:
        print(f"    需要解析 {len(lot_ids)} 个批号ID")
        lot_name_map = _load_generic_mapping("inv_lot_name_map")
        # 兼容旧格式：如果值是dict则转为字符串
        for k, v in list(lot_name_map.items()):
            if isinstance(v, dict):
                lot_name_map[k] = v.get("cn", str(v))
        missing_lots = [lid for lid in lot_ids if lid not in lot_name_map]
        if missing_lots:
            print(f"    需要查询 {len(missing_lots)} 个批号名称（STK_Inventory）")
            batch_size = 200
            for i in range(0, len(missing_lots), batch_size):
                batch = missing_lots[i:i + batch_size]
                filter_parts = [f"FLot = {lid}" for lid in batch]
                filter_str = " or ".join(filter_parts)
                try:
                    rows = client.execute_bill_query({
                        "FormId": "STK_Inventory",
                        "FieldKeys": "FLot,FLot.FNumber",
                        "FilterString": filter_str,
                        "OrderString": "",
                        "TopRowCount": 0,
                        "StartRow": 0,
                        "Limit": len(batch) + 10,
                    })
                    if rows:
                        for row in rows:
                            lot_id = str(row[0]).strip() if row[0] else ""
                            lot_number = str(row[1]).strip() if row[1] else ""
                            if lot_id and lot_number and lot_number not in ("None", "0"):
                                lot_name_map[lot_id] = lot_number
                except Exception:
                    pass
        # 检查仍未解析的批号ID，通过 view API 回退
        still_missing = [lid for lid in lot_ids if lid not in lot_name_map]
        if still_missing:
            print(f"    仍有 {len(still_missing)} 个批号未解析，尝试通过 view API 回退")
            lot_bill_map = {}
            for _, row in df.iterrows():
                lid = str(row.get("FLot", "")).strip()
                if lid in still_missing and lid not in lot_bill_map:
                    bill_no = str(row.get("FBillNo", "")).strip()
                    if bill_no:
                        lot_bill_map[lid] = bill_no
            for lid, bill_no in lot_bill_map.items():
                try:
                    result = client.view(INSTOCK_FORM_ID, {"Id": 0, "Number": bill_no, "CreateOrgId": 0})
                    inner = result.get("Result", {}).get("Result", {})
                    entries = inner.get("Entity", [])
                    for entry in entries:
                        lot_obj = entry.get("Lot", {})
                        if lot_obj:
                            entry_lot_id = str(lot_obj.get("Id", ""))
                            if entry_lot_id and entry_lot_id not in lot_name_map:
                                lot_number = str(lot_obj.get("Number", "")).strip()
                                if lot_number and lot_number not in ("None", "0"):
                                    lot_name_map[entry_lot_id] = lot_number
                except Exception:
                    pass
        _save_generic_mapping("inv_lot_name_map", lot_name_map)

    # 生产车间映射（FWorkShopId1 -> BD_Department）
    dept_code_map = {}
    dept_name_map = {}
    if workshop_ids:
        print(f"    需要解析 {len(workshop_ids)} 个生产车间ID")
        dept_code_map = fetch_generic_mapping(client, "BD_Department", "FDeptID", "FNumber", "dept_code_map", workshop_ids)
        dept_name_map = _fetch_name_map(client, "BD_Department", list(set(dept_code_map.values())), "dept_name_map")

    # 班组映射（FShiftGroupId -> BD_SHIFTGROUP）
    shift_group_name_map = {}
    if shift_group_ids:
        print(f"    需要解析 {len(shift_group_ids)} 个班组ID")
        shift_group_name_map = _load_generic_mapping("instock_shift_group_name_map")
        missing_sgs = [sid for sid in shift_group_ids if sid not in shift_group_name_map]
        if missing_sgs:
            print(f"    需要查询 {len(missing_sgs)} 个班组名称")
            # 尝试通过 BD_SHIFTGROUP 查询
            try:
                batch_size = 200
                for i in range(0, len(missing_sgs), batch_size):
                    batch = missing_sgs[i:i + batch_size]
                    filter_parts = [f"FID = {sid}" for sid in batch]
                    filter_str = " or ".join(filter_parts)
                    rows = client.execute_bill_query({
                        "FormId": "BD_SHIFTGROUP",
                        "FieldKeys": "FID,FNumber,FName",
                        "FilterString": filter_str,
                        "OrderString": "",
                        "TopRowCount": 0,
                        "StartRow": 0,
                        "Limit": len(batch) + 10,
                    })
                    if rows:
                        for row in rows:
                            sg_id = str(row[0]).strip() if row[0] else ""
                            sg_name = str(row[2]).strip() if row[2] else ""
                            if sg_id and sg_name and sg_name not in ("None", "0"):
                                shift_group_name_map[sg_id] = sg_name
            except Exception:
                # BD_SHIFTGROUP 查询失败，尝试通过 view API 解析
                print(f"    BD_SHIFTGROUP 查询失败，尝试通过 PRD_INSTOCK view API 解析班组")
                sg_bill_map = {}
                for _, row in df.iterrows():
                    sid = str(row.get("FShiftGroupId", "")).strip()
                    if sid in missing_sgs and sid not in sg_bill_map:
                        bill_no = str(row.get("FBillNo", "")).strip()
                        if bill_no:
                            sg_bill_map[sid] = bill_no
                for sid, bill_no in sg_bill_map.items():
                    try:
                        result = client.view(INSTOCK_FORM_ID, {"Id": 0, "Number": bill_no, "CreateOrgId": 0})
                        inner = result.get("Result", {}).get("Result", {})
                        entries = inner.get("Entity", [])
                        for entry in entries:
                            sg = entry.get("ShiftGroupId", {})
                            if sg:
                                entry_sg_id = str(sg.get("Id", ""))
                                if entry_sg_id and entry_sg_id not in shift_group_name_map:
                                    sg_name = ""
                                    name_list = sg.get("Name", [])
                                    if isinstance(name_list, list):
                                        for n in name_list:
                                            if isinstance(n, dict) and n.get("Key") == 2052:
                                                sg_name = n.get("Value", "")
                                                break
                                    elif isinstance(name_list, str):
                                        sg_name = name_list
                                    if sg_name:
                                        shift_group_name_map[entry_sg_id] = sg_name
                    except Exception:
                        pass
            _save_generic_mapping("instock_shift_group_name_map", shift_group_name_map)

    # 辅助属性映射（FAuxpropId -> 通过 view API 获取包装规格名称）
    aux_prop_name_map = {}
    if aux_prop_ids:
        print(f"    需要解析 {len(aux_prop_ids)} 个辅助属性ID")
        aux_prop_name_map = _load_generic_mapping("instock_aux_prop_name_map")
        # 兼容旧格式缓存
        for k, v in list(aux_prop_name_map.items()):
            if isinstance(v, str):
                aux_prop_name_map[k] = {"cn": v, "en": v}
        missing_aux_ids = [aid for aid in aux_prop_ids if aid not in aux_prop_name_map]
        if missing_aux_ids:
            print(f"    需要查询 {len(missing_aux_ids)} 个辅助属性名称")
            # 获取包含这些 FAuxpropId 的单据编号
            aux_bill_map = {}
            for _, row in df.iterrows():
                aid = str(row.get("FAuxpropId", "")).strip()
                if aid in missing_aux_ids and aid not in aux_bill_map:
                    bill_no = str(row.get("FBillNo", "")).strip()
                    if bill_no:
                        aux_bill_map[aid] = bill_no
            # 通过 view API 查询 PRD_INSTOCK 单据获取辅助属性信息
            for aid, bill_no in aux_bill_map.items():
                try:
                    result = client.view(INSTOCK_FORM_ID, {"Id": 0, "Number": bill_no, "CreateOrgId": 0})
                    inner = result.get("Result", {}).get("Result", {})
                    entries = inner.get("Entity", [])
                    for entry in entries:
                        aux_prop = entry.get("AuxpropId", {})
                        if aux_prop:
                            entry_aux_id = str(aux_prop.get("Id", ""))
                            if entry_aux_id and entry_aux_id not in aux_prop_name_map:
                                cn_spec = ""
                                en_spec = ""
                                f100001 = aux_prop.get("F100001", {})
                                if f100001:
                                    fdata_values = f100001.get("FDataValue", [])
                                    for dv in fdata_values:
                                        if isinstance(dv, dict):
                                            if dv.get("Key") == 2052:
                                                cn_spec = dv.get("Value", "")
                                            elif dv.get("Key") == 1033:
                                                en_spec = dv.get("Value", "")
                                cn_spec = cn_spec.strip().rstrip('.').strip()
                                en_spec = en_spec.strip().rstrip('.').strip()
                                aux_prop_name_map[entry_aux_id] = {
                                    "cn": cn_spec,
                                    "en": en_spec if en_spec else cn_spec,
                                }
                except Exception:
                    pass
            _save_generic_mapping("instock_aux_prop_name_map", aux_prop_name_map)

    # 状态映射
    if lang == "CN":
        status_map = {
            "FDocumentStatus": {"A": "已审核", "B": "暂存", "C": "已审核", "D": "重新审核"},
            "FProductType": {"1": "主产品", "2": "副产品", "3": "联产品"},
        }
    else:
        status_map = {
            "FDocumentStatus": {"A": "Approved", "B": "Draft", "C": "Approved", "D": "Re-audit"},
            "FProductType": {"1": "Main Product", "2": "By-product", "3": "Co-product"},
        }

    id_to_code_map = {
        "FMaterialId": material_code_map,
        "FUnitID": unit_code_map,
        "FStockOrgId": org_code_map,
        "FPrdOrgId": org_code_map,
        "FStockId": stock_code_map,
        "FLot": lot_name_map,
        "FWorkShopId1": dept_code_map,
        # FShiftGroupId 和 FAuxpropId 的映射值是dict格式，
        # 不放入id_to_code_map，由special_cols直接用原始ID查找
    }

    zero_to_empty_cols = {"FAuxpropId", "FWorkShopId1", "FShiftGroupId", "FLot"}

    special_cols = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if tpl_field in ("物料名称", "Material Name"):
            special_cols[tpl_idx] = "material_name"
        elif tpl_field in ("规格型号", "Specification"):
            special_cols[tpl_idx] = "material_spec"
        elif tpl_field in ("仓库", "Warehouse"):
            special_cols[tpl_idx] = "stock_name"
        elif tpl_field in ("单位", "UoM"):
            special_cols[tpl_idx] = "unit_name"
        elif tpl_field in ("入库组织", "Receipt Org."):
            special_cols[tpl_idx] = "stock_org_name"
        elif tpl_field in ("生产组织", "Production Org."):
            special_cols[tpl_idx] = "prd_org_name"
        elif tpl_field in ("生产车间", "Production Workshop"):
            special_cols[tpl_idx] = "dept_name"
        elif tpl_field in ("班组", "Team"):
            special_cols[tpl_idx] = "shift_group_name"
        elif tpl_field in ("辅助属性", "Secondary Property"):
            special_cols[tpl_idx] = "aux_prop_name"

    if lang == "CN":
        output_name = "生产入库单列表.xlsx"
        sheet_title = "Sheet1"
    else:
        output_name = "K3_Production_Receipts.xlsx"
        sheet_title = "Sheet1"
    output_path = OUTPUT_DIR / output_name
    for old_file in OUTPUT_DIR.glob(f"{output_name.split('.')[0]}*.xlsx"):
        try:
            old_file.unlink()
        except PermissionError:
            print(f"  [警告] 文件被占用，跳过: {old_file.name}")

    print(f"  [输出] {output_path.name}")

    workbook = openpyxl.Workbook()
    ws_out = workbook.active
    ws_out.title = sheet_title

    ws_out.append(h0)

    for _, row in df.iterrows():
        output_row = [""] * len(h0)

        for tpl_idx, data_col in field_mapping.items():
            val = row[data_col]
            if pd.isna(val) or str(val).strip() == "":
                continue

            str_val = str(val).strip()
            original_val = str_val  # 保存原始值，供special_cols查找dict格式映射

            if str_val in ("0", "0.0") and data_col in zero_to_empty_cols:
                continue

            if data_col in id_to_code_map and id_to_code_map[data_col]:
                str_val = _resolve_id(str_val, id_to_code_map[data_col])

            if data_col in status_map:
                str_val = status_map[data_col].get(str_val, str_val)

            if data_col in date_cols and str_val:
                str_val = _format_date(str_val)

            if tpl_idx in special_cols:
                sc = special_cols[tpl_idx]
                if sc == "material_name":
                    name = material_code_to_name.get(str_val, "")
                    if not name or name == str_val:
                        name = material_name_map.get(str_val, "")
                    if name and name != str_val:
                        str_val = name
                    else:
                        str_val = ""
                elif sc == "material_spec":
                    spec = material_spec_map.get(str_val, "")
                    if spec and spec != str_val:
                        str_val = spec
                    else:
                        str_val = ""
                elif sc == "stock_name":
                    str_val = stock_name_map.get(str_val, str_val)
                elif sc == "unit_name":
                    if lang == "CN":
                        str_val = unit_name_map.get(str_val, str_val)
                elif sc == "stock_org_name":
                    str_val = org_name_map.get(str_val, str_val)
                elif sc == "prd_org_name":
                    str_val = org_name_map.get(str_val, str_val)
                elif sc == "dept_name":
                    str_val = dept_name_map.get(str_val, str_val)
                elif sc == "shift_group_name":
                    # 使用原始ID查找，映射值可能是dict或字符串
                    resolved = shift_group_name_map.get(original_val, original_val)
                    if isinstance(resolved, dict):
                        str_val = resolved.get("cn" if lang == "CN" else "en", resolved.get("cn", original_val))
                    else:
                        str_val = resolved
                elif sc == "aux_prop_name":
                    # 使用原始ID查找，映射格式: {id: {"cn": "...", "en": "..."}}
                    info = aux_prop_name_map.get(original_val, {})
                    if isinstance(info, dict):
                        str_val = info.get("cn" if lang == "CN" else "en", info.get("cn", ""))
                    elif isinstance(info, str):
                        str_val = info
                    else:
                        str_val = ""

            output_row[tpl_idx] = _clean_string(str_val)

        ws_out.append(output_row)

    _apply_date_formats(ws_out, h0, date_cols, field_mapping)
    _save_workbook(workbook, output_path)
    print(f"  [完成] 文件大小: {output_path.stat().st_size / 1024 / 1024:.2f} MB")
    print(f"[完成] 生成生产入库单 Excel ({lang_label})")


# ══════════════════════════════════════════════════════════════════════════════
# STK_Inventory 即时库存 Excel 生成（使用模板+分批查询+合并）
# ══════════════════════════════════════════════════════════════════════════════


def generate_inv_excel(lang: str = "CN") -> None:
    """生成 STK_Inventory 即时库存 Excel（使用模板+分批查询+合并）。

    STK_Inventory 对多字段查询有限制（超过6-7个字段只返回1行），
    因此分批查询少量字段，以 FMaterialId+FStockId+FLot 为主键合并。
    """
    import json
    import time

    lang_label = "中文" if lang == "CN" else "英文"
    print(f"\n[开始] 生成即时库存 Excel ({lang_label})")
    OUTPUT_DIR.mkdir(exist_ok=True)

    # 库存状态英文映射（FNumber → 英文名称）
    _stock_status_en_map = {
        "KCZT01_SYS": "Available",
        "KCZT02_SYS": "Pending Inspection",
        "KCZT03_SYS": "Pending Freeze",
        "KCZT04_SYS": "Return Freeze",
        "KCZT05_SYS": "In Transit",
        "KCZT06_SYS": "Receipt Freeze",
        "KCZT07_SYS": "Scrap",
        "KCZT08_SYS": "Defective",
        "KCZT001": "Not accounting",
        "KCZT99_SYS": "On Loan",
        "KCZT09_SYS": "Manual Freeze",
        "KCZT002": "Defective",
    }

    # 读取模板表头
    template_suffix = "中文模板.xlsx" if lang == "CN" else "英文模板.xlsx"
    template_files = list(TEMPLATE_DIR.glob(f"{INV_FORM_ID}*{template_suffix}"))
    if not template_files:
        print(f"  [错误] 找不到即时库存{lang_label}模板文件")
        return

    template_path = template_files[0]
    print(f"  [模板] {template_path.name}")

    wb_tpl = openpyxl.load_workbook(template_path)
    ws_tpl = wb_tpl[wb_tpl.sheetnames[0]]
    h0 = []
    for row_idx, row in enumerate(ws_tpl.iter_rows(values_only=True)):
        if row_idx == 0:
            h0 = [str(c).strip() if c else "" for c in row]
            break
    print(f"  [表头] {len(h0)} 列")

    if lang == "CN":
        output_name = "即时库存.xlsx"
        sheet_title = "Sheet1"
    else:
        output_name = "K3_Inventory.xlsx"
        sheet_title = "Sheet1"
    output_path = OUTPUT_DIR / output_name
    for old_file in OUTPUT_DIR.glob(f"{output_name.split('.')[0]}*.xlsx"):
        try:
            old_file.unlink()
        except PermissionError:
            print(f"  [警告] 文件被占用，跳过: {old_file.name}")

    client = KingdeeClient()

    # ── 1. 加载映射 ──────────────────────────────────────────────────────────
    print("  [映射] 加载批号映射...")
    lot_map = _load_generic_mapping("inv_lot_name_map")
    # 兼容旧格式：dict值展平为字符串
    for k, v in list(lot_map.items()):
        if isinstance(v, dict):
            lot_map[k] = v.get("cn", str(v))

    print("  [映射] 加载仓位映射...")
    stock_loc_map = _load_generic_mapping("instock_stock_loc_name_map")
    for k, v in list(stock_loc_map.items()):
        if isinstance(v, dict):
            stock_loc_map[k] = v.get("cn", str(v))

    print("  [映射] 加载弹性域维度映射...")
    flex_entries = _load_generic_mapping("flex_values_entries_map")

    # 清理映射中的异常条目
    for k in list(lot_map.keys()):
        if "Result" in k and "ResponseStatus" in k:
            del lot_map[k]
    for k in list(stock_loc_map.keys()):
        if "Result" in k and "ResponseStatus" in k:
            del stock_loc_map[k]

    print(f"    批号映射: {len(lot_map)} 个, 仓位映射: {len(stock_loc_map)} 个, 弹性域: {len(flex_entries)} 个")

    # ── 2. 合并 flex_entries 到 stock_loc_map ─────────────────────────────────
    flex_added = 0
    for k, v in flex_entries.items():
        if k not in stock_loc_map:
            name = v.get("cn", v.get("number", "")) if isinstance(v, dict) else str(v)
            stock_loc_map[k] = name
            flex_added += 1
    print(f"    flex合并: 新增 {flex_added} 个仓位映射")

    # ── 3. 查询 BD_FLEXVALUESDETAIL 仓位组合ID映射 ───────────────────────────
    print("  [映射] 查询 BD_FLEXVALUESDETAIL 仓位组合映射...")
    detail_fields = "FId,FF100001,FF100003,FF100004,FF100005,FF100006"
    detail_rows = []
    start = 0
    while True:
        try:
            rows = client.execute_bill_query({
                "FormId": "BD_FLEXVALUESDETAIL",
                "FieldKeys": detail_fields,
                "StartRow": start,
                "Limit": 500,
            })
        except Exception as e:
            logger.error("BD_FLEXVALUESDETAIL 查询失败 at %s: %s", start, e, exc_info=True)
            print(f"    [警告] BD_FLEXVALUESDETAIL 查询失败 at {start}: {e}")
            break
        if not rows:
            break
        detail_rows.extend(rows)
        start += 500
        if len(rows) < 500:
            break
        time.sleep(0.2)

    print(f"    BD_FLEXVALUESDETAIL: {len(detail_rows)} 条")

    def _resolve_dim(dim_id: str) -> str:
        dim_str = str(dim_id).strip()
        if not dim_str or dim_str in ("0", "None", ""):
            return ""
        if dim_str in flex_entries:
            entry = flex_entries[dim_str]
            if isinstance(entry, dict):
                return entry.get("cn", entry.get("number", dim_str))
            return str(entry)
        return dim_str

    detail_added = 0
    # bin_detail_map: FId -> {"shelf": 成品仓位编码, "raw": 原料仓位编码}
    bin_detail_map = {}
    for r in detail_rows:
        fid = str(r[0]).strip() if len(r) > 0 and r[0] else ""
        if not fid or fid in ("None", "0"):
            continue
        dim_names = []
        for idx in range(1, 6):
            raw_val = r[idx] if len(r) > idx else 0
            name = _resolve_dim(raw_val)
            if name:
                dim_names.append(name)
        combined = "-".join(dim_names) if dim_names else ""
        if combined and fid not in stock_loc_map:
            stock_loc_map[fid] = combined
            detail_added += 1
        # 分别提取成品仓位(FF100001)和原料仓位(FF100003)的编码
        shelf_dim = str(r[1]).strip() if len(r) > 1 and r[1] else ""
        raw_dim = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        shelf_name = _resolve_dim(shelf_dim) if shelf_dim and shelf_dim not in ("0", "None", "") else ""
        raw_name = _resolve_dim(raw_dim) if raw_dim and raw_dim not in ("0", "None", "") else ""
        bin_detail_map[fid] = {"shelf": shelf_name, "raw": raw_name}

    print(f"    BD_FLEXVALUESDETAIL合并: 新增={detail_added}, 总计={len(stock_loc_map)}")

    # ── 4. 分批查询 STK_Inventory 数据 ──────────────────────────────────────
    # STK_Inventory 对多字段查询有限制（超过6-7个字段只返回1行），
    # 每批最多6个字段（含3个主键: FMaterialId,FStockId,FLot），即3个数据字段
    # 主键: FMaterialId + FStockId + FLot（唯一标识一行库存）
    print("  [数据] 分批查询 STK_Inventory 数据...")

    def _fetch_inv_page(field_keys: str, start_row: int = 0, limit: int = 500) -> list:
        """分页查询 STK_Inventory"""
        all_rows = []
        start = start_row
        while True:
            try:
                rows = client.execute_bill_query({
                    "FormId": INV_FORM_ID,
                    "FieldKeys": field_keys,
                    "FilterString": "",
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": start,
                    "Limit": limit,
                })
            except Exception as e:
                logger.error("查询失败 at %s: %s", start, e, exc_info=True)
                print(f"    [警告] 查询失败 at {start}: {e}")
                break
            if not rows:
                break
            all_rows.extend(rows)
            start += limit
            if start % 5000 == 0:
                print(f"    已查询 {start} 条")
            if len(rows) < limit:
                break
            time.sleep(0.2)
        return all_rows

    KEY_FIELDS = "FMaterialId,FStockId,FLot,FStockLocId"

    # 用主键构建数据字典
    # 主键: FMaterialId + FStockId + FLot + FStockLocId（仓位）
    # 同一物料+仓库+批号可能分布在多个仓位，每个仓位各有不同库存量
    def _make_key(material_id, stock_id, lot_id, stock_loc_id=""):
        return f"{material_id}|{stock_id}|{lot_id}|{stock_loc_id}"

    data_dict = {}

    # 批次1: 物料编码/名称 + 仓库编码/名称 + 批号 + 数量（9字段，含子字段）
    batch1_fields = "FMaterialId,FMaterialId.FNumber,FMaterialId.FName,FStockId,FStockId.FNumber,FStockId.FName,FLot,FLot.FNumber,FQty"
    print(f"    批次1(基础): {len(batch1_fields.split(','))} 字段")
    batch1_rows = _fetch_inv_page(batch1_fields)
    print(f"    批次1: {len(batch1_rows)} 条")

    # 临时3-key dict（batch2获取FStockLocId后重建为4-key）
    _tmp_dict = {}
    for r in batch1_rows:
        material_id = str(r[0]).strip() if len(r) > 0 and r[0] else ""
        material_code = str(r[1]).strip() if len(r) > 1 and r[1] else ""
        material_name = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        stock_id = str(r[3]).strip() if len(r) > 3 and r[3] else ""
        stock_code = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        stock_name = str(r[5]).strip() if len(r) > 5 and r[5] else ""
        lot_id = str(r[6]).strip() if len(r) > 6 and r[6] else ""
        lot_number = str(r[7]).strip() if len(r) > 7 and r[7] else ""
        qty = r[8] if len(r) > 8 else 0

        key3 = f"{material_id}|{stock_id}|{lot_id}"
        if key3 not in _tmp_dict:
            _tmp_dict[key3] = {
                "FMaterialId": material_id,
                "material_code": material_code,
                "material_name": material_name,
                "FStockId": stock_id,
                "stock_code": stock_code,
                "stock_name": stock_name,
                "FLot": lot_id,
                "lot_number": lot_number,
                "FQty": qty,
            }

    # 从批次1构建lot映射
    new_lot = 0
    for r in batch1_rows:
        lot_id = str(r[6]).strip() if len(r) > 6 and r[6] else ""
        lot_number = str(r[7]).strip() if len(r) > 7 and r[7] else ""
        if lot_id and lot_number and lot_number not in ("None", "0", ""):
            if lot_id not in lot_map:
                lot_map[lot_id] = lot_number
                new_lot += 1
    print(f"    新增批号映射: {new_lot} 个，总计: {len(lot_map)} 个")

    # 批次2: 库存组织 + 仓位（6字段）—— 先执行以获取FStockLocId
    batch2_fields = "FMaterialId,FStockId,FLot,FStockOrgId,FStockOrgId.FNumber,FStockLocId"
    print(f"    批次2(组织+仓位): {len(batch2_fields.split(','))} 字段")
    batch2_rows = _fetch_inv_page(batch2_fields)
    print(f"    批次2: {len(batch2_rows)} 条")

    # 用4-key（含FStockLocId）重建data_dict
    for r in batch2_rows:
        material_id = str(r[0]).strip() if len(r) > 0 and r[0] else ""
        stock_id = str(r[1]).strip() if len(r) > 1 and r[1] else ""
        lot_id = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        stock_org_id = str(r[3]).strip() if len(r) > 3 and r[3] else ""
        stock_org_code = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        stock_loc_id = str(r[5]).strip() if len(r) > 5 and r[5] else ""

        key4 = _make_key(material_id, stock_id, lot_id, stock_loc_id)
        if key4 not in data_dict:
            # 从_tmp_dict复制基础信息
            key3 = f"{material_id}|{stock_id}|{lot_id}"
            base = _tmp_dict.get(key3, {})
            data_dict[key4] = {
                "FMaterialId": base.get("FMaterialId", material_id),
                "material_code": base.get("material_code", ""),
                "material_name": base.get("material_name", ""),
                "FStockId": base.get("FStockId", stock_id),
                "stock_code": base.get("stock_code", ""),
                "stock_name": base.get("stock_name", ""),
                "FLot": base.get("FLot", lot_id),
                "lot_number": base.get("lot_number", ""),
                "FQty": base.get("FQty", 0),
                "FStockOrgId": stock_org_id,
                "stock_org_code": stock_org_code,
                "FStockLocId": stock_loc_id,
            }

    # 补充：batch2中可能有些3-key在_tmp_dict中但不在batch2中（FStockLocId为空的情况）
    for key3, base in _tmp_dict.items():
        # 检查这个3-key是否已经被batch2覆盖（即是否有对应的4-key在data_dict中）
        found = any(k.startswith(key3 + "|") for k in data_dict)
        if not found:
            # 该3-key没有在batch2中出现，用空FStockLocId创建4-key
            key4 = _make_key(base.get("FMaterialId", ""), base.get("FStockId", ""), base.get("FLot", ""), "")
            if key4 not in data_dict:
                data_dict[key4] = {
                    **base,
                    "FStockOrgId": "",
                    "stock_org_code": "",
                    "FStockLocId": "",
                }

    del _tmp_dict  # 释放临时dict
    print(f"    4-key data_dict: {len(data_dict)} 条")

    # 批次1b: 单独查询FQty（5字段，含FStockLocId）
    batch1b_fields = f"{KEY_FIELDS},FQty"
    print(f"    批次1b(FQty单独): {len(batch1b_fields.split(','))} 字段")
    batch1b_rows = _fetch_inv_page(batch1b_fields)
    print(f"    批次1b: {len(batch1b_rows)} 条")
    qty_fixed_count = 0
    for r in batch1b_rows:
        material_id = str(r[0]).strip() if len(r) > 0 and r[0] else ""
        stock_id = str(r[1]).strip() if len(r) > 1 and r[1] else ""
        lot_id = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        stock_loc_id = str(r[3]).strip() if len(r) > 3 and r[3] else ""
        qty_val = r[4] if len(r) > 4 else 0
        key = _make_key(material_id, stock_id, lot_id, stock_loc_id)
        if key in data_dict and qty_val and qty_val != 0:
            data_dict[key]["FQty"] = qty_val
            qty_fixed_count += 1
    print(f"    批次1b修复FQty: {qty_fixed_count} 条（从0恢复为真实值）")

    # 批次3: 库存主单位 + 基本单位（6字段）
    batch3_fields = f"{KEY_FIELDS},FStockUnitId,FStockUnitId.FNumber,FBaseUnitId,FBaseUnitId.FNumber"
    print(f"    批次3(单位): {len(batch3_fields.split(','))} 字段")
    batch3_rows = _fetch_inv_page(batch3_fields)
    print(f"    批次3: {len(batch3_rows)} 条")

    for r in batch3_rows:
        material_id = str(r[0]).strip() if len(r) > 0 and r[0] else ""
        stock_id = str(r[1]).strip() if len(r) > 1 and r[1] else ""
        lot_id = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        stock_loc_id = str(r[3]).strip() if len(r) > 3 and r[3] else ""
        stock_unit_id = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        stock_unit_code = str(r[5]).strip() if len(r) > 5 and r[5] else ""
        base_unit_id = str(r[6]).strip() if len(r) > 6 and r[6] else ""
        base_unit_code = str(r[7]).strip() if len(r) > 7 and r[7] else ""

        key = _make_key(material_id, stock_id, lot_id, stock_loc_id)
        if key in data_dict:
            data_dict[key]["FStockUnitId"] = stock_unit_id
            data_dict[key]["stock_unit_code"] = stock_unit_code
            data_dict[key]["FBaseUnitId"] = base_unit_id
            data_dict[key]["base_unit_code"] = base_unit_code

    # 批次4: 库存状态（6字段，含状态子字段）
    batch4_fields = f"{KEY_FIELDS},FStockStatusId,FStockStatusId.FNumber,FStockStatusId.FName"
    print(f"    批次4(库存状态): {len(batch4_fields.split(','))} 字段")
    batch4_rows = _fetch_inv_page(batch4_fields)
    print(f"    批次4: {len(batch4_rows)} 条")

    for r in batch4_rows:
        material_id = str(r[0]).strip() if len(r) > 0 and r[0] else ""
        stock_id = str(r[1]).strip() if len(r) > 1 and r[1] else ""
        lot_id = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        stock_loc_id = str(r[3]).strip() if len(r) > 3 and r[3] else ""
        stock_status_id = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        stock_status_number = str(r[5]).strip() if len(r) > 5 and r[5] else ""
        stock_status_fname = str(r[6]).strip() if len(r) > 6 and r[6] else ""

        key = _make_key(material_id, stock_id, lot_id, stock_loc_id)
        if key in data_dict:
            data_dict[key]["FStockStatusId"] = stock_status_id
            data_dict[key]["FStockStatusNumber"] = stock_status_number
            data_dict[key]["FStockStatusFName"] = stock_status_fname

    # 批次4b: 辅助属性 + 基本量（6字段）
    batch4b_fields = f"{KEY_FIELDS},FAuxPropId,FBaseQty,FAVBQty"
    print(f"    批次4b(辅助+基本量+可用量): {len(batch4b_fields.split(','))} 字段")
    batch4b_rows = _fetch_inv_page(batch4b_fields)
    print(f"    批次4b: {len(batch4b_rows)} 条")

    for r in batch4b_rows:
        material_id = str(r[0]).strip() if len(r) > 0 and r[0] else ""
        stock_id = str(r[1]).strip() if len(r) > 1 and r[1] else ""
        lot_id = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        stock_loc_id = str(r[3]).strip() if len(r) > 3 and r[3] else ""
        aux_prop_id = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        base_qty = r[5] if len(r) > 5 else 0
        avb_qty = r[6] if len(r) > 6 else 0

        key = _make_key(material_id, stock_id, lot_id, stock_loc_id)
        if key in data_dict:
            data_dict[key]["FAuxPropId"] = aux_prop_id
            # FBaseQty/FAVBQty: 同一key可能有多条API记录，
            # 部分记录返回0（API限制），部分返回真实值。
            # 只用非0值更新，避免0覆盖真实值。
            if base_qty and base_qty != 0:
                old_base_qty = data_dict[key].get("FBaseQty", 0)
                if not old_base_qty or old_base_qty == 0 or base_qty > old_base_qty:
                    data_dict[key]["FBaseQty"] = base_qty
            if avb_qty and avb_qty != 0:
                old_avb_qty = data_dict[key].get("FAVBQty", 0)
                if not old_avb_qty or old_avb_qty == 0 or avb_qty > old_avb_qty:
                    data_dict[key]["FAVBQty"] = avb_qty

    # 批次5: 生产日期 + 有效期至（5字段）
    batch5_fields = f"{KEY_FIELDS},FProduceDate,FExpiryDate"
    print(f"    批次5(日期): {len(batch5_fields.split(','))} 字段")
    batch5_rows = _fetch_inv_page(batch5_fields)
    print(f"    批次5: {len(batch5_rows)} 条")

    for r in batch5_rows:
        material_id = str(r[0]).strip() if len(r) > 0 and r[0] else ""
        stock_id = str(r[1]).strip() if len(r) > 1 and r[1] else ""
        lot_id = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        stock_loc_id = str(r[3]).strip() if len(r) > 3 and r[3] else ""
        produce_date = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        expiry_date = str(r[5]).strip() if len(r) > 5 and r[5] else ""

        key = _make_key(material_id, stock_id, lot_id, stock_loc_id)
        if key in data_dict:
            data_dict[key]["FProduceDate"] = produce_date
            data_dict[key]["FExpiryDate"] = expiry_date

    # 批次6: 基本可用量（5字段）
    batch6_fields = f"{KEY_FIELDS},FBaseAVBQty"
    print(f"    批次6(基本可用量): {len(batch6_fields.split(','))} 字段")
    batch6_rows = _fetch_inv_page(batch6_fields)
    print(f"    批次6: {len(batch6_rows)} 条")

    for r in batch6_rows:
        material_id = str(r[0]).strip() if len(r) > 0 and r[0] else ""
        stock_id = str(r[1]).strip() if len(r) > 1 and r[1] else ""
        lot_id = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        stock_loc_id = str(r[3]).strip() if len(r) > 3 and r[3] else ""
        base_avb_qty = r[4] if len(r) > 4 else 0

        key = _make_key(material_id, stock_id, lot_id, stock_loc_id)
        if key in data_dict:
            # FBaseAVBQty: 同一key可能有多条API记录，只用非0值更新
            if base_avb_qty and base_avb_qty != 0:
                old_val = data_dict[key].get("FBaseAVBQty", 0)
                if not old_val or old_val == 0 or base_avb_qty > old_val:
                    data_dict[key]["FBaseAVBQty"] = base_avb_qty

    print(f"    合并后总数据: {len(data_dict)} 条")

    # ── 4a. 从BD_MATERIAL获取库存单位→基本单位的换算率 ──────────────────────
    # 当FQty仍为0且stock_unit != base_unit时，需要换算率来计算FQty = FBaseQty / 换算率
    # 收集需要查询换算率的物料ID
    material_ids_needing_rate = set()
    for key, d in data_dict.items():
        qty = d.get("FQty", 0)
        base_qty = d.get("FBaseQty", 0)
        stock_unit_code = d.get("stock_unit_code", "")
        base_unit_code = d.get("base_unit_code", "")
        if (not qty or qty == 0) and base_qty and base_qty != 0 and stock_unit_code != base_unit_code:
            material_ids_needing_rate.add(d.get("FMaterialId", ""))

    unit_convert_map = {}  # material_id → convert_rate (stock_unit → base_unit)
    if material_ids_needing_rate:
        print(f"    查询单位换算率: {len(material_ids_needing_rate)} 个物料（stock_unit != base_unit 且 FQty=0）")
        mid_list = sorted([m for m in material_ids_needing_rate if m and m not in ("0", "None", "")])
        for i in range(0, len(mid_list), 100):
            batch_mids = mid_list[i:i+100]
            filter_str = " or ".join([f"FMaterialId = {mid}" for mid in batch_mids])
            try:
                rate_rows = client.execute_bill_query({
                    "FormId": "BD_MATERIAL",
                    "FieldKeys": "FMaterialId,FStoreUnitID.FConvertRate,FBaseUnitId.FNumber,FStoreUnitID.FNumber",
                    "FilterString": filter_str,
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": 200,
                })
                if rate_rows:
                    for rr in rate_rows:
                        mid = str(rr[0]).strip() if rr[0] else ""
                        convert_rate = rr[1] if len(rr) > 1 and rr[1] else None
                        if mid and convert_rate:
                            try:
                                unit_convert_map[mid] = float(convert_rate)
                            except (ValueError, TypeError):
                                pass
            except Exception as e:
                logger.error("查询单位换算率失败: %s", e, exc_info=True)
                print(f"    [警告] 查询单位换算率失败: {e}")
            time.sleep(0.2)
        print(f"    单位换算率: 获取到 {len(unit_convert_map)} 个")

    # 应用换算率修复FQty
    rate_fixed_count = 0
    fallback_count = 0
    for key, d in data_dict.items():
        qty = d.get("FQty", 0)
        base_qty = d.get("FBaseQty", 0)
        stock_unit_code = d.get("stock_unit_code", "")
        base_unit_code = d.get("base_unit_code", "")
        if (not qty or qty == 0) and base_qty and base_qty != 0:
            if stock_unit_code != base_unit_code:
                mid = d.get("FMaterialId", "")
                rate = unit_convert_map.get(mid)
                if rate and rate != 0:
                    d["FQty"] = base_qty / rate
                    rate_fixed_count += 1
                else:
                    d["FQty"] = base_qty  # 兜底：用FBaseQty确保行不被过滤
                    fallback_count += 1
            else:
                d["FQty"] = base_qty  # 单位相同时直接用FBaseQty
                fallback_count += 1
    if rate_fixed_count:
        print(f"    换算率修复FQty: {rate_fixed_count} 条（FBaseQty / 换算率）")
    if fallback_count:
        print(f"    兜底修复FQty: {fallback_count} 条（FBaseQty直接赋值）")

    # 统计FQty修复情况
    fqty_zero_after_repair = sum(1 for d in data_dict.values() if not d.get("FQty", 0) or d.get("FQty", 0) == 0)
    fqty_nonzero_after_repair = len(data_dict) - fqty_zero_after_repair
    print(f"    FQty修复后: {fqty_nonzero_after_repair} 条非0, {fqty_zero_after_repair} 条仍为0")

    # ── 4b. 通过 SAL_ExpectQtyQueryReport 获取真实可用量 ──────────────────────
    # 金蝶 ExecuteBillQuery 对 STK_Inventory 的 FAVBQty/FBaseAVBQty 始终返回0，
    # 但 SAL_ExpectQtyQueryReport（预计可发量查询报表）通过 GetSysReportData
    # 可以返回真实的可用量数据，且按物料+仓库组维度分组。
    # 关联方式：BD_STOCK.FName = 报表 FStockId（仓库组名），
    #          BD_STOCK.FNumber = STK_Inventory FStockId.FNumber（仓库编码）
    print("  [报表] 查询 SAL_ExpectQtyQueryReport 可用量数据...")
    try:
        sdk = client._ensure_sdk()
        rpt_avb_rows = []
        start = 0
        while True:
            result = sdk.GetSysReportData("SAL_ExpectQtyQueryReport", json.dumps({
                "FieldKeys": "FMasterId,FMaterialName,FStockId,FAVBQTY",
                "StartRow": start,
                "Limit": 5000,
                "FilterString": [],
                "Model": {
                    "FCheckStockId": "true",
                    "FMulMaterialIds": [],
                    "FMulStockIds": [],
                    "FStockOrgIds": [],
                }
            }))
            parsed = json.loads(result)
            result_data = parsed.get("Result", {})
            if not result_data.get("IsSuccess", False):
                errors = result_data.get("ResponseStatus", {}).get("Errors", [])
                err_msg = errors[0].get("Message", "") if errors else "unknown"
                print(f"    [警告] GetSysReportData 失败: {err_msg}")
                break
            batch_rows = result_data.get("Rows", [])
            rpt_avb_rows.extend(batch_rows)
            if len(batch_rows) < 5000:
                break
            start += 5000
            time.sleep(0.3)
        print(f"    报表返回 {len(rpt_avb_rows)} 条可用量数据")

        # 查询 BD_STOCK 建立 仓库编码→仓库组名(FName) 映射
        stock_code_to_group = {}
        stock_rows = client.execute_bill_query({
            "FormId": "BD_STOCK",
            "FieldKeys": "FNumber,FName",
            "FilterString": "",
            "OrderString": "",
            "TopRowCount": 0,
            "StartRow": 0,
            "Limit": 200,
        })
        if stock_rows:
            for sr in stock_rows:
                s_code = str(sr[0]).strip() if sr[0] else ""
                s_name = str(sr[1]).strip() if sr[1] else ""
                if s_code and s_name:
                    stock_code_to_group[s_code] = s_name
        print(f"    仓库编码→仓库组名映射: {len(stock_code_to_group)} 个")

        # 构建报表可用量索引: (物料编码, 仓库组名) → FAVBQTY
        rpt_avb_map = {}
        for row in rpt_avb_rows:
            mat_no = str(row[0]).strip() if row[0] else ""
            stock_group = str(row[2]).strip() if row[2] else ""
            avb_str = str(row[3]).strip() if row[3] else "0"
            try:
                avb_val = float(avb_str.replace(",", "")) if avb_str and avb_str not in ("None", "") else 0
            except (ValueError, TypeError):
                avb_val = 0
            if mat_no and stock_group:
                rpt_key = f"{mat_no}|{stock_group}"
                # 同一物料+仓库组可能有多条（不同辅助属性），取合计
                if rpt_key not in rpt_avb_map:
                    rpt_avb_map[rpt_key] = 0
                rpt_avb_map[rpt_key] += avb_val
        print(f"    报表可用量索引: {len(rpt_avb_map)} 条")

        # 将报表可用量匹配到 data_dict
        # 对于同一物料+仓库组下有多条库存记录（不同批号/仓位），
        # 按各记录的 FBaseQty 占比分配可用量
        avb_matched = 0
        avb_not_matched = 0

        # 先按 (物料编码, 仓库组名) 分组
        group_items = {}
        for key, d in data_dict.items():
            mat_no = d.get("material_code", "")
            stock_code = d.get("stock_code", "")
            group_name = stock_code_to_group.get(stock_code, stock_code)
            grp_key = f"{mat_no}|{group_name}"
            if grp_key not in group_items:
                group_items[grp_key] = []
            group_items[grp_key].append(key)

        for grp_key, keys in group_items.items():
            if grp_key not in rpt_avb_map:
                avb_not_matched += len(keys)
                continue
            total_avb = rpt_avb_map[grp_key]
            if not total_avb or total_avb == 0:
                avb_not_matched += len(keys)
                continue

            # 计算该组内各记录的 FBaseQty 合计，按比例分配可用量
            total_base = sum(data_dict[k].get("FBaseQty", 0) or 0 for k in keys)
            if not total_base or total_base == 0:
                # FBaseQty 也为0，均匀分配
                per_item = total_avb / len(keys)
                for k in keys:
                    data_dict[k]["_rpt_avb_qty"] = per_item
                    data_dict[k]["_rpt_base_avb_qty"] = per_item
                avb_matched += len(keys)
            else:
                for k in keys:
                    item_base = data_dict[k].get("FBaseQty", 0) or 0
                    ratio = item_base / total_base
                    data_dict[k]["_rpt_avb_qty"] = total_avb * ratio
                    data_dict[k]["_rpt_base_avb_qty"] = total_avb * ratio
                avb_matched += len(keys)

        print(f"    可用量匹配: {avb_matched} 条成功, {avb_not_matched} 条未匹配")
    except Exception as e:
        logger.error("SAL_ExpectQtyQueryReport 查询失败: %s", e, exc_info=True)
        print(f"    [警告] SAL_ExpectQtyQueryReport 查询失败: {e}")

    # ── 5. 解析ID映射 ──────────────────────────────────────────────────────
    # 收集需要映射的ID（库存状态已在批次4中直接查询FNumber/FName，无需额外映射）
    stock_org_ids = set()
    stock_unit_ids = set()
    base_unit_ids = set()
    aux_prop_ids = set()

    for d in data_dict.values():
        if d.get("FStockOrgId") and d["FStockOrgId"] not in ("0", "0.0", "None", ""):
            stock_org_ids.add(d["FStockOrgId"])
        if d.get("FStockUnitId") and d["FStockUnitId"] not in ("0", "0.0", "None", ""):
            stock_unit_ids.add(d["FStockUnitId"])
        if d.get("FBaseUnitId") and d["FBaseUnitId"] not in ("0", "0.0", "None", ""):
            base_unit_ids.add(d["FBaseUnitId"])
        if d.get("FAuxPropId") and d["FAuxPropId"] not in ("0", "0.0", "None", ""):
            aux_prop_ids.add(d["FAuxPropId"])

    # 组织映射
    org_code_map = _load_generic_mapping("org_code_map")
    org_name_map = _load_generic_mapping("org_name_map")
    missing_org_ids = [oid for oid in stock_org_ids if oid not in org_code_map]
    if missing_org_ids:
        print(f"    需要解析 {len(missing_org_ids)} 个库存组织ID")
        new_org_map = fetch_generic_mapping(client, "ORG_Organizations", "FOrgId", "FNumber", "org_code_map", missing_org_ids)
        org_code_map.update(new_org_map)
        new_org_names = _fetch_name_map(client, "ORG_Organizations", list(set(new_org_map.values())), "org_name_map")
        org_name_map.update(new_org_names)

    # 库存状态映射 - 已在批次4中直接查询 FStockStatusId.FNumber 和 FStockStatusId.FName
    # 无需额外映射查询，直接使用 data_dict 中的 FStockStatusNumber 和 FStockStatusFName

    # 单位映射
    all_unit_ids = list(set(stock_unit_ids) | set(base_unit_ids))
    unit_code_map = _load_generic_mapping("unit_code_map")
    unit_name_map = _load_generic_mapping("unit_name_map")
    missing_unit_ids = [uid for uid in all_unit_ids if uid not in unit_code_map]
    if missing_unit_ids:
        print(f"    需要解析 {len(missing_unit_ids)} 个单位ID")
        new_unit_map = fetch_generic_mapping(client, "BD_UNIT", "FUnitID", "FNumber", "unit_code_map", missing_unit_ids)
        unit_code_map.update(new_unit_map)
        new_unit_names = _fetch_name_map(client, "BD_UNIT", list(set(new_unit_map.values())), "unit_name_map")
        unit_name_map.update(new_unit_names)

    # 辅助属性映射（通过 BD_FLEXVALUESDETAIL 查询 F100001 维度值）
    aux_prop_name_map = {}
    if aux_prop_ids:
        print(f"    需要解析 {len(aux_prop_ids)} 个辅助属性ID")
        aux_prop_name_map = _load_generic_mapping("inv_aux_prop_name_map")
        for k, v in list(aux_prop_name_map.items()):
            if isinstance(v, dict):
                aux_prop_name_map[k] = v.get("cn", str(v))
        missing_aux = [aid for aid in aux_prop_ids if aid not in aux_prop_name_map]
        if missing_aux:
            print(f"    需要查询 {len(missing_aux)} 个辅助属性名称")
            # FAuxPropId 在 STK_Inventory 中是弹性域分录ID，查询 BD_FLEXVALUESDETAIL
            batch_size = 200
            for i in range(0, len(missing_aux), batch_size):
                batch = missing_aux[i:i + batch_size]
                filter_parts = [f"FId = {aid}" for aid in batch]
                filter_str = " or ".join(filter_parts)
                try:
                    rows = client.execute_bill_query({
                        "FormId": "BD_FLEXVALUESDETAIL",
                        "FieldKeys": "FId,FF100001",
                        "FilterString": filter_str,
                        "OrderString": "",
                        "TopRowCount": 0,
                        "StartRow": 0,
                        "Limit": len(batch) + 10,
                    })
                    if rows:
                        for row in rows:
                            aid = str(row[0]).strip() if row[0] else ""
                            f100001_val = str(row[1]).strip() if row[1] else ""
                            if aid and f100001_val and f100001_val not in ("None", "0"):
                                # f100001_val 是维度值ID，需要进一步解析
                                resolved = _resolve_dim(f100001_val)
                                if resolved:
                                    aux_prop_name_map[aid] = resolved
                except Exception:
                    pass
            _save_generic_mapping("inv_aux_prop_name_map", aux_prop_name_map)

    # 物料规格映射
    material_spec_map = _load_generic_mapping("material_spec_map")

    # 物料保质期映射（从BD_MATERIAL获取FExpPeriod，而非动态计算）
    material_expiry_period_map = _load_generic_mapping("material_expiry_period_map")
    all_material_codes = set(d.get("material_code", "") for d in data_dict.values())
    missing_expiry_codes = [c for c in all_material_codes if c and c not in material_expiry_period_map]
    if missing_expiry_codes:
        print(f"    查询物料保质期: {len(missing_expiry_codes)} 个物料")
        for i in range(0, len(missing_expiry_codes), 100):
            batch = missing_expiry_codes[i:i+100]
            filter_str = " or ".join([f"FNumber = '{c}'" for c in batch])
            try:
                rows = client.execute_bill_query({
                    "FormId": "BD_MATERIAL",
                    "FieldKeys": "FNumber,FExpPeriod",
                    "FilterString": filter_str,
                    "OrderString": "",
                    "TopRowCount": 0,
                    "StartRow": 0,
                    "Limit": 200,
                })
                if rows:
                    for row in rows:
                        code = str(row[0]).strip() if row[0] else ""
                        exp_period = row[1] if len(row) > 1 and row[1] else ""
                        if code:
                            material_expiry_period_map[code] = str(exp_period) if exp_period else ""
            except Exception:
                pass
            time.sleep(0.2)
        _save_generic_mapping("material_expiry_period_map", material_expiry_period_map)
    print(f"    物料保质期映射: {len(material_expiry_period_map)} 个")

    # ── 6. 构建字段映射 ──────────────────────────────────────────────────────
    # 模板22列到数据字段的映射
    inv_field_map = {
        "序号": "_seq", "S/N": "_seq",
        "物料编码": "material_code", "Material Code": "material_code",
        "物料名称": "material_name", "Material Name": "material_name",
        "规格型号": "_spec", "Specification": "_spec",
        "辅助属性": "_aux_prop", "Secondary Property": "_aux_prop",
        "仓库名称": "stock_name", "Warehouse Name": "stock_name",
        "批号": "_lot", "Lot No.": "_lot",
        "库存主单位": "_stock_unit_name", "Stock Unit (CUoM)": "_stock_unit_name",
        "仓库编码": "stock_code", "Warehouse Code": "stock_code",
        "库存量(主单位)": "FQty", "Inventory Qty  (CUoM)": "FQty",
        "库存组织": "_org_name", "Inventory Org.": "_org_name",
        "生产日期": "FProduceDate", "Production Date": "FProduceDate",
        "有效期至": "FExpiryDate", "Valid till": "FExpiryDate",
        "仓位": "_bin", "Bin": "_bin",
        "库存状态": "_stock_status", "Inventory Status": "_stock_status",
        "仓位.成品仓位.编码": "_bin_shelf_code", "Bin.Shelf.Code": "_bin_shelf_code",
        "仓位.原料仓位.编码": "_bin_raw_code", "Bin.原料仓位.Code": "_bin_raw_code",
        "保质期": "_guarantee_period", "Guarantee Period": "_guarantee_period",
        "基本单位": "_base_unit_name", "BUoM": "_base_unit_name",
        "库存量(基本单位)": "FBaseQty", "Inventory Qty (BUoM)": "FBaseQty",
        "可用量(主单位)": "FAVBQty", "Available Qty (CUoM))": "FAVBQty",
        "可用量(基本单位)": "FBaseAVBQty", "Available Qty (BUoM)": "FBaseAVBQty",
    }

    date_cols = {"FProduceDate", "FExpiryDate"}

    # ── 7. 处理数据并应用映射 ────────────────────────────────────────────────
    print("  [处理] 应用映射...")

    # 仓库名称英文映射（通过View API获取BD_STOCK的多语言名称）
    stock_name_en_map = {}
    if lang != "CN":
        print("    获取仓库名称英文翻译...")
        stock_name_en_map = _load_generic_mapping("stock_name_en_map")
        # 收集所有仓库编码
        all_stock_codes = set(d.get("stock_code", "") for d in data_dict.values())
        missing_stock_codes = [c for c in all_stock_codes if c and c not in stock_name_en_map]
        if missing_stock_codes:
            import json as _json
            for code in missing_stock_codes:
                try:
                    sid_rows = client.execute_bill_query({
                        "FormId": "BD_STOCK",
                        "FieldKeys": "FStockId,FNumber",
                        "FilterString": f"FNumber = '{code}'",
                        "OrderString": "", "TopRowCount": 0, "StartRow": 0, "Limit": 1,
                    })
                    if sid_rows:
                        result = client.view("BD_STOCK", {"Id": sid_rows[0][0]})
                        if isinstance(result, str):
                            result = _json.loads(result)
                        # 递归搜索Name字段中的英文值
                        def _find_en_name(obj):
                            if isinstance(obj, dict):
                                if "Name" in obj and isinstance(obj["Name"], list):
                                    for item in obj["Name"]:
                                        if item.get("Key") == 1033:
                                            return item.get("Value", "")
                                for v in obj.values():
                                    r = _find_en_name(v)
                                    if r:
                                        return r
                            elif isinstance(obj, list):
                                for item in obj:
                                    r = _find_en_name(item)
                                    if r:
                                        return r
                            return ""
                        en_name = _find_en_name(result)
                        if en_name:
                            stock_name_en_map[code] = en_name
                except Exception:
                    pass
            _save_generic_mapping("stock_name_en_map", stock_name_en_map)
        print(f"    仓库名称英文映射: {len(stock_name_en_map)} 个")

    # 排序数据，按物料编码排序
    sorted_items = sorted(data_dict.items(), key=lambda x: x[1].get("material_code", ""))

    lot_ok = lot_fail = bin_ok = bin_fail = zero_qty_skip = 0
    output_seq = 0  # 连续序号计数器

    workbook = openpyxl.Workbook()
    ws_out = workbook.active
    ws_out.title = sheet_title
    ws_out.append(h0)

    for key, d in sorted_items:
        output_row = [""] * len(h0)

        # 解析各字段值
        material_code = d.get("material_code", "")
        material_name = d.get("material_name", "")
        stock_name = d.get("stock_name", "")
        stock_code = d.get("stock_code", "")
        lot_id = d.get("FLot", "")
        lot_number = d.get("lot_number", "")
        qty = d.get("FQty", 0)
        base_qty = d.get("FBaseQty", 0)
        avb_qty = d.get("FAVBQty", 0)
        base_avb_qty = d.get("FBaseAVBQty", 0)
        stock_org_id = d.get("FStockOrgId", "")
        stock_org_code = d.get("stock_org_code", "")
        stock_loc_id = d.get("FStockLocId", "")
        stock_status_id = d.get("FStockStatusId", "")
        stock_unit_id = d.get("FStockUnitId", "")
        stock_unit_code = d.get("stock_unit_code", "")
        base_unit_id = d.get("FBaseUnitId", "")
        base_unit_code = d.get("base_unit_code", "")

        # 修复金蝶API对STK_Inventory数量字段返回0的问题：
        # ExecuteBillQuery 对 STK_Inventory 的 FQty/FAVBQty/FBaseAVBQty/FLockQty
        # 全部返回0，仅 FBaseQty 有真实数据。
        # 修复策略：
        # 1. 当 stock_unit == base_unit 时，FQty = FBaseQty
        # 2. 当 stock_unit != base_unit 时，通过BD_MATERIAL换算率计算 FQty = FBaseQty / rate
        # 3. 若换算率不可用，仍用 FBaseQty 作为 FQty（保证行不被过滤）
        if (not qty or qty == 0) and base_qty and base_qty != 0:
            if stock_unit_code == base_unit_code:
                qty = base_qty
            else:
                # 尝试用换算率计算，若不可用则用FBaseQty兜底
                mid = d.get("FMaterialId", "")
                rate = unit_convert_map.get(mid) if mid in unit_convert_map else None
                if rate and rate != 0:
                    qty = base_qty / rate
                else:
                    qty = base_qty  # 兜底：用FBaseQty确保行不被过滤
        # 可用量修复：优先使用报表数据
        rpt_avb = d.get("_rpt_avb_qty")
        rpt_base_avb = d.get("_rpt_base_avb_qty")
        if rpt_avb is not None and rpt_avb != 0:
            avb_qty = rpt_avb
        elif (not avb_qty or avb_qty == 0):
            if stock_unit_code == base_unit_code:
                avb_qty = qty  # FAVBQty = FQty（已修复的值）
        if rpt_base_avb is not None and rpt_base_avb != 0:
            base_avb_qty = rpt_base_avb
        elif (not base_avb_qty or base_avb_qty == 0) and base_qty and base_qty != 0:
            base_avb_qty = base_qty  # FBaseAVBQty = FBaseQty

        # 过滤：库存量为0的行不需要输出
        # 注意：FQty始终为0是API限制，应以FBaseQty为准判断是否有库存
        if (not base_qty or base_qty == 0) and (not qty or qty == 0):
            zero_qty_skip += 1
            continue

        aux_prop_id = d.get("FAuxPropId", "")
        produce_date = d.get("FProduceDate", "")
        expiry_date = d.get("FExpiryDate", "")

        # 仓库名称（英文版使用英文名称）
        if lang != "CN" and stock_name_en_map:
            en_stock_name = stock_name_en_map.get(stock_code, "")
            if en_stock_name:
                stock_name = en_stock_name

        # Lot映射
        lot_val = lot_number
        if lot_id and lot_id in lot_map:
            lot_val = lot_map[lot_id]
            lot_ok += 1
        elif lot_id and lot_id not in ("0", "0.0", "None", ""):
            lot_fail += 1

        # Bin映射
        bin_val = ""
        if stock_loc_id and stock_loc_id not in ("None", "0", ""):
            if stock_loc_id in stock_loc_map:
                bin_val = stock_loc_map[stock_loc_id]
                bin_ok += 1 if bin_val else 0
                bin_fail += 1 if not bin_val else 0
            else:
                bin_fail += 1

        # 组织名称
        org_name = org_name_map.get(stock_org_code, stock_org_code) if stock_org_code else ""

        # 库存状态名称（英文版使用英文名称，中文版使用FName名称）
        stock_status_name = ""
        if stock_status_id and stock_status_id not in ("0", "0.0", "None", ""):
            if lang == "CN":
                stock_status_name = d.get("FStockStatusFName", "")
            else:
                # 英文版：将FNumber编码映射为英文名称
                status_number = d.get("FStockStatusNumber", "")
                stock_status_name = _stock_status_en_map.get(status_number, status_number)

        # 单位名称（英文版使用FNumber编码，中文版使用FName名称）
        if lang == "CN":
            stock_unit_name = unit_name_map.get(stock_unit_code, stock_unit_code) if stock_unit_code else ""
            base_unit_name = unit_name_map.get(base_unit_code, base_unit_code) if base_unit_code else ""
        else:
            stock_unit_name = stock_unit_code if stock_unit_code else ""
            base_unit_name = base_unit_code if base_unit_code else ""

        # 辅助属性名称
        aux_prop_name = aux_prop_name_map.get(aux_prop_id, "") if aux_prop_id else ""

        # 物料规格
        spec = material_spec_map.get(material_code, "")

        # 仓位子字段（成品仓位编码、原料仓位编码）
        bin_shelf_code = ""
        bin_raw_code = ""
        if stock_loc_id and stock_loc_id not in ("None", "0", ""):
            if stock_loc_id in bin_detail_map:
                bin_shelf_code = bin_detail_map[stock_loc_id].get("shelf", "")
                bin_raw_code = bin_detail_map[stock_loc_id].get("raw", "")

        # 保质期：从BD_MATERIAL的FExpPeriod获取（物料主数据定义的保质期天数）
        guarantee_period = material_expiry_period_map.get(material_code, "")

        # 构建字段值字典
        output_seq += 1
        field_values = {
            "_seq": output_seq,
            "material_code": material_code,
            "material_name": material_name,
            "_spec": spec,
            "_aux_prop": aux_prop_name,
            "stock_name": stock_name,
            "_lot": lot_val,
            "_stock_unit_name": stock_unit_name,
            "stock_code": stock_code,
            "FQty": qty,
            "_org_name": org_name,
            "FProduceDate": _format_date(produce_date) if produce_date and produce_date not in ("None", "") else "",
            "FExpiryDate": _format_date(expiry_date) if expiry_date and expiry_date not in ("None", "") else "",
            "_bin": bin_val,
            "_stock_status": stock_status_name,
            "_bin_shelf_code": bin_shelf_code,
            "_bin_raw_code": bin_raw_code,
            "_guarantee_period": guarantee_period,
            "_base_unit_name": base_unit_name,
            "FBaseQty": base_qty,
            "FAVBQty": avb_qty,
            "FBaseAVBQty": base_avb_qty,
        }

        # 数值字段：始终输出（包括0.0）
        numeric_data_keys = {"FQty", "FBaseQty", "FAVBQty", "FBaseAVBQty"}

        # 按模板映射填充
        for tpl_idx, tpl_field in enumerate(h0):
            if not tpl_field:
                continue
            if tpl_field in inv_field_map:
                data_key = inv_field_map[tpl_field]
                val = field_values.get(data_key, "")
                if val is None:
                    continue
                # 数值字段始终输出
                if isinstance(val, (int, float)):
                    output_row[tpl_idx] = val
                elif isinstance(val, datetime):
                    output_row[tpl_idx] = val
                elif str(val).strip() not in ("None", "0.0", ""):
                    output_row[tpl_idx] = _clean_string(str(val).strip())

        ws_out.append(output_row)

    # 为日期列设置单元格格式
    for tpl_idx, tpl_field in enumerate(h0):
        if not tpl_field:
            continue
        if tpl_field in inv_field_map:
            data_key = inv_field_map[tpl_field]
            if data_key in date_cols:
                fmt = "yyyy/M/d"
                for row_idx in range(2, ws_out.max_row + 1):
                    cell = ws_out.cell(row=row_idx, column=tpl_idx + 1)
                    if isinstance(cell.value, datetime):
                        cell.number_format = fmt

    _save_workbook(workbook, output_path)
    print(f"    Lot: {lot_ok} 映射成功, {lot_fail} 未映射")
    print(f"    Bin: {bin_ok} 映射成功, {bin_fail} 未映射")
    if zero_qty_skip:
        print(f"    过滤FQty=0的行: {zero_qty_skip} 条（已跳过）")
    print(f"  [完成] {output_path.name} ({len(data_dict) - zero_qty_skip} 行)")

    # ── 8. 保存映射 ──────────────────────────────────────────────────────────
    _save_generic_mapping("inv_lot_name_map", lot_map)
    _save_generic_mapping("instock_stock_loc_name_map", stock_loc_map)

    print(f"[完成] 生成即时库存 Excel ({lang_label})")


# ══════════════════════════════════════════════════════════════════════════════
# PRD_PickMtrl 生产领料单 Excel 生成
# ══════════════════════════════════════════════════════════════════════════════


def generate_pickmtrl_excel(lang: str = "CN") -> None:
    """生成 PRD_PickMtrl 生产领料单 Excel（使用模板+CSV数据）。

    模板21列：
    中文: 单据编号, 领料人, 日期, 单据状态, 发料组织, 生产组织, 物料编码,
          物料名称, 车间, 规格型号, 单位, 成本价, 总成本, 申请数量,
          实发数量, 仓库, 批号, 辅助属性, 生产订单编号, 生产订单行号, 需求来源
    英文: Doc No., Picked by, Date, Doc Status, Issue Org., Production Org.,
          Material Code, Material Name, Workshop, Specification, UoM, U/P,
          Amount, Requested Qty, Actual Ship Qty, Warehouse, Lot No.,
          Secondary Property, MO No., MO Row No., Requirement Source
    """
    lang_label = "中文" if lang == "CN" else "英文"
    print(f"\n[开始] 生成生产领料单 Excel ({lang_label})")
    OUTPUT_DIR.mkdir(exist_ok=True)

    # 读取模板表头
    template_suffix = "中文模板.xlsx" if lang == "CN" else "英文模板.xlsx"
    template_files = list(TEMPLATE_DIR.glob(f"{PICKMTRL_FORM_ID}*{template_suffix}"))
    if not template_files:
        print(f"  [错误] 找不到生产领料单{lang_label}模板文件")
        return

    template_path = template_files[0]
    print(f"  [模板] {template_path.name}")

    wb_tpl = openpyxl.load_workbook(template_path)
    ws_tpl = wb_tpl[wb_tpl.sheetnames[0]]
    h0 = []
    for row_idx, row in enumerate(ws_tpl.iter_rows(values_only=True)):
        if row_idx == 0:
            h0 = [str(c).strip() if c else "" for c in row]
            break
    wb_tpl.close()
    print(f"  [表头] {len(h0)} 列: {h0}")

    # 查找最新的CSV数据文件
    csv_files = list(DATA_DIR.glob(f"{PICKMTRL_FORM_ID}_*.csv"))
    if not csv_files:
        print(f"  [错误] 找不到生产领料单数据文件，请先运行 --form PRD_PickMtrl 拉取数据")
        return
    csv_path = sorted(csv_files, key=lambda x: x.stat().st_mtime, reverse=True)[0]
    print(f"  [数据] {csv_path.name}")

    df = pd.read_csv(csv_path, dtype=str)
    data_cols = df.columns.tolist()
    print(f"  [数据] {len(df)} 行, {len(data_cols)} 列")

    if lang == "CN":
        output_name = "生产领料单列表.xlsx"
        sheet_title = "Sheet1"
    else:
        output_name = "K3_18015_PickingList.xlsx"
        sheet_title = "Sheet1"
    output_path = OUTPUT_DIR / output_name
    for old_file in OUTPUT_DIR.glob(f"{output_name.split('.')[0]}*.xlsx"):
        try:
            old_file.unlink()
        except PermissionError:
            print(f"  [警告] 文件被占用，跳过: {old_file.name}")

    client = KingdeeClient()

    # ── 1. 加载映射 ──────────────────────────────────────────────────────────
    org_code_map = _load_generic_mapping("org_code_map")
    org_name_map = _load_generic_mapping("org_name_map")
    material_code_map = _load_generic_mapping("material_code_map")
    material_spec_map = _load_generic_mapping("material_spec_map")
    unit_code_map = _load_generic_mapping("unit_code_map")
    unit_name_map = _load_generic_mapping("unit_name_map")
    lot_name_map = _load_generic_mapping("inv_lot_name_map")
    stock_loc_name_map = _load_generic_mapping("instock_stock_loc_name_map")
    dept_code_map = _load_generic_mapping("dept_code_map")
    dept_name_map = _load_generic_mapping("dept_name_map")
    user_name_map = _load_generic_mapping("user_name_map")

    # 兼容旧格式：dict值展平为字符串
    for k, v in list(lot_name_map.items()):
        if isinstance(v, dict):
            lot_name_map[k] = v.get("cn", str(v))
    for k, v in list(stock_loc_name_map.items()):
        if isinstance(v, dict):
            stock_loc_name_map[k] = v.get("cn", str(v))

    # ── 2. 收集需要映射的ID ─────────────────────────────────────────────────
    def _collect_ids(col_name):
        if col_name not in data_cols:
            return []
        ids = []
        for val in df[col_name].dropna().unique():
            s = str(val).strip()
            if s and s != "0" and s != "0.0" and s != "None":
                ids.append(s)
        return list(set(ids))

    org_ids = _collect_ids("FPrdOrgId") + _collect_ids("FStockOrgId")
    material_ids = _collect_ids("FMaterialId")
    unit_ids = _collect_ids("FUnitID") + _collect_ids("FBaseUnitId")
    stock_ids = _collect_ids("FStockId")
    lot_ids = _collect_ids("FLot")
    dept_ids = _collect_ids("FEntryWorkShopId")
    user_ids = _collect_ids("FPickerId") + _collect_ids("FSTOCKERID")
    aux_prop_ids = _collect_ids("FAuxPropId")

    # ── 3. 获取缺失的映射 ────────────────────────────────────────────────────
    # 组织映射
    missing_org_ids = [oid for oid in org_ids if oid not in org_code_map]
    if missing_org_ids:
        print(f"    需要解析 {len(missing_org_ids)} 个组织ID")
        new_org_map = fetch_generic_mapping(client, "ORG_Organizations", "FOrgId", "FNumber", "org_code_map", missing_org_ids)
        org_code_map.update(new_org_map)
        new_org_names = _fetch_name_map(client, "ORG_Organizations", list(set(new_org_map.values())), "org_name_map")
        org_name_map.update(new_org_names)

    # 物料映射
    missing_material_ids = [mid for mid in material_ids if mid not in material_code_map]
    if missing_material_ids:
        print(f"    需要解析 {len(missing_material_ids)} 个物料ID")
        new_material_map, _ = fetch_material_mapping(client, missing_material_ids)
        material_code_map.update(new_material_map)

    # 单位映射
    missing_unit_ids = [uid for uid in unit_ids if uid not in unit_code_map]
    if missing_unit_ids:
        print(f"    需要解析 {len(missing_unit_ids)} 个单位ID")
        new_unit_map = fetch_generic_mapping(client, "BD_UNIT", "FUnitID", "FNumber", "unit_code_map", missing_unit_ids)
        unit_code_map.update(new_unit_map)
        new_unit_names = _fetch_name_map(client, "BD_UNIT", list(set(new_unit_map.values())), "unit_name_map")
        unit_name_map.update(new_unit_names)

    # 仓库映射
    stock_code_map = _load_generic_mapping("stock_code_map")
    stock_name_map = _load_generic_mapping("stock_name_map")
    missing_stock_ids = [sid for sid in stock_ids if sid not in stock_code_map]
    if missing_stock_ids:
        print(f"    需要解析 {len(missing_stock_ids)} 个仓库ID")
        new_stock_map = fetch_generic_mapping(client, "BD_STOCK", "FStockId", "FNumber", "stock_code_map", missing_stock_ids)
        stock_code_map.update(new_stock_map)
        new_stock_names = _fetch_name_map(client, "BD_STOCK", list(set(new_stock_map.values())), "stock_name_map")
        stock_name_map.update(new_stock_names)

    # 批号映射（从STK_Inventory查询 + view API回退）
    if lot_ids:
        print(f"    需要解析 {len(lot_ids)} 个批号ID")
        missing_lots = [lid for lid in lot_ids if lid not in lot_name_map]
        if missing_lots:
            print(f"      查询 {len(missing_lots)} 个缺失批号（STK_Inventory）")
            batch_size = 200
            for i in range(0, len(missing_lots), batch_size):
                batch = missing_lots[i:i + batch_size]
                filter_parts = [f"FLot = {lid}" for lid in batch]
                filter_str = " or ".join(filter_parts)
                try:
                    rows = client.execute_bill_query({
                        "FormId": "STK_Inventory",
                        "FieldKeys": "FLot,FLot.FNumber",
                        "FilterString": filter_str,
                        "OrderString": "",
                        "TopRowCount": 0,
                        "StartRow": 0,
                        "Limit": len(batch) + 10,
                    })
                    if rows:
                        for row in rows:
                            lot_id = str(row[0]).strip() if row[0] else ""
                            lot_number = str(row[1]).strip() if row[1] else ""
                            if lot_id and lot_number and lot_number not in ("None", "0"):
                                lot_name_map[lot_id] = lot_number
                except Exception:
                    pass
        # 检查仍未解析的批号ID，通过 view API 回退
        still_missing = [lid for lid in lot_ids if lid not in lot_name_map]
        if still_missing:
            print(f"    仍有 {len(still_missing)} 个批号未解析，尝试通过 view API 回退")
            lot_bill_map = {}
            for _, row in df.iterrows():
                lid = str(row.get("FLot", "")).strip()
                if lid in still_missing and lid not in lot_bill_map:
                    bill_no = str(row.get("FBillNo", "")).strip()
                    if bill_no:
                        lot_bill_map[lid] = bill_no
            for lid, bill_no in lot_bill_map.items():
                try:
                    result = client.view(PICKMTRL_FORM_ID, {"Id": 0, "Number": bill_no, "CreateOrgId": 0})
                    inner = result.get("Result", {}).get("Result", {})
                    entries = inner.get("Entity", [])
                    for entry in entries:
                        lot_obj = entry.get("Lot", {})
                        if lot_obj:
                            entry_lot_id = str(lot_obj.get("Id", ""))
                            if entry_lot_id and entry_lot_id not in lot_name_map:
                                lot_number = str(lot_obj.get("Number", "")).strip()
                                if lot_number and lot_number not in ("None", "0"):
                                    lot_name_map[entry_lot_id] = lot_number
                except Exception:
                    pass
        _save_generic_mapping("inv_lot_name_map", lot_name_map)

    # 部门映射（车间）
    if dept_ids:
        missing_dept_ids = [did for did in dept_ids if did not in dept_code_map]
        if missing_dept_ids:
            print(f"    需要解析 {len(missing_dept_ids)} 个部门ID")
            new_dept_map = fetch_generic_mapping(client, "BD_Department", "FDeptId", "FNumber", "dept_code_map", missing_dept_ids)
            dept_code_map.update(new_dept_map)
            new_dept_names = _fetch_name_map(client, "BD_Department", list(set(new_dept_map.values())), "dept_name_map")
            dept_name_map.update(new_dept_names)

    # 用户映射（领料人、仓管员）
    if user_ids:
        missing_user_ids = [uid for uid in user_ids if uid not in user_name_map]
        if missing_user_ids:
            print(f"    需要解析 {len(missing_user_ids)} 个用户ID")
            new_user_map = fetch_generic_mapping(client, "BD_Empinfo", "FEmpID", "FNumber", "user_code_map", missing_user_ids)
            new_user_names = _fetch_name_map(client, "BD_Empinfo", list(set(new_user_map.values())), "user_name_map")
            user_name_map.update(new_user_names)

    # 辅助属性映射（FAuxPropId -> 通过 view API 获取包装规格名称）
    aux_prop_name_map = {}
    if aux_prop_ids:
        print(f"    需要解析 {len(aux_prop_ids)} 个辅助属性ID")
        aux_prop_name_map = _load_generic_mapping("pickmtrl_aux_prop_name_map")
        # 兼容旧格式缓存
        for k, v in list(aux_prop_name_map.items()):
            if isinstance(v, str):
                aux_prop_name_map[k] = {"cn": v, "en": v}
        missing_aux_ids = [aid for aid in aux_prop_ids if aid not in aux_prop_name_map]
        if missing_aux_ids:
            print(f"    需要查询 {len(missing_aux_ids)} 个辅助属性名称")
            # 获取包含这些 FAuxPropId 的单据编号
            aux_bill_map = {}
            for _, row in df.iterrows():
                aid = str(row.get("FAuxPropId", "")).strip()
                if aid in missing_aux_ids and aid not in aux_bill_map:
                    bill_no = str(row.get("FBillNo", "")).strip()
                    if bill_no:
                        aux_bill_map[aid] = bill_no
            # 通过 view API 查询 PRD_PickMtrl 单据获取辅助属性信息
            for aid, bill_no in aux_bill_map.items():
                try:
                    result = client.view(PICKMTRL_FORM_ID, {"Id": 0, "Number": bill_no, "CreateOrgId": 0})
                    inner = result.get("Result", {}).get("Result", {})
                    entries = inner.get("Entity", [])
                    for entry in entries:
                        aux_prop = entry.get("AuxPropId", {}) or entry.get("AuxpropId", {})
                        if aux_prop:
                            entry_aux_id = str(aux_prop.get("Id", ""))
                            if entry_aux_id and entry_aux_id not in aux_prop_name_map:
                                cn_spec = ""
                                en_spec = ""
                                f100001 = aux_prop.get("F100001", {})
                                if f100001:
                                    fdata_values = f100001.get("FDataValue", [])
                                    for dv in fdata_values:
                                        if isinstance(dv, dict):
                                            if dv.get("Key") == 2052:
                                                cn_spec = dv.get("Value", "")
                                            elif dv.get("Key") == 1033:
                                                en_spec = dv.get("Value", "")
                                cn_spec = cn_spec.strip().rstrip('.').strip()
                                en_spec = en_spec.strip().rstrip('.').strip()
                                aux_prop_name_map[entry_aux_id] = {
                                    "cn": cn_spec,
                                    "en": en_spec if en_spec else cn_spec,
                                }
                except Exception:
                    pass
            _save_generic_mapping("pickmtrl_aux_prop_name_map", aux_prop_name_map)

    # 物料名称映射
    material_code_to_name = {}
    if material_ids:
        print(f"    需要解析 {len(material_ids)} 个物料名称")
        _, material_name_map = fetch_material_mapping(client, material_ids)
        for mid, code in material_code_map.items():
            name = material_name_map.get(mid, "")
            if code and name:
                material_code_to_name[code] = name
        # 补充缺失的物料名称
        missing_name_codes = [c for c in set(material_code_map.values()) if c not in material_code_to_name]
        if missing_name_codes:
            print(f"      查询 {len(missing_name_codes)} 个缺失物料名称")
            batch_size = 50
            for i in range(0, len(missing_name_codes), batch_size):
                batch = missing_name_codes[i:i + batch_size]
                filter_parts = [f"FNumber = '{code}'" for code in batch]
                filter_str = " or ".join(filter_parts)
                try:
                    rows = client.execute_bill_query({
                        "FormId": "BD_MATERIAL",
                        "FieldKeys": "FNumber,FName",
                        "FilterString": filter_str,
                        "OrderString": "",
                        "TopRowCount": 0,
                        "StartRow": 0,
                        "Limit": len(batch) + 10,
                    })
                    if rows:
                        for r in rows:
                            fnumber = str(r[0]).strip() if r[0] else ""
                            fname_raw = r[1] if r[1] else ""
                            fname = ""
                            if fname_raw:
                                try:
                                    import json as _json
                                    if isinstance(fname_raw, str) and fname_raw.startswith("["):
                                        name_list = _json.loads(fname_raw)
                                        if name_list and isinstance(name_list, list) and len(name_list) > 0:
                                            fname = name_list[0].get("Value", str(name_list[0])) if isinstance(name_list[0], dict) else str(name_list[0])
                                    else:
                                        fname = str(fname_raw).strip()
                                except Exception:
                                    fname = str(fname_raw).strip()
                            if fnumber and fname:
                                material_code_to_name[fnumber] = fname
                except Exception:
                    pass

    # ── 4. 状态映射 ──────────────────────────────────────────────────────────
    if lang == "CN":
        status_map = {
            "FDocumentStatus": {"A": "已审核", "B": "暂存", "C": "已审核", "D": "重新审核"},
        }
    else:
        status_map = {
            "FDocumentStatus": {"A": "Approved", "B": "Draft", "C": "Approved", "D": "Re-audit"},
        }

    # ── 5. 字段映射 ──────────────────────────────────────────────────────────
    pickmtrl_field_map = {
        "单据编号": "FBillNo", "Doc No.": "FBillNo",
        "领料人": "FPickerId", "Picked by": "FPickerId",
        "日期": "FDate", "Date": "FDate",
        "单据状态": "FDocumentStatus", "Doc Status": "FDocumentStatus",
        "发料组织": "FStockOrgId", "Issue Org.": "FStockOrgId",
        "生产组织": "FPrdOrgId", "Production Org.": "FPrdOrgId",
        "物料编码": "FMaterialId", "Material Code": "FMaterialId",
        "物料名称": "_material_name", "Material Name": "_material_name",
        "车间": "FEntryWorkShopId", "Workshop": "FEntryWorkShopId",
        "规格型号": "_spec", "Specification": "_spec",
        "单位": "FUnitID", "UoM": "FUnitID",
        "成本价": "FPrice", "U/P": "FPrice",
        "总成本": "FAmount", "Amount": "FAmount",
        "申请数量": "FAppQty", "Requested Qty": "FAppQty",
        "实发数量": "FActualQty", "Actual Ship Qty": "FActualQty",
        "仓库": "FStockId", "Warehouse": "FStockId",
        "批号": "FLot", "Lot No.": "FLot",
        "辅助属性": "FAuxPropId", "Secondary Property": "FAuxPropId",
        "生产订单编号": "FMoBillNo", "MO No.": "FMoBillNo",
        "生产订单行号": "FMoEntrySeq", "MO Row No.": "FMoEntrySeq",
        "需求来源": "FReqSrc", "Requirement Source": "FReqSrc",
    }

    # 将模板表头映射到数据列
    field_mapping = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if not tpl_field:
            continue
        if tpl_field in pickmtrl_field_map:
            data_field = pickmtrl_field_map[tpl_field]
            if data_field.startswith("_") or data_field in data_cols:
                field_mapping[tpl_idx] = data_field

    print(f"  [映射] 找到 {len(field_mapping)} 个字段映射")

    id_to_code_map = {
        "FPrdOrgId": org_code_map,
        "FStockOrgId": org_code_map,
        "FMaterialId": material_code_map,
        "FUnitID": unit_code_map,
        "FStockId": stock_code_map,
        "FLot": lot_name_map,
        "FEntryWorkShopId": dept_code_map,
        "FPickerId": user_name_map,
        "FSTOCKERID": user_name_map,
        # FAuxPropId 映射值是dict格式，由special_cols用original_val查找
    }

    zero_to_empty_cols = {"FAuxPropId", "FLot"}

    special_cols = {}
    for tpl_idx, tpl_field in enumerate(h0):
        if tpl_field in ("生产组织", "Production Org.", "发料组织", "Issue Org."):
            special_cols[tpl_idx] = "org_name"
        elif tpl_field in ("物料名称", "Material Name"):
            special_cols[tpl_idx] = "material_name"
        elif tpl_field in ("规格型号", "Specification"):
            special_cols[tpl_idx] = "material_spec"
        elif tpl_field in ("单位", "UoM"):
            special_cols[tpl_idx] = "unit_name"
        elif tpl_field in ("仓库", "Warehouse"):
            special_cols[tpl_idx] = "stock_name"
        elif tpl_field in ("车间", "Workshop"):
            special_cols[tpl_idx] = "dept_name"
        elif tpl_field in ("领料人", "Picked by"):
            special_cols[tpl_idx] = "user_name"
        elif tpl_field in ("辅助属性", "Secondary Property"):
            special_cols[tpl_idx] = "aux_prop_name"

    # ── 6. 生成Excel ─────────────────────────────────────────────────────────
    print("  [生成] 写入Excel...")

    workbook = openpyxl.Workbook()
    ws_out = workbook.active
    ws_out.title = sheet_title

    if lang == "CN" and "序号" not in h0:
        h0 = ["序号"] + h0
        field_mapping = {k + 1: v for k, v in field_mapping.items()}
        special_cols = {k + 1: v for k, v in special_cols.items()}

    ws_out.append(h0)

    sn_counter = 0
    for _, row in df.iterrows():
        output_row = [""] * len(h0)
        sn_counter += 1

        # 预计算物料编码（用于名称和规格查找）
        material_code = ""
        if "FMaterialId" in data_cols:
            mid = str(row["FMaterialId"]).strip() if pd.notna(row["FMaterialId"]) else ""
            material_code = material_code_map.get(mid, mid)

        for tpl_idx, data_col in field_mapping.items():
            if tpl_idx >= len(output_row):
                continue

            # 特殊字段：物料名称
            if data_col == "_material_name":
                val = material_code_to_name.get(material_code, "")
                output_row[tpl_idx] = _clean_string(val)
                continue

            # 特殊字段：规格型号
            if data_col == "_spec":
                val = material_spec_map.get(material_code, "")
                output_row[tpl_idx] = _clean_string(val)
                continue

            if data_col not in data_cols:
                continue

            val = row[data_col]
            if pd.isna(val) or str(val).strip() == "":
                continue

            str_val = str(val).strip()
            original_val = str_val  # 保存原始值，供special_cols查找dict格式映射

            if str_val in ("0", "0.0") and data_col in zero_to_empty_cols:
                continue

            # ID映射
            if data_col in id_to_code_map and id_to_code_map[data_col]:
                str_val = _resolve_id(str_val, id_to_code_map[data_col])

            # 状态映射
            if data_col in status_map:
                str_val = status_map[data_col].get(str_val, str_val)

            # 日期格式化
            if data_col in ("FDate", "FProduceDate", "FExpiryDate"):
                str_val = _format_date(str_val)

            # special_cols处理
            if tpl_idx in special_cols:
                sc = special_cols[tpl_idx]
                if sc == "org_name":
                    str_val = org_name_map.get(str_val, str_val)
                elif sc == "unit_name":
                    if lang == "CN":
                        str_val = unit_name_map.get(str_val, str_val)
                elif sc == "stock_name":
                    str_val = stock_name_map.get(str_val, str_val)
                elif sc == "dept_name":
                    str_val = dept_name_map.get(str_val, str_val)
                elif sc == "user_name":
                    str_val = str_val  # 已通过id_to_code_map映射
                elif sc == "material_name":
                    str_val = material_code_to_name.get(str_val, str_val)
                elif sc == "material_spec":
                    spec = material_spec_map.get(str_val, "")
                    str_val = spec if spec else ""
                elif sc == "aux_prop_name":
                    # 使用原始ID查找，映射格式: {id: {"cn": "...", "en": "..."}}
                    info = aux_prop_name_map.get(original_val, {})
                    if isinstance(info, dict):
                        str_val = info.get("cn" if lang == "CN" else "en", info.get("cn", ""))
                    elif isinstance(info, str):
                        str_val = info
                    else:
                        str_val = ""

            output_row[tpl_idx] = _clean_string(str_val)

        if "序号" in h0:
            output_row[h0.index("序号")] = str(sn_counter)

        ws_out.append(output_row)

    date_cols = {"FDate", "FProduceDate", "FExpiryDate"}
    _apply_date_formats(ws_out, h0, date_cols, field_mapping)
    _save_workbook(workbook, output_path)
    print(f"  [完成] {output_path.name} ({len(df)} 行)")

    # ── 7. 保存映射 ──────────────────────────────────────────────────────────
    _save_generic_mapping("inv_lot_name_map", lot_name_map)

    print(f"[完成] 生成生产领料单 Excel ({lang_label})")


# ══════════════════════════════════════════════════════════════════════════════
# 主函数
# ══════════════════════════════════════════════════════════════════════════════

def _sync_to_network(network_root: str) -> None:
    """将本地 output 目录中的 xlsx 文件同步到网络共享目录（映射驱动器方式），每个文件对应独立子目录。"""
    import subprocess, tempfile

    if not OUTPUT_DIR.exists():
        print(f"[同步] 本地 output 目录不存在，跳过同步")
        return
    if not network_root:
        print(f"[同步] 未指定网络输出路径，跳过同步")
        return

    # 文件名 → 网络子目录映射
    FILE_NET_MAP = {
        "物料清单列表.xlsx": "K3_Data_BOM",
        "K3_BOM.xlsx": "K3_Data_BOM",
        "即时库存.xlsx": "K3_Data_Inventory",
        "K3_Inventory.xlsx": "K3_Data_Inventory",
        "生产领料单列表.xlsx": "K3_Data_PickingList",
        "K3_18015_PickingList.xlsx": "K3_Data_PickingList",
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
        "物料列表.xlsx": "K3_MaterialMaster",
        "K3_MaterialMaster.xlsx": "K3_MaterialMaster",
    }

    xlsx_files = list(OUTPUT_DIR.glob("*.xlsx"))
    print(f"[同步] 本地 output 共 {len(xlsx_files)} 个 xlsx 文件，目标根: {network_root}")

    if not xlsx_files:
        print(f"[同步] 无 xlsx 文件需要复制")
        return

    # 网络认证参数
    parts = Path(network_root).parts
    share_root = parts[0].rstrip("\\") if parts else network_root
    net_user = r"STERN\shenjj"
    net_pass = r"zH820303."
    root_sub_path = str(Path(network_root)).replace(share_root, "").strip("\\")

    # 生成临时批处理脚本：映射驱动器 → 创建子目录 → 复制 → 断开驱动器
    bat_lines = [
        f'net use * /delete /yes 2>nul',
        f'timeout /t 1 /nobreak >nul',
        f'net use Z: {share_root} {net_pass} /user:{net_user} /persistent:no',
    ]
    for f in xlsx_files:
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
        if result.returncode == 0:
            print(f"[同步] 完成，成功复制 {len(xlsx_files)} 个文件到 {network_root}")
        else:
            print(f"[同步] 批处理返回非零退出码: {result.returncode}")
    except subprocess.TimeoutExpired:
        print(f"[同步] 批处理超时（180秒）")
    except Exception as e:
        logger.error("批处理执行失败: %s", e, exc_info=True)
        print(f"[同步] 批处理执行失败: {e}")
    finally:
        try:
            bat_file.unlink()
        except OSError:
            pass


def main() -> None:
    parser = argparse.ArgumentParser(description="ENG_BOM & PLN_FORECAST & PLN_PLANORDER & PRD_MO & PUR_PurchaseOrder 数据同步与导入 Excel 生成")
    parser.add_argument("--skip-fetch", action="store_true", help="跳过金蝶拉取")
    parser.add_argument("--form", type=str, default="all",
                        choices=["all", "BD_MATERIAL", "ENG_BOM", "PLN_FORECAST", "PLN_PLANORDER", "PRD_MO", "PUR_PurchaseOrder", "PRD_INSTOCK", "STK_Inventory", "PRD_PickMtrl"],
                        help="指定要处理的表单，支持: all(全部), BD_MATERIAL(物料基础资料), ENG_BOM(物料清单), PLN_FORECAST(预测单), PLN_PLANORDER(计划订单), PRD_MO(生产订单), PUR_PurchaseOrder(采购订单), PRD_INSTOCK(生产入库单), STK_Inventory(即时库存), PRD_PickMtrl(生产领料单)")
    parser.add_argument("--lang", type=str, default="EN",
                        choices=["ALL", "CN", "EN"],
                        help="输出语言版本: ALL(中英文都输出), CN(仅中文), EN(仅英文，默认)")
    parser.add_argument("--netaddr", type=str, nargs='?',
                        default=r"\\192.168.1.181\12_SupplyChain\07_MPS\MPS_Upgrade",
                        const="",
                        help="网络输出根路径（SMB共享），留空则不同步到网络")
    args = parser.parse_args()

    langs = ["CN", "EN"] if args.lang == "ALL" else [args.lang]

    print("══════════════════════════════════════════════════════════════════════")
    print(f"ENG_BOM & PLN_FORECAST & PLN_PLANORDER & PRD_MO & PUR_PurchaseOrder 数据同步与导入 Excel 生成工具")
    print(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"表单: {args.form}  语言: {'/'.join(langs)}")
    print(f"日志: {_log_file}")
    print("══════════════════════════════════════════════════════════════════════")

    OUTPUT_DIR.mkdir(exist_ok=True)

    # 1. 从金蝶拉取数据
    fetch_all_from_kingdee(args.skip_fetch, args.form)

    # 2. 生成物料基础资料 Excel
    if args.form == "all" or args.form == "BD_MATERIAL":
        for lang in langs:
            generate_material_excel(lang)

    # 3. 生成物料清单导入 Excel
    if args.form == "all" or args.form == "ENG_BOM":
        for lang in langs:
            generate_import_excel(lang)

    # 4. 生成预测单导入 Excel
    if args.form == "all" or args.form == "PLN_FORECAST":
        for lang in langs:
            generate_forecast_excel(lang)

    # 5. 生成计划订单 Excel
    if args.form == "all" or args.form == "PLN_PLANORDER":
        for lang in langs:
            generate_planorder_excel(lang)

    # 6. 生成生产订单 Excel
    if args.form == "all" or args.form == "PRD_MO":
        for lang in langs:
            generate_mo_excel(lang)

    # 7. 生成采购订单 Excel
    if args.form == "all" or args.form == "PUR_PurchaseOrder":
        for lang in langs:
            generate_purchaseorder_excel(lang)

    # 8. 生成生产入库单 Excel
    if args.form == "all" or args.form == "PRD_INSTOCK":
        for lang in langs:
            generate_instock_excel(lang)

    # 9. 生成即时库存 Excel
    if args.form == "all" or args.form == "STK_Inventory":
        for lang in langs:
            generate_inv_excel(lang)

    # 10. 生成生产领料单 Excel（默认不生成，需显式指定 --form PRD_PickMtrl）
    if args.form == "PRD_PickMtrl":
        for lang in langs:
            generate_pickmtrl_excel(lang)

    # 11. 同步到网络地址
    if args.netaddr:
        _sync_to_network(args.netaddr)

    print("══════════════════════════════════════════════════════════════════════")
    print("完成!")
    print("══════════════════════════════════════════════════════════════════════")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.exception("程序异常退出: %s", e)
        print(f"\n程序异常退出，详见日志: {_log_file}", file=sys.stderr)
        sys.exit(1)
