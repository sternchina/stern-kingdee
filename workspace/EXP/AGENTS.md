# EXP Module — 金蝶数据导出通用模块 (stern-kingdee)

EXP 是 [斯顿食品原料（苏州）有限公司] 金蝶云星空集成项目的核心模块，专注从金蝶导出 **各种业务数据**（物料、BOM、预测单、计划/生产/采购订单、入库单、领料单共 8 种表单）。它通过 `ExecuteBillQuery` / `View` API 拉取数据、解析辅助资料/基础资料 GUID、生成标准 Excel 导入文件，并 SMB 同步到 Windows 共享目录。

---

## Project — 在项目中的位置

```
stern-kingdee/
├── src/                          # 共享核心库（所有 workpsace 模块共用）
│   ├── config.py                 # Pydantic Settings, 读 .env
│   ├── kingdee_client.py         # KingdeeClient: 封装官方 SDK（懒加载 .vendor/ wheel）
│   ├── export_excel.py           # DataFrame→xlsx 导出工具 + demo 生成器
│   ├── fetch_execute_bill_query.py
│   ├── fetch_form_field_keys.py
│   └── fetch_view.py
├── workspace/
│   ├── EXP/                      # ← 本模块：金蝶数据导出（8 种表单）
│   │   ├── sync_and_generate.py  # 主引擎（~6500 行）：fetch → resolve → generate → sync
│   │   ├── server.py             # Flask Web 后端，支持全部 8 种表单（端口 5000）
│   │   ├── index.html            # 单页 Web UI
│   │   ├── .progress.json        # 运行时进度（server.py 读写）
│   │   ├── map/                  # ID→code/name 缓存（JSON）
│   │   ├── template/             # 金蝶导入模板（CN/EN xlsx）
│   │   ├── data/                 # 原始 CSV 缓存
│   │   └── output/               # 生成的 Excel 文件
│   ├── BOM_API/                  # 原模块：BOM 及供应链数据同步（功能完整，持续维护）
│   ├── COA-COC2/                 # 质检主数据同步（5 种表单）
│   ├── COA-COC1/                 # Test Plan 生成（预留，未实现）
│   ├── Master-Data/              # 主数据处理（预留）
│   └── Inventory/                # 库存同步（独立废弃模块，未在 EXP 中）
├── .vendor/                      # kingdee.cdp.webapi.sdk-8.2.0-py3-none-any.whl
├── .env / .env.example           # 运行时参数
├── conf.ini / conf.ini.example   # 金蝶 SDK 凭证
├── main.py                       # 共享 CLI 入口（check / query / view / fields / export-demo）
├── requirements.txt              # 依赖：python-dotenv, pydantic-settings, pandas, openpyxl, flask, urllib3, requests
├── reasonix.toml                 # Reasonix 权限配置
└── CLAUDE.md                     # 项目级权威指南（比本文档新）
```

---

## Commands — EXP 模块

```bash
# === 共享 CLI (main.py) ===
python main.py check              # 验证金蝶 SDK 和 conf.ini 配置
python main.py fields BD_MATERIAL # 查询字段元数据
python main.py export-demo        # 生成示例 Excel（不连金蝶）

# === EXP 同步 ===
python workspace/EXP/sync_and_generate.py --form all --lang EN       # 导出全部表单（含领料单）
python workspace/EXP/sync_and_generate.py --form BD_MATERIAL --lang EN  # 指定表单+语言
python workspace/EXP/sync_and_generate.py --skip-fetch                   # 重用本地 CSV 缓存
python workspace/EXP/sync_and_generate.py --netaddr ""                   # 跳过 SMB 网络复制

# === EXP Web UI ===
python workspace/EXP/server.py     # 启动 Flask（http://localhost:5000）

# === 其他模块 ===
python workspace/BOM_API/sync_and_generate.py --form ENG_BOM --lang CN
python workspace/BOM_API/server.py
python workspace/COA-COC2/sync_and_generate.py --skip-fetch
```

`sync_and_generate.py` 支持的全部表单：`BD_MATERIAL`, `ENG_BOM`, `PLN_FORECAST`, `PLN_PLANORDER`, `PRD_MO`, `PUR_PurchaseOrder`, `PRD_INSTOCK`, `PRD_PickMtrl`。

---

## Architecture — EXP 模块架构

### 数据流
```
金蝶云星空 API
    ↓ ExecuteBillQuery / View
CSV 缓存 (data/*.csv)
    ↓ _resolve_assistant_guids / _resolve_base_field (GUID → code/name/desc)
地图缓存 (map/*.json)
    ↓ generate_*_excel() 填入模板
Excel 文件 (output/*.xlsx)
    ↓ SMB 同步 (net use Z:)
\\192.168.1.181\99_Share\04_IT\KingdeeExport\K3_*
```

### 关键函数
- `_resolve_assistant_guids(field, guids, cache_file)` → `(name_map, code_map)` — 解析 `F_PALW_Assistant*` 字段
- `_resolve_base_field(field, guids, cache_file)` → `(name_map, code_map, desc_map)` — 解析 `F_PALW_Base*` 字段（提取 Description）
- `_find_field_value` 每次运行最多处理 50 个新 GUID，首次同步需多次运行

### BD_MATERIAL 特有字段映射（相对于旧版 BOM_API 的变化）
- **颜色**: `F_PALW_Assistant_83g`（废弃）→ `F_PALW_Base_apv1`；`color code` → Description，`color` → DataValue
- **气味**: `F_PALW_Assistant_re5`（废弃）→ `F_PALW_Base_tzk`；`odor code` → DataValue，`odor` → 留空
- **Non-GMO**: `F_PALW_Text_uky`（错误）→ `F_PALW_CheckBox_uky`，显示 "Non-GMO with statement"
- **密度**: `F_WTRM_Text31`（空）→ `F_PALW_Text_w5c1`
- **过敏原编码**: 全名通过 `_ALLERGEN_COMPONENT_MAP` 缩写（如 "gluten"→"GLU"）
- **GB 2760**: `F_PALW_Text_tzk2`
- **SC**: `F_PALW_Text_qtr12`

### 网络同步路径
所有文件 → `\\192.168.1.181\99_Share\04_IT\KingdeeExport`

| 输出文件 | 子目录 |
|---|---|
| K3_BOM.xlsx / 物料清单列表.xlsx | `K3_Data_BOM` |
| K3_MaterialMaster.xlsx / 物料列表.xlsx | `K3_MaterialMaster` |
| K3_Forecast.xlsx | `K3_Data_SalesForecast` |
| K3_PlannedOrders.xlsx | `K3_Data_ProductionRequests` |
| K3_PurchaseOrders.xlsx | `K3_Data_PurchaseOrders` |
| K3_Production_Receipts.xlsx | `K3_Data_Production_Receipts_MTD` |
| K3_ManufacturingOrders.xlsx | `K3_Data_ProductionOrders` |
| K3_Pickinglist.xlsx / 生产领料单列表.xlsx | `K3_Data_PickingList` |

SMB 同步使用 `net use Z:` + `cmd.exe`。WSL 环境下 UNC 路径受限，需在原生 Windows 运行同步。

---

## Conventions

- **编码**: Python 3.12+, `from __future__ import annotations`, 4 空格缩进, `snake_case`/`PascalCase`/`UPPER_SNAKE_CASE`
- **类型注解**: PEP 604（`X | Y`），标准库 `annotations`
- **导入顺序**: 标准库 → 第三方 → 本地，每组空行分隔
- **无格式化工具**: 手动 PEP 8，行 ≤120
- **日志**: `workspace/<module>/log/run_YYYYMMDD_HHMMSS.log`
- **缓存**: `map/` 目录 JSON 文件持久化 ID 查询结果；金蝶数据大变更时手动删除刷新
- **同步**: `~$` 前缀（Excel 临时锁）自动过滤
- **无测试框架**: 手动验证，运行脚本检查输出 Excel
- **提交**: Conventional commits (`feat|fix|docs|refactor|chore: 描述`)
- **不要提交** `conf.ini` 或 `.env`（含凭证，已在 `.gitignore`）

---

## Notes

- 本文档是 EXP 模块专用指南；项目级权威文档见根目录 `CLAUDE.md`。
- `sync_and_generate.py` 在 EXP 和 BOM_API 中均约 6500 行，每个 `generate_*_excel()` 自包含，几乎没有共享抽象。
- Kingdee SDK wheel 提交在 `.vendor/kingdee.cdp.webapi.sdk-8.2.0-py3-none-any.whl`，勿直接修改。
- SMB 凭证硬编码在 `server.py` 中，对外分享仓库前需审查。
