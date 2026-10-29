# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Read-only integration with 金蝶云星空 (Kingdee K3 Cloud) ERP, using the official Python SDK. Pulls business data (materials, BOMs, inventory, production orders, purchase orders, etc.) via `ExecuteBillQuery` and `View` APIs, then generates properly formatted Excel import files matching Kingdee templates. Operated for 斯顿食品原料（苏州）有限公司.

## Architecture

```
src/                         # Shared core library (used by all workspace modules)
├── config.py                # Pydantic Settings, reads .env
├── kingdee_client.py        # KingdeeClient: wraps official SDK (lazy-loads from .vendor/ wheel)
├── export_excel.py          # DataFrame→xlsx export helper + demo generator
├── fetch_execute_bill_query.py  # ExecuteBillQuery helper scripts
├── fetch_form_field_keys.py     # Form field metadata fetcher scripts
└── fetch_view.py                # View API helper scripts

workspace/
├── BOM_API/                 # Original module: BOM & supply-chain data sync (9 form types)
│   ├── sync_and_generate.py # ~6500-line engine: fetch → resolve IDs → generate Excel → network sync
│   ├── server.py            # Flask web UI backend (port 5000), orchestrates sync_and_generate.py via subprocess
│   ├── index.html           # Single-page web UI
│   ├── .progress.json       # Runtime progress tracking (read/written by server.py)
│   ├── map/                 # ID→code/name lookup caches (~53 JSON files, persisted across runs)
│   ├── template/            # Kingdee import templates (CN/EN xlsx per form type)
│   ├── data/                # Raw CSV cache from Kingdee fetches
│   └── output/              # Generated Excel import files
├── COA-COC2/                # Quality inspection master data sync (5 form types)
│   └── sync_and_generate.py # Fetches QC data, compares with test plan Excel, generates missing-item imports
├── COA-COC1/                # Test Plan generation (placeholder, not yet implemented)
├── Master-Data/             # Master data processing (placeholder)
├── Inventory/               # Inventory sync (standalone deprecated module, not in EXP)
└── EXP/                     # New BD_MATERIAL-focused module (current development target)
    ├── sync_and_generate.py # Material master sync with updated field mappings
    ├── server.py            # Flask web UI backend
    ├── index.html           # Single-page web UI
    ├── map/                 # ID→code/name lookup caches
    ├── template/            # Kingdee import templates (CN/EN)
    ├── data/                # Raw CSV cache
    └── output/              # Generated Excel import files
```

## EXP Module — Key Differences from BOM_API

### Field Mapping Changes
- **Color**: Deprecated `F_PALW_Assistant_83g` (颜色（废弃）) → replaced by `F_PALW_Base_apv1` (颜色)
  - F_PALW_Base_apv1 has a `Description` field containing color codes: dark/light/medium/MS/NA
  - `color code` column → Description field
  - `color` column → DataValue (full display name like "taupe 灰褐色")
- **Odor**: Deprecated `F_PALW_Assistant_re5` (气味（废弃）) → replaced by `F_PALW_Base_tzk` (气味)
  - `odor code` column → DataValue (F/MS/N/O)
  - `odor` column → intentionally empty
- **Non-GMO clarified**: `F_PALW_Text_uky` (wrong — had energy values) → `F_PALW_CheckBox_uky` with special display "Non-GMO with statement"
- **Density**: `F_WTRM_Text31` (empty) → `F_PALW_Text_w5c1`
- **Allergen Code**: Full names abbreviated via `_ALLERGEN_COMPONENT_MAP` (e.g., "gluten"→"GLU", "大豆、麸质"→"SOY, GLU")
- **GB 2760 Code**: `F_PALW_Text_tzk2`
- **SC**: `F_PALW_Text_qtr12`

### Resolution Functions
- `_resolve_assistant_guids(field_name, guid_ids, cache_filename)` → returns `(name_map, code_map)` for F_PALW_Assistant* fields
- `_resolve_base_field(field_name, guid_ids, cache_filename)` → returns `(name_map, code_map, desc_map)` for F_PALW_Base* fields (extracts Description)
- `_find_field_value` in both functions caps at 50 new GUIDs per run per field

### Network Sync
- All files sync to `\\192.168.1.181\99_Share\04_IT\KingdeeExport` using single Z: drive mapping
- `~$` prefix files (Excel temp locks) are filtered out
- WSL dev environment has UNC path limitation; sync works natively on Windows
- `cmd.exe` is used for batch execution (works on both WSL and Windows)
- Original BOM_API uses `"cmd"` — keep that in BOM_API

## Common Commands

```bash
# Verify Kingdee connectivity and SDK config
python main.py check

# Query field metadata for a Kingdee business object
python main.py fields BD_MATERIAL

# CLI diagnostics & metadata queries
python main.py export-demo           # Generate demo Excel (no Kingdee connection)

# Run EXP sync (CLI) — specific form, single language
python workspace/EXP/sync_and_generate.py --form BD_MATERIAL --lang EN

# Run EXP sync — using cached data
python workspace/EXP/sync_and_generate.py --skip-fetch

# Run EXP sync — disable network copy
python workspace/EXP/sync_and_generate.py --netaddr ""

# Start EXP web UI
python workspace/EXP/server.py
# Then open http://localhost:5000

# BOM_API (original module)
python workspace/BOM_API/sync_and_generate.py --form ENG_BOM --lang CN

# BOM_API web UI
python workspace/BOM_API/server.py
# Then open http://localhost:5000

# Quality inspection sync
python workspace/COA-COC2/sync_and_generate.py --skip-fetch
```

## Configuration

- **`conf.ini`** (gitignored): Kingdee API credentials — `AcctID`, `UserName`, `AppID`, `AppSec`, `ServerUrl` under `[config]` section
- **`.env`** (gitignored): Non-sensitive runtime settings — `KINGDEE_SDK_WHEEL_PATH`, `KINGDEE_CONFIG_PATH`, `KINGDEE_CONFIG_NODE`, optional query/view parameters
- Templates available: `conf.ini.example`, `.env.example`
- SDK wheel vendored at: `.vendor/kingdee.cdp.webapi.sdk-8.2.0-py3-none-any.whl`

## EXP Network Sync Paths

All files → `\\192.168.1.181\99_Share\04_IT\KingdeeExport`

| Output File | Subdirectory |
|---|---|
| K3_BOM.xlsx / 物料清单列表.xlsx | K3_Data_BOM |
| K3_MaterialMaster.xlsx / 物料列表.xlsx | K3_MaterialMaster |
| K3_Forecast.xlsx | K3_Data_SalesForecast |
| K3_PlannedOrders.xlsx | K3_Data_ProductionRequests |
| K3_PurchaseOrders.xlsx | K3_Data_PurchaseOrders |
| K3_Production_Receipts.xlsx | K3_Data_Production_Receipts_MTD |
| K3_ManufacturingOrders.xlsx | K3_Data_ProductionOrders |
| K3_Pickinglist.xlsx | K3_Data_PickingList |
| Inventory files | Removed (not synced) |

## Notes

- The `workspace/EXP/AGENTS.md` file is a stale summary of this doc — consider removing or updating it.
- There are no automated tests. Verification is manual: run the scripts and inspect the output Excel files.
- The project runs directly from source — no build step, no `pyproject.toml`, no CI.
- Dependencies are listed in `requirements.txt`: `python-dotenv`, `pydantic-settings`, `pandas`, `openpyxl`, `flask`, `urllib3`, `requests`.
- Virtual environment at `.venv/` (gitignored).
- `sync_and_generate.py` in both BOM_API and EXP is very large (~6500 lines). Each `generate_*_excel()` function is self-contained; there is little shared abstraction between them.
- The `map/` directory caches ID lookups. If Kingdee data changes significantly, delete the relevant map files to force a fresh fetch.
- SMB network sync uses Windows `net use` for domain authentication. WSL dev environment has UNC path limitation — sync works on native Windows.
- Logs are written to `workspace/<module>/log/run_YYYYMMDD_HHMMSS.log` on each run.
- Color/odor resolution caps at 50 new GUIDs per run per field. Multiple runs needed for full resolution on first sync.
