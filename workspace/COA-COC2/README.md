# COA-COC 数据同步与导入工具
# COA-COC Data Sync and Import Tool

该工具从金蝶云星空拉取质检主数据，与 Test Plan 对比，并生成缺失数据的导入 Excel。
This tool pulls quality master data from Kingdee K3 Cloud, compares it with the Test Plan, and generates import Excel files for missing data.

## 当前实现范围
## Current Scope

- 拉取 `BD_UNIT`、`QM_InspectMethod`、`QM_InspectItem`、`QM_QCScheme`、`QM_QualityStd` 全量数据到 `data/*.csv`。
- Pull full data for `BD_UNIT`, `QM_InspectMethod`, `QM_InspectItem`, `QM_QCScheme`, and `QM_QualityStd` into `data/*.csv`.
- 解析 `input/` 中第一个 `xlsx`，自动识别包含必要列的 sheet。
- Parse the first `xlsx` file in `input/` and automatically detect the sheet that contains all required columns.
- 生成检验方法、质量标准、检验项目、质检方案 4 类导入文件。
- Generate four kinds of import files: inspection methods, quality standards, inspection items, and quality schemes.
- 如果发现缺失的计量单位，当前实现还会额外生成 `计量单位_{timestamp}_导入.xlsx`。
- If missing units are detected, the current implementation also generates `计量单位_{timestamp}_导入.xlsx`.
- 生成导入 Excel 前会删除 `output/` 下全部 `.xlsx`，再写入本次带时间戳的导入文件。
- Before generating import workbooks, the script deletes all `.xlsx` under `output/`, then writes the new timestamped files.

## 目录结构
## Directory Layout

- `COA-COC2/`：当前工作目录。
- `COA-COC2/`: The current workspace directory.
- `sync_and_generate.py`：主脚本。
- `sync_and_generate.py`: Main script.
- `input/`：放置待处理的 Test Plan Excel，脚本每次取第一个 `xlsx`。
- `input/`: Stores the Test Plan Excel file to process, and the script always picks the first `xlsx`.
- `template/`：金蝶导入模板目录，只读使用。
- `template/`: Directory for Kingdee import templates, used as read-only input.
- `data/`：运行后生成的金蝶全量 CSV 缓存。
- `data/`: Generated cache of full CSV exports fetched from Kingdee.
- `output/`：运行后生成的导入 Excel 文件目录。
- `output/`: Directory for generated import Excel files.

## 前置条件
## Prerequisites

1. 项目根目录的 `conf.ini` 已配置金蝶连接信息，例如 `AcctID`、`UserName`、`AppID`、`AppSec`、`ServerUrl`。
1. The project root `conf.ini` must already contain the Kingdee connection settings such as `AcctID`, `UserName`, `AppID`, `AppSec`, and `ServerUrl`.
2. 仓库当前已内置固定版本的金蝶官方 Python SDK wheel，默认无需再手动放置；如果需要切换版本，可改用其他本地 wheel 或全局安装包。
2. The repository currently bundles a pinned version of the official Kingdee Python SDK wheel, so you usually do not need to place it manually; if you need a different version, you can switch to another local wheel or an installed package.
3. 本地依赖已安装，至少包含 `openpyxl`、`pandas`、`pydantic-settings`。
3. Local dependencies must be installed, including at least `openpyxl`, `pandas`, and `pydantic-settings`.

## 使用方法
## Usage

在项目根目录运行以下命令。
Run the following commands from the project root.

```bash
python workspace/COA-COC2/sync_and_generate.py
python workspace/COA-COC2/sync_and_generate.py --skip-fetch
```

第一条命令会从金蝶拉最新数据并生成导入文件。
The first command fetches the latest data from Kingdee and generates import files.

第二条命令会跳过在线拉取，直接使用本地 `data/*.csv` 重新生成导入文件。
The second command skips remote fetching and regenerates import files directly from local `data/*.csv`.

## 处理流程
## Processing Flow

1. 从金蝶拉取 5 个业务对象的全量数据，并保存到 `data/*.csv`。
1. Fetch full data for five business objects from Kingdee and save them to `data/*.csv`.
2. 从 `input/` 读取 Test Plan，自动识别包含必要列的 sheet。
2. Read the Test Plan from `input/` and automatically detect the sheet with all required columns.
3. 将 `method`、`unit`、`target`、`test_item`、`material_code` 与金蝶主数据做匹配。
3. Match `method`, `unit`, `target`, `test_item`, and `material_code` against the Kingdee master data.
4. 找出金蝶中不存在的主数据和方案明细。
4. Identify master data and scheme rows that do not yet exist in Kingdee.
5. 按模板格式把差异写入 `output/` 中的导入 Excel。
5. Write the detected differences into import Excel files under `output/` using the template format.

## 输入要求
## Input Requirements

- `input/` 下至少有 1 个 `xlsx` 文件，如果有多个则按文件名排序后取第一个。
- `input/` must contain at least one `xlsx` file, and if multiple files exist the script selects the first one after filename sorting.
- Test Plan 至少要包含 `material_code`、`material_name`、`test_item`、`target`、`unit`、`method` 这些列。
- The Test Plan must include at least `material_code`, `material_name`, `test_item`, `target`, `unit`, and `method`.
- 脚本会遍历所有 sheet，优先选择 sheet 名包含 `test plan` 且必要列齐全的 sheet。
- The script scans all sheets and prefers the sheet whose name contains `test plan` and also includes all required columns.
- 所有文本字段在比对前会做首尾空白清理，空值、`nan`、`None` 会按空串处理。
- All text fields are trimmed before comparison, and empty values, `nan`, and `None` are normalized to empty strings.

## 匹配规则
## Matching Rules

- `method`、`unit`、`target`、`test_item` 当前都按名称匹配。
- `method`, `unit`, `target`, and `test_item` are currently matched by name.
- 名称列优先使用 `FName#2052`、`FName#1033`、`FName#3076`、`FName`，实际以当前 CSV 中存在的第一列为准。
- Name matching prefers `FName#2052`, `FName#1033`, `FName#3076`, and `FName`, using the first available column in the current CSV.
- `material_code` 按 `QM_QCScheme.FNumber` 匹配，相同物料会聚合成一个质检方案，多行明细写入同一个方案。
- `material_code` is matched against `QM_QCScheme.FNumber`, and rows for the same material are aggregated into one quality scheme.
- 生成导入文件时，跨对象引用会优先回填已有主数据的 `FNumber`，如果引用对象也是本次新增，则使用本次新生成的编码。
- When generating import files, cross-object references first reuse existing `FNumber` values, and if a referenced object is also new in this run, the newly generated code is used.

## 导入顺序
## Import Order

生成的文件存在依赖关系，必须按以下顺序导入金蝶。
The generated files have dependencies and must be imported into Kingdee in the following order.

1. `计量单位_{timestamp}_导入.xlsx`：可选，仅在检测到缺失单位时生成。
1. `计量单位_{timestamp}_导入.xlsx`: Optional and generated only when missing units are detected.
2. `检验方法_{timestamp}_导入.xlsx`：无依赖，优先导入。
2. `检验方法_{timestamp}_导入.xlsx`: Has no dependency and should be imported early.
3. `质量标准_{timestamp}_导入.xlsx`：无依赖，优先导入。
3. `质量标准_{timestamp}_导入.xlsx`: Has no dependency and should be imported early.
4. `检验项目_{timestamp}_导入.xlsx`：引用计量单位、检验方法、质量标准编码。
4. `检验项目_{timestamp}_导入.xlsx`: References unit, inspection method, and quality standard codes.
5. `质检方案_{timestamp}_导入.xlsx`：引用检验项目、计量单位、检验方法、质量标准编码。
5. `质检方案_{timestamp}_导入.xlsx`: References inspection item, unit, inspection method, and quality standard codes.

如果某类型在金蝶中全部已存在，对应文件不会生成。
If a given type already fully exists in Kingdee, the corresponding file is not generated.

脚本会主动为新增记录分配 `FNumber`，并把这些编码回填到下游依赖文件中。
The script assigns `FNumber` values to new records and backfills those codes into downstream dependent files.

## 字段默认值
## Default Field Values

- 创建组织 / 使用组织：`100`（斯顿食品原料（苏州）有限公司）。
- Create organization / use organization: `100` (Stern Food Ingredients Suzhou Co., Ltd.).
- 分析方法：`其他分析`。
- Analysis method: `其他分析`.
- 缺陷等级：`重缺陷`。
- Defect level: `重缺陷`.
- 重点检查：`True`。
- Key inspection: `True`.
- 状态（检验项目行）：`合格`。
- Status for inspection item rows: `合格`.
- 比较符：`=`。
- Compare symbol: `=`.
- 适用范围（质检方案）：`通用`。
- Apply scope for quality schemes: `通用`.
- 生效日期：脚本运行当天。
- Effective date: The day the script runs.
- 失效日期：`9999-12-31`。
- Expire date: `9999-12-31`.

## 已知限制
## Known Limitations

- `QM_InspectItem` 当前按 `test_item` 名称判重，不区分同名但 `unit`、`target`、`method` 不同的情况。
- `QM_InspectItem` currently de-duplicates by `test_item` name only and does not distinguish rows that share the same name but have different `unit`, `target`, or `method`.
- 如果 Test Plan 中存在同名多配置记录，生成结果可能被合并。
- If the Test Plan contains multiple configurations with the same item name, the generated result may merge them.
- 名称匹配使用单值映射，如果金蝶主数据里存在同名多编码，脚本只能命中其中一条。
- Name matching uses a single-value map, so if Kingdee master data contains multiple codes for the same name, the script can match only one of them.
- 脚本当前拉取的是全量数据，没有按组织过滤，但新生成记录固定写入组织 `100`。
- The script currently fetches full data without organization filtering, while all new records are written to organization `100`.
- `template/` 下如果同一类型有多个模板文件，脚本按文件名排序后取第一个。
- If `template/` contains multiple template files for the same type, the script picks the first one after filename sorting.

## 注意事项
## Notes

- `data/` 和 `output/` 目录内容均为自动生成，可随时重新运行覆盖。
- The contents of `data/` and `output/` are generated automatically and can be overwritten by rerunning the script.
- `input/` 目录放 Test Plan Excel，脚本自动取第一个 `xlsx`，更换 Test Plan 时直接替换该目录下文件即可。
- Put the Test Plan Excel file into `input/`, and the script will pick the first `xlsx`; to replace the Test Plan, just replace the file in that directory.
- 当前脚本兼容旧目录名 `import/`、`export/`，如果新旧目录同时存在，会优先使用 `input/`、`output/`。
- The current script is compatible with the legacy directory names `import/` and `export/`, but it prefers `input/` and `output/` when both sets exist.
- `template/` 目录存放金蝶导入模板，不要修改。
- The `template/` directory stores the Kingdee import templates and should not be modified.
- 本地运行若看到 `openpyxl` 关于 `Workbook contains no default style` 的警告，通常不影响导出结果，但仍建议抽样打开生成文件确认格式。
- If you see the `openpyxl` warning `Workbook contains no default style` during local runs, it usually does not affect the exported result, but you should still spot-check the generated files.
- 运行前建议先执行 `python main.py check` 确认金蝶连接正常。
- Before running the script, it is recommended to execute `python main.py check` to confirm the Kingdee connection is healthy.
