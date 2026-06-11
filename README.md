# stern-kingdee

金蝶云星空（K3 Cloud）只读集成，通过官方 Python SDK 调用。

This repository provides a read-only integration with Kingdee K3 Cloud using the official Python SDK.

## 项目概述

本仓库提供金蝶云星空的 Python 集成能力，支持通过官方 SDK 执行只读查询操作。主要功能包括：

- `ExecuteBillQuery`：执行单据查询
- `View`：查看业务对象详情
- `QueryBusinessInfo`：查询字段元数据

## 目录结构

```text
stern-kingdee/
├── .vendor/                    # 内置金蝶 SDK wheel
├── docs/                       # 文档目录
│   ├── kingdee-python-integration.md  # 详细集成说明
│   └── project-tutorial.md            # 项目入门教程
├── src/                        # 核心源码
│   ├── __init__.py
│   ├── config.py               # 配置管理
│   ├── export_excel.py         # Excel 导出工具
│   ├── fetch_execute_bill_query.py
│   ├── fetch_form_field_keys.py
│   ├── fetch_view.py
│   └── kingdee_client.py       # 金蝶客户端封装
├── workspace/                  # 业务工作区
│   ├── BOM_API/                # BOM 数据同步与导入
│   │   ├── server.py           # Web 界面后端服务
│   │   ├── index.html          # Web 操作面板
│   │   ├── sync_and_generate.py # 主程序脚本
│   │   ├── map/                # 映射文件
│   │   ├── template/           # 金蝶导入模板
│   │   ├── data/               # 数据缓存
│   │   └── output/             # 输出目录
│   ├── COA-COC1/               # Test Plan 生成（预留）
│   ├── COA-COC2/               # 质检主数据同步与导入
│   └── Master-Data/            # 主数据处理（预留）
├── .env.example                # 环境变量模板
├── .gitignore
├── conf.ini.example            # SDK 配置模板
├── main.py                     # CLI 入口
├── requirements.txt            # 依赖清单
└── README.md
```

## 快速开始

### 环境准备

```bash
# 创建虚拟环境
python -m venv .venv
.venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt

# 复制配置文件
copy .env.example .env
copy conf.ini.example conf.ini
```

### 配置说明

1. **编辑 `conf.ini`**：填写金蝶云星空连接信息

```ini
[config]
X-KDApi-AcctID = your_dbid
X-KDApi-UserName = your_username
X-KDApi-AppID = your_app_id
X-KDApi-AppSec = your_app_secret
X-KDApi-LCID = 2052
X-KDApi-ServerUrl = https://your-host/K3Cloud/
```

2. **编辑 `.env`**：设置非敏感运行参数（可选）

```env
# 指定本地 SDK wheel 路径（可选）
# KINGDEE_SDK_WHEEL_PATH=.vendor/kingdee.cdp.webapi.sdk-8.2.0-py3-none-any.whl

# 查询参数示例
# KINGDEE_QUERY_JSON={"FormId":"BD_MATERIAL","FieldKeys":"FNumber,FName","Limit":10}
```

### 验证配置

```bash
python main.py check
```

## CLI 命令

### `check` - 检查配置

```bash
python main.py check
python main.py check --dry-run
```

检查 SDK 可用性、配置文件是否存在、必需键是否缺失。

### `query` - 执行 ExecuteBillQuery

```bash
# 先在 .env 中设置 KINGDEE_QUERY_JSON
python main.py query
python main.py query --dry-run  # 只解析参数，不发起请求
```

### `view` - 执行 View

```bash
# 先在 .env 中设置 KINGDEE_VIEW_FORM_ID 和 KINGDEE_VIEW_PARA_JSON
python main.py view
python main.py view --dry-run
```

### `fields` - 查询字段元数据

```bash
python main.py fields QM_QualityStd
```

返回指定业务对象的字段信息和可用 FieldKeys。

### `export-demo` - 导出演示数据

```bash
python main.py export-demo
```

生成示例 Excel（不连接金蝶）。

## 工作区模块

### COA-COC2 - 质检主数据同步与导入

从金蝶拉取质检主数据，与 Test Plan 对比，生成缺失数据的导入 Excel。

```bash
# 从金蝶拉取最新数据并生成导入文件
python workspace/COA-COC2/sync_and_generate.py

# 跳过在线拉取，使用本地缓存重新生成
python workspace/COA-COC2/sync_and_generate.py --skip-fetch
```

### COA-COC1 - Test Plan 生成（预留）

当前为需求 + 模板状态，等待后续实现脚本。

### Master-Data - 主数据处理（预留）

预留目录，用于后续主数据相关专项脚本。

### BOM_API - BOM 数据同步与导入

从金蝶拉取 BOM 及相关业务数据，生成导入 Excel。支持 9 种表单类型和中文/英文双语输出。

**Web 界面**（推荐）：

```bash
# 启动 Web 服务
python workspace/BOM_API/server.py
# 浏览器访问 http://localhost:5000
```

Web 界面提供可视化操作面板，支持：
- 复选框选择表单类型（9 种）和输出语言（CN/EN）
- 一键运行，实时进度条显示
- 运行完成后显示输出文件路径

#### Web 界面使用说明

1. **启动服务**：在项目根目录执行 `python workspace/BOM_API/server.py`，终端会显示访问地址
2. **打开页面**：浏览器访问 `http://localhost:5000`
3. **选择表单**：在"表单选项"卡片中勾选需要导出的表单类型，支持多选和全选
4. **选择语言**：在"语言选项"卡片中勾选输出语言（CN 中文 / EN 英文），支持多选
5. **确认运行**：点击"确认运行"按钮，进度条区域会显示实时运行状态
6. **查看进度**：运行过程中进度条实时更新，显示当前步骤和百分比
7. **停止运行**：运行中可点击"停止运行"按钮终止任务
8. **查看输出**：运行完成后自动显示输出目录路径和生成的文件列表，点击"打开输出目录"可复制路径
9. **重置选项**：点击"重置"按钮清空所有选择，恢复初始状态

#### 后端 API 说明

Web 服务基于 Flask，提供以下接口：

| 接口 | 方法 | 说明 |
|------|------|------|
| `/` | GET | 返回操作面板页面 |
| `/api/forms` | GET | 获取表单类型和语言选项列表 |
| `/api/run` | POST | 启动主程序，参数：`{"forms": ["ENG_BOM"], "langs": ["CN"], "network_output": "\\\\server\\share\\path"}` |
| `/api/progress` | GET | 获取运行进度，返回进度百分比和当前步骤 |
| `/api/output` | GET | 获取输出目录路径和文件列表 |
| `/api/stop` | POST | 停止当前运行中的任务 |
| `/api/clear-output` | POST | 清空本地和网络输出目录，参数：`{"network_output": "\\\\server\\share\\path"}` |
| `/api/open-output` | POST | 打开输出目录 |
| `/api/open-log` | GET | 打开日志目录 |
| `/api/clear-log` | POST | 清空日志目录 |

#### 网络路径同步

Web 界面支持将输出文件同步到 SMB 网络共享路径：

1. **设置网络路径**：在"输出位置"输入框中填写 UNC 路径（如 `\\192.168.1.181\99_Share\04_IT\KingdeeExport`）
2. **自动同步**：运行完成后自动将本地 output 目录中的 xlsx 文件复制到网络路径
3. **域认证**：系统自动使用域账号进行 SMB 认证（配置于 `server.py` 中）

#### 日志功能

每次运行完成后自动生成日志文件，保存在 `workspace/BOM_API/log/` 目录：

- 日志文件命名：`run_YYYYMMDD_HHMMSS.log`
- 日志内容：运行步骤、查询耗时、同步结果、错误信息

**命令行**：

```bash
# 导出所有表单（中英文），并同步到默认网络路径
python workspace/BOM_API/sync_and_generate.py

# 指定表单和语言
python workspace/BOM_API/sync_and_generate.py --form ENG_BOM --lang CN

# 跳过金蝶拉取，使用本地缓存
python workspace/BOM_API/sync_and_generate.py --skip-fetch

# 指定网络输出路径
python workspace/BOM_API/sync_and_generate.py --netaddr "\\server\share\path"

# 不同步到网络
python workspace/BOM_API/sync_and_generate.py --netaddr ""
```

**参数说明**：

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--form` | 指定表单类型，支持多选 | `all` |
| `--lang` | 输出语言（CN/EN） | `all` |
| `--skip-fetch` | 跳过金蝶拉取，使用本地缓存 | - |
| `--netaddr` | 网络输出路径（SMB共享） | `\\192.168.1.181\99_Share\04_IT\KingdeeExport` |

支持的表单类型：`BD_MATERIAL`（物料基础资料）、`ENG_BOM`（物料清单）、`PLN_FORECAST`（预测单）、`PLN_PLANORDER`（计划订单）、`PRD_MO`（生产订单）、`PUR_PurchaseOrder`（采购订单）、`PRD_INSTOCK`（生产入库单）、`STK_Inventory`（即时库存）、`PRD_PickMtrl`（生产领料单）

## SDK 加载顺序

代码会按以下顺序查找金蝶 SDK：

1. 已安装的 Python 包 `k3cloud_webapi_sdk`
2. `.env` 中配置的 `KINGDEE_SDK_WHEEL_PATH`
3. 项目 `.vendor/kingdee*.whl`（当前默认提供 8.2.0 版本）

## 注意事项

- **前置条件**：必须先在金蝶后台配置"第三方系统登录授权"
- **敏感信息**：`conf.ini` 包含真实凭证，已加入 `.gitignore`，不要提交
- **官方 SDK**：已随仓库提供 `.vendor/kingdee.cdp.webapi.sdk-8.2.0-py3-none-any.whl`
- **API 文档**：接口定义以 [金蝶云星空开放平台 API 文档](https://openapi.open.kingdee.com/ApiDoc) 为准

## 详细文档

- [docs/kingdee-python-integration.md](docs/kingdee-python-integration.md) - 详细集成说明
- [docs/project-tutorial.md](docs/project-tutorial.md) - 项目入门教程（适合首次接手）