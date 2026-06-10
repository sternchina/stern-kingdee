# 金蝶云星空 Python SDK 集成说明
# Kingdee K3 Cloud Python SDK Integration Guide

## 当前实现
## Current Implementation

本仓库当前不再自行拼接登录请求，而是通过官方 Python SDK `k3cloud_webapi_sdk` 调用金蝶云星空接口。
This repository no longer builds login requests manually and now calls Kingdee K3 Cloud through the official Python SDK `k3cloud_webapi_sdk`.

现有封装位置如下。
The current wrappers are located in the following files.

- `src/kingdee_client.py`：懒加载 SDK，读取 `conf.ini`，封装 `ExecuteBillQuery`、`View`、`QueryBusinessInfo` 与字段提取。
- `src/kingdee_client.py`: Lazily loads the SDK, reads `conf.ini`, and wraps `ExecuteBillQuery`, `View`, `QueryBusinessInfo`, and field extraction.
- `main.py`：命令行入口。
- `main.py`: Command-line entry point.
- `.env.example`：非敏感运行参数模板。
- `.env.example`: Template for non-sensitive runtime parameters.
- `conf.ini.example`：SDK 登录配置模板。
- `conf.ini.example`: Template for SDK login configuration.

## 前置条件
## Prerequisites

在执行业务操作前，请先满足以下前提。
Before running business operations, make sure the following prerequisites are met.

1. 在金蝶云星空后台完成“第三方系统登录授权”配置。
1. Configure "Third-Party System Login Authorization" in the Kingdee K3 Cloud backend.
2. 仓库已自带固定版本的官方 Python SDK wheel；如需切换版本，再准备其他 wheel。
2. The repository already includes a pinned official Python SDK wheel; only prepare another wheel if you need to switch versions.

默认情况下，直接使用仓库内置的 wheel 即可。
By default, you can use the wheel bundled in the repository.

```text
.vendor/kingdee.cdp.webapi.sdk-8.2.0-py3-none-any.whl
```

如果你需要替换版本，仓库也支持以下两种覆盖方式。
If you need to replace that version, the repository also supports the following override options.

```bash
pip install F:\path\to\k3cloud_webapi_sdk-1.0.0-py3-none-any.whl
```

或者在 `.env` 中设置 `KINGDEE_SDK_WHEEL_PATH` 指向其他本地 wheel。
Or set `KINGDEE_SDK_WHEEL_PATH` in `.env` to point to another local wheel.

代码会按以下顺序查找 SDK。
The code looks for the SDK in the following order.

1. 已安装的 Python 包。
1. An installed Python package.
2. `.env` 中的 `KINGDEE_SDK_WHEEL_PATH`。
2. `KINGDEE_SDK_WHEEL_PATH` from `.env`.
3. 项目 `.vendor/kingdee*.whl`，当前仓库默认提供 `8.2.0`。
3. A matching `.vendor/kingdee*.whl` file in the project; this repository currently ships `8.2.0` by default.

## 配置方式
## Configuration

### `conf.ini`
### `conf.ini`

敏感信息不再放在 `.env`，而是放在 SDK 的登录配置文件中。
Sensitive values are no longer stored in `.env` and are instead placed in the SDK login configuration file.

先复制模板。
Copy the template first.

```bash
copy conf.ini.example conf.ini
```

示例如下。
An example is shown below.

```ini
[config]
X-KDApi-AcctID = your_dbid
X-KDApi-UserName = your_username
X-KDApi-AppID = your_app_id
X-KDApi-AppSec = your_app_secret
X-KDApi-LCID = 2052
X-KDApi-ServerUrl = https://your-host/K3Cloud/
```

默认读取配置文件路径 `conf.ini` 与节点 `[config]`。
The default configuration path is `conf.ini` and the default section is `[config]`.

### `.env`
### `.env`

`.env` 只放非敏感运行参数。
`.env` stores only non-sensitive runtime parameters.

先复制模板。
Copy the template first.

```bash
copy .env.example .env
```

支持的键如下。
The supported keys are listed below.

- `KINGDEE_SDK_WHEEL_PATH`：可选，用于覆盖仓库内置 wheel，指定其他本地 SDK wheel 路径。
- `KINGDEE_SDK_WHEEL_PATH`: Optional path to another local SDK wheel that overrides the bundled one.
- `KINGDEE_CONFIG_PATH`：SDK 配置文件路径，默认 `conf.ini`。
- `KINGDEE_CONFIG_PATH`: SDK configuration file path, defaulting to `conf.ini`.
- `KINGDEE_CONFIG_NODE`：配置节点，默认 `config`。
- `KINGDEE_CONFIG_NODE`: Configuration section name, defaulting to `config`.
- `KINGDEE_QUERY_JSON`：`ExecuteBillQuery` 的参数 JSON。
- `KINGDEE_QUERY_JSON`: JSON payload for `ExecuteBillQuery`.
- `KINGDEE_VIEW_FORM_ID`：`View` 的业务对象标识。
- `KINGDEE_VIEW_FORM_ID`: Form identifier used by `View`.
- `KINGDEE_VIEW_PARA_JSON`：`View` 的参数 JSON。
- `KINGDEE_VIEW_PARA_JSON`: JSON payload used by `View`.

## 命令行用法
## CLI Usage

### 检查配置
### Check Configuration

执行以下命令检查配置。
Run the following command to check the configuration.

```bash
python main.py check
```

该命令会检查以下内容。
This command checks the following items.

- 官方 SDK 是否已安装。
- Whether the official SDK is installed.
- 是否发现可用的本地 SDK wheel（包括仓库内置 wheel）。
- Whether an available local SDK wheel can be found, including the bundled wheel in the repository.
- `conf.ini` 是否存在。
- Whether `conf.ini` exists.
- 指定节名是否存在。
- Whether the specified config section exists.
- 必需键是否缺失。
- Whether required keys are missing.

这里不会单独发起登录测试，SDK 会在真正调用业务方法时隐式处理登录。
This command does not perform a standalone login test, because the SDK handles login implicitly when a business method is called.

### 执行 `ExecuteBillQuery`
### Run `ExecuteBillQuery`

`.env` 示例。
Example `.env` value.

```env
KINGDEE_QUERY_JSON={"FormId":"BD_MATERIAL","FieldKeys":"FNumber,FName","FilterString":"","OrderString":"","TopRowCount":0,"StartRow":0,"Limit":10,"SubSystemId":""}
```

执行以下命令发起请求。
Run the following command to send the request.

```bash
python main.py query
```

执行以下命令只看参数而不请求。
Run the following command to inspect the parameters without making the request.

```bash
python main.py query --dry-run
```

### 执行 `View`
### Run `View`

`.env` 示例。
Example `.env` value.

```env
KINGDEE_VIEW_FORM_ID=QM_QualityStd
KINGDEE_VIEW_PARA_JSON={"CreateOrgId":0,"Number":"","Id":"","IsSortBySeq":"false"}
```

执行以下命令发起请求。
Run the following command to send the request.

```bash
python main.py view
```

执行以下命令只看参数而不请求。
Run the following command to inspect the parameters without making the request.

```bash
python main.py view --dry-run
```

`View` 一般至少需要有效的 `Id` 或 `Number`，否则通常查不到具体数据。
`View` usually requires a valid `Id` or `Number`; otherwise it typically cannot return a concrete record.

### 通过 `FormId` 查询字段与 `FieldKeys`
### Query Fields and `FieldKeys` by `FormId`

执行以下命令。
Run the following command.

```bash
python main.py fields QM_QualityStd
```

返回结果会包含以下内容。
The result includes the following fields.

- `form_id`、`form_name`。
- `form_id` and `form_name`.
- `field_keys`：扁平字段列表。
- `field_keys`: A flattened list of fields.
- `field_keys_csv`：可直接复制给 `ExecuteBillQuery` / `BillQuery` 的逗号串。
- `field_keys_csv`: A comma-separated string that can be copied directly into `ExecuteBillQuery` or `BillQuery`.
- `entries`：按实体分组的字段明细。
- `entries`: Field details grouped by entity.

这里底层走的是 `QueryBusinessInfo({"FormId": "..."})`。
Under the hood, this uses `QueryBusinessInfo({"FormId": "..."})`.

它能查到字段元数据，但不代表所有字段都适合直接放进查询语句，复杂引用字段仍建议实测。
It can return field metadata, but that does not guarantee every field is suitable for direct querying, so complex reference fields should still be validated in practice.

## 代码调用示例
## Code Examples

### 直接使用仓库封装
### Use the Repository Wrappers Directly

```python
from src.fetch_execute_bill_query import fetch_execute_bill_query
from src.fetch_form_field_keys import fetch_form_field_keys
from src.fetch_view import fetch_view

rows = fetch_execute_bill_query(
    {
        "FormId": "BD_MATERIAL",
        "FieldKeys": "FNumber,FName",
        "FilterString": "",
        "OrderString": "",
        "TopRowCount": 0,
        "StartRow": 0,
        "Limit": 10,
        "SubSystemId": "",
    }
)

detail = fetch_view(
    "QM_QualityStd",
    {"CreateOrgId": 0, "Number": "STD0001", "Id": "", "IsSortBySeq": "false"},
)

field_info = fetch_form_field_keys("QM_QualityStd")
print(field_info["field_keys_csv"])
```

### 直接使用官方 SDK
### Use the Official SDK Directly

```python
import json
import configparser
from k3cloud_webapi_sdk.main import K3CloudApiSdk

conf = configparser.ConfigParser()
conf.read("conf.ini", encoding="utf-8")
server_url = conf.get("config", "X-KDApi-ServerUrl")

api_sdk = K3CloudApiSdk(server_url)
api_sdk.Init(config_path="conf.ini", config_node="config")

query_para = {
    "FormId": "BD_MATERIAL",
    "FieldKeys": "FNumber,FName",
    "FilterString": "",
    "OrderString": "",
    "TopRowCount": 0,
    "StartRow": 0,
    "Limit": 10,
    "SubSystemId": "",
}
query_response = api_sdk.ExecuteBillQuery(query_para)

view_para = {"CreateOrgId": 0, "Number": "STD0001", "Id": "", "IsSortBySeq": "false"}
view_response = api_sdk.View("QM_QualityStd", view_para)
field_response = api_sdk.QueryBusinessInfo({"FormId": "QM_QualityStd"})

print(json.dumps({"query": query_response, "view": view_response, "fields": field_response}, ensure_ascii=False, indent=2))
```

## 注意事项
## Notes

- `conf.ini` 含真实凭证，已加入 `.gitignore`，不要提交。
- `conf.ini` contains real credentials, is already in `.gitignore`, and must not be committed.
- 仓库当前已提交 `.vendor/kingdee.cdp.webapi.sdk-8.2.0-py3-none-any.whl`，通常不需要再手动放置 wheel。
- The repository currently includes `.vendor/kingdee.cdp.webapi.sdk-8.2.0-py3-none-any.whl`, so you usually do not need to place another wheel manually.
- 官方 SDK 的返回值在不同接口上可能是 JSON 字符串或已解析对象，仓库封装会尽量自动做一次 JSON 解析。
- The official SDK may return either a JSON string or an already parsed object depending on the API, and the repository wrapper tries to normalize it with one JSON parse step.
- 本仓库当前已封装 `ExecuteBillQuery`、`View`、`QueryBusinessInfo` 以及按 `FormId` 提取 `FieldKeys`。
- This repository currently wraps `ExecuteBillQuery`, `View`, `QueryBusinessInfo`, and `FieldKeys` extraction by `FormId`.
- 如果后续要扩展 `Save`、`Submit`、`Audit` 等接口，建议继续走同一个 SDK 包装层。
- If you later extend `Save`, `Submit`, `Audit`, or similar APIs, keep using the same SDK wrapper layer.
- 业务对象、参数结构、返回结构仍以金蝶开放平台 API 文档和你方实际环境为准。
- Business objects, request payloads, and response structures still depend on the Kingdee Open Platform documentation and your actual environment.
