# COA-COC1 使用说明
# COA-COC1 Usage Guide

`COA-COC1/` 用于“解析输入文件，并按指定模板生成 Test Plan”。
`COA-COC1/` is intended for “parsing input files and generating a Test Plan based on the specified template”.

当前目录已经包含模板、输入输出目录和需求说明，但仓库里还没有正式的可执行脚本入口。
This directory already contains the template, input/output folders, and requirement notes, but the repository does not yet include a formal runnable script entry point.

## 当前状态
## Current Status

- 当前是“需求 + 模板 + 目录结构”状态。
- It is currently in a “requirements + template + directory structure” state.
- 还没有现成的 `python xxx.py` 命令可以直接运行。
- There is not yet a ready-made `python xxx.py` command that can be run directly.
- 如果后续实现脚本，建议放在本目录下，并让输入输出继续沿用现有目录结构。
- If a script is implemented later, it should be placed in this directory and continue to use the existing input/output structure.

## 目录说明
## Directory Layout

- `input/`：放待解析的源文件。
- `input/`: Stores the source files to be parsed.
- `template/`：放目标模板，当前已有 `20251231_test plan_upload.xlsx`。
- `template/`: Stores the target template, and currently includes `20251231_test plan_upload.xlsx`.
- `output/`：放生成后的 Test Plan 文件。
- `output/`: Stores the generated Test Plan files.
- `需求.md`：当前业务需求说明。
- `需求.md`: The current business requirement description.

## 如何使用
## How to Use

当前推荐按下面方式使用本目录。
At the moment, the recommended way to use this directory is as follows.

1. 先阅读 `需求.md`，确认输入文件类型、解析规则和目标模板格式。
1. First read `需求.md` to confirm the input file type, parsing rules, and target template format.
2. 把待处理的源文件放进 `input/`。
2. Put the source files to be processed into `input/`.
3. 参考 `template/20251231_test plan_upload.xlsx` 设计输出字段和排版。
3. Use `template/20251231_test plan_upload.xlsx` as the target format for output fields and layout.
4. 在本目录下新增脚本后，再约定正式运行命令。
4. After a script is added to this directory, define the formal run command.

## 后续实现建议
## Suggested Next Step

建议后续在本目录下新增一个主脚本，例如 `generate_test_plan.py`，职责保持简单。
It is recommended to add a main script here later, for example `generate_test_plan.py`, while keeping its responsibility simple.

- 从 `input/` 读取源文件。
- Read source files from `input/`.
- 调用根目录公共能力或外部接口做解析。
- Use shared capabilities from the repository root or external APIs for parsing.
- 按 `template/` 模板格式生成文件到 `output/`。
- Generate files into `output/` using the format defined in `template/`.

## 注意事项
## Notes

- 当前目录还不能直接独立运行。
- This directory cannot yet be run independently.
- 如果要补实现，优先先写 `README.md` 中的运行命令，再补脚本本身。
- If you implement it later, define the run command in `README.md` first, then add the script itself.
- 如果生成逻辑依赖金蝶数据，建议先在项目根目录运行 `python main.py check`。
- If the generation logic will depend on Kingdee data, it is recommended to run `python main.py check` from the repository root first.
