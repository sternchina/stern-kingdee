# Master-Data 使用说明
# Master-Data Usage Guide

`Master-Data/` 当前是预留目录，用于后续放置主数据相关的专项脚本或处理工具。
`Master-Data/` is currently a reserved directory for future master-data-related scripts or processing tools.

当前仓库里，这个目录还没有脚本、模板、样例输入或固定的输入输出约定。
In the current repository, this directory does not yet contain a script, template, sample input, or a fixed input/output convention.

## 当前状态
## Current Status

- 这是一个占位目录。
- This is a placeholder directory.
- 目前不能直接运行。
- It cannot be run directly at the moment.
- 目前也没有特定业务流程写死在这里。
- There is also no fixed business flow defined here yet.

## 什么时候使用
## When to Use It

当后续出现“主数据拉取、清洗、比对、导入模板生成”一类的新需求时，可以优先考虑放在这里。
When new requirements appear for tasks such as “master data fetching, cleansing, comparison, or import template generation”, this directory is a reasonable place to put them.

## 建议目录结构
## Suggested Directory Structure

如果后续要在这里落地新工具，建议沿用现有 `workspace/` 的结构。
If a new tool is implemented here later, it is recommended to follow the existing `workspace/` structure.

```text
Master-Data/
  input/
  template/
  data/
  output/
  README.md
  需求.md
  your_script.py
```

## 建议使用方式
## Suggested Usage Pattern

后续若在这里新增脚本，建议 README 至少写清楚下面几项。
If a script is later added here, the README should at minimum document the following items.

- 这个工具处理哪类主数据。
- What type of master data the tool handles.
- 输入文件放在哪里。
- Where input files should be placed.
- 是否依赖金蝶连接。
- Whether it depends on a Kingdee connection.
- 如何运行。
- How to run it.
- 输出会生成到哪里。
- Where outputs will be generated.

## 注意事项
## Notes

- 当前目录为空是正常状态。
- It is normal that this directory is currently empty.
- 如果后续新增脚本，建议同步补 `README.md` 和 `需求.md`，不要只留代码不留使用说明。
- If a script is added later, update both `README.md` and `需求.md`; do not leave code without usage instructions.
