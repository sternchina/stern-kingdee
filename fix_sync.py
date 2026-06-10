import re

# 读取文件
with open(r'c:\Users\shenjj\Desktop\stern-kingdee-main\workspace\COA-COC2\sync_and_generate.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 找到 _REQUIRED_COLS 的位置并添加 _COL_ALIASES
old_text = '''_REQUIRED_COLS = {"material_code", "material_name", "test_item", "target", "unit", "method"}


def parse_test_plan()'''

new_text = '''_REQUIRED_COLS = {"material_code", "material_name", "test_item", "target", "unit", "method"}

_COL_ALIASES = {
    "items_checked": "test_item",
}


def parse_test_plan()'''

content = content.replace(old_text, new_text)

# 找到 cols = set(head.columns) 并在其后添加标准化逻辑
old_parse = '''            cols = set(head.columns)
            if not _REQUIRED_COLS.issubset(cols):'''

new_parse = '''            cols = set(head.columns)
            cols_normalized = {_COL_ALIASES.get(c, c) for c in cols}
            if not _REQUIRED_COLS.issubset(cols_normalized):'''

content = content.replace(old_parse, new_parse)

# 保存文件
with open(r'c:\Users\shenjj\Desktop\stern-kingdee-main\workspace\COA-COC2\sync_and_generate.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('修改完成')