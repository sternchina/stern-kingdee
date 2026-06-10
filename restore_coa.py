# 还原 COA-COC2 脚本到上一个版本

with open(r'c:\Users\shenjj\Desktop\stern-kingdee-main\workspace\COA-COC2\sync_and_generate.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 移除 _COL_ALIASES 字典
old_text1 = '''_REQUIRED_COLS = {"material_code", "material_name", "test_item", "target", "unit", "method"}

_COL_ALIASES = {
    "items_checked": "test_item",
}


def parse_test_plan()'''

new_text1 = '''_REQUIRED_COLS = {"material_code", "material_name", "test_item", "target", "unit", "method"}


def parse_test_plan()'''

content = content.replace(old_text1, new_text1)

# 移除列名标准化逻辑
old_text2 = '''            cols = set(head.columns)
            cols_normalized = {_COL_ALIASES.get(c, c) for c in cols}
            if not _REQUIRED_COLS.issubset(cols_normalized):'''

new_text2 = '''            cols = set(head.columns)
            if not _REQUIRED_COLS.issubset(cols):'''

content = content.replace(old_text2, new_text2)

# 移除 df.rename(columns=_COL_ALIASES)
old_text3 = '''    df = pd.read_excel(xl, sheet_name=matched_sheet, dtype=str)
    for col in df.columns:
        df[col] = df[col].fillna("").str.strip()
    df = df.rename(columns=_COL_ALIASES)
    print(f"Test plan ({path.name} / {matched_sheet}): {len(df)} 行，列：{df.columns.tolist()}")'''

new_text3 = '''    df = pd.read_excel(xl, sheet_name=matched_sheet, dtype=str)
    for col in df.columns:
        df[col] = df[col].fillna("").str.strip()
    print(f"Test plan ({path.name} / {matched_sheet}): {len(df)} 行，列：{df.columns.tolist()}")'''

content = content.replace(old_text3, new_text3)

# 保存文件
with open(r'c:\Users\shenjj\Desktop\stern-kingdee-main\workspace\COA-COC2\sync_and_generate.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('COA-COC2 已还原到上一个版本')