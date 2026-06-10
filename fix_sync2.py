import re

# 读取文件
with open(r'c:\Users\shenjj\Desktop\stern-kingdee-main\workspace\COA-COC2\sync_and_generate.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 在 df = pd.read_excel 之后添加列名标准化
old_text = '''    df = pd.read_excel(xl, sheet_name=matched_sheet, dtype=str)
    for col in df.columns:
        df[col] = df[col].fillna("").str.strip()
    print(f"Test plan ({path.name} / {matched_sheet}): {len(df)} 行，列：{df.columns.tolist()}")'''

new_text = '''    df = pd.read_excel(xl, sheet_name=matched_sheet, dtype=str)
    for col in df.columns:
        df[col] = df[col].fillna("").str.strip()
    df = df.rename(columns=_COL_ALIASES)
    print(f"Test plan ({path.name} / {matched_sheet}): {len(df)} 行，列：{df.columns.tolist()}")'''

content = content.replace(old_text, new_text)

# 保存文件
with open(r'c:\Users\shenjj\Desktop\stern-kingdee-main\workspace\COA-COC2\sync_and_generate.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('修改完成')