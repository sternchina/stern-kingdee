"""使用subprocess运行脚本"""
import subprocess
import sys

# 运行脚本
result = subprocess.run(
    [sys.executable, r'c:\Users\shenjj\Desktop\stern-kingdee-main\workspace\BOM_API\generate_excel.py'],
    capture_output=True,
    text=True,
    encoding='utf-8'
)

print("=== STDOUT ===")
print(result.stdout)

print("\n=== STDERR ===")
print(result.stderr)

print(f"\n=== 返回码: {result.returncode} ===")
