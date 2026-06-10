import sys
sys.path.insert(0, r'c:\Users\shenjj\Desktop\stern-kingdee-main')

# 运行脚本并捕获输出
from io import StringIO
old_stdout = sys.stdout
sys.stdout = StringIO()

try:
    from workspace.BOM_API.sync_and_generate import fetch_all_from_kingdee
    fetch_all_from_kingdee()
    
    output = sys.stdout.getvalue()
    with open('script_output.txt', 'w', encoding='utf-8') as f:
        f.write(output)
    
    print('脚本输出已保存到 script_output.txt')
finally:
    sys.stdout = old_stdout