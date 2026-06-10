import sys
sys.path.insert(0, r'c:\Users\shenjj\Desktop\stern-kingdee-main')

from src.kingdee_client import KingdeeClient
import json

client = KingdeeClient()

print("=== 解析物料清单中的值 ===\n")

# 1. WLQD01_SYS - 单据类型
print("1. WLQD01_SYS (单据类型)")
try:
    result = client.view('BOS_BILLTYPE', {'Number': 'WLQD01_SYS', 'Id': '', 'CreateOrgId': 0})
    bill_type = result.get('Result', {}).get('Result', {})
    print("   编码:", bill_type.get('Number'))
    name_info = bill_type.get('Name', '')
    if isinstance(name_info, list):
        for item in name_info:
            if item.get('Key') == 2052:
                print("   名称:", item.get('Value'))
    else:
        print("   名称:", name_info)
    form_id_info = bill_type.get('BillFormID', {})
    print("   表单ID:", form_id_info.get('Number'))
except Exception as e:
    print("   查询失败:", e)

print()

# 2. 1 - BOM类别
print("2. 1 (BOM类别)")
print("   --- 尝试 ExecuteBillQuery 查询所有数据 ---")
try:
    query = {
        "FormId": "ENG_BOMCATEGORY",
        "FieldKeys": "FNumber,FName,FId",
        "FilterString": "",
        "OrderString": "FId",
        "TopRowCount": 100
    }
    result = client.execute_bill_query(query)
    print("   返回数据结构:", type(result))
    print("   返回数据:", json.dumps(result, ensure_ascii=False, indent=2)[:1000])
except Exception as e:
    print("   查询失败:", e)

print()

# 3. 99 - BOM用途
print("3. 99 (BOM用途)")
print("   --- 尝试 ExecuteBillQuery 查询所有数据 ---")
try:
    query = {
        "FormId": "ENG_BOMUSE",
        "FieldKeys": "FNumber,FName,FId",
        "FilterString": "",
        "OrderString": "FId",
        "TopRowCount": 100
    }
    result = client.execute_bill_query(query)
    print("   返回数据结构:", type(result))
    print("   返回数据:", json.dumps(result, ensure_ascii=False, indent=2)[:1000])
except Exception as e:
    print("   查询失败:", e)

print("\n=== 查看BOM系统参数 ===")
print("   这些基础资料可能在系统参数中定义")
print("   BOM类别 1 通常表示 '标准BOM'")
print("   BOM用途 99 通常表示 '通用' 或 '其他'")