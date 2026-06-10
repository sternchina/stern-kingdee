import sys
sys.path.insert(0, r'c:\Users\shenjj\Desktop\stern-kingdee-main')

from src.kingdee_client import KingdeeClient

client = KingdeeClient()

print('正在查询所有单据类型...')
rows = client.execute_bill_query({
    "FormId": "BOS_BILLTYPE",
    "FieldKeys": "FBillFormID,FNumber,FName",
    "FilterString": "",
    "OrderString": "",
    "TopRowCount": 0,
    "StartRow": 0,
    "Limit": 1000,
    "SubSystemId": "",
})

print(f'共找到 {len(rows)} 个单据类型')

bill_type_mapping = {}
output_lines = []

for row in rows:
    bill_form_id = row[0]
    fnumber = row[1]
    fname = row[2]
    
    try:
        view_result = client.view("BOS_BILLTYPE", {"Number": fnumber, "Id": "", "CreateOrgId": 0})
        bill_id = view_result.get("Result", {}).get("Result", {}).get("Id", "")
        
        if bill_id:
            bill_type_mapping[bill_id] = fnumber
            output_lines.append(f"ID: {bill_id}")
            output_lines.append(f"  FNumber: {fnumber}")
            output_lines.append(f"  FName: {fname}")
            output_lines.append(f"  FBillFormID: {bill_form_id}")
            output_lines.append("")
    except Exception as e:
        output_lines.append(f"  查询失败: {fnumber} - {e}")
        output_lines.append("")

output_lines.append('-' * 80)
output_lines.append(f'共建立 {len(bill_type_mapping)} 条映射')
output_lines.append("")
output_lines.append('BILL_TYPE_MAPPING = {')
for bill_id, fnumber in bill_type_mapping.items():
    output_lines.append(f'    "{bill_id}": "{fnumber}",')
output_lines.append('}')

with open(r'c:\Users\shenjj\Desktop\stern-kingdee-main\bill_type_mapping.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(output_lines))

print(f'结果已保存到 bill_type_mapping.txt')
print(f'共建立 {len(bill_type_mapping)} 条映射')