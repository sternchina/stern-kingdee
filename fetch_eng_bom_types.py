import sys
sys.path.insert(0, r'c:\Users\shenjj\Desktop\stern-kingdee-main')

from src.kingdee_client import KingdeeClient

client = KingdeeClient()

print('查询 ENG_BOM 相关的单据类型...')
rows = client.execute_bill_query({
    "FormId": "BOS_BILLTYPE",
    "FieldKeys": "FBillFormID,FNumber,FName",
    "FilterString": "FBillFormID='ENG_BOM'",
    "OrderString": "",
    "TopRowCount": 0,
    "StartRow": 0,
    "Limit": 10,
    "SubSystemId": "",
})

print(f'找到 {len(rows)} 个 ENG_BOM 单据类型')

bill_type_mapping = {}
for row in rows:
    bill_form_id = row[0]
    fnumber = row[1]
    fname = row[2]
    
    print(f'\nFNumber: {fnumber}')
    print(f'FName: {fname}')
    print(f'FBillFormID: {bill_form_id}')
    
    try:
        view_result = client.view("BOS_BILLTYPE", {"Number": fnumber, "Id": "", "CreateOrgId": 0})
        bill_id = view_result.get("Result", {}).get("Result", {}).get("Id", "")
        
        if bill_id:
            bill_type_mapping[bill_id] = fnumber
            print(f'ID: {bill_id}')
        else:
            print('ID: 未找到')
    except Exception as e:
        print(f'查询失败: {e}')

print('\n' + '-' * 60)
print('BILL_TYPE_MAPPING = {')
for bill_id, fnumber in bill_type_mapping.items():
    print(f'    "{bill_id}": "{fnumber}",')
print('}')