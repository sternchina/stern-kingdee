import sys
sys.path.insert(0, r'c:\Users\shenjj\Desktop\stern-kingdee-main')

from src.kingdee_client import KingdeeClient
import pandas as pd

client = KingdeeClient()

df = pd.read_csv(r'c:\Users\shenjj\Desktop\stern-kingdee-main\workspace\BOM_API\data\ENG_BOM_20260429_165141.csv', encoding='utf-8-sig')

# 定义需要查询的字段及其对应的 FormId
field_mappings = {
    'FMATERIALID': 'BD_MATERIAL',
    'FUNITID': 'BD_UNIT',
    'FGroup': 'ENG_BOMGROUP',
    'FBOMCATEGORY': 'ENG_BOMCATEGORY',
    'FBOMUSE': 'ENG_BOMUSE',
}

all_mappings = {}

for field_name, form_id in field_mappings.items():
    if field_name not in df.columns:
        continue
    
    unique_ids = df[field_name].unique()
    print(f'正在查询 {field_name} ({form_id}): {len(unique_ids)} 个唯一值')
    
    mappings = {}
    for id_val in unique_ids:
        id_str = str(id_val).strip()
        if not id_str or id_str == 'nan':
            continue
            
        try:
            result = client.view(form_id, {'Id': id_str, 'Number': '', 'CreateOrgId': 0})
            data = result.get('Result', {}).get('Result', {})
            fnumber = data.get('Number', '')
            
            if fnumber:
                mappings[id_str] = fnumber
                print(f'  {id_str} -> {fnumber}')
            else:
                print(f'  {id_str} -> 未找到编码')
        except Exception as e:
            print(f'  {id_str} -> 查询失败: {str(e)[:30]}')
    
    if mappings:
        all_mappings[field_name] = mappings
    print()

print('=' * 80)
print('完整映射字典:')
print('=' * 80)
for field_name, mappings in all_mappings.items():
    print(f'{field_name}_MAPPING = {{')
    for id_val, code in mappings.items():
        print(f'    "{id_val}": "{code}",')
    print('}')
    print()