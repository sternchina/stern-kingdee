import sys
sys.path.insert(0, r'c:\Users\shenjj\Desktop\stern-kingdee-main')

from src.kingdee_client import KingdeeClient
import pandas as pd

client = KingdeeClient()

df = pd.read_csv(r'c:\Users\shenjj\Desktop\stern-kingdee-main\workspace\BOM_API\data\ENG_BOM_20260429_162137.csv', encoding='utf-8-sig')

# 查询计量单位
unit_ids = df['FUNITID'].unique()
print('查询计量单位 (BD_UNIT):')
print('-' * 60)

for unit_id in unit_ids:
    try:
        result = client.view('BD_UNIT', {'Id': str(unit_id), 'Number': '', 'CreateOrgId': 0})
        data = result.get('Result', {}).get('Result', {})
        fnumber = data.get('Number', '')
        fname = data.get('Name', [{}])[0].get('Value', '') if data.get('Name') else ''
        print(f"ID: {unit_id}")
        print(f"  FNumber: {fnumber}")
        print(f"  FName: {fname}")
        print()
    except Exception as e:
        print(f"ID: {unit_id} - 查询失败: {e}")
        print()

# 查询BOM分组
group_ids = df['FGroup'].unique()
print('查询BOM分组 (ENG_BOMGROUP):')
print('-' * 60)

for group_id in group_ids:
    try:
        result = client.view('ENG_BOMGROUP', {'Id': str(group_id), 'Number': '', 'CreateOrgId': 0})
        data = result.get('Result', {}).get('Result', {})
        fnumber = data.get('Number', '')
        fname = data.get('Name', [{}])[0].get('Value', '') if data.get('Name') else ''
        print(f"ID: {group_id}")
        print(f"  FNumber: {fnumber}")
        print(f"  FName: {fname}")
        print()
    except Exception as e:
        print(f"ID: {group_id} - 查询失败: {e}")
        print()