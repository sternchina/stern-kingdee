import sys
sys.path.insert(0, r'c:\Users\shenjj\Desktop\stern-kingdee-main')

from src.kingdee_client import KingdeeClient
import pandas as pd

client = KingdeeClient()

df = pd.read_csv(r'c:\Users\shenjj\Desktop\stern-kingdee-main\workspace\BOM_API\data\ENG_BOM_20260429_162137.csv', encoding='utf-8-sig')

# 查询物料编码
material_ids = df['FMATERIALID'].unique()[:5]
print('查询物料 (BD_MATERIAL):')
print('-' * 60)

for mat_id in material_ids:
    try:
        result = client.view('BD_MATERIAL', {'Id': str(mat_id), 'Number': '', 'CreateOrgId': 0})
        data = result.get('Result', {}).get('Result', {})
        fnumber = data.get('Number', '')
        fname = data.get('Name', [{}])[0].get('Value', '') if data.get('Name') else ''
        print(f"ID: {mat_id}")
        print(f"  FNumber: {fnumber}")
        print(f"  FName: {fname}")
        print()
    except Exception as e:
        print(f"ID: {mat_id} - 查询失败: {e}")
        print()