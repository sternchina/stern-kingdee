import zipfile
import os
from pathlib import Path
from datetime import datetime

project_root = Path(r'c:\Users\shenjj\Desktop\stern-kingdee-main')
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
backup_name = f'project_backup_{timestamp}.zip'
backup_path = project_root / backup_name

# 忽略的文件和目录
exclude_patterns = {
    '.venv', 'venv', '__pycache__', '.git', 'node_modules',
    '.pytest_cache', '.mypy_cache', '.coverage', 'htmlcov',
    '.env', '*.pyc', '*.pyo', '*.pyd', '.DS_Store', 'Thumbs.db',
    '*.log', '*.xlsx', '*.csv', '.idea', '.vscode'
}

def should_exclude(path):
    name = path.name
    for pattern in exclude_patterns:
        if pattern.startswith('*'):
            if name.endswith(pattern[1:]):
                return True
        elif name == pattern:
            return True
    return False

print(f'正在创建备份: {backup_name}')
print(f'项目根目录: {project_root}')

with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
    count = 0
    for root, dirs, files in os.walk(project_root):
        # 过滤目录
        dirs[:] = [d for d in dirs if not should_exclude(Path(d))]
        
        for file in files:
            file_path = Path(root) / file
            if should_exclude(file_path):
                continue
            
            arcname = file_path.relative_to(project_root)
            zipf.write(file_path, arcname)
            count += 1
            
            if count % 100 == 0:
                print(f'  已添加 {count} 个文件...')

print(f'\n✅ 备份完成！')
print(f'文件: {backup_name}')
print(f'大小: {backup_path.stat().st_size / 1024 / 1024:.2f} MB')
print(f'包含 {count} 个文件')