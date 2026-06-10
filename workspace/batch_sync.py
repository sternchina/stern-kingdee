"""
批量拉取多个 FormId 的数据并生成导入文件
每个 FormId 使用自己的模板并生成单独的文件
"""
import subprocess
import sys
from pathlib import Path

# 定义要批量处理的 FormId 列表
FORM_IDS = [
    "ENG_BOM",
    "BD_MATERIAL",
    "BD_UNIT",
    "QM_InspectMethod",
    "QM_InspectItem",
    "QM_QCScheme",
    "QM_QualityStd"
]

def main():
    print("开始批量拉取 FormId 数据...")
    print(f"共处理 {len(FORM_IDS)} 个 FormId")
    print("=" * 80)
    
    for form_id in FORM_IDS:
        print(f"\n处理 FormId: {form_id}")
        print("-" * 60)
        
        # 执行同步脚本
        cmd = [
            sys.executable,
            str(Path(__file__).parent / "TEMPLATE" / "sync_and_generate.py"),
            "--form-id", form_id
        ]
        
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            print(result.stdout)
            if result.stderr:
                print("错误信息:", result.stderr)
        except subprocess.CalledProcessError as e:
            print(f"处理 {form_id} 时出错:")
            print(e.stdout)
            print(e.stderr)
    
    print("\n" + "=" * 80)
    print("批量处理完成！")
    print("每个 FormId 的数据和导入文件已分别生成到对应目录")

if __name__ == "__main__":
    main()
