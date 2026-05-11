import win32com.client
import os
from pathlib import Path

source_dir = r"E:\Ashuxuejianmo\A题\附件1 人口普查数据集"
dest_dir = r"E:\Ashuxuejianmo\A题_整理后\附件1 人口普查数据集"

os.makedirs(dest_dir, exist_ok=True)

excel_files = []
for root, dirs, files in os.walk(source_dir):
    for f in files:
        if f.endswith('.xls') and not f.endswith('.xlsx'):
            excel_files.append(os.path.join(root, f))

print(f"找到 {len(excel_files)} 个 .xls 文件")
print("正在用同一个Excel实例转换...")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

success = 0
failed = 0

for i, xls_file in enumerate(excel_files):
    try:
        rel_path = os.path.relpath(xls_file, source_dir)
        csv_dir = os.path.join(dest_dir, os.path.dirname(rel_path))
        os.makedirs(csv_dir, exist_ok=True)
        
        base_name = os.path.splitext(os.path.basename(xls_file))[0]
        csv_file = os.path.join(csv_dir, base_name + '.csv')
        
        wb = excel.Workbooks.Open(xls_file)
        wb.SaveAs(csv_file, 6)
        wb.Close()
        
        success += 1
        print(f"[{i+1}/{len(excel_files)}] OK: {base_name}")
        
    except Exception as e:
        failed += 1
        print(f"[{i+1}/{len(excel_files)}] 失败: {os.path.basename(xls_file)}")

excel.Quit()

print(f"\n完成! 成功: {success}, 失败: {failed}")