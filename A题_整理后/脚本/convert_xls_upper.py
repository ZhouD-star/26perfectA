import win32com.client
import os

source_dir = r"E:\Ashuxuejianmo\A题\附件1 人口普查数据集\5.五普"
dest_dir = r"E:\Ashuxuejianmo\A题_整理后\附件1 人口普查数据集\5.五普"

os.makedirs(dest_dir, exist_ok=True)

excel_files = []
for root, dirs, files in os.walk(source_dir):
    for f in files:
        if f.upper().endswith('.XLS'):
            excel_files.append(os.path.join(root, f))

print(f"找到 {len(excel_files)} 个 .XLS 文件")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

for i, xls_file in enumerate(excel_files):
    try:
        base_name = os.path.splitext(os.path.basename(xls_file))[0]
        csv_file = os.path.join(dest_dir, base_name + '.csv')
        
        wb = excel.Workbooks.Open(xls_file)
        wb.SaveAs(csv_file, 6)
        wb.Close()
        
        print(f"[{i+1}/{len(excel_files)}] OK: {base_name}")
    except Exception as e:
        print(f"[{i+1}/{len(excel_files)}] 失败: {os.path.basename(xls_file)}")

excel.Quit()
print("完成!")