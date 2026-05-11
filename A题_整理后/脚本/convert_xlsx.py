import pandas as pd
from pathlib import Path

source_dir = r"E:\Ashuxuejianmo\A题\附件1 人口普查数据集\7.七普"
dest_dir = r"E:\Ashuxuejianmo\A题_整理后\附件1 人口普查数据集\7.七普"

Path(dest_dir).mkdir(parents=True, exist_ok=True)

xlsx_files = list(Path(source_dir).glob("*.xlsx"))
print(f"找到 {len(xlsx_files)} 个 xlsx 文件")

for xlsx_file in xlsx_files:
    try:
        excel_file = pd.ExcelFile(xlsx_file)
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            
            if len(excel_file.sheet_names) > 1:
                csv_file = Path(dest_dir) / f"{xlsx_file.stem}_{sheet_name}.csv"
            else:
                csv_file = Path(dest_dir) / f"{xlsx_file.stem}.csv"
            
            df.to_csv(csv_file, index=False, header=False, encoding='utf-8-sig')
            print(f"OK: {csv_file.name}")
            
    except Exception as e:
        print(f"失败: {xlsx_file.name} - {e}")

print("完成!")