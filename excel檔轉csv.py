import os
import pandas as pd
from openpyxl import load_workbook

# 資料夾路徑
folder_path = r'C:\_python 2024\2024final'

# 遍歷資料夾中所有的 Excel 檔案 (.xlsx, .xls)
for filename in os.listdir(folder_path):
    if filename.lower().endswith(('.xlsx', '.xls')):
        excel_path = os.path.join(folder_path, filename)
        base_name, _ = os.path.splitext(filename)

        # 1. 使用 openpyxl 讀取並重命名第一個工作表為 sheet1
        wb = load_workbook(excel_path)
        first_sheet = wb.worksheets[0]
        first_sheet.title = 'sheet1'
        wb.save(excel_path)  # 覆蓋原始 Excel 檔

        # 2. 使用 pandas 讀取剛剛重命名後的 sheet1
        df = pd.read_excel(excel_path, sheet_name='sheet1')

        # 3. 將 DataFrame 輸出成 CSV，保留原始檔名
        csv_path = os.path.join(folder_path, f"{base_name}.csv")
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')

        print(f"已將 '{filename}' 轉換為 '{base_name}.csv' 並將工作表重命名為 'sheet1'。")

print("所有檔案處理完成。")



------------------------------------------------------------------------------------------
# 掃描 C:\_python 2024\2024final 底下所有 .xlsx / .xls 檔

# 使用 openpyxl 將每個檔案的第一個工作表名稱改為 sheet1（小寫）並儲存原檔

# 用 pandas 讀取剛重命名的 sheet1，並輸出同名的 CSV 檔（UTF-8 BOM）

# 原始 Excel 與轉出的 CSV 檔皆保留在同一資料夾