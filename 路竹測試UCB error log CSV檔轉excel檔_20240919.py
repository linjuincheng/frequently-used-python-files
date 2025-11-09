import os
import pandas as pd
from openpyxl import load_workbook

# 定義 CSV 檔案的路徑
csv_folder = r"C:\_python 2024\2024final"

# 遍歷資料夾中的所有 CSV 檔案
for filename in os.listdir(csv_folder):
    if filename.endswith('.csv'):
        # 獲取完整的檔案路徑
        csv_file_path = os.path.join(csv_folder, filename)
        
        try:
            # 嘗試使用不同的編碼讀取 CSV 檔案，並將所有列都作為字串來保留格式
            df = pd.read_csv(csv_file_path, encoding='utf-8', dtype=str)  # 先嘗試使用 UTF-8
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(csv_file_path, encoding='big5', dtype=str)  # 再嘗試使用 Big5
            except UnicodeDecodeError:
                df = pd.read_csv(csv_file_path, encoding='ISO-8859-1', dtype=str)  # 最後使用 ISO-8859-1

        # 將 CSV 轉換為 Excel 檔案（保留原始檔名）
        excel_file_path = os.path.splitext(csv_file_path)[0] + '.xlsx'
        
        # 使用 Pandas 將資料寫入 Excel，保留字串格式
        df.to_excel(excel_file_path, index=False, engine='openpyxl')
        
        # 使用 openpyxl 修改工作表名稱為 'sheet1'
        workbook = load_workbook(excel_file_path)
        sheet = workbook.active
        sheet.title = 'sheet1'
        workbook.save(excel_file_path)
        
        print(f"{filename} 已成功轉換為 {excel_file_path} 並修改工作表名稱為 sheet1")



# 多個CSV檔(其預設路徑放在C:\_python 2024\2024final), 將這些CSV檔轉為EXCEL檔,並把第1個工作表的名稱改為sheet1(英文用小寫), 並保留原檔名, 原始CSV檔及轉出的EXCEL檔都在同一個資料夾 用Python語法

# 說明：
# dtype=str：這個參數會將所有資料都以字串的形式讀取，即使是時間或數字型資料，也能確保原始格式不變。
# engine='openpyxl'：這是指定在寫入 Excel 時使用的引擎，因為我們後面使用 openpyxl 來修改 Excel 檔案的工作表名稱。
# df.to_excel：保留原始的資料內容格式，不會進行自動轉換。
# 這樣的處理方式能保證所有 CSV 檔中的資料格式在轉換成 Excel 檔案時不會丟失，特別是時間格式的資料。