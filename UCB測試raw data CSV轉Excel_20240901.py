

import os
import pandas as pd

# 設定資料夾路徑
folder_path = r'C:\_python 2024\2024final'

# 遍歷資料夾內所有檔案
for filename in os.listdir(folder_path):
    # 如果檔案是CSV檔
    if filename.endswith('.csv'):
        # 讀取CSV檔案
        csv_path = os.path.join(folder_path, filename)
        df = pd.read_csv(csv_path)
        
        # 設定儲存的Excel檔名（保留原檔名）
        excel_filename = filename.replace('.csv', '.xlsx')
        excel_path = os.path.join(folder_path, excel_filename)
        
        # 將DataFrame寫入Excel，並設定工作表名稱為'sheet1'
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='sheet1')
        
        print(f"已將 {filename} 轉換為 {excel_filename}，並將工作表名稱改為 'sheet1'")

print("所有CSV檔案已成功轉換為Excel格式並設置工作表名稱。")





# 多個CSV檔(其預設路徑放在C:\_python 2024\2024final), 將這些CSV檔轉為EXCEL檔,並把第1個工作表的名稱改為sheet1(英文用小寫), 並保留原檔名, 原始CSV檔及轉出的EXCEL檔都在同一個資料夾 用Python語法





