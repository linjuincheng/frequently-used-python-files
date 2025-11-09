import os
import pandas as pd

# 設定資料夾路徑
folder_path = r'C:\_python 2024\2024final'

# 逐一處理資料夾中的 Excel 檔案
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(folder_path, filename)
        
        # 讀取 Excel 檔案的 Sheet1
        df = pd.read_excel(file_path, sheet_name='sheet1')
        
        # 確保 D 欄和 E 欄為日期時間格式
        df['D'] = pd.to_datetime(df['D'], errors='coerce')
        df['E'] = pd.to_datetime(df['E'], errors='coerce')
        
        # 計算 E 欄與 D 欄的秒數差，並將結果存入 F 欄
        df['F'] = (df['E'] - df['D']).dt.total_seconds()
        
        # 將更新後的資料儲存回原 Excel 檔案的 Sheet1
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)

        print(f"已處理並更新檔案: {filename}")




# Raw data與封包數比對(資料起始時間總秒數)_20241112
# 1. 多個excel檔案(其預設路徑放在C:\_python 2024\2024final), 這些excel檔的sheet1工作表的D欄儲存格資料為開始時間, E欄儲存格資料為結束時間
#  2. 計算E欄儲存格與D欄儲存格的時間資料的秒數差,並把秒數寫在F欄儲存格 
# 3. 用Python語法

# -----------------------------------------------------------------------------------------------------------------------------------------------
# 說明
# 讀取 Excel 檔案：遍歷指定資料夾中的每個 .xlsx 檔案。
# 轉換日期時間格式：將 D 和 E 欄轉為日期時間格式，便於時間計算。
# 計算秒數差：使用 (df['E'] - df['D']).dt.total_seconds() 計算每一列中 E 欄與 D 欄的秒數差，結果寫入 F 欄。
# 更新 Excel：將結果回寫至原 Excel 檔案的 Sheet1 工作表。
# 注意
# 此程式假設所有檔案的 D、E 欄均為有效的日期時間資料。

# -------------------------------------------------------------------------------------------------------------------------------------------------



