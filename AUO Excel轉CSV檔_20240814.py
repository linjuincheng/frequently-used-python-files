import pandas as pd
import os
from datetime import datetime

# 設定檔案路徑
folder_path = 'C:\\_python 2024\\2024final'  # 設定資料夾路徑

# 確保資料夾路徑存在
if not os.path.exists(folder_path):
    raise FileNotFoundError(f"資料夾 {folder_path} 不存在")

# 處理資料夾中的所有 Excel 檔案
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        # 讀取 Excel 檔案
        file_path = os.path.join(folder_path, file_name)
        excel_data = pd.read_excel(file_path, sheet_name='sheet1')

        # 確保有足夠的行數
        if len(excel_data) < 2:
            print(f"檔案 {file_name} 的資料行數不足，無法處理")
            continue

        # 提取第二列的 MAC 和 Time 欄位的資料
        mac_value = excel_data.iloc[1]['MAC']
        time_value = excel_data.iloc[1]['Time']

        # 轉換 Time 欄位資料到所需格式
        try:
            time_str = datetime.strptime(time_value, '%m/%d/%Y %H:%M:%S').strftime('%Y%m%d')
        except ValueError:
            print(f"Time 格式錯誤: {time_value} 在檔案 {file_name} 中")
            continue

        # 創建 CSV 檔案名稱
        csv_file_name = f"AUO_{mac_value}_TW_{time_str}.csv"
        csv_file_path = os.path.join(folder_path, csv_file_name)

        # 將 DataFrame 轉換為 CSV 格式，分隔符號為分號
        excel_data.to_csv(csv_file_path, sep=';', index=False)

print("所有檔案處理完成。")




# 數個excel檔,excel的資料欄位有4欄,要把4欄用分號;作欄位區隔並轉成CSV檔, 檔名依sheet1工作表的第一列的MAC及Time欄位裡的日期作相對應的變化, 檔名格式AUO_MAC_TW_Time 
# e.g.如果MAC欄位是的第二列是D83ADD861851, 而Time(format='%m/%d/%Y %H:%M:%S')欄位的第二列是08/06/2024 10:03:02 
# 則CSV檔名輸出檔名為AUO_D83ADD861851_TW_20240806 , 原始excel檔及轉出的CSV檔都在同一個資料夾
# 用Python語法



# 程式碼說明：
# 設定檔案路徑：

# folder_path：資料夾的路徑，包含 Excel 檔案和輸出 CSV 檔案。
# 處理 Excel 檔案：

# 讀取資料夾中的每個 Excel 檔案。
# 讀取 Sheet1 工作表中的資料。
# 提取 MAC 和 Time 欄位的第二列資料。
# 格式化 Time 欄位的日期成為 YYYYMMDD 格式。
# 生成 CSV 檔案：

# 根據格式生成 CSV 檔名。
# 將資料儲存為 CSV 檔案，使用分號 ; 作為分隔符。
# 錯誤處理：

# 如果 Excel 檔案的資料行數不足或日期格式錯誤，會顯示錯誤訊息並跳過該檔案