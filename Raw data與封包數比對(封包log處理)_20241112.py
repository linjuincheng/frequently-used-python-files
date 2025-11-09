import os
import pandas as pd
from datetime import datetime, timedelta

# 設定資料夾路徑
folder_path = r'C:\_python 2024\2024final'

# 遍歷資料夾中所有 .log 檔案
for filename in os.listdir(folder_path):
    if filename.endswith('.log'):
        file_path = os.path.join(folder_path, filename)
        
        # 讀取 .log 檔案，假設逗號作為分隔符
        df = pd.read_csv(file_path, header=None)
        
        # 增加一列 (第 1 欄及第 3 欄間新增 1 欄)
        df.insert(1, 'New Column', '')
        
        # 新增標題列
        df.columns = ['Log time', 'Raw Data Time', 'IP', 'Packets']
        
        # 將時間往前回推 15 分鐘並寫入到 "Raw Data Time" 欄位
        df['Raw Data Time'] = pd.to_datetime(df['Log time']) - timedelta(minutes=15)
        
        # 複製最後一筆時間到右邊的儲存格
        last_time = df['Log time'].iloc[-1]
        df.at[len(df) - 1, 'Raw Data Time'] = last_time
        
        # 將 .log 檔名轉為 .xlsx
        excel_filename = filename.replace('.log', '.xlsx')
        excel_path = os.path.join(folder_path, excel_filename)
        
        # 輸出成 Excel 檔案
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)




# Raw data與封包數比對(封包log處理)_20241112
# 1.	多個副名為.log檔案(其預設路徑放在C:\_python 2024\2024final),其資料欄位用逗號,作欄位區分,將這些檔案轉為EXCEL檔 , 並保留原檔名, 原始.log檔及轉出的EXCEL檔都在同一個資料夾
# 2.	在每個Excel檔的最上方新增一列,第1欄及第3欄間新增1欄
# 3.	在A1儲存格填入Log time, B1儲存格填入對應raw data始末時間,C1 儲存格填入IP,D1儲存格填入Packets
# 4.	A2儲存格內的資料為時間, 把其時間往前回推15分鐘的時間放到B2儲存格(e.g. A2儲存格裡的時間為2024-11-08 00:03:58, 則B2儲存格就放2024-11-07 23:48:58)
# 5.	將A欄位底下的最後一個裡面有資料的儲存格,其內容複製一份,放到其右邊的欄位(e.g. A41儲存格裡是2024-11-08 09:48:58, 則把2024-11-08 09:48:58複製到B41儲存格裡)
# 用Python語法

# ------------------------------------------------------------------------------------------------------------------------------------------------------------程式碼說明
# folder_path：設定您 .log 檔案的資料夾路徑。
# 遍歷 .log 檔案，將其讀取為 pandas.DataFrame。
# 在第 1 欄和第 3 欄之間新增一個空白欄位。
# 設定標題列，並按需求修改欄位名稱。
# 計算每一筆 Log time 欄位時間減去 15 分鐘，並填入 Raw Data Time 欄位。
# 複製最後一筆的 Log time 資料到相鄰的 Raw Data Time 欄位。
# 將每個 .log 檔案轉存為相同檔名的 .xlsx 檔案。

