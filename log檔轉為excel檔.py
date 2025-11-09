
import os
import pandas as pd

# 定義檔案目錄
folder_path = r"C:\_python 2024\2024final"

# 獲取所有副名為 .log 的檔案
log_files = [f for f in os.listdir(folder_path) if f.endswith('.log')]

# 逐個處理每個 .log 檔案
for log_file in log_files:
    # 讀取 .log 檔案
    file_path = os.path.join(folder_path, log_file)
    data = pd.read_csv(file_path, delimiter=",")  # 使用逗號作為分隔符號
    
    # 將檔案儲存為 Excel 檔案
    excel_file_path = os.path.join(folder_path, log_file.replace('.log', '.xlsx'))
    data.to_excel(excel_file_path, index=False)  # 保留原始檔名並移除索引
    
    
    
    
#     多個副名為.log檔案(其預設路徑放在C:\_python 2024\2024final),其資料欄位用逗號,作欄位區分,將這些檔案轉為EXCEL檔 , 並保留原檔名, 原始.log檔及轉出的EXCEL檔都在同一個資料夾 用Python語法

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# 這段程式碼會執行以下操作：

# 讀取指定資料夾內所有 .log 檔案。
# 將每個 .log 檔案的內容讀取進 pandas 的 DataFrame。
# 將每個 DataFrame 轉換並儲存成 Excel 格式，保留原始檔名，並放回同一資料夾中
