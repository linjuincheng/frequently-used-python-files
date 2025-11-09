
import os
import pandas as pd
from openpyxl import load_workbook

# 設定資料夾路徑
folder_path = r'C:\_python 2024\2024final'
summary_data = []  # 儲存 summary 表格的資料

# 逐一處理資料夾中的 .csv 檔案
for filename in os.listdir(folder_path):
    if filename.endswith('.csv'):
        file_path = os.path.join(folder_path, filename)
        
        # 讀取 .csv 檔案
        df = pd.read_csv(file_path)
        
        # 計算 A 欄位的資料筆數 (包括標題列)
        data_count = len(df)
        
        # 使用 A2 儲存格內容作為工作表名稱
        sheet_name = str(df.iloc[0, 0]) if len(df) > 0 else "Sheet1"
        
        # 將 .csv 檔案轉為 .xlsx 並保留原檔名
        excel_filename = filename.replace('.csv', '.xlsx')
        excel_path = os.path.join(folder_path, excel_filename)
        
        # 寫入 Excel 檔，並命名 Sheet1 為 A2 儲存格的內容
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # 記錄 summary 資訊
        summary_data.append([sheet_name, data_count])

# 建立 "raw data資料筆數統計" 的 Excel 檔案
combined_excel_path = os.path.join(folder_path, 'raw data資料筆數統計.xlsx')
with pd.ExcelWriter(combined_excel_path, engine='openpyxl') as writer:
    # 在所有工作表的最前面新增 summary 表格
    summary_df = pd.DataFrame(summary_data, columns=['deviceid', '資料筆數'])
    summary_df.to_excel(writer, sheet_name='summary', index=False)
    
    # 將每個 Excel 檔案中的工作表複製到主檔案
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx') and filename != 'raw data資料筆數統計.xlsx':
            file_path = os.path.join(folder_path, filename)
            wb = load_workbook(file_path)
            sheet = wb.active
            
            # 將工作表的資料讀入 DataFrame 並寫入主檔案
            sheet_df = pd.DataFrame(sheet.values)
            sheet_df.columns = sheet_df.iloc[0]  # 設定標題列
            sheet_df = sheet_df[1:]  # 移除標題列
            sheet_df.to_excel(writer, sheet_name=sheet.title, index=False)





# 1.	多個副名為.csv檔案(其預設路徑放在C:\_python 2024\2024final),將這些檔案轉為EXCEL檔 , 並保留原檔名, 原始.csv檔及轉出的EXCEL檔都在同一個資料夾
# 2.	統計每個檔案的Sheet1工作表的A欄位總共有多少列資料, 得出的數字,即為該Sheet1工作表的資料筆數
# 3.	將每個檔案的Sheet1工作表名稱給為其A2儲存格的內容
# 4.	新增一個名稱為raw data資料筆數統計的excel檔,並把C:\_python 2024\2024final裡的所有excel檔裡個工作表,複製到raw data資料筆數統計這個excel檔
# 5.	raw data資料筆數統計的excel檔的工作表的最前面新增一個名稱為summary的工作表, 第一個欄位為deviceid, 第二個欄位為資料筆數, 然後把每個工作表有多少資料筆數列在summary裡
# 6.	用Python語法

# ------------------------------------------------------------------------------------------------------------------------------------------------------------
# 程式碼說明
# 將 .csv 檔案轉為 Excel：程式會讀取資料夾中的 .csv 檔案，將其轉換為 .xlsx 檔案並保留原檔名。
# 計算資料筆數：計算每個檔案的 A 欄位資料列數（包括標題），並將其記錄在 summary_data 中。
# 設定工作表名稱：每個 Excel 檔案的 Sheet1 工作表名稱設置為 A2 儲存格的內容。
# 創建主檔案：創建一個名為 "raw data資料筆數統計.xlsx" 的 Excel 檔案。
# 新增 Summary 表：在主檔案的最前方添加名為 summary 的工作表，並將 deviceid 和 資料筆數 記錄到此表中。
# 合併工作表：將每個 Excel 檔案中的工作表複製到主檔案中。            
            
            
            
            
