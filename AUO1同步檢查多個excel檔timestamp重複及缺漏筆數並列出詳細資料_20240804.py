# 程式碼說明
# 導入必要的庫：

# os 和 glob 用於文件和目錄操作。
# pandas 用於數據處理和 DataFrame 操作。
# 定義 process_excel 函數：

# 讀取 Excel 檔案並確保時間欄位是日期時間格式。
# 計算總筆數、時間重複的筆數和時間有無缺漏。
# 將計算結果寫入同一個 Excel 檔案的不同工作表中，並覆蓋原來的工作表。
# 定義 main 函數：

# 設定需要處理的目錄並使用 glob 取得目錄下所有的 Excel 檔案。
# 遍歷所有 Excel 檔案並調用 process_excel 函數進行處理。
# 執行 main 函數：

# 確保腳本在執行時會處理目錄中的所有 Excel 檔案。
# 注意事項
# 備份原始檔案：在覆蓋原始檔案之前，建議先備份原始檔案，以防出現意外情況。
# 安裝庫：確保已安裝 pandas 和 openpyxl 庫。如果尚未安裝，可以使用 pip install pandas openpyxl 安裝。


import os
import glob
import pandas as pd

def process_excel(file_path):
    # 讀取Excel檔案，假設第5個欄位的名稱是'Time'
    df = pd.read_excel(file_path)

    # 確保時間欄位是日期時間格式
    df['Time'] = pd.to_datetime(df.iloc[:, 4], format='%m/%d/%Y %H:%M:%S')

    # 總筆數
    total_records = len(df)

    # 查找時間重複的筆數
    duplicates = df[df.duplicated(subset=['Time'], keep=False)]
    duplicate_count = len(duplicates)

    # 檢查時間有無缺漏
    df = df.sort_values(by='Time').reset_index(drop=True)
    time_diffs = df['Time'].diff().dt.total_seconds().fillna(1)  # 第1筆資料的差異值設為1秒

    # 計算缺漏時間段並生成缺漏時間的DataFrame
    missing_times_list = []
    for i in range(1, len(df)):
        if time_diffs[i] > 1:
            start_time = df['Time'][i-1]
            end_time = df['Time'][i]
            missing_times_list.append([start_time, end_time])

    missing_times = pd.DataFrame(missing_times_list, columns=['Start Time', 'End Time'])
    missing_count = len(missing_times)

    print(f"檔案: {os.path.basename(file_path)}")
    print(f"總共有 {total_records} 筆資料")
    print(f"時間重複的筆數: {duplicate_count} 筆")
    print(f"時間有缺漏的筆數: {missing_count} 筆")

    # 將結果寫入同一個Excel檔案的不同工作表
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        duplicates.to_excel(writer, sheet_name='重複筆數', index=False)
        missing_times.to_excel(writer, sheet_name='缺漏筆數', index=False)

def main():
    # 設定需要處理的目錄
    input_directory = 'C:\\_python 2024\\2024final'
    # 取得目錄下所有的Excel檔案
    excel_files = glob.glob(os.path.join(input_directory, '*.xlsx'))

    # 遍歷所有Excel檔案並處理
    for file_path in excel_files:
        process_excel(file_path)

if __name__ == "__main__":
    main()


