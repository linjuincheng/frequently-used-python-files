import pandas as pd
from googletrans import Translator
import os

# 設定資料夾路徑
input_folder = r"C:\_python 2024\2024final"

# 找出所有 Excel 檔案
excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.xls', '.xlsx'))]

if not excel_files:
    raise FileNotFoundError("❌ 找不到任何 Excel 檔案，請將檔案放在指定資料夾中。")

# 初始化翻譯器
translator = Translator()

# 逐一處理每個 Excel 檔案
for input_filename in excel_files:
    input_file = os.path.join(input_folder, input_filename)
    base_name, ext = os.path.splitext(input_filename)
    output_filename = f"{base_name}_Translated.xlsx"
    output_file = os.path.join(input_folder, output_filename)

    print(f"🔄 處理中：{input_filename}...")

    try:
        excel_data = pd.read_excel(input_file, sheet_name=None)
        translated_sheets = {}

        for sheet_name, df in excel_data.items():
            translated_df = df.copy()

            # 翻譯欄位名稱
            translated_columns = []
            for col in df.columns:
                if isinstance(col, str):
                    try:
                        translated_text = translator.translate(col, src='zh-tw', dest='en').text
                    except:
                        translated_text = col
                else:
                    translated_text = col
                translated_columns.append(translated_text)
            translated_df.columns = translated_columns

            # 翻譯每欄內容
            for col in translated_df.columns:
                if translated_df[col].dtype == object:
                    translated_df[col] = translated_df[col].apply(
                        lambda x: translator.translate(x, src='zh-tw', dest='en').text
                        if isinstance(x, str) and x.strip() else x
                    )

            # 翻譯工作表名稱
            try:
                translated_sheet_name = translator.translate(sheet_name, src='zh-tw', dest='en').text
            except:
                translated_sheet_name = sheet_name

            translated_sheets[translated_sheet_name[:31]] = translated_df

        # 寫入翻譯後的 Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in translated_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✅ 完成：{output_filename}")

    except Exception as e:
        print(f"❌ 無法處理 {input_filename}：{e}")

print("\n✅ 所有翻譯工作已完成！")
