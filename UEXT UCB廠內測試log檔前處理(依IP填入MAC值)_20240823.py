

import os
import pandas as pd

# 設定 Excel 文件所在的文件夾
folder_path = r'C:\_python 2024\2024final'

# IP 地址到對應值的映射字典
ip_to_value = {
    '192.9.120.75': '80C9555CD1E8',
    '192.9.120.89': '80C9555CBF6C',
    '192.9.120.79': '80C9555CB2FC',
    '192.9.120.73': '80C9555CC648',
    '192.9.131.45': '80C9555C3064',
    '192.9.120.91': '80C9555CAFB0',
    '192.9.120.88': '80C9555CC470',
    '192.9.120.90': '80C955431230',
    '192.9.131.42': '80C9555CC90C',
    '192.9.120.84': '80C9555CC3CC',
    '192.9.120.92': '80C955438E08',
    '192.9.120.86': '80C9555CB064',
    '192.9.120.93': '80C955438D5C',
    '192.9.120.80': '80C9555CCE40',
    '192.9.120.81': '80C955438D90',
    '192.9.120.82': '80C9555CB040',
    '192.9.120.87': '80C9555CB644',
    '192.9.120.90': '80C955439230',
    '192.9.120.160': '80C955438C4C',
    '192.9.120.74': '80C9555CB92C',
    '192.9.120.83': '80C9555CB8E8',
    '192.9.120.85': '80C9555CD15C',
    '192.9.120.165': '80C9555CB020',
    '192.9.120.135': '80C9555CC1AC',
    '192.9.120.139': '48E7DA58152F',
    '192.9.120.208': 'D8126542E637',
    '192.9.100.107': '48A472FBD78B',
    '192.9.120.61': '0050568AFCB5',
    '192.9.120.153': '089DF42D1036',
    '192.9.120.211': '0A57D83BDFEE',
    '192.9.120.213': '0C9A3C955673'
}

# 迭代文件夾中的每個 Excel 文件
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        file_path = os.path.join(folder_path, filename)
        
        # 讀取 Excel 文件的 sheet1
        df = pd.read_excel(file_path, sheet_name='sheet1')
        
        # 列印 DataFrame 的列名
        print(f"Columns in {filename}: {df.columns.tolist()}")

        # 根據實際列名設定變數
        ip_col = df.columns[0]  # 第 1 欄的列名
        value_col = df.columns[1]  # 第 2 欄的列名
        
        # 更新第 2 欄的值
        df[value_col] = df[ip_col].map(ip_to_value).fillna(df[value_col])
        
        # 保存修改後的 Excel 文件
        df.to_excel(file_path, sheet_name='sheet1', index=False)
        
        print(f'Updated file: {filename}')

