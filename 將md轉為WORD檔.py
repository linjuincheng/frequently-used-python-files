import os
import subprocess

input_dir = r'C:\_python 2024\2024final'
output_dir = os.path.join(input_dir, 'converted')
os.makedirs(output_dir, exist_ok=True)

for fname in os.listdir(input_dir):
    if fname.lower().endswith('.md'):
        md_path = os.path.join(input_dir, fname)
        docx_path = os.path.join(output_dir, os.path.splitext(fname)[0] + '.docx')
        subprocess.run(['pandoc', md_path, '-o', docx_path], check=True)
        print(f'Converted {fname} → {os.path.basename(docx_path)}')
        
        


# -----------------------------------------------------------------------------------------------------

# 多個.md檔(其預設路徑放在C:\_python 2024\2024final), 要把.md檔轉為word檔, 用Python


# --------------------------------------------------------------------------------------------------------
