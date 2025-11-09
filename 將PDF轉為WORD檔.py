import os
from pdf2docx import Converter

# PDF 源文件夹
input_dir = r'C:\_python 2024\2024final'
# 输出到子目录 converted
output_dir = os.path.join(input_dir, 'converted')
os.makedirs(output_dir, exist_ok=True)

for filename in os.listdir(input_dir):
    if filename.lower().endswith('.pdf'):
        pdf_path = os.path.join(input_dir, filename)
        docx_name = os.path.splitext(filename)[0] + '.docx'
        docx_path = os.path.join(output_dir, docx_name)

        # 创建 Converter 并执行转换
        cv = Converter(pdf_path)
        cv.convert(docx_path)   # 默认转换全部页面
        cv.close()

        print(f'Converted {filename} → {docx_name}')




# ----------------------------------------------------------------------------------------------------

# 多個PDF檔(其預設路徑放在C:\_python 2024\2024final), 要把PDF檔轉為word檔, 用Python


# --------------------------------------------------------------------------------------------------

# 这样，所有 PDF 就会被转换成 Word 文档并保存在 C:\_python 2024\2024final\converted 目录里。你也可以根据需要调整输出路径或 Converter 的参数（如 start/end 页码范围）。

# ----------------------------------------------------------------------------------------------------------