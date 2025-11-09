from pathlib import Path
import pandas as pd

# 1) 設定根資料夾
root = Path(r"C:\_python 2024\2024final")

# 2) 取得「含子資料夾」的所有檔案
files = [p for p in root.rglob("*") if p.is_file()]

# 3) 依相對路徑（不分大小寫）排序
files.sort(key=lambda p: str(p.relative_to(root)).lower())

# 4) 準備輸出資料（可加一些常用欄位）
rows = []
for i, p in enumerate(files, start=1):
    rel_path = p.relative_to(root).as_posix()       # 例：子夾/檔名.ext
    rows.append({
        "序號": i,
        "相對路徑": rel_path,
        "副檔名": p.suffix,
        "大小(Bytes)": p.stat().st_size,
        # 如需修改時間：p.stat().st_mtime 可再轉人類可讀時間
    })

df = pd.DataFrame(rows, columns=["序號", "相對路徑", "副檔名", "大小(Bytes)"])

# 5) 輸出到 Excel
output_path = root / "檔名清單_含子資料夾.xlsx"
df.to_excel(output_path, index=False, sheet_name="清單")
print(f"已輸出：{output_path}")




# 我想把放在電腦C:\_python 2024\2024final這資料夾裡的所有檔案(含子資料夾)的檔名抓出來,並依序條列在excel工作表裡,Python程式碼?