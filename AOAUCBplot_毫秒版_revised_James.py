
# -*- coding: utf-8 -*-
"""
Created on Tue Oct 15 13:10:33 2024

@author: UNEO
"""
import pickle
import numpy as np
import tkinter
from tkinter import filedialog
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime
import os
import pandas as pd
import openpyxl
from pathlib import Path

# ---------- Helpers ----------
def parse_time_any(x):
    """
    將各種可能型態的日期時間(含/不含毫秒、Y/M/D 或 M/D/Y、pandas Timestamp、float 等)
    轉成 Python datetime；失敗則拋出 ValueError。
    """
    if pd.isna(x):
        raise ValueError("NaT/NaN time")

    # 已是 Timestamp 或 datetime
    if isinstance(x, (pd.Timestamp, datetime)):
        return pd.to_datetime(x).to_pydatetime()

    # 可能是 Excel 序號或數字型
    if isinstance(x, (int, float)):
        # 交給 pandas 嘗試自動判斷（原本 Excel 讀進來常見）
        dt = pd.to_datetime(x, errors='coerce', unit='ns', origin='unix')
        if pd.isna(dt):
            # 嘗試當作 Excel 序列日期（1900 基準）
            dt = pd.to_datetime(x, errors='coerce', unit='D', origin='1899-12-30')
        if not pd.isna(dt):
            return dt.to_pydatetime()
        raise ValueError(f"Unrecognized numeric time: {x}")

    # 字串情況
    s = str(x).strip()
    # 常見格式依序嘗試（含毫秒 / 不含毫秒；Y/M/D 與 M/D/Y）
    for fmt in ('%m/%d/%Y %H:%M:%S.%f',
                '%m/%d/%Y %H:%M:%S',
                '%Y/%m/%d %H:%M:%S.%f',
                '%Y/%m/%d %H:%M:%S'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass

    # 最後交給 pandas 的高容錯自動解析
    dt = pd.to_datetime(s, errors='coerce', dayfirst=False)
    if not pd.isna(dt):
        return dt.to_pydatetime()

    raise ValueError(f"Unrecognized time string: {s}")

# ---------- Main ----------
datastr = []
eventls = []
adcsum = []
tmmkls = []

tkinter.Tk().withdraw()
directory = filedialog.askdirectory()
if not directory:
    raise SystemExit("未選擇目錄，程式結束。")

dirpath = Path(directory)
name = dirpath.name

# 只處理 Excel 檔
filelist = sorted([p for p in dirpath.iterdir() if p.suffix.lower() in ('.xlsx', '.xls')])

if not filelist:
    raise SystemExit("目錄內未找到 .xlsx / .xls 檔。")

for idx, xlfile in enumerate(filelist, start=1):
    try:
        # 讀所有工作表為 dict
        xlf = pd.read_excel(xlfile, sheet_name=None)
    except Exception as e:
        print(f"讀取失敗，略過：{xlfile.name}，原因：{e}")
        continue

    if '原始 data' not in xlf:
        print(f"檔案 {xlfile.name} 未找到工作表『原始 data』，略過。")
        continue

    xldata = xlf['原始 data']
    # 確保以連續整數索引
    xldata = xldata.reset_index(drop=True)

    nrow, ncol = xldata.shape
    print(f"Load file:{idx}/{len(filelist)} -> {xlfile.name}")

    for i in range(nrow):
        row = xldata.loc[i]

        # 事件欄（假設在 iloc[2]）
        evt_val = row.iloc[2] if len(row) > 2 else None
        try:
            evt_int = int(evt_val) if not pd.isna(evt_val) else 0
        except Exception:
            evt_int = 0
        eventls.append(evt_int)

        # 原始 hex 資料（假設在 iloc[3]）
        rawhex = row.iloc[3] if len(row) > 3 else None
        rawint = []

        if pd.isna(rawhex):
            # 若沒有資料，補空陣列與 0 和
            datastr.append(rawint)
            adcsum.append(0)
        else:
            # 有些欄位可能被讀成數字或含空白，統一字串化並去空白
            rawhex_str = str(rawhex).strip()
            try:
                rawbyte = bytes.fromhex(rawhex_str)
            except ValueError:
                # 若字串中可能有逗號或中括號等非十六進位字元，嘗試清洗
                cleaned = ''.join(ch for ch in rawhex_str if ch in '0123456789abcdefABCDEF ')
                try:
                    rawbyte = bytes.fromhex(cleaned)
                except Exception:
                    rawbyte = b''

            if len(rawbyte) == 72:
                # 只取前 60 bytes
                for k in range(60):
                    rawint.append(rawbyte[k])
            else:
                for b in rawbyte:
                    rawint.append(b)

            datastr.append(rawint)
            adcsum.append(int(np.array(rawint, dtype=np.int64).sum()) if rawint else 0)

        # 時間欄（假設在 iloc[4]）
        datetn = row.iloc[4] if len(row) > 4 else None
        try:
            time_mk = parse_time_any(datetn)
            tmmkls.append(time_mk)
        except Exception:
            # 若此列時間壞掉，仍推一個 None 以便對齊（也可選擇略過該列）
            # 這裡選擇「略過對應的 event/adc/time」避免繪圖長度不等
            eventls.pop()   # 回退剛 append 的事件
            adcsum.pop()    # 回退剛 append 的和
            # datastr 對繪圖不直接用到，可保留或回退均可；這裡也回退保持對齊
            datastr.pop()
            continue

# --- 繪圖 ---
fig = plt.figure(figsize=(10, 6))
tmformatter = mdates.DateFormatter('%m/%d/%Y %H:%M:%S')

ax1 = fig.add_subplot(211)
ax1.plot(tmmkls, eventls, '--', linewidth=0.5)
ax1.set_title(f"{name}_EVENT")
ax1.xaxis.set_major_formatter(tmformatter)

ax2 = fig.add_subplot(212)
ax2.plot(tmmkls, adcsum, '--', linewidth=0.5)
ax2.set_title(f"{name}_ADC")
ax2.sharex(ax1)
ax2.xaxis.set_major_formatter(tmformatter)

fig.autofmt_xdate()

# 儲存
figfn_pkl = dirpath / f"{name}_ADC.pickle"
figfn_png = dirpath / f"{name}_ADC.png"
pickle.dump(fig, open(figfn_pkl, "wb"))
plt.savefig(figfn_png, dpi=150, bbox_inches='tight')

print(f"已輸出圖檔：{figfn_png}")
print(f"已輸出 pickle：{figfn_pkl}")
