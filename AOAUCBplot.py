# -*- coding: utf-8 -*-
"""
Created on Tue Oct 15 13:10:33 2024

@author: UNEO
"""
import pickle
import numpy as np
import tkinter
from tkinter import filedialog
import csv
import matplotlib.pyplot as plt
import matplotlib.dates as dates
from datetime import datetime
import os
import pandas as pd
import openpyxl

datastr=[]
eventls=[]
imagedata=[]
adcsum=[]
tmmkls=[]
tkinter.Tk().withdraw()
directory=filedialog.askdirectory()
dirsplt=directory.split('/')
name=dirsplt[-1]
filelist=os.listdir(directory)
for j in range(0,len(filelist)):
    xlf=pd.read_excel(directory+"/"+filelist[j], sheet_name = None)
    xls=pd.ExcelFile(directory+"/"+filelist[j])
    xldata=xlf.get('原始 data')
    nrow,ncol=xldata.shape
    print("Load file:"+str(j)+"/"+str(len(filelist))+"->"+filelist[j])
    for i in range(0,nrow):
        row=xldata.loc[i]
        row1splt=row.iloc[2]
        eventls.append(int(row1splt))
        row2splt=row.iloc[3]
        rawint=[]
        rawbyte=bytes.fromhex(row2splt)
        if len(rawbyte)==72:
            for j in range(0,60):
                rawint.append(rawbyte[j])
            datastr.append(rawint)
            adcsum.append(np.array(rawint).sum())
        else:
            for data in rawbyte:
                rawint.append(data)
            datastr.append(rawint)
            adcsum.append(np.array(rawint).sum())
        datetnstr=row.iloc[4]
        # datetimestr=datestr[0]+" "+row3splt[2]
        time_mk=datetime.strptime(datetnstr,'%m/%d/%Y %H:%M:%S')
        tmmkls.append(time_mk)

# oldprmt=np.zeros([len(prmtdata),6])
# cnt=0
# for cnd in prmtdata:
#     for j in range(1,7):
#         oldprmt[cnt,j-1]=cnd[(j-1)*2+1]*256+cnd[(j-1)*2]
#     cnt +=1

fig=plt.figure()
tmformatter=dates.DateFormatter('%m/%d/%Y %H:%M:%S')
ax1 = fig.add_subplot(211)
ax1.plot(tmmkls,eventls,'r--',linewidth=0.5)
ax1.set_title(name+"_EVENT")
ax1.xaxis.set_major_formatter(tmformatter)
ax2 = fig.add_subplot(212)
ax2.plot(tmmkls,adcsum,'b--',linewidth=0.5)
ax2.set_title(name+"_ADC")
ax2.sharex(ax1)
ax2.xaxis.set_major_formatter(tmformatter)
figfn=name+"ADC.pickle"
pickle.dump(fig, open(figfn, "wb"))
#xlf=openpyxl.load_workbook(directory+"/"+filelist[1])
#datasheet=xlf['原始 data']
# for row in datasheet:
#     if not row[0].value=='Data select ':
#         datastr.append(row[2].value+row[3].value+row[4].value)
#         row1splt=row[2].value
#         if(row1splt=='Resting on the bed'):
#             eventls.append(1)
#         else:
#             eventls.append(4)
#         row2splt=row[3].value
#         if(row2splt):
#             strg=np.array(row2splt.split(',')).astype(np.uint8)
#             imagedata.append(strg[0:60])
#             prmtdata.append(strg[60:])
#         row3splt=row[4]
#         datetnstr=row3splt.value
#         # datetimestr=datestr[0]+" "+row3splt[2]
#         time_mk=datetime.strptime(datetnstr,'%m/%d/%Y %H:%M:%S')
#         tmmkls.append(time_mk)
#file_path=filedialog.askopenfilename()