# -*- coding: utf-8 -*-
"""
Created on Wed Aug 14 14:48:25 2024

@author: UNEO
"""
import numpy as np
import tkinter
from tkinter import filedialog
import csv
import matplotlib.pyplot as plt
import matplotlib.dates as dates
from datetime import datetime

datastr=[]
eventls=[]
imagedata=[]
prmtdata=[]
tmmkls=[]
tkinter.Tk().withdraw()
file_path=filedialog.askopenfilename()
namesp=file_path.split('_')
with open(file_path) as File: 
    Line_reader = csv.reader(File,delimiter=";")
    for row in Line_reader:
        if not row[0]=="MAC":
            datastr.append(row)
            row1splt=row[1]
            if(row1splt=='Resting on the bed'):
                eventls.append(1)
            else:
                eventls.append(4)
            row2splt=row[2]
            if(row2splt):
                strg=np.array(row2splt.split(',')).astype(np.int32)
                imagedata.append(strg[0:60])
                prmtdata.append(strg[60:])
            row3splt=row[3]
            datetnstr=row3splt
            # datetimestr=datestr[0]+" "+row3splt[2]
            time_mk=datetime.strptime(datetnstr,'%m/%d/%Y %H:%M:%S')
            tmmkls.append(time_mk)
        
oldprmt=np.zeros([len(prmtdata),6])
cnt=0
for cnd in prmtdata:
    for j in range(1,7):
        oldprmt[cnt,j-1]=cnd[(j-1)*2+1]*256+cnd[(j-1)*2]
    cnt +=1

fig = plt.figure()
ax = fig.add_subplot(211)
tmformatter=dates.DateFormatter('%m/%d/%Y %H:%M:%S')
# for i, axes in enumerate(ax.flatten()):
plt.title(namesp[1]+'_'+namesp[2])
if namesp[1]=="706":
    PMIO=200
else:
    PMIO=50
plt.plot(tmmkls,eventls,'r--',linewidth=0.5)
#plt.plot(tmmkls,eventckls,'g--',linewidth=0.5)
ax.set_title("Event")
ax.xaxis.set_major_formatter(tmformatter)
ax2=fig.add_subplot(212)
plt.plot(tmmkls,oldprmt[:,1]+PMIO,'g--',linewidth=0.5)
plt.plot(tmmkls,oldprmt[:,2],'b--',linewidth=0.5)
ax2.set_title("ADC")
ax2.xaxis.set_major_formatter(tmformatter)
plt.show()