# -*- coding: utf-8 -*-
"""
Created on Wed Jul 14 08:55:33 2021

@author: brendanlia
"""

import numpy
import xlsxwriter
import pandas
import openpyxl
from openpyxl import load_workbook
import time
import win32com.client

wDisp = pandas.read_excel(r'C:\Users\brendanlia\Desktop\Joint Data\Joint_DisplacementNew.xlsx', sheet_name = 'Joint Displacements')
# wDisp = pandas.DataFrame.transpose(wDisp)

StageZeros = []

for a in range(357):
    current = wDisp.loc[(a+1)*224-1]
    current1 = current.values.tolist()
    current1[4] = 0
    current1[6] = current1[6] * -1
    current1[7] = current1[7] * -1
    current1[8] = current1[8] * -1
    current1[9] = current1[9] * -1
    current1[10] = current1[10] * -1
    current1[11] = current1[11] * -1
    StageZeros.append(current1)
    
StageZeros = pandas.DataFrame(StageZeros)
# StageZeros = pandas.DataFrame.transpose(StageZeros)

Disp = wDisp

for b in range(357):
    Row = StageZeros.loc[[b]]
    Row1 = Row.values.tolist()
    Disp.loc[b*224-0.5+b] = Row1[0]
    Disp = Disp.sort_index().reset_index(drop=True)

for c in range(len(Disp)):
    if c%225 == 0:
        U1 = Disp.loc[c , 6]
        U2 = Disp.loc[c , '7']
        U3 = Disp.loc[c , 8]
        R1 = Disp.loc[c , '9']
        R2 = Disp.loc[c , 10]
        R3 = Disp.loc[c , '11']
    else:
        Disp.loc[c , 6] = Disp.loc[c , 6] + U1
        Disp.loc[c , '7'] = Disp.loc[c , '7'] + U2
        Disp.loc[c , 8] = Disp.loc[c , 8] + U3
        Disp.loc[c , '9'] = Disp.loc[c , '9'] + R1
        Disp.loc[c , 10] = Disp.loc[c , 10] + R2
        Disp.loc[c , '11'] = Disp.loc[c , '11'] + R3

Disp.to_excel(r'C:\Users\brendanlia\Desktop\Joint Data\Modified.xlsx', index = False)




