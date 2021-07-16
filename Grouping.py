# -*- coding: utf-8 -*-
"""
Created on Fri Jul 16 11:57:19 2021

@author: brendanlia
"""

import numpy
import xlsxwriter
import pandas
import openpyxl
from openpyxl import load_workbook
import time
import win32com.client
import sys
import os
import re
import io
import math

def distance(x1,y1,z1,x2,y2,z2):
    return math.sqrt((x1-x2)**2 + (y1-y2)**2 + (z1-z2)**2)


Coords = pandas.read_excel('position.xlsx')
Forces = pandas.read_csv('C:/Users/brendanlia/Desktop/Jacking Tower/CSVs/combined_csv.csv')

Forces.drop_duplicates(subset = ['Node ID','Result Case'], keep = 'first', inplace = True)
Forces = Forces.reset_index(drop = True)

JT = []
Typhoon = []

JTFF = []
TyphoonF = []

for a in range(len(Coords)):
    if len(Coords['Name'][a]) == 17:
        JT.append(Coords.loc[a])
    else: 
        pass
    
for b in range(len(Coords)):
    if len(Coords['Name'][b]) > 17:
        Typhoon.append(Coords.loc[b])
    else: 
        pass
    
for a in range(len(Forces)):
    if len(Forces['Node ID'][a]) == 17:
        JTFF.append(Forces.loc[a])
    else: 
        pass
    
for b in range(len(Forces)):
    if len(Forces['Node ID'][b]) > 17:
        TyphoonF.append(Forces.loc[b])
    else: 
        pass
    
JT = pandas.DataFrame(JT).reset_index(drop = True)
Typhoon = pandas.DataFrame(Typhoon).reset_index(drop = True)

JTForces = pandas.DataFrame(JTFF).reset_index(drop = True)
TyphoonF = pandas.DataFrame(TyphoonF).reset_index(drop = True)

JT['Name'] = JT['Name'].str[:-6]

JTForces['Node ID'] = JTForces['Node ID'].str[:-6]

JTMean = JT.groupby(['Name'], as_index = False).mean()

# Find the Typhoon Restraints that are attatched to each jacking tower
typhoonGroup = []

for a in range(len(Typhoon)):
    x2 = Typhoon['X (m)'][a]
    y2 = Typhoon['Y (m)'][a]
    z2 = Typhoon['Z (m)'][a]
    for b in range(len(JTMean)):
        x1 = JTMean['X (m)'][b]
        y1 = JTMean['Y (m)'][b]
        z1 = JTMean['Z (m)'][b]
        dist = distance(x1,y1,z1,x2,y2,z2)
        if dist < 16:
            typhoonGroup.append(JTMean['Name'][b])
        else:
            pass
        
Typhoon['JTName'] = typhoonGroup        
            
# Replace TyphBrc name with correct JT Group

for a in range(len(Typhoon)):
    # TyphoonF.replace(to_replace = Typhoon['Name'][a], value = Typhoon['JTName'][a],regex = True)
    TyphoonF['Node ID'] = TyphoonF['Node ID'].replace(Typhoon['Name'][a], Typhoon['JTName'][a])

Combined = pandas.concat([TyphoonF, JTForces])

CombinedSumGroup = Combined.groupby(['Node ID','Result Case'],as_index = False).sum()

# Min Max

MaxFx = CombinedSumGroup[CombinedSumGroup.groupby(['Node ID'])['Fx (kN)'].transform(max) == CombinedSumGroup['Fx (kN)']]
MaxFy = CombinedSumGroup[CombinedSumGroup.groupby(['Node ID'])['Fy (kN)'].transform(max) == CombinedSumGroup['Fy (kN)']]
MaxFz = CombinedSumGroup[CombinedSumGroup.groupby(['Node ID'])['Fz (kN)'].transform(max) == CombinedSumGroup['Fz (kN)']]

MinFx = CombinedSumGroup[CombinedSumGroup.groupby(['Node ID'])['Fx (kN)'].transform(min) == CombinedSumGroup['Fx (kN)']]
MinFy = CombinedSumGroup[CombinedSumGroup.groupby(['Node ID'])['Fy (kN)'].transform(min) == CombinedSumGroup['Fy (kN)']]
MinFz = CombinedSumGroup[CombinedSumGroup.groupby(['Node ID'])['Fz (kN)'].transform(min) == CombinedSumGroup['Fz (kN)']]

MaxFx = MaxFx.drop(['Mx (kNm)', 'My (kNm)', 'Mz (kNm)'], axis = 1)
MaxFy = MaxFy.drop(['Mx (kNm)', 'My (kNm)', 'Mz (kNm)'], axis = 1)
MaxFz = MaxFz.drop(['Mx (kNm)', 'My (kNm)', 'Mz (kNm)'], axis = 1)

MinFx = MinFx.drop(['Mx (kNm)', 'My (kNm)', 'Mz (kNm)'], axis = 1)
MinFy = MinFy.drop(['Mx (kNm)', 'My (kNm)', 'Mz (kNm)'], axis = 1)
MinFz = MinFz.drop(['Mx (kNm)', 'My (kNm)', 'Mz (kNm)'], axis = 1)

MaxFx.drop_duplicates(subset = ['Node ID'], keep = 'first', inplace = True)
MaxFy.drop_duplicates(subset = ['Node ID'], keep = 'first', inplace = True)
MaxFz.drop_duplicates(subset = ['Node ID'], keep = 'first', inplace = True)

MinFx.drop_duplicates(subset = ['Node ID'], keep = 'first', inplace = True)
MinFy.drop_duplicates(subset = ['Node ID'], keep = 'first', inplace = True)
MinFz.drop_duplicates(subset = ['Node ID'], keep = 'first', inplace = True)

for a in range(len(JTMean)):
    data = [MaxFx.iloc[a],MaxFy.iloc[a],MaxFz.iloc[a],MinFx.iloc[a],MinFy.iloc[a],MinFz.iloc[a]]
    data1 = pandas.DataFrame(data)
    globals()['%s' %JTMean['Name'][a]] = pandas.DataFrame(data1)
    

writer = pandas.ExcelWriter('JackingTowerExport.xlsx', engine='xlsxwriter')  


for a in range(len(JTMean)):
    globals()['%s' %JTMean['Name'][a]].to_excel(writer, sheet_name = '%s' %JTMean['Name'][a])
    
writer.save()















