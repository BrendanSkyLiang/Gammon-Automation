# -*- coding: utf-8 -*-
"""
Created on Wed Jul 14 08:55:33 2021

@author: brendanlia
"""

# import numpy
# import xlsxwriter
import pandas
# import openpyxl
# from openpyxl import load_workbook
# import time
# import win32com.client
# import sys
# import os
import re
import io

'----------------------------------------------------------------------------------------------------------------------------------------'
# Convert .s2k file to something useful. Aka, DataFrame

ShittyFile = open('ExportCompare.s2k') # <--- Needs Modification
LessShittyFile = ShittyFile.read()
TrimmedText = LessShittyFile.split('\n',6)[-1]

x = re.sub("   ",", ",TrimmedText.strip())
x = re.sub(" _",", ",x.strip())
x = re.sub(", ,   ","",x.strip())
x = re.sub("\n","",x.strip())
x = re.sub(", Joint","\nJoint",x.strip())
x = re.sub('TABLE:  "JOINT DISPLACEMENTS"','"Joint","OutputCase","CaseType","StepType","StepNum","StepLabel","U1","U2","U3","R1","R2","R3"\n',x.strip())
x = re.sub("Joint=","",x.strip())
x = re.sub("OutputCase=","",x.strip())
x = re.sub("CaseType=","",x.strip())
x = re.sub("StepType=","",x.strip())
x = re.sub("StepNum=","",x.strip())
x = re.sub("StepLabel=","",x.strip())
x = re.sub("U1=","",x.strip())
x = re.sub("U2=","",x.strip())
x = re.sub("U3=","",x.strip())
x = re.sub("R1=","",x.strip())
x = re.sub("R2=","",x.strip())
x = re.sub("R3=","",x.strip())
x = re.sub(" END TABLE DATA","",x.strip())

x = io.StringIO(x)
df = pandas.read_csv(x, dtype = {'Joint': str, 'OutputCase': str, 'CaseType': str, 'StepType': str,
                                'StepNum': int, 'StepLabel': str,
                                 'U1': float, 'U2': float, 'U3': float, 'R1': float, 'R2': float, 'R3': float
                                })

# df.to_excel('pandas_simple.xlsx', sheet_name='Sheet1',index = False)

'----------------------------------------------------------------------------------------------------------------------------------------'

# wDisp = pandas.read_excel('pandas_simple.xlsx', sheet_name = 'Sheet1')

# wDisp = pandas.DataFrame.transpose(wDisp)

wDisp = df

count = wDisp.nunique()[0]

StageZeros = []

for a in range(count):
    try:
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
        
    except:
        pass
    
    
StageZeros = pandas.DataFrame(StageZeros, columns=["Joint","OutputCase","CaseType","StepType","StepNum","StepLabel","U1_Updated","U2_Updated","U3_Updated","R1_Updated","R2_Updated","R3_Updated"])

StageZeros_expanded = StageZeros.loc[StageZeros.index.repeat(224)].reset_index(drop=True)

wDisp["U1_Updated"] = wDisp['U1'] + StageZeros_expanded["U1_Updated"]
wDisp["U2_Updated"] = wDisp['U2'] + StageZeros_expanded["U2_Updated"]
wDisp["U3_Updated"] = wDisp['U3'] + StageZeros_expanded["U3_Updated"]
wDisp["R1_Updated"] = wDisp['R1'] + StageZeros_expanded["R1_Updated"]
wDisp["R2_Updated"] = wDisp['R2'] + StageZeros_expanded["R2_Updated"]
wDisp["R3_Updated"] = wDisp['R3'] + StageZeros_expanded["R3_Updated"]
Disp = wDisp.drop(columns = {'U1', 'U2', 'U3', 'R1', 'R2', 'R3'})
Disp = pandas.concat([Disp, StageZeros])

sortedDisp = Disp.sort_values(by=['Joint', 'StepNum']).reset_index(drop=True)

'---------------------------------------------------------------------------------------------------------------------------------'

# Define the points that needs to be sorted out
Joints = [10075, 'TG_PrInf994', 5] # <---- Input Parameter

def filterPoint(list):
    string_list = []
    for i in list:
        string_list.append(str(i))
    a = sortedDisp[sortedDisp['Joint'].isin(string_list)]
    a.to_csv('modified_csv.csv', index = False) # <---- needs to change the name and match with the file name below

filterPoint(Joints)

'---------------------------------------------------------------------------------------------------------------------------------'
#Find overlapping points

jointNum = Joints
stepNum = 1  #input('Step Number: ')
#stepNum = str(stepNum)

FCoord = pandas.read_excel(r'C:\Users\brendanlia\Desktop\Joint Data\Joint_Coordinates.xlsx', sheet_name = 'Joint Coordinates') # <--- Needs Modification
FCoord = pandas.DataFrame(FCoord)
FDisp = pandas.read_csv('modified_csv.csv') # <---- Needs Modification, match with filterPoints name
FDisp = pandas.DataFrame(FDisp)

Coord = pandas.DataFrame.transpose(FCoord)
Disp = pandas.DataFrame.transpose(FDisp)


CoordExport = []
DispExport = []

for a in range(len(jointNum)):
    
    jointNumStr = str(jointNum[a])
    
    FCoordIndex = (FCoord[FCoord['TABLE:  Joint Coordinates'] == jointNumStr].index.values)
    FDispIndex1 = (FDisp[FDisp['Joint'] == jointNumStr].index.values)
    FDispIndex2 = (FDisp[FDisp['StepNum'] == stepNum].index.values)
    DispIndex = list(set(FDispIndex1) & set(FDispIndex2))

    CoordAmend = pandas.DataFrame.transpose(Coord[FCoordIndex][7:10])
    CoordAmend1 = CoordAmend.values.tolist()
    CoordExport.append(CoordAmend1[0])
    
    DispAmend = pandas.DataFrame.transpose(Disp[DispIndex][6:12])
    DispAmend1 = DispAmend.values.tolist()
    DispExport.append(DispAmend1[0])
    


CExport = pandas.DataFrame(CoordExport, columns = ['x','y','z']) 
DExport = pandas.DataFrame(DispExport, columns = ['U1','U2','U3','R1','R2','R3']) 

ConcatExport = pandas.concat([CExport,DExport],axis = 1)

ConcatExport.to_excel('GrasshopperImport.xlsx', sheet_name='Sheet1',index = False)









































# StageZeros = pandas.DataFrame.transpose(StageZeros)

# Disp = wDisp

# for b in range(count):
#     try:
#         Row = StageZeros.loc[b].tolist()
#         Disp.loc[b*224-0.5+b] = Row
        
#     except:
#         pass

# Disp = Disp.sort_index().reset_index(drop=True)

# for c in range(len(Disp)):
    
#     if c % 225 == 0:
        
#         U1 = Disp.loc[c , 'U1']
#         U2 = Disp.loc[c , 'U2']
#         U3 = Disp.loc[c , 'U3']
#         R1 = Disp.loc[c , 'R1']
#         R2 = Disp.loc[c , 'R2']
#         R3 = Disp.loc[c , 'R3']
        
#     else:
        
#         Disp.loc[c , 'U1'] = Disp.loc[c , 'U1'] + U1
#         Disp.loc[c , 'U2'] = Disp.loc[c , 'U2'] + U2
#         Disp.loc[c , 'U3'] = Disp.loc[c , 'U3'] + U3
#         Disp.loc[c , 'R1'] = Disp.loc[c , 'R1'] + R1
#         Disp.loc[c , 'R2'] = Disp.loc[c , 'R2'] + R2
#         Disp.loc[c , 'R3'] = Disp.loc[c , 'R3'] + R3

# #Disp.to_excel(r'C:\Users\brendanlia\Desktop\Joint Data\Modified.xlsx', index = False)




