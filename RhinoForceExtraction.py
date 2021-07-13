# -*- coding: utf-8 -*-
"""
Created on Tue Jul 13 09:37:53 2021

@author: brendanlia
"""

import numpy
import xlsxwriter
import pandas
import openpyxl
from openpyxl import load_workbook
import time
import win32com.client

jointNum = [148, 149, 450, 3611, 21491]#input('Joint Number: ')
#jointNum = str(jointNum)
stepNum = 1  #input('Step Number: ')
#stepNum = str(stepNum)

wCoord = pandas.read_excel(r'C:\Users\brendanlia\Desktop\Joint Data\Joint_Coordinates.xlsx', sheet_name = 'Joint Coordinates')
wCoord = pandas.DataFrame(wCoord)
wDisp = pandas.read_excel(r'C:\Users\brendanlia\Desktop\Joint Data\Joint_Displacement.xlsx', sheet_name = 'Joint Displacements')
wDisp = pandas.DataFrame(wDisp)

Coord = pandas.DataFrame.transpose(wCoord)
Disp = pandas.DataFrame.transpose(wDisp)


CoordExport = []
DispExport = []

for a in range(len(jointNum)):

    
    jointNumStr = str(jointNum[a])
    
    
    wCoordIndex = (wCoord[wCoord['TABLE:  Joint Coordinates'] == jointNumStr].index.values)
    wDispIndex1 = (wDisp[wDisp['TABLE:  Joint Displacements'] == jointNumStr].index.values)
    wDispIndex2 = (wDisp[wDisp['Unnamed: 4'] == stepNum].index.values)
    DispIndex = list(set(wDispIndex1) & set(wDispIndex2))

    CoordAmend = pandas.DataFrame.transpose(Coord[wCoordIndex][7:10])
    CoordAmend1 = CoordAmend.values.tolist()
    CoordExport.append(CoordAmend1[0])
    
    DispAmend = pandas.DataFrame.transpose(Disp[DispIndex][6:12])
    DispAmend1 = DispAmend.values.tolist()
    DispExport.append(DispAmend1[0])
    #DispExport.append(Disp[DispIndex][6:12])


CExport = pandas.DataFrame(CoordExport, columns = ['x','y','z']) 
DExport = pandas.DataFrame(DispExport, columns = ['U1','U2','U3','R1','R2','R3']) 







