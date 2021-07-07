# -*- coding: utf-8 -*-
"""
Created on Wed Jul  7 09:17:36 2021

@author: brendanlia
"""

import numpy
import xlsxwriter
import pandas
import easygui 
import openpyxl
from openpyxl import load_workbook
import time
import win32com.client

'--------------------------------------------------------------------------------------------------------------------------------------------------'
#Def Functions

def refreshExcel():
    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(r"C:\Users\brendanlia\Desktop\Airport Excel Data\19 tree base result check\BoltCalc1.xlsx")
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Save()
    xlapp.Quit()
    
'----------------------------------------------------------------------------------------------------------------------------------------------------'
    
def getMaxUR(ForceChoice, BoltCalcStr, sheetCalcStr, ForceCombinedStr, ForceStr):
    # ForceChoice = selction of MaxFx, MaxFy, ... , MinFy
    # BoltCalcStr = file name of the bolt calc spreadsheet, input as a string e.g. "BoltCalc.xlsx"
    # SheetCalcStr = the name of the sheet within BoltCalcStr that is the desired bolt configuration, input as string, e.g. "BP (CAP14)"
    # ForceCombinedStr = the file name of the combined forces xlsx doc, input as string, e.g. "CombinedMinMax.xlsx"
    #ForceStr = the Force Choice in string format, input as string, e.g. "MaxFx"
    
    ex = []
    EXX = []
    
    for a in range(len(ForceChoice.columns)): #Input new Force Values
        workbook = load_workbook(filename = BoltCalcStr)
        sheet = workbook[sheetCalcStr]
        sheet["F6"] = ForceChoice[a][6]
        sheet["F7"] = ForceChoice[a][6]
        sheet["F8"] = ForceChoice[a][4]
        sheet["F9"] = ForceChoice[a][5]
        sheet["F10"] = ForceChoice[a][7]
        sheet["F11"] = ForceChoice[a][8]
        workbook.save(filename = BoltCalcStr)
       
        #Allow Excel to refresh
        time.sleep(0.1)
        refreshExcel()
        
        #Get UR Values and Calculate Max
        wb = load_workbook(filename = BoltCalcStr, data_only = True)
        ws = wb[sheetCalcStr]
        your_data = pandas.DataFrame(ws.values)
        
        ex = [your_data[13][32], your_data[13][33], your_data[13][34],your_data[13][35], your_data[13][36], your_data[13][37], your_data[13][38], your_data[13][39]]
        ex1 = max(ex)
        EXX.append(ex1)
        print(EXX)
        
    #append to CombinedForces Excel
    workbook = load_workbook(filename = ForceCombinedStr)
    sheet = workbook[ForceStr]
    sheet["K1"] = "Max UR"
    for a in range(len(EXX)):
        b = str(a+2)
        c = "K"
        d = c+b
        sheet[d] = EXX[a]
    
    workbook.save(filename = ForceCombinedStr)

'----------------------------------------------------------------------------------------------------------------------------------------------------'
# Get Min Max Forces Data

MaxFx = pandas.read_excel(r'C:\Users\brendanlia\Desktop\Airport Excel Data\19 tree base result check\CombinedMinMax.xlsx', sheet_name = 'MaxFx')
MaxFy = pandas.read_excel(r'C:\Users\brendanlia\Desktop\Airport Excel Data\19 tree base result check\CombinedMinMax.xlsx', sheet_name = 'MaxFy')
MaxFz = pandas.read_excel(r'C:\Users\brendanlia\Desktop\Airport Excel Data\19 tree base result check\CombinedMinMax.xlsx', sheet_name = 'MaxFz')
MinFx = pandas.read_excel(r'C:\Users\brendanlia\Desktop\Airport Excel Data\19 tree base result check\CombinedMinMax.xlsx', sheet_name = 'MinFx')
MinFy = pandas.read_excel(r'C:\Users\brendanlia\Desktop\Airport Excel Data\19 tree base result check\CombinedMinMax.xlsx', sheet_name = 'MinFy')
MinFz = pandas.read_excel(r'C:\Users\brendanlia\Desktop\Airport Excel Data\19 tree base result check\CombinedMinMax.xlsx', sheet_name = 'MinFz')

MaxFx = pandas.DataFrame.transpose(MaxFx)
MaxFy = pandas.DataFrame.transpose(MaxFy)
MaxFz = pandas.DataFrame.transpose(MaxFz)
MinFx = pandas.DataFrame.transpose(MinFx)
MinFy = pandas.DataFrame.transpose(MinFy)
MinFz = pandas.DataFrame.transpose(MinFz)

getMaxUR(MaxFx, "BoltCalc1.xlsx", "BP (CAP14)", "CombinedMinMax.xlsx", "MaxFx")
getMaxUR(MaxFy, "BoltCalc1.xlsx", "BP (CAP14)", "CombinedMinMax.xlsx", "MaxFy")
getMaxUR(MaxFz, "BoltCalc1.xlsx", "BP (CAP14)", "CombinedMinMax.xlsx", "MaxFz")
getMaxUR(MaxFx, "BoltCalc1.xlsx", "BP (CAP14)", "CombinedMinMax.xlsx", "MinFx")
getMaxUR(MinFy, "BoltCalc1.xlsx", "BP (CAP14)", "CombinedMinMax.xlsx", "MinFy")
getMaxUR(MinFz, "BoltCalc1.xlsx", "BP (CAP14)", "CombinedMinMax.xlsx", "MinFz")

input('Press Enter to Close')
    









