# -*- coding: utf-8 -*-
"""
Created on Fri Jul  9 10:41:34 2021

@author: brendanlia
"""

import numpy
import xlsxwriter
import pandas
import openpyxl
from openpyxl import load_workbook
import time
import win32com.client

def refreshExcel():
    xlapp = win32com.client.DispatchEx("Excel.Application")
    # The File location will need to be edited for different file locations
    wb = xlapp.Workbooks.Open(r"C:\Users\brendanlia\Desktop\Airport Excel Data\CombineTest\CombinedMinMax.xlsx") 
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Save()
    xlapp.Quit()

'----------------------------------------------------------------------------'

MaxFx1 =  pandas.read_excel("Max Fx.xlsx")
MaxFy1 =  pandas.read_excel("Max Fy.xlsx")
MaxFz1 =  pandas.read_excel("Max Fz.xlsx")
MinFx1 =  pandas.read_excel("Min Fx.xlsx")
MinFy1 =  pandas.read_excel("Min Fy.xlsx")
MinFz1 =  pandas.read_excel("Min Fz.xlsx")
MaxMx1 =  pandas.read_excel("Max Mx.xlsx")
MaxMy1 =  pandas.read_excel("Max My.xlsx")
MinMx1 =  pandas.read_excel("Min Mx.xlsx")
MinMy1 =  pandas.read_excel("Min My.xlsx")

writer = pandas.ExcelWriter("CombinedMinMax.xlsx", engine = 'xlsxwriter')

MaxFx1.to_excel(writer, sheet_name = 'MaxFx')
MaxFy1.to_excel(writer, sheet_name = 'MaxFy')
MaxFz1.to_excel(writer, sheet_name = 'MaxFz')

MinFx1.to_excel(writer, sheet_name = 'MinFx')
MinFy1.to_excel(writer, sheet_name = 'MinFy')
MinFz1.to_excel(writer, sheet_name = 'MinFz')

MaxMx1.to_excel(writer, sheet_name = 'MaxMx')
MaxMy1.to_excel(writer, sheet_name = 'MaxMy')

MinMx1.to_excel(writer, sheet_name = 'MinMx')
MinMy1.to_excel(writer, sheet_name = 'MinMy')

writer.save()