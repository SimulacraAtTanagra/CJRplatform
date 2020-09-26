# -*- coding: utf-8 -*-
"""
Created on Mon Mar  9 09:40:27 2020

@author: sayers
"""
import xlwings as xw
from xlwings import constants

def xl_col_sort(sheet,col_num):
    try:
        sheet.range('A2:X99999').api.Sort(Key1=sheet.range((2,col_num)).api, Order1=1)
    except:    
        sheet.range('A2:X9999').api.Sort(Key1=sheet.range((2,col_num)).api, Order1=1)

def cleansheet(nfname):
    wb = xw.Book(nfname) 
    wb.sheets['Sheet1'].autofit()
    wb.save()
    try:
        xw.Range("A:A").api.Delete(constants.DeleteShiftDirection.xlShiftUp)
    except:
        print("Didn't work this time boss")
        pass
    xl_col_sort(wb.sheets['Sheet1'],2)
    wb.save()
    active_window = wb.app.api.ActiveWindow
    active_window.FreezePanes = False
    active_window.SplitColumn = 0
    active_window.SplitRow = 1
    active_window.FreezePanes = True
    app = xw.apps.active 
    wb.save()
    app.quit()
    
def dl_clean(filenamestring,df):
    if len(df) > 0:
        df.reset_index(drop=True, inplace=True).to_excel(filenamestring)
        cleansheet(filenamestring)

