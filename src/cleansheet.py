# -*- coding: utf-8 -*-
"""
Created on Mon Mar  9 09:40:27 2020

@author: sayers
"""
import pandas as pd
import xlwings as xw
from xlwings import constants

def xl_col_sort(sheet,col_num):
    try:
        sheet.range('A2:X99999').api.Sort(Key1=sheet.range((2,col_num)).api, Order1=1)
    except:    
        sheet.range('A2:X9999').api.Sort(Key1=sheet.range((2,col_num)).api, Order1=1)

def get_col_widths(dataframe):
    # First we find the maximum length of the index column   
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

def set_widths(dataframe,worksheet):
    for i, width in enumerate(get_col_widths(dataframe)):
        worksheet.set_column(i, i, width)
    return(worksheet)


def cleansheet(nfname):
    df=pd.read_excel(nfname, index=False)
    writer = pd.ExcelWriter(nfname, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Data')
    worksheet = writer.sheets['Data']
    worksheet.freeze_panes(1,1)
    worksheet.autofilter(0, 0, df.shape[0], df.shape[1]-1)
    worksheet=set_widths(df,worksheet)
    writer.save()
    writer.close()
    
def cleansheet2(nfname):
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
    

def quickopen(nfname):
    wb = xw.Book(nfname) 
    wb.save()
    wb.close()
    
def dl_clean(filenamestring,df):
    if len(df) > 0:
        df.to_excel(filenamestring,index=False)
        cleansheet(filenamestring)

