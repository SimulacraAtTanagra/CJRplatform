# -*- coding: utf-8 -*-
"""
Created on Sat Sep 26 13:08:58 2020

@author: shane
"""



import pandas as pd
def str_constr(cols,conds):
    str1=f"{cols}.astype('str')"
    str2=f".str.contains('{conds}')"
    str3=',engine="python"'
    str4='"'+str1+str2+'"'+str3
    return(str(str4))

def str_tran(i:list) -> str:    #generates string for subsetting
    cols=i[0]
    conds=i[1]
    strlist =[]
    try:
        if type(conds)==str:
            if conds=="notnull":
                strlist.append(f'{cols} == {cols}')
            elif "|" in conds:
                strlist.append(f'{cols}== {conds.split("|")[0]} or {cols}== {conds.split("|")[1]}')
            elif "~" in conds[0]:
                strlist.append(f"{cols} != '{conds[1:]}'")
            elif "^" in conds[0]:
                fstring=str_constr(cols,conds)
                strlist.append(fstring)
            elif "$" in conds[0]:
                fstring=str_constr(cols,conds[1:]+"$")
                strlist.append(fstring)
            elif "!" in conds[0]:
                fstring=str_constr(cols,conds[1:])
                strlist.append(fstring)
            else:
                strlist.append(f"{cols} == '{conds}'")
        elif type(conds)==list:
            if "~" in conds:
                strlist.append(f'{cols} not in {conds[:-1]}')
            else:
                strlist.append(f'{cols} in {conds}')
        else:
           strlist.append(f'{cols} != {cols} ') 
    except:
        strlist.append(f'{cols}  != {cols} ')
    if len(strlist)>1:
        return(' and '.join(strlist))
    else:
        return(strlist[0])

def subset(df:pd.DataFrame,sublist:list)-> pd.DataFrame:    #subset dataframe
    condition=str_tran(sublist)
    df=eval(f'''df.query("{condition}")''')
    return(df)
    
def subsetlist(df:pd.DataFrame,sublist:list,str1=None) -> pd.DataFrame: #for multiples
    """
    This function allows the procedural subsetting of a dataframe.
    It accepts list of lists. Each list is one or more pairs of columns and 
    values to filter by. 
    If value is Nonetype, will automatically filter out cells where column is empty. 
    Accepts ~ at beginning of value string to denote a not operation. 
    Accepts ^ at beginning of value string to denote a startswith operation
    Accepts $ at beginning of value string to denoate an endswith operation
    Accepts ! at beginning of value string to denoate a contains operation
    Above operators cannot currently be compounded. Use individually ONLY.
    string as str1 kwarg and will split on comma to filter columns.  
    """
    if type(sublist[0])==list:
        for i in sublist:
            df=subset(df,i)
    else:
        df=subset(df,sublist)
    if str1:
        if type(str1)==str:
            df1=df[[w for w in str1.split(',') if w in list(df.columns)]]
            if len(df1.columns)>0:
                df=df1
            return(df)
    return(df)
    