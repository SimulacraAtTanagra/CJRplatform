# -*- coding: utf-8 -*-
"""
Created on Wed Sep  9 08:12:46 2020

@author: sayers
"""

import pandas as pd
import os


#this is an administrative source file
#it holds code used in most, if not all, of my other work-related projects

def newest(path,fname):     #this function returns newest file in folder by name
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files if fname in basename]
    return max(paths, key=os.path.getmtime)

def colclean(df):           #this file make dataframe headers more manageable
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
    return(df)
    
def renamefile(path,fname,newname):
    newpath = path+newname
    os.rename(r''+newest(path,fname),r''+newpath)
    
def retrieve(df_name,fname):
    x=df_name
    df_name=pd.read_excel(fname)
    df_name.name=x
    return(df_name)
    
def mover(path,fname,dest):
    oldpath=path+fname
    if path[-2:]!="\\":
        path+="\\"
    if dest[-2:]!="\\":
        dest+="\\"
    newpath=dest+fname
    os.rename(oldpath,newpath)