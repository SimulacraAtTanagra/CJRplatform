# -*- coding: utf-8 -*-
"""
Created on Wed Jan 22 08:57:38 2020

@author: sayers
"""
from emailautosend import mailthis
from emailautosend import getemail
import os
import xlwings
import openpyxl
import pandas as pd
import re
from datetime import datetime 
from matplotlib import pyplot as plt

def newest(path,fname):
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files if basename.startswith(fname)]
    return max(paths, key=os.path.getmtime)

path = "S:\\Downloads\\"     # Give the location of the files
fname = "FULL_FILE"         # Give filename prefix
df = pd.read_excel(newest(path,fname))  #getting the newest of these files in the directory and converting to df
#stripping out the 2 metadata columns in CJR files
if re.match("R1013",df.columns.values[0]).group() == "R1013":
    new_header = df.iloc[1] #grab the first row for the header
    df = df[2:] #take the data less the header row
    df.columns = new_header #set the header row as the df header
#standardizing the column names
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')


"""
for i in sorted(list(df.columns.values)):
    print(i)
    
column_list = df.columns.values.tolist()
for column_name in column_list:
      print(df[column_name].unique())
"""

#removing the Federal Workstudy Records
df=df[df['company'] != "WSF"]

#finds not null end dates where the expired end date is before today
expired_end_dates = df[df.exp_job_end_dt.isnull() ==False][df['exp_job_end_dt'] < datetime.now()][df['empl_stat_cd']=="A"][['empl_id','person_nm','dept_descr_position','labor_job_ld','exp_job_end_dt']]
   
expired_leaves = df[df.return_dt.isnull() ==False][df['return_dt'] < datetime.now()][['empl_id','person_nm','dept_descr_position','labor_job_ld','empl_stat_ld','return_dt']]
possible_no_email = df[(~df.work_email.str.endswith('york.cuny.edu',na=False)) & (~df.work_email.isnull() == True)]

blank_email = df[(df.work_email.isnull() ==True)][df['empl_stat_cd']=="A"][['empl_id','empl_rcd','person_nm','dept_descr_position','labor_job_ld','pos_cd','empl_stat_cd']]
#blank_email.person_nm = blank_email.person_nm.str.replace(' ', '')
blank_email['global_address']=blank_email['person_nm'].apply(getemail)
blank_email[blank_email['global_address']!= '']
#if blank_email[blank_email['global_address']!= ''][df['jobcode_ld'].str.contains('Adj')].shape[0] >0:
    #mailthis('lolsson@york.cuny.edu','lwilkinson901@york.cuny.edu',blank_email[blank_email['global_address']!= ''][df['jobcode_ld'].str.contains('Adj')],'Please update these e-mail addresses in CF')
#if blank_email[blank_email['global_address']!= ''][df['jobcode_ld'].str.contains('College Assistant')].shape[0] >0:
    #mailthis('ajackson1@york.cuny.edu','lwilkinson901@york.cuny.edu',blank_email[blank_email['global_address']!= ''][df['jobcode_ld'].str.contains('College Assistant')],'Please update these e-mail addresses in CF')

remove_pos = df[df.pos_cd.isnull()==False][df['empl_stat_cd']=="T"][['empl_id','empl_rcd','person_nm','dept_descr_position','labor_job_ld','pos_cd','empl_stat_cd']]
remove_pos = remove_pos.append(df[df.pos_cd.isnull()==False][df['empl_stat_cd']=="R"][['empl_id','empl_rcd','person_nm','dept_descr_position','labor_job_ld','pos_cd','empl_stat_cd']])
#df = df[['empl_id','person_nm','dept_descr_position','labor_job_ld','empl_stat_ld','return_dt']]
#if remove_pos.shape[0] >0:
#    mailthis('lwilkinson901@york.cuny.edu','',remove_pos,"Position Numbers to Remove")

no_empl_class = df[df.empl_cls_ld.isnull() ==True].iloc[:, [0,1,5]]
citizenship_missing = df[df.citizenship_status.isnull() ==True].iloc[:, [0,1,5]]
citizenship_missing.shape[0]/df.shape[0]

#df.iloc[:,167]

#df[(df['Last_Name']=='AYERS') & (df['First_Name']=='SHANE ')].Ending.values[0]
#df.loc[(df['Last_Name']=="DAVIS" & df['First_Name']=='ALISHA '),'First_Name']

hrisgroup = "'lolsson@york.cuny.edu';'pcaceres901@york.cuny.edu';'ajackson1@york.cuny.edu';'mwilliams@york.cuny.edu';'lwilkinson901@york.cuny.edu'"

#mailthis(hrisgroup,'sayers@york.cuny.edu',expired_end_dates, 'Expired end date records (test)')

paf_empls = df[df.empl_stat_cd.isin(['A','S','R','L'])][['empl_id','person_nm','home_addr1', 'home_addr2', 'home_city','home_state','home_postal','jobcode_ld','labor_job_ld','budget_line_nbr','pos_cd']]
paf_empls.to_excel('Y:\PAF_report.xlsx')

active_emps = df[df['empl_stat_cd']=="A"][['empl_id','person_nm','jobcode_ld','labor_job_ld','dept_descr_job','comp_freq_job_ld']]
active_emps.to_excel('Z:\Registrar\Active_Employee_report.xlsx')
active_emps.to_excel('Z:\Security\Active_Employee_report.xlsx')

dfcross = pd.read_excel(newest(path,"HR_REPORT_PAYROLL_ID_LIST"))
dfpay= pd.read_excel(newest(path,"LOCKED_QUERY_1_"))
dfcross.columns = dfcross.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
dfpay.columns = dfpay.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')

df2 = dfcross[['nys_payserv_id','cf_empl_id']]
df3 = dfpay[['id','name','title','full/part','annual_salary','hrly._rate']]
df4 = df[['empl_id','person_nm','jobcode_ld','labor_job_ld','dept_descr_job','comp_freq_job_ld','comp_rt','appointment_hours','professional_hours']]

df2.columns = ['id','empl_id']
df4 = df4.astype({'empl_id':'int64'})

res = pd.merge(df2, df4, on='empl_id')
rest2 = pd.merge(res,df3, on='id')
rest2['cf_sal'] = rest2['comp_rt'] * (rest2['appointment_hours'] + rest2['professional_hours'])
rest2.cf_sal = rest2.cf_sal.astype('float64')
attemptedrest = rest2.loc[((rest2.cf_sal - rest2.annual_salary) > 1)&(rest2.cf_sal > 0)&(rest2.labor_job_ld.str.contains('Adj'))]
attemptedrest = attemptedrest[(~attemptedrest.labor_job_ld.str.contains('Non-'))&(~attemptedrest.labor_job_ld.str.contains('Tech')&(~attemptedrest.dept_descr_job.str.contains('EOC')))]
#rest2.to_excel('S:\\Downloads\\combine.xlsx')

#this is to track reports to information for people reporting to an ECP. Should expand to whole campus, ensure integrity
rtlist =['Dana Trimboli','Berenecea Johnson-Eanes', 'La Toro Yates']
df[df['person_nm'].isin(rtlist)][['empl_id','person_nm','jobcode_ld','labor_job_ld','dept_descr_job']]
ecp_group = df[df['paygroup_cd'] == '089'][['empl_id','person_nm','empl_stat_cd','jobcode_ld','labor_job_ld','dept_descr_job']]
rt_ecp = df[df['reports_to_emplid'].isin(ecp_group.empl_id)][['empl_id','effdt_job','person_nm','empl_stat_cd','jobcode_ld','labor_job_ld','dept_descr_job','reports_to_emplid']]
rt_inact = df[df['reports_to_emplid'].isin(df[df['empl_stat_cd']!='A'].empl_id)][['empl_id','effdt_job','person_nm','empl_stat_cd','jobcode_ld','labor_job_ld','dept_descr_job','reports_to_emplid']]
rt_inact = rt_inact[rt_inact.empl_stat_cd.isin(['A','S','R','P','L'])]
rt_update = rt_ecp[rt_ecp['reports_to_emplid'].isin(ecp_group[ecp_group['empl_stat_cd']!= 'A'].empl_id)]
#mailthis('sayers@york.cuny.edu','',rt_update,'Update these reports to values, Shane')

duplicatedrows = df[df.duplicated(['empl_id','jobcode_cd','dept_id_job','empl_stat_cd'],False)]
duplicatedrows = duplicatedrows[duplicatedrows['empl_stat_cd'].isin(['A','S','R','P','L'])][['empl_id','empl_rcd','effdt_job','person_nm','empl_stat_cd','jobcode_ld','labor_job_ld','dept_descr_job']]
#mailthis('lolsson@york.cuny.edu','',duplicatedrows,'Please consolidate these records with duplicate dept and title')

#dashboard summary segment
print('Our data currently has',expired_end_dates.shape[0],'records with expired end dates, comprising',int((expired_end_dates.shape[0]/active_emps.shape[0])*100), '% of all records')
print('Our data currently has',df[df.action_date > df.effdt_job][df.empl_stat_cd.isin(['A','S','R'])].shape[0],'records with effective dates earlier than the action date, comprising',int((df[df.action_date > df.effdt_job][df.empl_stat_cd.isin(['A','S','R'])].shape[0]/active_emps.shape[0])*100), '% of all records')
print('Our data currently has',blank_email.shape[0],'records with blank business e-mails, comprising',int((blank_email.shape[0]/active_emps.shape[0])*100), '% of all records')
"""
x=[df[df.action_date > df.effdt_job][df.empl_stat_cd.isin(['A','S','R'])].shape[0],blank_email.shape[0],expired_end_dates.shape[0],active_emps.shape[0]]
y=['Late','Email Missing','Expired','All Active']
plt.bar(y,x)
plt.show()

df51 = df[df['empl_stat_cd']=='A']
df51['realstatus'] = df['empl_id']*100
df51.loc['realstatus34'] = df51['empl_id']

# make this a function that accepts a dataframe and spits out a string
#this was for creating a population-specific distribution list
emails = df[df.empl_stat_cd.isin(['A','S','R','P','L'])][['first_nm','last_nm']]
emails['full_nm'] = emails['first_nm']+' '+emails['last_nm']
emails['address'] = emails['full_nm'].apply(getemail)
emails2 = df[(df.work_email.str.endswith('york.cuny.edu',na=False)) & (df.empl_stat_cd.isin(['A','S','R','P','L'])) & (df['job_family_ld']=='Faculty')][['person_nm','work_email']]
emails2 = emails2.dropna()

df[df['job_family_ld']=='Faculty']

emaillist= list(emails2['work_email'].unique())
completelist = ' ; '.join(emaillist)
employees = list(df['empl_id'].unique())

jobs = list(df.jobcode_ld.unique())
jobs2 = list(df.labor_job_ld.unique())
"""

currpay = df[['empl_id','empl_rcd', 'annl_rt']]
#currpay.to_excel('somefile.xlsx')
#process to run the birthdayreport from cjr, doing once a month at beginning of month
'''
bdays = df[(df.birth__dtmmdd.str.startswith('03/'))&(df.empl_stat_cd.isin(['A','S','R','P','L']))][['person_nm','work_email','birth__dtmmdd']]
bdaydist = list(bdays.work_email.dropna())
' ; '.join(bdaydist)
'''

