# -*- coding: utf-8 -*-
"""
Created on Wed Jan 22 08:57:38 2020

@author: sayers
"""

import pandas as pd
import re
import datetime
from matplotlib import pyplot as plt
import ast
from src.cleansheet import dl_clean
from src.admin import newest,colclean
from src.emailautosend import mailthis
from src.emailautosend import getemail
from src.subset import subsetlist as sl
  
def isValid(s): 
    
    # 1) Begins with 0 or 91 
    # 2) Then contains 7 or 8 or 9. 
    # 3) Then contains 9 digits 
    try:
        str(s).replace(' ','')
        Pattern = re.compile("[0-9]{10}") 
        if Pattern.match(s):
            return("valid")
        else:
            return("invalid")
    except:
        return("invalid")
def get_df_name(df):
    name =[x for x in globals() if globals()[x] is df][0]
    return name


def stripmail(x):
    try:
        if '@' not in str(x):
            return('')
    except:
        pass
    try:
        return(x.split('@')[0])
    except:
        return('')
def datedate(startdate,cutdate1,cutdate2):
    if startdate<datetime.date(cutdate1[0],cutdate1[1],cutdate1[2]):
        if startdate<datetime.date(cutdate2[0],cutdate2[1],cutdate2[2]):
           return(datetime.date(cutdate2[0],cutdate2[1],cutdate2[2]).strftime("%m/%d/%Y"))
        return(datetime.date(cutdate1[0],cutdate1[1],cutdate1[2]).strftime("%m/%d/%Y"))
        
path = "S:\\Downloads\\"     # Give the location of the files
fname = "FULL_FILE"         # Give filename prefix
df = pd.read_excel(newest(path,fname))  #getting the newest of these files in the directory and converting to df
#stripping out the 2 metadata columns in CJR files
if re.match("R1013",df.columns.values[0]).group() == "R1013":
    new_header = df.iloc[1] #grab the first row for the header
    df = df[2:] #take the data less the header row
    df.columns = new_header #set the header row as the df header
#standardizing the column names
df=colclean(df)
df['newname'] = df['first_nm'].str.cat(df['last_nm'], sep =" ") 
df=sl(df,['company',"~WSF"])

adjfile= sl(df,["df.empl_cls_ld","Adjuncts"],str1='empl_id,empl_rcd,dept_id_job,labor_job_ld')
adjfile.to_excel('s:\\downloads\\Adj_Records.xls')


#finds not null end dates where the expired end date is before today
expired_end_dates = df[df.exp_job_end_dt.isnull() ==False][df['exp_job_end_dt'] < datetime.datetime.now()][df['empl_stat_cd']=="A"][['empl_id','empl_rcd','person_nm','dept_descr_position','labor_job_ld','exp_job_end_dt']]
expired_end_dates.name="expired_end_dates"  
pttiles=['Non-Teaching Adjunct 3', 'College Assistant', 'Professor H',
       'Non-Teaching Adjunct 1', 'Lecturer H',
       'Continuing Ed Teacher-Hourly', 'Does Not Apply',  
       'Adjunct Assistant Professor', 'CLIP Instructor',
       'Adjunct Associate Professor',
       'Asst Professor Hourly', 'Non-Teaching Adjunct 5',
       'College Lab Tech', 'Assc Professor Hourly', 
       'Non-Teaching Adjunct 2', 'Non-Teaching Adjunct 4', 
       'Campus Security Asst', 'Adjunct Lecturer', 'IT Associate',
       'EOC Assistant', 'Adj Sr College Lab Tech', 'EOC Lecturer',
       'Custodial Assistant','IT Support Asst']
ftexp=sl(expired_end_dates,['labor_job_ld',pttiles])
dl_clean('s:\\downloads\\expired_end_ft.xls',ftexp)
leavestats=['Short Work Break', 'Leave of Absence',
       'Leave With Pay']
leavestr='empl_id,person_nm,dept_descr_position,labor_job_ld,empl_stat_ld,return_dt'
suspiciousleaves=sl(df,[['return_dt',None],['empl_stat_ld',leavestats]],str1=leavestr)
expired_leaves = df[df.return_dt.isnull() ==False][df['return_dt'] < datetime.datetime.now()][['empl_id','empl_rcd','person_nm','dept_descr_position','labor_job_ld','empl_stat_ld','return_dt']]
expired_leaves.name="expired_leaves"
ptleavetitles=['Non-Teaching Adjunct 2', 'Adjunct Assistant Professor',
       'College Assistant', 'Non-Teaching Adjunct 1', 'Adjunct Lecturer',
       'Adj College Lab Tech', 'Asst Professor Hourly',
       'Non-Teaching Adjunct 3', 'EOC Adjunct Lecturer',
       'Custodial Assistant', 'Continuing Ed Teacher-Hourly',
       'Assc Professor Hourly', 'Adjunct Associate Professor',
       'Campus Security Asst', 'Professor H', 'Adjunct Professor',
       'Non-Teaching Adjunct 5', 'Lecturer H',
       'College Lab Tech','Non-Teaching Adjunct 4']
ftleaves=sl(expired_leaves,['labor_job_ld',ptleavetitles+["~"]])
ftleaves=ftleaves.append(suspiciousleaves)
dl_clean('s:\\downloads\\ft_leaves.xls',ftleaves)
#this is specifically for use by the automated row insertion program
expired_end_dates.reset_index(drop=True, inplace=True)
expired_end_dates['startdate']=pd.DatetimeIndex(expired_end_dates.exp_job_end_dt) + pd.DateOffset(1)
expired_end_dates['returndt']= expired_end_dates.startdate.apply(datedate,args=((2020,8,25),(2020,6,30)))
expired_end_dates['enddate']= ''
expired_end_dates['actions']='Short Work Break'
expired_end_dates['reasons']='Short Work Break'

adjexp = expired_end_dates[expired_end_dates.labor_job_ld.str.contains("Adjunct")][['empl_id','empl_rcd','startdate','enddate','returndt','actions','reasons']]
adjexp.reset_index(drop=True, inplace=True)
adjexp['returndt']= adjexp.startdate.apply(datedate,args=((2020,8,25),(2020,5,31)))
adjexp.startdate= adjexp.startdate.dt.strftime("%m/%d/%Y")
caexp = expired_end_dates[expired_end_dates.labor_job_ld.str.contains("College Assistant")][['empl_id','empl_rcd','startdate','enddate','returndt','actions','reasons']]
caexp.reset_index(drop=True, inplace=True)
caexp.startdate=caexp.startdate.dt.strftime("%m/%d/%Y")

    

possible_no_email = df[(~df.work_email.str.endswith('york.cuny.edu',na=False)) & (~df.work_email.isnull() == True)]
work_email = df[df.empl_stat_cd.isin(['A','S','L'])][['empl_id','empl_rcd','work_email','newname','dept_descr_position','labor_job_ld','pos_cd','empl_stat_cd']]
try:
    work_email['global_address'] = work_email['newname'].apply(getemail)
except: 
    pass
work_email.work_email =work_email.work_email.str.lower()
work_email.global_address = work_email.global_address.str.lower()
try:
    work_email['username']= work_email.work_email.apply(stripmail)
except:
    pass
try:
    work_email['fixit'] = work_email.username.apply(getemail)
except Exception as e:
    print(e)
fix_email = work_email[(work_email.global_address.str.contains('york.cuny.edu'))&(work_email.work_email!= work_email.global_address)]
blank_email = sl(df,[['work_email',None]['empl_stat_cd',"A"]],'empl_id,empl_rcd,person_nm,dept_descr_position,labor_job_ld,pos_cd,empl_stat_cd,global_address')

try:
    blank_email[blank_email['global_address']!= '']
except:
    pass
blank_email.name="blank_email"
#if blank_email[blank_email['global_address']!= ''][df['jobcode_ld'].str.contains('Adj')].shape[0] >0:
    #mailthis('lolsson@york.cuny.edu','lwilkinson901@york.cuny.edu',blank_email[blank_email['global_address']!= ''][df['jobcode_ld'].str.contains('Adj')],'Please update these e-mail addresses in CF')
#if blank_email[blank_email['global_address']!= ''][df['jobcode_ld'].str.contains('College Assistant')].shape[0] >0:
    #mailthis('ajackson1@york.cuny.edu','lwilkinson901@york.cuny.edu',blank_email[blank_email['global_address']!= ''][df['jobcode_ld'].str.contains('College Assistant')],'Please update these e-mail addresses in CF')

remove_pos = sl(df,[['pos_cd','notnull'],['empl_stat_cd',"T|R"]],'empl_id,empl_rcd,person_nm,dept_descr_position,labor_job_ld,pos_cd,empl_stat_cd')
#remove_pos = remove_pos.append(df[df.pos_cd.isnull()==False][df['empl_stat_cd']=="R"][['empl_id','empl_rcd','person_nm','dept_descr_position','labor_job_ld','pos_cd','empl_stat_cd']])

no_empl_class = sl(df,['empl_cls_ld',None])
no_empl_class.name="no_empl_class"
citizenship_missing = sl(df,['citizenship_status',None])
citizenship_missing.name = "citizenship_missing"

hrisgroup = "'lolsson@york.cuny.edu';'pcaceres901@york.cuny.edu';'ajackson1@york.cuny.edu';'mwilliams@york.cuny.edu';'lwilkinson901@york.cuny.edu'"

#mailthis(hrisgroup,'sayers@york.cuny.edu',expired_end_dates, 'Expired end date records (test)')

paf_empls = sl(df,['empl_stat_cd',['A','S','P','R','L']],str1='empl_id,person_nm,home_addr1,home_addr2,home_city,home_state,home_postal,jobcode_ld,labor_job_ld,budget_line_nbr,pos_cd')
dl_clean('Y:\PAF_report.xlsx',paf_empls)

active_emps = sl(df,['empl_stat_cd',['A','S','P','L']],str1='empl_id,last_nm,person_nm,jobcode_ld,labor_job_ld,dept_descr_job,comp_freq_job_ld,comp_rt')
dl_clean('Z:\Registrar\Active_Employee_report.xlsx',active_emps)
dl_clean('Z:\Security\Active_Employee_report.xlsx',active_emps)
ethicsreport = active_emps[active_emps.comp_rt>101000]
dl_clean('Y:\ethics_report.xlsx',ethicsreport)

no_ethnicity = df[df.ethnicity_cuny == "NSPEC"]
no_ethnicity.name = "no_ethnicity"

df['valid_num']= df.home_phone.apply(isValid)
missing_phone_num = df[df.valid_num == "invalid"][['empl_id','person_nm','labor_job_ld','home_phone']]
missing_phone_num.name="wrongnums"
blank_reasons = df[df.action_reason_ld.isnull()==True][['empl_id','person_nm','labor_job_ld','action_ld','action_reason_ld']]
blank_reasons.name='blankreasons'
auditlist=[missing_phone_num,blank_reasons, citizenship_missing,no_empl_class,blank_email,expired_end_dates,expired_leaves,no_ethnicity]
for i in auditlist:
    print(f'Our data currently has {i.shape[0]} records with the issue of {i.name} comprising {int((i.shape[0]/active_emps.shape[0])*100)} % of all records')
'''dfx = df[['first_nm','last_nm','empl_id']]
dfx['code']= df.apply(f"{dfx['first_nm']}{dfx['last_nm']}{dfx['empl_id'][-3:]}")

dfcross = pd.read_excel(newest(path,"HR_REPORT_PAYROLL_ID_LIST"))
dfpay= pd.read_excel(newest(path,"LOCKED_QUERY_1_"))
dfcross=colclean(dfcross)
dfpay=colclean(dfpay)

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
'''
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

#ast.literal_eval('df[df.emplid="23134914"]')

#reports to information section
rtinfo=df[df.hr_status=='Active'][['dept_id_position', 'dept_descr_position',
 'dept_id_job', 'dept_descr_job', 'pos_cd', 'dept_mgr_pos_cd', 'dept_mgr_pos_ld', 'dept_mgr_id',
 'dept_mgr_name', 'dept_mgr_emplstatus',  'reports_to_position', 'reports_to_position_descr',
 'reports_to_emplid', 'reports_to_name', 'reports_to_emplstatus']]
rtinfo=rtinfo.drop_duplicates()
depttable=rtinfo[['dept_id_position', 'dept_descr_position']]
depttable=depttable.drop_duplicates()
depttable = depttable[depttable.dept_descr_position.isnull()==False]
depttable = depttable[depttable.dept_id_position.isnull()==False]
deptnums=list(depttable.dept_id_position.unique())
deptmgrs=list(rtinfo.dept_mgr_name.unique())
deptrts=list(rtinfo.reports_to_name.unique())
inactdeptmgrs=rtinfo[rtinfo.dept_mgr_emplstatus!='A'][['dept_id_position']]
inactdeptmgrs=inactdeptmgrs.drop_duplicates()
#del(deptnums[2])
def checkdeptnums(deptcode):
    return(len(depttable[depttable.dept_id_position==deptcode]))
depttable['countdepts']=depttable.dept_id_position.apply(checkdeptnums)
depttable[depttable.countdepts>1]
for i,x in enumerate(deptnums):
    if len(df[(df.dept_id_position==x)&(df.hr_status=='Active')])<1:
        print(x)
#mailthis('lolsson@york.cuny.edu','', fix_email[fix_email.labor_job_ld.str.contains("Adj")],'Adjuncts with mismatches','')

#dashboard summary segment
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
if datetime.datetime.now().day == 1:
    bdays = df[(df.birth__dtmmdd.str.startswith(f'0{datetime.datetime.now().month}'))&(df.empl_stat_cd.isin(['A','S','R','P','L']))][['person_nm','work_email','birth__dtmmdd']]
    bdaydist = list(bdays.work_email.dropna())
    print(' ; '.join(bdaydist))
'''
ca_reappt = df[(df.labor_job_ld.str.contains("College As"))&(df.hr_status.isin(['Active']))][['empl_id','person_nm','labor_job_ld','dept_descr_job','reports_to_emplid']]
ca_reappt.columns = ca_reappt.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')

peoplenames = df[['empl_id','person_nm']]
peoplenames = peoplenames.drop_duplicates()
deptmgrs = df[df.hr_status.isin(['Active'])][['dept_id_job','dept_descr_job','dept_mgr_id']]
deptmgrs = deptmgrs.drop_duplicates()
deptmgrs.columns = ['hcm_dept','dept_descr_job','dept_mgr_id']

rt = df[df.hr_status.isin(['Active'])][['empl_id','dept_id_job','reports_to_emplid']]
rt.columns = ['empl_id','hcm_dept','reports_to']
depts = pd.read_excel('S:\\Downloads\\prasstdepts.xlsx')
depts=colclean(depts)
#list(depts.columns.to_list())
depts1= depts[['code','hcm_dept']]

prasstdata = pd.read_excel("S:\\Downloads\\CrystalReportViewer1 - 2020-04-10T153115.706.xls")
prasstdata.columns = prasstdata.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
#list(prasstdata.columns.to_list())
list(set(prasstdata.title.to_list()))
active_ca= prasstdata[(prasstdata.status=="ACTIVE")&(prasstdata.title.isin(['EOC College Assistant','COLLEGE ASST','CUNY CAP COLLEGE ASST']))][['ss_#','dept']]
active_ca.columns =['empl_id','code']
new111= pd.merge(active_ca,peoplenames,on='empl_id',how='left')
new111 = pd.merge(new111,depts1,on='code', how='left')
new111.hcm_dept= new111.apply(lambda x : x['code'] if pd.isnull(x['hcm_dept']) else x['hcm_dept'], axis=1 )
deptmgrs.hcm_dept = deptmgrs.hcm_dept.astype('int64')
new111.hcm_dept =new111.hcm_dept .astype('int64')
new111 = pd.merge(new111,deptmgrs,on='hcm_dept',how='left')
new111 = pd.merge(new111,peoplenames,how='left',left_on='dept_mgr_id',right_on='empl_id')
new111.columns.to_list()
df.columns.to_list()
rt.empl_id = rt.empl_id.astype('int64')
rt.hcm_dept = rt.hcm_dept.astype('int64')
new111.empl_id_x=new111.empl_id_x.astype('int64')
new111 = pd.merge(new111,rt,how='left',left_on=['empl_id_x','hcm_dept'],right_on=['empl_id','hcm_dept'])
new111 = pd.merge(new111,peoplenames,how='left',left_on='reports_to',right_on='empl_id')
new111.columns.to_list()
ca_final_list = new111[['empl_id_x','person_nm_x','hcm_dept','dept_descr_job','person_nm_y','person_nm']]
ca_final_list.columns = ['empl1_id','empl_id2','person_nm','dept_code','dept_name','dept_head','reports_to']
ca_final_list = ca_final_list.drop('empl_id2',axis=1)
ca_final_list.to_excel("S:\\Downloads\\calist.xlsx")
'''