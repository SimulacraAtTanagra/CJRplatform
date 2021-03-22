# -*- coding: utf-8 -*-
"""
Created on Wed Jan 22 08:57:38 2020

@author: sayers
"""
#Wow, what a mess. Let's see if we can't fix this. 


#from src.emailautosend import mailthis
from src.emailautosend import getemail
from src.cleansheet import dl_clean
import pandas as pd
import re
import datetime
from src.admin import newest,colclean,rehead
from src.subset import src.subsetlist
import os
  
def isValid(s): #for phone number validation
    try:
        str(s).replace(' ','')
        Pattern = re.compile("[0-9]{10}") 
        if Pattern.match(s):
            return("valid")
        else:
            return("invalid")
    except:
        return("invalid")

def get_df_name(df):    #for establishing name variable for use in reporting
    name =[x for x in globals() if globals()[x] is df][0]
    return name

def stripmail(x):   #for validating e-mail addreses
    try:
        if '@' not in str(x):
            return('')
        return(x.split('@')[0])
    except:
        return('')     

def datedate(startdate,cutdate1,cutdate2):  #I would be lying if I said I remembered why I wrote this
    if startdate<datetime.date(cutdate1[0],cutdate1[1],cutdate1[2]):
        if startdate<datetime.date(cutdate2[0],cutdate2[1],cutdate2[2]):
           return(datetime.date(cutdate2[0],cutdate2[1],cutdate2[2]).strftime("%m/%d/%Y"))
        return(datetime.date(cutdate1[0],cutdate1[1],cutdate1[2]).strftime("%m/%d/%Y"))

def load_data(path,fname):
    #getting the newest of these files in the directory and converting to df
    #stripping out the 2 metadata columns in CJR files
    #standardizing the column names
    df = colclean(rehead(pd.read_excel(newest(path,fname)),2))
    return(df)


def add_calc_col(df,title,col,func,subs=False): #this abstracts applying a bit
    df[title]=df[col].apply(func)
    if subs:
        df=df.iloc[:,subs]
    return(df)
    
def filesubset(df,cols,conds,filename=None,subs=None,name=None):   #procedural subsetting and writing to excel
    if subs:
        subs=','.join(subs)
        df=subsetlist(df,[cols,conds],str1=subs)
    else:
        df=subsetlist(df,[cols,conds])
    if name:
        df.name=name
    if filename:
        df.to_excel(filename)
    else:
        return(df)
def multifilesubset(df:pd.DataFrame,basename:str,filenames:list,argsl:list):      
    for ix,file in enumerate(filenames):
        if ":" in file:
            filename=file
        else:
            filename=os.path.join(basename,file)
        col_cond_subs=argsl[ix]
        if type(col_cond_subs[0])==str:  #if this is a singular subset argument
            if len(col_cond_subs)>2:   #and there's three arguments
                filesubset(df,col_cond_subs[0],col_cond_subs[1],filename=filename)
            else:
                subs=','.join(col_cond_subs[2])
                filesubset(df,col_cond_subs[0],col_cond_subs[1],filename=filename,subs=subs)
        elif type(col_cond_subs[0])==list:
            df1=df
            for grouping in col_cond_subs:
                if type(grouping[0])==str:  #if this is a singular subset argument
                    if len(grouping)>2:   #and there's three arguments
                        df1=filesubset(df1,grouping[0],grouping[1])
                    else:
                        subs=','.join(grouping[2])
                        df1=filesubset(df1,grouping[0],grouping[1],subs=subs)
            df1.to_excel(filename)
        print(f"Written file {filename}")

def mass_subset(df:pd.DataFrame,names:list,argsl:list):
    #creates a list of named data frames for audit purposes
    auditlist=[]
    for ix,name in enumerate(names):
        col_cond_subs=argsl[ix]
        if type(col_cond_subs[0])==str:  #if this is a singular subset argument
            if len(col_cond_subs)>2:   #and there's three arguments
                df1=filesubset(df,col_cond_subs[0],col_cond_subs[1],name=name)
            else:
                subs=','.join(col_cond_subs[2])
                df1=filesubset(df,col_cond_subs[0],col_cond_subs[1],subs=subs,name=name)
        elif type(col_cond_subs[0])==list:
            df1=df
            for grouping in col_cond_subs:
                if type(grouping[0])==str:  #if this is a singular subset argument
                    if len(grouping)>2:   #and there's three arguments
                        df1=filesubset(df1,grouping[0],grouping[1],name=name)
                    else:
                        subs=','.join(grouping[2])
                        df1=filesubset(df1,grouping[0],grouping[1],subs=subs,name=name)
        auditlist.append(df1)
    return(auditlist)

def main(path): #literally just doing this so I can run stuff unmolested
    fname = "FULL_FILE"         # Give filename prefix
    df=load_data(path,fname)    
    
    #consolidating all of my lines of spaghetti into a few well-ordered functions
    #Adjunct Records is...
    #PAF Report is a document with only the information reuqired to complete a Personnel Action Form
    #Active Employee report is all active employees for Registrar and Security
    #^Possibly deprecated client-side
    
    #TODO complete list of files being written to disk
    filenames=['Adj_Records.xls','Y:\PAF_report.xlsx',
               'Z:\Registrar\Active_Employee_report.xlsx',
               'Z:\Security\Active_Employee_report.xlsx',
               'Y:\Current Data\emplids.xlsx'
               ]
    arguments=[['empl_cls_ld',"Adjuncts",['empl_id','empl_rcd','dept_id_job','labor_job_ld']],
     ["empl_stat_cd",["A","S","R","L"],['empl_id','person_nm','home_addr1', 'home_addr2', 'home_city','home_state','home_postal','jobcode_ld','labor_job_ld','budget_line_nbr','pos_cd']],
     ["empl_stat_cd",["A","S","P","L"],['empl_id','last_nm','person_nm','jobcode_ld','labor_job_ld','dept_descr_job','comp_freq_job_ld','comp_rt']],
     ["empl_stat_cd",["A","S","P","L"],['empl_id','last_nm','person_nm','jobcode_ld','labor_job_ld','dept_descr_job','comp_freq_job_ld','comp_rt']],
     [],
     [],
     ]
    multifilesubset(df,filenames,arguments)
    df['newname'] = df['first_nm'].str.cat(df['last_nm'], sep =" ") 
    
    #removing the Federal Workstudy Records
    df=df[df['company'] != "WSF"]
    
    
    
   
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
    ftexp=expired_end_dates[~expired_end_dates.labor_job_ld.isin(pttiles)]
    dl_clean('s:\\downloads\\expired_end_ft.xls',ftexp)
    
    suspiciousleaves=df[(df.return_dt.isnull()==True)&(df.empl_stat_ld.isin(['Short Work Break', 'Leave of Absence',
           'Leave With Pay']))][['empl_id','person_nm','dept_descr_position','labor_job_ld','empl_stat_ld','return_dt']]
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
    ftleaves=expired_leaves[~expired_leaves.labor_job_ld.isin(ptleavetitles)]
    ftleaves=ftleaves.append(suspiciousleaves)
    dl_clean('s:\\downloads\\ft_leaves.xls',ftleaves)
    
    #this is specifically for use by the automated row insertion program
    expired_end_dates.reset_index(drop=True, inplace=True)
    expired_end_dates['startdate']=pd.DatetimeIndex(expired_end_dates.exp_job_end_dt) + pd.DateOffset(1)
    expired_end_dates['returndt']= expired_end_dates.startdate.apply(datedate,args=((2020,8,25),(2020,6,30)))
    expired_end_dates['enddate']= ''
    expired_end_dates['actions']='Short Work Break'
    expired_end_dates['reasons']='Short Work Break'
    
    try:
        adjexp = expired_end_dates[expired_end_dates.labor_job_ld.str.contains("Adjunct")][['empl_id','empl_rcd','startdate','enddate','returndt','actions','reasons']]
        adjexp.reset_index(drop=True, inplace=True)
        adjexp['returndt']= adjexp.startdate.apply(datedate,args=((2020,8,25),(2020,5,31)))
        adjexp.startdate= adjexp.startdate.dt.strftime("%m/%d/%Y")
        caexp = expired_end_dates[expired_end_dates.labor_job_ld.str.contains("College Assistant")][['empl_id','empl_rcd','startdate','enddate','returndt','actions','reasons']]
        caexp.reset_index(drop=True, inplace=True)
        caexp.startdate=caexp.startdate.dt.strftime("%m/%d/%Y")
    except:
        pass
    
        
    
    possible_no_email = df[(~df.work_email.str.endswith('york.cuny.edu',na=False)) & (~df.work_email.isnull() == True)]
    
    work_email = df[df.empl_stat_cd.isin(['A','S','L'])][['empl_id','empl_rcd','work_email','newname','dept_descr_position','labor_job_ld','pos_cd','empl_stat_cd']]
    work_email=add_calc_col(work_email,'global_address','newname',getemail)
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
    try:
        blank_email = df[(df.work_email.isnull() ==True)][df['empl_stat_cd']=="A"][['empl_id','empl_rcd','person_nm','dept_descr_position','labor_job_ld','pos_cd','empl_stat_cd','global_address']]
    except:
        blank_email = df[(df.work_email.isnull() ==True)][df['empl_stat_cd']=="A"][['empl_id','empl_rcd','person_nm','dept_descr_position','labor_job_ld','pos_cd','empl_stat_cd']]
    
    #blank_email.person_nm = blank_email.person_nm.str.replace(' ', '')
    
    try:
        blank_email[blank_email['global_address']!= '']
    except:
        pass
    blank_email.name="blank_email"
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
    no_empl_class.name="no_empl_class"
    citizenship_missing = df[df.citizenship_status.isnull() ==True].iloc[:, [0,1,5]]
    citizenship_missing.name = "citizenship_missing"
    citizenship_missing.shape[0]/df.shape[0]
    
    #df.iloc[:,167]
    
    #df[(df['Last_Name']=='AYERS') & (df['First_Name']=='SHANE ')].Ending.values[0]
    #df.loc[(df['Last_Name']=="DAVIS" & df['First_Name']=='ALISHA '),'First_Name']
    
    hrisgroup = "'lolsson@york.cuny.edu';'pcaceres901@york.cuny.edu';'ajackson1@york.cuny.edu';'mwilliams@york.cuny.edu';'lwilkinson901@york.cuny.edu'"
    
    #mailthis(hrisgroup,'sayers@york.cuny.edu',expired_end_dates, 'Expired end date records (test)')
   
         #ethicsreport = active_emps[active_emps.comp_rt>101000]
    #cleansheet(ethicsreport,'s:\\downloads\\ethics_report.xlsx')
    
    no_ethnicity = df[(df.ethnicity_cuny.isin(["NSPEC",'Unknown']))|(df.ethnicity_federal.isin(["NSPEC",'Unknown']))][['empl_id','person_nm']].drop_duplicates()
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
            
    ca_nonreappt_audit = df[(df.comp_freq_job_id=='H')&(df.job_func_cd=='C')][['effdt_job','empl_id','person_nm','dept_id_job','dept_descr_job','jobcode_cd','labor_job_ld','action_reason_ld','lbr_job_cl_entry_dt']]
    ca_nonreappt_audit['combocode']=ca_nonreappt_audit.empl_id.astype('str')+ca_nonreappt_audit.dept_id_job.astype('str')
    #mailthis('lolsson@york.cuny.edu','', fix_email[fix_email.labor_job_ld.str.contains("Adj")],'Adjuncts with mismatches','')
    finaldf=colclean(pd.read_excel(newest(path,'finaldf')))
    finaldf['combocode']=finaldf.empl_id.astype('str')+finaldf.dept.astype('str')
    finalnonreappt=ca_nonreappt_audit[ca_nonreappt_audit.effdt_job >datetime.datetime(2020,6,25,0,0,0)][~ca_nonreappt_audit.combocode.isin(finaldf.combocode.unique())]
    finalnonreappt.to_excel('s://downloads//canonreappts.xls')
    finalreappt=ca_nonreappt_audit[ca_nonreappt_audit.combocode.isin(finaldf.combocode.unique())]
    finalreappt.to_excel('s://downloads//careappts.xls')
    
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
        bdays = df[(df.birth__dtmmdd.str.startswith(f'{datetime.datetime.now().month}'))&(df.empl_stat_cd.isin(['A','S','R','P','L']))][['person_nm','work_email','birth__dtmmdd']]
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
    
    
    

#finds not null end dates where the expired end date is before today
#expired_end_dates = df[df.exp_job_end_dt.isnull() ==False][df['exp_job_end_dt'] < datetime.now()]   
#expired_leaves = df[df.return_dt.isnull() ==False][df['return_dt'] < datetime.now()][['empl_id','person_nm','dept_descr_position','labor_job_ld','empl_stat_ld','return_dt']]
#df[(df.work_email.isnull() ==True) | (df[df['work_email'].str.endswith('york.cuny.edu')])]
#df[df.work_email.isnull() ==True]
#df.drop(df[df['work_email'].str.endswith('york.cuny.edu',na=False)],axis=1)

#df = df[['empl_id','full_name','dept_descr_position','labor_job_ld','empl_stat_ld','return_dt']]



if __name__=="__main__":
    path = "C:\\users\\shane\\Downloads\\"     # Give the location of the files
    fname = "FULL_FILE"         # Give filename prefix
    df=load_data(path,fname)
    filesubset(df,'empl_cls_ld',"Adjuncts",os.path.join(path,'adjfile.xls'),subs='empl_id,empl_rcd,dept_id_job,labor_job_ld')