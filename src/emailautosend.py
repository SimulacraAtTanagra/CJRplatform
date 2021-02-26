# -*- coding: utf-8 -*-
"""
Created on Wed Jan 22 12:17:23 2020

@author: sayers
"""
from re import search
import win32com.client as win32
from tabulate import tabulate

def getemail(search_string):
    try:    
        search_string = str(search_string)
        outlook = win32.Dispatch('outlook.application')
        gal = outlook.Session.GetGlobalAddressList()
        entries = gal.AddressEntries
        ae = entries[search_string]
        email_address = None
        
        if search(f'{search_string}$',str(ae)) != None:
           pass
        else:
           return('')
        
        if 'EX' == ae.Type:
            eu = ae.GetExchangeUser()
            email_address = eu.PrimarySmtpAddress
           
        
        if 'SMTP' == ae.Type:
            email_address = ae.Address
        return(email_address)
    except:
        return('')
def mailer(recipientlist,cclist,subject,bodytxt,**kwargs):
    df=kwargs.get('df')
    obj=kwargs.get('obj')
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipientlist
    mail.Cc = cc
    mail.Subject = subject
    greeting="Good Day,\n"
    txt=bodytxt+'\n'
    signoff="\nBest Regards,\nShane Ayers\nActing Human Resources Information Systems Manager\
    \nOffice of Human Resources\nYork College\nThe City University of New York"
    if df:
        tables='\n{table}\n'
    else:
        tables=''
    text=greeting+txt+tables+signoff
    hgreeting="<html><head><style> \
    table, th, td {{ border: 1px solid black; border-collapse: collapse; }} \
    th, td {{ padding: 10px; }}</style></head><body><p>Good Day,</p><p></p>"
    htxt="<p>"+bodytxt+'</p>'
    hsignoff="<p>Best Regards,</p><p>Shane Ayers</p>\
    <p>Acting Human Resources Information Systems Manager</p>\
    <p>Office of Human Resources</p><p>York College</p>\
    <p>The City University of New York</p></body></html>"
    if df:
        htables='<p></p>{table}<p></p>'
    else:
        htables=''
    html=hgreeting+htxt+htables+hsignoff
    if df:
        text = text.format(table=tabulate(df, headers=(list(df.columns.values)), tablefmt="grid"))
        html = html.format(table=tabulate(df, headers=(list(df.columns.values)), tablefmt="html"))
    mail.Body = text
    mail.HTMLBody = html
    if obj!='':
        mail.Attachments.Add(obj)
    mail.Send()
