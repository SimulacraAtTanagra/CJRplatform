# -*- coding: utf-8 -*-
"""
Created on Wed Jan 22 12:17:23 2020

@author: sayers
"""
from re import search
import win32com.client as win32
from tabulate import tabulate
def getemail(search_string):
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

def mailthis(recipientlist,cc, df, subject):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipientlist
    mail.Cc = cc
    mail.Subject = subject
    
    
    
    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)
    text = """
    Good Day,
    
    
    
    {table}
    
    Best Regards,
    Shane Ayers
    Acting Human Resources Information Systems Manager
    Office of Human Resources
    York College
    The City University of New York"""
    
    html = """
    <html>
    <head>
    <style>     
     table, th, td {{ border: 1px solid black; border-collapse: collapse; }}
      th, td {{ padding: 10px; }}
    </style>
    </head>
    <body><p>Good Day,</p>
    <p></p>
    {table}
    <p>Best Regards,</p>
    <p>Shane Ayers</p>
    <p>Acting Human Resources Information Systems Manager</p>
    <p>Office of Human Resources</p>
    <p>York College</p>
    <p>The City University of New York</p>
    </body></html>
    """
    
    # above line took every col inside csv as list
    text = text.format(table=tabulate(df, headers=(list(df.columns.values)), tablefmt="grid"))
    html = html.format(table=tabulate(df, headers=(list(df.columns.values)), tablefmt="html"))
    mail.Body = text
    mail.HTMLBody = html
    mail.Send()
    
if __name__ == "__main__":
    mailthis()