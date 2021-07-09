#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = "Denise Mao"

import win32com.client as win32
import pandas as pd
import pyodbc
import datetime
import shutil
import os
import glob
from pathlib import Path


# Start timer
start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
start = datetime.datetime.strptime(start_time, '%H:%M:%S')

# Retrive Reporting Email List
# Create the MSSQL connection with python
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=Server Name;'
                      'Database=Database;'
                      'Trusted_Connection=yes;')

# Set the cursor for conn
cursor = conn.cursor()

# Fetching data in dataframe
email_list = pd.read_sql_query('''SQL Query Statement''', conn)
                                          
            
cursor.close()
del cursor
conn.close()


# Create path strings for report files folders
# Set date format
today_string = datetime.date.today().strftime('%d/%m/%Y')

# Set folder targets for attachments and archiving
attachment_path = Path.cwd() / 'attachments'
archive_dir = Path.cwd() / 'archive'

# Create a list to store report file path
attachments = []
for ind in range(len(email_list)):
    filename= email_list.loc[ind, 'Report_Name'] + ".xlsx"
    attachment = os.path.join(attachment_path, filename)
    attachments.append((email_list.loc[ind, 'Report_Name'], attachment))

# Create a dataframe incl. report file path
att_df = pd.DataFrame(attachments, columns=['Report_Name', 'Report_File'])
combined = pd.merge(email_list, att_df, on='Report_Name', how='left')

# Separate df combined into two subsets with BACKORDER and CS CHECK respectively and reset index
combined_bi, combined_csk = combined[(mask:=combined['Report_Name'].str.contains("BACKORDER"))].copy().reset_index(drop=True), combined[~mask].copy().reset_index(drop=True)


## Retrive Daily Report

# Create the MSSQL connection with python
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=Server Name;'
                      'Database=Database;'
                      'Trusted_Connection=yes;')

# Set the cursor for conn
cursor = conn.cursor()

# Generate BACKORDER Report
for ind in range(len(combined_bi)):
    param = combined_bi.loc[ind, 'Report_Name'] 
    write_path = combined_bi.loc[ind, 'Report_File']
    
    # Execute the stored srocedure with parameters
    storedProc_inv = "Exec stored_procedure_inv @ReportName=" + "'" + param + "'"
    storedProc_oo = "Exec stored_procedure_oo @ReportName=" + "'" + param + "'"    
    storedProc_pk = "Exec stored_procedure_pk @ReportName=" + "'" + param + "'"
    
    # Fetching data in dataframe    
    report_inv = pd.read_sql_query(storedProc_inv, conn)
    report_oo = pd.read_sql_query(storedProc_oo, conn)
    report_pk = pd.read_sql_query(storedProc_pk, conn)
        
    # Specify an ExcelWriter object to write to more than one sheet in the workbook
    with pd.ExcelWriter(write_path) as writer:
        report_inv.to_excel(writer, sheet_name='Invoice Report', index=False)
        report_oo.to_excel(writer, sheet_name='Open Orders', index=False)        
        report_pk.to_excel(writer, sheet_name='In Pick Report', index=False)

# Generate CS CHECK Report
for ind in range(len(combined_csk)):
    param = combined_csk.loc[ind, 'Report_Name'] 
    write_path = combined_csk.loc[ind, 'Report_File']
    
    # Execute the stored srocedure with parameters
    storedProc_csk = "Exec stored_procedure_csk @ReportName=" + "'" + param + "'"
    
    # Fetching data in dataframe
    report_csk = pd.read_sql_query(storedProc_csk, conn)
    
    # Specify an ExcelWriter object to write to more than one sheet in the workbook
    with pd.ExcelWriter(write_path) as writer:  
        report_csk.to_excel(writer, sheet_name='CS Check Report', index=False)
        

cursor.close()
del cursor
conn.close()


## Email reports
class EmailsSender:
    def __init__(self):
        self.outlook = win32.Dispatch('Outlook.Application')

    def send_email(self, report_name, receiver_name, to_email_address, attachment_path):
        # choose sender account
        send_account = None
        for account in self.outlook.Session.Accounts:
            if account.DisplayName == 'sender@email.com':
                send_account = account
            break
                
        mail = self.outlook.CreateItem(0)    # 0: olMailItem
        mail.To = to_email_address           # or, mail.Recipients.Add(to_email_address)
        ##mail.CC = 'Sales.Operations@blackwoods.com.au'
        mail.Subject = report_name + ' ' + today_string
        mail.HTMLBody = '''
                        <p>Hi {},</p>

                        <p>Please find the daily report.</p>

                        <p>For any question please reach out to contacts@email.com</p>                    
                        '''.format(receiver_name)
        
        with open(attachment_path, 'r', encoding='utf8', errors='ignore') as my_attch:
            myfile = my_attch.read()
        mail.Attachments.Add(attachment_path)
        
        # Use this to show the email
        #mail.Display(True)
        
        # Uncomment to send
        mail.Send()
        
        
# Send emails
email_sender_bi = EmailsSender()
for index, row in combined_bi.iterrows():
    email_sender_bi.send_email(row['Report_Name'], row['FirstName'], row['Email'], row['Report_File'])

email_sender_csk = EmailsSender()
for index, row in combined_csk.iterrows():
    email_sender_csk.send_email(row['Report_Name'], row['FirstName'], row['Email'], row['Report_File'])


# Empty archive locaion before moving reports
if os.listdir(archive_dir):
    files = glob.glob(os.path.join(archive_dir,'*'))
    for f in files:
        os.remove(f)
else:
    pass

# Move the files to the archive location
for f in attachments:
    shutil.move(f[1], archive_dir)
    

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
end = datetime.datetime.strptime(end_time,'%H:%M:%S')
elapsed = end - start

print("Total Runnig Time: " + str(elapsed))


