# Program to complete email analysis
import os
import datetime
import sys

import emails
import storage

import win32com
from win32com import client
import openpyxl
from tkinter import filedialog
from tkinter import *
from bs4 import BeautifulSoup
from colorama import Fore


# Get the date
d = (datetime.date.today() - datetime.timedelta (days=1)).strftime("%d-%m-%y")

# Connect to outlook account
app = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts
account = accounts[0]
print("Account:", account.DisplayName)

# Grab folder for CS EMAILS
root_folder = app.Folders(account.DisplayName)
emails_folder = emails.get_folder_by_name("CS EMAILS", root_folder)

# Grab spreadsheet path
excelPath = os.path.normpath(filedialog.askopenfilename(title='Select File'))
storage.xl_db(excelPath)

# Get emails from folder and populate list of emails
messages = emails_folder.Items

# Get the email/s
msg = messages.GetLast()

# List of emails from day
new_messages = []

# Log keeping track of email objects
emails_log = []

# Parse spreadsheet to predict issue
prev_data = storage.generate_issue_data()

# Loop through emails
while msg:
    # Get email date 
    date = msg.SentOn.strftime("%d-%m-%y")

    # Get Subject Line of email
    sjl = msg.Subject

    # Only add emails from yesterday           
    if d == date:
        new_messages.append(msg)

        # Dictionary to store message info
        info = {}

        # Date message was received
        date = str(msg.SentOn).split(' ')[0]
        date = datetime.datetime.strptime(date, '%Y-%m-%d').date()
        info['date'] = date

        # Remove unecessary characters from msg html
        regex = msg.HTMLBody.replace('\r', '').replace('\n', '') 

        # Parse into html using soup
        soup = BeautifulSoup(regex, "html.parser") 

        # Create list of category + values
        texts = str(soup.find_all('font')[0].encode_contents(encoding='utf-8')).strip('b').strip('\'').strip('\"').replace('<br/>', '\n')
        texts = emails.replaceCharacters(texts)
        texts = texts.strip().split('\n')
        texts = list(filter(None, texts))

        # Create list of pairs to populate info dictionary
        pairs = []
        
        # Edit list for unwanted extra elements caused by extra break elements
        lastKey = ""
        for data in texts:
            pair = data.split(':', 1)
            if len(pair) == 1:
                info[lastKey] = info[lastKey] + pair[0]
            elif len(pair) == 2: 
                lastKey = pair[0].strip()
                info[lastKey] = pair[1].strip()

        # Generate summary of comment  
        comment = info['Comment Value']
        char_list = [comment[j] for j in range(len(comment)) if ord(comment[j]) in range(65536)]
        comment_fix=''.join(char_list)    
        info['Comment Value'] = comment_fix

        predicted_issue = storage.generate_issue(comment, prev_data)

        info['Issue Summary'] = predicted_issue[0]
        info['Product'] = predicted_issue[1]
        
        # Make new email object with info
        newEmail = emails.emailCreator(info)

        # Add email object to emails log
        emails_log.append(newEmail)
        
    msg = messages.GetPrevious()


# Add new entries to spreadsheet
print("Adding to log to spreadsheet...\n")
wb = openpyxl.load_workbook(excelPath) 
print("Spreadsheet opened successfully.")

sheet = wb.active 

data = []

for email in emails_log:
    row = (email.date, email.issueSummary, email.product, email.name, email.customerEmail, email.comment, email.ipAddress, email.cookies, email.followup)
    data.append(row)

for row in data:
    sheet.append(row)

wb.save(excelPath)
wb.close()

print("=========")
print("All done! Thank you for using AZ Email Analysis!")
print("=========")
