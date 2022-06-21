import tkinter
from tkinter import *
from tkinter import ttk
from tkinter.ttk import *
from tkinter import filedialog
from tkinter import messagebox

import datetime
import emails
import storage
import win32com
from win32com import client
from bs4 import BeautifulSoup
import os

import pandas as pd

# Import spreadsheet
def import_xl():  
    try:
        excelPath = os.path.normpath(filedialog.askopenfilename(title='Select File'))    
    except:
        messagebox.showerror('Error', 'File was not sucessfully chosen. Please try again.')
    try:
        storage.xl_db(excelPath)
        messagebox.showinfo('Success! You imported you spreadsheet into the database.')
    except:
        messagebox.showerror('Error', 'File was not successfully imported. Please try again.')
    update_treeview()

# Export spreadsheet
def export_xl():
    result = list_entries()
    df = pd.DataFrame(result)
    df.to_excel('cs_feedback.xlsx', sheet_name='CS Feedback', index=False, header=False)
    messagebox.showinfo('Success', 'Your spreadsheet was exported to cs_feedback.xlsx!')

# Import emails
def import_emails():
    # Get the date
    d = datetime.datetime.strptime(storage.get_last_date(), '%m/%d/%Y').date()
    
    # Connect to outlook account
    app = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts
    account = accounts[0]

    # Grab folder for CS EMAILS
    root_folder = app.Folders(account.DisplayName)
    emails_folder = emails.get_folder_by_name("CS EMAILS", root_folder)

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
        date = datetime.datetime.strptime(date, '%d-%m-%y').date()

        # Get Subject Line of email
        sjl = msg.Subject

        # Only add emails since last update           
        if d <= date:
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
    storage.add_emails(emails_log)
    update_treeview()

# View data
def list_entries():
    result = storage.get_emails()
    return result

# update treeview
def update_treeview():
    results = list_entries()
    for item in my_tree.get_children():
        my_tree.delete(item)
    num = 0
    for result in results:
        my_tree.insert(parent='', index='end', iid=num, text='', values=result)
        num+=1

master_window=Tk()
master_window.title('AZ Email Analysis')
master_window.iconbitmap('images/azlogo.ico')
master_window.geometry("1000x400")

excel = LabelFrame(master_window, text='Excel', padding=10)
excel.pack(anchor=CENTER)

outlook = LabelFrame(master_window, text='Outlook', padding=10)
outlook.pack(anchor=CENTER)

data = LabelFrame(master_window, text='Data', padding=10)
data.pack(anchor=CENTER)

my_tree_scroll_y = Scrollbar(data)
my_tree_scroll_y.pack(side=RIGHT, fill=Y)
my_tree_scroll_x = Scrollbar(data, orient='horizontal')
my_tree_scroll_x.pack(side=BOTTOM, fill=X)

my_tree = ttk.Treeview(data, yscrollcommand=my_tree_scroll_y.set, xscrollcommand=my_tree_scroll_x.set)
my_tree.pack()

my_tree_scroll_y.config(command=my_tree.yview)
my_tree_scroll_x.config(command=my_tree.xview)

my_tree['columns'] = ('Date', 'Issue', 'Product', 'Name', 'Email', 'Comment', 'Session', 'Followup')
my_tree.column("#0", width=0, minwidth=0)
my_tree.column('Date', anchor=W, width=100)
my_tree.column('Issue', anchor=W, width=100)
my_tree.column('Product', anchor=W, width=100)
my_tree.column('Name', anchor=W, width=100)
my_tree.column('Email', anchor=W, width=100)
my_tree.column('Comment', anchor=W, width=150)
my_tree.column('Session', anchor=W, width=100)
my_tree.column('Followup', anchor=W, width=80)

my_tree.heading('#0', text="", anchor=W)
my_tree.heading('Date', text='Date', anchor=W)
my_tree.heading('Issue', text='Issue', anchor=W)
my_tree.heading('Product', text='Product', anchor=W)
my_tree.heading('Name', text='Name', anchor=W)
my_tree.heading('Email', text='Email', anchor=W)
my_tree.heading('Comment', text='Comment', anchor=W)
my_tree.heading('Session', text='Session', anchor=W)
my_tree.heading('Followup', text='Followup', anchor=W)

update_treeview()

btn_sp=Button(excel, text="Import", command=import_xl)
btn_sp.grid(row=0, column=0)

btn_csv = Button(excel, text="Export", command = lambda: export_xl())
btn_csv.grid(row=0, column=1)

btn_outlook=Button(outlook, text="Retrieve emails", command = lambda: import_emails())
btn_outlook.pack(anchor=CENTER)

master_window.mainloop()