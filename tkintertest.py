import tkinter
from tkinter import *
from tkinter import ttk
from tkinter.ttk import *
from tkinter import filedialog
from tkinter import messagebox

from PIL import ImageTk, Image

import storage
import os
import csv

import pandas as pd

# Import spreadsheet
def import_xl():  
    try:
        excelPath = os.path.normpath(filedialog.askopenfilename(title='Select File'))    
    except:
        messagebox.showerror('Error', 'File was not sucessfully chosen. Please try again.')
    try:
        storage.xl_db(excelPath)
    except:
        messagebox.showerror('Error', 'File was not successfully imported. Please try again.')

# Export spreadsheet
def export_xl(result):
    df = pd.DataFrame(result)
    df.to_excel('cs_feedback.xlsx', sheet_name='CS Feedback', index=False, header=False)
    messagebox.showinfo('Success', 'Your spreadsheet was exported to cs_feedback.xlsx!')

# View data
def list_entries():
    og_result = storage.get_emails()
    result = og_result[-5:]
    fixed_result = []
    
    for index, entry in enumerate(result):
        num = 0
        for item in entry:
            item = str(item)
            char_list = [item[j] for j in range(len(item)) if ord(item[j]) in range(65536)]
            item_fix=''.join(char_list)            
            lookup_label = Label(data, text=item_fix, wraplength=100, justify=LEFT)
            lookup_label.grid(row=index, column=num)
            num+=1
            
    return og_result

# Query data
def search():
    return

master_window=Tk()
master_window.title('AZ Email Analysis')
master_window.iconbitmap('images/azlogo.ico')

nav = LabelFrame(master_window, text='Actions', padding=10)
nav.pack(anchor=W)

data = LabelFrame(master_window, text='Data', padding=10)
data.pack(anchor=W)

result = list_entries()

#Label(master_window, text="Welcome to AZ Email Analysis!", font=("Arial Bold", 20)).pack(anchor=CENTER)
btn_sp=Button(nav, text="Import", command=import_xl)
btn_sp.grid(row=0, column=0)

btn_csv = Button(nav, text="Export", command = lambda: export_xl(result))
btn_csv.grid(row=0, column=1)



master_window.geometry("1000x400")
master_window.mainloop()