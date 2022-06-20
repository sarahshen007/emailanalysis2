import tkinter
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from tkinter import messagebox

from PIL import ImageTk, Image

from functools import partial

import storage
import os

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

# View data
def query(filters):
    emails_list = storage.get_emails(filters)

master_window=Tk()
master_window.title('AZ Email Analysis')
master_window.iconbitmap('images/azlogo.ico')
style = Style()

style.configure('W.TButton', background='#345', foreground='black', font=('Arial', 14 ))
Label(master_window, text="Welcome to AZ Email Analysis!", font=("Arial Bold", 20)).pack(anchor=CENTER)
btn_sp=Button(master_window, text="Import", style = 'W.TButton', command=import_xl)
btn_sp.pack(anchor=CENTER)



master_window.geometry("600x300+10+10")
master_window.mainloop()