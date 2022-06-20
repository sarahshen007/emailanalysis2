from tkinter import *
window=Tk()
window.title('AZ Email Analysis')
btn=Button(window, text="This is Button widget", fg='red')
btn.place(x=350, y=350)
lbl=Label(window, text="This is Label widget", fg='red', font=("Helvetica", 16))
lbl.place(x=350, y=250)
window.geometry("700x700+10+10")
window.mainloop()