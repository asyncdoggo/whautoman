import tkinter as tk

import pandas

import main
import threading
from tkinter import filedialog


def send():
    b3.configure(state="disabled")
    xlsx = excel_file.cget("text")
    text = text_file.cget("text")

    if xlsx and text:
        with open(text,'r') as f:
            msg_data = f.read()

        data = pandas.read_excel(xlsx, sheet_name="Sheet1")
        numbers = data["Numbers"].to_list()

        obj = main.Automate(numbers)
        obj.send("textdata",msg_data)
    else:
        print("empty")




def BrowseExcelFile():
    filename = filedialog.askopenfilename(title="Select a File")
    excel_file.configure(text=filename)


def BrowseTextFile():
    filename = filedialog.askopenfilename(title="Select a File")
    text_file.configure(text=filename)


window = tk.Tk()
l1 = tk.Label(window, text="excel file")
l1.grid(row=0, column=0)

l3 = tk.Label(window, text="text file")
l3.grid(row=1, column=0)

b1 = tk.Button(window, text="select file", width=12, command=BrowseExcelFile)
b1.grid(row=0, column=2)
b2 = tk.Button(window, text="select file", width=12, command=BrowseTextFile)
b2.grid(row=1, column=2)

excel_file = tk.Label(window, text="", width=100, height=4, fg="blue")
excel_file.grid(row=0, column=1)
text_file = tk.Label(window, text="", width=100, height=4, fg="blue")
text_file.grid(row=1, column=1)

b3 = tk.Button(window, text="start", width=12, command=lambda: threading.Thread(target=send,daemon=True).start())
b3.grid(row=1, column=3)

list1 = tk.Listbox(window, height=6, width=35)
list1.grid(row=2, column=0, rowspan=6, columnspan=2)

sb1 = tk.Scrollbar(window)
sb1.grid(row=2, column=2, rowspan=6)

list1.configure(yscrollcommand=sb1.set)
sb1.configure(command=list1.yview)

# list1.bind('<<ListboxSelect>>', "get_selected_row")

b6 = tk.Button(window, text="Close", width=12, command=window.destroy)
b6.grid(row=7, column=3)

window.mainloop()
