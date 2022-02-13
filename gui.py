import threading
import tkinter as tk
from tkinter import filedialog
import pandas
import main

xlsx = ""
text = ""


def insert_to_list():
    log = (None, None)
    prev = ()
    while log[0] != 'END':
        log = main.cmsg
        if prev != log:
            prev = log
            s = str(log[0]) + str(log[1])
            list1.insert(tk.END, s)

    list1.insert(tk.END,"Finished sending")


def get_selected_row(event):
    try:
        global selected_tuple
        index = list1.curselection()[0]
        selected_tuple = list1.get(index)
        # TODO: add code
    except:
        pass


def send():
    btn_start.configure(state="disabled")

    if xlsx and text:
        with open(text, 'r') as f:
            msg_data = f.read()
        data = pandas.read_excel(xlsx, sheet_name="Sheet1")
        numbers = data["Numbers"].to_list()
        print(xlsx, text)

        t1 = threading.Thread(target=insert_to_list)
        t1.start()
        obj = main.Automate(numbers)
        obj.send("textdata", msg_data)
    else:
        print("empty")


def BrowseExcelFile():
    global xlsx
    filename = filedialog.askopenfilename(title="Select a File")
    xlsx = filename
    list1.insert(tk.END, ("Selected file", filename))


def BrowseTextFile():
    global text
    filename = filedialog.askopenfilename(title="Select a File")
    text = filename
    list1.insert(tk.END, ("Selected file", filename))


window = tk.Tk()
window.title("Whatsapp automation")
window.rowconfigure(0, minsize=800, weight=1)
window.columnconfigure(1, minsize=800, weight=1)

list1 = tk.Listbox(window)
fr_buttons = tk.Frame(window, relief=tk.RAISED, bd=2)
btn_excel = tk.Button(fr_buttons, text="Open Excel file", command=BrowseExcelFile)
btn_text = tk.Button(fr_buttons, text="Open text file", command=BrowseTextFile)
btn_start = tk.Button(fr_buttons, text="Start", command=lambda: threading.Thread(target=send, daemon=True).start())
btn_close = tk.Button(fr_buttons, text="close", command=window.destroy)
sb = tk.Scrollbar(window)

btn_excel.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
btn_text.grid(row=1, column=0, sticky="ew", padx=5)
btn_start.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
btn_close.grid(row=3, column=0, sticky="ew", padx=5)
sb.grid(row=0, column=2, rowspan=6, sticky="nse")

fr_buttons.grid(row=0, column=0, sticky="ns")
list1.grid(row=0, column=1, sticky="nsew")

list1.configure(yscrollcommand=sb.set)
sb.configure(command=list1.yview)

list1.bind('<<ListboxSelect>>', get_selected_row)

window.mainloop()
