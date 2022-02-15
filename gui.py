import threading
import tkinter as tk
from tkinter import filedialog
import pandas
import main

xlsx = ""
text = ""
image = ""

not_found = []


def insert_to_list():
    log = (None, None)
    prev = ()
    while log[0] != 'END':
        log = main.cmsg
        if prev != log:
            prev = log
            s = str(log[0]) + str(log[1])
            list1.insert(tk.END, s)
            if "not" in log[0]:
                not_found.append(log[1])

    list1.insert(tk.END, "Finished sending")

    with open("not_found.txt", 'w') as file:
        for i in not_found:
            file.write(f"{i}\n")

    list1.insert(tk.END, "Numbers that were not found have been put in the file \"not_found.txt\"")
    btn_start.configure(state="normal")


def get_selected_row(event):
    try:
        global selected_tuple
        index = list1.curselection()[0]
        selected_tuple = list1.get(index)
        # TODO: add code
    except:
        pass


def send():
    t1 = threading.Thread(target=insert_to_list)
    if xlsx:
        excel_data = pandas.read_excel(xlsx, sheet_name="Sheet1")
        numbers = excel_data["Numbers"].to_list()
        if text:
            with open(text, 'r') as f:
                data = f.read()
            t1.start()
            btn_start.configure(state="disabled")
            obj = main.Automate(numbers)
            obj.send('TEXT', data)

        elif image:
            data = image
            t1.start()
            btn_start.configure(state="disabled")
            obj = main.Automate(numbers)
            obj.send('IMAGE', data)

        else:
            list1.insert(tk.END, "Please select an image or text file to send")
    else:
        list1.insert(tk.END, "Please select an excel file")


# TODO: add file type assertion to browse
def browse_excel():
    global xlsx
    xlsx = filedialog.askopenfilename(title="Select a File")

    if xlsx:
        list1.insert(tk.END, f"Selected excel file {xlsx}")


def browse_text():
    global text
    global image
    text = filedialog.askopenfilename(title="Select a File")

    if text:
        list1.insert(tk.END, f"Selected text file {text}")
        if image:
            list1.insert(tk.END, "Text file was selected, unselecting Image Folder")
            image = ""


def browse_img():
    global image
    global text
    image = filedialog.askdirectory(title="Select a Folder")

    image = image.replace('/', '\\')

    if image:
        list1.insert(tk.END, f"Selected image/video Folder {image}")
        if text:
            list1.insert(tk.END, "Image Folder was selected, unselecting Text file")
            text = ""


window = tk.Tk()
window.title("Whatsapp automation")
window.rowconfigure(0, minsize=800, weight=1)
window.columnconfigure(1, minsize=800, weight=1)

list1 = tk.Listbox(window)
fr_buttons = tk.Frame(window, relief=tk.RAISED, bd=2)
btn_excel = tk.Button(fr_buttons, text="Open Excel file", command=browse_excel)
btn_text = tk.Button(fr_buttons, text="Open text file", command=browse_text)
btn_start = tk.Button(fr_buttons, text="Start", command=lambda: threading.Thread(target=send, daemon=True).start())
btn_close = tk.Button(fr_buttons, text="close", command=window.destroy)
btn_img = tk.Button(fr_buttons, text="image/video", command=browse_img)

sb = tk.Scrollbar(window)

btn_excel.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
btn_text.grid(row=1, column=0, sticky="ew", padx=5)
btn_img.grid(row=2, column=0, sticky="ew", padx=5)
btn_start.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
btn_close.grid(row=4, column=0, sticky="ew", padx=5)
sb.grid(row=0, column=2, rowspan=6, sticky="nse")

fr_buttons.grid(row=0, column=0, sticky="ns")
list1.grid(row=0, column=1, sticky="nsew")

list1.configure(yscrollcommand=sb.set)
sb.configure(command=list1.yview)

list1.bind('<<ListboxSelect>>', get_selected_row)

window.mainloop()
