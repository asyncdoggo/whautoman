import threading
import tkinter as tk
from tkinter import filedialog
import pandas
import main

xlsx = ""
text = ""
images = []
not_found = []
filetype = ".apng .avif .gif .jpg, .jpeg .jfif .pjpeg .pjp .png .svg .webp .bmp .ico .cur .tif .tiff .mp4 .mov .wmv " \
           ".avi .avchd .mkv .webm"


def insert_to_list() -> None:
    log = ""
    prev = ()
    while log != 'END':
        log = main.cmsg
        if prev != log:
            prev = log
            list1.insert(tk.END, log)
            if "not" in log:
                not_found.append(log)

    list1.insert(tk.END, "Finished sending")

    with open("not_found.txt", 'w') as file:
        for i in not_found:
            file.write(f"{i}\n")

    list1.insert(tk.END, "Numbers that were not found have been put in the file \"not_found.txt\"")
    btn_start.configure(state="normal")


def send() -> None:
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

        elif images:
            data = images
            t1.start()
            btn_start.configure(state="disabled")
            obj = main.Automate(numbers)
            obj.send('IMAGE', data)

        else:
            list1.insert(tk.END, "Please select an image or text file to send")
    else:
        list1.insert(tk.END, "Please select an excel file")


def browse_excel() -> None:
    global xlsx
    xlsx = filedialog.askopenfilename(title="Select a Excel File", filetypes=[("Excel file", ".xlsx .xls")])

    if xlsx:
        list1.insert(tk.END, f"Selected excel file {xlsx}")


def browse_text() -> None:
    global text
    global images
    text = filedialog.askopenfilename(title="Select a text File", filetypes=[("Text file", ".txt")])

    if text:
        list1.insert(tk.END, f"Selected text file {text}")
        if images:
            list1.insert(tk.END, "Text file was selected, unselecting Image Files")
            images = ()


def browse_img() -> None:
    global images
    global text
    images = list(
        filedialog.askopenfilenames(title="Select Images/videos", filetypes=[("Images and Videos", filetype)]))

    if images:
        for i in range(len(images)):
            images[i] = images[i].replace("/", "\\")
            list1.insert(tk.END, f"Selected file {images[i]}")
        if text:
            list1.insert(tk.END, "Image/video Files was selected, unselecting Text file")
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

window.mainloop()
