# Import the required libraries
from tkinter import *
from tkinter import ttk
import tkinter as tk
import datetime
import time
from datetime import timedelta
import json
import os

today_date = str(datetime.datetime.now().date()).replace("-","_")

# Create an instance of tkinter frame
win= Tk()

# Set the size of the tkinter window
win.geometry("375x150")
s = ttk.Style()
s.theme_use('clam')


# Configure the style of Heading in Treeview widget
s.configure('Treeview.Heading', background="#003035", foreground='#ffffff')

# Add a Treeview widget
tree= ttk.Treeview(win, column=("c1", "c2","c3"), show= 'headings', height= 8)

tree.column("# 1",anchor=CENTER, width=95)
tree.heading("# 1", text= "PCASE")
tree.column("# 2", anchor= CENTER, width=160)
tree.heading("# 2", text= "CSRNAME")
tree.column("# 3", anchor= CENTER, width=95)
tree.heading("# 3", text= "TOTAL TIME")

user = os.getlogin()
data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
tracker_file =data_folder + "\\nt-json-files\\ttimes_stored.json"
total_times = data_folder + "\\nt-json-files\\total_times.json"

def load_times():
    with open(total_times) as json_file:
        times_dict=  json.load(json_file)
    return times_dict


times = load_times()
cases = times[today_date]
rowCount = 1
for case in cases:
    captureRowCount = rowCount
    pcase_csrname = case.split("_")
    captureRowCount = rowCount
    tree.insert('', 'end',text= str(rowCount),values=(pcase_csrname[0],pcase_csrname[1], cases[case]))

    rowCount += 1



tree_scroll = tk.Scrollbar(win,command=tree.yview)
tree['yscrollcommand'] = tree_scroll.set
tree_scroll.pack(side=tk.RIGHT,fill = tk.Y)

tree.pack()

win.mainloop()