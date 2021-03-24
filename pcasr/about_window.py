import tkinter as tk
from tkinter import ttk

class aboutWindow:

    def __init__(self,parent):

        # save parent for further use
        self.the_parent = parent

        self.about_window = tk.Tk()
        self.about_window.title("About")
        self.about_window.geometry('+%d+%d'%(parent.toplevel.winfo_x(),parent.toplevel.winfo_y()))

        # Tell the window manager to give focus back after the X button is hit
        self.about_window.protocol("WM_DELETE_WINDOW", self.kill_window)

        # Set this window to appear on top of other windows
        self.about_window.attributes("-topmost",True)

        self.version_label = ttk.Label(self.about_window,text="v0.1.1.0")
        self.copyright_label = ttk.Label(self.about_window,text="Copyright © 2020 - 2021 Factor Systems Inc. Dba BillTrust")
        self.author_label = ttk.Label(self.about_window,text="Written By Matthew DeGenaro")

        self.version_label.pack()
        self.copyright_label.pack()
        self.author_label.pack()

    def kill_window(self):
        self.the_parent.toplevel.wm_attributes("-disabled",False)
        self.about_window.destroy()
