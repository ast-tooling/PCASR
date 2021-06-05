import tkinter as tk
from tkinter import ttk
import webbrowser

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

        self.version_label = ttk.Label(self.about_window,text="v0.1.4.0")
        self.git_label = tk.Label(self.about_window,text='https://github.com/ast-tooling/PCASR',fg='blue',cursor='hand2')
        self.copyright_label = ttk.Label(self.about_window,text="Copyright © 2020 - 2021 Factor Systems Inc. Dba BillTrust")
        self.author_label = ttk.Label(self.about_window,text="Orginal Source Code Written By: Matthew DeGenaro")
        self.maintainer_label = ttk.Label(self.about_window,text="Maintaince & Feature Updates By: Christopher Durham")
        
        self.git_label.bind("<Button-1>",lambda e: self.openWebsite('https://github.com/ast-tooling/PCASR'))

        self.version_label.pack()
        self.git_label.pack()
        self.copyright_label.pack()
        self.author_label.pack()
        self.maintainer_label.pack()
    
    def openWebsite(self,page):
        if page:
            webbrowser.get('chrome').open(page)
       


    def kill_window(self):
        self.the_parent.toplevel.wm_attributes("-disabled",False)
        self.about_window.destroy()
