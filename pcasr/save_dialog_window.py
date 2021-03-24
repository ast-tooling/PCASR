import tkinter as tk
import subprocess
from tkinter import messagebox, filedialog, RIGHT, RAISED, Listbox, END, MULTIPLE, TOP, BOTTOM, ttk, Frame, Label, Text, Scrollbar, Y,X, Message, Button, Menu, Entry, DISABLED,ACTIVE, BOTH
import webbrowser
import validators
import os
from shutil import copyfile
import time
import threading
import queue
import json
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
from simple_salesforce import Salesforce
import configparser


class saveDialogWindow:

    def __init__(self,parent,edit):

        # save parent for further use
        self.the_parent = parent

        # Disable parent window
        self.the_parent.toplevel.wm_attributes("-disabled", True)

        # Set Theme for Window
        self.style=ttk.Style()
        self.style.theme_use('vista')

        # Initialize tk instance and frame
        self.save_window = tk.Tk()
        self.save_window.title("New Entry")
        self.save_frame = Frame(self.save_window,borderwidth=1)

        # Control Location of Appearence
        self.save_window.geometry('+%d+%d'%(parent.toplevel.winfo_x(),parent.toplevel.winfo_y()))

        # Tell the window manager to give focus back after the X button is hit
        self.save_window.protocol("WM_DELETE_WINDOW", self.kill_window)

        # Set this window to appear on top of all other windows
        self.save_window.attributes("-topmost",True)

        # Add Frame to Window via Grid
        self.save_frame.pack()

        # Initialize Labels
        self.srd_link = Label(self.save_frame,text="SRD Link")
        self.sf_link = Label(self.save_frame,text="SF Link")

        # Initialize Entries
        self.srd_entry = Entry(self.save_frame,validate="focusout",validatecommand=self.validateSRD,width=50)
        self.sf_entry = Entry(self.save_frame,validate="focusout",validatecommand=self.validateSF,width=50)

        # ✔ ❌ ❓
        # Initialize Verify Labels, Label Messages, and Tooltips
        self.srd_valid = ttk.Label(self.save_frame,text="❓",width=3,foreground="#fcba03")
        self.sf_valid = ttk.Label(self.save_frame,text="❓",width=3,foreground="#fcba03")

        self.srd_tt = CreateToolTip(self.srd_valid,"SRD will be validated when text is entered and the mouse leaves this box.")
        self.sf_tt = CreateToolTip(self.srd_valid,"SalesForce link will be validated when text is entered and the mouse leaves this box.")


        # Initialize Command Buttons
        self.save_button = ttk.Button(self.save_window,text="Save", command=self.savePCase,state=DISABLED)
        self.cancel_button = ttk.Button(self.save_window,text="Cancel",command=self.kill_window)
        self.validate_button = ttk.Button(self.save_window,text="Validate")

        # Pack Command Buttons to Window
        self.save_button.pack(side="right", padx=5, pady=5)
        self.cancel_button.pack(side="right",padx=5)
        self.validate_button.pack(side="left",padx=5)

        # Configure Grid Layout
        self.sf_link.grid(row=2,column=0)

        self.sf_entry.grid(row=2,column=1)

        self.sf_valid.grid(row=2,column=2,padx=2,pady=2)

        self.srd_link.grid(row=3,column=0)

        self.srd_entry.grid(row=3,column=1)

        self.srd_valid.grid(row=3,column=2,padx=2,pady=2)


        # Initialize Data File Location
        user = os.getlogin()
        self.data_folder = "C:\\Users\\%s\\AppData\\Roaming\\PCASR" %user
        self.data_file = self.data_folder+"\\pcasr.json"

        # If pcase already exists in data, load and validate
        if edit:
            data = self.the_parent.json_data
            pcase = self.the_parent.pcase_list.item(self.the_parent.pcase_list.selection())['values'][0]
            if pcase in data:
                self.srd_entry.insert(tk.END,data[pcase]['srd_link'])
                self.sf_entry.insert(tk.END,data[pcase]['sf_link'])

                self.validateSRD()
                self.validateSF()

    def savePCase(self):

        if not os.path.exists(self.data_file):
            if not os.path.isdir(self.data_folder):
                os.mkdir(self.data_folder)

            data = {} 
            # Initialize Data File
            with open(self.data_file,'w') as outfile:
                json.dump(data,outfile)

        sf_link = self.sf_entry.get().strip()

        srd_link = ""
        if(self.validateSRD()):
            srd_link = self.srd_entry.get()

        case_info = self.getSFInfo(sf_link)

        if case_info:
            case_info['srd_link'] = srd_link
            case_info['sf_link'] = sf_link
            case_info['notes'] = ''
            pcase = case_info['pcase']

            json_file = open(self.data_file)
            data = json.load(json_file)
            if pcase not in data:
                data[pcase] = case_info
            else:
                messagebox.showwarning('Error', 'You have already added this case.')
            json_file.close()
            with open(self.data_file,'w') as outfile:
                json.dump(data,outfile)

            self.the_parent.json_data = data
            self.kill_window()
            self.the_parent.loadPCases()
            self.the_parent.pcase_list.selection_set(pcase)

    def getSFInfo(self,sf_link):

        if not os.path.exists(self.data_folder+"\\config.txt"):
            messagebox.showwarning('Error', 'You must first add your sf credentials to\nC:\\Users\\<you>\\AppData\\Roaming\\PCASR\\credentials.txt\nAn example can be found at Z:\\AST\\Utilities\\PCASR')
            return False
        else:
            config = configparser.ConfigParser()
            config.read(self.data_folder+"\\config.txt")

            client = Salesforce(
                username=config.get('credentials','username'),
                password=config.get('credentials','password'),
                security_token=config.get('credentials','security_token')
                )
            case_id = sf_link.split('/')[-1]

            print(sf_link)
            print(sf_link.split('/'))
            print(case_id)
            case_info = client.Case.get(case_id)

            case_dict = {
                'subject':case_info['Subject'],
                'last_modified':case_info['LastModifiedDate'],
                'case_owner':case_info['Case_Owner__c'],
                'case_number':case_info['CaseNumber'],
                'parent_case_owner':case_info['Parent_Case_Owner__c'],
                'classification':case_info['Case_Classification__c'],
                'cust_name':case_info['CSR_Username_Cases__c'],
                'pcase':case_info['Z_Case_NoPath__c']
                }

            print(case_dict)
            return case_dict

    def validateSF(self):
        url = self.sf_entry.get()
        if "https://" not in url:
            url = "https://" + url 
        if validators.url(url):
            self.sf_valid.config(text="✔",foreground="#198000")
            self.sf_tt.setText("SF Link is a valid website")
            self.stateHandler()
            return True
        else:
            self.sf_valid.config(text="❌",foreground="#b00505") 
            self.sf_tt.setText("SF Link is not a valid website")
            self.stateHandler()
            return False

    def validateSRD(self):
        url = self.srd_entry.get()
        if "https://" not in url:
            url = "https://" + url 
        if validators.url(url):
            self.srd_valid.config(text="✔",foreground="#198000")
            self.srd_tt.setText("SRD is a valid website")
            return True
        else:
            self.srd_valid.config(text="❌",foreground="#b00505") 
            self.srd_tt.setText("SRD is not a valid website")
            return False

    # ✔ ❌ ❓
    def stateHandler(self):
        if self.sf_valid.cget('text') == "✔":
            self.save_button.config(state=ACTIVE)
        else:
            self.save_button.config(state=DISABLED)

    def kill_window(self):
        self.the_parent.toplevel.wm_attributes("-disabled",False)
        self.save_window.destroy()