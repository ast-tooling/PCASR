import tkinter as tk
import pygubu
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
from salesforce_api import Salesforce
import configparser

class watchDog:

    def __init__(self,parent,path):
        patterns = "*"
        ignore_patterns = ""
        ignore_directories = False
        case_sensitive = True
        self.parent = parent
        my_event_handler = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)

        def on_any_event(event):
            parent.ftpListUpdate()

        my_event_handler.on_any_event = on_any_event

        go_recursively = False
        my_observer = Observer()
        my_observer.schedule(my_event_handler, path, recursive=go_recursively)

        my_observer.start()
        while self.parent.keepWatching():
            time.sleep(1)
        my_observer.stop()
        my_observer.join()

    def stopWatching(self):
        self.watching = False

class CreateToolTip(object):
    """
    create a tooltip for a given widget
    credit to vegaseat on stackoverflow
    edited to include a mutator for text by mdegenaro
    """
    def __init__(self, widget, text='widget info'):
        self.waittime = 500     #miliseconds
        self.wraplength = 180   #pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def setText(self,text="Default"):
        self.text = text

    def showtip(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = tk.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw, text=self.text, justify='left',
                       background="#ffffff", relief='solid', borderwidth=1,
                       wraplength = self.wraplength)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw= None
        if tw:
            tw.destroy()

class copyFilesWindow:

    the_queue = queue.Queue()

    def __init__(self,parent,subFolder_path,subDir,files):

        # save parent for further use
        self.the_parent = parent

        self.subFolder_path = subFolder_path
        self.subDir = subDir
        self.files = files
        self.already_copied = False

        self.copy_window = tk.Tk()
        self.copy_window.title("Copying Files")
        self.copy_window.geometry('+%d+%d'%(parent.toplevel.winfo_x(),parent.toplevel.winfo_y()))

        # Tell the window manager to give focus back after the X button is hit
        self.copy_window.protocol("WM_DELETE_WINDOW", self.kill_window)

        # Set this window to appear on top of other windows
        self.copy_window.attributes("-topmost",True)


        self.textbox = tk.Text(master=self.copy_window)
        self.textbox.pack(side=TOP)
        self.start = tk.Button(master=self.copy_window,text="Copy Files",command=self.copy_status)
        self.start.pack(side=BOTTOM)
        

    # wrapper method for textbox.insert to handle state disabling
    def text_insert(self,text):
        self.textbox.configure(state='normal')
        self.textbox.insert(END, text)
        self.textbox.configure(state='disabled')
        self.textbox.see(tk.END)  

    # method called by Copy button, handles opening threads
    def copy_status(self):
        if not self.already_copied:
            self.already_copied = True
            threadlist = []
            for file in self.files:
                self.text_insert("Copying " + file + '\n')

                if self.subDir != "\\SAMPLE_DATA":
                    for server in self.the_parent.server_list:
                        self.text_insert("       Copying to "+ "\\\\"+server+ '\n')
                        thread = threading.Thread(target=self.copy_files, args=[self.subFolder_path+"\\"+file,"\\\\"+server+self.subDir+"\\"+file])
                        threadlist.append(thread)
                        thread.start()
                else:
                    self.text_insert("       Copying to "+ "FTP ROOT"+ '\n')
                    server = "\\\\ssnj-netapp01\\imtest\\imstage01\\ftproot\\"+self.the_parent.cust_info.cget('text')
                    thread = threading.Thread(target=self.copy_files, args=[self.subFolder_path+"\\"+file,server+"\\"+file])
                    threadlist.append(thread)
                    thread.start()
            thread = threading.Thread(target=self.done_check, args=[threadlist])
            thread.start()
            self.textbox.after(100,self.thread_check)
        else:
             messagebox.showwarning('Error', 'Files already copied.  If you want to copy again, close this window.')


    def copy_files(self,src,dst):
        copyfile(src,dst)

    # works in conjunction with thread_check to provide updates to progress
    def done_check(self,threadlist):
        print(threadlist)
        while True:
            if all(not thread.isAlive() for thread in threadlist):
                self.the_queue.put("Done")
                break

    # works in conjunction with done_check to provide updates to progress
    def thread_check(self):
        try:
            message = self.the_queue.get(block=False)
        except queue.Empty:
            self.textbox.after(100, self.thread_check)
            self.text_insert(".")
            return

        if message == "Done":
            self.text_insert("\nDone!")
        else:
            self.textbox.after(100,self.thread_check)

    def kill_window(self):
        self.the_parent.toplevel.wm_attributes("-disabled",False)
        self.copy_window.destroy()

class filePickerWindow:

    def __init__(self,parent,subFolder_path,files):
        # save parent for further use
        self.the_parent = parent

        self.picker_window = tk.Tk()
        self.picker_window.title("File Selector")
        picker_frame = tk.Frame(self.picker_window,relief=RAISED, borderwidth=1)
        self.toplevel =parent.toplevel
        self.infiles = files
        self.files = []
        self.picker_window.geometry('+%d+%d'%(parent.toplevel.winfo_x(),parent.toplevel.winfo_y()))

        # Set this window to appear on top of all other windows
        self.picker_window.attributes("-topmost",True)

        # Tell the window manager to give focus back after the X button is hit
        self.picker_window.protocol("WM_DELETE_WINDOW", self.kill_window)

        self.listbox = Listbox(self.picker_window,selectmode=MULTIPLE,width =50,height=len(files))
        self.listbox.pack()
        for index, file in enumerate(files):
            self.listbox.insert(END,file)

        ok_button = tk.Button(self.picker_window,text="Select",command=self.ok_button)
        cancel_button = tk.Button(self.picker_window,text="Cancel", command=self.kill_window)

        picker_frame.pack()
        ok_button.pack(side=RIGHT, padx=5, pady=5)  
        cancel_button.pack(side=RIGHT)

    def ok_button(self):
        selected =  self.listbox.curselection()
        for i in selected:
            self.files.append(self.infiles[i])
        self.kill_window()

    def show(self):
        self.picker_window.wait_window()
        return self.files

    def kill_window(self):
        self.the_parent.toplevel.wm_attributes("-disabled",False)
        self.picker_window.destroy()

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
        self.data_folder = "C:\\Users\\%s\\AppData\\Roaming\\PCASR_DEV" %user
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
            case_info = client.sobjects.Case.get(case_id)

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

class PCaser:

    webbrowser.register('chrome',None,webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
    webbrowser.register('firefox',None,webbrowser.BackgroundBrowser("C://Program Files //Mozilla Firefox//Chrome//firefox.exe"))


    def __init__(self):

        self.initMainFrame()

        self.all_servers = ["ssnj-imbisc10","ssnj-imbisc11","ssnj-imbisc12","ssnj-imbisc13","ssnj-imbisc20","ssnj-imbisc21"]
        self.json_data = []
        self.notes_hash = ""
        self.server_list = []

        self.first_init = True

        # Call GUI initialization methods
        self.initMenuBar()
        self.initInfoFrame()
        self.initQuickButtons()
        self.initFTPFrame()
        self.initTabArea()

        self.initTreeView()



        # Populate the ListBox with PCases
        self.loadPCases()


        self.watching = False
        # Select the first listbox item if there is one
        try:
            topSelect = self.pcase_list.identify_row(0)
            print(topSelect)
            self.pcase_list.selection_set(topSelect)
            self.updateInfo(self.pcase_list.item(topSelect)['values'])

            self.threadStart()
        except:
            pass


        # Tell the window manager to give focus back after the X button is hit
        self.toplevel.protocol("WM_DELETE_WINDOW", self.kill_window)


    def kill_window(self):
        self.setWatch(False)
        self.notes_hash = self.notes.get("1.0","end")
        self.saveNotes()
        self.toplevel.destroy()
        quit()

    def initFTPFrame(self):
        self.ftpList = Listbox(self.ftp_root_frame,height=6,width=65)
        self.ftpList.pack(fill=BOTH)

       
    def initMainFrame(self):
        self.mainwindow = tk.Tk()
        self.mainwindow.title("PCASR")
        self.mainwindow.resizable(False, False)
        self.toplevel = self.mainwindow.winfo_toplevel()
        self.main_frame = ttk.Labelframe(self.mainwindow)

        self.pcase_frame = ttk.Labelframe(self.main_frame,text="PCase")
        self.pcase_list_frame2 = ttk.Labelframe(self.main_frame,text="PCase List")

        self.info_area_frame = ttk.Labelframe(self.pcase_frame,text="Info")
        self.quick_button_frame = ttk.Labelframe(self.pcase_frame,text="Quick Links")
        self.ftp_root_frame = ttk.Labelframe(self.pcase_frame,text="FTP Root")
        self.tab_frame = ttk.Frame(self.pcase_frame)

        self.info_area_frame.grid_propagate(True)
        self.ftp_root_frame.grid_columnconfigure(0,pad=165)

        self.info_area_frame.grid(row=0,column=0)
        self.quick_button_frame.grid(row=1,column=0)
        self.ftp_root_frame.grid(row=2,column=0)
        self.tab_frame.grid(row=3,column=0)


        self.pcase_frame.grid(row=0,column=1)

        self.pcase_list_frame2.grid(row=0,column=0,padx=10)

        self.main_frame.grid(row=0,column=0,padx=5)

    def initTabArea(self):
        # Put a Space in Before The Tab Area for Visual Cleanliness
        self.space_label = Label(self.tab_frame)
        self.space_label.pack()

        # Remove Ugly Dotted Lines From Selected Tabs
        self.style = ttk.Style()
        self.style.layout("Tab",
        [('Notebook.tab', {'sticky': 'nswe', 'children':
            [('Notebook.padding', {'side': 'top', 'sticky': 'nswe', 'children':
                #[('Notebook.focus', {'side': 'top', 'sticky': 'nswe', 'children':
                    [('Notebook.label', {'side': 'top', 'sticky': ''})],
                #})],
            })],
        })]
        )

        # Tab Definitions
        self.tab_area = ttk.Notebook(self.tab_frame,width=400,height=380)

        self.tab_1 = Frame(self.tab_area)
        self.tab_2 = Frame(self.tab_area)
        self.tab_3 = Frame(self.tab_area)
        self.tab_4 = Frame(self.tab_area)

        self.tab_area.add(self.tab_1,text="Quick View")
        self.tab_area.add(self.tab_2,text="Notes")
        self.tab_area.add(self.tab_3,text="Sync/Push")
        self.tab_area.add(self.tab_4,text="Tab 4")
        self.tab_area.pack(expand=1,fill="both")


        # Tab 1 - Notes
        self.notes = Text(self.tab_2)
        self.note_scroll = Scrollbar(self.tab_2,command=self.notes.yview)
        self.notes['yscrollcommand'] = self.note_scroll.set

        # Initialize Hash Value
        self.notehash = ""
        self.note_scroll.pack(side=RIGHT,fill = Y)
        self.notes.pack()


        # Tab 3 - Probably the Sync/Push Area


        self.top_frame = ttk.Labelframe(self.tab_3)
        self.bot_frame = ttk.Labelframe(self.tab_3)

        self.top_frame.pack(ipady=5,ipadx=5,padx=5)
        self.bot_frame.pack(ipady=5,ipadx=5,padx=5,expand=True,fill='both')

        # Top Half

        # Buttons
        self.button1 = ttk.Button(self.top_frame, text="Edit",command=self.editTemplates)
        self.button2 = ttk.Button(self.top_frame, text="Edit",command=self.editParsers)
        self.button3 = ttk.Button(self.top_frame, text="Edit",command=self.editScripts)
        self.button4 = ttk.Button(self.top_frame, text="Edit",command=self.editSamples)
        self.button5 = ttk.Button(self.top_frame, text="Edit",command=self.editRelease)
        self.button6 = ttk.Button(self.top_frame, text="Push",command=self.pushTemplates)
        self.button7 = ttk.Button(self.top_frame, text="Push",command=self.pushParsers)
        self.button8 = ttk.Button(self.top_frame, text="Push",command=self.pushScripts)
        self.button9 = ttk.Button(self.top_frame, text="Push",command=self.pushSamples)
        self.button10 = ttk.Button(self.top_frame, text="Push",command=self.pushRelease)

        # Labels
        self.label1 = ttk.Label(self.top_frame,text="Templates")
        self.label2 = ttk.Label(self.top_frame,text="ParserConfig")
        self.label3 = ttk.Label(self.top_frame,text="Script")
        self.label4 = ttk.Label(self.top_frame,text="SampleData")
        self.label5 = ttk.Label(self.top_frame,text="ReleaseBin")


        # Grid Everything

        self.button1.grid(row=1,column=1)
        self.button2.grid(row=2,column=1)
        self.button3.grid(row=3,column=1)
        self.button4.grid(row=4,column=1)
        self.button5.grid(row=5,column=1)
        self.button6.grid(row=1,column=2)
        self.button7.grid(row=2,column=2)
        self.button8.grid(row=3,column=2)
        self.button9.grid(row=4,column=2)
        self.button10.grid(row=5,column=2)

        self.label1.grid(row=1,column=0,padx=87)
        self.label2.grid(row=2,column=0)
        self.label3.grid(row=3,column=0)
        self.label4.grid(row=4,column=0)
        self.label5.grid(row=5,column=0)

        # Bottom Half
        self.bot_frame.grid_propagate(1)


        # Radio buttons

        self.radio_var = tk.IntVar()
        self.radio1 = tk.Radiobutton(self.bot_frame,text="Stack 1",variable=self.radio_var, value=0,command=self.setServerList)
        self.radio2 = tk.Radiobutton(self.bot_frame,text="Stack 2",variable=self.radio_var, value=1,command=self.setServerList)
        self.radio3 = tk.Radiobutton(self.bot_frame,text="Both",variable=self.radio_var, value=2,command=self.setServerList)

        # Buttons

        self.push_all_button = ttk.Button(self.bot_frame,text="Push Everything",width=25,command=self.pushEverything)

        # Grid Everything
        self.bot_frame.grid_columnconfigure(0,weight=2)
        self.bot_frame.grid_columnconfigure(1,weight=1)

        self.bot_frame.grid_rowconfigure(0,weight=1)
        self.bot_frame.grid_rowconfigure(1,weight=1)
        self.bot_frame.grid_rowconfigure(2,weight=1)


        self.radio1.grid(row=0,column=0,sticky='w',padx=10)
        self.radio2.grid(row=1,column=0,sticky='w',padx=10)
        self.radio3.grid(row=2,column=0,sticky='w',padx=10)
        self.radio3.invoke()

        self.push_all_button.grid(row=0,column=1,sticky='e',padx=10)



    def setServerList(self):
        if self.radio_var.get() == 2:
            self.server_list = self.all_servers
        elif self.radio_var.get() == 1:
            self.server_list = self.all_servers[-2:]
        else:
            self.server_list = self.all_servers[:4] 
        print(self.server_list)
        print(self.radio_var.get())

    def pushEverything(self):
        pass

    def initTreeView(self):
        # Define new treeview to hold PCases
        columns=["PCase","CSR Name"]#,"Date Created"]
        self.pcase_list = ttk.Treeview(self.pcase_list_frame2,height=30,columns=columns,show="headings")
        self.pcase_list.bind('<<TreeviewSelect>>',self.onselect)

        self.pcase_list.column('PCase',width=65,stretch=False,minwidth=65)
        self.pcase_list.column('CSR Name',width=150,stretch=False,minwidth=100)
        #self.pcase_list.column('Date Created',width=80,stretch=False,minwidth=80)

        self.pcase_list.heading('PCase',text="PCase")
        self.pcase_list.heading('CSR Name',text="CSR Name")
        #self.pcase_list.heading('Date Created',text="Date Created")

        for col in columns:
            self.pcase_list.heading(col,command=lambda _col=col: self.treeview_sort_column(_col, False))

        self.pcase_list.tag_configure('True',background='lightgray')

        self.pcase_list.pack()

    def treeview_sort_column(self,col, reverse):
        l = [(self.pcase_list.set(k, col), k) for k in self.pcase_list.get_children('')]
        l.sort(reverse=reverse)

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            self.pcase_list.move(k, '', index)

        # reverse sort next time
        self.pcase_list.heading(col, command=lambda _col=col: self.treeview_sort_column(_col, not reverse))

    def initQuickButtons(self):

        self.quick_button_frame.grid_propagate(True)

        self.pcase_button = ttk.Button(self.quick_button_frame,text="PCase",width=13,command=lambda: self.openDir("Z:\\IT Documents\\QA\\" + self.pcase_info.cget('text')))
        self.srd_button = ttk.Button(self.quick_button_frame,text="SRD",width=13,command=lambda: self.openWebsite(self.json_data[self.pcase_info.cget('text')]['srd_link']))
        self.ftp_root_button = ttk.Button(self.quick_button_frame,text="FTP Root",width=13,command=lambda: self.openDir("\\\\ssnj-netapp01\\imtest\\imstage01\\ftproot\\"+self.cust_info.cget('text') ))
        self.sf_button = ttk.Button(self.quick_button_frame,text="SalesForce",width=13,command=lambda: self.openWebsite(self.json_data[self.pcase_info.cget('text')]['sf_link']))

        self.pcase_button.grid(row=0,column=0,sticky="E",padx=5)
        self.srd_button.grid(row=0,column=3,sticky="E",padx=5)
        self.sf_button.grid(row=0,column=2,sticky="E",padx=5)
        self.ftp_root_button.grid(row=0,column=1,sticky="E",padx=5)


    def initInfoFrame(self):
        # Initialize Info Section contents
        self.info_frame = self.info_area_frame

        self.pcase_info = ttk.Label(self.info_frame,text="By Using The File Menu",width=23)#font=("Arial",10,"bold"),width=20)
        self.cust_info = ttk.Label(self.info_frame,text="File > New ",width=23)#font=("Arial",10,"bold"),width=20)
        self.sf_info = ttk.Label(self.info_frame,text="",width=23)#font=("Arial",10,"bold"),width=20)
        self.desc_info = ttk.Label(self.info_frame,text="",font=("Arial",10,"bold"),width=54)

        self.last_mod_info = ttk.Label(self.info_frame,text="",width=20)
        self.owner_info = ttk.Label(self.info_frame,text="",width=20)
        self.case_owner_info = ttk.Label(self.info_frame,text="",width=20)

        self.last_mod_label = ttk.Label(self.info_frame,text="Last Modified: ",width=15)
        self.owner_label = ttk.Label(self.info_frame,text="Current Owner: ",width=15)
        self.case_owner_label = ttk.Label(self.info_frame,text="Case Owner: ",width=15)



        # Define Grid for Info Section Items
        self.desc_info.grid(row=0,column=0,columnspan=3,padx=5,sticky="W")
        self.pcase_info.grid(row=1,column=0,padx=5,sticky="W")
        self.cust_info.grid(row=3,column=0,padx=5,sticky="W")
        self.sf_info.grid(row=2,column=0,padx=5,sticky="W")

        self.last_mod_label.grid(row=1,column=1,padx=5,sticky="E")
        self.owner_label.grid(row=2,column=1,padx=5,sticky="E")
        self.case_owner_label.grid(row=3,column=1,padx=5,sticky="E")

        self.last_mod_info.grid(row=1,column=2,padx=5,sticky="E")
        self.owner_info.grid(row=2,column=2,padx=5,sticky="E")
        self.case_owner_info.grid(row=3,column=2,padx=5,sticky="E")




    def initMenuBar(self):
        # Initialize Menu Bar and Menu Items
        self.menu_bar = Menu(self.mainwindow)
        self.toplevel.config(menu=self.menu_bar)
        self.menu_item_1 = Menu(self.menu_bar,tearoff=False)
        self.menu_item_2 = Menu(self.menu_bar,tearoff=False)
        self.menu_item_3 = Menu(self.menu_bar,tearoff=False)
        self.menu_item_4 = Menu(self.menu_bar,tearoff=False)

        self.toplevel.bind_all("<Control-n>",self.newWindowWrapper)
        self.toplevel.bind_all("<Control-q>",self.killWindowWrapper)

        # Add 1 - File Menu Subitems
        self.menu_item_1.add_command(label="New",accelerator="Ctrl+N",command=self.newWindow)
        self.menu_item_1.add_separator()
        self.menu_item_1.add_command(label="Quit",accelerator="Ctrl+Q",command=self.kill_window)

        # Add 2- Edit Menu Subitems
        self.menu_item_2.add_command(label="Change PCase Details",command=self.editWindow)
        self.menu_item_2.add_command(label="Preferences")

        # Add 3- Tools Menu Subitems
        self.menu_item_3.add_command(label="Open Parseamajig",command=lambda: self.openApplication("Z:\\AST\\Utilities\\Parseamajig\\Parseamajig.exe"))
        self.menu_item_3.add_command(label="Open XMLGenerator",command=lambda: self.openApplication("Z:\\AST\\Utilities\\XMLGenerator\\XmlGenerator.exe"))
        self.menu_item_3.add_command(label="Open BillGen Wrapper",command=lambda: self.openApplication("Z:\\AST\\Utilities\\BillGen\\BillGenWrapper.exe"))
        self.menu_item_3.add_separator()
        self.menu_item_3.add_command(label="Run PDF Version Checker",state=DISABLED)


        # Add 4- Help Menu Subitems
        self.menu_item_4.add_command(label="About", command=lambda:aboutWindow(self))
        self.menu_item_4.add_command(label="Confluence Page",command=lambda: self.openWebsite("https://billtrust.atlassian.net/wiki/spaces/AT/overview"))
        
        # Add Menu Options to Bar
        self.menu_bar.add_cascade(label="File", menu=self.menu_item_1)
        self.menu_bar.add_cascade(label="Edit", menu=self.menu_item_2)
        self.menu_bar.add_cascade(label="Tools", menu=self.menu_item_3)
        self.menu_bar.add_cascade(label="Help", menu=self.menu_item_4)


    def threadStart(self):
        self.server = "\\\\ssnj-netapp01\\imtest\\imstage01\\ftproot\\"+self.cust_info.cget('text') 
        self.saved_contents = []
        self.setWatch(True)
        thread = threading.Thread(target=watchDog, args=[self,self.server])
        self.ftpListUpdate()
        thread.start()

    def keepWatching(self):
        return self.watching

    def setWatch(self,boolWatch):
        self.watching = boolWatch


    def ftpListUpdate(self):
        current_contents = [f for f in os.listdir(self.server) if os.path.isfile(os.path.join(self.server, f))]
        self.ftpList.delete(0,self.ftpList.size())
        for item in current_contents:
            if item != "Thumbs.db":
                self.ftpList.insert(tk.END,item)


    def newWindowWrapper(self,parent):
        self.newWindow()

    def killWindowWrapper(self,parent):
        self.kill_window()

    def saveNotes(self):
        user = os.getlogin()
        data_folder = "C:\\Users\\%s\\AppData\\Roaming\\PCASR_DEV" %user
        data_file = data_folder+"\\pcasr.json"

        print(self.notes.get("1.0","end"))
        if os.path.isdir(data_folder) and self.notes_hash and self.notes_hash != hash(self.notes.get("1.0","end")):
            pcase = self.pcase_info.cget('text')
            notes = self.notes.get("1.0","end")

            data = self.json_data
            data[pcase].update({'notes':notes.rstrip()})

            with open(data_file,'w') as outfile:
                json.dump(data,outfile)
            self.notehash = hash(notes)

    def onselect(self,parent):
        self.saveNotes()
        self.setWatch(False)
        index = self.pcase_list.selection()
        values= self.pcase_list.item(index)['values']
        self.notes.delete(1.0, END)
        self.updateInfo(values)
        self.lastSelected = index

    def updateInfo(self,values):

        pcase = values[0]
        cust_name = values[1]
        data = self.json_data

        self.pcase_info.config(text=pcase)
        self.cust_info.config(text=cust_name)

        self.sf_info.config(text=data[pcase]['case_number'])
        self.desc_info.config(text=data[pcase]['subject'])
        self.last_mod_info.config(text=data[pcase]['last_modified'].split('.')[0][:-3].replace('T'," "))
        self.owner_info.config(text=data[pcase]['case_owner'])
        self.case_owner_info.config(text=data[pcase]['parent_case_owner'])

        self.ftpList.delete(0,self.ftpList.size())
        self.threadStart()

        self.pcase_tree()

        user = os.getlogin()
        data_folder = "C:\\Users\\%s\\AppData\\Roaming\\PCASR_DEV" %user
        data_file = data_folder+"\\pcasr.json"
        data = self.json_data
        notes = data[pcase]['notes']
        self.notes_hash = hash(notes)
        self.srd_info = data[pcase]['srd_link']
        self.notes.insert(END,notes)
        self.notehash = hash(self.notes.get("1.0","end"))


    def loadPCases(self):
        user = os.getlogin()
        data_file = "C:\\Users\\%s\\AppData\\Roaming\\PCASR_DEV\\pcasr.json" %user
        odd_even = True
        if os.path.exists(data_file):
            self.pcase_list.delete(*self.pcase_list.get_children())
            if not self.json_data:
                with open(data_file) as json_file:
                    data = json.load(json_file)
                    self.json_data = data
                    if 'notes' in data.keys():
                        self.notes_hash = hash(data['notes'])
                    for case in data:
                        self.pcase_list.insert('','end',iid=[data[case]['pcase']],values=[data[case]['pcase'],data[case]['cust_name']],tags=[str(odd_even)])
                        odd_even = not odd_even
            else:
                data = self.json_data
                for case in data:
                    self.pcase_list.insert('','end',iid=[data[case]['pcase']],values=[data[case]['pcase'],data[case]['cust_name']],tags=[str(odd_even)])
                    odd_even = not odd_even


    def newWindow(self):
        saveDialogWindow(self,False)

    def editWindow(self):
        saveDialogWindow(self,True)

    def run(self):
        self.mainwindow.mainloop()

    def pcase_tree(self):
        pcase = self.pcase_info.cget('text')
        if pcase:
            if self.first_init:
                self.nodes = dict()
                frame = self.tab_1
                self.first_init = False
            else:
                self.nodes = dict()
                frame = self.tab_1
                self.ysb.destroy()
                self.tree.destroy()

            self.tree = ttk.Treeview(frame,height=18)
            self.ysb = ttk.Scrollbar(frame, orient='vertical', command=self.tree.yview)
            self.tree.configure(yscroll=self.ysb.set)
            self.tree.heading('#0', text='PCase Folder', anchor='w')

            self.ysb.pack(side=RIGHT,fill=Y)
            self.tree.pack(fill=BOTH)
            

            abspath = "Z:\\IT Documents\\QA\\" + pcase
            self.insert_node('', abspath, abspath)
            self.tree.bind('<<TreeviewOpen>>', self.open_node)

            self.tree.focus(self.tree.get_children()[0])
            self.tree.item(self.tree.focus(),open=True)
            self.open_node('')



        #else:
        #    messagebox.showwarning('Error', 'Please Enter a Valid PCASE\nBefore Trying to View The PCase File Tree')

    def insert_node(self, parent, text, abspath):
        node = self.tree.insert(parent, 'end', text=text, open=False)
        if os.path.isdir(abspath):
            self.nodes[node] = abspath
            self.tree.insert(node, 'end')

    def open_node(self, event):
        node = self.tree.focus()
        abspath = self.nodes.pop(node, None)
        print(node)
        print(abspath)
        if abspath:
            self.tree.delete(self.tree.get_children(node))
            for p in os.listdir(abspath):
                self.insert_node(node, p, os.path.join(abspath, p))



    def filesInDir(self, direc,fileType=""):
        if fileType != "":
            return [(i) for i in os.listdir(direc) if fileType in i]
        else:
            return os.listdir(direc)


    def editFiles(self,subDir,fileType=""):
        pcase = self.pcase_info.cget('text')
        if pcase:
            subFolder_path = "Z:\\IT Documents\\QA\\" + pcase + subDir
            #file_path_string = filedialog.askopenfilenames(initialdir = fdt_path,filetypes = (("CSV files","*.csv"),("all files","*.*")))
            files = self.filesInDir(subFolder_path,fileType)
            if len(files) == 0:
                messagebox.showwarning('Error','No files found in directory')
            elif len(files) == 1:
                os.startfile(subFolder_path+"\\"+files[0], 'open')
            else:
                print(files)
                selected_files = filePickerWindow(self,subFolder_path,files).show()
                if selected_files:
                    for file in selected_files:
                        print(file)
                        os.startfile(subFolder_path+"\\"+file, 'open')
        #else:
        #    messagebox.showwarning('Error', 'Please Enter a Valid PCASE\nBefore Trying to Open or Edit Files')

    def pushFiles(self,subDir,fileType=""):
        pcase = self.pcase_info.cget('text')
        if pcase:
            subFolder_path = "Z:\\IT Documents\\QA\\" + pcase + subDir
            files = self.filesInDir(subFolder_path,fileType)
            if len(files) == 0:
                messagebox.showwarning('Error','No files found in directory')
            elif len(files) == 1:
                copyFilesWindow(self,subFolder_path,subDir,files)
            else:
                print(files)
                selected_files = filePickerWindow(self,subFolder_path,files).show()
                if selected_files:
                    copyFilesWindow(self,subFolder_path,subDir,selected_files)
        #else:
        #    messagebox.showwarning('Error', 'Please Enter a Valid PCASE\nBefore Trying to Open or Edit Files')

    def editTemplates(self):
        self.editFiles("\\FDT",".csv")
    def editParsers(self):
        self.editFiles("\\VC\\ParserConfigs",".xml")
    def editScripts(self):
        self.editFiles("\\VC\\Scripts",".py")
    def editSamples(self):
        self.editFiles("\\Sample Data")        
    def editRelease(self):
        self.editFiles("\\VC\\Release\\Bin",)
        
    def pushTemplates(self):
        self.pushFiles("\\FDT",".csv")
    def pushParsers(self):
        self.pushFiles("\\VC\\ParserConfigs",".xml")
    def pushScripts(self):
        self.pushFiles("\\VC\\Scripts",".py")
    def pushSamples(self):
        self.pushFiles("\\SAMPLE_DATA")        
    def pushRelease(self):
        self.pushFiles("\\VC\\Release\\Bin",)

    def openDir(self,path):
        expString = "explorer " + path
        print(expString)
        subprocess.Popen(expString)     

    def openSRD(self):
        srd = self.srd_info
        if srd:
            webbrowser.get('chrome').open(srd)
        else:
            messagebox.showwarning('Error', 'URL appears to be invalid.\nPlease enter a valid URL.')

    def openApplication(self,path):
        try:
            messagebox.showinfo('Just a Second...', 'This can  a minute to open, please be patient.')
            os.startfile(path)
        except:
            messagebox.showwarning('Error', 'You must be connected to the VPN to open this.')

    def openWebsite(self,page):
        if page:
            webbrowser.get('chrome').open(page)
        else:
            messagebox.showwarning('Error', 'There is no SRD saved for this case.\nYou can add via Edit > Change PCase Details') 
        

if __name__ == '__main__':
    app = PCaser()
    app.run()

