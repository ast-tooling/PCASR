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

# local imports
import save_dialog_window
import file_picker_window
import copy_files_window
import create_tool_tip
import about_window
import watch_dog

class PCaser:

    webbrowser.register('chrome',None,webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
    webbrowser.register('firefox',None,webbrowser.BackgroundBrowser("C://Program Files //Mozilla Firefox//Chrome//firefox.exe"))


    def __init__(self):

        self.initMainFrame()

        self.all_servers = ["ssnj-imbisc10","ssnj-imbisc11","ssnj-imbisc12","ssnj-imbisc13","ssnj-imbisc20","ssnj-imbisc21"]
        self.json_data = {}
        self.archive_data = {}
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
        self.loadArchive()


        self.watching = False
        # Select the first listbox item if there is one
        try:
            topSelect = topSelect = self.pcase_list.get_children()[0]
            print(topSelect)
            self.pcase_list.selection_set(topSelect)
            self.updateInfo(self.pcase_list.item(topSelect)['values'],self.json_data)

            self.threadStart()
        except:
            pass


        # Tell the window manager to give focus back after the X button is hit
        self.toplevel.protocol("WM_DELETE_WINDOW", self.kill_window)


    def kill_window(self):
        self.setWatch(False)
        self.notes_hash = self.notes.get("1.0","end")
        self.savePCase()
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

        #self.push_all_button = ttk.Button(self.bot_frame,text="Push Everything",width=25,command=self.pushEverything)

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

        #self.push_all_button.grid(row=0,column=1,sticky='e',padx=10)



    def setServerList(self):
        if self.radio_var.get() == 2:
            self.server_list = self.all_servers
        elif self.radio_var.get() == 1:
            self.server_list = self.all_servers[-2:]
        else:
            self.server_list = self.all_servers[:4] 
        print(self.server_list)
        print(self.radio_var.get())
        

    def initTreeView(self):

         # Tab Definitions
        self.pcase_tab_area = ttk.Notebook(self.pcase_list_frame2,width=245,height=650)
        self.pcase_tab_area.bind('<<NotebookTabChanged>>', self.on_tab_change)

        self.current_tab = Frame(self.pcase_tab_area)
        self.archive_tab = Frame(self.pcase_tab_area)

        self.pcase_tab_area.add(self.current_tab,text="            Current            ")
        self.pcase_tab_area.add(self.archive_tab,text="            Archive            ")


        # Define active tab contents
        columns=["PCase","CSR Name"]#,"Date Created"]
        self.pcase_list = ttk.Treeview(self.current_tab,height=32,columns=columns,show="headings")
        self.pcase_list.bind('<<TreeviewSelect>>',self.onselect)

        self.pcase_list.column('PCase',width=78,stretch=False,minwidth=78)
        self.pcase_list.column('CSR Name',width=163,stretch=False,minwidth=100)
        #self.pcase_list.column('Date Created',width=80,stretch=False,minwidth=80)

        self.pcase_list.heading('PCase',text="PCase")
        self.pcase_list.heading('CSR Name',text="CSR Name")
        #self.pcase_list.heading('Date Created',text="Date Created")

        # Emable sorting
        for col in columns:
            self.pcase_list.heading(col,command=lambda _col=col: self.treeview_sort_column(_col, False))

        self.pcase_list.grid(row=0,column=0)


        # Define archive tab contents
        columns=["PCase","CSR Name"]#,"Date Created"]
        self.archive_list = ttk.Treeview(self.archive_tab,height=32,columns=columns,show="headings")
        self.archive_list.bind('<<TreeviewSelect>>',self.onselect)

        self.archive_list.column('PCase',width=78,stretch=False,minwidth=78)
        self.archive_list.column('CSR Name',width=163,stretch=False,minwidth=100)
        #self.pcase_list.column('Date Created',width=80,stretch=False,minwidth=80)

        self.archive_list.heading('PCase',text="PCase")
        self.archive_list.heading('CSR Name',text="CSR Name")
        #self.pcase_list.heading('Date Created',text="Date Created")

        # Emable sorting
        for col in columns:
            self.archive_list.heading(col,command=lambda _col=col: self.treeview_sort_column(_col, False))

        self.archive_list.grid(row=0,column=0)

        # Add function buttons
        self.update_button = ttk.Button(self.pcase_list_frame2,text='Update',width=10,command=self.editWindow)
        self.new_button = ttk.Button(self.pcase_list_frame2,text='New',width=10,command=self.newWindow)
        self.archive_button = ttk.Button(self.pcase_list_frame2,text='Archive',width=10,command=self.archiveCase)

        # Add items to frame
        self.pcase_tab_area.grid(row=0,column=0,columnspan=3)
        self.new_button.grid(row=1,column=0,padx=1)
        self.update_button.grid(row=1,column=1,padx=1)
        self.archive_button.grid(row=1,column=2,padx=1,pady=2)

    def on_tab_change(self,event):
        tab = event.widget.tab('current')['text']
        if tab == "            Current            ":
            self.archive_button.config(text='Archive')
            topSelect = self.pcase_list.get_children()[0]
            self.pcase_list.selection_set(topSelect)
            self.updateInfo(self.pcase_list.item(topSelect)['values'],self.json_data)
            self.threadStart()
        elif tab == "            Archive            ":
            self.archive_button.config(text='UnArchive')
            topSelect = self.archive_list.get_children()[0]
            self.archive_list.selection_set(topSelect)
            #self.updateInfo(self.archive_list.item(topSelect)['values'],self.archive_data)

            self.threadStart()


    def archiveCase(self):

        """
        TODO: 
        -Create mutator and accessor methods for
            -data folder
            -data file
            -archive file
            -current selection
                -pcase, etc
        -Replace all updates and retrievals with new methods
        -finish archiveCase method
        """

        user = os.getlogin()
        self.data_folder = "C:\\Users\\%s\\AppData\\Roaming\\PCASR" %user
        self.data_file = self.data_folder+"\\archive.json"
        self.pcase_file = self.data_folder+"\\pcasr.json"

        pcase = self.pcase_info.cget('text')

        if not os.path.exists(self.data_file):
            data = {} 
            # Initialize Data File
            with open(self.data_file,'w') as outfile:
                json.dump(data,outfile)

        with open(self.data_file) as json_file:
            data = json.load(json_file)
            self.archive_data = data

        self.archive_data[pcase] = self.json_data[pcase]
        del self.json_data[pcase]

        with open(self.pcase_file,'w') as outfile:
            json.dump(self.json_data,outfile)

        with open(self.data_file,'w') as outfile:
            json.dump(self.archive_data,outfile)


        self.loadPCases()
        self.loadArchive()

        try:
            topSelect = self.pcase_list.get_children()[0]
            self.pcase_list.selection_set(topSelect)
            self.updateInfo(self.pcase_list.item(topSelect)['values'],self.json_data)

            self.threadStart()
        except:
           pass


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
        self.desc_info = ttk.Label(self.info_frame,text="",font=("Arial",10,"bold"),width=51)

        self.last_mod_info = ttk.Label(self.info_frame,text="",width=20)
        self.owner_info = ttk.Label(self.info_frame,text="",width=20)
        self.case_owner_info = ttk.Label(self.info_frame,text="",width=20)

        self.last_mod_label = ttk.Label(self.info_frame,text="Last Modified: ",width=15)
        self.owner_label = ttk.Label(self.info_frame,text="Current Owner: ",width=15)
        self.case_owner_label = ttk.Label(self.info_frame,text="Case Owner: ",width=15)

        self.refresh_button = ttk.Button(self.info_frame,text="↻",width="3",command=self.refreshSFInfo)



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

        self.refresh_button.grid(row=0,column=2,sticky="E",padx=5)




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
        thread = threading.Thread(target=watch_dog.watchDog, args=[self,self.server])
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

    def onselect(self,parent):
        self.savePCase()
        self.setWatch(False)
        index = self.pcase_list.selection()
        values= self.pcase_list.item(index)['values']
        self.notes.delete(1.0, END)
        self.updateInfo(values,self.json_data)
        self.lastSelected = index

    def archiveSelect(self,parent):
        self.savePCase()
        self.setWatch(False)
        index = self.pcase_list.selection()
        values= self.pcase_list.item(index)['values']
        self.notes.delete(1.0, END)
        self.updateInfo(values,self.archive_data)
        self.lastSelected = index

    def updateInfo(self,values,data):

        pcase = values[0]
        cust_name = values[1]

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
        data_folder = "C:\\Users\\%s\\AppData\\Roaming\\PCASR" %user
        data_file = data_folder+"\\pcasr.json"
        data = self.json_data
        notes = data[pcase]['notes']
        self.notes_hash = hash(notes)
        self.srd_info = data[pcase]['srd_link']
        self.notes.insert(END,notes)
        self.notehash = hash(self.notes.get("1.0","end"))


    def savePCase(self):
        user = os.getlogin()
        data_folder = "C:\\Users\\%s\\AppData\\Roaming\\PCASR" %user
        data_file = data_folder+"\\pcasr.json"
        data = self.json_data
        pcase = self.pcase_info.cget('text')
        notes = self.notes.get("1.0","end")

        try:
            self.archive_data[pcase]
            archive_flag = True
        except:
            archive_flag = False

        if data:
            if os.path.isdir(data_folder) and self.notes_hash != hash(self.notes.get("1.0","end")) and not archive_flag:

                
                data[pcase].update({'notes':notes.rstrip()})
                self.notehash = hash(notes)

            with open(data_file,'w') as outfile:
                json.dump(data,outfile)
            

    def loadPCases(self):
        user = os.getlogin()
        data_file = "C:\\Users\\%s\\AppData\\Roaming\\PCASR\\pcasr.json" %user

        if os.path.exists(data_file):
            self.pcase_list.delete(*self.pcase_list.get_children())
            if not self.json_data:
                with open(data_file) as json_file:
                    data = json.load(json_file)
                    self.json_data = data
                    if 'notes' in data.keys():
                        self.notes_hash = hash(data['notes'])
                    for case in data:
                        self.pcase_list.insert('','end',iid=[data[case]['pcase']],values=[data[case]['pcase'],data[case]['cust_name']])
            else:
                data = self.json_data
                for case in data:
                    self.pcase_list.insert('','end',iid=[data[case]['pcase']],values=[data[case]['pcase'],data[case]['cust_name']])

    def loadArchive(self):
        user = os.getlogin()
        data_file = "C:\\Users\\%s\\AppData\\Roaming\\PCASR\\archive.json" %user

        if os.path.exists(data_file):
            self.archive_list.delete(*self.archive_list.get_children())
            if not self.archive_data:
                with open(data_file) as json_file:
                    data = json.load(json_file)
                    self.archive_data = data
                    for case in data:
                        self.archive_list.insert('','end',iid=[data[case]['pcase']],values=[data[case]['pcase'],data[case]['cust_name']])
            else:
                data = self.archive_data
                for case in data:
                    self.archive_list.insert('','end',iid=[data[case]['pcase']],values=[data[case]['pcase'],data[case]['cust_name']])



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
                selected_files = filePickerWindow(self,subFolder_path).show()
                if selected_files:
                    for file in selected_files:
                        print(file)
                        os.startfile(subFolder_path+"\\"+file[1]+file[0], 'open')
        #else:
        #    messagebox.showwarning('Error', 'Please Enter a Valid PCASE\nBefore Trying to Open or Edit Files')

    def pushFiles(self,subDir,fileType=""):
        pcase = self.pcase_info.cget('text')
        if pcase:
            subFolder_path = "Z:\\IT Documents\\QA\\" + pcase + subDir
            files = self.filesInDir(subFolder_path,fileType)
            if len(files) == 0:
                messagebox.showwarning('Error','No files found in directory')
            elif len(files) == 1 and not os.path.isdir(subFolder_path+'\\'+files[0]):
                print(files)
                copyFilesWindow(self,subFolder_path,subDir,files)
            else:
                print(files)
                selected_files = filePickerWindow(self,subFolder_path).show()
                if selected_files:
                    copyFilesWindow(self,subFolder_path,subDir,selected_files)
        #else:
        #    messagebox.showwarning('Error', 'Please Enter a Valid PCASE\nBefore Trying to Open or Edit Files')

    def refreshSFInfo(self):
        print('does this do anything')
        self.data_file = "C:\\Users\\%s\\AppData\\Roaming\\PCASR\\config.txt" % os.getlogin()
        if not os.path.exists(self.data_file):
            messagebox.showwarning('Error', 'You must first add your sf credentials to\nC:\\Users\\<you>\\AppData\\Roaming\\PCASR\\credentials.txt\nAn example can be found at Z:\\AST\\Utilities\\PCASR')

        else:
            config = configparser.ConfigParser()
            config.read(self.data_file)

            client = Salesforce(
                username=config.get('credentials','username'),
                password=config.get('credentials','password'),
                security_token=config.get('credentials','security_token')
                )

            pcase = self.pcase_info.cget('text')
            if pcase != 'By Using The File Menu':
                case_id = self.json_data[pcase]['sf_link'].split('/')[-1]

                case_info = client.Case.get(case_id)

                self.json_data[pcase]['last_modified'] = case_info['LastModifiedDate']
                self.json_data[pcase]['case_owner'] = case_info['Case_Owner__c']
                self.json_data[pcase]['parent_case_owner'] = case_info['Parent_Case_Owner__c']

                index = self.pcase_list.selection()
                values= self.pcase_list.item(index)['values']
                self.updateInfo(values,self.json_data)



            

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
            messagebox.showinfo('Just a Second...', 'This can take a minute to open, please be patient.')
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

