import tkinter as tk
import subprocess
from tkinter import ttk
import webbrowser
import os
import threading
import json
from simple_salesforce import Salesforce
import configparser

# local imports
import save_dialog_window
import file_picker_window
import copy_files_window
import about_window
import watch_dog


'''The Main Class of the project.  

This both creates the UI for the window and 
drives the logic.
'''
class PCaser:

    webbrowser.register('chrome',None,webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
    webbrowser.register('firefox',None,webbrowser.BackgroundBrowser("C://Program Files //Mozilla Firefox//Chrome//firefox.exe"))

    ''' Configure initial variables and call all the methods that instantiate their sub-areas of the UI
    '''
    def __init__(self):

        self.initMainFrame()

        self.all_servers = ["ssnj-imbisc10","ssnj-imbisc11","ssnj-imbisc12","ssnj-imbisc13","ssnj-imbisc20","ssnj-imbisc21"]
        self.json_data = {}
        self.archive_data = {}
        self.notes_hash = ""
        self.server_list = []
        self.data_folder = ""
        self.data_file = ""
        self.archive_file = ""
        self.pcase = ""
        self.cust = ""
        
        user = os.getlogin()
        data_folder = "C:\\Users\\%s\\AppData\\Roaming\\PCASR" %user
        data_file = data_folder+"\\pcasr.json"
        archive_file = data_folder+"\\archive.json"
        
        self.setDataFolder(data_folder)
        self.setDataFile(data_file)
        self.setArchiveFile(archive_file)

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

    def initFTPFrame(self):
        self.ftpList = tk.Listbox(self.ftp_root_frame,height=6,width=65)
        self.ftpList.pack(fill=tk.BOTH)

       
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
        self.space_label = tk.Label(self.tab_frame)
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

        self.tab_1 = tk.Frame(self.tab_area)
        self.tab_2 = tk.Frame(self.tab_area)
        self.tab_3 = tk.Frame(self.tab_area)
        self.tab_4 = tk.Frame(self.tab_area)

        self.tab_area.add(self.tab_1,text="Quick View")
        self.tab_area.add(self.tab_2,text="Notes")
        self.tab_area.add(self.tab_3,text="Edit/Push")
        self.tab_area.add(self.tab_4,text="Tab 4")
        self.tab_area.pack(expand=1,fill="both")


        # Tab 1 - Notes
        self.notes = tk.Text(self.tab_2)
        self.note_scroll = tk.Scrollbar(self.tab_2,command=self.notes.yview)
        self.notes['yscrollcommand'] = self.note_scroll.set

        # Initialize Hash Value
        self.notehash = ""
        self.note_scroll.pack(side=tk.RIGHT,fill = tk.Y)
        self.notes.pack()


        # Tab 3 - Edit/Push Area


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




    def setServerList(self):
        if self.radio_var.get() == 2:
            self.server_list = self.all_servers
        elif self.radio_var.get() == 1:
            self.server_list = self.all_servers[-2:]
        else:
            self.server_list = self.all_servers[:4] 
        

    def initTreeView(self):

         # Tab Definitions
        self.pcase_tab_area = ttk.Notebook(self.pcase_list_frame2,width=245,height=650)
        self.pcase_tab_area.bind('<<NotebookTabChanged>>', self.on_tab_change)

        self.current_tab = tk.Frame(self.pcase_tab_area)
        self.archive_tab = tk.Frame(self.pcase_tab_area)

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

        # Enable sorting
        for col in columns:
            self.pcase_list.heading(col,command=lambda _col=col: self.treeview_sort_column(_col, False,self.pcase_list))

        self.pcase_list.grid(row=0,column=0)



        # Define archive tab contents
        columns=["PCase","CSR Name"]#,"Date Created"]
        self.archive_list = ttk.Treeview(self.archive_tab,height=32,columns=columns,show="headings")
        self.archive_list.bind('<<TreeviewSelect>>',self.archiveSelect)

        self.archive_list.column('PCase',width=78,stretch=False,minwidth=78)
        self.archive_list.column('CSR Name',width=163,stretch=False,minwidth=100)
        #self.pcase_list.column('Date Created',width=80,stretch=False,minwidth=80)

        self.archive_list.heading('PCase',text="PCase")
        self.archive_list.heading('CSR Name',text="CSR Name")
        #self.pcase_list.heading('Date Created',text="Date Created")

        # Enable sorting
        for col in columns:
            self.archive_list.heading(col,command=lambda _col=col: self.treeview_sort_column(_col, False,self.archive_list))

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
        

    def initQuickButtons(self):

        self.quick_button_frame.grid_propagate(True)

        self.pcase_button = ttk.Button(self.quick_button_frame,text="PCase",width=13,command=lambda: self.openDir("Z:\\IT Documents\\QA\\" + self.getPCaseString()))
        self.srd_button = ttk.Button(self.quick_button_frame,text="SRD",width=13,command=lambda: self.openWebsite(self.json_data[self.getPCaseString()]['srd_link']))
        self.ftp_root_button = ttk.Button(self.quick_button_frame,text="FTP Root",width=13,command=lambda: self.openDir("\\\\ssnj-netapp01\\imtest\\imstage01\\ftproot\\"+self.getCustString() ))
        self.sf_button = ttk.Button(self.quick_button_frame,text="SalesForce",width=13,command=lambda: self.openWebsite(self.json_data[self.getPCaseString()]['sf_link']))

        self.pcase_button.grid(row=0,column=0,sticky="E",padx=5)
        self.srd_button.grid(row=0,column=3,sticky="E",padx=5)
        self.sf_button.grid(row=0,column=2,sticky="E",padx=5)
        self.ftp_root_button.grid(row=0,column=1,sticky="E",padx=5)


    def initInfoFrame(self):
        # Initialize Info Section contents
        self.info_frame = self.info_area_frame

        self.pcase_info = ttk.Label(self.info_frame,text="By Using The File Menu",width=23)
        self.cust_info = ttk.Label(self.info_frame,text="File > New ",width=23)
        self.sf_info = ttk.Label(self.info_frame,text="",width=23)
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
        self.menu_bar = tk.Menu(self.mainwindow)
        self.toplevel.config(menu=self.menu_bar)
        self.menu_item_1 = tk.Menu(self.menu_bar,tearoff=False)
        self.menu_item_2 = tk.Menu(self.menu_bar,tearoff=False)
        self.menu_item_3 = tk.Menu(self.menu_bar,tearoff=False)
        self.menu_item_4 = tk.Menu(self.menu_bar,tearoff=False)

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
        self.menu_item_3.add_command(label="Run PDF Version Checker",state=tk.DISABLED)


        # Add 4- Help Menu Subitems
        self.menu_item_4.add_command(label="About", command=lambda:about_window.aboutWindow(self))
        self.menu_item_4.add_command(label="Confluence Page",command=lambda: self.openWebsite("https://billtrust.atlassian.net/wiki/spaces/AT/overview"))
        
        # Add Menu Options to Bar
        self.menu_bar.add_cascade(label="File", menu=self.menu_item_1)
        self.menu_bar.add_cascade(label="Edit", menu=self.menu_item_2)
        self.menu_bar.add_cascade(label="Tools", menu=self.menu_item_3)
        self.menu_bar.add_cascade(label="Help", menu=self.menu_item_4)
        
    '''Called by tk.notebook object on pcase_tab_area when different tabs are selected. 
     
    Configures new, update, archive,and delete functions for buttons.
    Determines which list to use to update PCASE info pane.
    Enables/Disables notes tab.
    '''
    def on_tab_change(self,event):
        tab = event.widget.tab('current')['text']
        if tab == "            Current            ":
            self.archive_button.config(text='Archive')
            self.archive_button.config(command=self.archiveCase)
            self.tab_area.tab(1,state="normal")
            topSelect = self.pcase_list.get_children()[0]
            self.pcase_list.selection_set(topSelect)
            self.onselect(self)
            self.threadStart()
        elif tab == "            Archive            ":
            self.savePCase()
            self.archive_button.config(text='UnArchive')
            self.archive_button.config(command=self.unArchiveCase)
            self.new_button.config(text='Delete')
            self.new_button.config(command=self.deleteCase)
            self.tab_area.tab(1,state="disabled")
            topSelect = self.archive_list.get_children()[0]
            self.archive_list.selection_set(topSelect)
            self.archiveSelect(self)
            self.threadStart()
            
    def unArchiveCase(self):
        archive_file = self.getArchiveFile()
        pcase_file = self.getDataFile()

        pcase = self.getPCaseString()
        
        with open(archive_file) as json_file:
            data = json.load(json_file)
            self.archive_data = data

        self.json_data[pcase] = self.archive_data[pcase]
        del self.archive_data[pcase]

        self.saveJSON(pcase_file,self.json_data)
        self.saveJSON(archive_file,self.archive_data)


        self.loadPCases()
        self.loadArchive()

        try:
            topSelect = self.pcase_list.get_children()[0]
            self.pcase_list.selection_set(topSelect)
            self.updateInfo(self.pcase_list.item(topSelect)['values'],self.json_data)

            self.threadStart()
        except:
           pass
        
    def archiveCase(self):
        archive_file = self.getArchiveFile()
        pcase_file = self.getDataFile()

        pcase = self.getPCaseString()
        
        if not os.path.exists(archive_file):
            data = {} 
            # Initialize Data File
            self.saveJSON(archive_file,{})

        with open(archive_file) as json_file:
            data = json.load(json_file)
            self.archive_data = data

        self.archive_data[pcase] = self.json_data[pcase]
        del self.json_data[pcase]

        self.saveJSON(pcase_file,self.json_data)

        self.saveJSON(archive_file,self.archive_data)


        self.loadPCases()
        self.loadArchive()

        try:
            topSelect = self.pcase_list.get_children()[0]
            self.pcase_list.selection_set(topSelect)
            self.updateInfo(self.pcase_list.item(topSelect)['values'],self.json_data)

            self.threadStart()
        except:
           pass
       
    def deleteCase(self):
        
        confirmDialog = tk.messagebox.askquestion('Delete PCase?','Are you sure you want to remove this PCase?',icon='warning')
        if confirmDialog == 'yes':
            archive_file = self.getArchiveFile()
            pcase = self.getPCaseString()
            del self.archive_data[pcase]
            
            self.saveJSON(archive_file,self.archive_data)
                
            self.loadArchive()

            try:
                topSelect = self.pcase_list.get_children()[0]
                self.pcase_list.selection_set(topSelect)
                self.updateInfo(self.pcase_list.item(topSelect)['values'],self.json_data)
    
                self.threadStart()
            except:
               pass
    '''Called by treeview instances to sort by column headers
    '''
    def treeview_sort_column(self,col, reverse,treeview):
        l = [(treeview.set(k, col), k) for k in treeview.get_children('')]
        l.sort(reverse=reverse)

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            treeview.move(k, '', index)

        # reverse sort next time
        treeview.heading(col, command=lambda _col=col: self.treeview_sort_column(_col, not reverse))

    '''Initializes thread to watch FTP root for file changes
    '''
    def threadStart(self):
        self.server = "\\\\ssnj-netapp01\\imtest\\imstage01\\ftproot\\"+self.getCustString() 
        self.saved_contents = []
        self.setWatch(True)
        thread = threading.Thread(target=watch_dog.watchDog, args=[self,self.server])
        self.ftpListUpdate()
        thread.start()
        
    '''Helper method for threadStart
    '''
    def keepWatching(self):
        return self.watching
    
    '''Helper method for threadStart
    '''
    def setWatch(self,boolWatch):
        self.watching = boolWatch

    '''Handles updating UI listbox with FTP root contents
    '''
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
        
    '''Called by the treeview when a new selection is made.
    
    Saves PCase data to json, disables ftp root watchdog, calls updateInfo
    '''
    def onselect(self,parent):
        self.savePCase()
        self.setWatch(False)
        index = self.pcase_list.selection()
        values= self.pcase_list.item(index)['values']
        self.notes.delete(1.0, tk.END)
        self.updateInfo(values,self.json_data)
        self.lastSelected = index
    
    '''Called by the archive treeview when a new selection is made.
    
    Disables ftp root watchdog, calls updateInfo
    '''      
    def archiveSelect(self,parent):
        self.setWatch(False)
        index = self.archive_list.selection()
        values= self.archive_list.item(index)['values']
        self.notes.delete(1.0, tk.END)
        self.updateInfo(values,self.archive_data)
        self.lastSelected = index


    '''Updates the PCase Info area, quick view, edit/push, and notes
    
    PARAMS:
        values - a 1x2 list containing the pcase and the customer name
        data - either json_data or archive_data, the json with case data in memory
    '''
    def updateInfo(self,values,data):

        pcase = values[0]
        cust_name = values[1]
        
        self.setPCaseString(pcase)
        self.setCustString(cust_name)

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

        notes = data[pcase]['notes']
        self.notes_hash = hash(notes)
        self.srd_info = data[pcase]['srd_link']
        self.notes.insert(tk.END,notes)
        self.notehash = hash(self.notes.get("1.0","end"))

    def savePCase(self):
        data_folder = self.getDataFolder()
        data_file = self.getDataFile()
        data = self.json_data
        pcase = self.getPCaseString()
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

            self.saveJSON(data_file,data)
            
    '''Loads JSON file into memory if not already done, updates pcase_list treeview
    '''
    def loadPCases(self):
        data_file = self.getDataFile()

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
    
    '''Loads Archive JSON file into memory if not already done, updates archive_list treeview
    '''
    def loadArchive(self):
        data_file = self.getArchiveFile()

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
        save_dialog_window.saveDialogWindow(self,False)

    def editWindow(self):
        save_dialog_window.saveDialogWindow(self,True)

    def run(self):
        self.mainwindow.mainloop()
        
    '''Loads 'Quick View' PCase treeview
    '''
    def pcase_tree(self):
        pcase = self.getPCaseString()
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

            self.ysb.pack(side=tk.RIGHT,fill=tk.Y)
            self.tree.pack(fill=tk.BOTH)
            

            abspath = "Z:\\IT Documents\\QA\\" + pcase
            self.insert_node('', abspath, abspath)
            self.tree.bind('<<TreeviewOpen>>', self.open_node)

            self.tree.focus(self.tree.get_children()[0])
            self.tree.item(self.tree.focus(),open=True)
            self.open_node('')

    '''Helper method for pcase_tree
    '''
    def insert_node(self, parent, text, abspath):
        node = self.tree.insert(parent, 'end', text=text, open=False)
        if os.path.isdir(abspath):
            self.nodes[node] = abspath
            self.tree.insert(node, 'end')
            
    '''Helper method for pcase_tree
    '''
    def open_node(self, event):
        node = self.tree.focus()
        abspath = self.nodes.pop(node, None)
        if abspath:
            self.tree.delete(self.tree.get_children(node))
            for p in os.listdir(abspath):
                self.insert_node(node, p, os.path.join(abspath, p))


    ''' Returns a list of the files in a given directory, optionally of a specified file type.
    
    PARAMS:
        direc - the directory to list files from
        fileType - (Optional) the file extension to exclusively return
    '''
    def filesInDir(self, direc,fileType=""):
        if fileType != "":
            return [(i) for i in os.listdir(direc) if fileType in i]
        else:
            return os.listdir(direc)


    ''' Given a subdirectory and a filetype, open files in that directory.
    
    There are three possible cases for the contents of the folder:
        -There are no files in the folder.
            --This will open an error dialog
        -There is one file in the folder
            --This will open that file without further user input
        --There are multiple files in the folder
            --This will call the filePickerWindow class and have the user
              choose which files to open
    
    PARAMS:
        subDir - the directory to look in for file to open
        fileType - the type of file to open
    '''
    def editFiles(self,subDir,fileType=""):
        pcase = self.getPCaseString()
        if pcase:
            subFolder_path = "Z:\\IT Documents\\QA\\" + pcase + subDir
            files = self.filesInDir(subFolder_path,fileType)
            if len(files) == 0:
                tk.messagebox.showwarning('Error','No files found in directory')
            elif len(files) == 1:
                os.startfile(subFolder_path+"\\"+files[0], 'open')
            else:
                selected_files = file_picker_window.filePickerWindow(self,subFolder_path).show()
                if selected_files:
                    for file in selected_files:
                        os.startfile(subFolder_path+"\\"+file[1]+file[0], 'open')
                        
    ''' Given a subdirectory and a filetype, call classes that handle file copying.
    
    There are three possible cases for the contents of the folder:
        -There are no files in the folder.
            --This will open an error dialog
        -There is one file in the folder
            --This will push that file without further user input
        --There are multiple files in the folder
            --This will call the filePickerWindow class and have the user
              choose which files to push
    
    PARAMS:
        subDir - the directory to look in for file to open
        fileType - the type of file to open
    '''
    def pushFiles(self,subDir,fileType=""):
        pcase = self.getPCaseString()
        if pcase:
            subFolder_path = "Z:\\IT Documents\\QA\\" + pcase + subDir
            files = self.filesInDir(subFolder_path,fileType)
            if len(files) == 0:
                tk.messagebox.showwarning('Error','No files found in directory')
            elif len(files) == 1 and not os.path.isdir(subFolder_path+'\\'+files[0]):
                copy_files_window.copyFilesWindow(self,subFolder_path,subDir,files)
            else:
                selected_files = file_picker_window.filePickerWindow(self,subFolder_path).show()
                if selected_files:
                    copy_files_window.copyFilesWindow(self,subFolder_path,subDir,selected_files)

    '''Updates Salesforce pulled data in the PCase Info Area
    '''
    def refreshSFInfo(self):
        data_file = self.getDataFile()
        if not os.path.exists(data_file):
            tk.messagebox.showwarning('Error', 'You must first add your sf credentials to\nC:\\Users\\<you>\\AppData\\Roaming\\PCASR\\credentials.txt\nAn example can be found at Z:\\AST\\Utilities\\PCASR')

        else:
            config = configparser.ConfigParser()
            config.read(data_file)

            client = Salesforce(
                username=config.get('credentials','username'),
                password=config.get('credentials','password'),
                security_token=config.get('credentials','security_token')
                )

            pcase = self.getPCaseString()
            if pcase != 'By Using The File Menu':
                case_id = self.json_data[pcase]['sf_link'].split('/')[-1]

                case_info = client.Case.get(case_id)

                self.json_data[pcase]['last_modified'] = case_info['LastModifiedDate']
                self.json_data[pcase]['case_owner'] = case_info['Case_Owner__c']
                self.json_data[pcase]['parent_case_owner'] = case_info['Parent_Case_Owner__c']

                index = self.pcase_list.selection()
                values= self.pcase_list.item(index)['values']
                self.updateInfo(values,self.json_data)
                
                
    def saveJSON(self,file,data):
        with open(file,'w') as outfile:
            json.dump(data,outfile)
            
    #Some Helpful mutator and accessor methods:
    
    def setDataFolder(self,folder):
        self.data_folder = folder
    def getDataFolder(self):
        return self.data_folder
    def setDataFile(self,file):
        self.data_file = file
    def getDataFile(self):
        return self.data_file
    def setArchiveFile(self,file):
        self.archive_file = file
    def getArchiveFile(self):
        return self.archive_file    

    def setPCaseString(self,pcase):
        self.pcase=pcase
    def getPCaseString(self):
        return self.pcase
    def setCustString(self,cust):
        self.cust = cust
    def getCustString(self):
        return self.cust     
    
    # Wrapper methods for each push/edit button

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
        subprocess.Popen(expString)     

    def openSRD(self):
        srd = self.srd_info
        if srd:
            webbrowser.get('chrome').open(srd)
        else:
            tk.messagebox.showwarning('Error', 'URL appears to be invalid.\nPlease enter a valid URL.')

    def openApplication(self,path):
        try:
            tk.messagebox.showinfo('Just a Second...', 'This can take a minute to open, please be patient.')
            os.startfile(path)
        except:
            tk.messagebox.showwarning('Error', 'You must be connected to the VPN to open this.')

    def openWebsite(self,page):
        if page:
            webbrowser.get('chrome').open(page)
        else:
            tk.messagebox.showwarning('Error', 'There is no SRD saved for this case.\nYou can add via Edit > Change PCase Details') 
        

if __name__ == '__main__':
    app = PCaser()
    app.run()

