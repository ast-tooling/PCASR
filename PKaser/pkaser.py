#from cProfile import label
from json.decoder import JSONDecodeError
import tkinter as tk
import subprocess
import multiprocessing
from tkinter import ttk
import tkinter
import webbrowser
import os
import datetime
import re
import json
from simple_salesforce import Salesforce
import configparser
from subprocess import call
from tkinter import messagebox
import time
import datetime
import dateutil.relativedelta
import keyring
import pyperclip

# Called Imports for PKase Notes Code merge.
from tkinter import *
from tkinter import filedialog, TclError
from tkinter.font import Font, families
from tkinter.messagebox import *
from tkinter import font
from textblob import TextBlob
from textblob import Word
from PIL import ImageTk, Image


# local imports
import pk_features.save_dialog_window as save_dialog_window
import pk_features.change_pcase_details_window as change_pcase_details_window
import pk_popouts.file_picker_window as file_picker_window
import pk_popouts.copy_files_window as copy_files_window
import pk_popouts.about_window as about_window
import pk_features.pk_options as pk_options
import pk_popouts.roll_schedule as roll_schedule
import pk_features.timeTracker as timeTracker
import pk_salesforce.salesforceUpdateCommands as sfCase

import pk_multiprocessing.pk_process as pk_process
from pk_wrappers.buttonsWrapper import TkinterCustomButton
from pk_features.codeConflict import *

from pk_features.winMergeU import *




'''The Main Class of the project.  
This both creates the UI for the window and 
drives the logic.
'''
class PCaser:

    #webbrowser.register('windows-default',None,webbrowser.WindowsDefault("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"))
    webbrowser.register('chrome',None,webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
    #webbrowser.register('firefox',None,webbrowser.BackgroundBrowser("C://Program Files //Mozilla Firefox//Chrome//firefox.exe"))
    

    ''' Configure initial variables and call all the methods that instantiate their sub-areas of the UI
    '''
    def __init__(self):

        self.initMainFrame()

        self.all_servers = ["ssnj-imbisc12","ssnj-imbisc13","ssnj-imbisc14","ssnj-imbisc15","ssnj-imbisc20","ssnj-imbisc21"]
        self.json_data = {}
        self.archive_data = {}
        self.server_list = []
        self.data_folder = ""
        self.document_folder = ""
        self.data_file = ""
        self.archive_file = ""
        self.pcase = ""
        self.cust = ""
        self.appIcon_file = ""
        self.config_file = ""
        self.__file__ = None
        self.__firstStartUp = True
        self.first_init = True
        self.b_opt_file_exists = pk_options.checkOptFileExists()
        
        user = os.getlogin()
        data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
        document_folder = r"C:\\Users\\%s\\Documents" %user
        data_file = data_folder+"\\nt-json-files\pkaser.json"
        archive_file = data_folder+"\\nt-json-files\\archive.json"
        appIcon_file = data_folder+"\\images\\billTrustIcon.png"
        branding_image_file = data_folder+"\\images\\inapp-brand-image.png"
        config_file = data_folder+"\\config.txt"
        words_file = data_folder+"\\nt-program-files\\words.txt"
        pcase_notes = document_folder+"\\pcasenotes\\"
        last_selected = data_folder + "\\nt-program-files\\last-selected.txt"
        total_time_spent = data_folder +"\\nt-program-files\\total-time-spent.txt"
        
        
        self.setDataFolder(data_folder)
        self.setDataFile(data_file)
        self.setAppIconFile(appIcon_file)
        self.setBrandingImage(branding_image_file)
        self.setConfigFile(config_file)
        self.setWordsFile(words_file)
        self.setPKaseNotesFolder(pcase_notes)
        self.setLastSelectedFile(last_selected)
        self.setTotalTimeSpent(total_time_spent)
        self.setServerList()


        


        # Call GUI initialization methods
        self.initMenuBar()
        self.initInfoFrame()
        self.initQuickButtons()
        self.initTabArea()
        self.initTreeView()
        self.initCodeConflictFrame()
        self.initNotePadFrame()
        self.initNotePadTextArea()
        self.initBrandingFrame()
        self.initPcaseSearchFrame()

        # Populate the ListBox with PCases
        try:
            self.loadPCases()
        except JSONDecodeError:
            pass

        #self.watching = False
        # Select the first listbox item if there is one
        try:
            with open(self.getLastSelectedFile()) as f:
                try:
                    index = int(f.readline())

                except ValueError as e:
                    tk.messagebox.showerror(message='Error: Value in last-selected.txt maybe NULL: "{}"'.format(e))
            topSelect = topSelect = self.pcase_list.get_children()[index]
            self.pcase_list.selection_set(topSelect)
            #today_date = timeTracker.get_today_date()
            last_selected = "SMILE-ITS_A-NEW-DAY"
            currently_selected = "%s_%s"%(self.pcase_list.item(topSelect)['values'][0],self.pcase_list.item(topSelect)['values'][1])
            
            timeTracker.update_time(currently_selected,last_selected)
            d_time_tracking = timeTracker.load_times()
            timeTracker.calculate_times(d_time_tracking)        
        except IndexError:
            try:
                topSelect = topSelect = self.pcase_list.get_children()[0]
                self.pcase_list.selection_set(topSelect)

                #today_date = timeTracker.get_today_date()
                last_selected = "SMILE-ITS_A-NEW-DAY"
                currently_selected = "%s_%s"%(self.pcase_list.item(topSelect)['values'][0],self.pcase_list.item(topSelect)['values'][1])
                
                timeTracker.update_time(currently_selected,last_selected)
                d_time_tracking = timeTracker.load_times()
                timeTracker.calculate_times(d_time_tracking)

            except IndexError:
                pass


        # Tell the window manager to give focus back after the X button is hit
        self.toplevel.protocol("WM_DELETE_WINDOW", self.kill_window)


    def force_kill_window(self):
        try:
            index_of_last_selected = self.pcase_list.index(self.getPCaseString())
        except:
            index_of_last_selected = "0"

        f = open(self.getLastSelectedFile(), "w")
        f.write(str(index_of_last_selected))
        f.close()
        self.savePCase()
        fileCheck = self.__fileCheck__()
        if fileCheck !="\n":
            currentFileName = "%s_%s.txt"%(self.getPCaseString(),self.getCustString())
            currentTextArea = self.text_area.get("1.0",END)
            self.__saveOnSelect__(currentFileName,currentTextArea)    
        self.toplevel.destroy()
    
    def kill_window(self):
        try:
            index_of_last_selected = self.pcase_list.index(self.getPCaseString())
        except:
            index_of_last_selected = "0"

        f = open(self.getLastSelectedFile(), "w")
        f.write(str(index_of_last_selected))
        f.close()
        self.savePCase()
        fileCheck = self.__fileCheck__()
        if fileCheck !="\n":
            currentFileName = "%s_%s.txt"%(self.getPCaseString(),self.getCustString())
            currentTextArea = self.text_area.get("1.0",END)
            self.__saveOnSelect__(currentFileName,currentTextArea)    
        self.toplevel.destroy()


    def initMainFrame(self):
        self.mainwindow = tk.Tk()
        self.mainwindow.title("The PKaser v2.6")
        self.mainwindow.resizable(height=None,width=None)
        self.mainwindow.geometry('%dx%d+%d+%d' % (1510, 750, 1200, 0))
        self.mainwindow.maxsize(1492, 730)
        
        self.mainwindow.grid_rowconfigure(0, weight=1)
        self.mainwindow.grid_columnconfigure(0, weight=1)
        self.toplevel = self.mainwindow.winfo_toplevel()
        self.userlabel = "Billtrust Team Member - %s"%(os.getlogin())
        self.notepad_frame_label = "%s - PCase ♫'s"%(os.getlogin())

        # Define Row Colors
        self.style = ttk.Style()
        self.style.element_create("plain.field", "from", "clam")
        self.style.layout("EntryStyle.TEntry",
                        [('Entry.plain.field', {'children': [(
                            'Entry.background', {'children': [(
                                'Entry.padding', {'children': [(
                                    'Entry.textarea', {'sticky': 'nswe'})],
                            'sticky': 'nswe'})], 'sticky': 'nswe'})],
                            'border':'2', 'sticky': 'nswe'})])
        
        self.style.configure(('Treeview'),
                background='#ffffff',
                foreground='#000000')
        self.style.configure(('TFrame'),
                background='#f2f2f2',
                foreground='#000000')
        self.style.configure(('TRadiobutton'),
                background='#e6e6e6',
                foreground='#000000')
        self.style.configure(('TLabelFrame'),
                background='#e6e6e6',
                foreground='#000000')
        self.style.configure(('TLabelFrame'),
                background='#e6e6e6',
                foreground='#000000')
        self.style.configure(('TMenubutton'),
                background='#f2f2f2',
                foreground='#000000')
        self.style.configure("EntryStyle.TEntry",
                fieldbackground="white") 
        self.style.map('Treeview',
                    background=[('selected','#003035')])

        # Creation of the FRAMES
        self.main_frame = ttk.Frame(self.mainwindow)
        self.pcase_frame = ttk.Frame(self.main_frame)
        self.pcase_list_frame2 = ttk.Frame(self.main_frame)
        self.info_area_frame = tk.Frame(self.main_frame)
        self.quick_button_frame = ttk.Frame(self.main_frame)

        self.code_conflict_frame = ttk.Frame(self.main_frame)
        self.tab_frame = ttk.Frame(self.main_frame)
        self.notepad_frame = ttk.Frame(self.main_frame)
        self.branding_frame = ttk.Label(self.main_frame)
        self.pcase_search_frame = ttk.Frame(self.main_frame)


        # Placement of Frames in main_frame
        self.info_area_frame.place(x=465, y=0)
        self.quick_button_frame.place(x=1, y=409)
        self.tab_frame.place(x=1, y=0)
        self.pcase_frame.place(x=1228, y=1)
        self.pcase_list_frame2.place(x=1228, y=22)
        self.code_conflict_frame.place(x=1, y=437)
        self.notepad_frame.place(width=802,height=700,x=423, y=50)
        self.main_frame.pack(expand=True, fill=tk.BOTH)
        self.branding_frame.place(x=1, y=665)
        self.pcase_search_frame.place(x=1229, y=0)

        
    def initPcaseSearchFrame(self):
        self.search_frame = self.pcase_search_frame
        self.search_box_entry = ttk.Entry(self.search_frame,style="EntryStyle.TEntry",width=25)
        self.search_box_entry.insert(0,"Type CSR Name or PCase")
        self.search_button = TkinterCustomButton(master=self.search_frame,text="Search ",corner_radius=5,bg_color="#ffffcc",fg_color="#003035",text_color="white",hover_color="#53ba65",width=85,height=23,command=self.search_pcase_list)
        self.search_button.grid(row=0,column=1,padx=10)
        self.search_box_entry.grid(row=0,column=0)

    def search_pcase_list(self):

        pcase_file = self.getDataFile()
        search = self.search_box_entry.get().strip().upper()
        # building out current cases list
        with open(pcase_file) as json_file:
            data = json.load(json_file)    
        current_cases = []
        for case in data:
            current_case = [data[case]['pcase'],data[case]['cust_name'],data[case]['case_number']]
            current_cases.append(current_case)

        x_cases = []
        for x_case in current_cases:
            joined = ','.join([str(i) for i in x_case])
            if search in x_case or joined.find(search) != -1:
                for item in self.pcase_list.get_children():
                    self.pcase_list.delete(item)
                x_cases.append(x_case)

        if x_cases:
            for i_c in x_cases:
                self.pcase_list.insert('','end',iid=i_c[0],values=[i_c[0],i_c[1]])
        else:
            tk.messagebox.showwarning('Information','Case Not Found!')

        if search == "":
            self.search_box_entry.insert(0,"Type CSR Name or PCase")
            for item in self.pcase_list.get_children():
                self.pcase_list.delete(item)
            for case in data:
                self.pcase_list.insert('','end',iid=[data[case]['pcase']],values=[data[case]['pcase'],data[case]['cust_name']])




  
    def initBrandingFrame(self):
        self.brand_frame = self.branding_frame

        loadImage = Image.open(self.getBrandingImage())
        renderImage = ImageTk.PhotoImage(loadImage)
        self.branding_image = TkinterCustomButton(master=self.brand_frame,bg_color="#ffffcc",border_width=1,fg_color="#003035",corner_radius=0,text_color="white",hover_color="#aadb1e",width=416,height=53,image=renderImage ,command=lambda: self.openWebsite("https://billtrust.kazoohr.com/dashboard"))
        self.branding_image.pack()

    def initTabArea(self):
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
        self.tab_area = ttk.Notebook(self.tab_frame,width=420,height=380)

        self.tab_1 = tk.Frame(self.tab_area)
        self.tab_2 = tk.Frame(self.tab_area)
        self.tab_3 = tk.Frame(self.tab_area)
        self.tab_4 = tk.Frame(self.tab_area)

        self.tab_area.add(self.tab_1,text="PCase Folder")
        self.tab_area.add(self.tab_2,text="Time Tracker")
        self.tab_area.add(self.tab_3,text="Edit/Push/Diff/Opts")
        self.tab_area.pack(expand=1,fill="both")
  
        # Tab 3 - Edit/Push/Diff/Opts Windows
        self.top_frame = ttk.LabelFrame(self.tab_3)
        self.opts_frame = ttk.LabelFrame(self.tab_3,text="PKaser Options")

        self.top_frame.pack(ipady=5,ipadx=5,padx=5)
        self.opts_frame.pack(ipady=5,ipadx=5,padx=5,expand=True,fill='both')

        #  start - Top Half -  Edit/Push/Diff 
        # Buttons
        self.editTemplateButton = ttk.Button(self.top_frame, text="Edit",command=self.editTemplates)
        #self.button1 = TkinterCustomButton(master=self.top_frame,text="Edit",bg_color="#e6e6e6",fg_color="#e6e6e6",corner_radius=10,text_color="white",hover_color="#53ba65",width=65,height=20,command=self.editTemplates)
        self.editParserButton = ttk.Button(self.top_frame, text="Edit",command=self.editParsers)
        #self.button2 = TkinterCustomButton(master=self.top_frame,text="Edit",bg_color="#ffffcc",fg_color="#003035",corner_radius=10,text_color="white",hover_color="#53ba65",width=65,height=20,command=self.editParsers)
        self.editScriptButton = ttk.Button(self.top_frame, text="Edit",command=self.editScripts)
        #self.button3 = TkinterCustomButton(master=self.top_frame,text="Edit",bg_color="#ffffcc",fg_color="#156184",corner_radius=10,text_color="white",hover_color="#53ba65",width=65,height=20,command=self.editScripts)
        self.editSampleButton = ttk.Button(self.top_frame, text="Edit",command=self.editSamples)
        #self.button4 = TkinterCustomButton(master=self.top_frame,text="Edit",bg_color="#ffffcc",fg_color="#003035",corner_radius=10,text_color="white",hover_color="#53ba65",width=65,height=20,command=self.editSamples)
        self.editReleaseButton = ttk.Button(self.top_frame, text="Edit",command=self.editRelease)
        #self.button5 = TkinterCustomButton(master=self.top_frame,text="Edit",bg_color="#ffffcc",fg_color="#156184",corner_radius=10,text_color="white",hover_color="#53ba65",width=65,height=20,command=self.editRelease)
        self.pushTemplateButton = ttk.Button(self.top_frame, text="Push",command=self.pushTemplates)
        #self.button6 = TkinterCustomButton(master=self.top_frame,text="Push",bg_color="#ffffcc",fg_color="#003035",corner_radius=10,text_color="white",hover_color="#53ba65",width=65,height=20,command=self.pushTemplates)
        self.pushParserButton = ttk.Button(self.top_frame, text="Push",command=self.pushParsers)
        #self.button7 = TkinterCustomButton(master=self.top_frame,text="Push",bg_color="#ffffcc",fg_color="#156184",corner_radius=10,text_color="white",hover_color="#53ba65",width=65,height=20,command=self.pushParsers)
        self.pushScriptsButton = ttk.Button(self.top_frame, text="Push",command=self.pushScripts)
        #self.button8 = TkinterCustomButton(master=self.top_frame,text="Push",bg_color="#ffffcc",fg_color="#003035",corner_radius=10,text_color="white",hover_color="#53ba65",width=65,height=20,command=self.pushScripts)
        self.pushSamplesButton = ttk.Button(self.top_frame, text="Push",command=self.pushSamples)
        #self.button9 = TkinterCustomButton(master=self.top_frame,text="Push",bg_color="#ffffcc",fg_color="#003035",corner_radius=10,text_color="white",hover_color="#53ba65",width=65,height=20,command=self.pushSamples)
        self.pushReleaseButton = ttk.Button(self.top_frame, text="Push",command=self.pushRelease)
        self.diffTemplateButton = ttk.Button(self.top_frame, text="Diff",command=self.diffTemplates)
        self.diffParserButton = ttk.Button(self.top_frame, text="Diff", command=self.diffParser)
        self.diffScriptButton = ttk.Button(self.top_frame, text="Diff",command=self.diffScripts)
        self.diffReleaseButton = ttk.Button(self.top_frame, text="Diff", command=self.diffReleaseBin)
        self.blankButton = ttk.Button(self.top_frame)

        # Labels - Top Half -  Edit/Push/Diff 
        self.templateLabel = ttk.Label(self.top_frame,text="Templates")
        self.parserLabel = ttk.Label(self.top_frame,text="ParserConfig")
        self.scriptLabel = ttk.Label(self.top_frame,text="Script")
        self.sampleDataLabel = ttk.Label(self.top_frame,text="SampleData")
        self.releaseBinLabel = ttk.Label(self.top_frame,text="ReleaseBin")

        # Grid Everything - Top Half -  Edit/Push/Diff 
        self.editTemplateButton.grid(row=1,column=1)
        self.editParserButton.grid(row=2,column=1)
        self.editScriptButton.grid(row=3,column=1)
        self.editSampleButton.grid(row=4,column=1)
        self.editReleaseButton.grid(row=5,column=1)
        self.pushTemplateButton.grid(row=1,column=2)
        self.pushParserButton.grid(row=2,column=2)
        self.pushScriptsButton.grid(row=3,column=2)
        self.pushSamplesButton.grid(row=4,column=2)
        self.pushReleaseButton.grid(row=5,column=2)
        self.diffTemplateButton.grid(row=1,column=3)
        self.diffParserButton.grid(row=2,column=3)
        self.diffScriptButton.grid(row=3,column=3)
        self.blankButton.grid(row=4,column=3)
        self.diffReleaseButton.grid(row=5,column=3)
        self.templateLabel.grid(row=1,column=0, padx=50)
        self.parserLabel.grid(row=2,column=0)
        self.scriptLabel.grid(row=3,column=0)
        self.sampleDataLabel.grid(row=4,column=0)
        self.releaseBinLabel.grid(row=5,column=0)
        #  end  - Top Half -  Edit/Push/Diff


        # Bottom Half - Opts
    
        self.opts_frame.grid_propagate(1)

        self.jobTitleLabel = Label(self.opts_frame, text="Job Title")
        self.jobclicked = StringVar()
        self.trackTimeLabel = Label(self.opts_frame, text="Track Time")
        self.trackTimeClicked = StringVar()

        if not self.b_opt_file_exists:
            print("if not self.b_opt_file_exists")
            self.jobclicked.set(pk_options.jobTitles()[0])
            self.trackTimeClicked.set(pk_options.trackTimeOpts()[0])
            self.jobTitleMenu = OptionMenu(self.opts_frame, self.jobclicked, *pk_options.jobTitles(), command=pk_options.getjobTitle)
            self.trackTimeMenu = OptionMenu(self.opts_frame, self.trackTimeClicked, *pk_options.trackTimeOpts(), command=pk_options.getTrackTimeValue)
            pk_options.createOptFile()

        elif self.b_opt_file_exists:
            print("elif self.b_opt_file_exists")
            opts = pk_options.loadOptFile()
            jobTitle = opts["JOB_TITLE"]
            trackTime = opts["TRACK_TIME"]
            self.jobclicked.set(jobTitle)
            self.trackTimeClicked.set(trackTime)
            self.jobTitleMenu = OptionMenu(self.opts_frame, self.jobclicked, *pk_options.jobTitles(), command=pk_options.getjobTitle)
            self.trackTimeMenu = OptionMenu(self.opts_frame, self.trackTimeClicked, *pk_options.trackTimeOpts(),command=pk_options.getTrackTimeValue)



        self.jobTitleLabel.grid(row=1,column=1)
        self.jobTitleMenu.grid(row=1,column=2)
        self.trackTimeLabel.grid(row=2,column=1)
        self.trackTimeMenu.grid(row=2, column=2)
        self.trackTimeMenu.configure(state='disabled')




    def initTreeView(self):
         # Tab Definitions

        self.pcase_tab_area = ttk.Frame(self.pcase_list_frame2,width=245,height=650)

        # Define active tab contents
        columns=["PCase","CSR Name"]
        #self.pcase_list = ttk.Treeview(self.current_tab,height=31,columns=columns,show="headings")
        self.pcase_list = ttk.Treeview(self.pcase_tab_area,height=34,columns=columns,show="headings")
        self.veritcalScrollbar = ttk.Scrollbar(self.pcase_list_frame2, orient='vertical',command=self.pcase_list.yview)
        self.pcase_list['yscrollcommand'] = self.veritcalScrollbar.set
        self.veritcalScrollbar.grid(row=0,column=3,sticky='ns')
        self.pcase_list.bind('<<TreeviewSelect>>',self.onselect)
        
        self.pcase_list.column('PCase',width=78,stretch=False,minwidth=78)
        self.pcase_list.column('CSR Name',width=163,stretch=False,minwidth=100)

        self.pcase_list.heading('PCase',text="PCase")
        self.pcase_list.heading('CSR Name',text="CSR Name")

        # Enable sorting
        for col in columns:
            self.pcase_list.heading(col,command=lambda _col=col: self.treeview_sort_column(_col, False,self.pcase_list))

        self.pcase_list.grid(row=0,column=0)

        # Add function buttons
        self.update_button = TkinterCustomButton(master=self.pcase_list_frame2,text='Update',bg_color="#ffffcc",fg_color="#003035",corner_radius=5,text_color="white",hover_color="#53ba65",width=80,height=24,command=self.editWindow)
        self.new_button = TkinterCustomButton(master=self.pcase_list_frame2,text='New',bg_color="#ffffcc",fg_color="#003035",corner_radius=5,text_color="white",hover_color="#53ba65",width=80,height=24,command=self.newWindow)
    
        # Add items to frame
        self.pcase_tab_area.grid(row=0,column=0,columnspan=3)




    def initQuickButtons(self):

        self.quick_button_frame.grid_propagate(True)
        self.pcase_button = TkinterCustomButton(master=self.quick_button_frame,text="PCase",bg_color="#ffffcc",fg_color="#003035",corner_radius=5,text_color="white",hover_color="#53ba65",width=90,height=24,command=lambda: self.openDir("Z:\\IT Documents\\QA\\" + self.getPCaseString()))
        self.srd_button = TkinterCustomButton(master=self.quick_button_frame,text="SRD",bg_color="#ffffcc",fg_color="#003035",corner_radius=5,text_color="white",hover_color="#53ba65",width=90,height=24,command=lambda: self.openWebsite(self.json_data[self.getPCaseString()]['srd_link']))
        self.ftp_root_button = TkinterCustomButton(master=self.quick_button_frame,text="FTP Root",bg_color="#ffffcc",fg_color="#003035",corner_radius=5,text_color="white",hover_color="#53ba65",width=90,height=24,command=lambda: self.openDir("\\\\ssnj-isilon02\\imtest\\imstage01\\ftproot\\"+self.getCustString() ))
        self.sf_button = TkinterCustomButton(master=self.quick_button_frame,text="SalesForce",bg_color="#ffffcc",fg_color="#003035",corner_radius=5,text_color="white",hover_color="#53ba65",width=110,height=24,command=lambda: self.openWebsite(self.json_data[self.getPCaseString()]['sf_link']))


        self.pcase_button.grid(row=0,column=0,sticky="E",padx=5)
        self.ftp_root_button.grid(row=0,column=1,sticky="E",padx=5)
        self.sf_button.grid(row=0,column=2,sticky="E",padx=5)
        self.srd_button.grid(row=0,column=3,sticky="E",padx=5)
        
    def initCodeConflictFrame(self):
        self.conflict_frame = self.code_conflict_frame
        self.code_conflict_label = tk.Label(self.conflict_frame, justify='center', wraplength=400)
        self.code_conflict_label.grid(row=1,column=0,padx=1,sticky="W")

    def initNotePadFrame(self):
        self.notepad_frame = self.notepad_frame
    

    def initInfoFrame(self):
        # Initialize Info Section contents
        self.info_frame = self.info_area_frame
        # Do NOT MOVE OR DELETE --- START
        self.pcase_info = ttk.Label(self.info_frame,text="Welcome to The PKaser",width=23)
        self.cust_info = ttk.Label(self.info_frame,text="Click PCase > New to Add",width=23)
        self.sf_info = ttk.Label(self.info_frame,text="",width=23)
        self.desc_info = tk.Label(self.info_frame,text="",font=("Arial",9,"bold"),width=75, justify="center")
        self.last_mod_info = ttk.Label(self.info_frame,text="",width=20)
        self.owner_info = ttk.Label(self.info_frame,text="",width=20)
        self.case_owner_info = ttk.Label(self.info_frame,text="",width=20)
        self.last_mod_label = ttk.Label(self.info_frame,text="Last Modified: ",width=15)
        self.owner_label = ttk.Label(self.info_frame,text="Current Owner: ",width=15)
        self.case_owner_label = ttk.Label(self.info_frame,text="Case Owner: ",width=15)
         # Do NOT MOVE OR DELETE --- END

    def updateInfoBox(self,values,data):
        try:
            index = self.pcase_list.selection()
            values= self.pcase_list.item(index)['values']
            pcase = values[0]
            sf_case_sub = data[pcase]['subject']
            sf_pcase_number = data[pcase]['pcase']
            sf_case_number = data[pcase]['case_number']
            sf_last_modified = data[pcase]['last_modified'].split('.')[0][:-3].replace('T'," ")
            sf_parent_case_owner = data[pcase]['parent_case_owner']
            sf_current_case_owner = data[pcase]['case_owner']
            sf_csrname = data[pcase]['cust_name']
        except KeyError:
            print("UpdateInfoBox Key Error. Need to fix this part.")



        sf_case_info = "Case Number: %s / Case Owner: %s / Queue: %s"%(sf_case_number,sf_parent_case_owner,sf_current_case_owner)
        
        self.pcase_sub_label = ttk.Label(self.info_frame, text ='Subject:')
        self.pcase_info_label = ttk.Label(self.info_frame, text ='Case Info:')
        self.csr_box_label = ttk.Label(self.info_frame, text ='Cust:')

        self.pcase_sub_box = ttk.Entry(self.info_frame,style="EntryStyle.TEntry",width=85)
        self.pcase_sub_box.insert(0,sf_case_sub)
        self.pcase_sub_box.configure(state='readonly')
        
        self.pcase_info_box = ttk.Entry(self.info_frame,style="EntryStyle.TEntry",width=85)
        self.pcase_info_box.insert(0,sf_case_info)
        self.pcase_info_box.configure(state='readonly')

        self.csr_box = ttk.Entry(self.info_frame,style="EntryStyle.TEntry",width=25)
        self.csr_box.insert(0,sf_csrname)
        self.csr_box.configure(state='readonly')


        self.refresh_button = TkinterCustomButton(master=self.info_frame,text="Refresh",corner_radius=5,bg_color="#ffffcc",fg_color="#003035",text_color="white",hover_color="#53ba65",width=85,height=24,command=self.refreshSFInfo)

        # Define Grid for Info Section Items
        self.pcase_sub_label.grid(row=1,column=0)
        self.pcase_sub_box.grid(row=1,column=1)
        self.csr_box_label.grid(row=1, column=2)
        self.csr_box.grid(row=1,column=3)
        self.pcase_info_label.grid(row=2,column=0)   
        self.pcase_info_box.grid(row=2,column=1)
        self.refresh_button.grid(row=2,column=3,padx=10, pady=2)
        


        


    def initMenuBar(self):
        # Initialize Menu Bar and PKase Options Menu
        self.menu_bar = tk.Menu(self.mainwindow)
        self.toplevel.config(menu=self.menu_bar)
        self.menu_item_1 = tk.Menu(self.menu_bar,tearoff=False)
        self.menu_item_2 = tk.Menu(self.menu_bar,tearoff=False)
        self.menu_item_3 = tk.Menu(self.menu_bar,tearoff=False)
        self.menu_item_4 = tk.Menu(self.menu_bar,tearoff=False)

        # PKase Notes Options Menu
        self.file_menu        = Menu(self.menu_bar, tearoff=0)
        self.edit_menu        = Menu(self.menu_bar, tearoff=0)
        self.tool_menu        = Menu(self.menu_bar, tearoff=0)
        self.help_menu        = Menu(self.menu_bar, tearoff=0)
        self.sf_menu          = Menu(self.menu_bar, tearoff=0)
        self.ts_menu          = Menu(self.menu_bar, tearoff=0)
        

        # Add Menu Options to Bar
        self.menu_bar.add_cascade(label="PKaser Options:")
        self.menu_bar.add_cascade(label="PCase", menu=self.menu_item_1)
        self.menu_bar.add_cascade(label="Edit", menu=self.menu_item_2)
        self.menu_bar.add_cascade(label="Apps", menu=self.menu_item_3)
        self.menu_bar.add_cascade(label="Help", menu=self.menu_item_4)
        self.menu_bar.add_cascade(label="                                                    ")
        self.menu_bar.add_cascade(label="PCase ♫'s Options:")
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        self.menu_bar.add_cascade(label="Edit", menu=self.edit_menu)
        self.menu_bar.add_cascade(label="Tools", menu=self.tool_menu)
        self.menu_bar.add_cascade(label="Update SF", menu=self.sf_menu)
        self.menu_bar.add_cascade(label="Timestamps", menu=self.ts_menu)

        # Additional Menu Options
        self.right_click_menu_notepad = Menu(self.mainwindow, tearoff=0)
        self.right_click_menu_pcase_folder = Menu(self.mainwindow, tearoff=0)



        # Add subitems for File Option
        data_folder = self.getDataFolder()
        user = os.getlogin()
        self.file_menu.add_command(label="Open PCase ♫'s Directory",command=lambda: self.openDir("C:\\Users\\%s\\Documents\\pcasenotes" %user))
        self.file_menu.add_command(label="Search PCase ♫'s Directory",command=lambda:self.openApplication("Z:\\AST\\Utilities\\FindAndReplace\\fnr\\fnr.exe"))
        # Add subitems for Edit Option
        self.edit_menu.add_command(label="Undo",accelerator="Ctrl+Z",command=self.__undo__)
        self.edit_menu.add_command(label="Redo",accelerator="Ctrl+Y",command=self.__redo__)
        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="Cut",accelerator="Ctrl+X", command=self.__cut__)
        self.edit_menu.add_command(label="Copy", accelerator="Ctrl+C",command=self.__copy2__)
        self.edit_menu.add_command(label="Paste", accelerator="Ctrl+V",command=self.__paste__)
        
        # Add subitems for Tools Option
        self.tool_menu.add_command(label="Copy PCase",accelerator="Alt+P", command=self.copyPCase)
        self.tool_menu.add_command(label="Copy SFCase",accelerator="Alt+F", command=self.copySFCase)
        self.tool_menu.add_command(label="Copy PCase_CSRNAME",accelerator="Alt+B",command=self.copyPCaseAndCSRName)
        self.tool_menu.add_command(label="Copy CSRNAME",accelerator="Alt+C",command=self.copyCSRName)
        self.tool_menu.add_separator()
        self.tool_menu.add_command(label="Python Comment",accelerator="Alt+O",command=self.pythonComment)
        
        # Add subitems for Update SF OPtion
        self.sf_menu.add_command(label="Text To Comment",command=lambda: sfCase.addCaseComment(self.returnParentId(),self.__returnSelectedText__())) # lambda:
        self.sf_menu.add_command(label="EIPP-Comment", command=self.__sentToEIPP__)
        self.sf_menu.add_command(label="FCI-Comment", command=self.__sentToFCI__)
        self.sf_menu.add_command(label="Committed-Comment", command=self.__CaseCommittedNoteStamp__)
        self.sf_menu.add_command(label="Reviewed OnCall-Comment", command=self.__callReviewCompleted__)
        self.sf_menu.add_command(label="Pending Code Review-Comment", command=self.__PendingASECodeReview__)
        

        # Add subitems for Timestamp Option
        self.ts_menu.add_command(label="Insert-Timestamp", command=self.__regularNoteStamp__)
        self.ts_menu.add_command(label="Insert-Slack-Comment", command=self.__checkSlackStamp__)
        self.ts_menu.add_command(label="Insert-SF-Comment", command=self.__salesforceComment__)
        self.ts_menu.add_command(label="Insert-FF-Comment", command=self.__FileFailureComment__)
  



        self.toplevel.bind_all("<Control-n>",self.newWindowWrapper)
        self.toplevel.bind_all("<Control-c>",self.copyWrapper)
        self.toplevel.bind_all("<Control-x>",self.cutWrapper)
        self.toplevel.bind_all("<Alt-p>",self.copyPCaseWrapper)
        self.toplevel.bind_all("<Alt-f>",self.copySFCaseWrapper)
        self.toplevel.bind_all("<Alt-b>",self.copyPCaseAndCSRNameWrapper)
        self.toplevel.bind_all("<Alt-c>",self.copyCSRNameWrapper)

        # Right click options
        #self.toplevel.bind_all("<Button-3>", self.rightClickMenu)

        # Add 1 - File Menu Subitems
        self.menu_item_1.add_command(label="New",accelerator="Ctrl+N",command=self.newWindow)
        #self.menu_item_1.add_command(label="Delete",command=self.deleteCase)
        self.menu_item_1.add_separator()
        self.menu_item_1.add_command(label="Quit",accelerator="Ctrl+Q",command=self.force_kill_window)

        # Add 2- Edit Menu Subitems
        self.menu_item_2.add_command(label="Take Ownership",command=lambda: sfCase.takeOwnership(self.json_data[self.getPCaseString()]['sf_link']))
        self.menu_item_2.add_command(label="Change PCase Details",command=self.editWindow)
        self.menu_item_2.add_command(label="Save Pkaser Options")

        # Add 3- Tools Menu Subitems
        self.menu_item_3.add_command(label="Exif Tool",command=lambda: self.openApplication("Z:\\AST\\Utilities\\exif_tool\\exif_tool.bat"))
        self.menu_item_3.add_command(label="Diff Pdf",command=lambda: self.openApplication("Z:\\AST\\DiffPDF\\DiffpdfPortable\\DiffpdfPortable.exe"))
        self.menu_item_3.add_command(label="Parseamajig",command=lambda: self.openApplication("Z:\\AST\\Utilities\\Parseamajig\\Parseamajig.exe"))
        self.menu_item_3.add_command(label="XMLGenerator",command=lambda: self.openApplication("Z:\\AST\\Utilities\\XMLGenerator\\XmlGenerator.exe"))
        self.menu_item_3.add_command(label="BillGen Wrapper",command=lambda: self.openApplication("Z:\\AST\\Utilities\\BillGen\\BillGenWrapper.exe"))
        self.menu_item_3.add_command(label="PrePost Compare",command=lambda: self.openWebsite("http://ssnj-devast01/prepost#"))
        self.menu_item_3.add_command(label="Make PCase Dir",command=lambda: self.openApplication("C:\\AST\\utils\\scripts\\case_dir.bat"))
                
        # Add 4- Help Menu Subitems
        self.menu_item_4.add_command(label="About", command=lambda:about_window.aboutWindow(self))
        self.menu_item_4.add_command(label="Confluence Page",command=lambda: self.openWebsite("https://billtrust.atlassian.net/wiki/spaces/AT/pages/29252616317/PKaser+FAQ+Page"))
        self.menu_item_4.add_command(label="Roll Schedule", command=lambda: roll_schedule.roll_schedule())


        #self.right_click_menu.add_command(label="Quick Comment")

    def initNotePadTextArea(self):
        self.scroll_bar       = Scrollbar(self.notepad_frame)
        self.text_area        = Text(self.notepad_frame,wrap=WORD,width=111,height=37,font=('Calibri',11),undo=True)
        self.text_area.edit_modified(True)
        self.file             = None

        self.text_area.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        self.scroll_bar.grid(row=0, column=1,sticky='ns')
        self.text_area['yscrollcommand'] = self.scroll_bar.set
        

        self.scroll_bar.columnconfigure(0,weight=1)
        self.scroll_bar.config(command=self.text_area.yview)

        self.bold_font = font.Font(self.text_area, self.text_area.cget("font"))
        self.bold_font.configure(weight="bold")
        self.text_area.tag_configure("misspelled",foreground="red", underline=True)
        #self.text_area.bind("<space>", self.__spellCheck__)



        # initialize the spell checking dictionary.
        self.words_file = open(self.getWordsFile()).read().split("\n")

        # Right Click Menu Options
        self.right_click_menu_notepad.add_command(label="Text To Comment",command=lambda: sfCase.addCaseComment(self.returnParentId(),self.__returnSelectedText__()))
        self.right_click_menu_notepad.add_command(label="Cut",command=self.__cut__)
        self.right_click_menu_notepad.add_command(label="Copy",command=self.__copy2__)
        self.right_click_menu_notepad.add_command(label="Paste",command=self.__paste__)
        self.text_area.bind("<Button-3>", self.rightClickMenuNotePad)

   
    def treeview_sort_column(self,col, reverse,treeview):
        l = [(treeview.set(k, col), k) for k in treeview.get_children('')]
        l.sort(reverse=reverse)

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            treeview.move(k, '', index)

        # reverse sort next time
        treeview.heading(col, command=lambda _col=col: self.treeview_sort_column(_col, not reverse))


    def newWindowWrapper(self,parent):
        self.newWindow()

    def killWindowWrapper(self,parent):
        self.kill_window()
    def copyWrapper(self, parent):
        self.__copy2__()
    
    def pasteWrapper(self, parent):
        self.__paste__()



    def cutWrapper(self, parent):
        self.__cut__
    def copyPCaseWrapper(self,parent):
        self.copyPCase()
    def copySFCaseWrapper(self,parent):
        self.copySFCase()
    def copyPCaseAndCSRNameWrapper(self,parent):
        self.copyPCaseAndCSRName()
    def copyCSRNameWrapper(self,parent):
        self.copyCSRName()

    

        
    '''Called by the treeview when a new selection is made.
    
    Saves PCase data to json, disables ftp root watchdog, calls updateInfo
    '''
    def onselect(self,parent):
        self.onselect = True
        index = self.pcase_list.selection()
        values= self.pcase_list.item(index)['values']
        pcase = values[0]
        if self.__firstStartUp == True:
            pk_process.runWithThreads(self.setPCaseString(pcase),self.updateInfoBox(values,self.json_data),self.updateInfo(values,self.json_data),self.__openOnSelect__())
            self.__firstStartUp = False
        else:
            previousFileName = "%s_%s.txt"%(self.getPCaseString(),self.getCustString())
            previousTextArea = self.text_area.get("1.0",END)

            #set up for timeTracker feature
            today_date = timeTracker.get_today_date()
            last_selected = "%s_%s"%(self.getPCaseString(),self.getCustString())
            currently_selected = "%s_%s"%(values[0],values[1])
            print("today_date: %s"%today_date)
            print("last_selected: %s"%last_selected)
            print("currently_selected: %s"%currently_selected)

            timeTracker.update_time(currently_selected,last_selected)
            d_time_tracking = timeTracker.load_times()
            timeTracker.calculate_times(d_time_tracking)
            pk_process.runWithThreads(self.setPCaseString(pcase),self.updateInfoBox(values,self.json_data),self.updateInfo(values,self.json_data),self.__saveOnSelect__(previousFileName,previousTextArea), self.__openOnSelect__(),self.savePCase())
            
            
            

    '''Updates the PCase Info area, quick view, edit/push, and notes
    
    PARAMS:
        values - a 1x2 list containing the pcase and the customer name
        data - either json_data or archive_data, the json with case data in memory
    '''
    def updateInfo(self,values,data):
        pcase = values[0]
        cust_name = values[1]
        case_number = data[pcase]['case_number']
        
        self.setPCaseString(pcase)
        self.setCustString(cust_name)
        self.setCaseString(case_number)

        self.pcase_info.config(text=pcase)
        self.cust_info.config(text=cust_name)

        self.sf_info.config(text=data[pcase]['case_number'])
        self.desc_info.config(text=data[pcase]['subject'])
        self.last_mod_info.config(text=data[pcase]['last_modified'].split('.')[0][:-3].replace('T'," "))
        self.owner_info.config(text=data[pcase]['case_owner'])
        self.case_owner_info.config(text=data[pcase]['parent_case_owner'])

        print("self.time_tracker_tree() ran in updateInfo function.-------------")

        # self.time_tracker_tree()
        # self.pcase_tree()
        # self.code_conflict_tree()
        pk_process.runWithThreads(self.time_tracker_tree(),self.pcase_tree(),self.code_conflict_tree())
        
        
    def savePCase(self):
        print("savePCase(self)- ran")
        data_folder = self.getDataFolder()
        data_file = self.getDataFile()
        data = self.json_data
        pcase = self.getPCaseString()


        try:
            self.archive_data[pcase]
            archive_flag = True
        except:
            archive_flag = False

        if data:
            self.saveJSON(data_file,data)
    '''Loads JSON file into memory if not already done, updates pcase_list treeview
    '''
    def loadPCases(self):
        try:
            data_file = self.getDataFile()

            if os.path.exists(data_file):
                self.pcase_list.delete(*self.pcase_list.get_children())
                if not self.json_data:
                    with open(data_file) as json_file:
                        data = json.load(json_file)
                        self.json_data = data
    
                        for case in data:
                            self.pcase_list.insert('','end',iid=[data[case]['pcase']],values=[data[case]['pcase'],data[case]['cust_name']])
                else:
                    data = self.json_data
                    for case in data:
                        self.pcase_list.insert('','end',iid=[data[case]['pcase']],values=[data[case]['pcase'],data[case]['cust_name']])
        except:
            tk.messagebox.showwarning('Error', 'Issue loading pkaser.json file located in nt-json files.\nTypically this is caused bya faulty load.')
            print("issues with json file.")
    

    def newWindow(self):
        save_dialog_window.saveDialogWindow(self,False)

    def editWindow(self):
        change_pcase_details_window.saveDialogWindow(self,True)

    def run(self):
        appIcon_file = self.getAppIconFile()
        iconPhotoImage = tk.PhotoImage(file = appIcon_file)

        self.mainwindow.iconphoto(False,iconPhotoImage)
        #self.mainwindow.attributes('-fullscreen',True)
        self.mainwindow.mainloop()
        
    '''Loads 'Quick View' PCase treeview
    '''
    def pcase_tree(self):
        pcase = self.getPCaseString()
        def copy_path():
            item = self.tree.selection()[0]
            parent_iid = self.tree.parent(item)
            node = []
            path = ""
           # go backward until reaching root
            while parent_iid != '':
                node.insert(0, self.tree.item(parent_iid)['text'])
                parent_iid = self.tree.parent(parent_iid)
                i = self.tree.item(item, "text")
                path = os.path.join(*node, i)
                self.copyInput(path)
                print(path)
            return path
        def copy_filename():
            item = self.tree.selection()[0]
            parent_iid = self.tree.parent(item)
            node = []
           # go backward until reaching root
            while parent_iid != '':
                node.insert(0, self.tree.item(parent_iid)['text'])
                parent_iid = self.tree.parent(parent_iid)
                i = self.tree.item(item, "text")
                path = os.path.join(*node, i)
                isFile = os.path.isfile(path)
                print(isFile)
                if isFile:
                    filename = i
                    self.copyInput(filename)
                    print(filename)
            
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
            text_display = "PCase Number: " + pcase
            self.tree.heading('#0', text=text_display, anchor='w')

            self.ysb.pack(side=tk.RIGHT,fill=tk.Y)
            self.tree.pack(fill=tk.BOTH)
            

            abspath = "Z:\\IT Documents\\QA\\" + pcase
            self.insert_node('', abspath, abspath)
            self.tree.bind('<<TreeviewOpen>>', self.open_node)

            self.tree.focus(self.tree.get_children()[0])
            self.tree.item(self.tree.focus(),open=True)
            self.open_node('')
        
        if self.__firstStartUp:
            print("***PCase Folder Right Click Menu Run***")
            self.right_click_menu_pcase_folder.add_command(label="Copy Path", command=lambda: copy_path())
            self.right_click_menu_pcase_folder.add_command(label="Copy Filename", command=lambda: copy_filename())
            self.right_click_menu_pcase_folder.add_command(label="Open Resource", command=lambda: self.openDir(str(copy_path())))
        self.tree.bind("<Button-3>",  self.rightClickMenuPCaseFolder)

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
            print(subDir)
            files = self.filesInDir(subFolder_path,fileType)
            if len(files) == 0:
                tk.messagebox.showwarning('Error','No files found in directory')
            elif len(files) == 1:
                print(files)
                print(subDir)
                print(files[0])
                os.startfile(subFolder_path+"\\"+files[0], 'open')
                print(subFolder_path+"\\"+files[0])
            else:
                selected_files = file_picker_window.filePickerWindow(self,subFolder_path).show()
                print(selected_files)
                if selected_files:
                    for file in selected_files:
                        os.startfile(subFolder_path+"\\"+file[1]+file[0], 'open')
                        print(subFolder_path+"\\"+file[1]+file[0])


    def diffFiles(self,subDir,fileType=""):
        pcase = self.getPCaseString()
        if pcase:
            subFolder_path = "Z:\\IT Documents\\QA\\" + pcase + subDir
            print(subFolder_path)
            if fileType == ".csv":
                astFolder_path = "C:\\AST\\FDT\\"
            elif fileType == ".xml":
                astFolder_path = "C:\\AST\\ParserConfigs\\"
            elif fileType == ".py":
                astFolder_path = "C:\\AST\\scripts\\"
            elif fileType == "":
                astFolder_path = "\\AST\\ReleaseBin\\"

            files = self.filesInDir(subFolder_path,fileType)
            print(files)
            if len(files) == 0:
                tk.messagebox.showwarning('Error','No files found in directory')
            elif len(files) == 1:
                winMergeOpen(astFolder_path+"\\"+files[0],subFolder_path+"\\"+files[0])
            else:
                selected_files = file_picker_window.filePickerWindow(self,subFolder_path).show()
                if selected_files:
                    for file in selected_files:
                        winMergeOpen(astFolder_path+"\\"+file[0],subFolder_path+"\\"+file[1]+file[0])
                        
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
                print("pkaser.py copy file window values: %s, %s, %s"%(subFolder_path,subDir, files))
            else:
                selected_files = file_picker_window.filePickerWindow(self,subFolder_path).show()
                if selected_files:
                    copy_files_window.copyFilesWindow(self,subFolder_path,subDir,selected_files)

    '''Updates Salesforce pulled data in the PCase Info Area
    '''
    def refreshSFInfo(self):
        config_file = self.getConfigFile()
        if not os.path.exists(config_file):
            tk.messagebox.showwarning('Error', 'You must first add your sf credentials to\nC:\\Users\\<you>\\AppData\\Roaming\\PKaser\\credentials.txt\nAn example can be found at Z:\\AST\\Utilities\\PKaser')

        else:
            config = configparser.ConfigParser()
            config.read(config_file)
            username = config.get('credentials','username')

            client = Salesforce(
                username= username,
                password=keyring.get_password("pkaser-userinfo", username),
                security_token=keyring.get_password("pkaser-token", username)
                )

            pcase = self.getPCaseString()
            if pcase != 'Welcome to The PKaser':
                case_id = self.json_data[pcase]['sf_link'].split('/')[-1]
                print(case_id)
                if 'view' in case_id:
                    print('lighting')
                    case_id = self.json_data[pcase]['sf_link'].split('/')[6]
                case_info = client.Case.get(case_id)

                self.json_data[pcase]['last_modified'] = case_info['LastModifiedDate']
                self.json_data[pcase]['case_owner'] = case_info['Case_Owner__c']
                self.json_data[pcase]['parent_case_owner'] = case_info['Parent_Case_Owner__c']

                index = self.pcase_list.selection()
                values= self.pcase_list.item(index)['values']
                self.updateInfo(values,self.json_data)
                self.updateInfoBox(values,self.json_data)
            else:
                pass

    def returnParentId(self):
        pcase = self.getPCaseString()
        if pcase != 'Welcome to The PKaser':
            parentId = self.json_data[pcase]['sf_link'].split('/')[-1]
            if 'view' in parentId:
                parentId = self.json_data[pcase]['sf_link'].split('/')[6]
            return parentId       
                
    def saveJSON(self,file,data):
        with open(file,'w') as outfile:
            json.dump(data,outfile)

    def time_tracker_tree(self):

        def select():
            curItems = self.time_tracker_t.selection()
            selected = "\n".join(["%s_%s"%(str(self.time_tracker_t.item(i)['values'][0]),str(self.time_tracker_t.item(i)['values'][1])) for i in curItems])            
            selectedTimes = []
            selectedTimes.append([str(self.time_tracker_t.item(i)['values'][2]) for i in curItems])
            total_time = str(timeTracker.calculate_selected_times(selectedTimes[0]))
            output = "%s\nTotalTime:%s"%(selected,total_time)
            print(output)
            self.copyInput(output)


        pcase = self.getPCaseString()
        if pcase:
            if self.first_init:
                frame_2 = self.tab_2

            else:
                frame_2 = self.tab_2
                self.tree_scroll.destroy()
                self.time_tracker_t.destroy()
        
        self.time_tracker_t = ttk.Treeview(frame_2, column=("c1", "c2","c3"), show= 'headings', height= 24)

        #self.time_tracker_t.delete( * self.time_tracker_t.get_children())
        self.time_tracker_t.column("#1",anchor=CENTER, width=95)
        self.time_tracker_t.heading("#1", text= "PCASE")
        self.time_tracker_t.column("#2", anchor= CENTER, width=160)
        self.time_tracker_t.heading("#2", text= "CSRNAME")
        self.time_tracker_t.column("#3", anchor= CENTER, width=95)
        self.time_tracker_t.heading("#3", text= "TOTAL TIME")


        self.tree_scroll = ttk.Scrollbar(frame_2,command=self.time_tracker_t.yview)
        self.time_tracker_t['yscrollcommand'] = self.tree_scroll.set
        self.tree_scroll.pack(side=tk.RIGHT,fill = tk.Y)
        self.time_tracker_t.pack(fill=tk.BOTH)

        self.times = timeTracker.load_total_times()
        print("time_tracker_tree - times: %s"%self.times)
        times = self.times
        today_date = timeTracker.get_today_date()

 
        cases = times[today_date]
        print("time_tracker_tree - cases: %s"%(cases))
        rowCount = 1
        for case in cases:
            if case:
                captureRowCount = rowCount
                pcase_csrname = case.split("_")
                captureRowCount = rowCount
                self.time_tracker_t.insert('', index='end',iid=rowCount,values=(pcase_csrname[0],pcase_csrname[1], cases[case]))
                rowCount += 1
        
        self.time_tracker_t.bind("<Return>", lambda e: select())
        

    def code_conflict_tree(self):
        def selectItem(a):
            curItem = self.conflictTree.focus()
            returnSelected = self.conflictTree.item(curItem).get("text")
            print(returnSelected)
            return str(returnSelected)

        self.conflicts = find_OpenCases(self.getCustString(), self.getPCaseString())
        #print(self.conflicts)
        conflicts = target=self.conflicts
        

        # Define Tree
        self.conflictTree = ttk.Treeview(self.code_conflict_frame)
        self.veritcalScrollbar = ttk.Scrollbar(self.code_conflict_label, orient='vertical', command=self.conflictTree.yview)
        self.conflictTree.configure(yscroll=self.veritcalScrollbar.set)
        self.veritcalScrollbar.grid(row=0,column=3,sticky='ns')
        #Define Columns
        self.conflictTree['columns'] = ()
        #format Columns
        self.conflictTree.column("#0",width=415, minwidth=25)

        #Create Heading
        self.conflictTree.heading("#0", text="Potential Conflicts", anchor="w")
        self.conflictTree.tag_configure('yes-conflict',background='#ffa6bb')
        self.conflictTree.tag_configure('no-conflict',background='#a6ffb6')

        # idk yet
        self.conflictTree.bind('<Double-Button-1>', selectItem)

        # Add Data
        rowCount = 1
        b_noconflict = False
        for iConflict in conflicts:
            for pcaseNumber in iConflict:
                if "NOCONFLICT" in pcaseNumber:
                    pass
                else: # insert conflicting case
                    b_noconflict = True
                    captureRowCount = rowCount
                    self.conflictTree.insert(parent='', index='end', iid=rowCount, text=pcaseNumber,tags=('yes-conflict',))
                    rowCount += 1
                    for i in range(len(iConflict[pcaseNumber])):
                        value = iConflict[pcaseNumber][i]
                        if i == 0:
                            text = "SFNumber ~ %s"%(value)
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1
                        elif i == 1:
                            value = value.strip()
                            if value != "":
                                text = "sfCaseSubject ~ %s"%(value[:48])
                                self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                                self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                                rowCount += 1
                            else:
                                text = "sfCaseSubject ~ None Provided"
                                self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                                self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                                rowCount += 1
                        elif i == 2:
                            text = "Current Owner ~ %s"%(value)
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1
                        elif i == 3:
                            text = "Roll Status ~ %s"%(value)
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1
                        elif i == 4:
                            text = "Case Status ~ %s"%(value)
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1

                        elif i == 5:
                            text = "PCase Path ~ %s"%(value)
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1

                        elif i == 6:
                            text = "Template(s) ~ %s"%(" ".join(value))
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1


                        elif i == 7:
                            text = "Image(s) ~ %s"%(" ".join(value))
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1

                        elif i == 8:
                            text = "Terms Backer(s) ~ %s"%(" ".join(value))
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1

                        elif i == 9:
                            text = "Script(s) ~ %s"%(" ".join(value))
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1
                        
                        elif i == 10:
                            text = "Parser(s) ~ %s"%(" ".join(value))
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=text)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1
                        else:
                            self.conflictTree.insert(parent='', index='end', iid=rowCount, text=value)
                            self.conflictTree.move(rowCount, captureRowCount, captureRowCount)
                            rowCount += 1
        if not b_noconflict:
                    captureRowCount = rowCount
                    self.conflictTree.insert(parent='', index='end', iid=rowCount, text="CURRENTLY NO CONCLICTS",tags=('no-conflict',))
                    rowCount += 1
        self.conflictTree.grid(row=1, column=0)



    def rightClickMenuNotePad(self, event):
        try:
            self.right_click_menu_notepad.tk_popup(event.x_root, event.y_root)
        finally:
            self.right_click_menu_notepad.grab_release()

    def rightClickMenuPCaseFolder(self, event):
        try:
            self.right_click_menu_pcase_folder.tk_popup(event.x_root, event.y_root)
        finally:
            self.right_click_menu_pcase_folder.grab_release()
   
    
    def copyPCase(self):
        pyperclip.copy(self.getPCaseString())
    def copySFCase(self):
        pyperclip.copy(self.getCaseString())
    def copyPCaseAndCSRName(self):
        stamp = "%s_%s"%(self.getPCaseString(), self.getCustString())
        pyperclip.copy(stamp)
    def copyCSRName(self):
        pyperclip.copy(self.getCustString())
    
    def copyInput(self, input):
        pyperclip.copy(input)
        

    def pythonComment(self):
        pyperclip.copy('''
        # PCASE:%s - SFCASE: %s-  cdurham - %s
        # Request:
        # Action:
        '''%(self.getPCaseString(), self.getCaseString(),str(datetime.datetime.now().date()))
        )


            
    #Some Helpful mutator and accessor methods:
    def setServerList(self):
        self.server_list = self.all_servers
    def setDataFolder(self,folder):
        self.data_folder = folder
    def getDataFolder(self):
        return self.data_folder
    def setDataFile(self,file):
        self.data_file = file
    def getDataFile(self):
        return self.data_file

    def setAppIconFile(self, file):
        self.appIcon_file = file  
    def getAppIconFile(self):
        return self.appIcon_file

    def setBrandingImage(self, file):
        self.branding_image_file = file
    def getBrandingImage(self):
        return self.branding_image_file
    def setConfigFile(self, file):
        self.config_file = file  
    def getConfigFile(self):
        return self.config_file

    def setWordsFile(self, file):
        self.words_file = file  
    def getWordsFile(self):
        return self.words_file

    def setPKaseNotesFolder(self, file):
        self.pcase_notes = file  
    def getPKaseNotesFolder(self):
        return self.pcase_notes

    def setLastSelectedFile(self, file):
        self.last_selected = file
    def getLastSelectedFile(self):
        return self.last_selected
    
    def setTotalTimeSpent(self, file):
        self.total_time_spent = file
    def getTotalTimeSpent(self):
        return self.total_time_spent


    def setPCaseString(self,pcase):
        self.pcase=pcase
    def getPCaseString(self):
        return self.pcase
    def setCustString(self,cust):
        self.cust = cust
    def getCustString(self):
        return self.cust

    def setCaseString(self,sfnumber):
        self.case_number = sfnumber
    def getCaseString(self):
        return self.case_number



    
    # Wrapper methods for each push/edit button
    def editTemplates(self):
        self.editFiles("\\FDT",".csv")
    def editParsers(self):
        self.editFiles("\\VC\\ParserConfigs",".xml")
    def editScripts(self):
        self.editFiles("\\VC\\Scripts",".py")
    def editSamples(self):
        try:
            self.editFiles("\\Sample Data")
        except FileNotFoundError:
            self.editFiles("\\SAMPLE_DATA")
    def editRelease(self):
        self.editFiles("\\VC\\Release\\Bin",)
        
    def pushTemplates(self):
        self.pushFiles("\\FDT",".csv")
    def pushParsers(self):
        self.pushFiles("\\VC\\ParserConfigs",".xml")
    def pushScripts(self):
        try:
            self.pushFiles("\\VC\\Scripts",".py")
        except FileNotFoundError:
            self.pushFiles("\\VC\\SCRIPTS",".py")
    def pushSamples(self):
        try:
            self.pushFiles("\\SAMPLE_DATA")
        except FileNotFoundError:
            self.pushFiles("\\Sample Data")   
    def pushRelease(self):
        self.pushFiles("\\VC\\Release\\Bin",)
    
    def diffScripts(self):
        try:
            self.diffFiles("\\VC\\Scripts",".py")
        except FileNotFoundError:
            self.diffFiles("\\VC\\SCRIPTS",".py")
    def diffTemplates(self):
        self.diffFiles("\\FDT",".csv")

    def diffParser(self):
        self.diffFiles("\\VC\\ParserConfigs",".xml")
    def diffReleaseBin(self):
        self.diffFiles("\\VC\\Release\\Bin",)



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
            os.startfile(path)
        except:
            tk.messagebox.showwarning('Error', 'You must be connected to the VPN to open this.')

    def openWebsite(self,page):
        if page:
            webbrowser.get('windows-default').open(page)
        else:
            tk.messagebox.showwarning('Error', 'There is no SRD saved for this case.\nYou can add via Edit > Change PCase Details') 
        
    def openViaPython(self,file):
        try:
            call(["python", file])
        except:
            tk.messagebox.showinfo('You need python installed and in your path file to open this.')


class PKaseNotesFuctions(PCaser):
    def __cut__(self):
        # Using Tkinter event binding for cut
        self.text_area.event_generate("<<Cut>>")
    
    def __copy2__(self):
        self.text_area.event_generate("<<Copy>>")

    def __paste__(self):
        print("Paste Function Ran.")
        self.text_area.event_generate("<<Paste>>")

    def __undo__(self):
        self.text_area.edit_undo()
    def __redo__(self):
        self.text_area.edit_redo()

    def __openFile__(self):
        self.text_area.configure(state="normal")
        self.__file__ = filedialog.askopenfilename(defaultextension=".txt", filetypes=[("Text Documents","*.txt")],initialdir=str(self.getPKaseNotesFolder()))
        if self.__file__ == "":
            self.__file__ = None
        else:
        # Try to open the file
        # set the window title
            self.location = str(self.__file__)
            #self.mainwindow.title(os.path.basename(self.location) + " - The PKaser")
            self.text_area.delete(2.0, END)
            try:
                with open(self.location, "r") as file:
                    self.text_area.insert(2.0,file.read())
                    file.close()
            except FileNotFoundError:
                self.__file__ = filedialog.askopenfilename(defaultextension=".txt", filetypes=[("Text Documents","*.txt")],initialdir=str(self.getPKaseNotesFolder()))
                if self.__file__ == "":
                    self.__file__ = None
                else:
                # Try to open the file
                # set the window title
                    self.location = str(self.__file__)
                    #self.main_frame.title("Untitled-Notepad")
                    self.text_area.delete(1.0, END)

    def __salesforceComment__(self):
        self.user = os.getlogin()
        self.today = datetime.datetime.now().date()
        self.now = datetime.datetime.now().time()
        self.time = self.now.strftime("%I:%M:%S%p")
        self.timestamp = "%s_%s_%s"%(self.user,self.today,self.time)
        self.salesforceComment= "%s\n\nQA Items:\nPending QA:\nBA Items:\nPM Items:\nData Needed:\n\nResolution:\n\nPre:\nPost:\n\nUpdates Made To:\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.timestamp)
        self.text_area.mark_set("insert",1.0)
        self.text_area.insert(INSERT, self.salesforceComment)
        self.text_area.configure(state="normal")
    

    def __FileFailureComment__(self):
        self.user = os.getlogin()
        self.today = datetime.datetime.now().date()
        self.now = datetime.datetime.now().time()
        self.time = self.now.strftime("%I:%M:%S%p")
        self.timestamp = "%s_%s_%s"%(self.user,self.today,self.time)
        self.salesforceComment= "%s\nFile Failure Case\nData File:\nLine Number:\nColumn:\nDescription:\n\nLinks to any screenshots saved in the pcase:\nDiagnosis:\n\nReplicated in IMS Y/N:\nSuccessful Batch:\nResolution:\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.timestamp)
        self.text_area.mark_set("insert",1.0)
        self.text_area.insert(INSERT, self.salesforceComment)
        self.text_area.configure(state="normal")

    def __regularNoteStamp__(self):
        self.user = os.getlogin()
        self.today = datetime.datetime.now().date()
        self.now = datetime.datetime.now().time()
        self.time = self.now.strftime("%I:%M:%S%p")
        self.timestamp = "%s_%s_%s\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.user,self.today,self.time)
        self.text_area.mark_set("insert",1.0)
        self.text_area.insert(INSERT, self.timestamp)
        self.text_area.configure(state="normal")

    def __checkSlackStamp__(self):
        self.user = os.getlogin()
        self.today = datetime.datetime.now().date()
        self.now = datetime.datetime.now().time()
        self.time = self.now.strftime("%I:%M:%S%p")
        self.timestamp = "%s_%s_%s\nSlack Convo:\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.user,self.today,self.time)
        self.text_area.mark_set("insert",1.0)
        self.text_area.insert(INSERT, self.timestamp)
        self.text_area.configure(state="normal")

    def __CaseCommittedNoteStamp__(self):
        self.user = os.getlogin()
        self.today = datetime.datetime.now().date()
        self.now = datetime.datetime.now().time()
        self.time = self.now.strftime("%I:%M:%S%p")
        self.timestamp = "%s_%s_%s\nCase Committed. Ready To Roll.\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.user,self.today,self.time)
        self.text_area.mark_set("insert",1.0)
        self.text_area.insert(INSERT, self.timestamp)
        self.text_area.configure(state="normal")
        sfCase.statusToCheckedIn(self.returnParentId())

    def __sentToEIPP__(self):
        self.user = os.getlogin()
        self.today = datetime.datetime.now().date()
        self.now = datetime.datetime.now().time()
        self.time = self.now.strftime("%I:%M:%S%p")
        self.timestamp = "%s_%s_%s\nSent Case To EIPP.\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.user,self.today,self.time)
        self.text_area.mark_set("insert",1.0)
        self.text_area.insert(INSERT, self.timestamp)
        self.text_area.configure(state="normal")
        sfCase.statusToEIPP(self.returnParentId())

    def __sentToFCI__(self):
        self.user = os.getlogin()
        self.today = datetime.datetime.now().date()
        self.now = datetime.datetime.now().time()
        self.time = self.now.strftime("%I:%M:%S%p")
        self.timestamp = "%s_%s_%s\nSent Case To FCI.\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.user,self.today,self.time)
        self.text_area.mark_set("insert",1.0)
        self.text_area.insert(INSERT, self.timestamp)
        self.text_area.configure(state="normal")
        sfCase.statusToFCI(self.returnParentId())

    def __callReviewCompleted__(self):
        self.user = os.getlogin()
        self.today = datetime.datetime.now().date()
        self.now = datetime.datetime.now().time()
        self.time = self.now.strftime("%I:%M:%S%p")
        self.timestamp = "%s_%s_%s\nCase has been reviewed on call with offshore team. Case Ready for ASE/QA Process.\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.user,self.today,self.time)
        self.text_area.mark_set("insert",1.0)
        self.text_area.insert(INSERT, self.timestamp)
        self.text_area.configure(state="normal")
        sfCase.commentCallReviewCompleted(self.returnParentId())

    def __PendingASECodeReview__(self):
        self.user = os.getlogin()
        self.today = datetime.datetime.now().date()
        self.now = datetime.datetime.now().time()
        self.time = self.now.strftime("%I:%M:%S%p")
        self.timestamp = "%s_%s_%s\nBT Engineer Peer Review Completed.Case Pending ASE Code Review.\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.user,self.today,self.time)
        self.text_area.mark_set("insert",1.0)
        self.text_area.insert(INSERT, self.timestamp)
        self.text_area.configure(state="normal")
        print(self.returnParentId())
        sfCase.statusToASECodeReview(self.returnParentId())
        
    def __newFile__(self):
        #self.root.title("Untitled - Notepad")
        self.fileCheck = self.text_area.get("1.0",END)
        if self.fileCheck != "":
            self.__saveFileAs__()

        self.__file__ = None
        self.user = os.getlogin()
        self.today = datetime.datetime.now().date()
        self.timestamp = "%s_%s\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.user,self.today)
        self.text_area.delete(1.0, END)
        self.text_area.insert(INSERT, self.timestamp)
        self.text_area.configure(state="normal")
    
    def __fileCheck__(self):
        fileCheck = self.text_area.get("1.0",END)
        return fileCheck

    def __saveFileAs__(self):
        self.text_area.configure(state="normal")
        intitialfile = "%s_%s"%(self.getPCaseString(), self.getCustString())
        self.__file__ = filedialog.asksaveasfilename(initialfile=intitialfile, defaultextension=".txt",filetypes=[("Text Documents","*.txt")],initialdir=str(self.getPKaseNotesFolder()))
        if self.__file__ == None:
            print("If Statement Ran")
            if self.__file__ == "":
                self.__file__ = None
            else:
                # try to save file
                while True:
                    try:
                        with open(self.__file__,"w", encoding="utf-8") as file:
                            file.write(self.text_area.get(1.0, END))
                            file.close()
                    except TypeError:
                        continue
        else:
            self.location = str(self.__file__)                
            with open(self.location, "w", encoding="utf-8") as file:
                file.write(self.text_area.get(1.0,END))
    
    def __saveOnSelect__(self,prevFileName,prevTextArea):
        print("__saveOnSelect__ Ran")
        prevFileName = prevFileName
        prevTextArea = prevTextArea
        self.text_area.configure(state="normal")
        pcaseFolder = self.getPKaseNotesFolder()
        #intitialfile = "%s_%s.txt"%(self.getPCaseString(), self.getCustString())
        self.location =  pcaseFolder+"\\" + prevFileName      
        with open(self.location, "w", encoding="utf-8") as file:
            file.write(prevTextArea)

    def __openOnSelect__(self):        
        self.text_area.delete(1.0, END)
        self.text_area.configure(state="normal")
        pcaseFolder = self.getPKaseNotesFolder()
        intitialfile = "%s_%s.txt"%(self.getPCaseString(), self.getCustString())
        self.location =  pcaseFolder+"\\" +intitialfile
        try:
            with open(self.location, "r", encoding='utf-8') as file:
                self.text_area.insert(1.0,file.read())
                file.close()
        except FileNotFoundError:
            print('File Not found we need to create it.')
            self.__file__ = None
            self.user = os.getlogin()
            self.today = datetime.datetime.now().date()
            self.timestamp = "%s_%s\nCase picked up.\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n"%(self.user,self.today)
            self.text_area.delete(1.0, END)
            self.text_area.insert(INSERT, self.timestamp)
            self.text_area.configure(state="normal")

    def __caseStartTime__(self):
        with open(self.getTotalTimeSpent(), "w") as file:
            csrname = self.getCustString()
            pcase = self.getPCaseString()
            start = time.time()
            file.write("%s,%s,%s,"%(pcase,csrname,str(start)))
            file.close()
    def __caseEndTime__(self):
        with open(self.getTotalTimeSpent(), "a") as file:
            end = time.time()
            file.write("%s"%str(end))
            file.close()
    
    def __totalTimeSpent__(self):
        with open(self.getTotalTimeSpent()) as f:
            times = f.readline().split(",")
        f.close()

        startTime = float(times[2])
        endTime = times[3].strip("\n")
        endTime = float(endTime)
        total_time_spent = startTime - endTime

        dt1 = datetime.datetime.fromtimestamp(startTime) # 1973-11-29 22:33:09
        dt2 = datetime.datetime.fromtimestamp(endTime) # 1977-06-07 23:44:50
        rd = dateutil.relativedelta.relativedelta(dt2, dt1)

        total_time_message = "You spent %s Hours %s minutes on %s."%(rd.hours,rd.minutes,self.getCustString())

        tk.messagebox.showinfo("Total Time",total_time_message)
    

    def __returnSelectedText__(self):
        try:
            selectedText = self.text_area.selection_get()
            return selectedText
        except tkinter.TclError:
            error_message = "No text Selected! Make sure you highlight the text you want to add as a comment."
            tk.messagebox.showerror(message='Error: "{}"'.format(error_message))
            return None



        

        





#   SPELL CHECK NOT IN USE CURRENTLY
    def __spellCheck__(self, event):
        '''Spellcheck the word preceeding the insertion point'''

        index = self.text_area.search(r'\s', "insert", backwards=True, regexp=True)
        if index == "":
            index ="1.0"
        else:
            index = self.text_area.index("%s+1c" % index)
        grab_word = self.text_area.get(index, "insert")
        regex = r"(\\n)|(\\r)|(\\t)|(\u00A9)|([!\"#$%&()*+,-./:;<=>?@\[\\\]^_`{|}~])"
        subst = ""
        grab_word_result = re.sub(regex, subst, grab_word, 0, re.IGNORECASE).lower()
        print(grab_word_result)
        if grab_word_result not in self.words_file:       
            last_index = "%s+%dc" % (index, len(grab_word))        
            self.text_area.delete(index, last_index)
            self.text_area.insert(index,grab_word)
            self.text_area.tag_add("misspelled",index, last_index)
            self.text_area.tag_config("misspelled",underline=True)
    

    
    

    


def main():
    PKaseNotes = PKaseNotesFuctions()
    PKaseNotes.run()

    
if __name__ == '__main__':
    main()
    

