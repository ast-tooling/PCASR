import os

# tkinter imports
import tkinter as tk
from tkinter import filedialog, TclError
from tkinter.font import Font, families
from tkinter.messagebox import *
from tkinter import font


# local Imports
import file_picker_window
from winMergeU import *
from pkaser import *
import copy_files_window

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
    pcase = getPCaseString()
    if pcase:
        subFolder_path = "Z:\\IT Documents\\QA\\" + pcase + subDir
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
    pcase = self.
    if pcase:
        subFolder_path = "Z:\\IT Documents\\QA\\" + pcase + subDir
        if fileType == ".csv":
            astFolder_path = "C:\\AST\\FDT\\"
        elif fileType == ".xml":
            astFolder_path = "C:\\AST\\ParserConfigs\\"
        elif fileType == ".py":
            astFolder_path = "C:\\AST\\scripts\\"
        elif fileType == "":
            astFolder_path = "\\VC\\Release\\Bin"

        files = self.filesInDir(subFolder_path,fileType)

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
        else:
            selected_files = file_picker_window.filePickerWindow(self,subFolder_path).show()
            if selected_files:
                copy_files_window.copyFilesWindow(self,subFolder_path,subDir,selected_files)

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
def diffTemplates():
    diffFiles("\\FDT",".csv")

def diffParser(self):
    self.diffFiles("\\VC\\ParserConfigs",".xml")
def diffReleaseBin(self):
    self.diffFiles("\\VC\\Release\\Bin",)