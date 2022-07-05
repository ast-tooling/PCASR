import os
import json

def jobTitles():
    aJobTitles = [ "AST ENG","AST QA"]
    return aJobTitles

def trackTimeOpts():
    aTrackTimeOpts = ["YES", "NO"]
    return aTrackTimeOpts

def checkOptFileExists():
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser\\nt-json-files" %user
    data_file = data_folder+"\\pk_options.json"
    b_check_value = os.path.exists(data_file)
    return b_check_value

def getTrackTimeValue(selection):
    message = "User changed Time Tracker Pkaser Option to: %s"%(selection)
    print(message)
    opts = loadOptFile()
    print("getTrackTimeValue: %s"%(opts))
    opts["TRACK_TIME"] = selection
    saveOptFile(opts)
    opts = loadOptFile()
    print("getTrackTimeValue: %s"%(opts))
    return selection


def getjobTitle(selection):
    message = "User changed Job Title Pkaser Option to: %s"%(selection)
    print(message)
    opts = loadOptFile()
    print("getjobTitle: %s"%(opts))
    opts["JOB_TITLE"] = selection
    saveOptFile(opts)
    opts = loadOptFile()
    print("getjobTitle: %s"%(opts))
    return selection

def createOptFile():
    user = os.getlogin()
    opts = {"JOB_TITLE":"AST ENG",
            "TRACK_TIME":"YES"}

    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser\\nt-json-files" %user
    data_file = data_folder+"\\pk_options.json"
    with open(data_file, 'w') as optFile:
        json.dump(opts, optFile)
    optFile.close()

def loadOptFile():
    user = os.getlogin()
    opts = {}
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser\\nt-json-files" %user
    data_file = data_folder+"\\pk_options.json"
    with open(data_file, 'r') as optFile:
        opts = json.load(optFile)
    print(opts)
    return opts

def saveOptFile(opts={}):
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser\\nt-json-files" %user
    data_file = data_folder+"\\pk_options.json"
    with open(data_file, 'w') as optFile:
        json.dump(opts, optFile)
    optFile.close()




