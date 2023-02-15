from requests.sessions import cookiejar_from_dict
from simple_salesforce import Salesforce, SalesforceLogin
import configparser
import keyring
import os
import ctypes
import requests
from tkinter import * 
from tkinter import messagebox
from tkinter.messagebox import askyesno

# local imports
from pk_salesforce.salesForceRequest import getSfUserId, getSfCaseInfo
import pk_features.pk_options as pk_options

def addCaseComment(sf_link="",comment=""):
    print(type(comment))
    print(comment)
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    data_file = data_folder+"\\config.txt"
    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    organizationId = config.get('credentials','organizationId')
    session = requests.Session()
    
    sf_link = sf_link 
    parentId = sf_link.split('/')[-1]
    if 'lightning' in parentId:
        parentId = sf_link.split('/')[6]
    comment = comment

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    
    case_comment = {'ParentId': parentId, 'IsPublished': False, 'CommentBody': comment}

    if comment is not None:
        sf.CaseComment.create(case_comment)
        messagebox.showinfo("Information", "Comment Added Successfully :)")


def statusToEIPP(sf_link=""):
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    data_file = data_folder+"\\config.txt"
    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    organizationId = config.get('credentials','organizationId')
    session = requests.Session()
    
    sf_link = sf_link 
    parentId = sf_link.split('/')[-1]
    if 'lightning' in parentId:
        parentId = sf_link.split('/')[6]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    params = {'Lifecycle_Status__c': 'Pending Dev'}
    sf.Case.update(parentId,params)
    case_comment = {'ParentId': parentId, 'IsPublished': False, 'CommentBody': "Sent Case To EIPP."}
    sf.CaseComment.create(case_comment)

def commentCallReviewCompleted(sf_link=""):
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    data_file = data_folder+"\\config.txt"
    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    organizationId = config.get('credentials','organizationId')
    session = requests.Session()
    
    sf_link = sf_link 
    parentId = sf_link.split('/')[-1]
    if 'lightning' in parentId:
        parentId = sf_link.split('/')[6]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    params = {'Lifecycle_Status__c': 'Peer Review'}
    sf.Case.update(parentId,params)
    case_comment = {'ParentId': parentId, 'IsPublished': False, 'CommentBody': "Case has been reviewed on call with offshore team. Case Ready for ASE/QA Process."}
    sf.CaseComment.create(case_comment)

def statusToFCI(sf_link=""):
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    data_file = data_folder+"\\config.txt"
    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    organizationId = config.get('credentials','organizationId')
    session = requests.Session()
    
    sf_link = sf_link 
    parentId = sf_link.split('/')[-1]
    if 'lightning' in sf_link:
        parentId = sf_link.split('/')[6]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    sf.Case.update(parentId,{'Lifecycle_Status__c': "With FCI"})
    case_comment = {'ParentId': parentId, 'IsPublished': False, 'CommentBody': "Sent Case To FCI."}
    sf.CaseComment.create(case_comment)


def statusToASECodeReview(sf_link=""):
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    data_file = data_folder+"\\config.txt"
    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    organizationId = config.get('credentials','organizationId')
    session = requests.Session()
    
    sf_link = sf_link 
    parentId = sf_link.split('/')[-1]
    if 'lightning' in sf_link:
        parentId = sf_link.split('/')[6]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    sf.Case.update(parentId,{'Lifecycle_Status__c': "Senior Code Review"})
    case_comment = {'ParentId': parentId, 'IsPublished': False, 'CommentBody': "BT Engineer Peer Review Completed.Case Pending ASE Code Review."}
    sf.CaseComment.create(case_comment)

def statusToPendQA(sf_link=""):
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    data_file = data_folder+"\\config.txt"
    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    organizationId = config.get('credentials','organizationId')
    session = requests.Session()
    
    sf_link = sf_link 
    parentId = sf_link.split('/')[-1]
    if 'lightning' in sf_link:
        parentId = sf_link.split('/')[6]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    sf.Case.update(parentId,{'Lifecycle_Status__c': "Pending QA"})




def statusToCheckedIn(sf_link=""):
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    data_file = data_folder+"\\config.txt"
    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    organizationId = config.get('credentials','organizationId')
    session = requests.Session()
    
    sf_link = sf_link 
    parentId = sf_link.split('/')[-1]
    if 'lightning' in sf_link:
        parentId = sf_link.split('/')[6]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    sf.Case.update(parentId,{'Roll_Status__c': "Checked In",'Lifecycle_Status__c': "Customer Approved (Pending Roll)"})
    case_comment = {'ParentId': parentId, 'IsPublished': False, 'CommentBody': "Case Committed. Ready To Roll!"}
    sf.CaseComment.create(case_comment)

def updateEngineer(sf_link="",bt_eng="",job_title= "AST ENG"):
    bt_eng = bt_eng
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    data_file = data_folder+"\\config.txt"
    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    organizationId = config.get('credentials','organizationId')
    session = requests.Session()
    
    sf_link = sf_link 
    parentId = sf_link.split('/')[-1]
    if 'lightning' in sf_link:
        parentId = sf_link.split('/')[6]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,
    )
    if "AST ENG" == job_title:
        sf.Case.update(parentId,{'Engineer__c': bt_eng})
    elif "AST QA" == job_title:
        sf.Case.update(parentId,{'QA_Analyst__c': bt_eng})


def updateCaseOwner(sf_link="",new_owner=""):
    new_owner = new_owner
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    data_file = data_folder+"\\config.txt"
    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    organizationId = config.get('credentials','organizationId')
    session = requests.Session()
    
    sf_link = sf_link 
    parentId = sf_link.split('/')[-1]
    if 'lightning' in sf_link:
        parentId = sf_link.split('/')[6]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,
    )
    sf.Case.update(parentId,{'ChangeCaseOwner__Case_Owner_Change': new_owner})


def codeRollChatterComment(sf_link=""):
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    data_file = data_folder+"\\config.txt"
    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    organizationId = config.get('credentials','organizationId')
    session = requests.Session()
    
    sf_link = sf_link 
    parentId = sf_link.split('/')[-1]
    if 'lightning' in sf_link:
        parentId = sf_link.split('/')[6]

    print(parentId)
    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    params = {"ParentId": parentId,"Body": "Christopher Durham Case Commmitted. Ready To Roll."}
    #sf.FeedItem.create(params)
    sf.FeedItem.search(params)

def takeOwnership(sf_link=""):
    # Set User Type
    opts = pk_options.loadOptFile()
    job_title = opts["JOB_TITLE"]

    sf_link = sf_link
    answer = askyesno(title="Please Confirm!",
                    message="Will you be taking ownership of this case?")
    if answer:
        def get_display_name(): # https://stackoverflow.com/questions/21766954/how-to-get-windows-users-full-name-in-python
            GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
            NameDisplay = 3
            size = ctypes.pointer(ctypes.c_ulong(0))
            GetUserNameEx(NameDisplay, None, size)
            nameBuffer = ctypes.create_unicode_buffer(size.contents.value)
            GetUserNameEx(NameDisplay, nameBuffer, size)
            return nameBuffer.value
        fullName = get_display_name()
        print(fullName)
        updateEngineer(sf_link,getSfUserId(),job_title)
    else:
        updateEngineer(sf_link,"",job_title)
#codeRollChatterComment("https://billtrust.my.salesforce.com/5001M00001isgRV")
