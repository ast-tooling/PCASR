from requests.sessions import cookiejar_from_dict
from simple_salesforce import Salesforce, SalesforceLogin
import configparser
import keyring
import os
import requests
from tkinter import * 
from tkinter import messagebox

def addCaseComment(sf_link="",comment=""):
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
    if comment != "":
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
    sf.Case.update(parentId,{'Roll_Status__c': "Checked In",'Lifecycle_Status__c': "QA Complete"})
    case_comment = {'ParentId': parentId, 'IsPublished': False, 'CommentBody': "Case Committed. Ready To Roll!"}
    sf.CaseComment.create(case_comment)

def updateEngineer(sf_link="",bt_eng=""):
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
    sf.Case.update(parentId,{'Engineer__c': bt_eng})

#updateEngineer('https://billtrust.my.salesforce.com/5001M00001isgRV','Christopher Durham')

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
    sf.FeedItem.create(params)
