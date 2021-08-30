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
    start = sf_link.rfind("/") + 1
    parentId = sf_link[start:]
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
    start = sf_link.rfind("/") + 1
    parentId = sf_link[start:]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    sf.Case.update(parentId,{'Lifecycle_Status__c': "With EIPP"})


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
    start = sf_link.rfind("/") + 1
    parentId = sf_link[start:]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    sf.Case.update(parentId,{'Lifecycle_Status__c': "With FCI"})



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
    start = sf_link.rfind("/") + 1
    parentId = sf_link[start:]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    sf.Case.update(parentId,{'Lifecycle_Status__c': "Pending QA"})


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
    start = sf_link.rfind("/") + 1
    parentId = sf_link[start:]

    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username),
        organizationId=organizationId,
        session=session,

    )
    params = {"ParentId":parentId,"Body": "Christopher Durham Case Commmitted. Ready To Roll.","Type":"LinkPost","Visibility":"InternalUsers","LinkUrl":"Christopher Durham"}
    sf.FeedItem.create(params)

#codeRollChatterComment("https://billtrust.my.salesforce.com/5001M00001isgRV")