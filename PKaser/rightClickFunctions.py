from requests.sessions import cookiejar_from_dict
from simple_salesforce import Salesforce, SalesforceLogin
import configparser
import keyring
import os
import requests
from tkinter import * 
from tkinter import messagebox

def addComment(sf_link="",comment=""):
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

#addComment("https://billtrust.my.salesforce.com/5001M00001isgRV","Test Case Comment")