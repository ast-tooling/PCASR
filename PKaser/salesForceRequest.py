from simple_salesforce import Salesforce, SalesforceLogin
import json
import os
import configparser
import keyring

def get_SalesForceReport():
    global report_results_temp
    user = os.getlogin()
    
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    report_results_temp = data_folder+"\\nt-json-files\\temp_report_results.json"
    data_file = data_folder+"\\config.txt"


    config = configparser.ConfigParser()
    config.read(data_file)
    username = config.get('credentials','username')
    sf = Salesforce(
        username= username,
        password=keyring.get_password("pkaser-userinfo", username),
        security_token=keyring.get_password("pkaser-token", username)
    )


    report_id = '00O1M000007vhEO'

    report_results = sf.restful('analytics/reports/{}'.format(report_id))

    with open(report_results_temp, 'w') as fp:
        json.dump(report_results, fp)


def load_OpenCases():
    # Opening JSON file 
    f = open(report_results_temp) 

    data = json.load(f)
    # Closing file 
    f.close()

    caseInfo = data['factMap']["T!T"]['rows']
    # Closing file 
    f.close()
    caseCount = -1
    for case in caseInfo: 
        caseCount +=1

    openCaseList = []
    while caseCount >= 0:

        strCount = str(caseCount)
        csrUsername = data['factMap']["T!T"]['rows'][caseCount]['dataCells'][15]['label']
        #print(csrUsername)
        sfCaseNumber = data['factMap']["T!T"]['rows'][caseCount]['dataCells'][0]['label']
        #print(sfCaseNumber)
        pcaseNumber = data['factMap']["T!T"]['rows'][caseCount]['dataCells'][2]['label']
        #print(pcaseNumber)
        rollStatus = data['factMap']["T!T"]['rows'][caseCount]['dataCells'][3]['label']
        #print(rollStatus)
        caseType = data['factMap']["T!T"]['rows'][caseCount]['dataCells'][8]['label']
        #print(caseType)
        caseStatus = data['factMap']["T!T"]['rows'][caseCount]['dataCells'][5]['label']
        #print(caseStatus)
        currentQueue = data['factMap']["T!T"]['rows'][caseCount]['dataCells'][7]['label']
        #print(currentQueue)
        pcasePath = "Z:\IT Documents\QA\%s"%(pcaseNumber)
        #print(pcasePath)
        caseSubject = data['factMap']["T!T"]['rows'][caseCount]['dataCells'][9]['label']

        caseDict = {'CSRUSERNAME': csrUsername, 'SFCASENUMBER':sfCaseNumber, 'PCASENUMBER':pcaseNumber,
                    'CASETYPE':caseType, 'CASESTATUS':caseStatus, 'CURRENTQ':currentQueue, 
                    'ROLLSTATUS':rollStatus, 'PCASEPATH':pcasePath, 'CASESUBJECT':caseSubject}

        openCaseList.append(caseDict)
        #print(caseDict)
        #print("---------------------------")   
        caseCount -=1
    
    return openCaseList



