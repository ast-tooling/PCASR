from salesForceRequest import *

def find_OpenCases(csr_name, current_pcase):

    get_SalesForceReport()
    openCaseList = load_OpenCases()

    current_case = csr_name
    current_pcase = current_pcase
    matches = 0

    finalTreeViewList = []

    for case in openCaseList:

        imagesList = []
        templateList = []
        fdtMiscList = []
        termsBackerList = []
        pythonScriptList = []
        scriptMiscList = []  
        parserFileList = []
        parserMiscList = []
        releaseBinList = []
        releaseBinMiscList = []

        matchOnCsr = case['CSRUSERNAME']
        matchOnPCase = case['PCASENUMBER']
        rollStatus = case['ROLLSTATUS']
        if current_case == matchOnCsr and current_pcase != matchOnPCase and rollStatus != "Rolled"and current_pcase != "P2929255": # Need pcase Number for if Condition to compare.
            matches += 1
            #print("-----------Start Case %s Details----------------"%(matches))
            
            sfCaseNumber = case['SFCASENUMBER']
            #print('Case Number: %s'%(sfCaseNumber))

            caseReason = case["CASESUBJECT"]
        

            whoHasCase = case['CURRENTQ']
            #print("Currently Assigned to: %s"%(whoHasCase))
            
            rollyet = case['ROLLSTATUS']
            if rollyet == "-":
                rollyet = "Hasn't Rolled"
                #print("Roll Status: " + rollyet)
            else:
                pass


            pcasePath = case['PCASEPATH']
            checkpath = pcasePath

            caseStatus = case['CASESTATUS']
            if caseStatus == "-":
                caseStatus = "Not Set"
            #print("Case Status: %s"%(caseStatus))

            pcaseNumber = case['PCASENUMBER']
            #print("Pcase Number: %s"%(pcaseNumber))
            
            if os.path.exists(checkpath):
                #print("Pcase Path: %s"%(checkpath))
            
                fdtFolder = "%s\FDT"%(checkpath)
                parserConfigFolder = "%s\VC\PARSERCONFIGS"%(checkpath)
                releaseBinFolder = "%s\VC\RELEASE\BIN"%(checkpath)
                scriptsFolder = "%s\VC\SCRIPTS"%(checkpath)

                try:
                    # Find Logos, or images.
                    
                    if os.listdir(fdtFolder) != []:
                        imgSet = ('.png','.jpg','.jpeg')

                        for file in os.listdir(fdtFolder):

                            
                            if file.endswith(imgSet):
                                #print("Image Found: %s"%(file))
                                imagesList.append(file)
                            
                            if file.endswith(".csv"):
                                #print("Template Found: %s"%(file))
                                templateList.append(file)
  
                            secondCheck = os.path.join(fdtFolder, file)
                            if os.path.isdir(secondCheck) and os.path.join(fdtFolder,file) != os.path.join(fdtFolder,"TERMS"):
                                #print("%s in FDT"%(file))
                                fdtMiscList.append(file)

                            pdf = '.pdf'
                            if os.path.isdir(secondCheck) and os.path.join(fdtFolder,file) == os.path.join(fdtFolder,"TERMS"):
                                if os.listdir(os.path.join(fdtFolder,"TERMS")) != []:
                                    for file in os.listdir(os.path.join(fdtFolder,"TERMS")):
                                        if file.endswith(pdf):
                                            #print("Terms Backer Found: %s"%(file))
                                            termsBackerList.append(file)

                except FileNotFoundError:
                    pass
                
                try:
                    # Find Scripts
                    if os.listdir(scriptsFolder) != []:
 
                        for file in os.listdir(scriptsFolder):

                            if file.endswith(".py"):
                                #print("Script Found: %s"%(file))
                                pythonScriptList.append(file)

                             
                            secondCheck = os.path.join(scriptsFolder, file)
                            os.path.isdir(secondCheck)
                            if os.path.isdir(secondCheck):
                                #print("%s in SCRIPTS"%(file))
                                scriptMiscList.append(file)

                except FileNotFoundError:
                    pass

                try: # Find Parser
                    if os.listdir(parserConfigFolder) != []:

                        for file in os.listdir(parserConfigFolder):

                            if file.endswith(".xml"):
                                #print("Parser Found: %s"%(file))
                                parserFileList.append(file)

                            secondCheck = os.path.join(parserConfigFolder, file)
                            os.path.isdir(secondCheck)
                            if os.path.isdir(secondCheck):
                                #print("%s in PARSERCONFIGS"%(file))
                                parserMiscList.append(file)
                            
                except FileNotFoundError:
                    pass

                
                try:# Check ReleaseBin Folder
                    if os.listdir(releaseBinFolder) != []:

                        for file in os.listdir(releaseBinFolder):
                            fileExts = (".txt",".csv",".xml", ".dat")
                            if file.endswith(fileExts):
                                #print("File Found in ReleaseBin: %s"%(file))
                                releaseBinList.append(file)

                            secondCheck = os.path.join(releaseBinFolder, file)
                            os.path.isdir(secondCheck)
                            if os.path.isdir(secondCheck):
                                #print("%s in RELEASE\BIN"%(file))
                                releaseBinMiscList.append(file)

                except FileNotFoundError:
                    pass



                #print("-----------End Case %s Details----------------"%(matches))

            finalTrewViewDict = {pcaseNumber:[sfCaseNumber,caseReason,whoHasCase,rollyet, caseStatus, pcasePath, templateList, imagesList, termsBackerList, pythonScriptList,parserFileList]}#,[fdtMiscList,scriptMiscList,parserMiscList,releaseBinMiscList]]}
            finalTreeViewList.append(finalTrewViewDict)

    if int(matches) == 0:
        noConflictMessage = "find_OpenCases Notification: No Current Conflicts"
        finalTrewViewDict = {"NOCONFLICTS":"CURRENTLY NO CONFLICTS FOR"}
        finalTreeViewList.append(finalTrewViewDict)



    return finalTreeViewList












