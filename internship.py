import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import zipfile

def openFile():
    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_experimental_option("prefs", {"download.default_directory":"//Users//advikmareedu//Desktop//Internship"})
    username= "CP_nmule"
    password = "cenergypower123"
    driver = webdriver.Chrome('/Users/advikmareedu/Desktop/Internship/chromedriver_mac_arm64/chromedriver',options=chromeOptions)
    driver.get("https://www.pge.com/b2b/distribution-resource-planning/grid-needs-assessment-map.html")
    driver.maximize_window()
    time.sleep(8)

    user = driver.find_element(By.ID,"username")
    user.send_keys(username)

    passW = driver.find_element(By.XPATH,'//*[@id="loginForm"]/div[2]/input')
    passW.send_keys(password)

    driver.find_element(By.ID,"submit").click()
    time.sleep(5)
    driver.find_element(By.XPATH,'/html/body/div[9]/div[3]/div/button').click()
    time.sleep(3)
    driver.find_element(By.XPATH,'//*[@id="help"]/a[2]').click()
    time.sleep(70)
    with zipfile.ZipFile('/Users/advikmareedu/Desktop/Internship/PGE_DIDF_Tables_Public.zip', 'r') as zip_ref:
        zip_ref.extractall('/Users/advikmareedu/Desktop/Internship')







def getSubNames():
    appendix=pd.read_excel('/Users/advikmareedu/Desktop/Internship/DIDF_data_public/Appendices/PGE_2022_GNA_Appendix_D-F_Public.xlsx', 'GNA Bank & Feeder Capacity', header=5)
    facilityNames = appendix.loc[:,'Facility Name']
    facilityTypes = appendix.loc[:,'Facility Type']
    subNames = []

    for i in range(len(facilityNames)):
        name = facilityNames[i]
        if "POTRERO (SF A)" in name or "MISSION (SF X)" in name or "MARTIN (SF H)" in name:
            name= name.replace("POTRERO (SF A)","POTRERO PP (A)")
            name = name.replace("MISSION (SF X)","MISSION (X)")
            name = name.replace("MARTIN (SF H)","SF H")
            if facilityTypes[i] == "Bank":
                ind = facilityNames[i].find("BANK")
                indSub = facilityNames[i].find("SUB")
                name = facilityNames[i][:max(ind,indSub)-1]
            else:
                name = name[:len(name)-5]
                
        else:
            if facilityTypes[i] == "Feeder":
                name = name.replace(" NEW ", " ")
                sPar = name.find("(")
                ePar = name.find(")")
                if sPar != -1:
                    name = name[:sPar] + name[ePar+1:]
                    name.strip()
                name = name[:len(name)-5]
            else:
                ind = facilityNames[i].find("BANK")
                indSub = facilityNames[i].find("SUB")
                name = facilityNames[i][:max(ind,indSub)-1]
        if facilityNames[i] == "MISSION (SF X) 1126 (Warriors Arena)":
            name = "MISSION (X)"
        if "HUNTERS POINT" in name:
            name = "HUNTERS POINT (P)"
        if "POTRERO" in name:
            name = "POTRERO PP (A)"
        if "MISSION X" in name:
            name = "MISSION (X)"    
        name = name.strip()
        subNames.append(name)
    appendix['Substation Names'] = subNames
    appendix.to_excel('Expanded_PGE_2022_GNA_Data.xlsx', index = False)



def getBankNumbers():
    appendix=pd.read_excel('Expanded_PGE_2022_GNA_Data.xlsx', 'Sheet1')
    facilityNames = appendix.loc[:,'Facility Name']
    facilityTypes = appendix.loc[:,'Facility Type']
    subNames = appendix.loc[:,'Substation Names']
    bankNumbers = []
    for i in range(appendix.shape[0]):
        if "Bank" in facilityTypes[i]:
            bankNumbers.append("null")
        else:
            c = i
            while (c >= 0 and facilityTypes[c] == "Feeder") or (c>= 0 and subNames[i]!= subNames[c]):
                c-=1
            
            if c!=-1:
                bankNumbers.append(facilityNames[c])
            else:
                bankNumbers.append("null")
    appendix['Bank #'] = bankNumbers
    appendix.to_excel('Expanded_PGE_2022_GNA_Data.xlsx', index = False)




def findDG(facilityName:str):
    master=pd.read_excel('California_Master.xlsx', 'PG&E Feeder Details')
    feederNames = master['Feeder Name']
    dgCol = master['Total DG (kW)']
    for i in range(master.shape[0]):
        if facilityName == feederNames[i]:
            return dgCol[i]
    return 0   

def getAggregateDG():
    appendix=pd.read_excel('Expanded_PGE_2022_GNA_Data.xlsx')
    facilityNames = appendix['Facility Name']
    facilityTypes = appendix['Facility Type']
    bankNumbers = appendix['Bank #']
    aggregateDG = []
    for i in range(appendix.shape[0]):
        count = 0
        if facilityTypes[i] == "Bank":
            for c in range(i+1,appendix.shape[0]):
                if bankNumbers[c] == facilityNames[i]:
                    count += findDG(facilityNames[c])
            aggregateDG.append(count)
        else:
            aggregateDG.append(0)
    appendix['Aggregate DG'] = aggregateDG
    appendix.to_excel('Expanded_PGE_2022_GNA_Data.xlsx', index = False)   



###########################

def findLoad(facilityName:str,):
    appendix=pd.read_excel('/Users/advikmareedu/Desktop/Internship/DIDF_data_public/Appendices/PGE_2022_GNA_Appendix_D-F_Public.xlsx', 'GNA Bank & Feeder Capacity', header = 5)
    facilityNames = appendix.loc[:,'Facility Name']
    loadCol = appendix.loc[:,'Facility Loading (MW)']
    for i in range(len(facilityNames)):
        if facilityName == facilityNames[i]:
            return loadCol[i]
    return 0   

# def getAggregateLoad():
#     appendix=pd.read_excel('tester.xlsx')
#     facilityNames = appendix['Facility Name']
#     facilityTypes = appendix['Facility Type']
#     bankNumbers = appendix['Bank #']
#     aggregateLoad = []
#     for i in range(appendix.shape[0]): 
#         count = 0
#         if facilityTypes[i] == "Bank":
#             for c in range(i+1,appendix.shape[0]):
#                 if bankNumbers[c] == facilityNames[i]:
#                     newLoad = findLoad(facilityNames[c])
#                     if type(newLoad) is not str:
#                         count += findLoad(facilityNames[c])
#             aggregateLoad.append(count)
#         else:
#             aggregateLoad.append(0)
#         print(i)
#     appendix['Aggregate Load'] = aggregateLoad
#     appendix.to_excel('tester.xlsx', index = False)   


#########################

def getSubstationSheet():
    appendix=pd.read_excel('Expanded_PGE_2022_GNA_Data.xlsx')
    subNames = appendix['Substation Names']
    subCol = []
    for i in range(appendix.shape[0]): 
        if subNames[i] not in subCol:
            subCol.append(subNames[i])
    names = {'Substation Name' : subCol }
    newSheet = pd.DataFrame(names)
    newSheet.to_excel('SubstationData.xlsx', index = False)  


  

def aggDGSubSheet():
    appendix= pd.read_excel('Expanded_PGE_2022_GNA_Data.xlsx')
    subSheet= pd.read_excel('SubstationData.xlsx')
    subSheetNames = subSheet['Substation Name']
    appendixNames = appendix['Substation Names']
    aggDG = appendix['Aggregate DG']
    subDG = []
    for i in range(subSheet.shape[0]):
        count = 0
        for j in range(appendix.shape[0]):
            if subSheetNames[i] == appendixNames[j]:
                count+= aggDG[j]
        subDG.append(count)
    subSheet['Aggregate DG'] = subDG
    subSheet.to_excel('SubstationData.xlsx', "Substation Data", index = False)

  

def aggLoadSubSheet():
    appendix= pd.read_excel('Expanded_PGE_2022_GNA_Data.xlsx')
    subSheet= pd.read_excel('SubstationData.xlsx')
    subSheetNames = subSheet['Substation Name']
    appendixNames = appendix['Substation Names']
    facilityNames = appendix['Facility Name']
    facilityTypes = appendix['Facility Type']
    subLoad = []
    for i in range(subSheetNames.shape[0]): 
        count = 0
        for j in range(appendixNames.shape[0]):
            if facilityTypes[j]== "Bank" and subSheetNames[i] == appendixNames[j]:
                newLoad = findLoad(facilityNames[j])
                if type(newLoad) is not str:
                    count += newLoad
        subLoad.append(count)
    subSheet['Aggregate Load'] = subLoad
    subSheet.to_excel('SubstationData.xlsx', "Substation Data", index = False)



def aggSubSheet():
    aggDGSubSheet()
    aggLoadSubSheet() 


def getData():
    openFile()
    getSubNames()
    getBankNumbers()
    getAggregateDG()
    getSubstationSheet()
    aggSubSheet()

getData()

