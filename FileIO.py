''' 
RUN THIS SCRIPT:

SYNTAX:
1. To add a new Spreadsheet
python3 FileIO.py -n NewSpreadSheetName FileToBeProcessed
2. To add to the previously written SpreadSheet
python3 FileIO.py -p FileToBeProcessed

'''

import sys
import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
import json

NotCancelled=[]

'''
Function to create a spreadsheet of the mentioned Title.
And share it with personal ID
'''

def create_spreadsheet(title):
    key_path = "google-json-cred.json"  # Path to your service account JSON key

    # Create credentials using the service account file
    creds = service_account.Credentials.from_service_account_file(key_path, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])

    try:
        sheets_service = build("sheets", "v4", credentials=creds)

        spreadsheet = {"properties": {"title": title}}

        response = sheets_service.spreadsheets().create(body=spreadsheet, fields="spreadsheetId").execute()
        
        spreadsheet_id = response.get("spreadsheetId")
        #print(f"Spreadsheet created with ID: {spreadsheet_id}")

        drive_service = build("drive", "v3", credentials=creds)
        permission = {
            "type": "user",
            "role": "writer",  
            "emailAddress": ""  # Share with your email
        }

        drive_service.permissions().create(fileId=spreadsheet_id, body=permission, fields="id").execute()
        

        return spreadsheet_id

    except HttpError as error:
        #print(f"An error occurred: {error}")
        return None


'''
Funtion to create a new Sheet with the fileName 
'''
def AddNewSheet(spreadsheetId, worksheetName):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = 'google-json-cred.json'
    SPREADSHEET_ID = spreadsheetId

    creds = None
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    service = build("sheets", "v4", credentials=creds)

    batch_update_values_request_body = {
            'requests': [
                {
                    'addSheet': {
                        'properties': {
                            'title': worksheetName
                        }
                    }
                }
            ]
        }
    request = service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=batch_update_values_request_body)
    response = request.execute()


'''
Deprecated - Still in Progress to delete the first sheet 
'''
def deleteSheet1(spreadsheetId):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = 'google-json-cred.json'
    SPREADSHEET_ID = spreadsheetId

    creds = None
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    service = build("sheets", "v4", credentials=creds)

    batch_update_values_request_body = {
            'requests': [
                {
                    "deleteSheet": {
                        "sheetId": 0
                    }
                }
            ]
        }
    request = service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=batch_update_values_request_body)
    response = request.execute()


'''
TO check is the spreadsheet ID is valid 
'''
def CheckIfSpreadsheetExists(spreadSheet_ID):
    key_path = "google-json-cred.json" 
    creds = service_account.Credentials.from_service_account_file(key_path, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    try:
        sheets_service = build("sheets", "v4", credentials=creds)
        response= sheets_service.spreadsheets().get(
            spreadsheetId=spreadSheet_ID
        ).execute()

        if(response):
            #print(response)
            #print("SpreadSheet Already Exist")
            return True
        else:
            #print("Need to create a new spreadsheet")
            return False

    except HttpError as error:
        print(f"An error occurred: {error}")



'''
Check if the particular sheet name is valid 
'''
def checkIfSheetExist(spreadSheet_ID, sheetName):
    key_path = "google-json-cred.json" 
    creds = service_account.Credentials.from_service_account_file(key_path, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    try:
        sheets_service = build("sheets", "v4", credentials=creds)
        response= sheets_service.spreadsheets().get(
            spreadsheetId=spreadSheet_ID
        ).execute()
        #print(response.get('sheets'))
        #print([i.get('properties').get('title') for i in response.get('sheets')])
        if(sheetName in [i.get('properties').get('title') for i in response.get('sheets')]):
            #print(response)
            #print("Sheet Already Exist")
            return True
        else:
            #print("Need to create a new spreadsheet")
            return False

    except HttpError as error:
        print(f"An error occurred: {error}")



''' 
deprecated
'''
def writeFile():
    # outputFile="processed.txt"
    # writeinFile=open(outputFile,"w")
    # writeinFile.writelines(NotCancelled)
    print(NotCancelled)



'''
deprecated 
'''
def printList():
    for x in NotCancelled:
        print(x)



'''
Function to process the file - seperate blocks that are valid 
'''
def processBlock(temp_list):
    if(temp_list[-1][:8]!="Canceled"):
                temporary_append_list=[]
                str=""
                count=0
                for stringComponents in temp_list:
                    count+=1
                    if(count==2):
                        continue
                    temporary_append_list.append(stringComponents)
                    str+=stringComponents+" | "

                NotCancelled.append(temporary_append_list)
                # NotCancelled.append(str[:-2]+"\n")

'''
Function to process file and send to processBlock
'''
def processFile(file):
    i=0
    temp_list=[]
    for x in file:
        if(len(x)==1):
            processBlock(temp_list)
            temp_list=[]
        else:
            temp_list.append(x[:-1])
    
    if(len(temp_list)>0):
        processBlock(temp_list)
    

'''
Read the file to be processed 
'''
def readFile(FileName):
    with open(FileName, "r") as file:
        processFile(file)


'''
Append the data to a specific sreadsheet 
'''
def append_data_to_specific_sheet(spreadsheet_id, sheet_name):
    key_path = "google-json-cred.json" 
    creds = service_account.Credentials.from_service_account_file(key_path, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    try:
        sheets_service = build("sheets", "v4", credentials=creds)
        range_ = f"{sheet_name}" 
        body = {
            "values": NotCancelled
        }

        sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_,
            valueInputOption="RAW",  
            body=body,
            insertDataOption="INSERT_ROWS"  
        ).execute()

        print("Success")
        
        #print(f"Data appended to {sheet_name} successfully!")

    except HttpError as error:
        print(f"An error occurred: {error}")



def GetspreadSheetId():
    file = open('RobinhoodSpreadSheetIDStore.txt','r')
    sID=file.readline().strip()
    file.close()
    return sID


def WriteSpreadSheetID(spreadSheetID):
    file = open('RobinhoodSpreadSheetIDStore.txt','w')
    file.writelines(spreadSheetID)
    file.close()
    

if __name__=="__main__":

    
    command=sys.argv[1]

    if(command=="-n"):
        spreadSheet_ID=create_spreadsheet(sys.argv[2])
        WriteSpreadSheetID(spreadSheet_ID)
        FileName=sys.argv[3]
        readFile(FileName)
        if(CheckIfSpreadsheetExists(spreadSheet_ID)):
            #print("SpreadSheet already exist")
            if(checkIfSheetExist(spreadSheet_ID, FileName)):
                print("Sheet already exist")
            else:
                AddNewSheet(spreadSheet_ID,FileName)
                append_data_to_specific_sheet(spreadSheet_ID, FileName)

    if(command=="-p"):
        spreadSheet_ID=GetspreadSheetId()
        FileName=sys.argv[2]
        readFile(FileName)
        if(CheckIfSpreadsheetExists(spreadSheet_ID)):
            #print("SpreadSheet already exist")
            if(checkIfSheetExist(spreadSheet_ID, FileName)):
                print("Sheet already exist")
            else:
                AddNewSheet(spreadSheet_ID,FileName)
                append_data_to_specific_sheet(spreadSheet_ID, FileName)
                

