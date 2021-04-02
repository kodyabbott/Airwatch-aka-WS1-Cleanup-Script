#!/opt/homebrew/bin/python3
import requests
import json
import datetime
import os
import concurrent.futures
import time
### added logging import ###
### import logging
from google.oauth2 import service_account
from googleapiclient.discovery import build
from ldap3 import Server, \
    Connection, \
    AUTO_BIND_NO_TLS, \
    SUBTREE, \
    ALL_ATTRIBUTES
###from ldap3 import Server, Connection, ALL, NTLM
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
from pathlib import Path
env_path = Path.cwd() / '.env'
load_dotenv(dotenv_path=env_path)

time1 = time.perf_counter()
### added for logging connections ###
### logging.basicConfig(filename='client_application.log', level=logging.DEBUG)
### from ldap3.utils.log import set_library_log_detail_level, get_detail_level_name, EXTENDED
### set_library_log_detail_level(EXTENDED)
## end logging ###
### Functions ###

def myconverter(o):
    if isinstance(o, datetime.date):
        return o.__str__()

def getFileVault(i):
    id = i['Id']
    r = requests.get(f'{mdm_url}/api/mdm/devices/{id}/security', headers= {'aw-tenant-code': aw_tenant_code, 'Accept': 'application/json'}, auth=authWrkSpc)
    iSec = ''
    iSec = r.json()
    if 'PersonalRecoveryKey' in iSec:
        if iSec['PersonalRecoveryKey'] != '':
            i['EncrypKey'] = True
        else:
            if i['LastSeen'] < hundredEightyThreeDays:
                i['Reason'] = 'Apple OSX device unseen for 183 days, no filevault stored'
    else:
        if i['LastSeen'] < hundredEightyThreeDays:
            i['Reason'] = 'Apple OSX device unseen for 183 days, no filevault stored'
    if i['EncrypKey'] == True and i['LastSeen'] < thousandDays:
        i['Reason'] = 'Apple OSX device unseen for 1000 days with filevault'

def getBitLocker(i):
    id = i['Id']
    r = requests.get(f'{mdm_url}/api/mdm/devices/{id}/security', headers= {'aw-tenant-code': aw_tenant_code, 'Accept': 'application/json'}, auth=authWrkSpc)
    iSec = ''
    iSec = r.json()
    if 'PersonalRecoveryKey' in iSec:
        if iSec['PersonalRecoveryKey'] != '':
            i['EncrypKey'] = True
        else:
            if i['LastSeen'] < fourTwentySixDays:
                i['Reason'] = 'Windows 8+ device unseen for 426 days, no bitlock stored'
    else:
        if i['LastSeen'] < fourTwentySixDays:
            i['Reason'] = 'Windows 8+ device unseen for 426 days, no bitlock stored'
    if i['EncrypKey'] == True and i['LastSeen'] < thousandDays:
        i['Reason'] = 'Windows 8+ device unseen for 1000 days with Bitlocker'

def deleteDevice(d):
    id = d['Id']
    r = requests.delete(f'{mdm_url}/api/mdm/devices/{id}', headers= {'aw-tenant-code': aw_tenant_code, 'Accept': 'application/json'}, auth=authWrkSpc)
    if r.status_code == 200:
        d['Deleted'] = f'SUCCESS {today}'
    else:
        iSec = r.json()
        d['Deleted'] = f'FAIL {iSec["message"]}'

def deleteUser(u):
    id = u['Id']
    r = requests.delete(f'{mdm_url}/api/system/users/{id}/delete', headers= {'aw-tenant-code': aw_tenant_code, 'Accept': 'application/json'}, auth=authWrkSpc)
    if r.status_code == 200:
        u['Deleted'] = f'SUCCESS {today}'
    else:
        iSec = r.json()
        u['Deleted'] = f'FAIL {iSec["message"]}'

#Google Drive functions below
def getSpreadsheet(spreadName):
    service = build('drive', 'v3', credentials=credentials)

    #Getting list of files in "Airwatch Deletion" team drive to see if Google UMD OrgUnits spreadsheet exists.  If not, create it
    results = service.files().list(q='trashed = false', corpora='drive', pageSize=50, driveId=driveID, supportsAllDrives=True, includeItemsFromAllDrives=True, fields='files/id, files/name').execute() # pylint: disable=E1101
    
    spreadsheets = results.get('files', [])

    #Getting/Creating current year Devices spreadsheet ID
    SpreadId = ''
    for h in spreadsheets:
        if h['name'] == spreadName:
            SpreadId = h['id']
    if SpreadId == '':
        file_metadata = {
            'name': spreadName,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': [
                driveID
                ]
        }
        file = service.files().create(body=file_metadata, supportsAllDrives=True).execute() # pylint: disable=E1101
        SpreadId = file['id']
    return SpreadId

def getSheet(SpreadId):
    service = build('sheets', 'v4', credentials=credentials)
    getSheets = service.spreadsheets().get(spreadsheetId=SpreadId).execute() # pylint: disable=E1101
    SheetId = ''
    getSheets = getSheets['sheets']
    for sheet in getSheets:
        if sheet['properties']['sheetId'] == 0 and sheet['properties']['title'] == "Sheet1":
            body = {
                "requests": [{
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": "0",
                            "title": f"{month}"
                        },
                        "fields": "title"
                    }
                }]
            }
            sheetNameChange = service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101
            SheetId = 0
        elif sheet['properties']['title'] == f"{month}":
            SheetId = sheet['properties']['sheetId']
    if SheetId == '': 
        body = {"requests": [{"addSheet": {"properties": {"title": f"{month}"}}}]}
        sheetNameChange = service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101
        SheetId = sheetNameChange['replies'][0]['addSheet']['properties']['sheetId']
    return SheetId

def formatSheet(SpreadId, SheetId, Columns):
    service = build('sheets', 'v4', credentials=credentials)
    getHeaderRowInfo = service.spreadsheets().values().get(spreadsheetId=SpreadId, range=f"{month}!1:1").execute() # pylint: disable=E1101
    if not 'values' in getHeaderRowInfo :
        #Set font for all used Columns to "Proxima Nova"
        body = {"requests": [{"repeatCell": {"cell": {"userEnteredFormat": {"textFormat": {"fontFamily": "Proxima Nova"}}},"range": {"sheetId": SheetId,"startColumnIndex": 0,"endColumnIndex": 6},"fields": "userEnteredFormat(textFormat)"}}]}
        #formatFont = service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute()  
        service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute()   # pylint: disable=E1101
        #Create row ledger where even rows are slightly different color
        body = {"requests": [{"addConditionalFormatRule": {"index": 1,"rule": {"ranges": [{"sheetId": SheetId,"startRowIndex": 1}],"booleanRule": {"condition": {"type": "CUSTOM_FORMULA","values": [{"userEnteredValue": "=ISEVEN(ROW())"}]},"format": {"backgroundColor": {"blue": 0.953,"green": 0.945,"red": 0.941},"textFormat": {"foregroundColor": {"blue": 0,"green": 0,"red": 0}}}}}}}]}
        service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101
        #Format several parts of the Header Row
        body = {"requests": [{"repeatCell": {"range": {"sheetId": f"{SheetId}","startRowIndex": 0,"endRowIndex": 1},"cell": {"userEnteredFormat": {"backgroundColor": {"blue": 0.375,"red": 0.27,"green": 0.327},"textFormat": {"fontFamily": "Proxima Nova","bold": True,"fontSize": 12,"foregroundColor": {"blue": 1,"green": 1,"red": 1}},"horizontalAlignment": "CENTER","verticalAlignment": "MIDDLE"}},"fields": "userEnteredFormat(backgroundColor, textFormat, horizontalAlignment, verticalAlignment)"}}]}
        service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101
        #Increase size of Header row
        body = {"requests": [{"updateDimensionProperties": {"fields": "pixelSize","properties": {"pixelSize": 30},"range": {"sheetId": f"{SheetId}","startIndex": 0,"endIndex": 1,"dimension": "ROWS"}}}]}
        service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101
        #Populate the Column values of the Header row
        body = {"valueInputOption": "USER_ENTERED","data": [{"range": f"{month}","majorDimension": "ROWS","values": [Columns]}]}
        service.spreadsheets().values().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101
        #Freeze the Header row to prevent sorting
        body = {"requests": [{"updateSheetProperties": {"properties": {"gridProperties": {"frozenRowCount": 1},"sheetId": SheetId},"fields": "gridProperties.frozenRowCount"}}]}
        service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101
        if 'Deleted' in Columns:
            indexDeleted = Columns.index('Deleted')
            #Adding green background to cells in FlaggedForAutoFix column marked Yes
            body = { "requests": [ { "addConditionalFormatRule": { "index": 0, "rule": { "booleanRule": { "condition": { "values": [ { "userEnteredValue": "SUCCESS" } ], "type": "TEXT_STARTS_WITH" }, "format": { "backgroundColor": { "red": 0.717, "green": 0.882, "blue": 0.804 } } }, "ranges": [ { "startColumnIndex": indexDeleted, "endColumnIndex": indexDeleted+1, "startRowIndex": 1, "sheetId": SheetId } ] } } } ] }
            service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101
            #Adding red background to cells in FlaggedForAutoFix column marked No
            body = { "requests": [ { "addConditionalFormatRule": { "index": 0, "rule": { "booleanRule": { "condition": { "values": [ { "userEnteredValue": "FAIL" } ], "type": "TEXT_STARTS_WITH" }, "format": { "backgroundColor": { "red": 0.957, "green": 0.780, "blue": 0.765 } } }, "ranges": [ { "startColumnIndex": indexDeleted, "endColumnIndex": indexDeleted+1, "startRowIndex": 1, "sheetId": SheetId } ] } } } ] }
            service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101

def addSheetCells(SpreadId, cellValues):
    service = build('sheets', 'v4', credentials=credentials)
    body = {"valueInputOption": "USER_ENTERED","data": [{"range": f"{month}!A2","values": cellValues}]}
    service.spreadsheets().values().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101

def resizeSheetColumns(SpreadId, SheetId):
    service = build('sheets', 'v4', credentials=credentials)
    body = {"requests": [{"autoResizeDimensions": {"dimensions": {"dimension": "COLUMNS", "sheetId": SheetId}}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=SpreadId, body=body).execute() # pylint: disable=E1101


# Get expiration days for Objects
#fourTwentySix will delete Windows workstations not seen in 426 days without a bitlocker key stored in Airwatch
fourTwentySixDays = datetime.date.today() - datetime.timedelta(426)
#thousandDays will delete Windows workstations with a bitlocker key and MacOSX with Filevault unseen for 1000 days
thousandDays = datetime.date.today() - datetime.timedelta(1000)
# 183 days will delete MacOSX (without Filevault), Android, IOS, Chrome devices not seen in 183 days (6 months)
hundredEightyThreeDays = datetime.date.today() - datetime.timedelta(183)

# Getting strings for current year, month, and date for various queries
thisYear = datetime.date.today().year
month = datetime.date.today().strftime('%b')
today = datetime.date.today().strftime('%Y-%m-%d')

### Environment Variables ###
#Get creds for Workspace One API
aduser = os.environ.get('USER_AD')
adpwd = os.environ.get('USER_ADPWD')
aw_tenant_code = os.environ.get('aw-tenant-code')
mdm_url = os.environ.get('MDM_URL')
#Create auth token for Workspace One API
authWrkSpc = HTTPBasicAuth(aduser, adpwd)

#Get creds for Active Directory
user = os.environ.get('ADuser')
password = os.environ.get('ADpwd')
ADserver = os.environ.get('ADserver')
ADpath = os.environ.get('ADpath')
ADobjectCategory = os.environ.get('ADobjectCategory')

#Get Service Now creds
SNuser = os.environ.get('SN_User')
SNpwd = os.environ.get('SN_Pwd')
SN_URL = os.environ.get('SN_URL')
#Create auth token for Service Now
authSN = HTTPBasicAuth(SNuser, SNpwd)

#Google Drive environment variables
driveID = os.environ.get('driveID')
SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'service.json'
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES, subject='airwatch@gserviceaccount.com')

#Request all devices from Workspace One API
r = requests.get(f'{mdm_url}/api/mdm/devices/search?pagesize=50000', headers= {'aw-tenant-code': aw_tenant_code, 'Accept': 'application/json'}, auth=authWrkSpc)
airWatchDevs = r.json()
devices = airWatchDevs['Devices']

#Request all users from Workspace One API
r = requests.get(f'{mdm_url}/api/system/users/search?pagesize=50000', headers= {'aw-tenant-code': aw_tenant_code, 'Accept': 'application/json'}, auth=authWrkSpc)
airWatchUsers = r.json()

#Limit list of users for AD synced only
users = [item for item in airWatchUsers['Users'] if item['SecurityType'] == 1]

#Sort users list by UserName
users.sort(key=lambda x: x['UserName'])

#Get users from AD, post results to matching Users pulled from Airwatch
server = Server(ADserver, use_ssl=True)
conn = Connection(server, user=user, password=password, auto_bind=AUTO_BIND_NO_TLS)
conn.search (ADpath, ADobjectCategory, attributes=['sAMAccountName', 'userAccountControl', 'whenChanged', 'objectclass'], paged_size=1000)
cookie = conn.result['controls']['1.2.840.113556.1.4.319']['value']['cookie']
ADresults = conn.entries
while cookie:
    conn.search (ADpath, ADobjectCategory, attributes=['sAMAccountName', 'userAccountControl', 'whenChanged', 'objectclass'], paged_size=1000, paged_cookie=cookie)
    ADresults += conn.entries
    cookie = conn.result['controls']['1.2.840.113556.1.4.319']['value']['cookie']

ADresultsFinal = []
for tempDev in ADresults:
    som = json.loads(tempDev.entry_to_json())
    test = {}
    if som['attributes']['whenChanged'] != []:
        test['whenChanged'] = som['attributes']['whenChanged'].__str__()[2:-8]
        test['whenChanged'] = datetime.datetime.strptime(test['whenChanged'], "%Y-%m-%d %H:%M:%S").date()
    else:
        test['whenChanged'] = ''
    test['sAMAccountName'] = som['attributes']['sAMAccountName'].__str__()[2:-2]
    test['userAccountControl'] = som['attributes']['userAccountControl'].__str__()[1:-1]
    ADresultsFinal.append(test)

#Appending key/value pairs to users dictionary
for user in users:
    user['Id'] = user['Id']['Value']
    user['Reason'] = ''
    user['ADusername'] = ''
    user['AD_Disabled'] = False
    user['AD_Date'] = ''
    user['Deleted'] = ''
    temp = [item for item in ADresultsFinal if item['sAMAccountName'] == user['UserName']]
    if temp != []:
        user['ADusername'] = temp[0]['sAMAccountName']
        if temp[0]['userAccountControl'] == '514':
            user['AD_Disabled'] = True
        user['AD_Date'] = temp[0]['whenChanged']
    else:
        temp = [item for item in ADresultsFinal if item['sAMAccountName'] == f'old-{user["UserName"]}']
        if temp != []:
            user['ADusername'] = temp[0]['sAMAccountName']
            if temp[0]['userAccountControl'] == '514':
                user['AD_Disabled'] = True
            user['AD_Date'] = temp[0]['whenChanged']
        else:    
            temp = [item for item in ADresultsFinal if item['sAMAccountName'] == f'old-old-{user["UserName"]}']
            if temp != []:
                user['ADusername'] = temp[0]['sAMAccountName']
                if temp[0]['userAccountControl'] == '514':
                    user['AD_Disabled'] = True
                user['AD_Date'] = temp[0]['whenChanged']

#Create devices object filtering api response data with key/values needed
devResults = []

#Formatting fields with datetime stamp to filter devices
for tempDev in devices:
    tempDev['LastSeen'] = datetime.datetime.strptime(tempDev['LastSeen'], "%Y-%m-%dT%H:%M:%S.%f").date() #.strftime("%Y-%m-%d")
    tempDev['LastEnrolledOn'] = datetime.datetime.strptime(tempDev['LastEnrolledOn'], "%Y-%m-%dT%H:%M:%S.%f").date()
    #Getting object key/values pairs for each device
    devObj = {
        'Id': tempDev['Id']['Value'],
        'DeviceFriendlyName': tempDev['DeviceFriendlyName'],
        'SerialNumber': tempDev['SerialNumber'],
        'UserEmailAddress': tempDev['UserEmailAddress'],
        'Ownership': tempDev['Ownership'],
        'PlatformId': tempDev['PlatformId']['Id']['Value'],
        'Platform': tempDev['Platform'],
        'Model': tempDev['Model'],
        'OperatingSystem': tempDev['OperatingSystem'],
        'LastSeen': tempDev['LastSeen'],
        'LastEnrolledOn': tempDev['LastEnrolledOn'],
        'UserName': tempDev['UserName'],
        'Reason': '',
        'EncrypKey': False,
        'SNowRetired': '',
        'SNowInventoried': '',
        'SNowSubstate': '',
        'Deleted': ''
    }
    devResults.append(devObj)

#get AppleOsx devices
appleOSXdevices = [item for item in devResults if "AppleOsX" in item['Platform']]

#Run getFileVault function to determine if Airwatch has recovery token
#getFileVault function will tag AppleOsx devices for deletion if criteria met
with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(getFileVault, appleOSXdevices)


#tag ios/Droid/Chrome devices not seen in past 183 days
iosDroidExpired = [item for item in devResults if ('Android' in item['Platform'] or 'Chrome' in item['Platform'] or item['Platform'] == 'Apple') and item['LastSeen'] < hundredEightyThreeDays]
for i in iosDroidExpired:
    i['Reason'] = 'ios/Android/Chrome device LastSeen over 183 days'

#Get Windows 7 devices that haven't been seen over 425 days
win7 = [item for item in devResults if 'Win' in item['Platform'] and '6.1.' in item['OperatingSystem'] and item['LastSeen'] < fourTwentySixDays]
for i in win7:
    i['Reason'] = 'Windows 7 device LastSeen over 425 days'


#Getting bitlocker token for Windows 8+ devices and marking for deletion if lastseen over 426 days without bitlocker,
#and over 1000 days if it has bitlocker token
win8plus = [item for item in devResults if 'Win' in item['Platform'] and not '6.1.' in item['OperatingSystem']]

with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(getBitLocker, win8plus)

# Getting Service Now info, appending to matching device in devices object
r = requests.get(f'{SN_URL}/api/now/table/alm_hardware?sysparm_query=substatus%3Ddisposed%5EORsubstatus%3Dpending_disposal%5Einstall_status%3D7&sysparm_fields=serial_number%2Csubstatus%2Cretired%2Cu_csi_date_inventoried', headers= {"Content-Type":"application/json","Accept":"application/json"}, auth=authSN)
data = r.json()
SNdevices = data['result']

for tempDev in SNdevices:
    if tempDev['u_csi_date_inventoried'] != "":
        tempDev['u_csi_date_inventoried'] = datetime.datetime.strptime(tempDev['u_csi_date_inventoried'], "%Y-%m-%d %H:%M:%S").date()
    if tempDev['retired'] != "":
        tempDev['retired'] = datetime.datetime.strptime(tempDev['retired'], "%Y-%m-%d").date()

for SNdev in SNdevices:
    deviceSerial = [item for item in devResults if item['SerialNumber'] == SNdev['serial_number']]
    if deviceSerial != []:
        for devWithMatchingSerial in deviceSerial:
            devWithMatchingSerial['SNowRetired'] = SNdev['retired']
            devWithMatchingSerial['SNowInventoried'] = SNdev['u_csi_date_inventoried']
            devWithMatchingSerial['SNowSubstate'] = SNdev['substatus']
            if devWithMatchingSerial['SNowSubstate'] == "disposed" and devWithMatchingSerial['Reason'] == '' and devWithMatchingSerial['EncrypKey'] == False:
                devWithMatchingSerial['Reason'] = "Device disposed in ServiceNow, no bitlock/Filevault key"

#Look for duplicate devices by serial numbers
dupSer = []
for obj in devResults:
    if obj not in dupSer:
        objCheck = [item for item in devResults if item['SerialNumber'] == obj['SerialNumber'] and item['Reason'] == '']
        if len(objCheck) > 1:
            dupSer.append(objCheck)

for devDupPair in dupSer:
    devDupPairCount = len(devDupPair)
    dateEnrollList = []
    lastEnrolled = []
    lastSeen = []
    for dupDev in devDupPair:
        lastEnrolled.append(dupDev['LastEnrolledOn'])
        lastSeen.append(dupDev['LastSeen'])
        outputToList = f'{dupDev["LastEnrolledOn"]}, {dupDev["LastSeen"]}, {dupDev["EncrypKey"]}'
        dateEnrollList.append(outputToList)
    newestLastEnrolled = max(lastEnrolled)
    newestLastSeen = max(lastSeen)
    #if len([item for item in item2 if item['LastEnrolledOn'] == newestLastEnrolled]) < 2:
    for item4 in devDupPair:
        if item4['LastEnrolledOn'] != newestLastEnrolled and item4['EncrypKey'] == False and item4['Reason'] == "":
            item4['Reason'] = 'Old duplicate, no encryption token'

###   Evaluate deleted or disabled AD Users to determine if they can be deleted   ###
#Mark user accounts missing or disabled in AD with no enrolled devices
ADuserMissing = [item for item in users if item['ADusername'] == "" and item['EnrolledDevicesCount'] == ""]
for ADusermiss in ADuserMissing:
    ADusermiss['Reason'] = 'AD account deleted, no Enrolled Devices'
ADuserDisabled = [item for item in users if item['AD_Disabled'] == True and item['EnrolledDevicesCount'] == ""]
for ADuserdisa in ADuserDisabled:
    ADuserdisa['Reason'] = 'Disabled AD account, no Enrolled Devices'
ADuserswithDevices = [item for item in users if (item['AD_Disabled'] == True or item['ADusername'] == "") and item['EnrolledDevicesCount'] != ""]
for ADuserwDevice in ADuserswithDevices:
    #Get end users enrolled devices
    q = [item for item in devResults if item['UserName'] == ADuserwDevice['UserName']]
    devWithEncrypKey = False
    for i in q:
        if i['Reason'] == '':
            if i['EncrypKey'] == True:
                devWithEncrypKey = True
            else:
                i['Reason'] = 'Deleted/Disabled AD account enrolled device without bitlock/filevault'
    if devWithEncrypKey == False:
        ADuserwDevice['Reason'] = 'Deleted/Disabled AD account, enrolled devices deleted'

#  ###   Delete devices with Reason tag   ###
devs2Delete = [item for item in devResults if item['Reason'] != '']
with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(deleteDevice, devs2Delete)

gsheetDevs = [{'LastSeen': myconverter(tempDev['LastSeen']), 'DeviceFriendlyName': tempDev['DeviceFriendlyName'], 'UserEmailAddress': tempDev['UserEmailAddress'], 'SerialNumber': tempDev['SerialNumber'], 'DevId': tempDev['Id'], 'Reason': tempDev['Reason'], 'Deleted': tempDev['Deleted']} for tempDev in devs2Delete]

#  ###   Delete users with Reason tag   ###
users2Delete = [item for item in users if item['Reason'] != '']
with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(deleteUser, users2Delete)

gsheetUsrs = [{'UserName': tempUsr['UserName'], 'FirstName': tempUsr['FirstName'], 'LastName': tempUsr['LastName'], 'Email': tempUsr['Email'], 'Id': tempUsr['Id'], 'Reason': tempUsr['Reason'], 'Deleted': tempUsr['Deleted']} for tempUsr in users2Delete]

#Get/Create Google Spreadsheet for devices and users
devsSpreadSheetName = f'{thisYear} Airwatch Devices to Delete'
devSpreadId = getSpreadsheet(devsSpreadSheetName)

usrSpreadSheetName = f'{thisYear} Airwatch Users to Delete'
usrSpreadId = getSpreadsheet(usrSpreadSheetName)

#Get/Create Google Sheet tab for devices and users
devSheetId = getSheet(devSpreadId)
usrSheetId = getSheet(usrSpreadId)

#Checking for Header row in this months sheet tab of spreadsheet, if not exist, create and format
devColumns = [item for item in gsheetDevs[0]]
formatSheet(devSpreadId, devSheetId, devColumns)
usrColumns = [item for item in gsheetUsrs[0]]
formatSheet(usrSpreadId, usrSheetId, usrColumns)

#Alter the python dictionary objects to python list for gsheet upload
gsheetDevs = [list(item.values()) for item in gsheetDevs]
gsheetUsrs = [list(item.values()) for item in gsheetUsrs]

#Add the lists of users and devices to the appropriate gsheet
addSheetCells(devSpreadId, gsheetDevs)
addSheetCells(usrSpreadId, gsheetUsrs)

#Resize the Columns to a width that fits the column data
resizeSheetColumns(devSpreadId, devSheetId)
resizeSheetColumns(usrSpreadId, usrSheetId)

time2 = time.perf_counter()
print(f'ADresults created in {time2-time1} seconds')