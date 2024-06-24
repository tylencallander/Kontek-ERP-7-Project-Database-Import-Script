#
# Kontek Spare Parts / Inventory Program
# Project XLSX Import to Database
#
# Coryright Kontek Ecology Systems Inc 2021
#
# Email: cwelch@kontekecology.com
#
# Date: 2021-05-14
#

# Imports
import openpyxl
import psycopg2
import json

# Global Variables
DATABASE_HOST = "192.168.1.120"
DATABASE_PORT = "5432"
DATABASE_DATABASE = "spareparts"
DATABASE_USERNAME = "postgres"
DATABASE_PASSWORD = "4450"

IMPORT_FILE_NAME = "ProjectImport.xlsx"
IMPORT_FILE_SHEET_NAME = "NewProject"

DEBUG_MODE = True

# Functions
def dbquery(query):
    global db
    mycursor = db.cursor()
    mycursor.execute(query)
    result = mycursor.fetchall()
    mycursor.close()
    return result

# Exceptions
class ProjectAlreadyExists(Exception):
	pass
class FailedToInsertProject(Exception):
	pass

# Open File and Import Data
wb = openpyxl.load_workbook(IMPORT_FILE_NAME)
if DEBUG_MODE:
    print(wb.sheetnames)
sheet = wb[IMPORT_FILE_SHEET_NAME]
if DEBUG_MODE:
    print(sheet['A2'].value)

projectNumber = str(sheet['B2'].value)

billingSame = True
billingaddress = {}
if str(sheet['B17'].value) == "No":
    billingSame = False
    billingaddress = {
        "streetnumber": str(sheet['B20'].value),
        "streetname": str(sheet['B21'].value),
        "line2": str(sheet['B22'].value),
        "city": str(sheet['B23'].value),
        "prov": str(sheet['B24'].value),
        "provAbbr": str(sheet['B25'].value),
        "postalcode": str(sheet['B26'].value),
        "country": str(sheet['B27'].value),
        "countryAbbr": str(sheet['B28'].value)
    }
else:
    billingaddress = {
        "streetnumber": str(sheet['B8'].value),
        "streetname": str(sheet['B9'].value),
        "line2": str(sheet['B10'].value),
        "city": str(sheet['B11'].value),
        "prov": str(sheet['B12'].value),
        "provAbbr": str(sheet['B13'].value),
        "postalcode": str(sheet['B14'].value),
        "country": str(sheet['B15'].value),
        "countryAbbr": str(sheet['B16'].value)
    }

address = {
    "streetnumber": str(sheet['B8'].value),
    "streetname": str(sheet['B9'].value),
    "line2": str(sheet['B10'].value),
    "city": str(sheet['B11'].value),
    "prov": str(sheet['B12'].value),
    "provAbbr": str(sheet['B13'].value),
    "postalcode": str(sheet['B14'].value),
    "country": str(sheet['B15'].value),
    "countryAbbr": str(sheet['B16'].value),
    "sameBilling": billingSame,
}

details = {
    "name":str(sheet['B3'].value),
    "altnames":str(sheet['B4'].value).split(','),
    "description":str(sheet['B5'].value),
    "address": address,
    "billingaddress":billingaddress
}

print(IMPORT_FILE_NAME+' Successfully Parsed')
if DEBUG_MODE:
    print(details)

# Connect to DB
global db
db = psycopg2.connect(host=DATABASE_HOST, port=DATABASE_PORT, database=DATABASE_DATABASE, user=DATABASE_USERNAME, password=DATABASE_PASSWORD)

# Make Sure Project Doesn't Already Exist
checkQuery = "select id from project where projectnumber = '"+projectNumber+"';"
if DEBUG_MODE:
    print(checkQuery)
checkResult = dbquery(checkQuery)
if DEBUG_MODE:
    print(checkResult)
if len(checkResult) >= 1:
    print("Project "+projectNumber+" Already Exists. Exiting!")
    raise ProjectAlreadyExists
    


# Create Project
insertQuery = "insert into project (projectnumber, details) values ('"+projectNumber+"','"+json.dumps(details)+"') returning id;"
if DEBUG_MODE:
    print(insertQuery)
insertResult = dbquery(insertQuery)
if DEBUG_MODE:
    print(insertResult)
if len(insertResult) != 1:
    print("Project "+projectNumber+" Failed to Import. Exiting!")
    raise FailedToInsertProject
else:
    print("Project "+projectNumber+" Successfully Imported!")

# Close Database Connection
db.commit() #!!!!!!
db.close()