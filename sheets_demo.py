import os
from Google import Create_Service
from pprint import pprint

FOLDER_PATH = r'<C:\Users\asus\Downloads>'
CLIENT_SECRET_FILE = os.path.join('Client_Secret.json')
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

"""
Blank Spreadsheet File
"""
"""
dict_keys(['spreadsheetId','properties','sheets','spreadsheetUrl'])
"""
sheets_file1 = service.spreadsheets().create().execute()


pprint(sheets_file1['spreadsheetUrl'])
pprint(sheets_file1['sheets'])
pprint(sheets_file1['spreadsheetId'])
pprint(sheets_file1['properties'])

"""
Advanced Example: Spreadsheet File with some default settings
"""

sheet_body = {
    'properties': {
        'title': 'My First Google Sheet File',
        'locale': 'en_US',
        'timeZone': 'America/Los_Angeles',
        'autoRecalc': 'HOUR'
    },
    'sheets': [
        {
            'properties': {
                'title': 'Sales North'
            }
        },
        {
            'properties': {
                'title': 'Sales East'
            }
        },
        {
            'properties': {
                'title': 'Sales West'
            }
        },
        {
            'properties': {
                'title': 'Sales South'
            }
        },
        {
            'properties': {
                'title': 'ChartData'
            }
        },
    ]
}

sheets_file2 = service.spreadsheets().create(
    body = sheet_body
).execute()

print(sheets_file2)

pprint(sheets_file2['spreadsheetUrl'])
pprint(sheets_file2['sheets'])
pprint(sheets_file2['spreadsheetId'])
pprint(sheets_file2['properties'])


