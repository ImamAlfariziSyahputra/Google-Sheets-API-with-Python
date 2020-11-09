import os
from Google import Create_Service
from pprint import pprint

FOLDER_PATH = r'<C:\Users\asus\Downloads>'
CLIENT_SECRET_FILE = os.path.join('Client_Secret.json')
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

spreadsheet_id = '1z1OtLpjSd7KJnnzTyjBPUs6x7Ei_ELmhH2B09wQlvjU'

sheet_id = '2094130777'

request_body = {
    'requests': [
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,

                    # row 3 to ro 11
                    'startRowIndex': 2,
                    'endRowIndex': 11,

                    # column E to column F
                    'startColumnIndex': 4,
                    'endColumnIndex': 6
                },
                'cell': {
                    'userEnteredFormat': {
                        'numberFormat': {
                            'type': 'CURRENCY',
                            'pattern': '$#,##0'
                        },
                        'backgroundColor': {
                            'red': 120,
                            'green': 60,
                            'blue': 70
                        },
                        'textFormat': {
                            'fontSize': 14,
                            'bold': True
                        }
                    }
                },
                'fields': 'userEnteredFormat(numberFormat, backgroundColor, textFormat)'
            }
        },
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 2,
                    'endRowIndex': 11,
                    'startColumnIndex': 3,
                    'endColumnIndex': 4
                },
                'cell': {
                    'userEnteredFormat': {
                        'numberFormat': {
                            'type': 'Number',
                            'pattern': '#,##0'
                        }
                    }
                },
                'fields': 'userEnteredFormat(numberFormat)'
            }
        }
    ]
}

response = service.spreadsheets().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body=request_body
).execute()