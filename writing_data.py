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
mySpreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

"""
------------------------------
"""

worksheet_name = 'Sales North!'
cell_range_insert = 'B2'
values = (
    ('Col A', 'Col B', 'Col C', 'Col D'),
    ('Apple', 'Orange', 'Watermelon', 'Banana')
)
value_range_body = {
    'majorDimension': 'COLUMNS',
    'values': values
}

service.spreadsheets().values().update(
    spreadsheetId=spreadsheet_id,
    valueInputOption='USER_ENTERED',
    range=worksheet_name + cell_range_insert,
    body=value_range_body
).execute()

# service.spreadsheets().values().clear(
#     spreadsheetId=spreadsheet_id,
#     range='Sales North'
# ).execute()

"""
------------------------------
"""

worksheet_name = 'Sales North!'
cell_range_insert = 'B2'
values = (
    ('Col E', 'Col F', 'Col G', 'Col H'),
    ('Toyota', 'Honda', 'Tesla', 'BMW')
)
value_range_body = {
    'majorDimension': 'COLUMNS',
    'values': values
}

service.spreadsheets().values().append(
    spreadsheetId=spreadsheet_id,
    valueInputOption='USER_ENTERED',
    range=worksheet_name + cell_range_insert,
    body=value_range_body
).execute()

# service.spreadsheets().values().clear(
#     spreadsheetId=spreadsheet_id,
#     range='Sales North'
# ).execute()