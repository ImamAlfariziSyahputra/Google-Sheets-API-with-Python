import os
from Google import Create_Service
from pprint import pprint
import pandas as pd

FOLDER_PATH = r'<C:\Users\asus\Downloads>'
CLIENT_SECRET_FILE = os.path.join('Client_Secret.json')
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

spreadsheet_id = '1z1OtLpjSd7KJnnzTyjBPUs6x7Ei_ELmhH2B09wQlvjU'

"""
Example 1. Get method (single cell range of values)
"""

# response = service.spreadsheets().values().get(
#     spreadsheetId=spreadsheet_id,
#     majorDimension='ROWS',
#     range='Sales North!B2:F5'
# ).execute()
#
# pprint(response)
# # pprint(response.keys())
# # pprint(response['range'])
# # pprint(response['majorDimension'])
# # pprint(response['values'])
#
# columns = response['values'][0]
# data = response['values'][1:]
# df = pd.DataFrame(data, columns=columns)
#
# pprint(df)
#
# df2 = df.set_index('Invoice')
# pprint(df2)

"""
Example 1. Get method 
"""

response = service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id,
    majorDimension='ROWS',
    range='Sales North'
).execute()

pprint(response['values'])

columns = response['values'][1][1:]
data = [item[1:]for item in response['values'][2:]]
df = pd.DataFrame(data, columns=columns)
pprint(df)


