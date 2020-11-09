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
sheet_id = '2003463487'

"""
Example 1. Single Bar Chart (New Worksheet)
"""

# request_body = {
#     'requests': [
#         {
#             'addChart': {
#                 'chart': {
#                     'spec': {
#                         'title': 'Regional Sales Summary',
#                         'basicChart': {
#                             'chartType': 'COLUMN',
#                             'legendPosition': 'BOTTOM_LEGEND',
#                             'axis': [
#                                 # X-AXIS
#                                 {
#                                     'position': 'BOTTOM_AXIS',
#                                     'title': 'Sales Regions'
#                                 },
#                                 # Y-AXIS
#                                 {
#                                     'position': 'LEFT_AXIS',
#                                     'title': 'Total Revenue'
#                                 }
#                             ],
#
#                             # chart labels
#                             'domains': [
#                                 {
#                                     'domain': {
#                                         'sourceRange': {
#                                             'sources': [
#                                                 {
#                                                     'sheetId': sheet_id,
#                                                     'startRowIndex': 0,
#                                                     'endRowIndex': 11,
#                                                     'startColumnIndex': 0,
#                                                     'endColumnIndex': 1
#                                                 }
#                                             ]
#                                         }
#                                     }
#                                 }
#                             ],
#
#                             'series': [
#                                 {
#                                     'series': {
#                                         'sourceRange': {
#                                             'sources': [
#                                                 {
#                                                     'sheetId': sheet_id,
#                                                     'startRowIndex': 0,  # row number 1
#                                                     'endRowIndex': 11,  # row number 10
#                                                     'startColumnIndex': 1,  # column B
#                                                     'endColumnIndex': 2
#                                                 }
#                                             ]
#                                         }
#                                     },
#                                     'targetAxis': 'LEFT_AXIS'
#                                 }
#                             ]
#                         }
#                     },
#                     'position': {
#                         'newSheet': True
#                     }
#                 }
#             }
#         }
#     ]
# }
#
#
#
# response = service.spreadsheets().batchUpdate(
#     spreadsheetId=spreadsheet_id,
#     body=request_body
# ).execute()

"""
Example 1.b Single Bar Chart (Existing Worksheet)
"""

# request_body = {
#     'requests': [
#         {
#             'addChart': {
#                 'chart': {
#                     'spec': {
#                         'title': 'Regional Sales Summary',
#                         'basicChart': {
#                             'chartType': 'COLUMN',
#                             'legendPosition': 'BOTTOM_LEGEND',
#                             'axis': [
#                                 # X-AXIS
#                                 {
#                                     'position': 'BOTTOM_AXIS',
#                                     'title': 'Sales Regions'
#                                 },
#                                 # Y-AXIS
#                                 {
#                                     'position': 'LEFT_AXIS',
#                                     'title': 'Total Revenue'
#                                 }
#                             ],
#
#                             # chart labels
#                             'domains': [
#                                 {
#                                     'domain': {
#                                         'sourceRange': {
#                                             'sources': [
#                                                 {
#                                                     'sheetId': sheet_id,
#                                                     'startRowIndex': 0,
#                                                     'endRowIndex': 11,
#                                                     'startColumnIndex': 0,
#                                                     'endColumnIndex': 1
#                                                 }
#                                             ]
#                                         }
#                                     }
#                                 }
#                             ],
#
#                             'series': [
#                                 {
#                                     'series': {
#                                         'sourceRange': {
#                                             'sources': [
#                                                 {
#                                                     'sheetId': sheet_id,
#                                                     'startRowIndex': 0,  # row number 1
#                                                     'endRowIndex': 11,  # row number 10
#                                                     'startColumnIndex': 1,  # column B
#                                                     'endColumnIndex': 2
#                                                 }
#                                             ]
#                                         }
#                                     },
#                                     'targetAxis': 'LEFT_AXIS'
#                                 }
#                             ]
#                         }
#                     },
#                     'position': {
#                         'overlayPosition': {
#                             'anchorCell': {
#                                 'sheetId': '165129500',
#                                 'rowIndex': 1,
#                                 'columnIndex': 1
#                             },
#                             'offsetXPixels': 0,
#                             'offsetYPixels': 0,
#                             'widthPixels': 800,
#                             'heightPixels': 600
#                         }
#                     }
#                 }
#             }
#         }
#     ]
# }
#
#
#
# response = service.spreadsheets().batchUpdate(
#     spreadsheetId=spreadsheet_id,
#     body=request_body
# ).execute()

"""
Example 2 Group Data together (Single Chart)
"""

# request_body = {
#     'requests': [
#         {
#             'addChart': {
#                 'chart': {
#                     'spec': {
#                         'title': 'Regional Sales Summary',
#                         'basicChart': {
#                             'chartType': 'COLUMN',
#                             'legendPosition': 'BOTTOM_LEGEND',
#                             'axis': [
#                                 # X-AXIS
#                                 {
#                                     'position': 'BOTTOM_AXIS',
#                                     'title': 'Sales Regions'
#                                 },
#                                 # Y-AXIS
#                                 {
#                                     'position': 'LEFT_AXIS',
#                                     'title': 'Total Revenue'
#                                 }
#                             ],
#
#                             # chart labels
#                             'domains': [
#                                 {
#                                     'domain': {
#                                         'sourceRange': {
#                                             'sources': [
#                                                 {
#                                                     'sheetId': sheet_id,
#                                                     'startRowIndex': 0,
#                                                     'endRowIndex': 11,
#                                                     'startColumnIndex': 0,
#                                                     'endColumnIndex': 1
#                                                 }
#                                             ]
#                                         }
#                                     }
#                                 }
#                             ],
#
#                             # chart data
#                             'series': [
#                                 # North Region
#                                 {
#                                     'series': {
#                                         'sourceRange': {
#                                             'sources': [
#                                                 {
#                                                     'sheetId': sheet_id,
#                                                     'startRowIndex': 0,  # row number 1
#                                                     'endRowIndex': 11,  # row number 10
#                                                     'startColumnIndex': 1,  # column B
#                                                     'endColumnIndex': 2
#                                                 }
#                                             ]
#                                         }
#                                     },
#                                     'targetAxis': 'LEFT_AXIS'
#                                 },
#                                 # East
#                                 {
#                                     'series': {
#                                         'sourceRange': {
#                                             'sources': [
#                                                 {
#                                                     'sheetId': sheet_id,
#                                                     'startRowIndex': 0,  # row number 1
#                                                     'endRowIndex': 11,  # row number 10
#                                                     'startColumnIndex': 2,  # column C
#                                                     'endColumnIndex': 3
#                                                 }
#                                             ]
#                                         }
#                                     },
#                                     'targetAxis': 'LEFT_AXIS'
#                                 },
#                                 # West
#                                 {
#                                     'series': {
#                                         'sourceRange': {
#                                             'sources': [
#                                                 {
#                                                     'sheetId': sheet_id,
#                                                     'startRowIndex': 0,  # row number 1
#                                                     'endRowIndex': 11,  # row number 10
#                                                     'startColumnIndex': 3,  # column D
#                                                     'endColumnIndex': 4
#                                                 }
#                                             ]
#                                         }
#                                     },
#                                     'targetAxis': 'LEFT_AXIS'
#                                 },
#                                 # South
#                                 {
#                                     'series': {
#                                         'sourceRange': {
#                                             'sources': [
#                                                 {
#                                                     'sheetId': sheet_id,
#                                                     'startRowIndex': 0,  # row number 1
#                                                     'endRowIndex': 11,  # row number 10
#                                                     'startColumnIndex': 4,  # column E
#                                                     'endColumnIndex': 5
#                                                 }
#                                             ]
#                                         }
#                                     },
#                                     'targetAxis': 'LEFT_AXIS'
#                                 }
#                             ]
#                         }
#                     },
#                     'position': {
#                         'newSheet': True
#                     }
#                 }
#             }
#         }
#     ]
# }
#
#
#
# response = service.spreadsheets().batchUpdate(
#     spreadsheetId=spreadsheet_id,
#     body=request_body
# ).execute()

"""
Example 3. Graph Horizontal Bar Chart
"""

request_body = {
    'requests': [
        {
            'addChart': {
                'chart': {
                    'spec': {
                        'title': 'Regional Sales Summary',
                        'basicChart': {
                            'chartType': 'BAR',
                            'legendPosition': 'BOTTOM_LEGEND',
                            'axis': [
                                # X-AXIS
                                {
                                    'position': 'BOTTOM_AXIS',
                                    'title': 'Sales Regions'
                                },
                                # Y-AXIS
                                {
                                    'position': 'LEFT_AXIS',
                                    'title': 'Total Revenue'
                                }
                            ],

                            # chart labels
                            'domains': [
                                {
                                    'domain': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 0,
                                                    'endRowIndex': 11,
                                                    'startColumnIndex': 0,
                                                    'endColumnIndex': 1
                                                }
                                            ]
                                        }
                                    }
                                }
                            ],

                            # chart data
                            'series': [
                                # North Region
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 0,  # row number 1
                                                    'endRowIndex': 11,  # row number 10
                                                    'startColumnIndex': 1,  # column B
                                                    'endColumnIndex': 2
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'BOTTOM_AXIS'
                                },
                                # East
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 0,  # row number 1
                                                    'endRowIndex': 11,  # row number 10
                                                    'startColumnIndex': 2,  # column C
                                                    'endColumnIndex': 3
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'BOTTOM_AXIS'
                                },
                                # West
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 0,  # row number 1
                                                    'endRowIndex': 11,  # row number 10
                                                    'startColumnIndex': 3,  # column D
                                                    'endColumnIndex': 4
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'BOTTOM_AXIS'
                                },
                                # South
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 0,  # row number 1
                                                    'endRowIndex': 11,  # row number 10
                                                    'startColumnIndex': 4,  # column E
                                                    'endColumnIndex': 5
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'BOTTOM_AXIS'
                                }
                            ]
                        }
                    },
                    'position': {
                        'newSheet': True
                    }
                }
            }
        }
    ]
}



response = service.spreadsheets().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body=request_body
).execute()

print('chart created succesfully')
