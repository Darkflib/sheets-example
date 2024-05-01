import random
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle

# Scopes required by the script
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def authenticate_google_docs():
    """Authenticate and return the Google Sheets service object."""
    creds = None
    # Check if the token.pickle exists
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If no valid credentials available, log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return build('sheets', 'v4', credentials=creds)

def main():
    # Authenticate and create the service object
    service = authenticate_google_docs()
    # Create a new spreadsheet
    spreadsheet = {
        'properties': {
            'title': 'Python Test'
        }
    }
    spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                    fields='spreadsheetId').execute()
    print(f"Spreadsheet ID: {spreadsheet.get('spreadsheetId')}")

    # Generate random data
    data = [[random.randint(1, 100) for _ in range(3)] for _ in range(12)]
    # Adding the column headers
    data.insert(0, ['','',''])
    data.insert(1, ['x', 'y', 'z'])

    # Calculate min and max for bounding box
    data.append(['=MIN(A3:A14)', '=MIN(B3:B14)', '=MIN(C2:C14)'])
    data.append(['=MAX(A3:A14)', '=MAX(B3:B14)', '=MAX(C2:C14)'])

    # Batch update values
    body = {
        'values': data
    }
    range_name = 'A1:C16' # Range to update
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet.get('spreadsheetId'), range=range_name,
        valueInputOption='USER_ENTERED', body=body).execute()

    print(f"Updated cells: {result}")

    # Update font size of the title
    requests = [{
        "updateCells": {
            "rows": {
                "values": [{
                    "userEnteredValue": {
                        "stringValue": "Python Test"
                    },
                    "userEnteredFormat": {
                        "textFormat": {
                            "fontSize": 24
                        }
                    }
                }]
            },
            "fields": "userEnteredValue,userEnteredFormat.textFormat.fontSize",
            "range": {
                "sheetId": 0,
                "startRowIndex": 0,
                "endRowIndex": 1,
                "startColumnIndex": 0,
                "endColumnIndex": 1
            }
        }
    }]
    body = {
        'requests': requests
    }
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet.get('spreadsheetId'),
        body=body).execute()

    print(f"Updated cells response: {response}")

    # Update the format of the results rows a15 to c16
    requests = [{
        "repeatCell": {
            "range": {
                "sheetId": 0,
                "startRowIndex": 14,
                "endRowIndex": 16,
                "startColumnIndex": 0,
                "endColumnIndex": 3
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {
                        "red": 0.0,
                        "green": 0.0,
                        "blue": 1.0
                    },
                    "textFormat": {
                        "foregroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 1.0
                        },
                        "fontSize": 12
                    }
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }
    }]
    body = {
        'requests': requests
    }
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet.get('spreadsheetId'),
        body=body).execute()
    
    print(f"Updated cells response: {response}")
    


if __name__ == '__main__':
    main()
    # Apply additional formatting and summarizing functions via batchUpdate
#    requests += [
#        {
#            'repeatCell': {
#                'range': {
#                    'sheetId': 0,
#                    'startRowIndex': 0,
#                    'endRowIndex': 1,
#                    'startColumnIndex': 0,
#                    'endColumnIndex': 3
#                },
#                'cell': {
#                    'userEnteredFormat': {
#                        'horizontalAlignment': 'CENTER',
#                        'textFormat': {
#                            'fontSize': 16,
#                            'bold': True
#                        }
#                    }
#                },
#                'fields': 'userEnteredFormat(horizontalAlignment,textFormat)'
#            }
#        }
#    ]
#
#    body = {
#        'requests': requests
#    }
#    response = service.spreadsheets().batchUpdate(
#        spreadsheetId=spreadsheet['spreadsheetId'], body=body).execute()

    print('Sheet successfully created and formatted. Check your Google Sheets.')


