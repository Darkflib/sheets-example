# Read the google sheets spreadsheet ID from the sheet.json file
# for each sheet in the spreadsheet, download the sheet as a CSV file
# and save it in the data/ folder

import random
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import json
import time
import logging
from datetime import datetime

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# setup logging to console with timestamp
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console.setFormatter(formatter)
logger.addHandler(console)


def get_credentials():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        # if creds and creds.expired and creds.refresh_token:
        #     creds.refresh(Request())
        # else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def get_sheet_id():
    # {"spreadsheetId": "1EjOtwFQGTBHV4CZIcZQ6VcFkEcxCGYKcxlXjqrmY5Gk"}
    with open('sheet.json', 'r') as f:
        sheet = json.load(f)
    return sheet['spreadsheetId']


def get_sheet_names(service, spreadsheet_id):
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    sheet_names = [sheet['properties']['title'] for sheet in sheets]
    return sheet_names


def rename_sheet(service, spreadsheet_id, sheet_id, new_title):
    body = {
        'requests': [
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': sheet_id,
                        'title': new_title
                    },
                    'fields': 'title'
                }
            }
        ]
    }
    request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body)
    response = request.execute()
    return response


def download_sheet(service, spreadsheet_id, sheet_name):
    range_name = f'{sheet_name}!A1:ZZZ'
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    if not values:
        logger.info(f'No data found in {sheet_name}')
    else:
        filename = f'data/{sheet_name}.csv'
        with open(filename, 'w') as f:
            for row in values:
                f.write(','.join(row) + '\n')
        logger.info(f'{sheet_name} saved to {filename}')
    return


def main():
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    spreadsheet_id = get_sheet_id()
    print(f"Spreadsheet ID: {spreadsheet_id}")
    sheet_names = get_sheet_names(service, spreadsheet_id)
    print(sheet_names)
    #for sheet_name in sheet_names:
    #    download_sheet(service, spreadsheet_id, sheet_name)
    #return
    rename_sheet(service, spreadsheet_id, 0, 'Mikes Awesome Sheet')


if __name__ == '__main__':
    main()
    


