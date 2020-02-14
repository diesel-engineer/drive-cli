from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import sys
from datetime import datetime
from oauth2client import file, client, tools

# If modifying these scopes, delete the file token.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
dirpath = os.path.dirname(os.path.realpath(__file__))

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = sys.argv[1]
RANGE_NAME = sys.argv[2]
VARIABLES_COUNT = sys.argv[3]
SHEET_ID = sys.argv[4]

BUILD_STATUS = sys.argv[5]
BUILD_DURATION = sys.argv[6]
BUILD_VARIANT = sys.argv[7]
BUILD_LOCATION = sys.argv[8]

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    token = os.path.join(dirpath, 'token.json')
    store = file.Storage(token)
    creds = store.get()
    
    if not creds or creds.invalid:
        client_id = os.path.join(dirpath, 'oauth.json')
        flow = client.flow_from_clientsecrets(client_id, SCOPES)
        flags = tools.argparser.parse_args(args=[])
        creds = tools.run_flow(flow, store, flags)

    service = build('sheets', 'v4', credentials = creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    
    result = sheet.values().get(spreadsheetId = SPREADSHEET_ID,
                                range = VARIABLES_COUNT).execute()
    values = result.get('values', [])

    count = 0
    if values:
        count = int(values[0][0])
    
    date_time = datetime.now().strftime("%m/%d/%Y %H:%M").split(" ")
    date = date_time[0]
    time = date_time[1]
    
    body = {
        'values': [
            [
                date, 
                time, 
                BUILD_STATUS, 
                BUILD_DURATION, 
                '=CONVERT(D%s, "sec", "min")' % str(count + 2), 
                '=LOOKUP(D%s,{0,Variables!$B$4,Variables!$B$5,Variables!$B$6,Variables!$B$7},{"Fast","Normal","Slow","Too Slow"})&" Build"' % str(count + 2), 
                BUILD_VARIANT, 
                BUILD_LOCATION
            ]
        ]
    }
    
    result = sheet.values().append(spreadsheetId = SPREADSHEET_ID,
                               range = RANGE_NAME,
                               body = body,
                               valueInputOption = "USER_ENTERED",
                               insertDataOption="INSERT_ROWS").execute()
                               
if __name__ == '__main__':
    main()
