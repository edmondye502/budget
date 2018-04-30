
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import json
import datetime, time

# Setup the Sheets API
SCOPES_SHEETS = 'https://www.googleapis.com/auth/spreadsheets'
store_sheets = file.Storage('/home/scripts/budget_configs/credentials_sheets.json')
creds_sheets = store_sheets.get()
if not creds_sheets or creds_sheets.invalid:
    flow = client.flow_from_clientsecrets('/home/scripts/budget_configs/client_secret_sheets.json', SCOPES_SHEETS)
    creds_sheets = tools.run_flow(flow, store_sheets)
service_sheets = build('sheets', 'v4', http=creds_sheets.authorize(Http()))

with open('/home/scripts/budget_configs/config.json', 'r') as f:
    config = json.load(f)

SPREADSHEET_ID = config['EDMOND']['SPREADSHEET_ID']
RANGE_NAME = config['EDMOND']['RANGE_NAME']


# Setup the Gmail API
SCOPES_GMAIL = 'https://www.googleapis.com/auth/gmail.modify'
store_gmail = file.Storage('/home/scripts/budget_configs/credentials_gmail.json')
creds_gmail = store_gmail.get()
if not creds_gmail or creds_gmail.invalid:
    flow = client.flow_from_clientsecrets('/home/scripts/budget_configs/client_secret_gmail.json', SCOPES_GMAIL)
    creds_gmail = tools.run_flow(flow, store_gmail)
service_gmail = build('gmail', 'v1', http=creds_gmail.authorize(Http()))


def main():
    checkEmails()


def checkEmails():

    results = service_gmail.users().messages().list(userId='me', labelIds=['Label_1', 'UNREAD']).execute()
    messages = []
    if 'messages' in results:
      messages.extend(results['messages'])

    for msg in messages:
        msg_results = service_gmail.users().messages().get(userId='me', id=msg['id']).execute()

        newValues = setupBudget(msg_results)
        if newValues:
            addEntry(newValues)
        service_gmail.users().messages().modify(userId='me', id=msg['id'], body={'removeLabelIds': ['UNREAD']}).execute()


def setupBudget(msg_results):
    msg_date = datetime.datetime.fromtimestamp(float(msg_results['internalDate']) / 1000.0).strftime('%m/%d/%Y')
    budget_data = msg_results['snippet'].split()
    try:
        newValues = {'budget_date': msg_date, 'budget_type': budget_data[0], 'budget_amount': budget_data[1], 'budget_note': budget_data[2]}
        return newValues
    except IndexError:
        return False
    

def addEntry(newValues):

    values = [[newValues['budget_date'], newValues['budget_type'], newValues['budget_amount'], newValues['budget_note']]]
    body = {'values': values}

    #determine sheet
    budget_date = datetime.datetime.strptime(newValues['budget_date'], '%m/%d/%Y')
    budget_month = budget_date.strftime("%B")
    sheet =  budget_month + '!'

    result = service_sheets.spreadsheets().values().append(spreadsheetId=SPREADSHEET_ID, range=sheet+RANGE_NAME, valueInputOption='USER_ENTERED', body=body).execute()


if __name__ == "__main__":
    main()