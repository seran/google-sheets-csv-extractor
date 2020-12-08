from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import sys, getopt
import json, csv

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

SPREADSHEET_ID = ''
SHEET_NAME = 'Data'
LIMIT = 0
HAS_HEADER = False
OPERATION = None

def index_to_notation(index):
    string = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        string = chr(65 + remainder) + string
    return string

def extract():
    sheet = connect()

    range_values = []

    with open("metadata.json", "r") as metadata_file:
        options = json.load(metadata_file)
        
        for key, value in options.items():
            if value['enabled'] == True:
                range_string = SHEET_NAME + "!"+index_to_notation(value["index"]) + "1:" + index_to_notation(value["index"])
                if LIMIT > 0:
                    range_string += ":" + LIMIT
                range_values.append(range_string)

    result = sheet.values().batchGet(spreadsheetId=SPREADSHEET_ID, ranges=range_values, majorDimension="COLUMNS").execute()
    values = result.get("valueRanges", [])

    output = []

    if not values:
        print("No data found.")
    else:
        for index, row in enumerate(values):
            output.append(tuple(row['values'][0]))
    final = zip(*output)
    column_names = final[0]
    del final[0]
    if len(final) > 0:
        if HAS_HEADER == True:
            write_to_csv(final,column_names)
        else:
            write_to_csv(final)
    else:
        print("No data to write")

def write_to_csv(data,fields=None):
    with open("data.csv", "w") as csv_file:
        csvwiter = csv.writer(csv_file)

        if HAS_HEADER == True and fields:
            csvwiter.writerow(fields)

        csvwiter.writerows(data)
    print("CSV file successfully written.")

def create_metadata():
    sheet = connect()

    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
            range="A1:K1").execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        output = {}
        for index, cell in enumerate(values[0]):
            output[cell] = {
                "hash": False,
                "type": type(cell).__name__,
                "enabled": True,
                "description": "",
                "primary_key": False,
                "index": index + 1
            }
        with open("metadata.json", "w") as json_file:
            json.dump(output, json_file)
        print("Metadata file written successfully")

def connect():
    creds = None

    if not SPREADSHEET_ID:
        print("No spreadsheet id provided")
        sys.exit()

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    sheet = service.spreadsheets()

    return sheet

if __name__ == '__main__':
    try:
        opts, args = getopt.getopt(sys.argv[1:],"h:o:s:l:t",["operation=", "sheet=","limit=","title="])
    except getopt.GetoptError:
        print('connector.py -o <meta|extract> -s <spreadsheet id> -l <record limit> -t <has title>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h' or opt == 'h':
            print('connector.py -o <meta|extract> -s <spreadsheet id> -l <record limit> -t <has title>')
            sys.exit()
        elif opt in ("-s", "--sheet"):
            if not arg:
                print("No spreadsheet id provided.")
                sys.exit()
            else:
                SPREADSHEET_ID = arg
        elif opt in ("-l", "--limit"):
            if arg:
                LIMIT = arg
        elif opt in ("-t", "--title"):
            if arg:
                LIMIT = True
        elif opt in ("-o", "--operation"):
            if arg == 'meta':
                OPERATION = "metadata"
            elif arg == 'extract':
                OPERATION = "extract"
            else:
                print("No operation provided")

    if OPERATION == "metadata":
        create_metadata()
    elif OPERATION == "extract":
        extract()