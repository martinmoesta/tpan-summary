from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import sys
import datetime

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'


def main():

    #print out TPAN summary (formatted for slack)
    def print_summary(row):   
        print("\n")
        print("*",row[0], " TPAN Summary*")
        print("```Created: ", datetime.datetime.now())
        print("Tanium Version: ", ts_version)
        print("Number of clients: ", client_num)
        print("Active Clients Estimate: ", active_client_num)
        print("Leader Ratio: ", leader_percent, "%")
        print("\n")
        print("Question Performance:")
        print("50%: ", fifty_t, " seconds")
        print("90%: ", ninety_t, " seconds")
        print(tail_p,"% (tail): ", tail_t, " seconds")
        print("\n")
        print("Findings:")
        print(findings)
        print("\n")
        print("Additional findings that you may want to examine:")
        print("\n")
        print("```")

    if len(sys.argv) != 2:
        print("Please enter a customer name via the command line.")
        return 2;

    #determines which row of the TPAN master sheet to scrape
    customer = sys.argv[1]

    #Creds for google sheet connection
    store = file.Storage('../creds/token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('../creds/credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))

    # Find and scrape the spreadsheet
    SPREADSHEET_ID = '1O7opJmfbTCrI8jKh2_obyxKYXZ9SfqGFh03qh1g3sX4'
    RANGE_NAME = 'Broken out Data!A2:Z100'
    result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,
                                                range=RANGE_NAME).execute()
    values = result.get('values', [])

    #Assign cells to summary values
    good_row = []
    for row in values:
        if customer.lower() in str(row[0]).lower():
            #print("Found the customer!")
            good_row = row
            ts_version = good_row[5]
            client_num = good_row[6]
            active_client_num = good_row[7]
            leader_percent = good_row[8]
            fifty_t = good_row[14]
            ninety_t = good_row[15]
            tail_t = good_row[16]
            tail_p = good_row[17]
            findings = good_row[18]
            #print(good_row)
            break;
    
    if not good_row:
        print("Could not find customer ", customer)
        return 1

    print_summary(good_row)
    return 0

if __name__ == '__main__':
    main()

