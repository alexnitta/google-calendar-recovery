
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import datetime
import unicodecsv as csv

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """Exports a CSV file of a GET request to the Google Calendar API.

    To authorize this script, follow Step 1 from this page:
    https://developers.google.com/google-apps/calendar/quickstart/python

    This code is modified from the sample code in the Quickstart guide.

    See eventsResult for the details of the query.

    The CSV file contains columns which can be re-imported into Google Calendar per this page:
    https://support.google.com/calendar/answer/37118?hl=en

    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    calendar_id = raw_input("Enter the calendar id, or 'primary' for your default calendar: ")

    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    print('Getting the upcoming 1000 events')
    eventsResult = service.events().list(
        calendarId=calendar_id,
        timeMin=now,
        maxResults=1000,
        singleEvents=True,
        showDeleted=True,
        orderBy='updated').execute()
    events = eventsResult.get('items', [])

    if not events:
        print('No upcoming events found.')


    with open('output.csv', 'wb') as csvfile:

        headers = [
            "Subject",
            "Start Date",
            "End Date",
            "All Day Event",
            "Description",
            "Private",
            "NON-IMPORT ROWS ->",
            "Updated",
            "ID"]

        writer = csv.writer(csvfile, delimiter=',', encoding='utf-8')
        writer.writerow(headers)

        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            end = event['end'].get('dateTime', event['end'].get('date'))

            writer.writerow([
                event['summary'],
                start,
                end,
                "True",
                "Re-imported 10/28/16 - status: " + event['status'],
                "False",
                "",
                event['updated'],
                event['id']
                ])

            print("id: ", event['id'])
            print("start: ", start)
            print("end: ", end)
            print("summary: ", event['summary'])
            print("updated: ", event['updated'])
            print("status: ", event['status'])
            print("-----------")

if __name__ == '__main__':
    main()
