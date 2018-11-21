from __future__ import print_function
import re
import requests
import string
import sys, getopt
from datetime import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from lxml import html
from lxml.etree import tostring

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/calendar'

#payload static data
ACCOUNT = "jaycargroup"
USERNAME = ""
PASSWORD = ""
SIGNIN = "signin"

#program global variables
TIMEZONE = ""
DATE = ""
MANUALSHIFT = False
LOGIN_URL = "https://jaycargroup.workplaceonline.com.au/a/"
URL = "https://jaycargroup.workplaceonline.com.au/wb2/timesheet/"
URL2PART1 = "https://jaycargroup.workplaceonline.com.au/wb2/timesheet/,"
URL2PART2 = ",,,,,,1,"


YEAR = str(datetime.now().year)

def main(argv):
    global MANUALSHIFT
    global DATE
    global TIMEZONE
    try:
        opts, args = getopt.getopt(argv,"hu:p:d:",["username=","password=", "date="])
    except getopt.GetoptError:
        print('login_timesheet.py -u <username> -p <password> [-d <date>]')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print("Jaycar Timesheet to Google Calendar Converter")
            print("Usage:")
            print('\tlogin_timesheet.py -u <username> -p <password> [-d <date>]\n')
            print("-u <username> | your staff code or other login username")
            print("-p <password> | your plaintext password used to log in")
            print("-d <date>     | takes the format YYYY-MM-DD and returns the week that the date appears in.")
            print("\nIf no date is specified, the current week will be retrieved.")
            sys.exit()
        elif opt in ("-u", "--username"):
            USERNAME = arg
        elif opt in ("-p", "--password"):
            PASSWORD = arg
        elif opt in ("-d", "--date"):
            DATE = arg
    print("Starting")

    session_requests = requests.session()

    # Create payload
    payload = {
        "task" : SIGNIN,
        "account" : ACCOUNT,
        "username": USERNAME, 
        "password": PASSWORD 
    }

    # Perform login
    result = session_requests.post(LOGIN_URL, data = payload)
    
    # Scrape url
    if DATE is None:
        result = session_requests.get(URL, headers = dict(referer = URL))
    else:
        result = session_requests.get(URL2PART1 + DATE + URL2PART2, headers = dict(referer = URL))

    tree = html.fromstring(result.content)
    shift_dates = tree.xpath("//tr[contains(@class, 'tc_nap') or contains(@class, 'tc_app')]/td[position()=1]")
    shift_division = tree.xpath("//tr[contains(@class, 'tc_nap') or contains(@class, 'tc_app')]/td[position()=2]")
    shift_time_start = tree.xpath("//tr[contains(@class, 'tc_nap') or contains(@class, 'tc_app')]/td[position()=5]")
    shift_time_end = tree.xpath("//tr[contains(@class, 'tc_nap') or contains(@class, 'tc_app')]/td[position()=6]")
    shift_pay_start = tree.xpath("//tr[contains(@class, 'tc_nap') or contains(@class, 'tc_app')]/td[position()=9]")
    shift_pay_end = tree.xpath("//tr[contains(@class, 'tc_nap') or contains(@class, 'tc_app')]/td[position()=10]")
    
    print("Gathered shifts")

    ''' Get Google Calendar API Token '''
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))

    calendar = service.calendars().get(calendarId='primary').execute()


    TIMEZONE = calendar['timeZone']
    '''
    Check for shifts that have no Roster Start or Roster End times.
    if a shift has a Pay Start and Pay End, then it can be added to a calendar for posterity
    '''
    for dates, division, start, end, pay_start, pay_end in zip(shift_dates, shift_division, shift_time_start,shift_time_end, shift_pay_start, shift_pay_end):
        if (start.text is None) or (end.text is None):
            MANUALSHIFT = True
        if MANUALSHIFT is not False:
            print("Manual shift found:")
            print(dates.text)
            print(division.text)
            if (pay_start.text is not None) and (pay_end.text is not None):
                print("Shift has passed, I'll add it anyway")
                start.text = pay_start.text
                end.text = pay_end.text
            else:
                print("Shift has no pay start and end, Ignoring.")
                continue
        print(dates.text + " " + division.text + " " + start.text + " " + end.text)
        start.text = start.text.replace(":"," ") # convert : to whitespace, remove strftime reserved characters
        end.text = end.text.replace(":"," ") # convert : to whitespace, remove strftime reserved characters
        splitDate = dates.text.split(" ") #split date
        splitDate = [x for x in splitDate if x != ''] # remove whitespace / doublespace from some entries
        splitDate = [index.replace(',', '') for index in splitDate] # remove commas
        splitDate[1] = re.sub("[^0-9]", "", splitDate[1]); # change 1st to 1, 2nd to 2, etc.
        dates.text = " ".join(splitDate); # merge it all together
        # format the shift details into a complete date object
        shift_start_date_object = datetime.strptime(YEAR + " " + dates.text + " " + start.text, '%Y %A %d %B %H %M')
        shift_start_date_object = shift_start_date_object.strftime("%Y-%m-%dT%H:%M:%S")
        shift_end_date_object = datetime.strptime(YEAR + " " + dates.text + " " + end.text, '%Y %A %d %B %H %M')
        shift_end_date_object = shift_end_date_object.strftime("%Y-%m-%dT%H:%M:%S")
        ''' strptime object hates the ":" symbol '''
        start.text = start.text.replace(" ",":")
        end.text = end.text.replace(" ",":")
        ''' create the json scaffold for a calendar event '''
        event = {
            'summary': division.text + ": " + start.text + " - " + end.text ,
            'location': division.text,
            'description': 'Work. get your ass up now.',
            'start': {
                'dateTime': str(shift_start_date_object),
                'timeZone': TIMEZONE,
            },
            'end': {
                'dateTime': str(shift_end_date_object),
                'timeZone': TIMEZONE,
            },
            'reminders': {
                'useDefault': False,
                'overrides': [
                    #{'method': 'email', 'minutes': 24 * 60},
                    {'method': 'popup', 'minutes': 2 * 60},
                    ],
                },
            }
        print("Shift Created on the " + dates.text + " at " + division.text + "\n")
        event = service.events().insert(calendarId='primary', body=event).execute()
        MANUALSHIFT = False

if __name__ == "__main__":
    main(sys.argv[1:])