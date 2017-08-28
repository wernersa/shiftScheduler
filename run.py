# -*- coding: utf-8 -*-
"""
Created on Wed Jan 13 15:13:14 2016

@author: Werner
"""

# ----------------------
# --- IMPORTS
# ----------------------

import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime
import dateutil.parser
import collections
import pprint

import csv

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# ----------------------
# --- CONSTANTS
# ----------------------

SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'shiftScheduler python'

TZ = dateutil.tz.gettz('Europe/Oslo') 
WEEKDAY_NAMES = ["Monday", 
                 "Tuesday", 
                 "Wednesday", 
                 "Thursday", 
                 "Friday", 
                 "Saturday",
                 "Sunday"]
                 
SHIFT_START = 8
SHIFT_END   = 20

# ----------------------
# --- SECONDARY FUNCTIONS
# --- Helper functions, not directly associated with main program
# ----------------------

def flatten(l):
    for el in l:
        if isinstance(el, collections.Iterable) and not isinstance(el, basestring):
            for sub in flatten(el):
                yield sub
        else:
            yield el

def roundTime(dt=None, roundTo=60):
   """Round a datetime object to any time laps in seconds
   dt : datetime.datetime object, default now.
   roundTo : Closest number of seconds to round to, default 1 minute.
   Author: Thierry Husson 2012 - Use it as you want but don't blame me.
   """
   if dt == None : dt = datetime.datetime.now()
   seconds = (dt - dt.min.replace(tzinfo=dt.tzinfo)).seconds #Workaround "TypeError: can't subtract offset-naive and offset-aware datetimes"
   # // is a floor division, not a comment on following line:
   rounding = (seconds+roundTo/2) // roundTo * roundTo
   return dt + datetime.timedelta(0,rounding-seconds,-dt.microsecond)            

# ----------------------
# --- PRIMARY FUNCTIONS
# --- Primary functions, directly associated with main program
# ----------------------     

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
                                   'calendar-python.json')

    store = oauth2client.file.Storage(credential_path)
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


def get_calendars(service):
    """
    Get a list over all available Google API calendars.
    """

    calendar_list = {}
    i = 0

    page_token = None
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            i += 1
            print(i, calendar_list_entry['id'], calendar_list_entry['summary'])
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break

    return calendar_list


def get_events(service, selected_calendars):
    """
    Get a list over all available Google API events, for each calendar
    supplied by input.
    """
    events = []   
    
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    #TODO: Days should be a variable. Arbitrary 150 chosen for now.
    until = (datetime.datetime.utcnow() + datetime.timedelta(150)).isoformat() + 'Z'

    print("Calendars selected: ")
    for calendar in selected_calendars:
        print(calendar['summary'])
        
        events_Result = []
        page_token = ""
        
        while True:
            if not page_token:
                events_Result = service.events().list(
                    calendarId=calendar['id'], timeMin=now, timeMax=until,
                    maxResults=2500, singleEvents=True, orderBy='startTime').execute()
            else:
                #Handle cases where more than 2500 events are pending:
                events_Result = service.events().list(pageToken=page_token).execute()
            events.extend(events_Result.get('items', []))
            page_token = events_Result.get('nextPageToken')
            if not page_token:
                break

    if not events:
        print('No upcoming events found.')
        return False
    
    events.sort(key=lambda x: x['start'].get('dateTime', x['start'].get('date')))
    
    purge = []
    for i, event in enumerate(events):
        #TODO: Exclusion criterias
        if event['summary'] == 'Jobb PSYK':
            purge.append(i)

        try:
            if event['transparency'] == 'transparent':
                purge.append(i)
                continue
        except:
            pass
        
        if 'date' not in event['start']:
            
            event['start'] = roundTime(dateutil.parser.parse(event['start'].get('dateTime'), ignoretz=False),roundTo=60*60).astimezone(TZ)
            event['end'] = roundTime(dateutil.parser.parse(event['end'].get('dateTime'), ignoretz=False),roundTo=60*60).astimezone(TZ)
            #TODO: If the event starts and end on different days, handle it in another way
        else:      
            event['start'] = dateutil.parser.parse(event['start'].get('date'))
            event['end'] = dateutil.parser.parse(event['end'].get('date'))        
        
        event['weekday'] = event['start'].weekday()
    
    for i in reversed(purge):
        del events[i]
        
    events.sort(key=lambda x: x['weekday'])
    
    return events


def get_hourly_penalty(service, events):
    """
    Calculate how many times each weekly shift hour is scheduled for an event,
    standarize as percentage of total hours for future comparison.
    Returns:
        A list of weekdays list of hours penalty.
    """
    weekdays = [[] for x in range(7)]
    
    for event in events:
        i = 0
        #Round to nearest whole hour. TODO: Make it possible to decide on this
        ticker = event['start']
        
        while ticker < event['end']:
            weekdays[event['weekday']].append(ticker.hour)
            i += 1
            ticker = (event['start'] + datetime.timedelta(hours=i))
    

    print(weekdays)
    weekdays_Counter = []
    hours = range(SHIFT_START,SHIFT_END+1) #The range of dayly hours we are interested in
    largest_count = 0
    
    for i, day in enumerate(weekdays):
        day_Counter = []
        
        for x in hours:
            day_Counter.append(day.count(x))
        
        weekdays_Counter.append(day_Counter)
    
    #If individual schedules should be relatively weighted. This could be considered to be most utilitaristic as schedules with a high baseline hour count would be much worse off otherwise.
    largest_count = max([num for elem in weekdays_Counter for num in elem])
    weekdays_weigthed = [list(map(lambda y: round(y/float(largest_count)*100,2), x)) for x in weekdays_Counter]
    
    return weekdays_weigthed, largest_count



def get_shift_windows(service, hourly_penalty):
    """
    Temporary function for future integration of flexible shift start and end
    times.
    
    Returns:
        List:
           + Weekday n
               + Permutation set n
                   + Window 1/n
                       + {'penalty','hours'}
    """
    
    #TODO: for shift_windows in shift_partitions
    
    #As of now, it returns three possible shift setups:
    return [[[{'penalty':0,'hours':[8,14]},{'penalty':0,'hours':[14,20]},{'penalty':0,'hours':[8,18]}]] for x in range(7)]



def main():
    """Simple scheduler for finding office hours least occupied by other events

    Creates a Google Calendar API service object, user selects which of the
    returned calendars should be used. Successive weekday hours with the
    highest rate of unavailability due to other events is penalized
    accordingly.
    
    See function get_shift_windows for changing the timespan of shifts.    
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)
        
    #Simulate persons:    
    #for person in persons:    
    
    available_calendars = get_calendars(service)
    
    #Select calendars with input separated by commas
    selected_i = None
    while not selected_i:
        try:    
            selected_i = input("Please select the calendar numbers listed above, separated by commas\n").split(",")
        except:
            print("Invalid input.")
    #selected_i = [2, 3, 4, 6, 10]
    
    selected_calendars = [available_calendars['items'][int(x)-1] for x in selected_i]
    print(selected_calendars)

    events = get_events(service, selected_calendars)
    
    penalty, largest_hour_count = get_hourly_penalty(service, events) 
    
    shift_windows = get_shift_windows(service, penalty)
    
    for i_day, day in enumerate(shift_windows):
        for permutation_set in day:
            for shift in permutation_set:
                ticker = shift['hours'][0]
                shift_duration = shift['hours'][1] - shift['hours'][0]
                while ticker < shift['hours'][1]:
                    shift['penalty'] += penalty[i_day][ticker-SHIFT_START]
                    ticker += 1
                #Average penalty per hour window -> total hours through semestre unavailable
                shift['penalty'] = shift['penalty'] / shift_duration
                
                #Is it interesting to see total hours "lost" due to unavailability with the current shift?
                #shift['penalty'] = shift['penalty'] * largest_hour_count / 100
    
    print("Shift with lowest penalty is preferable:")
    pprint.pprint(shift_windows)
    
    with open("output.csv", "w", newline='') as f:
        #Zip transposes. Columns is days. Rows is shift-hours.
        print(penalty)
        export = list(zip(*[[x for x in range(SHIFT_START,SHIFT_END+1)], *penalty]))
        export.insert(0, ["", *WEEKDAY_NAMES])
        
        #Export column-named csv, sep=",", decimal="."
        writer = csv.writer(f, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        writer.writerows(export) 
    
if __name__ == '__main__':
    main()