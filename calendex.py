from datetime import datetime
from datetime import timedelta
import json
import os
import time
import pytz
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from configparser import ConfigParser

SCOPES = ['https://www.googleapis.com/auth/calendar.events', 'https://www.googleapis.com/auth/calendar.readonly']
delta = 978307200 # Difference of seconds between Cocoa and Unix epoch
managedDescription = 'Managed by Calendex'


def cocoa_to_datetime(cocoa_time, timezone):
    return datetime.fromtimestamp(cocoa_time + delta, timezone)


def parse_widget(input_filename, calendar_id):
    event_dates = dict()
    with open(input_filename, "r") as input_file:
        input = json.load(input_file)
        timezone = pytz.timezone(input["calendar"]["timeZone"]["identifier"])
        curdate = datetime.fromtimestamp(0)
        for dayToAppointments in input["dayToAppointments"]:
            if isinstance(dayToAppointments, int):
                curdate = cocoa_to_datetime(dayToAppointments, timezone)
                event_dates[curdate] = []
                continue
            for event in dayToAppointments:
                if event["calendar"]["id"] != calendar_id:
                    continue
                event_dates[curdate] += [event]
    
    for date in event_dates:
        events = event_dates[date]
        for event in events:
            event["startTime"] = cocoa_to_datetime(event["startTime"], timezone)
            event["endTime"] = cocoa_to_datetime(event["endTime"], timezone)
    return event_dates, timezone


def authorize():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def get_events_in_date(date, service, timezone, destination_calendar_id):
    utcStartTime = date.astimezone(pytz.utc)
    endTime = date + timedelta(days=1,seconds=1)
    utcEndTime = endTime.astimezone(pytz.utc)
    eventsResult = service.events().list(
        calendarId=destination_calendar_id, 
        timeMin=utcStartTime.isoformat(),
        timeMax=utcEndTime.isoformat(),
        timeZone=timezone).execute()
    events = eventsResult.get('items', [])
    # Filter only events having a description field and having it set as "Managed by Calendex"
    events = [event for event in events if 'description' in event and event['description'] == managedDescription]
    # Filter out events having a start time before the date we are looking for
    events = [event for event in events if datetime.fromisoformat(event['start']['dateTime']) >= date]
    return events


def are_events_different(events, googleEvents):
    if len(events) != len(googleEvents):
        print(f"Number of events: {len(events)} != {len(googleEvents)}")
        return True
    for event in events:
        matched = False
        for googleEvent in googleEvents:
            # Add empty location if not present
            if 'location' not in googleEvent:
                googleEvent['location'] = ''
            if event['location'] != googleEvent['location']:
                continue
            if event['subject'] != googleEvent['summary']:
                continue
            if event['startTime'] != datetime.fromisoformat(googleEvent['start']['dateTime']):
                continue
            if event['endTime'] != datetime.fromisoformat(googleEvent['end']['dateTime']):
                continue
            matched = True
            break
        if not matched:
            print(f"Cound not match event: {event}")
            return True
    return False


def delete_events(events, service, destination_calendar_id):
    for event in events:
        print(f"Deleting event: {event['summary']}")
        service.events().delete(calendarId=destination_calendar_id, eventId=event['id']).execute()


def create_events(events, service, destination_calendar_id):
    for event in events:
        print(f"Creating event: {event['subject']}")
        startDate = event['startTime'].isoformat()
        endDate = event['endTime'].isoformat()
        google_event = {
            'summary': event['subject'],
            'location': event['location'],
            'description': managedDescription,
            'start': {
                'dateTime': startDate,
            },
            'end': {
                'dateTime': endDate,
            },
        }
        service.events().insert(calendarId=destination_calendar_id, body=google_event).execute()


def update_calendar(event_dates, timezone, credentials, destination_calendar_id):
    try:
        service = build('calendar', 'v3', credentials=credentials)

        for date in event_dates:
            events = event_dates[date]
            googleEvents = get_events_in_date(date, service, timezone, destination_calendar_id)
            if are_events_different(events, googleEvents):
                print(f"Events are different on {date}")
                delete_events(googleEvents, service, destination_calendar_id)
                create_events(events, service, destination_calendar_id)

    except HttpError as error:
        print('An error occurred: %s' % error)


def print_calendar_id(service):
    calendar_list = service.calendarList().list().execute()
    for calendar_list_entry in calendar_list['items']:
        print(calendar_list_entry['summary'], calendar_list_entry['id'])


def validate_config(userinfo, credentials):
    if 'input' not in userinfo:
        raise Exception("Missing input field in config.ini")
    if 'source_calendar_id' not in userinfo:
        raise Exception("Missing source_calendar_id field in config.ini")
    if 'destination_calendar_id' not in userinfo:
        print("Missing destination_calendar_id field in config.ini")
        print("Please select the calendar you want to use as destination:")
        print_calendar_id(build('calendar', 'v3', credentials=credentials))
        print("Copy the calendar id and paste it in the config.ini file")
        exit(1)


def main():
    config_object = ConfigParser()
    config_object.read("config.ini")

    credentials = authorize()
    userinfo = config_object["CALENDEX"]
    validate_config(userinfo, credentials)
    input_filename = userinfo["input"]
    source_calendar_id = userinfo["source_calendar_id"]
    destination_calendar_id = userinfo["destination_calendar_id"]
    while True:
        try:
            event_dates, timezone = parse_widget(input_filename, source_calendar_id)
            update_calendar(event_dates, timezone, credentials, destination_calendar_id)
        except Exception as message:
            print(f"Failed to convert widget to calendar: {message}")
        print("Waiting for next update...")
        time.sleep(30 * 60)


if __name__ == "__main__":
    main()