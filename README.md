# Microsoft calendar widget to google calendar

## Setup

1. Clone this repository
2. Follow the guide here to get a google calendar API key: https://developers.google.com/calendar/api/quickstart/python#set_up_your_environment
3. Copy the `credentials.json` file to the root of this repository
4. (Optional) Create a virtual environment: `python3 -m venv venv`
5. Install dependencies: `pip install -r requirements.txt`
6. Create a configuration file: `cp config.example.ini config.ini` and fill in the values, see the section below for more information
7. Run the script: `python3 main.py`

## Configuration

The following values are required in the configuration file:
 - `source_calendar_id`: The ID of the widget calendar to take events from
 - `destination_calendar_id`: The ID of the google calendar to push events to. If unknown, run the script once and the ID will be printed in the console
 - `input`: The path to the widget calendar store file

**Note on the source calendar ID**: The identifier can be found looking into the events of the store.json file. The identifier is the value of the `calendar.id` key in the event object.

### Potential input sources

The widget calendar store file can be found in one of the following locations, depending on the version of macOS and the version of Microsoft Office:

- `/Users/{username}/Library/Group Containers/{identifier}.Office/Library/Application Support/Calendar Widget/store.json`
- `/System/Volumes/Data/Users/{username}/Library/Group Containers/{identifier}.Office/Library/Application Support/Calendar Widget/store.json`

## Development notes

### Time format

Time in the widget is expressed in apple cocoa time format, which is the number of seconds since 1 January 2001, GMT.