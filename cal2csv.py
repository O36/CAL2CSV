import sys
import os.path
from icalendar import Calendar
import calendar
import csv
from datetime import datetime, timedelta, date

# Check for verbose flag
verbose = '-v' in sys.argv
if verbose:
    sys.argv.remove('-v') # Remove it so it doesn't interfere with other argument parsing

filename = sys.argv[1]
filename_noext = filename[:-4]
file_extension = str(sys.argv[1])[-3:]
headers = ('Week', 'Summary', 'Start Time', 'End Time', 'Hours', 'Location', 'Description')

if len(sys.argv) < 3:
    print("No month provided, proceeding with current month")
    month = date.today().month
else:
    month = sys.argv[2]

class CalendarEvent:
    """Calendar event class"""
    summary = ''
    start = ''
    end = ''
    hours = ''
    description = ''
    location = ''

    def __init__(self, name):
        self.name = name

weeks = []
events = []


def open_cal():
    if os.path.isfile(filename):
        if file_extension == 'ics':
            print("Extracting events from file:", filename, "\n")
            f = open(sys.argv[1], 'rb')
            gcal = Calendar.from_ical(f.read())
            for component in gcal.walk():
                event = CalendarEvent("event")
                if component.get('STATUS') != "CONFIRMED":
                    continue
                if component.get('TRANSP') == 'TRANSPARENT' or component.get('TRANSP') == None:
                    continue #skip event that have not been accepted
                if component.get('SUMMARY') == None: continue #skip blank items
                if verbose:
                    print(f"Processing: {component.get('SUMMARY')}") 
                event.summary = component.get('SUMMARY')
                if hasattr(component.get('dtstart'), 'dt'):
                    event.start = component.get('dtstart').dt
                if hasattr(component.get('dtend'), 'dt'):
                    event.end = component.get('dtend').dt
                    if verbose:
                        print(f" Start: {event.start}, End: {event.end}")
#                now = date.today()
#                req_date_from = date(now.year, int(month), 1)
#                last_day = calendar.monthrange(now.year, int(month))[1]
#                req_date_to = date(now.year, int(month), last_day)
#                event_start_date = event.start.date() if isinstance(event.start, datetime) else event.start
#                if event_start_date < req_date_from or event_start_date > req_date_to:
#                    print(f"  Date check passed! event_start_date={event_start_date}, looking for month {month}")
#                    continue
                event.hours = event.end - event.start
                secs = event.hours.seconds
                minutes = ((secs/60)%60)/60.0
                hours = secs/3600
                event.hours = hours + minutes
                event.description = component.get('DESCRIPTION', '') # Extract description, empty string if blank
                event.location = component.get('LOCATION', '') # Extract location details, empty string if blank
                events.append(event)
            f.close()
            if verbose:
                print(f"\nTotal events added: {len(events)}")
        else:
            print("You entered ", filename, ". ")
            print(file_extension.upper(), " is not a valid file format. Looking for an ICS file.")
            exit(0)
    else:
        print("I can't find the file ", filename, ".")
        print("Please enter an ics file located in the same folder as this script.")
        exit(0)


#def csv_write(icsfile):
#    try:
#        count = 0
#        for week in weeks:
#            count = count + 1
#            csvfile = "week" + str(count) + ".csv"
#            with open(csvfile, 'w') as myfile:
#                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
#                wr.writerow(headers)
#                for event in week:
#                    values = (event.summary, event.start, event.end, event.hours, event.description)
#                    wr.writerow(values)
#    except IOError:
#        print("Could not open file! Please close Excel!")
#        exit(0)

def csv_write(icsfile):
    try:
        for year, year_events in weeks:
            csvfile = f"year_{year}.csv"
            with open(csvfile, 'w') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
                wr.writerow(headers)
                for week_num, event in year_events:
                    # Format start and end times to strip seconds
                    start_formatted = event.start.strftime('%Y-%m-%d %H:%M') if isinstance(event.start, datetime) else str(event.start)
                    end_formatted = event.end.strftime('%Y-%m-%d %H:%M') if isinstance(event.end, datetime) else str(event.end)
                    values = (week_num, event.summary, start_formatted, end_formatted, event.hours, event.location, event.description)
                    wr.writerow(values)
    except IOError:
        print("Could not open file! Please close Excel!")
        exit(0)

#def sort_by_weekly(events):
#    week = []
#    end_of_week = None
#    for event in events:
#        # Convert event.start to datetime for comparison and strip timezone
#        event_start_dt = event.start if isinstance(event.start, datetime) else datetime.combine(event.start, datetime.min.time())
#        event_start_dt = event_start_dt.replace(tzinfo=None)
#        
#        if end_of_week == None or event_start_dt > end_of_week:
#            if week:  # Only append if week has events
#                weeks.append(week)
#            week = [event]  # Start new week with this event
#            # Calculate new end_of_week
#            end_of_week = event_start_dt - timedelta(days=event_start_dt.weekday()) + timedelta(days=6)
#            end_of_week = end_of_week.replace(hour=23, minute=59)
#        else:
#            week.append(event)
#    if week:  # Don't forget the last week
#        weeks.append(week)

def sort_by_yearly(events):
    years = {}
    for event in events:
        # Convert event.start to datetime and strip timezone
        event_start_dt = event.start if isinstance(event.start, datetime) else datetime.combine(event.start, datetime.min.time())
        event_start_dt = event_start_dt.replace(tzinfo=None)
        
        year = event_start_dt.year
        week_num = event_start_dt.isocalendar()[1]  # Get ISO week number
        
        if year not in years:
            years[year] = []
        years[year].append((week_num, event))
    
    # Convert to sorted list by year
    for year in sorted(years.keys()):
        weeks.append((year, years[year]))


open_cal() # Sort events chronologically (Google Calendar exports aren always in order)
sortedevents=sorted(events, key=lambda obj: (obj.start.replace(tzinfo=None) if isinstance(obj.start, datetime) else datetime.combine(obj.start, datetime.min.time())))
#sort_by_weekly(sortedevents)
sort_by_yearly(sortedevents)
if verbose:
    print(f"Number of weeks created: {len(weeks)}")
for i, week in enumerate(weeks):
    if verbose:
        print(f"Week {i+1}: {len(week)} events")
csv_write(filename)

from pyexcel.cookbook import merge_all_to_a_book
import glob

merge_all_to_a_book(glob.glob("./*.csv"), filename_noext+".xlsx")
print("Done, your file: "+filename_noext+".xlsx")

#count = 0
#for week in weeks:
#    count = count + 1
#    csvfile = "week" + str(count) + ".csv"
#    os.remove(csvfile)

count = 0
for year, year_events in weeks:
    csvfile = f"year_{year}.csv"
    os.remove(csvfile)
