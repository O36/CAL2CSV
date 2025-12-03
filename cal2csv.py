import sys
import os.path
from icalendar import Calendar
import calendar
import csv
from datetime import datetime, timedelta, date
from dateutil.rrule import rrulestr

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

def expand_recurring_event(component):
    """Expand a recurring event into individual occurrences"""
    occurrences = []

    # Check if event has recurrence rule
    if not component.get('RRULE'):
        return None # not a recurring event

    # Get the recurrence rule
    rrule_str = component.get('RRULE').to_ical().decode('utf-8')
    dtstart = component.get('dtstart').dt

    # Get event duration
    if component.get('dtend'):
        dtend = component.get('dtend').dt
        duration = dtend - dtstart
    else:
        duration = timedelta(hours=1) # Default 1 hour if no end time

    # Get exception dates (cancelled occurences)
    exdates = []
    if component.get('EXDATE'):
        exdate = component.get('EXDATE')
        if isinstance(exdate, list):
            for ex in exdate:
                exdates.extend(ex.dts)
        else:
            exdates = exdate.dts

    # Parse and expand the recurrence rule
    # Limit to reasonable number (e.g., 10 years from start)
    try:
        if isinstance(dtstart, datetime):
            until = dtstart.replace(tzinfo=None) + timedelta(days=3650) # 10 years
            dtstart_clean = dtstart.replace(tzinfo=None)
        else:
            until = dtstart + timedelta(days=3650)
            dtstart_clean = dtstart

        rrule = rrulestr(rrule_str, dtstart=dtstart_clean)

        for occurence_start in rrule:
            if occurrence_start > until:
                break

            # Check if this occurence is in exception dates
            if any(occurrence_start.date() == ex.dt.date() if hasattr(ex.dt, 'date') else occurence_start.date() == ex.dt for ex in exdates):
                if verbose:
                    print(f" Skipping exception date: {occurrence_start}")
                continue

            occurence_end = occurrence_start + duration
            occurrences.append((occurrence_start, occurrence_end))

    except Exception as e:
        if verbose:
            print(f" Warning: Could not parse RRULE: {e}")
        return None
    return occurrences

def open_cal():
    if os.path.isfile(filename):
        if file_extension == 'ics':
            print("Extracting events from file:", filename, "\n")
            f = open(sys.argv[1], 'rb')
            gcal = Calendar.from_ical(f.read())
            for component in gcal.walk():
                if component.name != "VEVENT":
                    continue

#                event = CalendarEvent("event")
                if component.get('STATUS') != "CONFIRMED":
                    continue
                if component.get('TRANSP') == 'TRANSPARENT' or component.get('TRANSP') == None:
                    continue # skip event that have not been accepted
                if component.get('SUMMARY') == None:
                    continue # skip blank items

                summary = component.get('SUMMARY')
                description = component.get('DESCRIPTION', '')
                location = component.get('LOCATION', '')

                # Check if this is a recurring event
                occurrences = expand_recurring_event(component)

                if occurrences:
                    # Recurring event - create an event for each occurrence
                    if verbose:
                        print(f"Processing recurring: {summary} ({len(occurrences)} occurrences)")

                    for occur_start, occur_end in occurrences:
                        event = CalendarEvent("event")
                        event.summary = summary
                        event.start = occur_start
                        event.end = occur_end
                        event.description = description
                        event.location = location

                        # Calculate duration per event
                        event.hours = event.end - event.start
                        if isinstance(event.start, date) and not isinstance(event.start, datetime):
                            event.hours = event.hours.days * 24
                        else:
                            secs = event.hours.seconds
                            minutes = ((secs/60)%60)/60.0
                            hours = secs/3600
                            event.hours = hours + minutes

                        events.append(event)
                else:
                    # Non-recurring event - process normally
                    if verbose:
                        print(f"Processing: {summary}")
                    
                    event = CalendarEvent("event")
                    event.summary = summary
                    event.description
                    event.location = location

                    if hasattr(component.get('dtstart'), 'dt'):
                        event.start = component.get('dtstart').dt
                    if hasattr(component.get('dtend'), 'dt'):
                        event.end = component.get('dtend').dt
                        if verbose:
                            print(f" Start: {event.start}, End: {event.end}")

                    event.hours = event.end - event.start
                    if isinstance(event.start, date) and not isinstance(event.start, datetime):
                        event.hours = event.hours.days * 24
                    else:
                        secs = event.hours.seconds
                        minutes = ((secs/60)%60)/60.0
                        hours = secs/3600
                        event.hours = hours + minutes

                    events.append(event)

            f.close()
            if verbose:
                print(f"\nTotal events added: {len(events)}")
        else:
            print("You entered ", filename, ". ")
            print(file_extension.upper(), " is not a valid file format. Looking for an ICS file.")
            exit(0)
    else:
        print("File ", filename, " not found.")
        print("Please enter a valid ics-file.")
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
