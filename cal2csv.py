import sys
import os.path
from icalendar import Calendar
import calendar
import csv
import glob
from datetime import datetime, timedelta, date
from dateutil.rrule import rrulestr
from pyexcel.cookbook import merge_all_to_a_book

# Check for verbose flag
verbose = '-v' in sys.argv
if verbose:
    sys.argv.remove('-v') # Remove it so it doesn't interfere with other argument parsing

# Check for start date flag
start_date = None   # Need to set this before flagcheck, otherwise the code will crash if the flag wasn't given
if '-s' in sys.argv:
    idx = sys.argv.index('-s')
    if idx + 1 < len(sys.argv):
        try:
            start_date = datetime.strptime(sys.argv[idx + 1], '%Y-%m-%d').date()
            sys.argv.pop(idx)   # removes flag
            sys.argv.pop(idx)   # removes date value
        except ValueError:
            print("Error: Invalid start date format. use YYYY-MM-DD")
            exit(1)

# Check for end date flag
end_date = None # Need to set this before flagcheck, otherwise the code will crash if the flag wasn't given
if '-e' in sys.argv:
    idx = sys.argv.index('-e')
    if idx + 1 < len(sys.argv):
        try:
            end_date = datetime.strptime(sys.argv[idx + 1], '%Y-%m-%d').date()
            sys.argv.pop(idx)   # removes flag
            sys.argv.pop(idx)   # removes date value
        except ValueError:
            print("Error: Invalid end date format. use YYYY-MM-DD")
            exit(1)

# Display date range info
if start_date or end_date:
    if start_date and end_date and end_date < start_date:
        print(f"Excluding events between {end_date} and {start_date}")
    else:
        date_range_msg = "Filtering events"
        if start_date:
            date_range_msg += f" from {start_date}"
        if end_date:
            date_range_msg += f" to {end_date}"
        print(date_range_msg)

filename = sys.argv[1]
filename_noext = filename[:-4]
file_extension = str(sys.argv[1])[-3:]
headers = ('Week', 'Summary', 'Start Time', 'End Time', 'Hours', 'Location', 'Description')

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
        # Always work with timezone-naive datetimes for consistency
        if isinstance(dtstart, datetime):
            until = dtstart.replace(tzinfo=None) + timedelta(days=3650) # 10 years
            dtstart_clean = dtstart.replace(tzinfo=None)
        else:
            until = dtstart + timedelta(days=3650)
            dtstart_clean = dtstart
        # Remove any UNTIL parameter from RRULE to avoid timezone conflicts
        # We'll apply our own limit with the 'until' variable
        if 'UNTIL=' in rrule_str:
            # Strip out UNTIL parameter - we have our own limit
            import re
            rrule_str = re.sub(r';UNTIL=[^;]+', '', rrule_str)
            rrule_str = re.sub(r'UNTIL=[^;]+;?', '', rrule_str)

        rrule = rrulestr(rrule_str, dtstart=dtstart_clean)

        for occurrence_start in rrule:
            # Convert both to datetime for comparison
            occurrence_dt = occurrence_start if isinstance(occurrence_start, datetime) else datetime.combine(occurrence_start, datetime.min.time())
            until_dt = until if isinstance(until, datetime) else datetime.combine(until, datetime.min.time())
            
            if occurrence_dt > until_dt:
                break

            # Check if this occurence is in exception dates
            if any(occurrence_start.date() == ex.dt.date() if hasattr(ex.dt, 'date') else occurrence_start.date() == ex.dt for ex in exdates):
                if verbose:
                    print(f" Skipping exception date: {occurrence_start}")
                continue

            occurrence_end = occurrence_start + duration
            occurrences.append((occurrence_start, occurrence_end))

    except Exception as e:
        if verbose:
            print(f" Warning: Could not parse RRULE: {e}")
        return None

    return occurrences

def filter_by_date_range(events_list):
    """Filter events by date range if specified"""
    if not start_date and not end_date:
        return events_list  # No filtering needed

    # Check for exlusion dates (end date before start date)
    exclude_mode = False
    if start_date and end_date and end_date < start_date:
        exclude_mode = True
        if verbose:
            print(f"Excluding events between {end_date} and {start_date}")

    filtered = []
    for event in events_list:
        # Convert event start to date for comparison
        event_date = event.start.date() if isinstance(event.start, datetime) else event.start

        if exclude_mode:
            # Exclude mode: keep events OUTSIDE the range
            if event_date < end_date or event_date > start_date:
                filtered.append(event)
        else:
            # Normale mode: keep events INSIDE the range
            if start_date and event_date < start_date:
                continue
            if end_date and event_date > end_date:
                continue
            filtered.append(event)

    return filtered

def open_cal():
    if os.path.isfile(filename):
        if file_extension == 'ics':
            print("Extracting events from file:", filename, "\n")
            f = open(sys.argv[1], 'rb')
            gcal = Calendar.from_ical(f.read())
            for component in gcal.walk():
                if component.name != "VEVENT":
                    continue

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
                    event.description = description
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


open_cal()
# Apply date range filter if specified
if start_date or end_date:
    events = filter_by_date_range(events)
    if verbose:
        print(f"Events after date filtering: {len(events)}")

# Sort events chronologically (Google Calendar exports aren't always in order)
sortedevents=sorted(events, key=lambda obj: (obj.start.replace(tzinfo=None) if isinstance(obj.start, datetime) else datetime.combine(obj.start, datetime.min.time())))
sort_by_yearly(sortedevents)
if verbose:
    print(f"Number of years created: {len(weeks)}")
for year, year_events in weeks:
    if verbose:
        print(f"Year {year}: {len(year_events)} events")
csv_write(filename)

merge_all_to_a_book(sorted(glob.glob("./*.csv")), filename_noext+".xlsx")
print("Done, your file: "+filename_noext+".xlsx")

count = 0
for year, year_events in weeks:
    csvfile = f"year_{year}.csv"
    os.remove(csvfile)
