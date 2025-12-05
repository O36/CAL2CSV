# CAL2CSV
Convert '.ics' calendar files to '.xlsx' spreadsheets ('.ics' > '.csv' > '.xlsx')

## About this project:
I've been using my calendar to keep track of events as they happen rather than for planning future events.
What i assumed would span a few weeks, maybe months ended up covering several years.
To analyze the frequency and patterns of these events, i needed to extract all the calendar data.
Unfortunately google removed the option to export calendars to CSV, only offering '.ics' format instead.
This tool bridges that gap.

**credits:** This project is based on [erikcox's ical2csv] (https://github.com/erikcox/ical2csv). Thanks for the starting point, Erik!

## Features
✅ Converts '.ics' files to '.xlsx' Excel spreadsheets  
✅ Groups events by year (one sheet per year)  
✅ Includes week numbers for easy reference  
✅ Extracts event summaries, descriptions, locations, start/end times, and duration  
✅ Handles recurring appointments (expands them into individual occurrences)  
✅ Properly calculates duration for all-day events (24h+)  
✅ removes seconds from timestamps for a cleaner output  
✅ Optional verbose mode for debugging  

## Installation

### Requierments

- Python 3.x
- Required packages:
    - 'icalendar' (for parsing '.ics' files)
    - 'pyexcel' (for creating '.xlsx' files)
    - 'python-dateutil' (for recurring event expansion)

### Install Dependencies
```bash
pip install icalendar pyexcel pyexcel-xlsx python-dateutil
```

## Usage

### Basic Usage

Extract your entire calendar to an Excel file:
```bash
python cal2csv.py MyCalendar.ics
```

This creates 'MyCalendar.xlsx' with one sheet per year.

### Verbose Mode

Run with debugging output to see what's being processed:
```bash
python cal2csv.py MyCalendar.ics -v
```

### Date Range Filtering

Filter events to a specific data range:
```bash
# Events from a start date onwards
python cal2csv.py MyCalendar.ics -s 2023-01-01

# Events up to an end date
python cal2csv.py MyCalendar.ics -e 2024-12-31

# Events within a specific range
python cal2csv.py MyCalendar.ics -s 2023-01-01 -e 2024-12-31
```

### Exclusion Mode (Inverted Date Range)

When the end date is before the start date, the filter inverts - it excludes events within that ranke and keeps everything outside it:
```bash
# Export everything EXCEPT events between 2020 and 2026
python cal2csv.py MyCalendar.ics -s 2025-12-31 -e 2020-01-01
```

### Combining Flags

All flags are optional and can be combined:
```bash
python cal2csv.py MyCalendar.ics -v -s 2023-01-01 -e 2024-12-31
```

## Output Format

The generated Excel file contains:

| Week | Summary | Start Time | End Time | Hours | Location | Description |
|------|---------|------------|----------|-------|----------|-------------|
| 42 | Team Meeting | 2024-10-15 14:00 | 2024-10-15 15:30 | 1.5 | Room 301 | Quarterly review |

- **One sheet per year** (e.g., "2023", "2024", "2025")
- **Week numbers** using ISO week numbering
- **Clean timestamps** without seconds or timezone info
- **Accurate durations** including proper handling of all-day events

## Roadmap

- [✅] Add date range filtering (start/end dates)  
       (also implemented excludemode (end date < start date)
- [✅] Handle recurring appointments 
- [✅] Fix duration counter for all-day events
- [✅] Remove seconds from timestamps
- [✅] Add event location to export
- [✅] Add verbose debugging flag
- [✅] Fix verbose output for years/events count

## How to get your .ics file

### Google Calendar

1. Go to [Google Calendar](https://calendar.google.com)
2. Click the gear icon → Settings
3. Select your calendar from the left sidebar
4. Scroll to "Integrate calendar"
5. Click the "Export" button or copy the "Secret address in iCal format"

## Contributing

Contributions are welcome! Feel free to:
- Report bugs via Issues
- Suggest features
- Submit pull requests

## License

MIT License - see [LICENSE](LICENSE) file for details

## Acknowledgments

- Original concept: [erikcox/ical2csv](https://github.com/erikcox/ical2csv)
- Built during late-night troubleshooting sessions with Claude! 🤝

