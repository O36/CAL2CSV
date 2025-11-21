# CAL2CSV
Convertor .ics file to .xlsx file (.ics > .csv > .xlsx)

# About this project:
I have been using my calendar to keep track of certain events as they happened instead of planning events that were going to happen.
Assuming this would be needed for several weeks, maybe months. This ended up spanning several years instead.
In order to be able to analyse how often these events happened i needed to be able to extract all the data from the calendar;
unfortunately google in all its "wisdom" decided to remove the option to export a calendar to CSV, opting for .ics instead.

I based my code on the work of erikcox (https://github.com/erikcox/ical2csv), thanks for giving me a starting off point erik!

First i fixed the code up to the point where it would no longer spit out errors.
After that i updated the code to fix the exporter to .xlsx so it would actually generate a .xlsx with content.
Thirdly i reworked the code so it would include the description of the appointment.
Fourthly i reworked the code so it would give me a sheet per year, instead of per week (my .ics file supplied me with 472 weeks) and supply me with a column displaying the weeknumber of that year.

Finally i learned how to make a github account so i could opensource this project as i figured other people might have a use for it.

# Roadmap:
- add option to include start and enddates to be included in the report
- add system to fix reoccuring appointments (currently only exports first instance)
- fix duration counter to display 24 hours on whole day events
- remove seconds variable from the date.time output
- add location to the export

# Dependencies:
- icalendar (for parsing .ics files)
- openpyxl (for writing .xlsx files)
- pytz (for timezone handling)
pip install icalendar openpyxl pytz
- python3

# Usage:
Run the script followed by the location of the .ics file.
The script will run and spit out an .xlsx file.

Example: python cal2csv.py MyCalendar.ics


