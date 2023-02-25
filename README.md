- Read the data from your Excel file using the SheetJS library:
- Create an iCalendar event for each row of data:
- Write the iCalendar events to an .ics file:

In the iCal file, the timezone can be specified using the "TZID" property. For example, if you want to specify the "Europe/London" timezone, you can use the following format:

`DTSTART;TZID=Europe/London:20230215T140000`

- if there is an \__EMPTY_[i] in all 3 objects create array[date, task]
- separate file into sheets
- seperate file into users
- collate users
- convert into format for ical
- ical for each user
- use writeFile to make .ics
- test
