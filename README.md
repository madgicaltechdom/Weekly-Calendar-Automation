
# Calendar Automation Tool V-2
## Introduction 
The Calendar Automation Tool is a useful tool for managing your Google Calendar more efficiently. It can do several important things to help you stay organized. First, it can fetch events and data from your Google Calendar and neatly organize them in a spreadsheet. Second, you can export events from your spreadsheet back to your Google Calendar with ease. Third, making changes to your events is a breeze in the spreadsheet, and those changes will be automatically updated in your Google Calendar. Fourth, it can generate reports that show how you're using your time, helping you understand your schedule better. Fifth, if you forget to create an event or don't follow specific rules you've set, the tool can send you a reminder via email.

## Create a new Google Sheet
To begin, go to Google Sheets, an online tool for creating spreadsheets. Create a new sheet and customize its name for meaningful identification.

## Prerequisite
* You must have the events on your Google calendar.
* In the sheet that you have created, you need to create two sub-sheets based on the names mentioned below:

    1. *Setting* She - This sheet contains all the necessary credentials for event types, as well as the start and end dates for the calendar. Within this timeframe, events will be retrieved and generated, including Start Date, End Date, Event Type Separator, Non-separator Event type, Required Event's Type, Min Count, Max Count, Min Avg Time, Max Avg Time, Min Time, Max Time, Min Total Time and Max Total Time.

    2. *example@gmail.com* Sheet - We will use this sheet to generate duplicates for this sheet. The columns it has are: Title, Description, Location, Start Time, End Time, Guests, Color, Id, Creator Name, Hour, Minutes, Total Time, Event Type, Date, Week of, Day of the week, Meeting or Alone, and RSVP Status.
  

## Function Applied on Calendar Automation V-2
## Sync From Calendar
* The syncFromCalendar() function will take the start date, end date, and user email ID from the sheet name and then fetch the Google Calendar with the email ID, start date, and end date, then get all events in between the mentioned dates and set them in the Google Sheet.

* setDateInColumnN() function will fetch events, find the event date, and set it in the "Date" column.

* setAdjustedDateInColumnO() function will fetch events, find the current week date of Monday, and set it in the "Week of" column.

* getWeekdayName() function will fetch events find the event day and set it in the "Day of week" column.

* extractTitles() function will fetch events, find that this event is the meeting of alone time spent, and set it in the "Meeting or Alone" column.

* determinemeetingStatus() function will fetch events find the status of meetings and set it in the "RSVP Status" column.


## Update to Calendar
* The syncToCalendar() function will take the start date, end date, and user email ID from the sheet name and the event data from the sheet, then fetch the Google Calendar with the email ID, start date, and end date, then create the new events on the calendar using event data in between the mentioned dates and set the rest of the information automatic in the Google Sheet, like the event ID, creator name, etc.

* createIdxMap() function will find the index of the event and return the event data.

* missingFields() function will check all fields with event data, store the missing fields and return them.

* eventAddedTrue() function retrieves user mai ID data from a spreadsheet and updates the event fields in the Calendar.

* setEventData() function retrieves user mai ID data from a spreadsheet, finds the event type, and sets it in the "Event Type" column.
  
## Generate Report
* The createPivotTables() function will take the event's data from the email sheet and create two reports. First, this report provides details related to different event types recorded in the Google Calendar. including the total time, title count, average time, maximum time, and minimum time. The second report presents details of event types categorised as "meeting" or "alone," including the total time for each category and the grand total.
  
## Validate Events
* The validationEvent() function will take the report data from the report sheet and the validation data from the setting sheet. The function retrieves report data and validation data, iterates over each data point, and If you don't create the event for the required event and your event count does not match what is mentioned in the setting sheet, then you will get the message in the mail.


## Explaining Unit testing files
* *UnitTestingApp.min.js* and *MockData.min.js* files are common files that are used in unit testing. We need to write manual data in the *MockData* File for checking unit testing offline or fetch data from a spreadsheet.

## Links
* [Document Link](https://docs.google.com/document/d/1cI38RjcCSG3LHQDfrOZu2rRD7uOv4kzGzoh5ddFrDR8/edit?usp=sharing)
* [Sheet Link](https://docs.google.com/spreadsheets/d/1yhGdYYijaVlwwRiGYh-MRt1_YktRRO72pUlGfcaQU98/edit?usp=sharing)
* [Demo Video Link](https://drive.google.com/file/d/1JdCEf-9G677x59uFAvxHp3u-jEgif8_A/view)

## Contributing

We welcome contributions to this repository. Please submit a pull request with your changes and we will review them as soon as possible.

README (1).md
Displaying README (1).md.
