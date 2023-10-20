
# Weekly Calendar Automation Tool
## Introduction 
The Calendar Automation Tool is a useful tool for managing your Google Calendar more efficiently. It can do several important things to help you stay organized. First, it can fetch events and data from your Google Calendar and neatly organize them in a spreadsheet. Second, you can export events from your spreadsheet back to your Google Calendar with ease. Third, making changes to your events is a breeze in the spreadsheet, and those changes will be automatically updated in your Google Calendar. Fourth, it can generate reports that show how you're using your time, helping you understand your schedule better. Fifth, if you forget to create an event or don't follow specific rules you've set, the tool can send you a reminder via email.

## Create a new Google Sheet
To begin, go to Google Sheets, an online tool for creating spreadsheets. Create a new sheet and customize its name for meaningful identification.

## Prerequisite
* You must have the events on your Google calendar.
* In the sheet that you have created, you need to create two sub-sheets based on the names mentioned below:

    1. *Weekly Schedule CheckList* Sheet - This sheet contains all the rules that were created by madgicaltechdom. This sheet defines the eleven points related to the calendar events.

    2. *Team's Member Details* Sheet - This sheet contains all details of the team name, team lead name, and team member name as well. 

    3. *example@gmail.com* Sheet - We will use this sheet to generate duplicates for this sheet. The columns it has are: Title, Description, Location, Start Time, End Time, Guests, Color, Id, Creator Name, Hour, Minutes, Total Time.
  

## Function Applied on weekly calendar Automation Tool
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
  
## Validate My Sheet
* The TaskPivot() function will take the event's data from the email sheet and create many reports. First, this report provides details related to different task events recorded in the Google Calendar. including the total time, title count, average time, maximum time, and minimum time. All reports show the same data for video events, reading events, PPTs, etc.


* The validationEvent() After creating the all pivot, this function will take the report data from the report sheet and the validation data from the Weekly Schedule CheckList sheet. The function retrieves report data and validation data, iterates over each data point, and If you don't create the event for the required event and your event count does not match what is mentioned in the Weekly Schedule CheckList sheet, then you will get the message in the mail.

## send an email notification to the lead if team members miss the rules while preparing a calendar.
*validate_to_all.js,In this file -validateToAll()  function will run every week on saturday, this function when run it will send notification to team lead, if team members are properly preparing their calendar. In the event that a team member fails to do so, The team lead will recieve an email notification containing the necessary details.


## Explaining Unit testing files
* *UnitTestingApp.min.js* and *MockData.min.js* files are common files that are used in unit testing. We need to write manual data in the *MockData* File for checking unit testing offline or fetch data from a spreadsheet.

## Links
* [Sheet Link](https://docs.google.com/spreadsheets/d/1RYmXstNNk_fxi9Ao5nEdaFDvyZlN1CqYf6HMogVXTAg/edit?usp=sharing)


## Contributing

We welcome contributions to this repository. Please submit a pull request with your changes and we will review them as soon as possible.

README (1).md
Displaying README (1).md.
