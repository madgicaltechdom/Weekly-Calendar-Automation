// Script to synchronize a calendar to a spreadsheet and vice versa.
//
// See https://github.com/Davepar/gcalendarsync for instructions on setting this up.
//

// Set this value to match your calendar!!!
// Calendar ID can be found in the "Calendar Address" section of the Calendar Settings.
let calendarId = "";

// Set the beginning and end dates that should be synced. beginDate can be set to Date() to use
// today. The numbers are year, month, date, where month is 0 for Jan through 11 for Dec.
// let today = new Date();
// let beginDate = new Date(today.setDate(today.getDate() - today.getDay() + 1));
// beginDate.setHours(0);
// beginDate.setMinutes(1);
// beginDate.setSeconds(1);
// beginDate.setMilliseconds(1);

// let endDate = new Date(beginDate.getFullYear(), beginDate.getMonth(), beginDate.getDate() + 5);
// endDate.setHours(24);
// endDate.setMinutes(0);
// endDate.setSeconds(0);
// endDate.setMilliseconds(0);


let today = new Date();
const dayName = today.toLocaleDateString('en-US', { weekday: 'long' });
console.log(dayName);
let beginDate;
let endDate;

switch (dayName) {
  case "Saturday":
    beginDate = new Date(today.getTime() + 48 * 60 * 60 * 1000); 
    endDate = new Date(beginDate.getTime() + 120 * 60 * 60 * 1000);
    break;
  case "Sunday":
    beginDate = new Date(today.getTime() + 24 * 60 * 60 * 1000); 
    endDate = new Date(beginDate.getTime() + 120 * 60 * 60 * 1000);
    break;
  default:
    beginDate = new Date(today.setDate(today.getDate() - today.getDay() + 1));
    endDate = new Date(beginDate.getFullYear(), beginDate.getMonth(), beginDate.getDate() + 5);
}
beginDate.setHours(0);
beginDate.setMinutes(1);
beginDate.setSeconds(1);
beginDate.setMilliseconds(1);

endDate.setHours(23);
endDate.setMinutes(59);
endDate.setSeconds(59);
endDate.setMilliseconds(0);
console.log("start **************** ",beginDate,"\n end date ****************",endDate);

// Date format to use in the spreadsheet.
let dateFormat = "M/d/yyyy H:mm";

let titleRowMap = {
  title: "Title",
  description: "Description",
  location: "Location",
  starttime: "Start Time",
  endtime: "End Time",
  guests: "Guests",
  color: "Color",
  id: "Id",
  time: "Time",
  hour: "Hour",
  minutes: "Minutes",
  totaltime: "Total Time",
};
let titleRowKeys = [
  "title",
  "description",
  "location",
  "starttime",
  "endtime",
  "guests",
  "color",
  "id",
  "time",
  "hour",
  "minutes",
  "totaltime",
];
let requiredFields = ["id", "title", "starttime", "endtime"];

// This controls whether email invites are sent to guests when the event is created in the
// calendar. Note that any changes to the event will cause email invites to be resent.
let SEND_EMAIL_INVITES = false;

// Setting this to true will silently skip rows that have a blank start and end time
// instead of popping up an error dialog.
let SKIP_BLANK_ROWS = false;

// Updating too many events in a short time period triggers an error. These values
// were successfully used for deleting and adding 240 events. Values in milliseconds.
let THROTTLE_SLEEP_TIME = 200;
let MAX_RUN_TIME = 5.75 * 60 * 1000;

// Special flag value. Don't change.
let EVENT_DIFFS_WITH_GUESTS = 999;

// Adds the custom menu to the active spreadsheet.
function onOpen() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let menuEntries = [
    {
      name: "Update from Calendar",
      functionName: "syncFromCalendar",
    },
    {
      name: "Update to Calendar",
      functionName: "syncToCalendar",
    },
  ];
  spreadsheet.addMenu("Calendar Sync", menuEntries);
  onOpen1()
  onOpen2()
}
function onOpen1() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let menuEntries = [
    {
      name: "Validate my sheet",
      functionName: "TaskPivot",
    },
  ];
  spreadsheet.addMenu("Validate Task", menuEntries);
}
function onOpen2() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let menuEntries = [
    {
      name: "For Calendar Sync",
      functionName: "createTrigger1",
    },
    {
      name: "For Validation",
      functionName: "createTrigger2",
    }
  ];
  spreadsheet.addMenu("Apply Trigger", menuEntries);
}
// Creates a mapping array between spreadsheet column and event field name
function createIdxMap(row) {
  let idxMap = [];
  for (let idx = 0; idx < row.length; idx++) {
    let fieldFromHdr = row[idx];
    for (let titleKey in titleRowMap) {
      if (titleRowMap[titleKey] == fieldFromHdr) {
        idxMap.push(titleKey);
        break;
      }
    }
    if (idxMap.length <= idx) {
      // Header field not in map, so add null
      idxMap.push(null);
    }
  }
  return idxMap;
}

// Converts a spreadsheet row into an object containing event-related fields
function reformatEvent(row, idxMap, keysToAdd) {
  let reformatted = row.reduce(function (event, value, idx) {
    if (idxMap[idx] != null) {
      event[idxMap[idx]] = value;
    }
    return event;
  }, {});
  for (let k in keysToAdd) {
    reformatted[keysToAdd[k]] = "";
  }
  return reformatted;
}

// Converts a calendar event to a psuedo-sheet event.
function convertCalEvent(calEvent) {
  let convertedEvent = {
    id: calEvent.getId(),
    title: calEvent.getTitle(),
    description: calEvent.getDescription(),
    location: calEvent.getLocation(),
    guests: calEvent
      .getGuestList()
      .map(function (x) {
        return x.getEmail();
      })
      .join(","),
    color: calEvent.getColor(),
  };
  if (calEvent.isAllDayEvent()) {
    convertedEvent.starttime = calEvent.getAllDayStartDate();
    let endtime = calEvent.getAllDayEndDate();
    if (endtime - convertedEvent.starttime === 24 * 3600 * 1000) {
      convertedEvent.endtime = "";
    } else {
      convertedEvent.endtime = endtime;
      if (endtime.getHours() === 0 && endtime.getMinutes() == 0) {
        convertedEvent.endtime.setSeconds(endtime.getSeconds() - 1);
      }
    }
  } else {
    convertedEvent.starttime = calEvent.getStartTime();
    convertedEvent.endtime = calEvent.getEndTime();
  }
  convertedEvent.totaltime =
    (convertedEvent.endtime - convertedEvent.starttime) / (1000 * 3600);
  convertedEvent.hour = Math.floor(
    (convertedEvent.endtime - convertedEvent.starttime) / (1000 * 3600)
  );
  convertedEvent.minutes =
    ((convertedEvent.endtime - convertedEvent.starttime) % (1000 * 3600)) /
    (1000 * 60);
  return convertedEvent;
}

// Converts calendar event into spreadsheet data row
function calEventToSheet(calEvent, idxMap, dataRow) {
  let convertedEvent = convertCalEvent(calEvent);

  for (let idx = 0; idx < idxMap.length; idx++) {
    if (idxMap[idx] !== null) {
      dataRow[idx] = convertedEvent[idxMap[idx]];
    }
  }
}

// Returns empty string or time in milliseconds for Date object
function getEndTime(ev) {
  return ev.endtime === "" ? "" : ev.endtime.getTime();
}

// Determines the number of field differences between a calendar event and
// a spreadsheet event
function eventDifferences(convertedCalEvent, sev) {
  let eventDiffs =
    0 +
    (convertedCalEvent.title !== sev.title) +
    (convertedCalEvent.description !== sev.description) +
    (convertedCalEvent.location !== sev.location) +
    (convertedCalEvent.starttime.toString() !== sev.starttime.toString()) +
    (getEndTime(convertedCalEvent) !== getEndTime(sev)) +
    (convertedCalEvent.guests !== sev.guests) +
    (convertedCalEvent.color !== "" + sev.color);
  if (eventDiffs > 0 && convertedCalEvent.guests) {
    // Use a special flag value if an event changed, but it has guests.
    eventDiffs = EVENT_DIFFS_WITH_GUESTS;
  }
  return eventDiffs;
}

// Determine whether required fields are missing
function areRequiredFieldsMissing(idxMap) {
  return requiredFields.some(function (val) {
    return idxMap.indexOf(val) < 0;
  });
}

// Returns list of fields that aren't in spreadsheet
function missingFields(idxMap) {
  return titleRowKeys.filter(function (val) {
    return idxMap.indexOf(val) < 0;
  });
}

// Set up formats and hide ID column for empty spreadsheet
function setUpSheet(sheet, fieldKeys) {
  sheet
    .getRange(1, fieldKeys.indexOf("starttime") + 1, 999)
    .setNumberFormat(dateFormat);
  sheet
    .getRange(1, fieldKeys.indexOf("endtime") + 1, 999)
    .setNumberFormat(dateFormat);
  sheet.hideColumns(fieldKeys.indexOf("id") + 1);
}

// Display error alert
function errorAlert(msg, evt, ridx) {
  let ui = SpreadsheetApp.getUi();
  if (evt) {
    ui.alert(
      "Skipping row: " +
      msg +
      ' in event "' +
      evt.title +
      '", row ' +
      (ridx + 1)
    );
  } else {
    ui.alert(msg);
  }
}

// Updates a calendar event from a sheet event.
function updateEvent(calEvent, convertedCalEvent, sheetEvent) {
  const numChanges = updateCalEvent(calEvent, convertedCalEvent, sheetEvent);
  Utilities.sleep(THROTTLE_SLEEP_TIME * numChanges);
  return numChanges;
}

function updateCalEvent(calEvent, convertedCalEvent, sheetEvent) {
  const updatedFields = getUpdatedFields(convertedCalEvent, sheetEvent);
  applyChanges(calEvent, updatedFields);
  return updatedFields.length;
}

function getUpdatedFields(convertedCalEvent, sheetEvent) {
  const updatedFields = [];
  if (
    convertedCalEvent.starttime.toString() !==
    sheetEvent.starttime.toString() ||
    getEndTime(convertedCalEvent) !== getEndTime(sheetEvent)
  ) {
    updatedFields.push("time");
  }
  if (convertedCalEvent.title !== sheetEvent.title) {
    updatedFields.push("title");
  }
  if (convertedCalEvent.description !== sheetEvent.description) {
    updatedFields.push("description");
  }
  if (convertedCalEvent.location !== sheetEvent.location) {
    updatedFields.push("location");
  }
  if (convertedCalEvent.color !== "" + sheetEvent.color) {
    if (sheetEvent.color > 0 && sheetEvent.color < 12) {
      updatedFields.push("color");
    }
  }
  if (convertedCalEvent.guests !== sheetEvent.guests) {
    updatedFields.push("guests");
  }
  return updatedFields;
}

function applyChanges(calEvent, updatedFields) {
  sheetEvent.sendInvites = SEND_EMAIL_INVITES;
  if (updatedFields.includes("time")) {
    if (sheetEvent.endtime === "") {
      calEvent.setAllDayDate(sheetEvent.starttime);
    } else {
      calEvent.setTime(sheetEvent.starttime, sheetEvent.endtime);
    }
  }
  if (updatedFields.includes("title")) {
    calEvent.setTitle(sheetEvent.title);
  }
  if (updatedFields.includes("description")) {
    calEvent.setDescription(sheetEvent.description);
  }
  if (updatedFields.includes("location")) {
    calEvent.setLocation(sheetEvent.location);
  }
  if (updatedFields.includes("color")) {
    calEvent.setColor("" + sheetEvent.color);
  }
  if (updatedFields.includes("guests")) {
    updateGuests(calEvent, sheetEvent.guests);
  }
}

function updateGuests(calEvent, guestListString) {
  const guestCal = calEvent
    .getGuestList()
    .map((x) => ({ email: x.getEmail(), added: false }));
  const sheetGuests = guestListString || "";
  const guests = sheetGuests
    .split(",")
    .map((x) => x.trim())
    .filter((x) => x);
  for (const g of guestCal) {
    const index = guests.indexOf(g.email);
    if (index >= 0) {
      g.added = true;
      guests.splice(index, 1);
    }
  }
  for (const guest of guests) {
    calEvent.addGuest(guest);
  }
  for (const g of guestCal) {
    if (!g.added) {
      calEvent.removeGuest(g.email);
    }
  }
}

// Synchronize from calendar to spreadsheet.
function syncFromCalendar() {
  console.info("Starting sync from calendar");
  // Get spreadsheet and data
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getActiveSheet();
  let dataColumn =sheet.getRange('M1').getValue();
  console.log("********************",dataColumn);
  if(sheet.getLastRow()>1 && dataColumn !== ""){
    // sheet.deleteRows(1, sheet.getLastRow());
    sheet.clear();
    sheet.appendRow([ 'Title',
    'Description',
    'Location',
    'Start Time',
    'End Time',
    'Guests',
    'Color',
    'Id',
    'Time',
    'Hour',
    'Minutes',
    'Total Time'])
  }
  else if(sheet.getLastRow()>1 && dataColumn === ""){
    sheet.deleteRows(2,sheet.getLastRow()-1);
  }else{
    sheet.deleteRows(2,sheet.getLastRow());
  }
  let range = sheet.getDataRange();
  let data = range.getValues();
  let eventFound = new Array(data.length);
  Logger.log(beginDate + " : " + endDate + " : " + sheet.getSheetName());
  // Get calendar and events
  let calendar = CalendarApp.getCalendarById(sheet.getSheetName());

  let calEvents = calendar.getEvents(beginDate, endDate);

  // Check if spreadsheet is empty and add a title row
  let titleRow = [];
  titleRowKeysIdx(titleRowKeys, titleRow);

  dataLenthtiTleRowKeys(data, range, sheet, titleRow, titleRowKeys);

  dataLenthIsOne(data, range, sheet, titleRow, titleRowKeys);

  // Map spreadsheet headers to indices
  let idxMap = createIdxMap(data[0]);
  let idIdx = idxMap.indexOf("id");

  // Verify header has all required fields
  allRequiedFiled(idxMap, requiredFields, titleRowMap);

  // Array of IDs in the spreadsheet
  let sheetEventIds = data.slice(1).map(function (row) {
    return row[idIdx];
  });

  // Loop through calendar events
  updateSpreadsheetData(calEvents, sheetEventIds, idxMap, data, eventFound);

  // Remove any data rows not found in the calendar
  let rowsDeleted = 0;
  rowsDeleted = deleteNonexistentEvents(
    data,
    eventFound,
    sheetEventIds,
    rowsDeleted
  );

  // Save spreadsheet changes
  range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);
  rowDelete(rowsDeleted, data, sheet);
}
function titleRowKeysIdx(titleRowKeys, titleRow) {
  for (let idx = 0; idx < titleRowKeys.length * 1; idx++) {
    titleRow.push(titleRowMap[titleRowKeys[idx]]);
  }
}
function dataLenthtiTleRowKeys(data, range, sheet, titleRow, titleRowKeys) {
  if (data.length < 1) {
    data.push(titleRow);
    range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    setUpSheet(sheet, titleRowKeys);
  }
}
function dataLenthIsOne(data, range, sheet, titleRow, titleRowKeys) {
  if (data.length == 1 && data[0].length == 1 && data[0][0] === "") {
    data[0] = titleRow;
    range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    setUpSheet(sheet, titleRowKeys);
  }
}
function allRequiedFiled(idxMap, requiredFields, titleRowMap) {
  if (areRequiredFieldsMissing(idxMap)) {
    let reqFieldNames = requiredFields
      .map(function (x) {
        return titleRowMap[x];
      })
      .join(", ");
    errorAlert("Spreadsheet must have " + reqFieldNames + " columns");
  }
}
function deleteNonexistentEvents(data, eventFound, sheetEventIds, rowsDeleted) {
  for (let idx = eventFound.length - 1; idx > 0; idx--) {
    // event doesn't exist and has an event id
    if (!eventFound[idx] && sheetEventIds[idx - 1]) {
      data.splice(idx, 1);
      rowsDeleted++;
    }
  }

  return rowsDeleted;
}

function updateSpreadsheetData(
  calEvents,
  sheetEventIds,
  idxMap,
  data,
  eventFound
) {
  for (let cidx = 0; cidx < calEvents.length * 1; cidx++) {
    let calEvent = calEvents[cidx];
    let calEventId = calEvent.getId();

    let ridx = sheetEventIds.indexOf(calEventId) + 1;
    if (ridx < 1) {
      // Event not found, create it
      ridx = data.length;
      let newRow = [];
      let rowSize = idxMap.length;
      while (rowSize--) newRow.push("");
      data.push(newRow);
    } else {
      eventFound[ridx] = true;
    }
    // Update event in spreadsheet data
    calEventToSheet(calEvent, idxMap, data[ridx]);
  }
}

function rowDelete(rowsDeleted, data, sheet) {
  if (rowsDeleted > 0) {
    sheet.deleteRows(data.length + 1, rowsDeleted);
  }
}

// Synchronize from spreadsheet to calendar.
function syncToCalendar() {
  console.info("Starting sync to calendar");

  // Get spreadsheet and data
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getActiveSheet();
  let range = sheet.getDataRange();
  let data = range.getValues();
  if (data.length < 2) {
    errorAlert("Spreadsheet must have a title row and at least one data row");
    return;
  }

  let scriptStart = Date.now();
  // Get calendar and events
  let calendar = CalendarApp.getCalendarById(sheet.getSheetName());
  if (!calendar) {
    errorAlert("Cannot find calendar. Check instructions for set up.");
  }
  let calEvents = calendar.getEvents(beginDate, endDate);
  let calEventIds = calEvents.map(function (val) {
    return val.getId();
  });

  // Map headers to indices
  let idxMap = createIdxMap(data[0]);
  let idIdx = idxMap.indexOf("id");
  let idRange = range.offset(0, idIdx, data.length, 1);
  let idData = idRange.getValues();

  // Verify header has all required fields
  if (areRequiredFieldsMissing(idxMap)) {
    let reqFieldNames = requiredFields
      .map(function (x) {
        return titleRowMap[x];
      })
      .join(", ");
    errorAlert("Spreadsheet must have " + reqFieldNames + " columns");
    return;
  }

  let keysToAdd = missingFields(idxMap);

  // Loop through spreadsheet rows
  let numAdded = 0;
  let numUpdates = 0;
  let eventsAdded = false;
  ({ numUpdates, eventsAdded, numAdded } = forLoopForSyncToCalender(
    data,
    idxMap,
    keysToAdd,
    calEventIds,
    calEvents,
    numUpdates,
    scriptStart,
    calendar,
    idData,
    eventsAdded,
    numAdded,
    idRange
  ));

  // Save spreadsheet changes
  eventAddedTrue(eventsAdded, idRange, idData);

  // Remove any calendar events not found in the spreadsheet
  let numToRemove = calEventIds.reduce(function (prevVal, curVal) {
    prevVal = curValueIsNotNull(curVal, prevVal);
    return prevVal;
  }, 0);
  numToRemoveIsGreaterThenZero(
    numToRemove,
    calEventIds,
    calEvents,
    scriptStart
  );
}

function curValueIsNotNull(curVal, prevVal) {
  if (curVal !== null) {
    prevVal++;
  }
  return prevVal;
}

function numToRemoveIsGreaterThenZero(
  numToRemove,
  calEventIds,
  calEvents,
  scriptStart
) {
  if (numToRemove > 0) {
    let ui = SpreadsheetApp.getUi();
    let response = ui.alert(
      "Delete " + numToRemove + " calendar event(s) not found in spreadsheet?",
      ui.ButtonSet.YES_NO
    );
    resposeEqualToUiButtonYes(
      response,
      ui,
      calEventIds,
      calEvents,
      scriptStart
    );
  }
}

function resposeEqualToUiButtonYes(
  response,
  ui,
  calEventIds,
  calEvents,
  scriptStart
) {
  if (response == ui.Button.YES) {
    let numRemoved = 0;
    calEventIds.forEach(function (id, idx) {
      numRemoved = idIsNotEqualToNull(
        id,
        calEvents,
        idx,
        numRemoved,
        scriptStart
      );
    });
  }
}

function idIsNotEqualToNull(id, calEvents, idx, numRemoved, scriptStart) {
  if (id != null) {
    calEvents[idx].deleteEvent();
    Utilities.sleep(THROTTLE_SLEEP_TIME);
    numRemoved++;
    numRemovedMpersentTenEqualToZero(numRemoved, scriptStart);
  }
  return numRemoved;
}

function numRemovedMpersentTenEqualToZero(numRemoved, scriptStart) {
  if (numRemoved % 10 === 0) {
    console.info(
      "%d events removed, time: %d msecs",
      numRemoved,
      Date.now() - scriptStart
    );
  }
}

function eventAddedTrue(eventsAdded, idRange, idData) {
  if (eventsAdded) {
    idRange.setValues(idData);
  }
}

function forLoopForSyncToCalender(
  data,
  idxMap,
  keysToAdd,
  calEventIds,
  calEvents,
  numUpdates,
  scriptStart,
  calendar,
  idData,
  eventsAdded,
  numAdded,
  idRange
) {
  for (let ridx = 1; ridx < data.length; ridx++) {
    let sheetEvent = reformatEvent(data[ridx], idxMap, keysToAdd);

    // Do some error checking first
    // if (!sheetEvent.title) {
    //   errorAlert("must have title", sheetEvent, ridx);
    //   continue;
    // }
    const newLocal = !sheetEvent.title;
    if (newLocal) {
      errorAlert("must have title", sheetEvent, ridx);
      continue;
    } else if (!(sheetEvent.starttime instanceof Date)) {
      errorAlert("start time must be a date/time", sheetEvent, ridx);
      continue;
    } else if (
      sheetEvent.endtime !== "" &&
      !(sheetEvent.endtime instanceof Date)
    ) {
      errorAlert("end time must be empty or a date/time", sheetEvent, ridx);
      continue;
    } else if (
      sheetEvent.endtime !== "" &&
      sheetEvent.endtime < sheetEvent.starttime
    ) {
      errorAlert(
        "end time must be after start time for event",
        sheetEvent,
        ridx
      );
      continue;
    }

    // Ignore events outside of the begin/end range desired. && If enabled, skip rows with blank/invalid start and end times
    else if (
      (SKIP_BLANK_ROWS &&
        (!(sheetEvent.starttime instanceof Date) ||
          !(sheetEvent.endtime instanceof Date))) ||
      sheetEvent.starttime > endDate ||
      (sheetEvent.endtime === "" && sheetEvent.starttime < beginDate) ||
      (sheetEvent.endtime !== "" && sheetEvent.endtime < beginDate)
    ) {
      continue;
    }

    // Determine if spreadsheet event is already in calendar and matches
    let addEvent = true;
    addEvent = sheetEventId(
      sheetEvent,
      calEventIds,
      addEvent,
      calEvents,
      numUpdates
    );
    console.info(
      "%d updates, time: %d msecs",
      numUpdates,
      Date.now() - scriptStart
    );

    ({ eventsAdded, numAdded } = addEventIsTrue(
      addEvent,
      sheetEvent,
      calendar,
      idData,
      ridx,
      eventsAdded,
      numAdded,
      scriptStart
    ));
    // If the script is getting close to timing out, save the event IDs added so far to avoid lots
    // of duplicate events.
    scriptStartIsGreaterThanStartRunTime(scriptStart, idRange, idData);
  }
  return { numUpdates, eventsAdded, numAdded };
}
function addEventIsTrue(
  addEvent,
  sheetEvent,
  calendar,
  idData,
  ridx,
  eventsAdded,
  numAdded,
  scriptStart
) {
  if (addEvent) {
    let newEvent;
    sheetEvent.sendInvites = SEND_EMAIL_INVITES;
    newEvent = sheetEventTime(sheetEvent, newEvent, calendar);
    // Put event ID back into spreadsheet
    idData[ridx][0] = newEvent.getId();
    eventsAdded = true;

    // Set event color
    sheetEventColor(sheetEvent, newEvent);

    // Throttle updates.
    numAdded++;
    Utilities.sleep(THROTTLE_SLEEP_TIME);
    numAddedMpercentTenEqualToZero(numAdded, scriptStart);
  }
  return { eventsAdded, numAdded };
}

function sheetEventTime(sheetEvent, newEvent, calendar) {
  if (sheetEvent.endtime === "") {
    newEvent = calendar.createAllDayEvent(
      sheetEvent.title,
      sheetEvent.starttime,
      sheetEvent
    );
  } else {
    newEvent = calendar.createEvent(
      sheetEvent.title,
      sheetEvent.starttime,
      sheetEvent.endtime,
      sheetEvent
    );
  }
  return newEvent;
}

function sheetEventColor(sheetEvent, newEvent) {
  if (sheetEvent.color > 0 && sheetEvent.color < 12) {
    newEvent.setColor("" + sheetEvent.color);
  }
}

function numAddedMpercentTenEqualToZero(numAdded, scriptStart) {
  if (numAdded % 10 === 0) {
    console.info(
      "%d events added, time: %d msecs",
      numAdded,
      Date.now() - scriptStart
    );
  }
}

function scriptStartIsGreaterThanStartRunTime(scriptStart, idRange, idData) {
  if (Date.now() - scriptStart > MAX_RUN_TIME) {
    idRange.setValues(idData);
  }
}

function sheetEventId(
  sheetEvent,
  calEventIds,
  addEvent,
  calEvents,
  numUpdates
) {
  if (sheetEvent.id) {
    let eventIdx = calEventIds.indexOf(sheetEvent.id);
    addEvent = eventIdxPositive(
      eventIdx,
      calEventIds,
      addEvent,
      calEvents,
      sheetEvent,
      numUpdates
    );
  }
  return addEvent;
}

function eventIdxPositive(
  eventIdx,
  calEventIds,
  addEvent,
  calEvents,
  sheetEvent,
  numUpdates
) {
  if (eventIdx >= 0) {
    calEventIds[eventIdx] = null; // Prevents removing event below
    addEvent = false;
    let calEvent = calEvents[eventIdx];
    let convertedCalEvent = convertCalEvent(calEvent);
    let eventDiffs = eventDifferences(convertedCalEvent, sheetEvent);
    eventDiffsIsPositive(
      eventDiffs,
      numUpdates,
      addEvent,
      calEvent,
      convertedCalEvent,
      sheetEvent,
      calEventIds,
      eventIdx
    );
  }
  return addEvent;
}

function eventDiffsIsPositive(
  eventDiffs,
  numUpdates,
  addEvent,
  calEvent,
  convertedCalEvent,
  sheetEvent,
  calEventIds,
  eventIdx
) {
  if (eventDiffs > 0) {
    // When there are only 1 or 2 event differences, it's quicker to
    // update the event. For more event diffs, delete and re-add the event. The one
    // exception is if the event has guests (eventDiffs=99). We don't
    // want to force guests to re-confirm, so go through the slow update
    // process instead.
    ({ numUpdates, addEvent } = eventDiffsIsSmallerThanThree(
      eventDiffs,
      numUpdates,
      calEvent,
      convertedCalEvent,
      sheetEvent,
      addEvent,
      calEventIds,
      eventIdx
    ));
  }
}

function eventDiffsIsSmallerThanThree(
  eventDiffs,
  numUpdates,
  calEvent,
  convertedCalEvent,
  sheetEvent,
  addEvent,
  calEventIds,
  eventIdx
) {
  if (eventDiffs < 3 && eventDiffs !== EVENT_DIFFS_WITH_GUESTS) {
    numUpdates += updateEvent(calEvent, convertedCalEvent, sheetEvent);
  } else {
    addEvent = true;
    calEventIds[eventIdx] = sheetEvent.id;
  }
  return { numUpdates, addEvent };
}

// Set up a trigger to automatically update the calendar when the spreadsheet is
// modified. See the instructions for how to use this.
function createSpreadsheetEditTrigger() {
  let ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("syncToCalendar").forSpreadsheet(ss).onEdit().create();
}

// Delete the trigger. Use this to stop automatically updating the calendar.
function deleteTrigger() {
  // Loop over all triggers.
  let allTriggers = ScriptApp.getProjectTriggers();
  for (let idx = 0; idx < allTriggers.length * 1; idx++) {
    if (allTriggers[idx].getHandlerFunction() === "syncToCalendar") {
      ScriptApp.deleteTrigger(allTriggers[idx]);
    }
  }
}
