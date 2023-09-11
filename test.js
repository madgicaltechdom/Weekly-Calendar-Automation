function onOpen() {
  let menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Sync From Calendar', 'syncFromCalendar');
  menu.addItem('Update to Calendar', 'syncToCalendar');
  menu.addItem('Generate Report', 'createPivotTables');
  menu.addItem('Validate Events','validationEvent');
  menu.addItem('Trigger To Sync','triggerForSync');
  menu.addItem('Trigger To Report','triggerForReport');
  menu.addItem('Trigger To Validate','triggerForValidation');

  menu.addToUi();
}



function extractTitles() {
  const sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setting").getDataRange().getValues();
  console.log(sheet1[3][1]);

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let startRow = 2;
  let dataColumn = 1;
  let targetColumn = 13;
  let dataRange = sheet.getRange(startRow, dataColumn, sheet.getLastRow() - startRow + 1);
  let data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    let title = data[i][0];
    let extractedTitle = sheet1[4][1];

    let searchValue = sheet1[3][1];
    let colonIndex = title.indexOf(searchValue);

    if (colonIndex !== -1) {
      extractedTitle = title.substring(0, colonIndex).trim();
    }

    sheet.getRange(startRow + i, targetColumn).setValue(extractedTitle);
  }
}



function setDateInColumnN() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let startRow = 2;
  let dataColumn = 4;
  let targetColumn = 14;
  let dataRange = sheet.getRange(startRow, dataColumn, sheet.getLastRow() - startRow + 1);
  let data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    let dateValue = data[i][0];
    let formattedDate = Utilities.formatDate(new Date(dateValue), sheet.getParent().getSpreadsheetTimeZone(), "MM/dd/yyyy");
    sheet.getRange(startRow + i, targetColumn).setValue(formattedDate);
  }
}



function setAdjustedDateInColumnO() {
      console.log("column o")
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    console.log(sheet.getSheetName())
  let startRow = 2;
  let sourceColumn = 4;
  let targetColumn = 15;
  let dataRange = sheet.getRange(startRow, sourceColumn, sheet.getLastRow() - startRow + 1);
  let data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    let dateValue = new Date(data[i][0]);
    let adjustedDate = new Date(dateValue.getTime());
    adjustedDate.setDate(dateValue.getDate() - (dateValue.getDay() + 6) % 7); 
    let formattedDate = Utilities.formatDate(adjustedDate, sheet.getParent().getSpreadsheetTimeZone(), "yyyy-MM-dd"); // Format the date as desired
    sheet.getRange(startRow + i, targetColumn).setValue(formattedDate);
  }
}


function getWeekdayName() {
 console.log("column P")
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let startRow = 2;
  let sourceColumn = 4;
  let targetColumn = 16;
  let dataRange = sheet.getRange(startRow, sourceColumn, sheet.getLastRow() - startRow + 1);
  let data = dataRange.getValues();

  let weekdayNames = [];

  for (let i of data) {
    let dateValue = new Date(i[0]);
    let weekdayIndex = (dateValue.getDay() + 6) % 7;

    switch (weekdayIndex) {
      case 0:
        weekdayNames.push(["Monday"]);
        break;
      case 1:
        weekdayNames.push(["Tuesday"]);
        break;
      case 2:
        weekdayNames.push(["Wednesday"]);
        break;
      case 3:
        weekdayNames.push(["Thursday"]);
        break;
      case 4:
        weekdayNames.push(["Friday"]);
        break;
      case 5:
        weekdayNames.push(["Saturday"]);
        break;
      case 6:
        weekdayNames.push(["Sunday"]);
        break;
      default:
        weekdayNames.push([""]);
    }
  }

  let targetRange = sheet.getRange(startRow, targetColumn, weekdayNames.length, 1);
  console.log(weekdayNames)
  targetRange.setValues(weekdayNames);
}

function determinemeetingStatus() {
  console.log("column Q")
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let startRow = 2;
  let sourceColumn = 6;
  let targetColumn = 17;
  let dataRange = sheet.getRange(startRow, sourceColumn, sheet.getLastRow() - startRow + 1);
  let data = dataRange.getValues();

  let statuses = [];

  for (let i of data) {
    let value = i[0];
    let emails = value.split(",");
    let status = (emails.length === 1) ? "Alone" : "Meeting";
    statuses.push([status]);
  }

  let targetRange = sheet.getRange(startRow, targetColumn, statuses.length, 1);
  targetRange.setValues(statuses);
}






function setEventData(userMailId,data,sheet){
  let calendar = CalendarApp.getCalendarById(userMailId);

  for (let i =1;i<data.length;i++){

    let calEvents = calendar.getEvents(data[i][3], data[i][4]);
    for(let j of calEvents){
      if(j.getTitle() === data[i][0]){
        sheet.getRange(i+1,9).setValue(j.getCreators())
        if(data[i][6] === ""){
          sheet.getRange(i+1,7).setValue(7)
        }
        sheet.getRange(i+1,18).setValue(j.getMyStatus())
        
        let startTime = j.getStartTime();
        let endTime = j.getEndTime();
        let timeDifferenceMs = endTime - startTime;
        let hours1 = Math.floor(timeDifferenceMs / (1000 * 60 * 60));
        let hour = hours1;
        let minutes = Math.floor((timeDifferenceMs % (1000 * 60 * 60)) / (1000 * 60));
        let totaltime = hour+(minutes/60);
        sheet.getRange(i + 1, 10).setValue(hour);
        sheet.getRange(i + 1, 11).setValue(minutes);
        sheet.getRange(i + 1, 12).setValue(totaltime);
      
        break;
      }
    }
  }
}
