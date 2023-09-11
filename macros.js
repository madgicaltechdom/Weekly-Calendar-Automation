if (typeof require !== "undefined") {
  UnitTestingApp = require("./unitTestingApp.min.js");
  CreatePivotTableHelper = require("./createPivoteTableHelper.js");
  SendMailNodification = require("./sendMailNodificationHelper.js");
}
let textMsg = "";


function TaskPivot() {
  const doc = SpreadsheetApp.getActive();
  const sheet = doc.getActiveSheet();
  sheet.deleteColumns(13, 35);
  const headers = sheet.getRange("A1:L1").getValues()[0];
  const colIdx = headers.reduce((o, k, i) => {
    o[k] = i + 1;
    return o;
  }, {});
  const data = sheet.getDataRange();

  sheet.insertColumnsAfter(13, 35);
  //Create Pivot Table for Tasks
  createPivotTable("M1", "Task", sheet, colIdx, data);
  //Create Pivot Table for Videos
  createPivotTable("R1", "Video", sheet, colIdx, data);
  //Create Pivot Table for PPT
  createPivotTable("W1", "Reading", sheet, colIdx, data);
  //Create Pivot Table for PPT
  createPivotTable("AB1", "PPT", sheet, colIdx, data);
  //Create Pivot Table for Weekly Schedule
  createPivotTable("AG1", "Weekly Schedule", sheet, colIdx, data);
  //Create Pivot Table for Feedback
  createPivotTable("AL1", "Feedback", sheet, colIdx, data);
  //Create Pivot Table for Weekly Schedule
  createPivotTable("AQ1", "Action Item", sheet, colIdx, data);
  let textSearch = sheet.createTextFinder("Grand Total").findAll();
  validateTasks(textSearch, sheet);
}

function createPivotTable(startCol, filterTask, sheet, colIdx, data) {
  let pivotTable = sheet.getRange(startCol).createPivotTable(data);
  pivotTable.addRowGroup(colIdx["Title"]).showTotals(true);
  pivotTable.addRowGroup(colIdx["Description"]).showTotals(false);
  pivotTable.addRowGroup(colIdx["Start Time"]).showTotals(false);
  let criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextContains(filterTask)
    .build();
  pivotTable.addFilter(colIdx["Title"], criteria);
  pivotTable
    .addPivotValue(
      colIdx["Total Time"],
      SpreadsheetApp.PivotTableSummarizeFunction.SUM
    )
    .setDisplayName(filterTask + " Sum");
  pivotTable
    .addPivotValue(
      colIdx["Total Time"],
      SpreadsheetApp.PivotTableSummarizeFunction.COUNT
    )
    .setDisplayName(filterTask + " Count");
  return pivotTable;
}

function validateTasks(textSearch, sheet) {
  for (let i = 0; i < textSearch.length * 1; i++) {
    let row = textSearch[i].getRow();
    let column = textSearch[i].getColumn();
    console.log(row)
    console.log(column)
    console.log(sheet)
    if (column == 13) {
      taskColumn(sheet, column, row);
    } else if (column == 18) {
      videoTaskValidation(sheet, row, column);
    } else if (column == 23) {
      readingTaskValidation(sheet, row, column);
    } else if (column == 28) {
      pptTaskValidation(sheet, row, column);
    } else if (column == 33) {
      weeklyScheduleValidation(sheet, row, column);
    } else if (column == 38) {
      feedbackTaskValidation(sheet, row, column);
    } else if (column == 43) {
      actionItemValidation(sheet, row, column);
    }
  }
  const sendMailNodification = new SendMailNodification();
  let sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let email = sheetName.getName();
  let sendEmailData = sendMailNodification.sendEmailNotification(
    textMsg,
    email
  );
  MailApp.sendEmail(email, sendEmailData["textSub"], sendEmailData["message"]);
}

function videoTaskValidation(sheet, row, column) {
  videoTicketCountLessThenFiveInWeek(sheet, row, column);
  const pivotV_data = sheet
    .getRange(`R1:V${sheet.getRange(row, column + 4).getValue() + 2}`)
    .getValues();

  for (let i = 1; i < pivotV_data.length; i++) {
    let videoLoop = i;
    videoPresentationIsScheduledBeforeSixPM(pivotV_data, videoLoop);
  }
}

function readingTaskValidation(sheet, row, column) {
  readingTicketCount(sheet, row, column);
  Logger.log(sheet.getRange(row, column + 3).getValue());
}

function pptTaskValidation(sheet, row, column) {
  pptTaskCountForWeek(sheet, row, column);
  Logger.log(sheet.getRange(row, column + 3).getValue());
}

function weeklyScheduleValidation(sheet, row, column) {
  weeklyScheduleTicketCountIsOne(sheet, row, column);
  Logger.log(sheet.getRange(row, column + 3).getValue());
}

function feedbackTaskValidation(sheet, row, column) {
  feedBackCountForWeek(sheet, row, column);
  Logger.log(sheet.getRange(row, column + 3).getValue());
  const pivotF_data = sheet
    .getRange(`AL1:AP${sheet.getRange(row, column + 4).getValue() + 2}`)
    .getValues();

  for (let i = 1; i < pivotF_data.length; i++) {
    let feedbackLoop = i;
    scheduleFeedbackTask(pivotF_data, feedbackLoop);
  }
}

function actionItemValidation(sheet, row, column) {
  actionItemTicketCount(sheet, row, column);
  Logger.log(sheet.getRange(row, column + 3).getValue());
}

function averageTaskHourIsGreaterThenThreeHour(taskHours, taskCount) {
  if (taskHours / taskCount > 3) {
    Logger.log("Average tasks hours is greater than 3 hours.");
  } else {
    textMsg = textMsg + "Average tasks hours is smaller than 3 hours. \n";
  }
}
function oneWeekTaskHourIsThirtySix(taskHours) {
  if (taskHours >= 36) {
    Logger.log(
      `The total task hours are greater than or equal to 36 for the week.`
    );
  } else {
    textMsg =
      textMsg + "The total task hours are smaller than to 36 for the week. \n";
  }
}

function taskCountIsGeaterThenTen(taskCount) {
  if (taskCount >= 10) {
    Logger.log("task count is greater than 10.");
  } else {
    textMsg = textMsg + "Task count is smaller than 10. \n";
  }
}

function taskHourAreLessThenFour(data, loop) {
  if (data[loop][3] < 4 && data[loop][0] != "Grand Total") {
    Logger.log("Task hours are less than 4");
  } else {
    textMsg = textMsg + "Task hours are greater than 4. \n";
  }
}

function taskDiscription(data, loop) {
  if (
    // data[loop][0] != "" &&
    data[loop][0] != "Grand Total" &&
    data[loop][1] != ""
  ) {
    Logger.log("The task has description in it.");
  } else {
    textMsg = textMsg + "The task doesn't has description in it. \n";
  }
}

function videoTicketCountLessThenFiveInWeek(sheet, row, column) {
  if (sheet.getRange(row, column + 4).getValue() >= 5) {
    Logger.log("The video ticket count is >=5  for a week");
    Logger.log(
      `The video ticket count is ${sheet.getRange(row, column + 4).getValue()}`
    );
  } else {
    textMsg = textMsg + "The video ticket count is less then 5 in a week. \n";
  }
}

function videoPresentationIsScheduledBeforeSixPM(pivotV_data, videoLoop) {
  if (
    new Date(pivotV_data[videoLoop][2]).getDay() === 4 &&
    new Date(pivotV_data[videoLoop][2]).getHours() < 18 &&
    pivotV_data[videoLoop][0].toUpperCase().includes("VIDEO" && "PPT")
  ) {
    Logger.log(
      `The Video presentation is scheduled before 6:00 PM ${new Date(
        pivotV_data[videoLoop][2]
      ).toLocaleString("en-us", { weekday: "long" })}`
    );
  } else {
    textMsg =
      textMsg + "The Video presentation is not scheduled before 6:00 PM. \n";
  }
}

function readingTicketCount(sheet, row, column) {
  if (sheet.getRange(row, column + 4).getValue() >= 5) {
    Logger.log(
      `The reading ticket count is ${sheet
        .getRange(row, column + 4)
        .getValue()}`
    );
  } else {
    textMsg = textMsg + "The reading ticket count is Less than 5. \n";
  }
}

function pptTaskCountForWeek(sheet, row, column) {
  if (sheet.getRange(row, column + 4).getValue() >= 4) {
    Logger.log("The PPT task count is >=4 for a week");
  } else {
    textMsg = textMsg + "The PPT task count is less then 4 in a week. \n";
  }
}

function weeklyScheduleTicketCountIsOne(sheet, row, column) {
  if (sheet.getRange(row, column + 4).getValue() === 1) {
    Logger.log("The weekly schedule ticket count is 1 for the week");
  } else {
    textMsg =
      textMsg +
      "The weekly schedule ticket count is not equal to 1 in the week. \n";
  }
}
function feedBackCountForWeek(sheet, row, column) {
  if (sheet.getRange(row, column + 4).getValue() === 1) {
    Logger.log("The feedback count is 1 for a week.");
  } else {
    textMsg = textMsg + "The feedback count is not equal to 1 in a week. \n";
  }
}

function scheduleFeedbackTask(pivotF_data, feedbackLoop) {
  if (
    new Date(pivotF_data[feedbackLoop][2]).getDay() === 5 &&
    new Date(pivotF_data[feedbackLoop][2]).getHours() < 20
  ) {
    Logger.log(
      `The Feedback task is scheduled before 8:00 PM ${new Date(
        pivotF_data[feedbackLoop][2]
      ).toLocaleString("en-us", { weekday: "long" })}`
    );
  } else {
    textMsg = textMsg + "The Feedback task is not scheduled before 8:00 PM. \n";
  }
}

function actionItemTicketCount(sheet, row, column) {
  if (sheet.getRange(row, column + 4).getValue() === 2) {
    Logger.log("The Action Item ticket count is 2 for a week.");
  } else {
    textMsg =
      textMsg + "The Action Item ticket count is not equal to 2 in a week. \n";
  }
}

function taskColumn(sheet, column, row) {
  const taskCount = sheet.getRange(row, column + 4).getValue();
  const data = sheet.getRange(`M1:Q${taskCount + 2}`).getValues();
  const taskHours = sheet.getRange(row, column + 3).getValue();
  averageTaskHourIsGreaterThenThreeHour(taskHours, taskCount);
  oneWeekTaskHourIsThirtySix(taskHours);
  taskCountIsGeaterThenTen(taskCount);
  for (let j = 1; data.length-1 * 1 > j; j++) {
    let loop = j;
    taskHourAreLessThenFour(data, loop);
    taskDiscription(data, loop);
  }
}


