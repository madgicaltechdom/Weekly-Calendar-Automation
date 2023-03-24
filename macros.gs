function TaskPivot() {
  const doc = SpreadsheetApp.getActive();
  const sheet = doc.getActiveSheet();
  sheet.deleteColumns(13,35);
  
  const headers = sheet.getRange('A1:L1').getValues()[0];
  const colIdx = headers.reduce((o, k, i) => {
      o[k] = i+1;
      return o;
  }, {});
  const data = sheet.getDataRange();
 
  sheet.insertColumnsAfter(13,35);
  //Create Pivot Table for Tasks
  createPivotTable("M1", "Task",sheet,colIdx,data);
  //Create Pivot Table for Videos
  createPivotTable("R1", "Video",sheet,colIdx,data);
  //Create Pivot Table for PPT
  createPivotTable("W1", "Reading",sheet,colIdx,data);
  //Create Pivot Table for PPT
  createPivotTable("AB1", "PPT",sheet,colIdx,data);
  //Create Pivot Table for Weekly Schedule
  createPivotTable("AG1", "Weekly Schedule",sheet,colIdx,data);
  //Create Pivot Table for Feedback
  createPivotTable("AL1", "Feedback",sheet,colIdx,data);
  //Create Pivot Table for Weekly Schedule
  createPivotTable("AQ1", "Action Item",sheet,colIdx,data);

  let textSearch = sheet.createTextFinder('Grand Total').findAll();

  validateTasks(textSearch, sheet);
  
}

function createPivotTable(startCol, filterTask, sheet,colIdx,data){
  let  pivotTable = sheet.getRange(startCol).createPivotTable(data);
  pivotTable.addRowGroup(colIdx['Title']).showTotals(true);
  pivotTable.addRowGroup(colIdx['Description']).showTotals(false);
  pivotTable.addRowGroup(colIdx['Start Time']).showTotals(false);
 
  let criteria = SpreadsheetApp.newFilterCriteria()
                             .whenTextContains(filterTask)
                             .build();
  pivotTable.addFilter(colIdx['Title'],criteria);
  pivotTable.addPivotValue(colIdx['Total Time'], SpreadsheetApp.PivotTableSummarizeFunction.SUM).setDisplayName(filterTask + " Sum");
  pivotTable.addPivotValue(colIdx['Total Time'], SpreadsheetApp.PivotTableSummarizeFunction.COUNT).setDisplayName(filterTask + " Count");
  return pivotTable;

}

let textMsg = "";
function validateTasks(textSearch, sheet) {
  const tasks = [
    { column: 13, fn: validateTaskColumn },
    { column: 18, fn: validateVideoColumn },
    { column: 23, fn: validateReadingColumn },
    { column: 28, fn: validatePptColumn },
    { column: 33, fn: validateWeeklyScheduleColumn },
    { column: 38, fn: validateFeedbackColumn },
    { column: 43, fn: validateActionItemColumn },
  ];

  tasks.forEach((task) => {
    const column = task.column;
    const row = textSearch.find((cell) => cell.getColumn() === column).getRow();
    task.fn(sheet, row, column);
  });

  const sheetName = SpreadsheetApp.getActiveSpreadsheet()
    .getActiveSheet()
    .getName();
  const email = sheetName;
  const textSub = "Pivot Table Rule Alert.";
  const message = textMsg
    .split(". ")
    .filter((sentence, i, arr) => arr.indexOf(sentence) === i)
    .join(". ");

  MailApp.sendEmail(email, textSub, message);
}

function validateTaskColumn(sheet, row, column) {
  const taskCount = sheet.getRange(row, column + 4).getValue();
  const data = sheet.getRange(`M1:Q${taskCount + 2}`).getValues();
  const taskHours = sheet.getRange(row, column + 3).getValue();

  taskHoursIsGreaterThanThree(taskHours, taskCount);
  weekTaskHoursLessThanThirtySix(taskHours);
  taskCountLessThanTen(taskCount);

  data.slice(1).forEach((dataLoop, i) => {
    taskHourGreaterOrEqualToFourHour(data, i + 1);
    aboutTaskDiscription(data, i + 1);
  });
}

function validateVideoColumn(sheet, row, column) {
  videoTaskCountIsLessThanFive(sheet, row, column);
  const pivotV_data = sheet
    .getRange(`R1:V${sheet.getRange(row, column + 4).getValue() + 2}`)
    .getValues();
  pivotV_data.slice(1).forEach((pivotV_dataLoop, i) => {
    scheduledVideoPresentationBeforeSix(pivotV_data, i + 1);
  });
}

function validateReadingColumn(sheet, row, column) {
  readingTicketCount(sheet, row, column);
  Logger.log(sheet.getRange(row, column + 3).getValue());
}

function validatePptColumn(sheet, row, column) {
  pptTaskCount(sheet, row, column);
  Logger.log(sheet.getRange(row, column + 3).getValue());
}

function validateWeeklyScheduleColumn(sheet, row, column) {
  weeklyScheduledTicketCount(sheet, row, column);
  Logger.log(sheet.getRange(row, column + 3).getValue());
}

function validateFeedbackColumn(sheet, row, column) {
  feedBackCount(sheet, row, column);
  Logger.log(sheet.getRange(row, column + 3).getValue());
  const pivotF_data = sheet
    .getRange(`AL1:AP${sheet.getRange(row, column + 4).getValue() + 2}`)
    .getValues();
  pivotF_data.slice(1).forEach((pivotF_dataLoop, i) => {
    feedbackTaskScheduledBeforeEight(pivotF_data, i + 1);
  });
}

function validateActionItemColumn(sheet, row, column) {
  actionItemTicketCount(sheet, row, column);
  Logger.log(sheet.getRange(row, column + 3).getValue());
}

function taskHoursIsGreaterThanThree(taskHours, taskCount) {
  if (taskHours / taskCount > 3) {
    Logger.log("Average tasks hours is greater than 3 hours.");
  } else {
    textMsg = textMsg + "Average tasks hours is less than 3 hours. \n";
  }
}

function weekTaskHoursLessThanThirtySix(taskHours) {
  if (taskHours >= 36) {
    Logger.log(
      `The total task hours are greater than or equal to 36 for the week.`
    );
  } else {
    textMsg =
      textMsg + "The total task hours are less than 36 for the week. \n";
  }
}
function taskCountLessThanTen(taskCount) {
  if (taskCount >= 10) {
    Logger.log("Task count is greater than 10.");
  } else {
    textMsg = textMsg + "Task count is less than 10. \n";
  }
}

function taskHourGreaterOrEqualToFourHour(data, dataLoop) {
  if (data[dataLoop][3] < 4 && data[dataLoop][0] != "Grand Total") {
    Logger.log("Task hours are less than 4");
  } else {
    textMsg = textMsg + "Task hours are greater than or equal to 4. \n";
  }
}

function aboutTaskDiscription(data, dataLoop) {
  if (
    data[dataLoop][0] != "" &&
    data[dataLoop][0] != "Grand Total" &&
    data[dataLoop][1] != ""
  ) {
    Logger.log("The task has description in it.");
  } else {
    textMsg = textMsg + "The task has not description in it. \n";
  }
}

function videoTaskCountIsLessThanFive(sheet, row, column) {
  if (sheet.getRange(row, column + 4).getValue() >= 5) {
    Logger.log("The video ticket count is >=5  for a week");
    Logger.log(
      `The video ticket count is ${sheet.getRange(row, column + 4).getValue()}`
    );
  } else {
    textMsg = textMsg + "The video ticket count is less than 5 for a week. \n";
  }
}

function scheduledVideoPresentationBeforeSix(pivotV_data, pivotV_dataLoop) {
  if (
    new Date(pivotV_data[pivotV_dataLoop][2]).getDay() === 4 &&
    new Date(pivotV_data[pivotV_dataLoop][2]).getHours() < 18 &&
    pivotV_data[pivotV_dataLoop][0].toUpperCase().includes("VIDEO" && "PPT")
  ) {
    Logger.log(
      `The Video presentation is scheduled before 6:00 PM ${new Date(
        pivotV_data[pivotV_dataLoop][2]
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
    textMsg = textMsg + `The reading ticket count is less than five. \n`;
  }
}
function pptTaskCount(sheet, row, column) {
  if (sheet.getRange(row, column + 4).getValue() >= 4) {
    Logger.log("The PPT task count is >=4 for a week");
  } else {
    textMsg = textMsg + "The PPT task count is less than 4 for a week. \n";
  }
}
function weeklyScheduledTicketCount(sheet, row, column) {
  if (sheet.getRange(row, column + 4).getValue() === 1) {
    Logger.log("The weekly schedule ticket count is 1 for the week");
  } else {
    textMsg =
      textMsg +
      "The weekly schedule ticket count is not equal to 1 for the week. \n";
  }
}
function feedBackCount(sheet, row, column) {
  if (sheet.getRange(row, column + 4).getValue() === 1) {
    Logger.log("The feedback count is 1 for a week.");
  } else {
    textMsg = textMsg + "The feedback count is not equal to 1 for a week. \n";
  }
}
function feedbackTaskScheduledBeforeEight(pivotF_data, pivotF_dataLoop) {
  if (
    new Date(pivotF_data[pivotF_dataLoop][2]).getDay() === 5 &&
    new Date(pivotF_data[pivotF_dataLoop][2]).getHours() < 20
  ) {
    Logger.log(
      `The Feedback task is scheduled before 8:00 PM ${new Date(
        pivotF_data[pivotF_dataLoop][2]
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
      textMsg + "The Action Item ticket count is not equal to 2 for a week. \n";
  }
}
