if (typeof require !== "undefined") {
  UnitTestingApp = require("./unitTestingApp.min.js");
  CreatePivotTableHelper = require("./createPivoteTableHelper.js");
  SendMailNodification = require("./sendMailNodificationHelper.js");
}
let textMsg = "";
let objArray = []; // Initialize an array to store objArrays
let datas;
let teamMemberData;
function TaskPivot() {
  clearSpecificRange();
  const check_list_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Weekly Schedule CheckList").getDataRange().getValues();

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
  createPivotTable("AV1", "Review", sheet, colIdx, data);
  let textSearch = sheet.createTextFinder("Grand Total").findAll();

  validateTasks(textSearch, sheet, check_list_sheet);
}

function clearSpecificRange() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = spreadsheet.getActiveSheet().getName();
  let sheet = spreadsheet.getSheetByName(sheetName);

  let rangeToClear = sheet.getRange(1, 13, sheet.getLastRow() - 1 + 1, sheet.getLastColumn() - 12 + 1);
  // console.log(rangeToClear.getA1Notation());
  rangeToClear.clearContent();
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

function validateTasks(textSearch, sheet, check_list_sheet, flag) {
  for (let i = 0; i < textSearch.length * 1; i++) {
    let row = textSearch[i].getRow();
    let column = textSearch[i].getColumn();
    // console.log(row)
    // console.log(column)
    // console.log(sheet)
    if (column == 13) {
      taskColumn(sheet, column, row, check_list_sheet);
    } else if (column == 18) {
      videoTaskValidation(sheet, row, column, check_list_sheet);
    } else if (column == 23) {
      readingTaskValidation(sheet, row, column, check_list_sheet);
    } else if (column == 28) {
      pptTaskValidation(sheet, row, column, check_list_sheet);
    } else if (column == 33) {
      weeklyScheduleValidation(sheet, row, column, check_list_sheet);
    } else if (column == 38) {
      feedbackTaskValidation(sheet, row, column, check_list_sheet);
    } else if (column == 43) {
      actionItemValidation(sheet, row, column, check_list_sheet);
    } else if (column == 48) {
      reviewTicketsValidation(sheet, row, column, check_list_sheet);
    }
  }

  let email;
  if (flag === true) {
    email = sheet.getSheetName();
  }
  else {
    let sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    email = sheetName.getName();
  }
  const sendMailNodification = new SendMailNodification();
  let sendEmailData = sendMailNodification.sendEmailNotification(
    textMsg,
    email
  );


  const team_member_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team's Member Details").getDataRange().getValues();
  if (flag !== true) {
    if (sendEmailData["message"] !== "") {
      MailApp.sendEmail(email, sendEmailData["textSub"], sendEmailData["message"]);
    }
    else {
      console.log("you follow the all 👍 checklist points.  😀 Good!");
    }

  }
  else if (flag === true) {

    function check() {
      for (let t = 1; t < team_member_sheet.length; t++) {

  


        
        teamMemberData = team_member_sheet[t][2].split(',');

        if (teamMemberData.includes(email) && sendEmailData["message"] !== "") {
          if (email !== team_member_sheet[t][1]) {
            let obj = {
              'TeamLeadEmail': team_member_sheet[t][1],
              'teamMember': email,
              'message': sendEmailData["message"]
            };
            objArray.push(obj);

          }
        }
      }
      return objArray
    }
    let newdata = check(team_member_sheet)
    const uniqueEmails = new Set();
    newdata.forEach(obj => {
      uniqueEmails.add(obj.TeamLeadEmail);
    });

    const resultArray = [];

    [...uniqueEmails].forEach(email => {
      const teamMembers = newdata
        .filter(obj => obj.TeamLeadEmail === email)
        .map(obj => ({ teamMember: obj.teamMember, message: obj.message }));

      resultArray.push({
        TeamLeadEmail: email,
        teamMembers: teamMembers
      });
    });

    resultArray.forEach((member) => {

      const teamLead = member.TeamLeadEmail
      console.log("*****************",teamLead);

      const teamMembers = member.teamMembers;
      const subject = "Team Member Updates";
      let messageBody = "Hello,\n\n";

      teamMembers.forEach(member => {
        messageBody += `Team Member: ${member.teamMember}\n`;
        messageBody += `Missing rules:\n${member.message}\n\n`;
      });

      if (teamMemberData.length == teamMembers.length) {
        console.log("hhhhhhhhhhhheeeeeeeeeee");
        MailApp.sendEmail({
          to: teamLead,
          subject: subject,
          body: messageBody,
        });
      }

    })
  }
}


function videoTaskValidation(sheet, row, column, check_list_sheet) {
  if (check_list_sheet[4][1] === true) {
    videoTicketCountLessThenFiveInWeek(sheet, row, column);
  }
  if (check_list_sheet[6][1] === true) {
    const pivotV_data = sheet
      .getRange(`R1:V${sheet.getRange(row, column + 4).getValue() + 2}`)
      .getValues();
    for (let i = 1; i < pivotV_data.length - 1; i++) {
      let videoLoop = i;
      let flag = videoPresentationIsScheduledBeforeSixPM(pivotV_data, videoLoop);
      if (flag === false) {
        textMsg = textMsg + "The Video presentation is not scheduled before 6:00 PM. \n";
        break;
      }
    }
  }
}

function readingTaskValidation(sheet, row, column, check_list_sheet) {
  if (check_list_sheet[3][1] === true) {
    readingTicketCount(sheet, row, column);
    Logger.log(sheet.getRange(row, column + 3).getValue());
  }
  else {
    console.log(check_list_sheet[3][0], "this check list is False.");
  }
}

function pptTaskValidation(sheet, row, column, check_list_sheet) {
  if (check_list_sheet[5][1] === true) {
    pptTaskCountForWeek(sheet, row, column);
    Logger.log(sheet.getRange(row, column + 3).getValue());
  }
}

function weeklyScheduleValidation(sheet, row, column, check_list_sheet) {
  if (check_list_sheet[9][1] === true) {
    weeklyScheduleTicketCountIsOne(sheet, row, column);
    Logger.log(sheet.getRange(row, column + 3).getValue());
  }
}

function feedbackTaskValidation(sheet, row, column, check_list_sheet) {
  if (check_list_sheet[7][1] === true) {
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
}

function actionItemValidation(sheet, row, column, check_list_sheet) {
  if (check_list_sheet[8][1] === true) {
    actionItemTicketCount(sheet, row, column);
    Logger.log(sheet.getRange(row, column + 3).getValue());
  }
}

function reviewTicketsValidation(sheet, row, column, check_list_sheet) {
  if (check_list_sheet[10][1] === true) {
    reviewTicketCount(sheet, row, column);
    Logger.log(sheet.getRange(row, column + 3).getValue());
  }
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

  if ((data[loop][3] <= 4) && (data[loop][0] != "Grand Total")) {
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
    return true;
  } else {
    return false;
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
    // new Date(pivotV_data[videoLoop][2]).getDay() === 4 &&
    new Date(pivotV_data[videoLoop][2]).getHours() < 18
    // pivotV_data[videoLoop][0].toUpperCase().includes("VIDEO" && "PPT")
  ) {
    Logger.log(
      `The Video presentation is scheduled before 6:00 PM ${new Date(
        pivotV_data[videoLoop][2]
      ).toLocaleString("en-us", { weekday: "long" })}`
    );
    return true;

  } else {
    return false;
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
    textMsg = textMsg + "The reading ticket count is Less then 5. \n";
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

function reviewTicketCount(sheet, row, column) {

  if (sheet.getRange(row, column + 4).getValue() >= 10) {
    Logger.log("The Review ticket count is Grater then or equal to 10 for a week.");
  } else {
    textMsg =
      textMsg + "The Review ticket count is Less then 10 for a week.\n";
  }
}

function taskColumn(sheet, column, row, check_list_sheet) {
  const taskCount = sheet.getRange(row, column + 4).getValue();
  const data = sheet.getRange(`M1:Q${taskCount + 2}`).getValues();
  const taskHours = sheet.getRange(row, column + 3).getValue();
  if (check_list_sheet[0][1] === true) {
    averageTaskHourIsGreaterThenThreeHour(taskHours, taskCount);
    oneWeekTaskHourIsThirtySix(taskHours);
    taskCountIsGeaterThenTen(taskCount);
  }
  if (check_list_sheet[1][1] && check_list_sheet[2][1] === true) {
    for (let j = 1; j < data.length - 1; j++) {
      let loop = j;
      taskHourAreLessThenFour(data, loop);
      let flag = taskDiscription(data, loop);
      if (flag === false) {
        textMsg = textMsg + "The task doesn't has description in it. \n";
        break;
      }
    }
  }
}

function checkAndSendEmailForNotValidate() {
  function getAllSheetNames() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var sheetNames = [];

    for (var i = 0; i < sheets.length; i++) {
      sheetNames.push(sheets[i].getName());
    }

    return sheetNames;
  }

  var userSheetNames = getAllSheetNames();

  for (var i = 0; i < userSheetNames.length; i++) {
    var sheetName = userSheetNames[i];

    if (
      sheetName !== "Weekly Schedule CheckList" &&
      sheetName !== "Team's Member Details" &&
      sheetName !== "example@gmail.com"
    ) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var valueM1 = sheet.getRange("M1").getValue(); 

      if (!valueM1) {
        var subject = 'Not Validate sheet';
        var body = `Dear Team Member,\n\nYou have not yet validated your tasks. Please review and validate them as soon as possible.\n\nThank you,\n`+sheetName;

        MailApp.sendEmail({
          to: sheetName, // Replace with your email address
          subject: subject,
          body: body,
        });
      }
    }
  }
}
