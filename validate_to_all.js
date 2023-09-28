function validateToAll() {
  const check_list_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Weekly Schedule CheckList").getDataRange().getValues();
  
  const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for(let sheet of allSheets){
    let sheetName = sheet.getSheetName();
    if (sheetName !== "Weekly Schedule CheckList" && sheetName !== "Team's Member Details" && sheetName !== "example@gmail.com"){
      let userSheetData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(2,4).getValue();
      let date = new Date();
      textMsg = "";
      if(userSheetData-date>=0){

        let textSearch = sheet.createTextFinder("Grand Total").findAll();
        validateTasks(textSearch, sheet, check_list_sheet,true);
      }
      else{
        let team_member_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team's Member Details").getDataRange().getValues();
        let bodyText = "One of your team members "+ sheetName+", has not yet scheduled anything on Google Calendar for the upcoming week.";
        let subjectText =  "Weekly calendar Alert of ";

        for(let t=1; t<team_member_sheet.length ; t++){
          let teamMemberData = team_member_sheet[t][2].split(',');
          if (teamMemberData.includes(sheetName)){
            if(sheetName !== team_member_sheet[t][1]){
              MailApp.sendEmail(team_member_sheet[t][1],subjectText+sheetName, bodyText);
              break
            }
          }
        }
      }
    }
  }
}

