function validationEventToAll() {
  
  let maiSubject = "Pivot Table Rule Alert.";
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let getAllSheets = spreadsheet.getSheets();
  for(let sheet of getAllSheets){
    let sheet_name = sheet.getName();
    let firstWord = sheet_name.split(' ')[0];
  
    if(firstWord === "Report"){
      console.log(sheet_name,"********************",firstWord);
      message = "";
      let email = sheet_name.split(' ')[2];

      const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getDataRange().getValues();

      let grandToatalIndex = -1;
      for (let index=0; index<data.length; index++){
        if(data[index][0] === "Grand Total" && grandToatalIndex === -1){
          grandToatalIndex = index;
          break;
        }
      }
      console.log("grand total index: ",grandToatalIndex);

      const pivot_table_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange(1,1,grandToatalIndex
      +1,6).getValues();
      console.log(pivot_table_data);

      let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setting");
      const check_list_data = sheet1.getRange(7,1,sheet1.getLastRow()-6,9).getValues();
      console.log(check_list_data);
      
      checkList(pivot_table_data,check_list_data)
      

      if(message !==""){
        MailApp.sendEmail(email, maiSubject, message);
      }else{
        console.log("All Rules are Follow..... Good 👍");
      }
    }
  }
}