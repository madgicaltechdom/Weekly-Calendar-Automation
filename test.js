function onOpen() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Sync From Calendar', 'syncFromCalendar');
  menu.addItem('Update to Calendar', 'syncToCalendar');
  menu.addItem('Validate my sheet', 'TaskPivot');
  menu.addItem('Trigger For Calendar Sync', 'createTrigger1');
  menu.addItem('Trigger For Validation', 'createTrigger2');
  menu.addToUi();
}

// function clearSpecificRange() {
//   let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   let sheetName = spreadsheet.getActiveSheet().getName();
//   let sheet = spreadsheet.getSheetByName(sheetName);

//   let rangeToClear = sheet.getRange(1,13, sheet.getLastRow() -1 +1, sheet.getLastColumn() - 12 + 1);
//   console.log(rangeToClear.getA1Notation());
//   rangeToClear.clearContent();
// }













