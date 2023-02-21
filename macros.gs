function TaskPivot() {
  const doc = SpreadsheetApp.getActive();
  const sheet = doc.getActiveSheet();
  if (sheet.getMaxColumns()>35){
    sheet.deleteColumns(13,35);
  }else{
    var lastCol = sheet.getLastColumn();
    sheet.insertColumns(lastCol +1 , (50-lastCol));
  }
  
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

  var textSearch = sheet.createTextFinder('Grand Total').findAll();

  validateTasks(textSearch, sheet);
  
}

function createPivotTable(startCol, filterTask, sheet,colIdx,data){
    pivotTable = sheet.getRange(startCol).createPivotTable(data);
  pivotTable.addRowGroup(colIdx['Title']).showTotals(true);
  pivotTable.addRowGroup(colIdx['Description']).showTotals(false);
  pivotTable.addRowGroup(colIdx['Start Time']).showTotals(false);
 
  var criteria = SpreadsheetApp.newFilterCriteria()
                             .whenTextContains(filterTask)
                             .build();
  pivotTable.addFilter(colIdx['Title'],criteria);
  pivotTable.addPivotValue(colIdx['Total Time'], SpreadsheetApp.PivotTableSummarizeFunction.SUM).setDisplayName(filterTask + " Sum");
  pivotTable.addPivotValue(colIdx['Total Time'], SpreadsheetApp.PivotTableSummarizeFunction.COUNT).setDisplayName(filterTask + " Count");
  return pivotTable;

}

function validateTasks(textSearch, sheet){
  for (var i=0; i < textSearch.length; i++) {
    var row = textSearch[i].getRow();
    var column = textSearch[i].getColumn();
    // Task Column
    if(column == 13) {
      Logger.log(sheet.getRange(row,(column+4)).getValue());
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }else if (column == 18) // Videos
    {
      Logger.log(sheet.getRange(row,(column+4)).getValue());
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }else if (column == 23) // Reading
    {
      Logger.log(sheet.getRange(row,(column+4)).getValue());
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }
    else if (column == 28) // PPT
    {
      Logger.log(sheet.getRange(row,(column+4)).getValue());
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }
    else if (column == 33) // Weekly Schedule
    {
      Logger.log(sheet.getRange(row,(column+4)).getValue());
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }else if (column == 38) // Feedback
    {
      Logger.log(sheet.getRange(row,(column+4)).getValue());
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }else if (column == 43) // Action Items
    {
      Logger.log(sheet.getRange(row,(column+4)).getValue());
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }

  }
}