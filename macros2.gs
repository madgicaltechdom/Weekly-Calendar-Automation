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
      const taskCount = sheet.getRange(row,(column+4)).getValue();
      const data = sheet.getRange(`M1:Q${taskCount+2}`).getValues();
      const taskHours = sheet.getRange(row,(column+3)).getValue();
      if ((taskHours/taskCount) > 3){
        Logger.log("Average tasks hours is greater than 3 hours.")
      }
      if (taskHours >= 36){
        Logger.log(`The total task hours are greater than or equal to 36 for the week.`)
      }
      if(taskCount >= 10){
        Logger.log("task count is greater than 10.");
      }
      for (i=1; data.length > i; i++){
        if (data[i][3]< 4 && data[i][0]!= "Grand Total"){
          Logger.log("Task hours are less than 4")
        }
        if (data[i][0]!= "" && data[i][0]!= "Grand Total"){
          if( data[i][1]!= ""){
            Logger.log("The task has description in it.")
          }
        }
      }
    }
    else if (column == 18) // Videos
    {
      if(sheet.getRange(row,(column+4)).getValue() >= 5){
        Logger.log("The video ticket count is >=5  for a week")
        Logger.log(`The video ticket count is ${sheet.getRange(row,(column+4)).getValue()}`)
      };

      const pivotV_data = sheet.getRange(`R1:V${sheet.getRange(row,(column+4)).getValue()+2}`).getValues()
      for(let i = 1;pivotV_data.length > i; i++){
      if (new Date(pivotV_data[i][2]).getDay() === 4 && new Date(pivotV_data[i][2]).getHours() < 18 && pivotV_data[i][0].toUpperCase().includes("VIDEO" && "PPT")){
        Logger.log(`The Video presentation is scheduled before 6:00 PM ${new Date(pivotV_data[i][2]).toLocaleString('en-us', {weekday:'long'})}`)
      }
      }
    }
    else if (column == 23) // Reading
    {
      if(sheet.getRange(row,(column+4)).getValue() >= 5){
        Logger.log(`The reading ticket count is ${sheet.getRange(row,(column+4)).getValue()}`)
      };
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }
    else if (column == 28) // PPT
    {
      if(sheet.getRange(row,(column+4)).getValue() >= 4){
        Logger.log("The PPT task count is >=4 for a week")
      };
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }
    else if (column == 33) // Weekly Schedule
    {
      if(sheet.getRange(row,(column+4)).getValue() === 1){
        Logger.log("The weekly schedule ticket count is 1 for the week")
      }
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }
    else if (column == 38) // Feedback
    {
      if(sheet.getRange(row,(column+4)).getValue()=== 1){
        Logger.log("The feedback count is 1 for a week.")
      };
      Logger.log(sheet.getRange(row,(column+3)).getValue());
      const pivotF_data = sheet.getRange(`AL1:AP${sheet.getRange(row,(column+4)).getValue()+2}`).getValues()
      for(let i = 1;pivotF_data.length > i; i++){
        if (new Date(pivotF_data[i][2]).getDay() === 5 && new Date(pivotF_data[i][2]).getHours() < 20){
          Logger.log(`The Feedback task is scheduled before 8:00 PM ${new Date(pivotF_data[i][2]).toLocaleString('en-us', {weekday:'long'})}`)
        }
      }

    }
    else if (column == 43) // Action Items
    {
      if(sheet.getRange(row,(column+4)).getValue() === 2){
        Logger.log("The Action Item ticket count is 2 for a week.")
      };
      Logger.log(sheet.getRange(row,(column+3)).getValue());
    }

  }
}
