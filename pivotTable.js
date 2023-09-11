function createPivotTables() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let scoreDetailsSheet = ss.getSheets();
  for(let sheet of scoreDetailsSheet){

    let activeSheetName = sheet.getName();
    if(activeSheetName !== "Setting" && activeSheetName !== "example@gmail.com"){
      console.log("*************",activeSheetName);
      let pivotTableName = "Report of " + activeSheetName;
      let pivotTableSheet = ss.getSheetByName(pivotTableName);
      let inde = ss.getSheetByName(activeSheetName).getIndex();
      console.log("sheet index : ",inde);
      
    
      if (!pivotTableSheet) {
        pivotTableSheet = ss.insertSheet(pivotTableName,inde);
      }

      let range1 = sheet.getDataRange();
      let pivotTable1 = pivotTableSheet.getRange("A1").createPivotTable(range1);

      pivotTable1.addRowGroup(13);
      let sumOfTotalTime1 = pivotTable1.addPivotValue(12, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
      sumOfTotalTime1.setDisplayName("SUM of Total Time");
      
      let countOfTitle1 = pivotTable1.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);

      countOfTitle1.setDisplayName("Number of Events");
      let avgOfTotalTime1 = pivotTable1.addPivotValue(12, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
      avgOfTotalTime1.setDisplayName("AVERAGE of Total Time");
      let maxOfTotalTime1 = pivotTable1.addPivotValue(12, SpreadsheetApp.PivotTableSummarizeFunction.MAX);
      maxOfTotalTime1.setDisplayName("MAX of Total Time");
      let minOfTotalTime1 = pivotTable1.addPivotValue(12, SpreadsheetApp.PivotTableSummarizeFunction.MIN);
      minOfTotalTime1.setDisplayName("MIN of Total Time");

      let dataRange1 = pivotTableSheet.getRange("B2:F9");
      dataRange1.setNumberFormat("#0.0");
      let typeColumn1 = pivotTableSheet.getRange("A2:A9");
      typeColumn1.setHorizontalAlignment("left"); 
      let headerColumn1 = pivotTableSheet.getRange("H1");
      headerColumn1.setHorizontalAlignment("left"); 

      pivotTableSheet.getRange(pivotTableSheet.getLastRow() + 1, 1).activate();

      let range2 = sheet.getDataRange();
      let pivotTable2 = pivotTableSheet.getRange("H1").createPivotTable(range2);

      pivotTable2.addRowGroup(13);
      pivotTable2.addColumnGroup(17);

      let sumOfTotalTime2 = pivotTable2.addPivotValue(12, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
      sumOfTotalTime2.setDisplayName("SUM of Total Time");

      let dataRange2 = pivotTableSheet.getRange("I2:J9");
      dataRange2.setNumberFormat("#0.0");
      let typeColumn2 = pivotTableSheet.getRange("H2:H9");
      typeColumn2.setHorizontalAlignment("left"); 
      let headerColumn2 = pivotTableSheet.getRange("A1");
      headerColumn2.setHorizontalAlignment("left"); 

      pivotTableSheet.getRange(pivotTableSheet.getLastRow() + 1, 8).activate();
    }
  }

}


