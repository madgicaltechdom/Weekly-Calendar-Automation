let message = "";
function validationEvent() {
  
  let maiSubject = "Pivot Table Rule Alert.";
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let report_sheet = sheet.getName();
  console.log("*******************",report_sheet);
  let email = report_sheet.split(' ')[2];

  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(report_sheet).getDataRange().getValues();

  let grandToatalIndex = -1;
  for (let index=0; index<data.length; index++){
    if(data[index][0] === "Grand Total" && grandToatalIndex === -1){
      grandToatalIndex = index;
      break;
    }
  }
  console.log("grand total index: ",grandToatalIndex);

  const pivot_table_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(report_sheet).getRange(1,1,grandToatalIndex
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
  
  
function checkList(pivot_table_data,check_list_data){

  for (let point = 1; point < check_list_data.length; point++) {
    
    let taskType = check_list_data[point][0];
    let typesData = check_list_data[point];

    pivotTable(pivot_table_data,typesData,taskType)


  }
}


function pivotTable(pivot_table_data,typesData,taskType){
  let flag = false;
  for (let type = 1; type < pivot_table_data.length-1; type++) {
    if (pivot_table_data[type][0] === taskType) {
      flag = true;

      let taskType = pivot_table_data[type][0];
      let sumOfToatalTime = pivot_table_data[type][1];
      let totalCount = pivot_table_data[type][2];
      let averageOfTotalTime = pivot_table_data[type][3];
      let maxOfTotalTime = pivot_table_data[type][4];

      typeDataCkeck(typesData,taskType,sumOfToatalTime,totalCount,averageOfTotalTime,maxOfTotalTime)

    }
  }
  
  if (flag === false){
    message += ("You haven't created an event for the "+taskType+" ."+'\n');
  }

}

function typeDataCkeck(typesData,taskType,sumOfToatalTime,totalCount,averageOfTotalTime,maxOfTotalTime){

  for (let p = 1; p < typesData.length; p++) {
    emptyEvent1(typesData[p],p,taskType,totalCount)
    emptyEvent2(typesData[p],p,taskType,totalCount)
    emptyEvent3(typesData[p],p,taskType,averageOfTotalTime)
    emptyEvent4(typesData[p],p,taskType,averageOfTotalTime)
    emptyEvent5(typesData[p],p,taskType,maxOfTotalTime)
    emptyEvent6(typesData[p],p,taskType,maxOfTotalTime)
    emptyEvent7(typesData[p],p,taskType,totalCount)
    emptyEvent8(typesData[p],p,taskType,totalCount)

  }
}

function emptyEvent1(data1,p,taskType,totalCount){
  if (data1 !== '' && p === 1){
    totalCountEvent(totalCount,data1,taskType)
  }
}
function emptyEvent2(data2,p,taskType,totalCount){
  if (data2 !== '' && p === 2){
    totalCountEvent1(totalCount,data2,taskType)
  }
}
function emptyEvent3(data3,p,taskType,averageOfTotalTime){
  if (data3 !== '' && p === 3){
    averageOfTotalTimeEvent(averageOfTotalTime,data3,taskType)
  }
}
function emptyEvent4(data4,p,taskType,averageOfTotalTime){
  if (data4 !== '' && p === 4){
    averageOfTotalTimeEvent1(averageOfTotalTime,data4,taskType)
  }
}
function emptyEvent5(data5,p,taskType,maxOfTotalTime){
  if (data5 !== '' && p === 5){
    maxOfTotalTimeEvent(maxOfTotalTime,data5,taskType)
  }
}
function emptyEvent6(data6,p,taskType,maxOfTotalTime){
  if (data6 !== '' && p === 6){
    maxOfTotalTimeEvent1(maxOfTotalTime,data6,taskType)
  }
}
function emptyEvent7(data7,p,taskType,sumOfToatalTime){
  if (data7 !== '' && p === 7){
    sumOfToatalTimeEvent(sumOfToatalTime,data7,taskType)
  }
}
function emptyEvent8(data8,p,taskType,sumOfToatalTime){
  if (data8 !== '' && p === 8){
    sumOfToatalTimeEvent1(sumOfToatalTime,data8,taskType)
  }
}


function totalCountEvent(totalCount,typesData1,taskType){
  if(totalCount >= typesData1){
    console.log(taskType+" Minimum Count is grater then "+typesData1);
  }else{
    message += (taskType+" Minimum Count is less then : "+String(typesData1)+'\n');
  }
}

function totalCountEvent1(totalCount,typesData2,taskType){
  if(totalCount <= typesData2){
    console.log(taskType+" Maximum Count is less then "+typesData2);
  }else{
    message += (taskType+" Maximum Count is Grater then : "+String(typesData2)+'\n');
  }
}

function averageOfTotalTimeEvent(averageOfTotalTime,typesData3,taskType){
  if(averageOfTotalTime >= typesData3){
    console.log(taskType+" Minimum Average Time is grater then "+typesData3);
  }else{
    message += (taskType+" Minimum Average Time is less then : "+String(typesData3)+'\n');
  }
}

function averageOfTotalTimeEvent1(averageOfTotalTime,typesData4,taskType){
  if(averageOfTotalTime >= typesData4){
    console.log(taskType+" Maximum Average Time is less then "+typesData4);
  }else{
    message += (taskType+" Maximum Average Time is grater then : "+String(typesData4)+'\n');
  }
}

function maxOfTotalTimeEvent(maxOfTotalTime,typesData5,taskType){
  if(maxOfTotalTime >= typesData5){
    console.log(taskType+" Minimum Time Of Event is grater then "+typesData5);
  }else{
    message += (taskType+" Minimum Time Of Event is less then : "+String(typesData5)+'\n');
  }
}

function maxOfTotalTimeEvent1(maxOfTotalTime,typesData6,taskType){
  if(maxOfTotalTime <= typesData6){
    console.log(taskType+" Maximum Time Of Event is less then "+typesData6);
  }else{
    message += (taskType+" Maximum Time Of Event is Grater then : "+String(typesData6)+'\n');
  }
}

function sumOfToatalTimeEvent(sumOfToatalTime,typesData7,taskType){
  if(sumOfToatalTime >= typesData7){
    console.log(taskType+" Minimum Sum of Total Time is grater then "+typesData7);
  }else{
    message += (taskType+" Minimum Sum of Total Time is less then : "+String(typesData7)+'\n');
  }
}

function sumOfToatalTimeEvent1(sumOfToatalTime,typesData8,taskType){
  if(sumOfToatalTime <= typesData8){
    console.log(taskType+" Maximum Sum of Total Time is less then "+typesData8);
  }else{
    message += (taskType+" Maximum Sum Of Total Time is Grater then : "+String(typesData8)+'\n');
  }
}
