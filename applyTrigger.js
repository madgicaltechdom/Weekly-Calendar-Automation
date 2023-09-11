function triggerForSync() {
  ScriptApp.newTrigger('syncFromCalendar')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}
function triggerForReport() {
  ScriptApp.newTrigger('createPivotTables')
    .timeBased()
    .everyDays(1) 
    .atHour(17)    
    .nearMinute(5) 
    .create();
}

function triggerForValidation() {
  ScriptApp.newTrigger('validationEvent')
    .timeBased()
    .everyDays(1)
    .atHour(8)    
    .nearMinute(10)
    .create();
}

