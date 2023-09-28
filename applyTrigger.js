function createTrigger1() {
  ScriptApp.newTrigger('syncFromCalendar')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();
}

function createTrigger2() {
  ScriptApp.newTrigger('TaskPivot')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();
}

function triggerOnValidateToAll() {
  ScriptApp.newTrigger("validateToAll")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(22)
    .everyWeeks(1)
    .create();
}
