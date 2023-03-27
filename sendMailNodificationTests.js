if (typeof require !== "undefined") {
  UnitTestingApp = require("./unitTestingApp.min.js.js");
  SendMailNodification = require("./sendMailNodificationHelper.js");
}

function runTests() {
  const test = new UnitTestingApp();
  test.enable();
  test.clearConsole();
  const sendMailNodification = new SendMailNodification();
  test.runInGas(false);
  test.printHeader("LOCAL TESTS");
  let sendEmailData = sendMailNodification.sendEmailNotification("", "");
  test.assert(
    () => sendEmailData["email"] === "devendra.choukiker@medgicaltechdom.com",
    "Send Mail Id working Fine."
  );
  test.assert(
    () => sendEmailData["textSub"] === "Pivot Table Rule Alert.",
    "Send Mail Subject working Fine."
  );
  test.assert(
    () =>
      sendEmailData["message"] ===
      "The reading ticket count is Less then 5. The weekly schedule ticket count is not equal to 1 in the week. The Action Item ticket count is not equal to 2 in a week. The Feedback task is not scheduled before 8:00 PM. The PPT task count is less then 4 in a week. Average tasks hours is smaller than 3 hours. The total task hours are smaller than to 36 for the week. Task count is smaller than 10. The task doesn't has description in it. Task hours are greater than 4. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM.",
    "Send Mail Message working Fine."
  );

  test.runInGas(true);
  test.printHeader("ONLINE TESTS");
  let textMsg =
    "The reading ticket count is Less then 5. The weekly schedule ticket count is not equal to 1 in the week. The Action Item ticket count is not equal to 2 in a week. The Feedback task is not scheduled before 8:00 PM. The Feedback task is not scheduled before 8:00 PM. The PPT task count is less then 4 in a week. Average tasks hours is smaller than 3 hours. The total task hours are smaller than to 36 for the week. Task count is smaller than 10. The task doesn't has description in it. The task doesn't has description in it. The task doesn't has description in it. The task doesn't has description in it. The task doesn't has description in it. The task doesn't has description in it. Task hours are greater than 4. The task doesn't has description in it. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM.";
  sendEmailData = sendMailNodification.sendEmailNotification(
    textMsg,
    "devendra.choukiker@medgicaltechdom.com"
  );
  test.assert(
    () => sendEmailData["email"] === "devendra.choukiker@medgicaltechdom.com",
    "Send Mail Id working Fine."
  );
  test.assert(
    () => sendEmailData["textSub"] === "Pivot Table Rule Alert.",
    "Send Mail Subject working Fine."
  );

  test.assert(
    () =>
      sendEmailData["message"] ===
      "The reading ticket count is Less then 5. The weekly schedule ticket count is not equal to 1 in the week. The Action Item ticket count is not equal to 2 in a week. The Feedback task is not scheduled before 8:00 PM. The PPT task count is less then 4 in a week. Average tasks hours is smaller than 3 hours. The total task hours are smaller than to 36 for the week. Task count is smaller than 10. The task doesn't has description in it. Task hours are greater than 4. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM.",
    "Send Mail Message working Fine."
  );
}

(function () {
  const IS_GAS_ENV = typeof ScriptApp !== "undefined";
  if (!IS_GAS_ENV) runTests();
})();
