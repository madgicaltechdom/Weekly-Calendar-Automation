if (typeof require !== "undefined") {
  MockData = require("./MockData.min.js");
}

let SendMailNodification = (function () {
  const _textMessageData =
    "The reading ticket count is Less then 5. The weekly schedule ticket count is not equal to 1 in the week. The Action Item ticket count is not equal to 2 in a week. The Feedback task is not scheduled before 8:00 PM. The Feedback task is not scheduled before 8:00 PM. The PPT task count is less then 4 in a week. Average tasks hours is smaller than 3 hours. The total task hours are smaller than to 36 for the week. Task count is smaller than 10. The task doesn't has description in it. The task doesn't has description in it. The task doesn't has description in it. The task doesn't has description in it. The task doesn't has description in it. The task doesn't has description in it. Task hours are greater than 4. The task doesn't has description in it. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM. The Video presentation is not scheduled before 6:00 PM.";
  const _sendMailId = "devendra.choukiker@medgicaltechdom.com";

  let mockData;
  if (typeof SpreadsheetApp === "undefined") {
    mockData = new MockData().addData("textMessageData", _textMessageData);
    mockData.addData("sendMailId", _sendMailId);
  }

  let _SendMailNodification = new WeakMap();
  class SendMailNodification {
    constructor() {
      if (mockData) {
        _SendMailNodification.set(this, mockData);
        return this;
      }
    }
    sendEmailNotification(textMsg, email) {
      if (mockData && textMsg === "" && email === "") {
        email = _SendMailNodification.get(this).getData("sendMailId");
        textMsg = _SendMailNodification.get(this).getData("textMessageData");
      }
      let sentences;
      if (textMsg && typeof textMsg === "string") {
        sentences = textMsg.split(". ");
      }
      let uniqueSentences = [...new Set(sentences)];
      let message = uniqueSentences.join(". ");
      let textSub = "Pivot Table Rule Alert ";
      return { email, textSub, message };
    }

    print() {
      console.log(JSON.stringify(this.get(row)));
      return true;
    }
  }
  return SendMailNodification;
})();
if (typeof module !== "undefined") module.exports = SendMailNodification;
