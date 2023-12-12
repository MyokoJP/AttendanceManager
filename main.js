const sheetID = '1Kt6YjXmluf7QgYlA1ZCpg-c-lrHcACoX7YceDPE_toI';
const spreadSheet = SpreadsheetApp.openById(sheetID);

function main(e) {
  //const res = e.response.getItemResponses();
  const res = FormApp.getActiveForm().getResponses()[5].getItemResponses();

  const date = Utilities.parseDate(res[0].getResponse(), "JST", "yyyy-MM-dd");
  const part = res[1].getResponse();

  let attendance = {};
  for (let i = 2; i < res.length; i++) {
    attendance[res[i].getItem().getTitle()] = res[i].getResponse();
  }

  const sheet = new Attendance(spreadSheet, "12月");
  if (!sheet.exists()) {
    sheet.initialize();
  }

  /** GASのDate型の月は, 0から始まる(1月の場合は0)になるので、+1をする */
  if (!sheet.getMonth(date.getMonth() + 1)) {
    sheet.insertMonth(date.getMonth() + 1);
  }
  if (!sheet.getDate(date)) {
    sheet.insertDate(date);
  }

  sheet.setAttendance(part, attendance, date);
}
