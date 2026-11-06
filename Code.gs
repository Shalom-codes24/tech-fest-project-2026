const RESPONSE_SHEET_NAME = "Registrations";
const RESPONSE_HEADERS = [
  "Submitted at",
  "Full name",
  "Email",
  "WhatsApp",
  "Attendee type",
  "Attendance",
  "Event updates",
];

function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("Open this script from the Google Sheet you want to use.");
  }

  let sheet = spreadsheet.getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(RESPONSE_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).setValues([RESPONSE_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).setFontWeight("bold");
  }
}

function doPost(event) {
  const parameters = event && event.parameter;
  if (!parameters) {
    return submissionResult(false);
  }

  const name = cleanValue(parameters.name, 120);
  const email = cleanValue(parameters.email, 254);
  const whatsapp = cleanValue(parameters.whatsapp, 40);
  const attendeeType = cleanValue(parameters.attendeeType, 80);
  const attendance = cleanValue(parameters.attendance, 30);
  const validAttendance = ["November 26", "November 27", "Both days"];

  if (
    !name ||
    !email ||
    !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
    !whatsapp ||
    !attendeeType ||
    validAttendance.indexOf(attendance) === -1
  ) {
    return submissionResult(false);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("No response spreadsheet is connected to this script.");
  }

  const sheet = spreadsheet.getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() === 0) {
    throw new Error("Run setup() before accepting registrations.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.appendRow([
      new Date(),
      safeCellValue(name),
      safeCellValue(email),
      safeCellValue(whatsapp),
      safeCellValue(attendeeType),
      attendance,
      parameters.updates === "Yes" ? "Yes" : "No",
    ]);
  } finally {
    lock.releaseLock();
  }

  return submissionResult(true);
}

function cleanValue(value, maxLength) {
  return String(value || "").trim().slice(0, maxLength);
}

function safeCellValue(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function submissionResult(success) {
  const message = JSON.stringify({
    type: "JTF_REGISTRATION_RESULT",
    success: success,
  });
  return HtmlService.createHtmlOutput(
    "<!doctype html><html><body><script>window.parent.postMessage(" +
      message +
      ', "*");</script></body></html>',
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
