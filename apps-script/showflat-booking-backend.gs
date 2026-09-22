// Thomson Reserve — Showflat Booking Backend
//
// Paste into a Google Apps Script project bound to a Google Sheet, then
// Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
// Use the resulting /exec URL as APPS_SCRIPT_URL in showflat.html.
//
// GET  ?date=YYYY-MM-DD        -> { booked: ["10:00am – 11:00am", ...] }
// POST { date, time, name,     -> { success: true } or
//        mobile, email }          { success: false, message: "Slot already booked" }

const SHEET_NAME = 'Bookings';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Date', 'Time', 'Full Name', 'Mobile Number', 'Email', 'Booked At']);
  }
  return sheet;
}

function doGet(e) {
  var date = e.parameter.date;
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var booked = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === date) {
      booked.push(data[i][1]);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ booked: booked }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var body = JSON.parse(e.postData.contents);
    var date = body.date;
    var time = body.time;

    if (!date || !time) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Missing date or time' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === date && String(data[i][1]) === time) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Slot already booked' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    sheet.appendRow([date, time, body.name || '', body.mobile || '', body.email || '', new Date()]);
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
