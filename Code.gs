/**
 * Best of Post-Acute Care Awards – Nomination form handler
 *
 * Paste this into your Google Sheet: Extensions > Apps Script.
 * Then run setup() once, and deploy as a Web app.
 */

// Optional: who gets an email for each new nomination. Leave as "" to turn emails off.
// You can list more than one address, separated by commas.
const NOTIFY_EMAIL = "";

const SHEET_NAME = "Nominations";

// Each column in the sheet: [header shown in the sheet, field name from the form]
const COLUMNS = [
  ["Submitted At", null],
  ["Category", "category"],
  ["Nomination Type", "nominationType"],
  ["Nominee Organization", "nomineeOrg"],
  ["Nominator Name", "name"],
  ["Nominator Title", "title"],
  ["Nominator Email", "email"],
  ["Nominator Organization", "nominatorOrg"],
  ["Why They Deserve Recognition", "why"],
  ["Example of Impact", "impact"],
  ["Anything Else", "anythingElse"],
  ["Status", null],
];

/** Run this once from the editor (select "setup" and click Run). It creates the tab and headers. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const headers = COLUMNS.map(c => c[0]);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(9, 3, 350);
}

/** Runs every time someone submits the form. */
function doPost(e) {
  const p = e.parameter;

  // Bots fill in the hidden "website" field; real people never see it
  if (p.website) return ok();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const row = COLUMNS.map(([header, field]) => {
      if (header === "Submitted At") return new Date();
      if (header === "Status") return "New";
      // A leading ' stops Sheets from treating answers like "=..." as formulas
      return field && p[field] ? "'" + String(p[field]).trim() : "";
    });
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_EMAIL) {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: "New Best of Post-Acute Care nomination: " + (p.nomineeOrg || "Unknown") + " – " + (p.category || ""),
      body:
        "Category: " + p.category + "\n" +
        "Nominee: " + p.nomineeOrg + " (" + p.nominationType + " nomination)\n" +
        "Nominated by: " + p.name + ", " + p.title + " at " + p.nominatorOrg + " (" + p.email + ")\n\n" +
        "Why:\n" + p.why + "\n\n" +
        "Open the sheet: " + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    });
  }

  return ok();
}

function ok() {
  return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
