// Paste this into the "Wedding guest addresses" sheet: Extensions > Apps Script.
// Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
// Each postcard sent from the questionnaire becomes a new row on the Responses tab.

const SHEET_NAME = 'Responses';

function doPost(e) {
  const p = (e && e.parameter) || {};

  // The honeypot field is hidden from people; bots fill it in.
  if (p._gotcha) return reply({ ok: true });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    sheet.appendRow([
      new Date(),
      clean(p.name),
      clean(p.partner_name),
      clean(p.address),
      clean(p.email),
      clean(p.lang)
    ]);
  } finally {
    lock.releaseLock();
  }
  return reply({ ok: true });
}

// Keeps a guest's text as plain text, so nothing typed can run as a formula.
function clean(v) {
  const s = String(v || '').trim().slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
