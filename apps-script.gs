// Google Apps Script backend for Maryke's birthday site.
// 1. Create a Google Sheet. Extensions > Apps Script. Paste this file. Save.
// 2. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
// 3. Copy the web app URL into CONFIG.backendURL in index.html.
// RSVPs land in the "RSVP" tab, notes in "Notes". Delete a row in "Notes" to remove a note.

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var clip = function (s, n) { return String(s || '').slice(0, n); };
    if (d.type === 'rsvp') {
      sheet_('RSVP', ['Time', 'Name', 'Attending', 'Guests', 'Events', 'Diet', 'Note'])
        .appendRow([new Date(), clip(d.name, 60), clip(d.attending, 5), Number(d.guests) || 0,
                    (d.events || []).join(', '), clip(d.diet, 80), clip(d.note, 300)]);
    } else if (d.type === 'note') {
      sheet_('Notes', ['Time', 'Name', 'Text', 'Color'])
        .appendRow([new Date(), clip(d.name, 40), clip(d.text, 400), clip(d.color, 7)]);
    }
    return ContentService.createTextOutput('ok');
  } finally { lock.releaseLock(); }
}

function doGet(e) {
  var sh = sheet_('Notes', ['Time', 'Name', 'Text', 'Color']);
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
  var notes = rows.slice(-60).reverse().map(function (r) { return { name: r[1], text: r[2], color: r[3] }; });
  return ContentService.createTextOutput(JSON.stringify({ notes: notes }))
    .setMimeType(ContentService.MimeType.JSON);
}
