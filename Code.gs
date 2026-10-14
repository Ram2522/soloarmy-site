// Solo Army waitlist store. Bound to a Google Sheet named "Solo Army waitlist".
// Deploy as web app: execute as me, access anyone.
function doPost(e) {
  var out = function (o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); };
  try {
    var body = JSON.parse(e.postData.contents);
    var email = String(body.email || "").trim().toLowerCase();
    if (body.website) return out({ ok: true });
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || email.length > 200) return out({ ok: false });
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var last = sh.getLastRow();
    var seen = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
    if (seen.indexOf(email) < 0) sh.appendRow([email, new Date()]);
    return out({ ok: true, count: Math.max(sh.getLastRow() - 1, 0) });
  } catch (err) { return out({ ok: false }); }
}
function doGet() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  return ContentService.createTextOutput(JSON.stringify({ count: Math.max(sh.getLastRow() - 1, 0) })).setMimeType(ContentService.MimeType.JSON);
}
