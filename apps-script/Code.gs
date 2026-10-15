// Receives the coach survey (docs/survey.html) and appends one row per answer
// to the first sheet of the spreadsheet this script is bound to.
//
// Columns are not hardcoded: the form sends `_labels`, a list of
// [field, column title] pairs in the order the columns should appear.
// A missing column is added at the end of the header row, so the form's
// questions can change without redeploying this script.

var MAX_COLUMNS = 200;
var SKIP = { _labels: true, website: true };

function doPost(e) {
  var p = (e && e.parameters) || {};
  // Honeypot: bots fill the hidden field, people don't.
  if (p.website && p.website[0]) return done();

  var pairs = [];
  try {
    pairs = JSON.parse((p._labels || ['[]'])[0]);
  } catch (err) {
    pairs = [];
  }
  var titles = {};
  var order = [];
  (Array.isArray(pairs) ? pairs : []).forEach(function (pair) {
    if (!Array.isArray(pair) || pair.length !== 2) return;
    var key = String(pair[0]).slice(0, 60);
    if (titles[key]) return;
    titles[key] = String(pair[1]).slice(0, 120);
    order.push(key);
  });
  Object.keys(p).forEach(function (key) {
    if (!SKIP[key] && !titles[key]) {
      titles[key] = key.slice(0, 60);
      order.push(key);
    }
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var width = sheet.getLastColumn();
    var header = width ? sheet.getRange(1, 1, 1, width).getValues()[0].map(String) : [];
    if (header.length === 0) {
      header = ['زمان'];
      sheet.getRange(1, 1).setValue('زمان');
    }
    var row = header.map(function () { return ''; });
    row[0] = new Date();

    order.forEach(function (key) {
      var title = titles[key];
      var col = header.indexOf(title);
      if (col === -1) {
        if (header.length >= MAX_COLUMNS) return;
        header.push(title);
        row.push('');
        col = header.length - 1;
        sheet.getRange(1, col + 1).setValue(clean(title));
      }
      if (p[key]) row[col] = clean(p[key].join('، '));
    });

    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return done();
}

// Keeps values as plain text: stops formula injection and keeps the leading 0 of phone numbers.
function clean(v) {
  v = String(v).slice(0, 2000);
  return /^[=+\-@0-9]/.test(v) ? "'" + v : v;
}

function done() {
  return ContentService.createTextOutput('ok');
}
