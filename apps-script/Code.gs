// Receives the coach survey (docs/survey.html) and appends one row per answer
// to the first sheet of the spreadsheet this script is bound to.

var FIELDS = [
  'name', 'phone', 'instagram', 'city', 'teaching', 'needs', 'needs_other',
  'payment', 'pain', 'completion', 'medical', 'dm_time', 'tools', 'call', 'call_time'
];

var HEADERS = [
  'زمان', 'اسم', 'موبایل', 'اینستاگرام', 'شهر', 'نوع کلاس', 'هدف شاگردها', 'هدف دیگر',
  'روش دریافت پول', 'بزرگ‌ترین درد', 'تکمیل ترم از ۱۰ نفر', 'شاگرد با مشکل جسمی',
  'وقت روزانه برای پیام‌ها', 'ابزارهای پولی', 'تماس ۲۰ دقیقه‌ای', 'زمان تماس'
];

function doPost(e) {
  var p = (e && e.parameters) || {};
  // Honeypot: bots fill the hidden field, people don't.
  if (p.website && p.website[0]) return done();

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
    var row = [new Date()].concat(FIELDS.map(function (f) {
      return clean((p[f] || []).join('، '));
    }));
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
