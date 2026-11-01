/**
 * Live RSVP list for the BUK Computer Science 2021 Reunion site.
 *
 * Paste into Extensions → Apps Script of a Google Sheet and deploy as a web app
 * (Execute as: Me, Who has access: Anyone). The script keeps RSVPs on a tab
 * named "RSVPs". The site reads the public list with GET and saves RSVPs with POST.
 *
 * Public responses never include phone numbers, and hidden RSVPs are only
 * returned as counts. To remove or hide someone, edit or delete their row.
 */

var SHEET_NAME = 'RSVPs';
var HEADERS = ['Updated', 'Name', 'WhatsApp', 'Status', 'Extra guests', 'Hidden'];

function doGet() {
  return json_(publicList_());
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'That request could not be read. Please try again.' });
  }

  // Bots tend to fill every field, including the invisible one.
  if (data.website) return json_(Object.assign({ ok: true }, publicList_()));

  var name = clean_(data.name).slice(0, 60);
  var phone = phoneKey_(data.phone);
  var status = data.status === 'maybe' ? 'maybe' : 'yes';
  var guests = Number(data.guests) === 1 ? 1 : 0;
  var hidden = data.hidden === true;

  if (name.length < 2) return json_({ ok: false, error: 'Please enter your full name.' });
  if (phone.length < 10 || phone.length > 15) return json_({ ok: false, error: 'Please enter a valid WhatsApp number.' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    var row = [new Date(), safeCell_(name), "'" + phone, status, guests, hidden ? 'Yes' : 'No'];
    var found = -1;
    for (var i = 1; i < values.length; i++) {
      if (phoneKey_(values[i][2]) === phone) { found = i + 1; break; }
    }
    if (found > 0) {
      sh.getRange(found, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
  } finally {
    lock.releaseLock();
  }
  return json_(Object.assign({ ok: true }, publicList_()));
}

function publicList_() {
  var values = sheet_().getDataRange().getValues().slice(1);
  var people = [];
  var hidden = { yes: 0, maybe: 0, guests: 0 };
  values.forEach(function (r) {
    var name = clean_(r[1]).replace(/^'/, '');
    if (!name) return;
    var status = String(r[3]).toLowerCase() === 'maybe' ? 'maybe' : 'yes';
    var guests = Number(r[4]) === 1 ? 1 : 0;
    var isHidden = /^(yes|true)$/i.test(String(r[5]).trim());
    if (isHidden) {
      hidden[status] += 1;
      if (status === 'yes') hidden.guests += guests;
    } else {
      people.push({ name: name, status: status, guests: guests });
    }
  });
  return { people: people, hidden: hidden };
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function clean_(v) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim();
}

// Stops names like "=HYPERLINK(...)" being run as spreadsheet formulas.
function safeCell_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function phoneKey_(v) {
  var d = String(v == null ? '' : v).replace(/\D/g, '');
  if (d.charAt(0) === '0') d = '234' + d.slice(1);
  return d;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
