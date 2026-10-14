/**
 * RSVP backend · Birthday Thanksgiving Studio Session (Sun 18 Oct 2026, 3:30 PM)
 *
 * Lives inside the RSVP Google Sheet (Extensions → Apps Script).
 * Deploy → New deployment → Web app → Execute as: Me · Who has access: Anyone.
 * Paste the /exec URL into RSVP_ENDPOINT in index.html.
 *
 * - Saves each RSVP as a row on the "RSVPs" tab.
 * - The same WhatsApp number updates its existing row instead of adding a new one.
 * - Caps total attendance (guests + their plus-ones) at CAPACITY; overflow can join a waitlist.
 * - Records each guest's category (Family, Friends, Musicians, Special Guest).
 * - Emails the sheet owner on every response.
 */

var CONFIG = {
  CAPACITY: 35,        // total people allowed in the studio
  MAX_GUESTS: 0,       // invitations are personal: no plus-ones
  SHEET_NAME: 'RSVPs',
  SUMMARY_NAME: 'Summary',
  NOTIFY_EMAIL: true,  // email the owner each time someone responds
  EVENT_NAME: 'Birthday Thanksgiving Studio Session'
};

var CATEGORIES = ['Family', 'Friends', 'Musicians', 'Special Guest'];
var HEADERS = ['Submitted', 'Last updated', 'Name', 'Category', 'WhatsApp', 'Email', 'Response', 'Guests', 'Party size', 'Memory / wish', 'Phone key'];
var COL = { submitted: 1, updated: 2, name: 3, category: 4, phone: 5, email: 6, response: 7, guests: 8, party: 9, note: 10, key: 11 };
var STATUS = { yes: 'Attending', no: 'Not attending', wait: 'Waitlist' };

/* ------------------------------------------------------------------ */
/* Web endpoints                                                       */
/* ------------------------------------------------------------------ */

function doGet(e) {
  try {
    var sheet = getSheet_();
    var confirmed = confirmedCount_(readRows_(sheet), -1);
    return json_({ ok: true, capacity: CONFIG.CAPACITY, confirmed: confirmed, remaining: Math.max(0, CONFIG.CAPACITY - confirmed) });
  } catch (err) {
    return json_({ ok: false, code: 'server', message: 'Could not load seat availability.' });
  }
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, code: 'bad_request', message: 'We could not read your RSVP. Please try again.' });
  }

  // Spam trap: bots fill the hidden "website" field. Pretend success, save nothing.
  if (data.website) return json_({ ok: true, status: STATUS.no, partySize: 0, updated: false });

  var v = validate_(data);
  if (!v.ok) return json_(v);
  var r = v.value;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return json_({ ok: false, code: 'busy', message: 'Lots of RSVPs coming in right now. Please try again in a moment.' });
  }

  try {
    var sheet = getSheet_();
    var rows = readRows_(sheet);
    var existing = findByKey_(rows, r.key);              // index into rows, or -1
    var confirmedOthers = confirmedCount_(rows, existing);
    var remainingBefore = CONFIG.CAPACITY - confirmedOthers;

    var status, guests, party;
    if (r.response === 'yes') {
      guests = r.guests;
      party = 1 + guests;
      var heldSeats = existing >= 0 && rows[existing].response === STATUS.yes ? rows[existing].party : 0;
      if (party <= remainingBefore) {
        status = STATUS.yes;
      } else if (r.waitlist && heldSeats === 0) {
        status = STATUS.wait;
      } else {
        // Someone who already holds seats never gets bumped to the waitlist by an edit.
        return json_({ ok: false, code: 'capacity', remaining: Math.max(0, remainingBefore), held: heldSeats, message: 'Not enough seats left for your party.' });
      }
    } else {
      status = STATUS.no;
      guests = 0;
      party = 0;
    }

    var now = new Date();
    var record = [
      existing >= 0 ? rows[existing].submitted : now,
      now,
      safe_(r.name),
      r.category,
      "'" + r.phone,                       // keep as text so "+234…" / "0803…" display as typed
      safe_(r.email),
      status,
      status === STATUS.no ? 0 : guests,
      status === STATUS.yes ? party : 0,  // only confirmed seats count toward the cap
      safe_(r.note),
      "'" + r.key
    ];

    if (existing >= 0) {
      sheet.getRange(rows[existing].rowNumber, 1, 1, HEADERS.length).setValues([record]);
    } else {
      sheet.appendRow(record);
    }

    var confirmedAfter = confirmedOthers + (status === STATUS.yes ? party : 0);
    var remainingAfter = Math.max(0, CONFIG.CAPACITY - confirmedAfter);

    notify_(r, status, party, existing >= 0, confirmedAfter, remainingAfter);

    return json_({ ok: true, status: status, partySize: status === STATUS.yes ? party : 0, updated: existing >= 0, remaining: remainingAfter });
  } catch (err) {
    return json_({ ok: false, code: 'server', message: 'Something went wrong saving your RSVP. Please try again.' });
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function validate_(d) {
  var response = d.response === 'yes' || d.response === 'no' ? d.response : '';
  var name = String(d.name || '').trim().replace(/\s+/g, ' ').slice(0, 100);
  var phone = String(d.phone || '').trim().slice(0, 24);
  var email = String(d.email || '').trim().slice(0, 120);
  var note = String(d.note || '').trim().slice(0, 600);
  var category = CATEGORIES.indexOf(d.category) >= 0 ? d.category : '';  // older cached pages send none
  var guests = parseInt(d.guests, 10);
  if (isNaN(guests) || guests < 0) guests = 0;
  if (guests > CONFIG.MAX_GUESTS) guests = CONFIG.MAX_GUESTS;

  if (!response) return { ok: false, code: 'invalid', message: "Let us know if you'll attend." };
  if (name.length < 2) return { ok: false, code: 'invalid', message: 'Please enter your name.' };
  var key = phoneKey_(phone);
  if (key.length < 7) return { ok: false, code: 'invalid', message: 'Please enter a valid WhatsApp number.' };
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, code: 'invalid', message: "That email address doesn't look right." };

  return { ok: true, value: { response: response, name: name, phone: phone, email: email, note: note, guests: guests, waitlist: !!d.waitlist, key: key, category: category } };
}

/** Same person, however they type their number: +234 803…, 234803…, 0803… all match. */
function phoneKey_(phone) {
  var d = String(phone || '').replace(/\D/g, '');
  if (d.indexOf('00') === 0) d = d.slice(2);                   // 00234… → 234…
  if (d.indexOf('234') === 0 && d.length >= 13) d = d.slice(3); // 234803… → 803…
  if (d.charAt(0) === '0' && d.length === 11) d = d.slice(1);   // 0803… → 803…
  return d;
}

/** Stop names/notes like "=HYPERLINK(...)" from being run as formulas in the sheet. */
function safe_(s) {
  s = String(s || '');
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0
      ? ss.getSheets()[0].setName(CONFIG.SHEET_NAME)
      : ss.insertSheet(CONFIG.SHEET_NAME, 0);
  }
  if (sheet.getLastRow() > 0) migrate_(ss, sheet);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#4d0a0b').setFontColor('#f3d47e');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(COL.name, 200);
    sheet.setColumnWidth(COL.note, 320);
    sheet.getRange(2, COL.submitted, sheet.getMaxRows() - 1, 2).setNumberFormat('ddd d mmm, h:mm am/pm');
    sheet.hideColumns(COL.key);
    buildSummary_(ss);
  }
  return sheet;
}

/** Sheets created before the Category column: insert it after Name. Data and Summary formulas shift with it. */
function migrate_(ss, sheet) {
  var header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 4)).getValues()[0];
  if (header[COL.category - 1] === 'Category' || header[COL.name - 1] !== 'Name') return;
  sheet.insertColumnAfter(COL.name);
  sheet.getRange(1, COL.category).setValue('Category');
  sheet.setColumnWidth(COL.category, 130);
  addCategorySummary_(ss);
}

function buildSummary_(ss) {
  if (ss.getSheetByName(CONFIG.SUMMARY_NAME)) return;
  var s = ss.insertSheet(CONFIG.SUMMARY_NAME);
  var n = CONFIG.SHEET_NAME;
  s.getRange('A1:B7').setValues([
    ['Capacity', CONFIG.CAPACITY],
    ['Confirmed headcount', '=SUMIF(' + n + '!G:G,"Attending",' + n + '!I:I)'],
    ['Seats remaining', '=MAX(0,B1-B2)'],
    ['RSVPs attending', '=COUNTIF(' + n + '!G:G,"Attending")'],
    ['Declined', '=COUNTIF(' + n + '!G:G,"Not attending")'],
    ['Waitlist', '=COUNTIF(' + n + '!G:G,"Waitlist")'],
    ['Waitlist headcount', '=SUMIF(' + n + '!G:G,"Waitlist",' + n + '!H:H)+B6']
  ]);
  s.getRange('A1:A7').setFontWeight('bold');
  s.setColumnWidth(1, 200);
  addCategorySummary_(ss);
}

/** Confirmed headcount per category (plus-ones counted with the guest who brings them). */
function addCategorySummary_(ss) {
  var s = ss.getSheetByName(CONFIG.SUMMARY_NAME);
  if (!s || s.getRange('A9').getValue() === 'Headcount by category') return;
  var n = CONFIG.SHEET_NAME;
  var rows = [['Headcount by category', 'Attending']];
  CATEGORIES.concat(['(none given)']).forEach(function (c) {
    var crit = c === '(none given)' ? '""' : '"' + c + '"';
    rows.push([c, '=SUMIFS(' + n + '!I:I,' + n + '!G:G,"Attending",' + n + '!D:D,' + crit + ')']);
  });
  s.getRange(9, 1, rows.length, 2).setValues(rows);
  s.getRange('A9:B9').setFontWeight('bold');
}

function readRows_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  return values.map(function (v, i) {
    return {
      rowNumber: i + 2,
      submitted: v[COL.submitted - 1],
      response: v[COL.response - 1],
      party: Number(v[COL.party - 1]) || 0,
      key: String(v[COL.key - 1]).replace(/^'/, '')
    };
  });
}

function findByKey_(rows, key) {
  for (var i = 0; i < rows.length; i++) if (rows[i].key === key) return i;
  return -1;
}

function confirmedCount_(rows, skipIndex) {
  var total = 0;
  for (var i = 0; i < rows.length; i++) {
    if (i === skipIndex) continue;
    if (rows[i].response === STATUS.yes) total += rows[i].party;
  }
  return total;
}

function notify_(r, status, party, updated, confirmed, remaining) {
  if (!CONFIG.NOTIFY_EMAIL) return;
  try {
    var to = Session.getEffectiveUser().getEmail();
    if (!to) return;
    var what = status === STATUS.yes ? 'is coming' + (party > 1 ? ' (party of ' + party + ')' : '')
             : status === STATUS.wait ? 'joined the waitlist' + (party > 1 ? ' (party of ' + party + ')' : '')
             : "can't make it";
    var subject = (updated ? 'RSVP updated: ' : 'RSVP: ') + r.name + ' ' + what;
    var body = [
      r.name + ' ' + what + '.',
      '',
      r.category ? 'Category: ' + r.category : null,
      'WhatsApp: ' + r.phone,
      r.email ? 'Email: ' + r.email : null,
      r.note ? 'Message: ' + r.note : null,
      '',
      'Headcount: ' + confirmed + ' of ' + CONFIG.CAPACITY + ' · ' + remaining + ' seats left',
      'Sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    ].filter(function (x) { return x !== null; }).join('\n');
    MailApp.sendEmail(to, subject, body, { name: 'RSVP · ' + CONFIG.EVENT_NAME });
  } catch (err) {
    // Email quota or permission issue; the RSVP is already saved.
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor to create the tabs and approve permissions. */
function setup() {
  getSheet_();
}
