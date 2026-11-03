/**
 * ZÉLL-V Masterclass — registration handler
 * Saves every registration from index.html into a Google Sheet — into TWO tabs:
 *
 *   "Registrations"          The original record. Leave it untouched.
 *   "Copy of Registrations"  Your working copy. Add remark / status / follow-up columns,
 *                            insert or reorder columns, highlight rows — all safe.
 *
 * New rows in the working copy are placed by matching the column HEADINGS, so as long
 * as you don't rename the 8 original headings (Timestamp, Full name, WhatsApp, Email,
 * Country / city, Profile, Source page, Referrer), every registration lands in the right
 * column. Your extra columns are left blank on new rows for you to fill in.
 *
 * ── SETUP (about 5 minutes) ─────────────────────────────────────────────
 * 1. Create a new Google Sheet (sheets.new). Name it e.g. "Masterclass Registrations".
 * 2. In that sheet: Extensions ▸ Apps Script. Delete the sample code.
 * 3. Paste this whole file in. Change SHARED_SECRET below to your own random string.
 * 4. Save, then Deploy ▸ New deployment ▸ type "Web app".
 *       Execute as:        Me
 *       Who has access:    Anyone                 ← must be "Anyone", not "Anyone with Google account"
 *    Click Deploy and authorise when prompted (choose your account ▸ Advanced ▸ Go to … ▸ Allow).
 * 5. Copy the Web app URL (ends in /exec).
 * 6. In index.html, set:
 *       var FORM_ENDPOINT = 'https://script.google.com/macros/s/XXXXXXXX/exec';
 *       var FORM_SECRET   = 'the same string you set below';
 *
 * After editing this script later, you must Deploy ▸ Manage deployments ▸ edit ▸
 * Version: New version ▸ Deploy, or the live URL keeps running the old code.
 * ────────────────────────────────────────────────────────────────────────
 */

// Must match FORM_SECRET in index.html. Change it to any random string.
var SHARED_SECRET = 'change-me-to-something-random';

// The original record — rows are appended here in fixed column order. Don't edit this tab.
var SHEET_NAME = 'Registrations';

// The working copy — rows are placed by heading, so you can add/insert columns freely.
// Must match the tab name EXACTLY (Google names a duplicated tab "Copy of Registrations").
// If this tab doesn't exist yet, the script creates it as a full duplicate of the original.
// Set to '' to switch the working copy off.
var COPY_SHEET_NAME = 'Copy of Registrations';

var COLUMNS = [
  'Timestamp',
  'Full name',
  'WhatsApp',
  'Email',
  'Country / city',
  'Profile',
  'Source page',
  'Referrer'
];

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({ ok: false, error: 'empty request' });
    }

    var data = JSON.parse(e.postData.contents);

    if (SHARED_SECRET && data.secret !== SHARED_SECRET) {
      return json({ ok: false, error: 'bad secret' });
    }

    var row = buildRow(data);

    // Keep concurrent submissions from overwriting each other.
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      // Make sure the working copy exists BEFORE writing: if it has to be created as a
      // duplicate of the original, that must happen before this row is added, or the
      // new registration would end up in the copy twice.
      var copy = null;
      try {
        copy = getWorkingCopy();
      } catch (copyErr) {
        console.error('Working copy unavailable: ' + copyErr);
      }

      // 1) Original — the source of truth. If this fails, the request fails.
      getSheet().appendRow(row);

      // 2) Working copy — best effort. A problem here must never lose a registration,
      //    because the original above already holds it.
      if (copy) {
        try {
          appendToWorkingCopy(copy, row);
        } catch (copyErr) {
          console.error('Working copy not updated: ' + copyErr);
        }
      }
    } finally {
      lock.releaseLock();
    }

    notify(data);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// One registration, in COLUMNS order.
function buildRow(data) {
  return [
    new Date(),
    data.name || '',
    // Leading apostrophe keeps "+60…" as text instead of a broken formula.
    data.phone ? "'" + data.phone : '',
    data.email || '',
    data.country || '',
    data.profile || '',
    data.page || '',
    data.referrer || ''
  ];
}

// Writes each value under its matching heading in the working copy, on the next empty row.
// Only the 8 registration cells are written — your own columns on that row are left alone.
function appendToWorkingCopy(sheet, row) {
  var colOf = headingColumns(sheet);
  var target = sheet.getLastRow() + 1;

  COLUMNS.forEach(function (heading, i) {
    var col = colOf[normalise(heading)];
    if (col) sheet.getRange(target, col).setValue(row[i]);
  });
}

// Map of normalised heading text -> 1-based column number, from row 1.
function headingColumns(sheet) {
  var lastCol = sheet.getLastColumn();
  var map = {};
  if (lastCol === 0) return map;
  sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].forEach(function (h, i) {
    var key = normalise(h);
    if (key && !map[key]) map[key] = i + 1;
  });
  return map;
}

function normalise(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

// Open the /exec URL in a browser to check the deployment — it reports which spreadsheet
// and tabs this script writes to, how many rows each holds, and any headings the working
// copy is missing (a renamed heading means that value can't be placed there).
function doGet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return json({
        ok: false,
        error: 'This script is not bound to a spreadsheet. Open your Google Sheet and use ' +
               'Extensions > Apps Script (do not create a standalone script at script.google.com).'
      });
    }

    var sheet = getSheet();
    var report = {
      ok: true,
      message: 'Registration endpoint is running.',
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      allTabs: ss.getSheets().map(function (s) { return s.getName(); }),
      original: {
        tab: SHEET_NAME,
        rowsStored: Math.max(0, sheet.getLastRow() - 1),
        lastRow: sheet.getLastRow() > 1
          ? sheet.getRange(sheet.getLastRow(), 1, 1, COLUMNS.length).getDisplayValues()[0]
          : null
      }
    };

    var copy = COPY_SHEET_NAME ? ss.getSheetByName(COPY_SHEET_NAME) : null;
    if (copy) {
      var colOf = headingColumns(copy);
      report.workingCopy = {
        tab: COPY_SHEET_NAME,
        rowsStored: Math.max(0, copy.getLastRow() - 1),
        missingHeadings: COLUMNS.filter(function (h) { return !colOf[normalise(h)]; })
      };
    } else if (COPY_SHEET_NAME) {
      report.workingCopy = 'No tab named "' + COPY_SHEET_NAME + '" yet — it will be created ' +
        'as a duplicate of "' + SHEET_NAME + '" on the next registration.';
    } else {
      report.workingCopy = 'switched off (COPY_SHEET_NAME is empty)';
    }

    return json(report);
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(4, 220);
  }

  return sheet;
}

// Returns the working-copy tab, creating it as a duplicate of the original if it's missing.
function getWorkingCopy() {
  if (!COPY_SHEET_NAME) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var copy = ss.getSheetByName(COPY_SHEET_NAME);

  if (!copy) {
    copy = getSheet().copyTo(ss).setName(COPY_SHEET_NAME);
  }

  return copy;
}

/**
 * Optional: email yourself on each registration.
 * Put your address in NOTIFY_EMAIL to switch it on.
 */
var NOTIFY_EMAIL = '';

function notify(data) {
  if (!NOTIFY_EMAIL) return;
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New masterclass registration: ' + (data.name || 'unknown'),
      body: [
        'Name:      ' + (data.name || ''),
        'WhatsApp:  ' + (data.phone || ''),
        'Email:     ' + (data.email || ''),
        'Country:   ' + (data.country || ''),
        'Profile:   ' + (data.profile || '')
      ].join('\n')
    });
  } catch (err) {
    // Never let a mail failure lose the row.
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
