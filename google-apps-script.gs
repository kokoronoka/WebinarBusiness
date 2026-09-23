/**
 * ZÉLL-V Masterclass — registration handler
 * Saves every registration from index.html into a Google Sheet.
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

// Tab inside the spreadsheet where rows are appended (created automatically).
var SHEET_NAME = 'Registrations';

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

    var sheet = getSheet();

    // Keep concurrent submissions from overwriting each other.
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      sheet.appendRow([
        new Date(),
        data.name || '',
        // Leading apostrophe keeps "+60…" as text instead of a broken formula.
        data.phone ? "'" + data.phone : '',
        data.email || '',
        data.country || '',
        data.profile || '',
        data.page || '',
        data.referrer || ''
      ]);
    } finally {
      lock.releaseLock();
    }

    notify(data);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// Lets you open the /exec URL in a browser to check the deployment is live.
function doGet() {
  return json({ ok: true, message: 'Registration endpoint is running.' });
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
