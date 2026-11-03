/**
 * ============================================================
 *  IMEN · M.Sc. Semiconductor Engineering
 *  Marketability Study — Google Apps Script backend
 * ============================================================
 *
 *  SETUP
 *  1. Create a Google Sheet and copy its ID from the URL:
 *     https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit
 *  2. Paste it into SHEET_ID below.
 *  3. Set the dashboard password once, from the Apps Script editor:
 *     → run  setAdminPassword('YourStrongPassword')
 *  4. Deploy → New deployment → Web app
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. Copy the /exec URL into assets/data.js → APP_CONFIG.SCRIPT_URL
 *     and set DEMO_MODE to false.
 * ============================================================
 */

const SHEET_ID = '1Utai_Q0oAeJWkJeL9ux8SXCovXTUYProwbCMup-gUN8';
const SHEET_NAME = 'Responses_MSc_Semicond';
const ADMIN_KEY_PROPERTY = 'ADMIN_KEY_SHA256';

const HEADERS = [
  'timestamp',
  'email',
  'name',
  'language',
  'category',
  'category_other',
  'criteria',
  'study_mode',
  'fee_range',
  'fee_other',
  'interest',
  'feedback',
  'user_agent',
  // per-course suitability ratings (1-5), one per course card
  'core_1', 'core_2', 'core_3', 'core_4', 'core_5', 'core_6', 'core_7', 'core_8', 'core_9',
  'elective_1', 'elective_2', 'elective_3', 'elective_4', 'elective_5'
];

/* ------------------------------------------------------------
   GET — ping + token-gated dashboard data
------------------------------------------------------------ */
function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.action === 'ping') {
    return json({ ok: true, service: 'msc-semicond-survey', time: new Date().toISOString() });
  }

  if (params.action === 'data') {
    return dashboardData(params.k);
  }

  return json({ ok: false, error: 'Unknown action' });
}

/* ------------------------------------------------------------
   POST — survey submission
------------------------------------------------------------ */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({ ok: false, error: 'Empty request body' });
    }

    const body = JSON.parse(e.postData.contents);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);

    try {
      const sheet = getSheet();
      const row = HEADERS.map(function (key) {
        if (key === 'timestamp') return new Date();
        const value = body[key];
        return value === undefined || value === null ? '' : value;
      });
      sheet.appendRow(row);
    } finally {
      lock.releaseLock();
    }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* ------------------------------------------------------------
   Helpers
------------------------------------------------------------ */
function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  } else if (sheet.getLastColumn() < HEADERS.length) {
    // sheet created before the per-course rating columns existed: extend the header row
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  }
  return sheet;
}

function dashboardData(key) {
  const expected = PropertiesService.getScriptProperties().getProperty(ADMIN_KEY_PROPERTY);
  if (!expected) {
    return json({ ok: false, error: 'ADMIN_KEY_SHA256 is not configured. Run setAdminPassword() first.' });
  }
  if (!key || String(key).toLowerCase() !== String(expected).toLowerCase()) {
    return json({ ok: false, error: 'Unauthorized' });
  }

  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  if (!values.length) return json({ ok: true, count: 0, responses: [] });

  const headers = values.shift().map(String);
  const responses = values
    .filter(function (row) {
      return row.join('').trim() !== '';
    })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) {
        const v = row[i];
        obj[h] = v instanceof Date ? v.toISOString() : v;
      });
      return obj;
    });

  return json({ ok: true, count: responses.length, responses: responses });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------
   Run once to set / change the dashboard password.
   The plaintext password is never stored — only its SHA-256 hash.
------------------------------------------------------------ */
function setAdminPassword(password) {
  if (!password || String(password).length < 8) {
    throw new Error('Use a dashboard password of at least 8 characters.');
  }
  const hash = sha256Hex(password);
  PropertiesService.getScriptProperties().setProperty(ADMIN_KEY_PROPERTY, hash);
  return 'Admin password updated. Hash: ' + hash;
}

function sha256Hex(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8)
    .map(function (b) {
      return ('0' + (b & 0xff).toString(16)).slice(-2);
    })
    .join('');
}

/* ------------------------------------------------------------
   Optional: quick health check from the editor.
------------------------------------------------------------ */
function testSetup() {
  const sheet = getSheet();
  return {
    sheet: sheet.getName(),
    headers: HEADERS.length,
    adminKeyConfigured: !!PropertiesService.getScriptProperties().getProperty(ADMIN_KEY_PROPERTY)
  };
}
