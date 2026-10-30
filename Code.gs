/**
 * STUDY LEDGER — Google Apps Script backend (Drive sync)
 * ------------------------------------------------------------------
 * Stores the whole app state ({ subjects, dailyLog }) in a Google Sheet
 * and serves it back to the Study Ledger page.
 *
 * HOW IT WORKS
 *  - GET  -> returns the stored JSON exactly as it was saved.
 *  - POST -> saves the JSON body the app sends.
 *  The app verifies every save by reading it back and comparing it
 *  character-for-character, so this script deliberately stores the raw
 *  text and never re-formats it.
 *  Google Sheets cells hold at most 50,000 characters, so the data is
 *  split into chunks, one chunk per row (column A).
 *
 * ---------------------------------------------------------------
 * ONE-TIME SETUP
 * ---------------------------------------------------------------
 * 1. Go to https://sheets.new and create a blank spreadsheet. Name it
 *    "Study Ledger Data" (any name works).
 * 2. In that spreadsheet: Extensions -> Apps Script.
 * 3. Delete everything in the editor, paste this whole file, and save.
 * 4. Click Deploy -> New deployment -> gear icon -> "Web app":
 *      Execute as:      Me
 *      Who has access:  Anyone
 *    Click Deploy and approve the permissions when asked
 *    (Advanced -> "Go to ... (unsafe)" is normal for your own script).
 * 5. Copy the Web app URL (ends in /exec) and paste it into
 *    Study Ledger -> Settings -> Drive sync. Use "Test connection".
 *
 * IF YOU EDIT THIS SCRIPT LATER
 *   Deploy -> Manage deployments -> pencil icon -> Version: "New version"
 *   -> Deploy. The URL stays the same.
 * ---------------------------------------------------------------
 */

var SHEET_NAME = 'StudyLedgerData';
var CHUNK_SIZE = 45000; // safely under the 50,000-character cell limit
var EMPTY_STATE = '{"subjects":[],"dailyLog":[]}';

function doGet(e) {
  try {
    return jsonOut_(readState_());
  } catch (err) {
    return jsonOut_(JSON.stringify({ error: 'Read failed: ' + err.message }));
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var body = (e && e.postData && e.postData.contents) ? e.postData.contents : '';

    // Reject anything that isn't the shape the app sends, so a stray
    // request can never wipe the sheet.
    var parsed;
    try {
      parsed = JSON.parse(body);
    } catch (parseErr) {
      return jsonOut_(JSON.stringify({ error: 'Body is not valid JSON.' }));
    }
    if (!parsed || !Array.isArray(parsed.subjects)) {
      return jsonOut_(JSON.stringify({ error: 'Body must contain a "subjects" array.' }));
    }

    writeState_(body);
    return jsonOut_(JSON.stringify({ ok: true }));
  } catch (err) {
    return jsonOut_(JSON.stringify({ error: 'Save failed: ' + err.message }));
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* ---------------------------- storage ---------------------------- */

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('This script must be created from inside a spreadsheet (Extensions -> Apps Script).');
  }
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1').setValue('Study Ledger data (do not edit by hand). Row 1 is a label; data starts in row 2.');
  }
  return sheet;
}

function readState_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return EMPTY_STATE;

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var text = '';
  for (var i = 0; i < values.length; i++) {
    text += values[i][0];
  }
  return text === '' ? EMPTY_STATE : text;
}

function writeState_(text) {
  var sheet = getSheet_();

  // Clear old data rows (keep the row-1 label).
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  }

  var chunks = [];
  for (var i = 0; i < text.length; i += CHUNK_SIZE) {
    chunks.push([text.substring(i, i + CHUNK_SIZE)]);
  }
  if (chunks.length === 0) chunks.push(['']);

  var range = sheet.getRange(2, 1, chunks.length, 1);
  // Plain-text format stops Sheets from "helpfully" converting anything
  // that looks like a number, date or formula, which would corrupt the JSON.
  range.setNumberFormat('@');
  range.setValues(chunks);
  SpreadsheetApp.flush();
}

/* ----------------------------- output ---------------------------- */

function jsonOut_(text) {
  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}
