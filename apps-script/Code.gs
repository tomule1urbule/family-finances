/**
 * Family Finances 2026 — bridge between the phone app and this spreadsheet.
 *
 * Setup (once):
 *   1. Go to script.google.com → New project (works on a phone in Chrome's Desktop site),
 *      or in the sheet: Extensions → Apps Script.
 *   2. Replace everything in Code.gs with this file and press Save.
 *   3. Set KEY below to a long random value (the app asks for the same value).
 *   4. Deploy → New deployment → type "Web app".
 *        Execute as: Me      Who has access: Anyone
 *      Authorize when asked, then copy the Web app URL (ends with /exec).
 *   5. Open the app, paste the URL and the KEY.
 *
 * Anyone with BOTH the URL and the KEY can read and change the budget,
 * so share them only with the people who use the app.
 */
const KEY = 'CHANGE_ME_TO_A_LONG_RANDOM_VALUE';

function doGet() {
  return json_({ ok: true, app: 'family-finances', hint: 'POST only' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!KEY || KEY.indexOf('CHANGE_ME') === 0) return json_({ error: 'The KEY in Code.gs has not been set.', code: 'not_configured' });
    if (req.key !== KEY) return json_({ error: 'Wrong access key.', code: 'unauthorized' });
    const input = req.input || {};
    switch (req.tool) {
      case 'get_values':     return json_(getValues_(input.range));
      case 'update_values':  return locked_(() => updateValues_(input.range, input.values));
      case 'update_formulas':return locked_(() => updateFormulas_(input.range, input.formulas));
      case 'append_tx':      return locked_(() => appendTx_(input.rows || []));
      case 'append_row':     return locked_(() => appendRow_(input.sheet, input.values));
      default:               return json_({ error: 'Unknown operation: ' + req.tool, code: 'bad_request' });
    }
  } catch (err) {
    return json_({ error: String(err && err.message || err), code: 'tool_error' });
  }
}

const SHEET_ID = '13ecuoVtA3oRFgHIi3gHDH7msAtxfuUfhwV2EV57j7CA'; // Family Finances 2026
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function locked_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try { const r = fn(); SpreadsheetApp.flush(); return json_(r); }
  finally { lock.releaseLock(); }
}
function range_(a1) {
  return ss_().getRange(a1);
}

// Same shape as the Google Sheets connector: formatted cell text, trailing empty rows/cells trimmed.
function getValues_(a1) {
  const vals = range_(a1).getDisplayValues();
  let last = vals.length - 1;
  while (last >= 0 && vals[last].every(v => v === '')) last--;
  const out = vals.slice(0, last + 1).map(r => { let n = r.length; while (n > 0 && r[n - 1] === '') n--; return r.slice(0, n); });
  return { range: a1, values: out };
}

// null cells are skipped (left untouched), so formulas between edited cells survive.
function updateValues_(a1, values) {
  const rg = range_(a1), sh = rg.getSheet(), r0 = rg.getRow(), c0 = rg.getColumn();
  let n = 0;
  values.forEach((row, i) => {
    let start = -1;
    for (let j = 0; j <= row.length; j++) {
      const skip = j === row.length || row[j] === null || row[j] === undefined;
      if (!skip && start < 0) start = j;
      if (skip && start >= 0) { sh.getRange(r0 + i, c0 + start, 1, j - start).setValues([row.slice(start, j)]); n += j - start; start = -1; }
    }
  });
  return { status: 'success', updatedCells: n };
}
function updateFormulas_(a1, formulas) {
  range_(a1).setFormulas(formulas);
  return { status: 'success' };
}

// Adds transactions after the last used row, all under one lock so two phones can't collide.
function appendTx_(rows) {
  const sh = ss_().getSheetByName('Transakcijos');
  const start = lastRowInCol_(sh, 1) + 1;
  rows.forEach((r, i) => {
    const row = start + i;
    sh.getRange(row, 1, 1, 6).setValues([[r.serial, r.type, r.cat, r.amt, r.note || '', r.vendor || '']]);
    sh.getRange(row, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(row, 4).setNumberFormat('[$€]#,##0.00');
    if (r.periodOverride) sh.getRange(row, 7).setValue(r.periodOverride);
    else sh.getRange(row, 7).setFormula('=TEXT(IF(DAY(A' + row + ') < 10, A' + row + ' - DAY(A' + row + '), A' + row + '), "YYY, MMMM")');
  });
  return { status: 'success', start: start };
}
function appendRow_(sheetName, values) {
  const sh = ss_().getSheetByName(sheetName);
  if (!sh) throw new Error('No tab named ' + sheetName);
  const row = lastRowInCol_(sh, 1) + 1;
  sh.getRange(row, 1, 1, values.length).setValues([values]);
  return { status: 'success', row: row };
}
function lastRowInCol_(sh, col) {
  const n = sh.getLastRow(); if (n === 0) return 0;
  const v = sh.getRange(1, col, n, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '' && v[i][0] !== null) return i + 1;
  return 0;
}
