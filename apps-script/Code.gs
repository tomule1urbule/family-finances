/**
 * Family Finances 2026 — bridge between the phone app and the spreadsheet.
 *
 * Setup (once):
 *   1. script.google.com → New project (or in the sheet: Extensions → Apps Script).
 *   2. Replace everything in Code.gs with this file and press Save.
 *   3. Set KEY below to a long random value (the app asks for the same value).
 *   4. Deploy → New deployment → type "Web app".  Execute as: Me   Who has access: Anyone
 *      Authorize when asked, then copy the Web app URL (ends with /exec).
 *   5. Open the app, paste the URL and the KEY.
 *
 * Updating this script later: Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy.
 * That keeps the same URL, so the phones keep working.
 *
 * Anyone with BOTH the URL and the KEY can read and change the budget,
 * so share them only with the people who use the app.
 */
const KEY = 'CHANGE_ME_TO_A_LONG_RANDOM_VALUE';
const SHEET_ID = '13ecuoVtA3oRFgHIi3gHDH7msAtxfuUfhwV2EV57j7CA'; // Family Finances 2026
const TZ = 'Europe/Vilnius';
const PENDING = 'Laukiancios';
const PENDING_HEADERS = ['ID', 'Gauta', 'Data', 'Tipas', 'Kategorija', 'Suma', 'Komentarai', 'Vendorius', 'Pranesimas', 'Pastabos'];

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function doGet() {
  return json_({ ok: true, app: 'family-finances', hint: 'POST only' });
}

function doPost(e) {
  try {
    // JSON body from the app, or plain form/query parameters (e.g. from MacroDroid).
    let req = {};
    try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { req = {}; }
    const p = (e && e.parameter) || {};
    if (!req.key && p.key) req = { key: p.key, tool: p.tool || 'ingest_notification', input: { title: p.title || '', text: p.text || '', app: p.app || '' } };
    if (!KEY || KEY.indexOf('CHANGE_ME') === 0) return json_({ error: 'The KEY in Code.gs has not been set.', code: 'not_configured' });
    if (req.key !== KEY) return json_({ error: 'Wrong access key.', code: 'unauthorized' });
    const input = req.input || {};
    switch (req.tool) {
      case 'get_values':          return json_(getValues_(input.range));
      case 'update_values':       return locked_(() => updateValues_(input.range, input.values));
      case 'update_formulas':     return locked_(() => updateFormulas_(input.range, input.formulas));
      case 'append_tx':           return locked_(() => appendTx_(input.rows || []));
      case 'append_row':          return locked_(() => appendRow_(input.sheet, input.values));
      case 'ingest_notification': return locked_(() => ingest_(input));
      case 'approve_pending':     return locked_(() => approvePending_(input.id, input.row));
      case 'discard_pending':     return locked_(() => discardPending_(input.id));
      default:                    return json_({ error: 'Unknown operation: ' + req.tool, code: 'bad_request' });
    }
  } catch (err) {
    return json_({ error: String(err && err.message || err), code: 'tool_error' });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { const r = fn(); SpreadsheetApp.flush(); return json_(r); }
  finally { lock.releaseLock(); }
}
function range_(a1) { return ss_().getRange(a1); }

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
  rows.forEach((r, i) => writeTx_(sh, start + i, r));
  return { status: 'success', start: start };
}
function writeTx_(sh, row, r) {
  sh.getRange(row, 1, 1, 6).setValues([[r.serial, r.type, r.cat, r.amt, r.note || '', r.vendor || '']]);
  sh.getRange(row, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(row, 4).setNumberFormat('[$€]#,##0.00');
  if (r.periodOverride) sh.getRange(row, 7).setValue(r.periodOverride);
  else sh.getRange(row, 7).setFormula('=TEXT(IF(DAY(A' + row + ') < 10, A' + row + ' - DAY(A' + row + '), A' + row + '), "YYY, MMMM")');
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

// ---------- Pending transactions from Revolut notifications ----------
function pendingSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(PENDING);
  if (!sh) { sh = ss.insertSheet(PENDING); sh.getRange(1, 1, 1, PENDING_HEADERS.length).setValues([PENDING_HEADERS]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function todaySerial_() {
  const d = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd').split('-').map(Number);
  return Math.round((Date.UTC(d[0], d[1] - 1, d[2]) - Date.UTC(1899, 11, 30)) / 86400000);
}
// "€12.34", "12,34 €", "EUR 1,234.56", "-€5" → number; null if no amount in the text.
function parseAmount_(s) {
  const m = s.match(/(?:€|EUR)\s*-?\s*(\d[\d\s.,]*\d|\d)/i) || s.match(/(-?\d[\d\s.,]*\d|\d)\s*(?:€|EUR)/i);
  if (!m) return null;
  let t = m[1].replace(/\s/g, '');
  if (/,\d{2}$/.test(t)) t = t.replace(/\./g, '').replace(',', '.'); else t = t.replace(/,/g, '');
  const n = parseFloat(t);
  return isNaN(n) ? null : Math.abs(n);
}
// MacroDroid sends a variable name like [not_title] literally when it was typed instead of picked.
function cleanMagic_(s) { return String(s || '').replace(/\[[a-z_]+\]|\{[a-z_]+\}/gi, ' ').replace(/\s+/g, ' ').trim(); }
function guessVendor_(title, text) {
  // Revolut titles look like "Joint · Eurovaistinė" (account · merchant): take the part after the dot.
  const tp = String(title || '').split(/\s*[·•|]\s*/).map(x => x.trim()).filter(Boolean);
  const last = tp.length > 1 ? tp[tp.length - 1] : '';
  if (last && !/€|\d/.test(last) && !GENERIC_TITLE.test(last.replace(/[^\p{L}\s]/gu, '').trim())) return last;
  const m = text.match(/\b(?:at|pas|in|to|from|iš|is)\s+([^,.\n€]+?)(?:\s+(?:with|using|card|kortel|from|iš|is)\b|[,.\n!]|$)/i);
  if (m && m[1].trim().length > 1) return m[1].trim().replace(/^(UAB|AB|MB|VŠĮ|VSI|IĮ)\s+/i, '');
  const t = String(title || '').trim();
  if (t && !/^(revolut|joint|personal|bendr\S*)$/i.test(t) && !/€|\d/.test(t)) return t;
  return '';
}
function guessCategory_(vendor) {
  if (!vendor) return '';
  const sh = ss_().getSheetByName('Transakcijos');
  const n = lastRowInCol_(sh, 1); if (n < 2) return '';
  const v = sh.getRange(2, 3, n - 1, 4).getValues(); // C..F
  const key = vendor.toLowerCase(), count = {};
  v.forEach(r => { if (String(r[3]).trim().toLowerCase() === key && r[0]) count[r[0]] = (count[r[0]] || 0) + 1; });
  let best = '', bn = 0; Object.keys(count).forEach(c => { if (count[c] > bn) { best = c; bn = count[c]; } });
  return best;
}
// Revolut titles that describe the event, not the merchant ("Joint · Money sent").
const GENERIC_TITLE = /^(money sent|payment|card payment|transfer|pinigai i[sš]si[uų]sti|mok[eė]jimas|pervedimas)$/i;
// "Bumbis ❤️ has sent €20.35 to UAB AZETA VAISTINE from your joint account": the other joint-account
// holder paying someone. Returns the payer's name so it can go into the comment.
function jointPayer_(text) {
  const m = text.match(/^(.*?)\s+(?:has sent|sent|paid|spent)\s+(?:€|EUR)?\s*\d[\d.,]*\s*(?:€|EUR)?\s+(?:to|at)\s+.+\bjoint account/i);
  return m ? m[1].replace(/[^\p{L}\p{N}\s.'-]/gu, '').trim() : '';
}
// Only card payments / purchases become pending transactions. Everything else Revolut sends
// (referral bonuses, offers, transfers received, top-ups, savings moves) is skipped and logged.
const EXPENSE_WORDS = /\b(paid|spent|payment|purchase|charged|declined)\b|\bsent\s+(?:€|EUR)?\s*\d[\d.,]*\s*(?:€|EUR)?\s+to\s|apmok[eė]j|sumok[eė]j|i[sš]leid|mok[eė]jim|pirkin|nura[sš]yt/i;
const EXPENSE_SHAPE = /(?:€|EUR)\s*\d[\d.,]*\s+(?:at|pas)\s+\S|\d[\d.,]*\s*(?:€|EUR)\s+(?:at|pas)\s+\S/i;
const NOT_EXPENSE = /bonus|premij|referr|refer a|invite|friend|draug|pakviesk|earn|u[zž]dirb|reward|win\b|laim[eė]|offer|pasi[uū]lym|promo|akcij|up to|iki\s*(?:€|\d)|get\s*(?:€|\d)|gauk|gauti|received|gavai|gavote|gauta|sent you\s*(?:€|EUR|\d)|atsiunt|top.?up|papild|refund|gr[aą][zž]in|returned|savings|taupym|vault|instant access|exchanged|keitimas|pocket|declined|atmest/i;
function ingest_(input) {
  const title = cleanMagic_(input.title), text = cleanMagic_(input.text);
  const all = (title + ' ' + text).trim();
  if (!all) return logNotif_(all, 'skipped: empty');
  // Only the joint (family) account is tracked: personal-account spending is ignored.
  if (!/\bjoint\b/i.test(all)) return logNotif_(all, 'skipped: not joint account');
  if (NOT_EXPENSE.test(all)) return logNotif_(all, 'skipped: not an expense');
  if (!EXPENSE_WORDS.test(all) && !EXPENSE_SHAPE.test(all)) return logNotif_(all, 'skipped: no payment wording');
  const amt = parseAmount_(all);
  if (amt === null || amt === 0) return logNotif_(all, 'skipped: no amount');
  const sh = pendingSheet_();
  // MacroDroid can fire twice for one notification: skip an identical one from the last 3 minutes.
  const last = lastRowInCol_(sh, 1), now = new Date();
  if (last >= 2) {
    const recent = sh.getRange(Math.max(2, last - 10), 1, Math.min(11, last - 1), 9).getValues();
    if (recent.some(r => r[8] === all && r[1] instanceof Date && now - r[1] < 180000)) return logNotif_(all, 'skipped: duplicate notification');
  }
  const vendor = guessVendor_(title, text);
  const cat = guessCategory_(vendor);
  const id = 'p' + now.getTime().toString(36);
  const row = last + 1;
  const payer = jointPayer_(text);
  sh.getRange(row, 1, 1, 10).setValues([[id, now, todaySerial_(), 'Islaidos', cat, amt, payer ? 'Paid by ' + payer : '', vendor, all, '']]);
  sh.getRange(row, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(row, 3).setNumberFormat('yyyy-mm-dd');
  sh.getRange(row, 6).setNumberFormat('[$€]#,##0.00');
  logNotif_(all, 'added ' + amt + (vendor ? ' · ' + vendor : ''));
  return { status: 'success', id: id, amount: amt, vendor: vendor, category: cat };
}
// Keeps the last 200 Revolut notifications with what happened to them, so the filter can be tuned.
function logNotif_(text, result) {
  const ss = ss_(); let sh = ss.getSheetByName('Pranesimai');
  if (!sh) { sh = ss.insertSheet('Pranesimai'); sh.getRange(1, 1, 1, 3).setValues([['Gauta', 'Pranesimas', 'Rezultatas']]).setFontWeight('bold'); sh.setFrozenRows(1); }
  sh.insertRowAfter(1); sh.getRange(2, 1, 1, 3).setValues([[new Date(), text, result]]); sh.getRange(2, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  const n = sh.getLastRow(); if (n > 201) sh.deleteRows(202, n - 201);
  return { status: result.indexOf('skipped') === 0 ? 'skipped' : 'success', reason: result };
}
function findPending_(sh, id) {
  const last = lastRowInCol_(sh, 1); if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}
// Saves one pending transaction to Transakcijos and removes it from the pending tab, in one locked step.
function approvePending_(id, r) {
  const sh = pendingSheet_(), row = findPending_(sh, id);
  if (!row) return { status: 'gone', reason: 'Already approved or discarded on the other phone.' };
  const tx = ss_().getSheetByName('Transakcijos');
  const at = lastRowInCol_(tx, 1) + 1;
  writeTx_(tx, at, r);
  sh.deleteRow(row);
  return { status: 'success', start: at };
}
function discardPending_(id) {
  const sh = pendingSheet_(), row = findPending_(sh, id);
  if (row) sh.deleteRow(row);
  return { status: 'success' };
}

// Run this once from the editor (select testIngest → Run) to see a sample appear in the app.
function testIngest() {
  Logger.log(JSON.stringify(ingest_({ title: 'Maxima', text: 'Paid €12.34 at Maxima' })));
}
