/** Shared helpers for every 10RosterHR script (version 1.0).
 * Other files call HR10CFG only inside their functions, never while a file is
 * first loading, so the order of files in the Apps Script project does not matter.
 */
const HR10CFG = (() => {
  const ID = '1flssTRbZ74t1FFnoKlamhiSURcM-Byq5wraQHuI5TDw';
  const ROSTER_TABS = Object.freeze(['Master', 'SCC', 'ECC', 'STAR']);
  const PLAIN_TEXT_FORMAT = '@';
  // Text that Sheets would turn into a formula, number, date, time, percent,
  // currency or TRUE/FALSE if written as-is. Examples: =x, +1, -- note, 3/4,
  // 10-12, 00123, 5:06 PM, 50%, $5, (916), TRUE, Oct 5. A leading apostrophe
  // keeps it as text; Sheets hides the apostrophe.
  const COERCIBLE = /^[=+\-@']|^\s*[\d$.(]|^\s*(true|false)\s*$|^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d/i;

  /** Value as trimmed text. */
  function text(value) { return value == null ? '' : String(value).trim(); }

  /** Header/label comparison used everywhere: inner whitespace collapses to one space. */
  function normalize(value) { return String(value == null ? '' : value).replace(/\s+/g, ' ').trim(); }

  function blank(value) { return value === '' || value == null; }

  function same(a, b) {
    return a instanceof Date && b instanceof Date ? a.getTime() === b.getTime() : a === b;
  }

  /** Student Number as a digit string, '' for blank. Throws on anything else. */
  function studentId(value) {
    if (blank(value)) return '';
    if (typeof value === 'number' && (!Number.isSafeInteger(value) || value < 0)) {
      throw new Error('A Student Number is not a valid whole-number identifier: ' + value);
    }
    const key = String(value).trim();
    if (!key) return '';
    if (!/^\d+$/.test(key)) throw new Error('A Student Number contains unexpected characters: ' + key);
    return key;
  }

  /** Protects text from reinterpretation by setValues/appendRow. Cells formatted
   * as Plain text already keep text, so no apostrophe is added there. */
  function literal(value, numberFormat) {
    if (typeof value !== 'string' || value === '' || numberFormat === PLAIN_TEXT_FORMAT) return value;
    return COERCIBLE.test(value) ? "'" + value : value;
  }

  /** Values ready for range.setValues. formulaMask marks cells that are real formulas. */
  function prepareValues(range, values, formulaMask) {
    const formats = range.getNumberFormats();
    return values.map((row, r) => row.map((value, c) =>
      formulaMask && formulaMask[r] && formulaMask[r][c] ? value : literal(value, formats[r][c])));
  }

  function columnLetter(column) {
    let result = '';
    while (column) {
      result = String.fromCharCode(65 + (column - 1) % 26) + result;
      column = Math.floor((column - 1) / 26);
    }
    return result;
  }

  function workbook(ss) {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    if (!ss || ss.getId() !== ID) throw new Error('Run these tools only in 10RosterHR. No other file was changed.');
    return ss;
  }

  /** One document lock shared by every roster write. Never nest it. */
  function withLock(action, busyMessage) {
    const lock = LockService.getDocumentLock() || LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error(busyMessage || 'Roster tools are busy. Please retry.');
    try { return action(); }
    finally { lock.releaseLock(); }
  }

  /** Row-1 headers. columns maps header -> 1-based column number.
   * strict: a blank header inside the used width is an error. */
  function headerColumns(sheet, strict) {
    if (!sheet) throw new Error('A required tab is missing.');
    const width = sheet.getLastColumn();
    if (!width) throw new Error(sheet.getName() + ' has no headers.');
    const headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(normalize);
    const columns = new Map();
    headers.forEach((header, i) => {
      if (!header) {
        if (strict) throw new Error(sheet.getName() + ': a column header is blank.');
        return;
      }
      if (columns.has(header)) throw new Error(sheet.getName() + ': duplicate header ' + header + '.');
      columns.set(header, i + 1);
    });
    return {headers, columns, width};
  }

  /** 1-based column of one required header. */
  function findColumn(sheet, header) {
    const column = headerColumns(sheet).columns.get(normalize(header));
    if (!column) throw new Error(sheet.getName() + ': missing header ' + header + '.');
    return column;
  }

  /** Whole sheet as data. NOTE: columns here map header -> 0-based index into
   * each rows[] array (not a sheet column number). */
  function table(sheet, required) {
    const info = headerColumns(sheet);
    const columns = new Map();
    info.columns.forEach((column, header) => columns.set(header, column - 1));
    (required || []).forEach(header => {
      if (!columns.has(header)) throw new Error(sheet.getName() + ': missing header ' + header + '.');
    });
    const height = sheet.getLastRow();
    const rows = height > 1 ? sheet.getRange(2, 1, height - 1, info.width).getValues() : [];
    return {sheet, headers: info.headers, columns, rows, width: info.width};
  }

  /** Student Number from the selected row of a roster tab. */
  function studentOnRow(sheet, row) {
    const key = studentId(sheet.getRange(row, findColumn(sheet, 'Student Number')).getValue());
    if (!key) throw new Error('The selected row has no valid Student Number.');
    return key;
  }

  /** Read-only imports workbook named in Info!H2. */
  function sourceWorkbook(ss) {
    const info = ss.getSheetByName('Info');
    if (!info) throw new Error('Info tab is missing.');
    const url = text(info.getRange('H2').getValue());
    const match = url.match(/^https:\/\/docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]+)(?:[/?#].*)?$/);
    if (!match) throw new Error('Put the imports workbook Google Sheets URL in Info!H2.');
    if (match[1] === ID) throw new Error('Info!H2 must point to the separate imports workbook, not 10RosterHR.');
    try { return SpreadsheetApp.openById(match[1]); }
    catch (error) {
      throw new Error('Cannot read the imports workbook in Info!H2. Check the URL and this account\u2019s access. ' + error.message);
    }
  }

  /** Activity record in Extensions > Apps Script > Executions (replaces the
   * Automation Log tab). Never writes to the workbook, so it costs no sheet time.
   * Messages can include Student Numbers; Executions is visible only to project editors. */
  function log(level, action, studentNumber, message, details) {
    const entry = JSON.stringify({level: level || 'INFO', action, studentNumber: studentNumber || '', message,
      details: details || undefined});
    if (level === 'ERROR') console.error(entry); else console.log(entry);
  }
  function errorDetails(error) {
    return error ? (error.message || String(error)) + (error.stack ? '\n\n' + error.stack : '') : 'Unknown error.';
  }
  /** Stopwatch for the Executions log: lap('name') records ms since the last lap. */
  function timer() {
    const start = Date.now(), laps = {};
    let last = start;
    return {
      lap(name) { const now = Date.now(); laps[name] = now - last; last = now; },
      done() { laps.total = Date.now() - start; return laps; }
    };
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
  }

  return Object.freeze({ID, ROSTER_TABS, text, normalize, blank, same, studentId, literal, prepareValues,
    columnLetter, workbook, withLock, headerColumns, findColumn, table, studentOnRow, sourceWorkbook, escapeHtml,
    log, errorDetails, timer});
})();
