/** LogECC for 10RosterHR v2.0. Requires Config.gs, RosterSync 1.4, EccPowerSchool.
 * Menus are built by Menus.gs.
 *
 * ECC Contact History is one timeline, newest entry first. Conversations and
 * attempts both go there, each starting with a header line:
 *     -- 10/5/2026 [Conversation]        -- 10/2/2026 [Attempt]
 * A new entry is placed by its own date (ECC Date or ECC Attempt Date), so an
 * older attempt archived late still lands in the right spot. Old Attempts
 * Student Contact is no longer written; run merge10EccHistory once to move its
 * entries into the history (preview10EccHistoryMerge shows the result first).
 *
 * PowerSchool gets its own one-line version of the note (see EccPowerSchool).
 * Current notes and New Attempt Student Contact stay visible after archiving.
 * All roster writes go through update10RosterFields(planner), which reads Master
 * under the synchronization lock and writes all four rosters together.
 */
const HR10ECC = (() => {
  const TABS = ['Master', 'SCC', 'ECC', 'STAR'];
  const FIELDS = {
    date: 'ECC Date', history: 'ECC Contact History',
    notes: [
      ['Current ECC Notes Overall', 'Overall'],
      ['Current ECC Notes Classes', 'Classes'],
      ['Current ECC Notes Grades', 'Grades'],
      ['Current ECC Notes Socially', 'Socially'],
      ['Current ECC Notes Actions / Follow up', 'Actions / Follow up', 'Follow-up']
    ],
    attemptDate: 'ECC Attempt Date', attemptNote: 'New Attempt Student Contact',
    oldAttempts: 'Old Attempts Student Contact' // read only by the one-time merge
  };
  const CELL_LIMIT = 50000, CELL_WARNING = 45000;
  // "-- 9/14/2026 [Conversation]" and the older "-- 9/14/2026: Good..." both start entries.
  const HEADER_LINE = /^-- (\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/;
  const LEADING_DATE = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/;
  const text = value => HR10CFG.text(value);

  // ---- dates and history text -------------------------------------------
  function dateNumber(match) {
    let year = Number(match[3]);
    if (match[3].length === 2) year += 2000;
    const month = Number(match[1]), day = Number(match[2]);
    if (month < 1 || month > 12 || day < 1 || day > 31 || year < 2000 || year > 2100) return null;
    return year * 10000 + month * 100 + day;
  }
  function dateLabel(number) {
    return Math.floor(number / 100) % 100 + '/' + number % 100 + '/' + Math.floor(number / 10000);
  }
  function contains(history, entry) {
    return ('\n\n' + String(history || '').replace(/\r\n?/g, '\n').trim() + '\n\n')
      .includes('\n\n' + entry.trim() + '\n\n');
  }
  /** Puts entry above the first dated entry on or before its date (newest first). */
  function insertByDate(history, entry, entryDate) {
    const existing = String(history == null ? '' : history).replace(/\r\n?/g, '\n').trim();
    if (!existing) return entry;
    let offset = 0, insertAt = -1, anyHeader = false;
    for (const line of existing.split('\n')) {
      const match = line.match(HEADER_LINE), date = match && dateNumber(match);
      if (date) {
        anyHeader = true;
        if (date <= entryDate) { insertAt = offset; break; }
      }
      offset += line.length + 1;
    }
    if (insertAt === -1) return anyHeader ? existing + '\n\n' + entry : entry + '\n\n' + existing;
    const above = existing.slice(0, insertAt).trimEnd();
    return (above ? above + '\n\n' : '') + entry + '\n\n' + existing.slice(insertAt);
  }
  /** Splits a history or Old Attempts cell into entries for the one-time merge. */
  function entries(cell) {
    const all = String(cell == null ? '' : cell).replace(/\r\n?/g, '\n').trim();
    if (!all) return [];
    const lines = all.split('\n');
    if (lines.some(line => HEADER_LINE.test(line))) {
      const out = [];
      let current = null;
      lines.forEach(line => {
        if (HEADER_LINE.test(line)) { if (current) out.push(current); current = [line]; }
        else { if (!current) current = []; current.push(line); }
      });
      if (current) out.push(current);
      return out.map(group => group.join('\n').trim()).filter(Boolean);
    }
    // Free text: a paragraph that starts with a date begins a new entry.
    const out = [];
    all.split(/\n\s*\n/).map(p => p.trim()).filter(Boolean).forEach(paragraph => {
      if (!out.length || LEADING_DATE.test(paragraph)) out.push(paragraph);
      else out[out.length - 1] += '\n\n' + paragraph;
    });
    return out;
  }
  function entryDate(entry) {
    const header = entry.match(HEADER_LINE) || entry.match(LEADING_DATE);
    return header ? dateNumber(header) : null;
  }

  // ---- reading -----------------------------------------------------------
  function selectedStudent(ss) {
    const range = SpreadsheetApp.getActiveRange();
    if (!range || !TABS.includes(range.getSheet().getName()) ||
        range.getNumRows() !== 1 || range.getRow() <= 1) {
      throw new Error('Select one student row on Master, SCC, ECC, or STAR.');
    }
    const sheet = range.getSheet();
    if (sheet.getParent().getId() !== ss.getId()) throw new Error('Wrong workbook selection.');
    const key = HR10CFG.studentOnRow(sheet, range.getRow());
    return {key, sheet, row: range.getRow()};
  }
  /** Whole row in two calls; get(header) refuses formulas in editable fields. */
  function rowReader(sheet, row, columns) {
    const width = Math.max(...columns.values());
    const range = sheet.getRange(row, 1, 1, width);
    const values = range.getValues()[0], formulas = range.getFormulas()[0];
    return {
      get(header) {
        const column = columns.get(header);
        if (!column) throw new Error(sheet.getName() + ' is missing the field: ' + header);
        if (formulas[column - 1]) throw new Error('Use a value in the editable field: ' + header + ' (' + sheet.getName() + ').');
        return values[column - 1];
      }
    };
  }

  // ---- archive -----------------------------------------------------------
  function archive(type, openPowerSchool) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(), clock = HR10CFG.timer();
    let payload, archived = false, studentNumber = '';
    const action = openPowerSchool ? 'ECC Archive & Handoff' : 'ECC Archive';
    try {
      if (typeof update10RosterFields !== 'function') {
        throw new Error('Replace RosterSync with version 1.4 before using LogECC.');
      }
      const selection = selectedStudent(ss), key = selection.key;
      studentNumber = key;
      // Validate the optional handoff before changing any roster data.
      let settings;
      if (openPowerSchool) {
        if (typeof HR10ECCPS === 'undefined') throw new Error('Add EccPowerSchool.gs to use PowerSchool handoffs.');
        settings = HR10ECCPS.settings(ss, 'ECC ' + type);
      }
      clock.lap('prepare');
      update10RosterFields(master => {
        const attempt = type === 'Attempt';
        const masterRow = master.rows.get(key);
        if (!masterRow) throw new Error('Student ' + key + ' is not on Master.');
        const central = rowReader(master.sheet, masterRow, master.columns);
        const dateHeader = attempt ? FIELDS.attemptDate : FIELDS.date;
        const relevant = attempt ? [FIELDS.attemptDate, FIELDS.attemptNote, FIELDS.history] :
          [FIELDS.date, FIELDS.history, ...FIELDS.notes.map(pair => pair[0])];
        if (selection.sheet.getName() !== 'Master') {
          const local = rowReader(selection.sheet, selection.row, HR10CFG.headerColumns(selection.sheet).columns);
          relevant.forEach(header => {
            if (!HR10CFG.same(local.get(header), central.get(header))) {
              throw new Error('Selected tab does not yet match Master for ' + header +
                '. Wait for sync and try again; if it persists, check Apps Script Executions.');
            }
          });
        }
        const date = central.get(dateHeader);
        if (!(date instanceof Date) || isNaN(date.getTime())) {
          throw new Error('Enter a valid date in ' + dateHeader + '.');
        }
        const label = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'M/d/yyyy');
        const lines = ['-- ' + label + ' [' + type + ']'], powerSchoolLines = [];
        if (attempt) {
          const note = text(central.get(FIELDS.attemptNote));
          if (!note) throw new Error('Enter New Attempt Student Contact first.');
          lines.push(note); powerSchoolLines.push(note);
        } else {
          FIELDS.notes.forEach(([header, category, powerSchoolLabel]) => {
            const note = text(central.get(header));
            if (!note) return;
            lines.push(category + ': ' + note);
            powerSchoolLines.push((powerSchoolLabel || category) + ': ' + note);
          });
          if (lines.length === 1) throw new Error('Enter at least one Current ECC Note.');
        }
        const note = lines.join('\n');
        const history = central.get(FIELDS.history);
        const duplicate = contains(history, note);
        const updated = duplicate ? '' : insertByDate(history, note, dateNumber(label.match(/^(\d+)\/(\d+)\/(\d+)$/)));
        if (updated.length > CELL_LIMIT) {
          throw new Error('ECC Contact History for ' + key + ' would pass the 50,000-character cell limit. Move older entries elsewhere first.');
        }
        payload = {studentNumber: key, date: label, outcome: 'ECC ' + type, duplicate,
          note: openPowerSchool ? HR10ECCPS.withOpening(settings, powerSchoolLines.join('\n')) : note,
          nearLimit: updated.length > CELL_WARNING};
        return duplicate ? [] : [{studentNumber: key, header: FIELDS.history, value: updated}];
      });
      archived = true;
      clock.lap('save');
      // Open PowerSchool first so nothing else delays the dialog.
      if (openPowerSchool) { HR10ECCPS.open(payload, settings); clock.lap('dialog'); }
      ss.toast((payload.duplicate ? 'This ECC entry was already archived.' : 'ECC entry saved to Master, SCC, ECC, and STAR.') +
        (payload.nearLimit ? ' History is close to the cell size limit.' : ''), 'ECC Tools', 6);
      HR10CFG.log('INFO', action, key, payload.duplicate ? 'Entry already archived; no duplicate added.' :
        'Entry archived to all four rosters.', JSON.stringify({type, ms: clock.done()}));
      return payload;
    } catch (error) {
      const message = (archived ? 'Entry is saved; PowerSchool did not open. ' : 'ECC archive failed. ') + error.message;
      HR10CFG.log('ERROR', action, studentNumber, message, HR10CFG.errorDetails(error));
      ss.toast(message, 'ECC needs attention', 12);
      throw error;
    }
  }

  // ---- one-time merge of Old Attempts into the history --------------------
  /** Plans the merge for every student. Throws, writing nothing, if any tab
   * differs from Master in these two columns (unsynced edit) or a cell would overflow. */
  function mergePlan(master, ss) {
    if (!master.columns.has(FIELDS.history)) throw new Error('Missing header: ' + FIELDS.history);
    const hasOld = master.columns.has(FIELDS.oldAttempts);
    const headers = [FIELDS.history].concat(hasOld ? [FIELDS.oldAttempts] : []);
    const mirrors = TABS.slice(1).map(name => HR10CFG.table(ss.getSheetByName(name), ['Student Number'].concat(headers)));
    const central = HR10CFG.table(master.sheet, ['Student Number'].concat(headers));
    const rowOf = t => {
      const map = new Map(), idx = t.columns.get('Student Number');
      t.rows.forEach(row => { const key = HR10CFG.studentId(row[idx]); if (key) map.set(key, row); });
      return map;
    };
    const centralRows = rowOf(central), mirrorRows = mirrors.map(rowOf);
    const fields = [], students = [];
    for (const [key, row] of centralRows) {
      const value = header => row[central.columns.get(header)];
      mirrors.forEach((t, i) => headers.forEach(header => {
        const other = mirrorRows[i].get(key);
        if (!other || !HR10CFG.same(other[t.columns.get(header)], value(header))) {
          throw new Error(t.sheet.getName() + ' does not match Master for ' + key + ' in ' + header +
            '. Wait for sync (or run Check roster sync) and try again. Nothing was changed.');
        }
      }));
      const items = [];
      entries(value(FIELDS.history)).forEach((entry, i) => items.push({entry, date: entryDate(entry), source: 0, i}));
      let moved = 0;
      if (hasOld) entries(value(FIELDS.oldAttempts)).forEach((entry, i) => {
        const date = entryDate(entry);
        // Give Old Attempts entries the standard header; their own text is kept unchanged.
        const withHeader = HEADER_LINE.test(entry) ? entry :
          (date ? '-- ' + dateLabel(date) + ' [Attempt]' : '-- [Attempt, no date]') + '\n' + entry;
        if (!items.some(item => item.entry === withHeader || item.entry === entry)) {
          items.push({entry: withHeader, date, source: 1, i}); moved++;
        }
      });
      // Newest first; same day: history before attempts, later-written first; undated last.
      items.sort((a, b) => (a.date === null) - (b.date === null) || (b.date || 0) - (a.date || 0) ||
        a.source - b.source || (a.date === null ? a.i - b.i : b.i - a.i));
      const merged = items.map(item => item.entry).join('\n\n');
      const currentHistory = String(value(FIELDS.history) == null ? '' : value(FIELDS.history));
      const oldText = hasOld ? text(value(FIELDS.oldAttempts)) : '';
      if (merged === currentHistory && !oldText) continue;
      if (merged.length > CELL_LIMIT) throw new Error('Merged history for ' + key + ' is over the 50,000-character cell limit. Nothing was changed.');
      fields.push({studentNumber: key, header: FIELDS.history, value: merged});
      if (oldText) fields.push({studentNumber: key, header: FIELDS.oldAttempts, value: ''});
      students.push({key, entries: items.length, moved, undated: items.filter(item => item.date === null).length,
        top: items.length ? items[0].entry.split('\n')[0] : ''});
    }
    return {fields, students, hasOld};
  }
  function merge(preview) {
    const ss = HR10CFG.workbook();
    let report;
    update10RosterFields(master => {
      report = mergePlan(master, ss);
      return {fields: report.fields, bulk: true, dryRun: preview};
    });
    const esc = HR10CFG.escapeHtml;
    const undated = report.students.filter(s => s.undated);
    const html = '<div style="font:13px Arial,sans-serif">' +
      '<p><b>' + (preview ? 'Preview only. Nothing was changed.' : 'Merged and re-sorted on all four rosters.') + '</b></p>' +
      '<p>' + report.students.length + ' student histories ' + (preview ? 'would change' : 'changed') + '; ' +
      report.students.reduce((n, s) => n + s.moved, 0) + ' Old Attempts entries ' + (preview ? 'would move' : 'moved') +
      ' into ECC Contact History' + (report.hasOld ? ' and Old Attempts Student Contact ' + (preview ? 'would be' : 'was') + ' cleared for them.' : '.') + '</p>' +
      (undated.length ? '<p style="color:#b45f06">' + undated.length + ' student(s) have entries with no date at the start; they are kept at the bottom of the history: ' +
        undated.map(s => esc(s.key)).join(', ') + '</p>' : '') +
      '<table border="1" cellpadding="4" style="border-collapse:collapse"><tr><th>Student</th><th>Entries</th><th>Moved</th><th>Newest entry</th></tr>' +
      report.students.map(s => '<tr><td>' + esc(s.key) + '</td><td>' + s.entries + '</td><td>' + s.moved + '</td><td>' + esc(s.top) + '</td></tr>').join('') +
      '</table></div>';
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(640).setHeight(460),
      preview ? 'ECC history merge preview' : 'ECC history merge done');
    if (!preview) HR10CFG.log('INFO', 'ECC History Merge', '', report.students.length + ' histories merged.');
    return report;
  }

  /** Checks all four header/student sets and every LogECC field without writing. */
  function validate() {
    if (typeof update10RosterFields !== 'function') throw new Error('Replace RosterSync with version 1.4 first.');
    update10RosterFields(master => {
      [FIELDS.date, FIELDS.history, FIELDS.attemptDate, FIELDS.attemptNote,
        ...FIELDS.notes.map(pair => pair[0])].forEach(header => {
        if (!master.columns.has(header)) throw new Error('Missing LogECC header: ' + header);
      });
      return [];
    });
  }
  return Object.freeze({archive, validate, merge});
})();

function archiveCurrentECCNote() { return HR10ECC.archive('Conversation', false); }
function logCurrentEccRowAndOpenPowerSchool() { return HR10ECC.archive('Conversation', true); }
function archiveAndOpenPowerSchoolECC() { return logCurrentEccRowAndOpenPowerSchool(); }
function archiveCurrentECCAttempt() { return HR10ECC.archive('Attempt', false); }
function logCurrentEccAttemptAndOpenPowerSchool() { return HR10ECC.archive('Attempt', true); }
/** One-time move of Old Attempts Student Contact into ECC Contact History. */
function preview10EccHistoryMerge() { return HR10ECC.merge(true); }
function merge10EccHistory() { return HR10ECC.merge(false); }
