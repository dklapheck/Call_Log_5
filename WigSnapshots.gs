/** 10RosterHR WIG snapshots v1.5. Replace the old WigSnapshots Script file.
 * @NotOnlyCurrentDoc
 * Requires Config.gs. Menus are built by Menus.gs.
 * v1.2: Failing Math / English / CTE cells hold the count, then one student per
 * line in the same cell: first initial, last name, Student Number.
 * v1.3: List of Students Chronically Absent uses the same one-per-line format.
 * A student who appears in 2, 3 or 4 of those four list cells in the same week
 * is shown in bold violet, pumpkin orange or red text. (Sheets cannot shade part of a
 * cell, so the student's line is colored instead of the cell background.)
 * v1.4: in Failing Math / English / CTE, the count is font size 10 and the
 * student lines are font size 7.
 * Master defines roster membership/names. Info!H2 points to the imports workbook.
 * Reads Upstream and CourseGrades directly. No Dash, StudentData, Config or Helpers.
 * Saves available metrics and lists missing students/data in column L.
 * Installation does not enable the optional Tuesday schedule.
 */
const HR10WIG = (() => {
  const METRICS = ['HR Students', 'Chronically Absent', 'List of Students Chronically Absent',
    'Failing Math', 'Failing English', 'Failing CTE Courses'];
  const SCHEDULE = 'scheduledWigSnapshot';
  const MISSING_COLUMN = 12, MISSING_HEADER = 'Missing Students / Data';
  // The four student-list cells compared for repeats, found by header.
  const LIST_HEADERS = ['List of Students Chronically Absent', 'Failing Math', 'Failing English', 'Failing CTE Courses'];
  // Text colors for a student appearing in 2, 3 or 4 list cells: violet, pumpkin, red.
  const REPEAT_COLORS = {2: '#8E24AA', 3: '#E06C00', 4: '#CC0000'};
  const STUDENT_LINE = /^.+, \d+$/; // A Aguirre, 9487734
  // Failing cells only: count line size, student line size.
  const SIZED_HEADERS = ['Failing Math', 'Failing English', 'Failing CTE Courses'];
  const COUNT_FONT_SIZE = 10, NAME_FONT_SIZE = 7;
  const text = v => HR10CFG.normalize(v);
  const blank = v => HR10CFG.blank(v);
  const id = v => HR10CFG.studentId(v);
  function bool(v) {
    const s = text(v).toUpperCase();
    if (['Y', 'YES', 'TRUE', '1'].includes(s)) return true;
    if (['N', 'NO', 'FALSE', '0'].includes(s)) return false;
    return null;
  }
  function number(v) {
    if (blank(v) || typeof v === 'boolean') return null;
    const n = Number(v);
    return Number.isSafeInteger(n) && n >= 0 ? n : null;
  }
  function workbook() {
    return HR10CFG.workbook(SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(HR10CFG.ID));
  }
  const lock = action => HR10CFG.withLock(action, 'Roster/WIG is busy. Retry the snapshot.');
  const table = (sheet, required) => HR10CFG.table(sheet, required);
  function roster(ss) {
    const t = table(ss.getSheetByName('Master'), ['Student Number', 'FIRST NAME', 'LAST NAME']);
    const students = new Map();
    t.rows.forEach((row, offset) => {
      const key = id(row[t.columns.get('Student Number')]);
      if (!key) {
        if (row.some(v => !blank(v))) throw new Error('Master row ' + (offset + 2) + ' has information but no Student Number.');
        return;
      }
      if (students.has(key)) throw new Error('Master has duplicate Student Number ' + key);
      const first = text(row[t.columns.get('FIRST NAME')]), last = text(row[t.columns.get('LAST NAME')]);
      const preferred = t.columns.has('Preferred Name') ? text(row[t.columns.get('Preferred Name')]) : '';
      const name = (preferred || (first + ' ' + last).trim()) || key;
      // Failing lists use the Subject Line style: A Aguirre, 9487734
      const short = [first.charAt(0).toUpperCase(), last].filter(Boolean).join(' ');
      students.set(key, {name, first, last, label: (short ? short + ', ' : '') + key});
    });
    if (!students.size) throw new Error('Master has no students.');
    return students;
  }
  function weekKey(now, tz) {
    const parts = Utilities.formatDate(now, tz, 'yyyy-MM-dd').split('-').map(Number);
    const day = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    day.setUTCDate(day.getUTCDate() - (day.getUTCDay() - 2 + 7) % 7);
    return day.toISOString().slice(0, 10); // Most recent Tuesday in spreadsheet timezone.
  }
  function dateKey(value, tz) {
    if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
    if (blank(value)) return '';
    const raw = text(value);
    const m = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s*\([^)]*\))?$/) || raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (!m) throw new Error('Unrecognized date in WIG column A: ' + raw);
    const iso = raw.includes('-') && !raw.includes('/');
    const y = +(iso ? m[1] : m[3]), mo = +(iso ? m[2] : m[1]), d = +(iso ? m[3] : m[2]);
    const result = new Date(Date.UTC(y, mo - 1, d));
    if (result.getUTCFullYear() !== y || result.getUTCMonth() !== mo - 1 || result.getUTCDate() !== d) throw new Error('Invalid WIG date: ' + raw);
    return result.toISOString().slice(0, 10);
  }
  function destination(ss, key) {
    const t = table(ss.getSheetByName('WIG'), METRICS), tz = ss.getSpreadsheetTimeZone();
    METRICS.forEach(h => { if (t.columns.get(h) === 0) throw new Error('Keep WIG week dates in column A; a metric header occupies column A.'); });
    if (METRICS.some(h => t.columns.get(h) + 1 === MISSING_COLUMN)) throw new Error('WIG column L is reserved for missing students/data; move the metric header out of L.');
    if (t.headers[MISSING_COLUMN - 1] !== MISSING_HEADER) throw new Error('Set WIG!L1 to ' + MISSING_HEADER + '.');
    const matches = [];
    t.rows.forEach((row, i) => { if (dateKey(row[0], tz) === key) matches.push(i + 2); });
    if (matches.length !== 1) throw new Error(matches.length ? 'WIG has duplicate week dates for ' + key + '.' : 'Add the date ' + key + ' to WIG column A first. No row was added automatically.');
    const row = matches[0];
    const occupied = METRICS.some(h => {
      const range = t.sheet.getRange(row, t.columns.get(h) + 1);
      return !blank(range.getValue()) || !!range.getFormula();
    });
    return {...t, row, occupied};
  }
  function source(ss) {
    const ssSource = HR10CFG.sourceWorkbook(ss);
    const u = table(ssSource.getSheetByName('Upstream'), ['STUDENT_NUMBER', 'CURRENT_ACTIVE_STUDENT',
      'PAR_CHRONICALLY_ABSENT', 'MATH_FAILING', 'MATH_CLASSES', 'ELA_FAILING', 'ELA_CLASSES']);
    const c = table(ssSource.getSheetByName('CourseGrades'), ['STUDENT_IDENTITY_ID', 'CANVAS_PASSING_FLAG']);
    if (!c.columns.has('HS_COURSE_OFFERINGS_IS_CTE_COURSE') && !c.columns.has('HS_COURSE_OFFERINGS_CONTENT_CATEGORY')) {
      throw new Error('CourseGrades needs a CTE classification header.');
    }
    const upstream = new Map(), courses = new Map();
    u.rows.forEach((row, i) => {
      const key = id(row[u.columns.get('STUDENT_NUMBER')]);
      if (!key) {
        if (row.some(v => !blank(v))) throw new Error('Upstream row ' + (i + 2) + ' has no STUDENT_NUMBER.');
        return;
      }
      if (upstream.has(key)) throw new Error('Upstream has duplicate STUDENT_NUMBER ' + key);
      upstream.set(key, row);
    });
    c.rows.forEach((row, i) => {
      const key = id(row[c.columns.get('STUDENT_IDENTITY_ID')]);
      if (!key) {
        if (row.some(v => !blank(v))) throw new Error('CourseGrades row ' + (i + 2) + ' has no STUDENT_IDENTITY_ID.');
        return;
      }
      if (!courses.has(key)) courses.set(key, []);
      courses.get(key).push(row);
    });
    return {u, c, upstream, courses};
  }
  function calculate(students, data, week, row) {
    const issues = [], chronic = [], missing = [], inactive = [], details = new Map();
    const issue = (label, reason) => {
      issues.push(label + ': ' + reason);
      if (!details.has(label)) details.set(label, []);
      details.get(label).push(reason);
    };
    const failing = {MATH: [], ELA: [], CTE: []};
    let matched = 0;
    for (const [key, student] of students) {
      const name = student.name, label = name + ' (' + key + ')', r = data.upstream.get(key);
      if (!r) { missing.push(label); issue(label, 'missing in Upstream'); }
      else {
        matched++;
        const get = h => r[data.u.columns.get(h)], active = bool(get('CURRENT_ACTIVE_STUDENT'));
        if (active !== true) {
          inactive.push(label); issue(label, 'source active status is ' + (active === false ? 'inactive' : 'missing/invalid') + '; check roster membership');
        }
        const absent = bool(get('PAR_CHRONICALLY_ABSENT'));
        if (absent === null) issue(label, 'chronic-absence flag is missing/invalid');
        else if (absent) chronic.push(student);
        for (const [prefix, title] of [['MATH', 'Math'], ['ELA', 'English']]) {
          const raw = get(prefix + '_FAILING'), count = number(raw), classes = number(get(prefix + '_CLASSES'));
          if (count !== null) {
            if (count > 0) failing[prefix].push(student);
          } else if (!blank(raw) || classes !== 0) issue(label, title + ' failing count is missing/invalid');
          // Blank failing count with explicit zero classes means not enrolled.
        }
      }
      const rows = data.courses.get(key);
      if (!rows?.length) { issue(label, 'no CourseGrades rows; CTE coverage cannot be verified'); continue; }
      let fails = false, unknown = false;
      rows.forEach(course => {
        const get = h => data.c.columns.has(h) ? course[data.c.columns.get(h)] : '';
        const flag = bool(get('HS_COURSE_OFFERINGS_IS_CTE_COURSE'));
        const category = text(get('HS_COURSE_OFFERINGS_CONTENT_CATEGORY')).toLowerCase();
        const cte = flag === true || category === 'cte courses';
        if (!cte) { if (flag === null && !category) unknown = true; return; }
        const passing = bool(get('CANVAS_PASSING_FLAG'));
        if (passing === false) fails = true;
        else if (passing === null) unknown = true;
      });
      if (fails) failing.CTE.push(student); // Each student once, even with several failing CTE courses.
      else if (unknown) issue(label, 'CTE classification/passing data is missing or invalid');
    }
    const sortedLabels = list => list.slice().sort((a, b) => a.last.localeCompare(b.last) ||
      a.first.localeCompare(b.first) || a.label.localeCompare(b.label)).map(s => s.label);
    // Count stays in Chronically Absent; the list cell holds one student per line.
    const list = sortedLabels(chronic).join('\n');
    if (list.length > 45000) throw new Error('The chronic-absence list is too long for one WIG cell.');
    const missingText = Array.from(details).sort(([a], [b]) => a.localeCompare(b))
      .map(([label, reasons]) => label + ': ' + reasons.join('; ')).join('\n');
    if (missingText.length > 45000) throw new Error('The missing-data list is too long for WIG column L.');
    // Count first, then one student per line inside the same cell. Zero stays a number.
    const failingCell = (list, title) => {
      if (!list.length) return 0;
      const lines = sortedLabels(list);
      const value = [String(list.length)].concat(lines).join('\n');
      if (value.length > 45000) throw new Error('The ' + title + ' student list is too long for one WIG cell.');
      return value;
    };
    return {week, row, ready: true, complete: issues.length === 0, students: students.size, matched, missing, inactive,
      issues, missingText, missingDataStudents: details.size, labels: METRICS.slice(),
      values: [students.size, chronic.length, list, failingCell(failing.MATH, 'Failing Math'),
        failingCell(failing.ELA, 'Failing English'), failingCell(failing.CTE, 'Failing CTE')],
      failingCounts: [failing.MATH.length, failing.ELA.length, failing.CTE.length]};
  }
  /** How many of the given cell texts list each student line. */
  function repeatCounts(texts) {
    const counts = new Map();
    texts.forEach(t => new Set(String(t).split('\n').map(line => line.trim()).filter(line => STUDENT_LINE.test(line)))
      .forEach(line => counts.set(line, (counts.get(line) || 0) + 1)));
    return counts;
  }
  function repeatColor(line, counts) { return REPEAT_COLORS[Math.min(counts.get(line.trim()) || 0, 4)] || ''; }
  function richText(text, counts, sized) {
    const builder = SpreadsheetApp.newRichTextValue().setText(text);
    let offset = 0;
    text.split('\n').forEach((line, i) => {
      const color = repeatColor(line, counts);
      const isCount = sized && i === 0 && /^\s*\d+\s*$/.test(line);
      const isStudent = STUDENT_LINE.test(line.trim());
      const size = sized ? (isCount ? COUNT_FONT_SIZE : isStudent ? NAME_FONT_SIZE : 0) : 0;
      if (line.length && (color || size)) {
        const style = SpreadsheetApp.newTextStyle();
        if (color) style.setForegroundColor(color).setBold(true);
        if (size) style.setFontSize(size);
        builder.setTextStyle(offset, offset + line.length, style.build());
      }
      offset += line.length + 1;
    });
    return builder.build();
  }
  /** Colors repeated students in one WIG row's four list cells and sizes the failing cells. Skips formulas,
   * numbers and cells without student lines. Returns the number of cells styled. */
  function colorRow(sheet, columns, row) {
    const cells = LIST_HEADERS.map(h => sheet.getRange(row, columns.get(h) + 1));
    const texts = cells.map(cell => {
      const value = cell.getValue();
      return typeof value === 'string' && !cell.getFormula() && value.split('\n').some(l => STUDENT_LINE.test(l.trim())) ? value : '';
    });
    const counts = repeatCounts(texts);
    let styled = 0;
    cells.forEach((cell, i) => {
      if (!texts[i]) return;
      cell.setRichTextValue(richText(texts[i], counts, SIZED_HEADERS.includes(LIST_HEADERS[i])));
      styled++;
    });
    return styled;
  }
  /** Re-applies repeat colors and font sizes to every saved WIG week (for example, after editing a list by hand). */
  function recolor() {
    const ss = workbook();
    const rows = lock(() => {
      const t = table(ss.getSheetByName('WIG'), LIST_HEADERS);
      let count = 0;
      t.rows.forEach((row, i) => { if (colorRow(t.sheet, t.columns, i + 2)) count++; });
      return count;
    });
    ss.toast('Repeat colors applied to ' + rows + ' WIG week' + (rows === 1 ? '' : 's') + '.', 'WIG', 6);
    return rows;
  }
  function save(wig, report, ss) {
    const cells = METRICS.map((h, i) => ({column: wig.columns.get(h) + 1, value: report.values[i]})).sort((a, b) => a.column - b.column);
    const note = 'WIG snapshot saved ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss z') +
      '. Week ' + report.week + '. Master roster ' + report.students + '; matched Upstream ' + report.matched +
      '. Metrics read from Info!H2 Upstream/CourseGrades at snapshot time.' +
      (report.complete ? ' Complete coverage.' : ' Known counts only; missing data for ' + report.missingDataStudents + ' students listed in column L.');
    const ops = [];
    for (let i = 0; i < cells.length;) {
      const first = cells[i], values = [first.value]; i++;
      while (i < cells.length && cells[i].column === first.column + values.length) values.push(cells[i++].value);
      const range = wig.sheet.getRange(wig.row, first.column, 1, values.length);
      const before = range.getValues(), notes = range.getNotes();
      if (range.getFormulas()[0].some(Boolean) || before[0].some(v => !blank(v))) throw new Error('WIG snapshot cells changed during calculation; existing values were preserved.');
      ops.push({range, before, notes, values: HR10CFG.prepareValues(range, [values]),
        wrap: values.some(v => typeof v === 'string' && v.includes('\n')),
        afterNotes: [notes[0].map(n => n ? n + '\n\n' + note : note)]});
    }
    const missingRange = wig.sheet.getRange(wig.row, MISSING_COLUMN);
    if (missingRange.getFormula()) throw new Error('WIG column L contains a formula on the target row; preserve it by moving it before saving.');
    const before = missingRange.getValues(), notes = missingRange.getNotes();
    const existing = blank(before[0][0]) ? '' : String(before[0][0]);
    const missingValue = existing + (existing && report.missingText ? '\n\n' : '') + report.missingText;
    if (missingValue.length > 45000) throw new Error('WIG column L text is too long; snapshot was not changed.');
    ops.push({range: missingRange, before, notes, values: HR10CFG.prepareValues(missingRange, [[missingValue]]), wrap: false,
      afterNotes: [[notes[0][0] ? notes[0][0] + '\n\n' + note : note]]});
    const attempted = [];
    try {
      ops.forEach(op => {
        attempted.push(op); op.range.setValues(op.values); op.range.setNotes(op.afterNotes);
        if (op.wrap) op.range.setWrap(true).setVerticalAlignment('top'); // Keep each student on its own line.
      });
      SpreadsheetApp.flush();
    } catch (e) {
      const errors = [];
      attempted.reverse().forEach(op => {
        try { op.range.setValues(HR10CFG.prepareValues(op.range, op.before)); } catch (x) { errors.push(x.message); }
        try { op.range.setNotes(op.notes); } catch (x) { errors.push(x.message); }
      });
      try { SpreadsheetApp.flush(); } catch (x) { errors.push(x.message); }
      throw new Error('WIG snapshot failed: ' + e.message + (errors.length ? '. Restoration failed; review WIG row ' + wig.row + '.' : '. Snapshot cells and notes were restored.'));
    }
    // Coloring is cosmetic: a failure here keeps the saved snapshot.
    try { colorRow(wig.sheet, wig.columns, wig.row); SpreadsheetApp.flush(); }
    catch (e) { console.error('WIG repeat colors could not be applied: ' + e.message); }
  }
  function log(level, message, detail) {
    HR10CFG.log(level, 'WigSnapshot', '', message, detail || '');
  }
  function display(report, title) {
    const esc = v => HR10CFG.escapeHtml(v);
    let html = '<div style="font:14px Arial"><p>Week ' + esc(report.week) + '; WIG row ' + report.row + '.</p>';
    if (report.skipped) html += '<p>' + esc(report.message) + '</p>';
    else {
      html += '<p>' + (report.saved ? 'Snapshot saved.' : 'Ready to save. Preview changes no cells.') +
        ' Master students: ' + report.students + '; matched Upstream: ' + report.matched + '.</p>';
      if (!report.complete) html += '<p>The following metrics are known counts from available data. Missing data does not pause the update; affected students are listed in column L.</p>';
      const listIndexes = LIST_HEADERS.map(h => report.labels.indexOf(h));
      const counts = repeatCounts(listIndexes.map(i => typeof report.values[i] === 'string' ? report.values[i] : ''));
      const cellHtml = (value, i) => !listIndexes.includes(i) || typeof value !== 'string' ? esc(value) :
        value.split('\n').map(line => {
          const color = repeatColor(line, counts);
          return color ? '<b style="color:' + color + '">' + esc(line) + '</b>' : esc(line);
        }).join('\n');
      html += '<table border="1" cellpadding="6" style="border-collapse:collapse">' + report.labels.map((h, i) => '<tr><th>' + esc(h) + '</th><td style="white-space:pre-wrap">' + cellHtml(report.values[i], i) + '</td></tr>').join('') + '</table>';
      html += '<p>Repeats across the four lists: <b style="color:' + REPEAT_COLORS[2] + '">2 lists</b>, <b style="color:' + REPEAT_COLORS[3] + '">3 lists</b>, <b style="color:' + REPEAT_COLORS[4] + '">4 lists</b>.</p>';
      html += '<p>Column L: ' + (report.missingText ? '</p><pre style="white-space:pre-wrap">' + esc(report.missingText) + '</pre>' : 'No missing students/data.</p>');
    }
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html + '</div>').setWidth(850).setHeight(600), title);
  }
  function run(preview, interactive) {
    const ss = workbook(), key = weekKey(new Date(), ss.getSpreadsheetTimeZone());
    let report;
    try {
      // Occupied weeks can skip without opening the external source workbook.
      const occupied = lock(() => destination(ss, key));
      if (occupied.occupied) report = {week: key, row: occupied.row, skipped: true, saved: false,
        message: 'WIG row ' + occupied.row + ' already contains snapshot data or formulas. It was preserved.'};
      else {
        const data = source(ss); // Read externally before taking the roster lock.
        report = lock(() => {
          const wig = destination(ss, key);
          if (wig.occupied) return {week: key, row: wig.row, skipped: true, saved: false, message: 'The week was filled by another edit; existing data was preserved.'};
          const result = calculate(roster(ss), data, key, wig.row);
          if (!preview) { save(wig, result, ss); result.saved = true; }
          return result;
        });
      }
      if (!preview) {
        if (report.skipped) log('INFO', report.message);
        else if (report.saved) log('SUCCESS', 'Saved WIG week ' + key + ', row ' + report.row + '; ' + report.students + ' Master students; missing data for ' + report.missingDataStudents + ' students listed in column L.');
      }
    } catch (e) {
      if (!preview) log('ERROR', e.message, e.stack || '');
      if (interactive && report && !report.skipped) {
        try { display(report, 'WIG snapshot needs review'); } catch (x) { console.error(x.message); }
      }
      throw e;
    }
    console.log(JSON.stringify({week: report.week, row: report.row, saved: !!report.saved, skipped: !!report.skipped,
      ready: !!report.ready, complete: report.complete, students: report.students, matched: report.matched,
      missingDataStudents: report.missingDataStudents, issues: report.issues?.length || 0}));
    if (interactive) {
      try { display(report, preview ? 'WIG snapshot preview' : 'WIG snapshot'); }
      catch (e) { console.error('Could not display WIG report: ' + e.message); }
    }
    return report;
  }
  function matching(handler, type) {
    return ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === handler && t.getEventType() === type);
  }
  function enable() {
    const ss = workbook(), tz = ss.getSpreadsheetTimeZone();
    lock(() => {
      const old = matching(SCHEDULE, ScriptApp.EventType.CLOCK);
      // Create successfully before removing old triggers so failure keeps the prior schedule.
      ScriptApp.newTrigger(SCHEDULE).timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(16).inTimezone(tz).create();
      old.forEach(t => ScriptApp.deleteTrigger(t));
    });
    ss.toast('Tuesday WIG snapshots enabled during the 4–5 PM hour (' + tz + ').', 'WIG', 8);
  }
  function disable() { workbook(); lock(() => matching(SCHEDULE, ScriptApp.EventType.CLOCK).forEach(t => ScriptApp.deleteTrigger(t))); }
  return Object.freeze({run, enable, disable, recolor});
})();
function previewWigSnapshot() { return HR10WIG.run(true, true); }
function saveWigSnapshot() { return HR10WIG.run(false, true); }
function scheduledWigSnapshot() { return HR10WIG.run(false, false); }
function snapshotWig_() { return HR10WIG.run(false, false).row; }
function enableTuesdayWigSnapshots() { return HR10WIG.enable(); }
function disableTuesdayWigSnapshots() { return HR10WIG.disable(); }
function recolorWigStudentRepeats() { return HR10WIG.recolor(); }
