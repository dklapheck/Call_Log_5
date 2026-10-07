/** 10RosterHR WIG snapshots v1.9. Replace the old WigSnapshots Script file.
 * v1.6: Failing Math / English / CTE come from CourseGrades, one row per course.
 * Credit-recovery courses (HS_COURSE_OFFERINGS_IS_CR_COURSE = Yes) are not counted.
 * Each failing line ends with the grade: "J Garcia, 11304174 — 37%"; CTE lines name
 * the course first: "A Aguirre, 9487734 — CS CAR017A Bus Marketing Expl CA 45%".
 * Snapshots no longer add a cell note.
 * v1.7: chronic-absence lines end with days needed and last activity:
 * "J Garcia, 11304174 — needs 4, active 10/5". The List of Students Chronically
 * Absent cell gets a note with each student's full attendance details.
 * v1.8: no dash before the details ("J Garcia, 11304174 59%"), and CTE course
 * titles are shortened ("CAR015A Arts, AV, Comm, Principles" -> "Arts AV Comm").
 * Add any title you want written differently to COURSE_SHORT.
 * v1.9: Failing Math / English use Upstream MATH_PASSING_FLAG / ELA_PASSING_FLAG
 * (0 = failing) and show MATH_GRADE / ELA_GRADE. Failing CTE still comes from
 * CourseGrades, skipping credit-recovery courses.
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
  // A Aguirre, 9487734   or   A Aguirre, 9487734 45%   (group 1 is the student key).
  // Older weeks written as "A Aguirre, 9487734 — 45%" still match.
  const STUDENT_LINE = /^(.+?, \d+)(?:\s.*)?$/;
  // Exact course titles -> what to show. Checked before the automatic shortening.
  const COURSE_SHORT = {
    // 'CS CAR017A Bus Marketing Expl CA': 'Bus Marketing',
  };
  const COURSE_WORDS = [[/\bBusiness\b/g, 'Bus'], [/\bManagement\b/g, 'Mgmt'], [/\bTechnology\b/g, 'Tech'],
    [/\bInformation\b/g, 'Info'], [/\band\b/g, '&']];
  /** "CAR015A Arts, AV, Comm, Principles" -> "Arts AV Comm"; "CS CAR019A PBL Healthcare Expl CA" -> "PBL Healthcare". */
  function shortCourse(title) {
    const full = HR10CFG.normalize(title);
    if (COURSE_SHORT[full]) return COURSE_SHORT[full];
    let s = full.replace(/^CS\s+/i, '')                       // district prefix
      .replace(/^[A-Z]{2,4}\d{2,4}[A-Z]{0,3}\d?\s+/, '')       // course code (CAR015A, SCI330A)
      .replace(/\s+(CA|CR)$/, '')                               // trailing program tag
      .replace(/,/g, ' ')
      .replace(/\b(Principles of|Principles|Introduction to|Intro to|Exploration|Expl)\b/gi, ' ')
      .replace(/\s+[AB]$/, '');                                 // semester segment
    COURSE_WORDS.forEach(([pattern, word]) => { s = s.replace(pattern, word); });
    s = s.replace(/\s+/g, ' ').trim();
    return s || full;
  }
  const lineKey = line => { const m = String(line).trim().match(STUDENT_LINE); return m ? m[1] : ''; };
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
  /** Short date "10/5" from a Sheets date or M/D/YYYY / YYYY-MM-DD text; '' if blank. */
  function shortDate(v, tz) {
    if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, tz, 'M/d');
    const s = text(v);
    let m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/\d{2,4})?$/);
    if (m) return +m[1] + '/' + +m[2];
    m = s.match(/^\d{4}-(\d{1,2})-(\d{1,2})/);
    return m ? +m[1] + '/' + +m[2] : s;
  }
  const num = v => blank(v) || typeof v === 'boolean' || !Number.isFinite(Number(v)) ? null : Number(v);
  /** Chronic-absence line detail and full note for one student. get(header) reads Upstream. */
  function attendance(student, get, tz) {
    const needs = num(get('DAYS_NEEDED_TO_BECOME_COMPLIANT'));
    const active = shortDate(get('DAR_MOST_RECENT_ACTIVITY_DATE'), tz);
    const detail = [needs === null ? '' : 'needs ' + needs, active ? 'active ' + active : ''].filter(Boolean).join(', ');
    const pct = num(get('PAR_ATTENDANCE_PCT')), present = num(get('PRESENT_DAYS')), enrolled = num(get('ENROLLED_DAYS'));
    const absent = num(get('TOTAL_ABSENT_DAYS')) ?? num(get('ABSENT_DAYS')), excused = num(get('EXCUSED_DAYS'));
    const last5 = num(get('MISSING_LAST_5_SCHOOL_DAYS')), prev5 = num(get('MISSING_5_TO_10_SCHOOL_DAYS_AGO'));
    const lessons = num(get('DAR_TOTAL_LESSONS_ATTEMPTED_LAST_5_DAYS'));
    const lines = [student.label];
    if (pct !== null) lines.push('Attendance ' + pct + '%' + (present !== null && enrolled !== null ? ' (' + present + ' of ' + enrolled + ' days)' : ''));
    if (absent !== null) lines.push('Absent ' + absent + (excused !== null ? ' (' + excused + ' excused)' : ''));
    if (last5 !== null || prev5 !== null) lines.push('Missed last 5 days: ' + (last5 ?? '?') + '; 5–10 days ago: ' + (prev5 ?? '?'));
    if (!blank(get('MISSING_ATTENDANCE_DATES'))) lines.push('Missing dates: ' + text(get('MISSING_ATTENDANCE_DATES')));
    if (needs !== null) lines.push('Needs ' + needs + (needs === 1 ? ' day' : ' days') + ' to become compliant');
    if (active || lessons !== null) lines.push('Last activity ' + (active || '?') + (lessons !== null ? '; lessons attempted last 5 days: ' + lessons : ''));
    for (const n of ['NC1', 'NC2']) {
      const date = shortDate(get(n + '_MOST_RECENT'), tz), count = num(get(n + '_NUMBEROFLETTERS'));
      if (date || count) lines.push(n + ' letter ' + (date || '?') + (count ? ' (' + count + ' sent)' : ''));
    }
    const ge = shortDate(get('GE_CONNECTION_CALL'), tz), due = shortDate(get('GE_NEXT_CALL_DUE'), tz);
    if (ge || due) lines.push('GE connection ' + (ge || '?') + (due ? '; next call due ' + due : ''));
    return {student, detail, note: lines.join('\n')};
  }
  /** Course grade as "37%": posted grade, else unposted; '' if neither is a number. */
  function percent(posted, unposted) {
    for (const v of [posted, unposted]) {
      if (blank(v) || typeof v === 'boolean') continue;
      const n = Number(v);
      if (Number.isFinite(n)) return Math.round(n) + '%';
    }
    return '';
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
      'PAR_CHRONICALLY_ABSENT', 'MATH_GRADE', 'MATH_PASSING_FLAG', 'ELA_GRADE', 'ELA_PASSING_FLAG']);
    const c = table(ssSource.getSheetByName('CourseGrades'), ['STUDENT_IDENTITY_ID', 'CANVAS_PASSING_FLAG',
      'COURSE_NAME', 'HS_COURSE_OFFERINGS_CONTENT_CATEGORY', 'HS_COURSE_OFFERINGS_IS_CR_COURSE']);
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
    return {u, c, upstream, courses, tz: ss.getSpreadsheetTimeZone()};
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
        else if (absent) chronic.push(attendance(student, get, data.tz));
        for (const [prefix, title] of [['MATH', 'Math'], ['ELA', 'English']]) {
          const passing = bool(get(prefix + '_PASSING_FLAG'));
          if (passing === false) failing[prefix].push({student, detail: percent(get(prefix + '_GRADE'))});
          else if (passing === null && number(get(prefix + '_CLASSES')) !== 0) issue(label, title + ' passing flag is missing/invalid');
          // A blank flag with zero classes means the student is not enrolled in that subject.
        }
      }
      const rows = data.courses.get(key);
      if (!rows?.length) { issue(label, 'no CourseGrades rows; CTE cannot be verified'); continue; }
      const fails = {CTE: []}, unknown = new Set();
      rows.forEach(course => {
        const get = h => data.c.columns.has(h) ? course[data.c.columns.get(h)] : '';
        if (bool(get('HS_COURSE_OFFERINGS_IS_CR_COURSE')) === true) return; // Credit recovery is not counted.
        const category = text(get('HS_COURSE_OFFERINGS_CONTENT_CATEGORY')).toLowerCase();
        const cteFlag = bool(get('HS_COURSE_OFFERINGS_IS_CTE_COURSE'));
        const bucket = cteFlag === true || category === 'cte courses' ? 'CTE' : '';
        if (!bucket) return;
        const passing = bool(get('CANVAS_PASSING_FLAG'));
        if (passing === null) { unknown.add(bucket); return; }
        if (passing) return;
        fails[bucket].push({course: text(get('COURSE_NAME')),
          pct: percent(get('CANVAS_POSTED_COURSE_GRADE'), get('CANVAS_UNPOSTED_COURSE_GRADE'))});
      });
      fails.CTE.length && failing.CTE.push({student, detail: fails.CTE.map(f => (shortCourse(f.course) + ' ' + f.pct).trim()).join('; ')});
      if (unknown.size && !fails.CTE.length) issue(label, 'CTE passing data is missing');
    }
    const order = (a, b) => a.last.localeCompare(b.last) || a.first.localeCompare(b.first) || a.label.localeCompare(b.label);
    const sortedLabels = list => list.slice().sort(order).map(s => s.label);
    const sortedLines = list => list.slice().sort((a, b) => order(a.student, b.student))
      .map(e => e.student.label + (e.detail ? ' ' + e.detail : ''));
    // Count stays in Chronically Absent; the list cell holds one student per line.
    const list = sortedLines(chronic).join('\n');
    const chronicNote = chronic.slice().sort((a, b) => order(a.student, b.student)).map(e => e.note).join('\n\n');
    if (list.length > 45000) throw new Error('The chronic-absence list is too long for one WIG cell.');
    const missingText = Array.from(details).sort(([a], [b]) => a.localeCompare(b))
      .map(([label, reasons]) => label + ': ' + reasons.join('; ')).join('\n');
    if (missingText.length > 45000) throw new Error('The missing-data list is too long for WIG column L.');
    // Count first, then one student per line inside the same cell. Zero stays a number.
    const failingCell = (list, title) => {
      if (!list.length) return 0;
      const lines = sortedLines(list);
      const value = [String(list.length)].concat(lines).join('\n');
      if (value.length > 45000) throw new Error('The ' + title + ' student list is too long for one WIG cell.');
      return value;
    };
    return {week, row, ready: true, complete: issues.length === 0, students: students.size, matched, missing, inactive,
      issues, missingText, missingDataStudents: details.size, labels: METRICS.slice(), chronicNote,
      values: [students.size, chronic.length, list, failingCell(failing.MATH, 'Failing Math'),
        failingCell(failing.ELA, 'Failing English'), failingCell(failing.CTE, 'Failing CTE')],
      failingCounts: [failing.MATH.length, failing.ELA.length, failing.CTE.length]};
  }
  /** How many of the given cell texts list each student line. */
  function repeatCounts(texts) {
    const counts = new Map();
    texts.forEach(t => new Set(String(t).split('\n').map(lineKey).filter(Boolean))
      .forEach(key => counts.set(key, (counts.get(key) || 0) + 1)));
    return counts;
  }
  function repeatColor(line, counts) { return REPEAT_COLORS[Math.min(counts.get(lineKey(line)) || 0, 4)] || ''; }
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
    const ops = [];
    for (let i = 0; i < cells.length;) {
      const first = cells[i], values = [first.value]; i++;
      while (i < cells.length && cells[i].column === first.column + values.length) values.push(cells[i++].value);
      const range = wig.sheet.getRange(wig.row, first.column, 1, values.length);
      const before = range.getValues(), notes = range.getNotes();
      if (range.getFormulas()[0].some(Boolean) || before[0].some(v => !blank(v))) throw new Error('WIG snapshot cells changed during calculation; existing values were preserved.');
      ops.push({range, before, notes, values: HR10CFG.prepareValues(range, [values]),
        wrap: values.some(v => typeof v === 'string' && v.includes('\n'))});
    }
    const missingRange = wig.sheet.getRange(wig.row, MISSING_COLUMN);
    if (missingRange.getFormula()) throw new Error('WIG column L contains a formula on the target row; preserve it by moving it before saving.');
    const before = missingRange.getValues(), notes = missingRange.getNotes();
    const existing = blank(before[0][0]) ? '' : String(before[0][0]);
    const missingValue = existing + (existing && report.missingText ? '\n\n' : '') + report.missingText;
    if (missingValue.length > 45000) throw new Error('WIG column L text is too long; snapshot was not changed.');
    ops.push({range: missingRange, before, notes, values: HR10CFG.prepareValues(missingRange, [[missingValue]]), wrap: false});
    const noteRange = wig.sheet.getRange(wig.row, wig.columns.get('List of Students Chronically Absent') + 1);
    if (report.chronicNote.length > 45000) throw new Error('The chronic-absence note is too long; snapshot was not changed.');
    ops.push({range: noteRange, noteOnly: true, notes: noteRange.getNotes(), afterNotes: [[report.chronicNote]]});
    const attempted = [];
    try {
      ops.forEach(op => {
        attempted.push(op);
        if (op.noteOnly) { op.range.setNotes(op.afterNotes); return; }
        op.range.setValues(op.values);
        if (op.wrap) op.range.setWrap(true).setVerticalAlignment('top'); // Keep each student on its own line.
      });
      SpreadsheetApp.flush();
    } catch (e) {
      const errors = [];
      attempted.reverse().forEach(op => {
        if (op.noteOnly) { try { op.range.setNotes(op.notes); } catch (x) { errors.push(x.message); } return; }
        try { op.range.setValues(HR10CFG.prepareValues(op.range, op.before)); } catch (x) { errors.push(x.message); }
      });
      try { SpreadsheetApp.flush(); } catch (x) { errors.push(x.message); }
      throw new Error('WIG snapshot failed: ' + e.message + (errors.length ? '. Restoration failed; review WIG row ' + wig.row + '.' : '. Snapshot cells were restored.'));
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
      if (report.chronicNote) html += '<p>Note on List of Students Chronically Absent:</p><pre style="white-space:pre-wrap">' + esc(report.chronicNote) + '</pre>';
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
