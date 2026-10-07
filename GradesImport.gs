/** Grades tab import v1.0. Requires Config.gs. Called by RefreshStudentData.gs
 * (Student refresh menu) after Master is refreshed; it has no menu of its own.
 * Source: the CourseGrades tab of the imports workbook in Info!H2 (read only).
 *
 * Grades layout (row 1): Preferred Name, Student Number, LAST NAME, FIRST NAME,
 * GRADE, START DATE, then course slots. A slot is any header followed by a
 * "Percentage" column:
 *   - a course code (2-4 capital letters, e.g. MTH, ENG, ORN, HST, ART) takes
 *     the course whose COURSE_NAME starts with that code (MTH208A Geometry CA);
 *   - "Course 1", "Course 2", ... take every other course, in name order.
 * Adding a new code slot (e.g. SCI + Percentage) needs no code change.
 * Columns that are not slots (Week 6, photo columns) are never written.
 *
 * Each refresh rewrites the course name and current Canvas percentage in every
 * slot for students found in CourseGrades. Students missing from CourseGrades
 * keep their previous grades. Master students missing from Grades are added at
 * the bottom; A:F follows Master. Grades rows not in Master are kept.
 */
const HR10GRADEIMPORT = (() => {
  const SOURCE_TAB = 'CourseGrades';
  // First non-blank wins. Unposted = the teacher's current gradebook score;
  // swap the order to prefer what students see (posted).
  const GRADE_HEADERS = ['CANVAS_UNPOSTED_COURSE_GRADE', 'CANVAS_POSTED_COURSE_GRADE'];
  const ID_FIELDS = ['Preferred Name', 'Student Number', 'LAST NAME', 'FIRST NAME', 'GRADE', 'START DATE'];
  const text = v => HR10CFG.text(v);
  const blank = v => HR10CFG.blank(v);
  const same = (a, b) => HR10CFG.same(a, b);

  /** Course slots from normalized row-1 headers (0-based column indexes). */
  function slots(headers) {
    ID_FIELDS.forEach((h, i) => {
      if (headers[i] !== h) throw new Error('Grades A:F headers must be: ' + ID_FIELDS.join(', ') + '.');
    });
    const named = new Map(), extra = [];
    headers.forEach((h, i) => {
      if (i < ID_FIELDS.length || headers[i + 1] !== 'Percentage') return;
      if (/^Course \d+$/i.test(h)) extra.push({course: i, pct: i + 1, label: h});
      else if (/^[A-Z]{2,4}$/.test(h)) {
        if (named.has(h)) throw new Error('Grades: duplicate course column ' + h + '.');
        named.set(h, {course: i, pct: i + 1, label: h});
      }
    });
    if (!named.size && !extra.length) throw new Error('Grades has no course columns (a header followed by Percentage).');
    const all = [...named.values(), ...extra];
    return {named, extra, all, last: Math.max(...all.map(s => s.pct))};
  }

  /** Course code at the start of a course name: "MTH208A Geometry" -> MTH. */
  function code(courseName) {
    const match = text(courseName).match(/^([A-Z]{2,4})\d/);
    return match ? match[1] : '';
  }

  function number(value) {
    if (typeof value === 'number') return value;
    const s = text(value).replace(/%$/, '');
    return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : '';
  }

  /** CourseGrades rows grouped by Student Number. */
  function readSource(sheet) {
    const t = HR10CFG.table(sheet, ['STUDENT_IDENTITY_ID', 'COURSE_NAME']);
    const gradeColumns = GRADE_HEADERS.filter(h => t.columns.has(h)).map(h => t.columns.get(h));
    if (!gradeColumns.length) throw new Error(SOURCE_TAB + ' has none of ' + GRADE_HEADERS.join(', ') + '.');
    const col = h => t.columns.get(h);
    const byId = new Map(), issues = [], seen = new Set();
    t.rows.forEach((row, i) => {
      if (row.every(blank)) return;
      const id = HR10CFG.studentId(row[col('STUDENT_IDENTITY_ID')]);
      const name = text(row[col('COURSE_NAME')]);
      if (!id || !name) { issues.push(SOURCE_TAB + ' row ' + (i + 2) + ': no student or course name; skipped.'); return; }
      if (t.columns.has('CLASSROOM_ENROLLMENT_STATUS')) {
        const status = text(row[col('CLASSROOM_ENROLLMENT_STATUS')]);
        if (status && status.toLowerCase() !== 'active') return;
      }
      const key = id + '|' + name;
      if (seen.has(key)) { issues.push(id + ': ' + name + ' listed twice in ' + SOURCE_TAB + '; first row used.'); return; }
      seen.add(key);
      let grade = '';
      for (const c of gradeColumns) { grade = number(row[c]); if (grade !== '') break; }
      const drop = t.columns.has('FLAG_FUTURE_COURSE_DROP') && ['1', 'TRUE', 'Y', 'YES'].includes(text(row[col('FLAG_FUTURE_COURSE_DROP')]).toUpperCase());
      if (!byId.has(id)) byId.set(id, []);
      byId.get(id).push({name, grade, drop});
    });
    const dateColumn = t.columns.get('IMPORT_DATE');
    const importDate = dateColumn === undefined ? '' : t.rows.map(r => r[dateColumn]).find(v => !blank(v)) || '';
    return {byId, issues, importDate};
  }

  /** Slot index -> value for one student's courses. */
  function assign(courses, layout) {
    const cells = new Map(), used = new Set(), overflow = [];
    let next = 0;
    [...courses].sort((a, b) => a.name.localeCompare(b.name)).forEach(course => {
      const c = code(course.name);
      let slot = layout.named.get(c);
      if (slot && !used.has(c)) used.add(c);
      else slot = layout.extra[next++];
      if (!slot) { overflow.push(course.name); return; }
      cells.set(slot.course, course.name);
      cells.set(slot.pct, course.grade);
    });
    return {cells, overflow};
  }

  /** Pure planner. values/formulas: Grades rows 2.. (width >= layout.last + 1).
   * master: [{id, values: A:F}]. Returns contiguous row writes for columns 1..last+1. */
  function plan(headers, values, formulas, master, source) {
    const layout = slots(headers);
    const width = layout.last + 1;
    const guarded = [...ID_FIELDS.keys(), ...layout.all.flatMap(s => [s.course, s.pct])];
    const rows = values.map(r => { const copy = r.slice(0, width); while (copy.length < width) copy.push(''); return copy; });
    const byId = new Map(), issues = [], changes = [];
    let lastUsed = 0;
    rows.forEach((row, i) => {
      if (row.every(blank) && !(formulas[i] || []).some(Boolean)) return;
      lastUsed = i + 1;
      const id = HR10CFG.studentId(row[1]);
      if (!id) throw new Error('Grades row ' + (i + 2) + ' has no Student Number.');
      if (byId.has(id)) throw new Error('Grades: duplicate Student Number ' + id + '.');
      byId.set(id, i);
      guarded.forEach(c => { if ((formulas[i] || [])[c]) throw new Error('Grades row ' + (i + 2) + ' has a formula in ' + headers[c] + '; remove it first.'); });
    });
    const before = rows.map(r => r.slice());
    const masterIds = new Set();
    const added = [];
    master.forEach(student => {
      masterIds.add(student.id);
      let i = byId.get(student.id);
      if (i === undefined) {
        i = lastUsed++;
        if (i >= rows.length) { rows.push(new Array(width).fill('')); before.push(new Array(width).fill('')); }
        byId.set(student.id, i);
        added.push(student.id);
      }
      const row = rows[i];
      student.values.forEach((v, c) => {
        // Keep the existing Student Number representation when it is the same ID.
        if (c === 1 && HR10CFG.studentId(row[1]) === student.id) return;
        row[c] = v;
      });
      const courses = source.byId.get(student.id);
      if (!courses) { issues.push(student.id + ': no courses in ' + SOURCE_TAB + '; previous grades retained.'); return; }
      const {cells, overflow} = assign(courses, layout);
      layout.all.forEach(s => { row[s.course] = ''; row[s.pct] = ''; });
      cells.forEach((v, c) => { row[c] = v; });
      if (overflow.length) issues.push(student.id + ': no free Course column for ' + overflow.join(', ') + '. Add a Course N + Percentage pair to show it.');
      courses.filter(c => c.drop).forEach(c => issues.push(student.id + ': ' + c.name + ' is flagged for a future drop.'));
      courses.filter(c => c.grade === '').forEach(c => issues.push(student.id + ': ' + c.name + ' has no Canvas grade yet.'));
    });
    byId.forEach((i, id) => { if (!masterIds.has(id)) issues.push('Grades row ' + (i + 2) + ' (' + id + ') is not on Master; retained.'); });
    const changed = [];
    rows.forEach((row, i) => {
      let any = false;
      row.forEach((v, c) => {
        if (same(before[i][c], v)) return;
        any = true;
        if (c >= ID_FIELDS.length) changes.push({studentNumber: HR10CFG.studentId(row[1]), row: i + 2,
          field: HR10CFG.columnLetter(c + 1) + ' ' + headers[c], before: before[i][c], after: v});
      });
      if (any) changed.push(i);
    });
    const writes = [];
    for (let k = 0; k < changed.length;) {
      const start = changed[k], block = [rows[start]];
      k++;
      while (k < changed.length && changed[k] === start + block.length) block.push(rows[changed[k++]]);
      writes.push({row: start + 2, values: block});
    }
    return {writes, changes, issues: source.issues.concat(issues), added, width,
      studentsUpdated: new Set(changes.map(c => c.studentNumber)).size, importDate: source.importDate};
  }

  /** Reads Master + Grades + CourseGrades, writes Grades unless preview. */
  function run(ss, sourceBook, preview) {
    const sourceSheet = sourceBook.getSheetByName(SOURCE_TAB);
    if (!sourceSheet) throw new Error('The imports workbook has no ' + SOURCE_TAB + ' tab.');
    const source = readSource(sourceSheet);
    return HR10CFG.withLock(() => {
      const sheet = ss.getSheetByName('Grades');
      if (!sheet) throw new Error('Grades tab is missing.');
      const lastColumn = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(HR10CFG.normalize);
      const width = slots(headers).last + 1;
      const height = Math.max(sheet.getLastRow() - 1, 0);
      const values = height ? sheet.getRange(2, 1, height, width).getValues() : [];
      const formulas = height ? sheet.getRange(2, 1, height, width).getFormulas() : [];
      const m = HR10CFG.table(ss.getSheetByName('Master'), ID_FIELDS);
      const master = m.rows.filter(r => !blank(r[m.columns.get('Student Number')])).map(r => {
        const vals = ID_FIELDS.map(h => r[m.columns.get(h)]);
        return {id: HR10CFG.studentId(vals[1]), values: vals};
      });
      const result = plan(headers, values, formulas, master, source);
      if (preview || !result.writes.length) return result;
      const needed = Math.max(...result.writes.map(w => w.row + w.values.length - 1));
      if (needed > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows());
      result.writes.forEach(w => {
        const range = sheet.getRange(w.row, 1, w.values.length, width);
        range.setValues(HR10CFG.prepareValues(range, w.values));
      });
      SpreadsheetApp.flush();
      return result;
    }, 'Roster tools are busy; Grades was not updated. Please retry.');
  }

  return Object.freeze({run, plan, slots, code, assign, readSource});
})();
