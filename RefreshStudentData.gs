/** 10RosterHR source refresh v1.5 (course grades into Master's MTH/MTH% ... columns and the Grades tab). Replace the old RefreshStudentData file.
 * @NotOnlyCurrentDoc
 * Requires Config.gs, RosterSync v1.3, RemoveStudents.gs and GradesImport.gs.
 * Course grades from the CourseGrades tab are part of the Master refresh
 * (synced to SCC, ECC and STAR). The Grades tab is written after Master is
 * saved, under its own lock; a Grades tab problem is reported in the dialog and
 * never undoes the Master refresh. Menus are built by Menus.gs. Info!H2 holds the separate imports workbook URL.
 * Source is READ ONLY. Planner performs no writes; shared helper saves everything.
 */
const HR10REFRESH = (() => {
  const MAP = [
    ['STUDENT_LAST_NAME', 'LAST NAME'], ['STUDENT_FIRST_NAME', 'FIRST NAME'],
    ['GRADE', 'GRADE'], ['ENROLL_DATE_PS', 'START DATE'], ['SCHOOL', 'SCHOOL'],
    ['SPED504PLAN', 'SPED'], ['EMAIL_O365', 'Student Email'],
    ['COUNSELOR', 'Counselor'], ['COUNSELOR_EMAIL', 'Counselor Email']
  ];
  const CONTACT_HEADERS = ['Student Number', 'Student Source Name', 'Gender', 'Student Phone',
    'Contact Name', 'Relationship', 'Source Phone', 'Source Email', 'Effective Phone',
    'Effective Email', 'Learning Coach', 'Contact Key'];
  const text = v => HR10CFG.text(v);
  const blank = v => HR10CFG.blank(v);
  const same = (a, b) => HR10CFG.same(a, b);
  const id = v => HR10CFG.studentId(v);
  const book = () => HR10CFG.workbook();
  const table = (sheet, required) => HR10CFG.table(sheet, required);
  function source(ss) {
    const upstream = HR10CFG.sourceWorkbook(ss);
    const u = table(upstream.getSheetByName('Upstream'), ['STUDENT_NUMBER', 'CURRENT_ACTIVE_STUDENT']);
    const p = table(upstream.getSheetByName('ProRoster'),
      ['Student Number', 'Name', 'Gender', 'Phone', 'Contact Name', 'Contact Phone', 'Contact Email']);
    const students = new Map();
    u.rows.forEach((row, offset) => {
      const key = id(row[u.columns.get('STUDENT_NUMBER')]);
      if (!key) {
        if (row.some(v => !blank(v))) throw new Error('Upstream row ' + (offset + 2) + ' has no STUDENT_NUMBER.');
        return;
      }
      if (students.has(key)) throw new Error('Upstream has duplicate STUDENT_NUMBER ' + key);
      const active = ['Y', 'YES', 'TRUE', '1'].includes(text(row[u.columns.get('CURRENT_ACTIVE_STUDENT')]).toUpperCase());
      students.set(key, {row, active});
    });
    let grades = null, gradesError = '';
    try { grades = HR10GRADEIMPORT.sourceFrom(upstream); }
    catch (e) { gradesError = e.message; }
    return {u, p, students, book: upstream, grades, gradesError};
  }
  function date(v, label) {
    if (v instanceof Date && !isNaN(v.getTime())) return new Date(v.getTime());
    const s = text(v);
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/), y, mo, d;
    if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
    else {
      m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) { y = +m[3]; mo = +m[1]; d = +m[2]; }
    }
    if (m) {
      const check = new Date(Date.UTC(y, mo - 1, d));
      if (check.getUTCFullYear() === y && check.getUTCMonth() === mo - 1 && check.getUTCDate() === d) {
        return Utilities.parseDate(y + '-' + mo + '-' + d, book().getSpreadsheetTimeZone(), 'yyyy-M-d');
      }
    }
    throw new Error('Invalid date for ' + label + '. Use a Sheets date or YYYY-MM-DD/MM/DD/YYYY.');
  }
  function nameKey(v) {
    return text(v).replace(/\([^)]*\)/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().split(/\s+/).sort().join(' ');
  }
  function phoneKey(v) { return text(v).replace(/\([^)]*\)/g, '').replace(/\D/g, ''); }
  function emailKey(v) { return text(v).replace(/\([^)]*\)/g, '').trim().toLowerCase(); }
  function roleKey(v) { return text(v).toLowerCase().replace(/\s+/g, ' '); }
  function planContacts(ss, masterIds, data, report) {
    const t = table(ss.getSheetByName('Contacts'), CONTACT_HEADERS);
    const old = new Map(), fallback = new Map();
    const append = (map, key, value) => {
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(value);
    };
    t.rows.forEach((row, i) => {
      if (row.every(blank)) return;
      const key = id(row[t.columns.get('Student Number')]);
      if (!key) throw new Error('Contacts row ' + (i + 2) + ' has no Student Number.');
      const ck = text(row[t.columns.get('Contact Key')]);
      if (ck && !ck.startsWith(key + '|')) throw new Error('Contacts row ' + (i + 2) + ': Contact Key does not match Student Number.');
      const fk = JSON.stringify([key, nameKey(row[t.columns.get('Contact Name')])]);
      const entry = {row, number: i + 2, role: roleKey(row[t.columns.get('Relationship')]),
        phone: phoneKey(row[t.columns.get('Source Phone')]), email: emailKey(row[t.columns.get('Source Email')])};
      // The old import can use one Contact Key for several phone/email rows.
      // Keep every row; uniqueness is established when matching a source record.
      if (ck) append(old, ck, entry);
      append(fallback, fk, entry);
    });
    const proposed = [], seenRecords = new Set(), sourceCounts = new Map(), matched = new Set();
    let key = '', studentName = '', gender = '', phone = '';
    data.p.rows.forEach((row, i) => {
      const get = h => row[data.p.columns.get(h)];
      const next = id(get('Student Number'));
      if (next) { key = next; studentName = ''; gender = ''; phone = ''; }
      if (!key) {
        if (row.some(v => !blank(v))) throw new Error('ProRoster row ' + (i + 2) + ' starts without a Student Number.');
        return;
      }
      if (!blank(get('Name'))) studentName = get('Name');
      if (!blank(get('Gender'))) gender = get('Gender');
      if (!blank(get('Phone'))) phone = get('Phone');
      if (!masterIds.has(key) || !data.students.get(key)?.active || blank(get('Contact Name'))) return;
      const contact = text(get('Contact Name'));
      const contactKey = key + '|' + contact.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
      const fk = JSON.stringify([key, nameKey(contact)]);
      const recordKey = JSON.stringify([contactKey, phoneKey(get('Contact Phone')), emailKey(get('Contact Email'))]);
      if (seenRecords.has(recordKey)) {
        report.issues.push('ProRoster row ' + (i + 2) + ' for ' + key + ': repeated contact/channel record skipped. Existing contacts retained.');
        return;
      }
      seenRecords.add(recordKey);
      sourceCounts.set(contactKey, (sourceCounts.get(contactKey) || 0) + 1);
      const coachColumn = data.u.columns.get('LC_COACH_NAME');
      const coach = coachColumn === undefined ? '' : data.students.get(key).row[coachColumn];
      const relationship = (contact.match(/\(([^)]*)\)/) || [])[1] || '';
      const values = {'Student Number': key, 'Student Source Name': studentName, 'Gender': gender,
        'Student Phone': phone, 'Contact Name': contact, 'Relationship': relationship,
        'Source Phone': get('Contact Phone'), 'Source Email': get('Contact Email'),
        'Learning Coach': blank(coach) ? '' : nameKey(contact) === nameKey(coach), 'Contact Key': contactKey};
      proposed.push({key, values, fk, sourceRow: i + 2});
    });
    const writes = [];
    let nextRow = t.rows.length + 2;
    proposed.forEach(item => {
      const {key, values, fk} = item;
      const contactKey = values['Contact Key'];
      const exact = old.get(contactKey) || [];
      const role = roleKey(values.Relationship);
      const candidates = exact.length ? exact : (fallback.get(fk) || []).filter(e => !role || e.role === role);
      const phone = phoneKey(values['Source Phone']), email = emailKey(values['Source Email']);
      const scored = candidates.map(entry => ({entry, score: (phone && phone === entry.phone ? 1 : 0) +
        (email && email === entry.email ? 1 : 0)})).sort((a, b) => b.score - a.score);
      let existing;
      if (scored.length && scored[0].score > 0 && (scored.length === 1 || scored[0].score > scored[1].score)) {
        existing = scored[0].entry;
      } else if (candidates.length === 1 && sourceCounts.get(contactKey) === 1) {
        existing = candidates[0];
      }
      if (candidates.length && !existing || existing && matched.has(existing.number)) {
        report.issues.push('Student ' + key + ', source ProRoster row ' + item.sourceRow + ' (' + values['Contact Name'] +
          '): contact match is uncertain against Contacts rows ' + candidates.map(e => e.number).join(', ') +
          '. Source contact skipped; existing rows/overrides retained. Review phone/email details.');
        return;
      }
      item.existing = existing;
      if (existing) matched.add(existing.number);
      const effective = (sourceHeader, effectiveHeader) => {
        const previousSource = existing ? existing.row[t.columns.get(sourceHeader)] : '';
        const previousEffective = existing ? existing.row[t.columns.get(effectiveHeader)] : '';
        // Preserve manual overrides, including an intentionally cleared effective value.
        if (existing && !same(previousSource, previousEffective)) return previousEffective;
        return blank(values[sourceHeader]) ? previousSource : values[sourceHeader];
      };
      values['Effective Phone'] = effective('Source Phone', 'Effective Phone');
      values['Effective Email'] = effective('Source Email', 'Effective Email');
      if (existing) CONTACT_HEADERS.forEach(h => {
        if (!['Effective Phone', 'Effective Email'].includes(h) && blank(values[h])) values[h] = existing.row[t.columns.get(h)];
      });
      const row = item.existing ? item.existing.number : nextRow++;
      const changes = [];
      const cells = [];
      CONTACT_HEADERS.forEach(h => {
        const before = item.existing ? item.existing.row[t.columns.get(h)] : '';
        let after = item.values[h];
        // ID representation is retained when the numeric/text identifier matches.
        if (h === 'Student Number' && item.existing && id(before) === item.key) after = before;
        if (!same(before, after)) {
          cells.push({column: t.columns.get(h) + 1, value: after});
          changes.push({field: h, before, after});
        }
      });
      cells.sort((a, b) => a.column - b.column);
      for (let i = 0; i < cells.length;) {
        const first = cells[i], values = [first.value];
        i++;
        while (i < cells.length && cells[i].column === first.column + values.length) values.push(cells[i++].value);
        writes.push({sheet: 'Contacts', row, column: first.column, values: [values]});
      }
      if (changes.length) report.contacts.push({studentNumber: item.key, row, added: !item.existing, changes});
    });
    t.rows.forEach((row, i) => {
      const key = id(row[t.columns.get('Student Number')]);
      if (key && masterIds.has(key) && !matched.has(i + 2)) {
        report.issues.push('Existing Contacts row ' + (i + 2) + ' for ' + key + ' was not matched in active source contacts; retained.');
      }
    });
    return writes;
  }
  function build(ss, master, data) {
    const report = {changes: [], contacts: [], issues: [], sourceOnly: [], missing: [], inactive: []};
    const rows = master.sheet.getRange(2, 1, master.lastRow - 1, master.headers.length).getValues();
    const formulas = master.sheet.getRange(2, 1, master.lastRow - 1, master.headers.length).getFormulas();
    const fields = [];
    const mirrors = ['SCC', 'ECC', 'STAR'].map(name => {
      const t = table(ss.getSheetByName(name), ['Student Number']);
      const byId = new Map(t.rows.map((row, i) => [id(row[t.columns.get('Student Number')]), i]));
      const formulas = t.rows.length ? t.sheet.getRange(2, 1, t.rows.length, t.width).getFormulas() : [];
      return {name, ...t, byId, formulas};
    });
    MAP.forEach(([src, dest]) => {
      if (!data.u.columns.has(src)) report.issues.push('Upstream is missing ' + src + '; ' + dest + ' retained.');
    });
    const gradeLayout = HR10GRADEIMPORT.findSlots(master.headers, 0);
    if (data.gradesError) report.issues.push('Course grades not refreshed: ' + data.gradesError);
    else if (!gradeLayout) report.issues.push('Master has no course columns (e.g. MTH and MTH%); course grades not copied to Master.');
    function add(key, rowNumber, header, value) {
      const c = master.columns.get(header);
      if (!c) throw new Error('Master is missing ' + header);
      if (formulas[rowNumber - 2][c - 1]) throw new Error('Master has a formula in source field ' + header);
      const before = rows[rowNumber - 2][c - 1];
      if (same(before, value)) return;
      // A pending edit in a mirror must be resolved before source refresh overwrites it.
      mirrors.forEach(t => {
        const r = t.byId.get(key), c = t.columns.get(header);
        if (t.formulas[r][c] || !same(t.rows[r][c], before)) {
          throw new Error(t.name + ' differs from Master for student ' + key + ', ' + header + '. Resolve/synchronize this field first.');
        }
      });
      fields.push({studentNumber: key, header, value});
      report.changes.push({studentNumber: key, field: header, before, after: value});
    }
    for (const [key, r] of master.rows) {
      const student = data.students.get(key);
      if (!student) { report.missing.push(key); continue; }
      if (!student.active) { report.inactive.push(key); continue; }
      MAP.forEach(([src, dest]) => {
        if (!data.u.columns.has(src)) return;
        let value = student.row[data.u.columns.get(src)];
        if (dest === 'SPED') {
          // SPED = Yes if the student has an IEP (STUDENT_WITH_DISABILITIES) or a 504 plan (SPED504PLAN).
          const yes = v => ['Y', 'YES', 'TRUE', '1'].includes(text(v).toUpperCase());
          const swdCol = data.u.columns.get('STUDENT_WITH_DISABILITIES');
          const swd = swdCol === undefined ? '' : student.row[swdCol];
          if (yes(swd) || yes(value)) value = 'Yes';
          else if (blank(value) && !blank(swd)) value = 'No';
        }
        if (blank(value)) { report.issues.push(key + ': blank source ' + src + '; existing value retained.'); return; }
        if (dest === 'START DATE') value = date(value, key + ' START DATE');
        add(key, r, dest, value);
      });
      for (const subject of ['Reading', 'Math']) {
        const prefix = 'STAR_2TO12_' + subject.toUpperCase();
        const dh = prefix + '_DATE', sh = prefix + '_UNIFIED_SCORE';
        const targetDate = 'Latest ' + subject + ' Date', targetScore = 'Latest ' + subject + ' Score';
        if (!data.u.columns.has(dh) || !data.u.columns.has(sh)) {
          report.issues.push(key + ': source ' + subject + ' date/score headers missing; pair retained.'); continue;
        }
        const dv = student.row[data.u.columns.get(dh)], sv = student.row[data.u.columns.get(sh)];
        if (blank(dv) || blank(sv)) { report.issues.push(key + ': incomplete source ' + subject + ' date/score; pair retained.'); continue; }
        const parsedDate = date(dv, key + ' ' + dh);
        const score = typeof sv === 'number' ? sv : /^\d+(\.\d+)?$/.test(text(sv)) ? Number(text(sv)) : NaN;
        if (!Number.isFinite(score) || score < 0) throw new Error('Invalid ' + subject + ' score for ' + key);
        const previous = rows[r - 2][master.columns.get(targetDate) - 1];
        if (!blank(previous) && parsedDate.getTime() < date(previous, key + ' existing ' + targetDate).getTime()) {
          report.issues.push(key + ': older source ' + subject + ' assessment retained existing pair.'); continue;
        }
        add(key, r, targetDate, parsedDate); add(key, r, targetScore, score);
      }
      // Course names and Canvas percentages; students without CourseGrades rows keep theirs.
      const gradeFields = data.grades && HR10GRADEIMPORT.masterFields(gradeLayout, data.grades.byId.get(key));
      if (gradeFields) gradeFields.forEach(f => add(key, r, f.header, f.value));
    }
    data.students.forEach((v, key) => { if (v.active && !master.rows.has(key)) report.sourceOnly.push(key); });
    const relatedWrites = planContacts(ss, master.rows, data, report);
    return {fields, relatedWrites, report};
  }
  function display(report, title) {
    const esc = v => HR10CFG.escapeHtml(v instanceof Date ? Utilities.formatDate(v, book().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : v);
    let html = '<div style="font:14px Arial"><p>' + report.changes.length + ' Master field changes; ' + report.contacts.length + ' changed/added contact rows.</p>';
    html += '<p>Source-only active students (not added): ' + esc(report.sourceOnly.join(', ') || 'None') + '<br>Master students missing in source (retained): ' + esc(report.missing.join(', ') || 'None') + '<br>Inactive source students (retained): ' + esc(report.inactive.join(', ') || 'None') + '</p>';
    html += '<table border="1" cellpadding="5" style="border-collapse:collapse"><tr><th>Student / location</th><th>Field</th><th>Before</th><th>After</th></tr>';
    report.changes.forEach(x => { html += '<tr><td>' + esc(x.studentNumber) + '</td><td>' + esc(x.field) + '</td><td>' + esc(x.before) + '</td><td>' + esc(x.after) + '</td></tr>'; });
    report.contacts.forEach(x => x.changes.forEach(c => {
      html += '<tr><td>' + esc(x.studentNumber + ' / Contacts row ' + x.row + (x.added ? ' (new)' : '')) + '</td><td>' + esc(c.field) + '</td><td>' + esc(c.before) + '</td><td>' + esc(c.after) + '</td></tr>';
    }));
    html += '</table>';
    const g = report.grades;
    if (g) {
      html += '<h3>Grades tab</h3><p>' + esc(g.studentsUpdated) + ' students with grade changes; ' + esc(g.added.length) +
        ' rows added' + (g.importDate ? '; CourseGrades import date ' + esc(g.importDate) : '') + '.</p>';
      if (g.changes.length) {
        html += '<table border="1" cellpadding="5" style="border-collapse:collapse"><tr><th>Student / row</th><th>Column</th><th>Before</th><th>After</th></tr>';
        g.changes.forEach(x => { html += '<tr><td>' + esc(x.studentNumber + ' / Grades row ' + x.row) + '</td><td>' + esc(x.field) + '</td><td>' + esc(x.before) + '</td><td>' + esc(x.after) + '</td></tr>'; });
        html += '</table>';
      }
    }
    html += HR10REMOVE.section(report.missingDetails || []);
    html += '<h3>Missing information and review notes</h3><ul>' + report.issues.map(x => '<li>' + esc(x) + '</li>').join('') + '</ul></div>';
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(950).setHeight(650), title);
  }
  function run(preview) {
    const ss = book();
    let report, data;
    try {
      data = source(ss);
      update10RosterFields(master => {
        const result = build(ss, master, data);
        report = result.report;
        result.relatedWrites.push({sheet: 'Info', row: 3, column: 8,
          values: [[Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss z')]]});
        return {fields: result.fields, relatedWrites: result.relatedWrites, dryRun: preview, bulk: true};
      });
    } catch (e) {
      if (!preview) HR10CFG.log('ERROR', 'RefreshStudentData', '', e.message, HR10CFG.errorDetails(e));
      throw e;
    }
    if (!preview) HR10CFG.log('SUCCESS', 'RefreshStudentData', '',
      report.changes.length + ' Master fields; ' + report.contacts.length + ' contact rows refreshed.', '');
    try {
      if (!data.grades) throw new Error('see the Course grades note above');
      report.grades = HR10GRADEIMPORT.run(ss, data.grades, preview);
      report.grades.issues.forEach(x => report.issues.push('Grades: ' + x));
      if (!preview) HR10CFG.log('SUCCESS', 'GradesImport', '', report.grades.changes.length + ' Grades cells changed.', '');
    } catch (e) {
      report.issues.push('Grades import stopped' + (preview ? '' : '; check the Grades tab') + ': ' + e.message);
      if (!preview) HR10CFG.log('ERROR', 'GradesImport', '', e.message, HR10CFG.errorDetails(e));
    }
    try { report.missingDetails = HR10REMOVE.details(ss, report.missing, data.book); }
    catch (e) { report.issues.push('AddDrop details unavailable: ' + e.message); report.missingDetails = []; }
    console.log(JSON.stringify(report));
    try { display(report, preview ? 'Student refresh preview — no data saved' : 'Student refresh saved'); }
    catch (e) { console.error('Refresh report could not be displayed: ' + e.message); }
    return report;
  }
  return Object.freeze({run});
})();
function preview10StudentRefresh() { return HR10REFRESH.run(true); }
function refreshStudentData() { return HR10REFRESH.run(false); }
function refreshStudentDataAndCounselors() { return refreshStudentData(); }
