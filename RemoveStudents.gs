/** 10RosterHR: remove Master students who are no longer in Upstream (v1.0). Requires Config.gs.
 * @NotOnlyCurrentDoc
 * Called from the Student refresh report. Each removal:
 *  - re-checks Upstream and refuses any student who is still there,
 *  - copies every row it deletes to the "Removed Students" tab first (values, as JSON),
 *  - deletes the student from Master, SCC, ECC and STAR together (and Contacts if chosen).
 * WIG history is never changed.
 */
const HR10REMOVE = (() => {
  const ARCHIVE = 'Removed Students';
  const ARCHIVE_HEADERS = ['Removed At', 'Tab', 'Student Number', 'Name', 'Row Values (JSON)'];
  const text = v => HR10CFG.text(v);
  const id = v => {
    try { return HR10CFG.studentId(typeof v === 'number' ? Math.round(v) : v); } catch (e) { return ''; }
  };
  const fmt = (v, ss) => v instanceof Date
    ? Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : text(v);

  /** AddDrop events for the given Student Numbers, one entry per drop/add (days merged). */
  function addDrop(source, keys, ss) {
    const result = new Map(keys.map(k => [k, []]));
    const sheet = source.getSheetByName('AddDrop');
    if (!sheet) return {result, note: 'AddDrop tab not found in the imports workbook.'};
    const t = HR10CFG.table(sheet, ['ADD / DROP', 'STUDENT NUMBER']);
    const col = h => t.columns.get(h);
    const get = (row, h) => col(h) === undefined ? '' : row[col(h)];
    const groups = new Map();
    t.rows.forEach(row => {
      const event = text(get(row, 'ADD / DROP')).toUpperCase();
      if (event !== 'ADD' && event !== 'DROP') return;   // skips repeated header rows and notes
      const key = id(get(row, 'STUDENT NUMBER'));
      if (!result.has(key)) return;
      const when = text(get(row, 'DATE / TIME'));
      const course = text(get(row, 'COURSE NAME'));
      const g = JSON.stringify([key, event, when, course]);
      if (!groups.has(g)) {
        const entry = {event, when, course, studentName: text(get(row, 'STUDENT NAME')),
          enrolled: fmt(get(row, 'ENROLLMENT DATE'), ss), hasLeft: text(get(row, 'HAS LEFT')),
          deleted: text(get(row, 'DELETED STUDENT')), days: []};
        groups.set(g, entry);
        result.get(key).push(entry);
      }
      const day = text(get(row, 'DAY'));
      if (day && !groups.get(g).days.includes(day)) groups.get(g).days.push(day);
    });
    return {result, note: ''};
  }

  /** Name, grade, start date and AddDrop history for students missing from Upstream. */
  function details(ss, keys, source) {
    if (!keys.length) return [];
    const m = HR10CFG.table(ss.getSheetByName('Master'), ['Student Number']);
    const byId = new Map(m.rows.map(r => [id(r[m.columns.get('Student Number')]), r]));
    const get = (row, h) => !row || m.columns.get(h) === undefined ? '' : row[m.columns.get(h)];
    const ad = addDrop(source, keys, ss);
    return keys.map(key => {
      const row = byId.get(key);
      return {key, name: (text(get(row, 'FIRST NAME')) + ' ' + text(get(row, 'LAST NAME'))).trim(),
        grade: text(get(row, 'GRADE')), start: fmt(get(row, 'START DATE'), ss),
        events: ad.result.get(key) || [], note: ad.note};
    });
  }

  function archiveSheet(ss) {
    let sheet = ss.getSheetByName(ARCHIVE);
    if (!sheet) {
      sheet = ss.insertSheet(ARCHIVE);
      sheet.getRange(1, 1, 1, ARCHIVE_HEADERS.length).setValues([ARCHIVE_HEADERS]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  function remove(ids, includeContacts) {
    const ss = HR10CFG.workbook();
    const keys = [...new Set((ids || []).map(v => HR10CFG.studentId(v)).filter(Boolean))];
    if (!keys.length) throw new Error('No students were selected.');

    // Safety: never remove a student who is in Upstream.
    const up = HR10CFG.table(HR10CFG.sourceWorkbook(ss).getSheetByName('Upstream'), ['STUDENT_NUMBER']);
    const inUpstream = new Set(up.rows.map(r => id(r[up.columns.get('STUDENT_NUMBER')])));
    const blocked = keys.filter(k => inUpstream.has(k));
    if (blocked.length) throw new Error('Still in Upstream, not removed: ' + blocked.join(', '));

    return HR10CFG.withLock(() => {
      const tabs = [...HR10CFG.ROSTER_TABS, ...(includeContacts ? ['Contacts'] : [])];
      const plan = tabs.map(name => {
        const sheet = ss.getSheetByName(name);
        if (!sheet) throw new Error(name + ' tab is missing.');
        const t = HR10CFG.table(sheet, ['Student Number']);
        const rows = [];
        t.rows.forEach((row, i) => {
          const key = id(row[t.columns.get('Student Number')]);
          if (keys.includes(key)) rows.push({number: i + 2, key, values: row});
        });
        if (HR10CFG.ROSTER_TABS.includes(name)) {
          const missing = keys.filter(k => !rows.some(r => r.key === k));
          if (missing.length) throw new Error(name + ' does not contain ' + missing.join(', ') + '. Nothing was removed.');
        }
        return {name, sheet, t, rows};
      });

      const masterPlan = plan[0];
      const nameOf = new Map(masterPlan.rows.map(r => [r.key,
        (text(r.values[masterPlan.t.columns.get('FIRST NAME')]) + ' ' +
         text(r.values[masterPlan.t.columns.get('LAST NAME')])).trim()]));

      // 1) Archive everything first.
      const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      const archive = [];
      plan.forEach(p => p.rows.forEach(r => {
        const record = {};
        p.t.headers.forEach((h, c) => { if (h) record[h] = r.values[c] instanceof Date ? fmt(r.values[c], ss) : r.values[c]; });
        archive.push([stamp, p.name, "'" + r.key, nameOf.get(r.key) || '', JSON.stringify(record)]);
      }));
      const sheet = archiveSheet(ss);
      sheet.getRange(sheet.getLastRow() + 1, 1, archive.length, ARCHIVE_HEADERS.length).setValues(archive);
      SpreadsheetApp.flush();

      // 2) Delete bottom-up on each tab so row numbers stay valid.
      const counts = {};
      plan.forEach(p => {
        p.rows.map(r => r.number).sort((a, b) => b - a).forEach(n => p.sheet.deleteRow(n));
        counts[p.name] = p.rows.length;
      });
      HR10CFG.log('SUCCESS', 'RemoveStudents', keys.join(','), 'Removed from ' + tabs.join(', '), JSON.stringify(counts));
      return {removed: keys.map(k => k + (nameOf.get(k) ? ' ' + nameOf.get(k) : '')), counts};
    }, 'Roster tools are busy. Please retry the removal.');
  }

  /** HTML section for the refresh report: one checkbox per missing student. */
  function section(list) {
    const esc = HR10CFG.escapeHtml;
    if (!list.length) return '';
    let html = '<h3>Master students not in Upstream</h3>' +
      '<p>Check the students to remove. Their rows are copied to the <b>' + ARCHIVE +
      '</b> tab, then deleted from Master, SCC, ECC and STAR. WIG history is not changed.</p>' +
      '<table border="1" cellpadding="5" style="border-collapse:collapse"><tr><th>Remove</th><th>Student</th><th>AddDrop</th></tr>';
    list.forEach(s => {
      const events = s.events.length ? s.events.map(e =>
        '<b>' + esc(e.event) + '</b> ' + esc(e.when) + ' — ' + esc(e.course) +
        '<br><small>Enrolled ' + esc(e.enrolled || '?') + '; Has left: ' + esc(e.hasLeft || '?') +
        '; Deleted: ' + esc(e.deleted || '?') + (e.days.length ? '; ' + esc(e.days.length >= 5 ? 'all days' : e.days.join(', ')) : '') + '</small>'
      ).join('<br>') : '<i>' + esc(s.note || 'No AddDrop record found.') + '</i>';
      html += '<tr><td style="text-align:center"><input type="checkbox" class="rm" value="' + esc(s.key) + '"></td>' +
        '<td>' + esc(s.key) + '<br>' + esc(s.name) + '<br><small>Grade ' + esc(s.grade) + '; start ' + esc(s.start) +
        '</small></td><td>' + events + '</td></tr>';
    });
    html += '</table><p><label><input type="checkbox" id="rmContacts" checked> Also remove their Contacts rows</label></p>' +
      '<button id="rmBtn" onclick="rmClick()">Remove checked students</button> <span id="rmMsg"></span>' +
      '<script>let rmArmed=false;function rmMsg(t){document.getElementById("rmMsg").textContent=t;}' +
      'function rmClick(){const ids=[...document.querySelectorAll("input.rm:checked")].map(x=>x.value);' +
      'const b=document.getElementById("rmBtn");if(!ids.length){rmMsg("Check at least one student.");return;}' +
      'if(!rmArmed){rmArmed=true;b.textContent="Click again to remove "+ids.length+" student(s)";return;}' +
      'b.disabled=true;rmMsg("Removing…");google.script.run.withSuccessHandler(r=>{rmMsg("Removed: "+r.removed.join("; ")+' +
      '". Rows saved in ' + ARCHIVE + '.");document.querySelectorAll("input.rm:checked").forEach(x=>{x.disabled=true;x.checked=false;});' +
      'b.textContent="Remove checked students";b.disabled=false;rmArmed=false;})' +
      '.withFailureHandler(e=>{rmMsg("Not removed: "+e.message);b.textContent="Remove checked students";b.disabled=false;rmArmed=false;})' +
      '.remove10MissingStudents(ids,document.getElementById("rmContacts").checked);}</script>';
    return html;
  }

  return Object.freeze({details, remove, section});
})();

function remove10MissingStudents(ids, includeContacts) { return HR10REMOVE.remove(ids, includeContacts); }
