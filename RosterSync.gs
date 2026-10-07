/**
 * 10RosterHR synchronization, version 1.4. Requires Config.gs.
 * @NotOnlyCurrentDoc
 *
 * INSTALL: In 10RosterHR choose Extensions > Apps Script. Add a script file
 * named RosterSync, paste this entire file, save, then run install10RosterSync
 * once and authorize it. Only one person should install the trigger.
 *
 * Normal cell edits and clears synchronize Master, SCC, ECC, and STAR by
 * Student Number and header. Column order, row order, grouping, filters,
 * formatting, and hidden columns remain local to each tab.
 * Name and Subject Line are calculated locally from the matching headers.
 * Keep header names and Student Numbers intact. This handles existing
 * students: adding/removing students or fields requires updating all four tabs.
 * Paste editable fields only, excluding the header and Student Number column.
 *
 * Script/API writes and formula recalculation do not fire edit triggers.
 * Preferred script writes: call update10RosterFields(planner). The planner
 * reads Master under the same lock and returns [{studentNumber, header, value}].
 * This helper writes all four tabs and attempts rollback if a write fails.
 * Do not wrap it in another document lock or write cells inside the planner.
 * Legacy scripts that already wrote a range may call sync10RosterRange(range).
 * There is deliberately no onEdit or onOpen function to collide with other code.
 *
 * Other manual functions: check10RosterSync (read-only),
 * refresh10RosterFromMaster (Master values overwrite corresponding fields in
 * the other tabs; it never changes layout), uninstall10RosterSync.
 * If two tabs edit the same field at nearly the same time, the last successfully
 * processed edit wins. Review Apps Script Executions if an error is reported.
 * v1.3: text that Sheets would reinterpret (3/4, 00123, TRUE, =x) is written
 * as text on every tab, except in cells already formatted as Plain text.
 * v1.4 (speed): Name and Subject Line formulas are rewritten only for students
 * whose FIRST NAME, LAST NAME or SCHOOL changed, and rollback formats are read
 * only if a save fails.
 * Local tests cover the logic with Apps Script service mocks; activation must
 * be verified in Google Sheets after installation.
 */

const HR10SYNC = (() => {
  const TABS = ['Master', 'SCC', 'ECC', 'STAR'];
  const ID_HEADER = 'Student Number';
  const DERIVED = ['Name', 'Subject Line'];
  // Fields the Name / Subject Line formulas read (Student Number never changes).
  const DERIVED_INPUTS = ['FIRST NAME', 'LAST NAME', 'SCHOOL'];
  const derivedKeysFor = changes => new Set(changes.filter(c => DERIVED_INPUTS.includes(c.header)).map(c => c.key));
  const HANDLER = 'handle10RosterEdit';

  const workbook = ss => HR10CFG.workbook(ss);
  const normalizeHeader = value => HR10CFG.normalize(value);
  const studentKey = value => HR10CFG.studentId(value);
  const letter = column => HR10CFG.columnLetter(column);

  function readIndex(sheet) {
    if (!sheet) throw new Error('A roster tab is missing. Expected Master, SCC, ECC, and STAR.');
    const lastColumn = sheet.getLastColumn();
    if (!lastColumn) throw new Error(sheet.getName() + ' has no headers.');
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(normalizeHeader);
    const columns = new Map();
    headers.forEach((header, i) => {
      if (!header) throw new Error(sheet.getName() + ': a column header is blank.');
      if (columns.has(header)) throw new Error(sheet.getName() + ': duplicate header ' + header + '.');
      columns.set(header, i + 1);
    });
    for (const header of [ID_HEADER, 'FIRST NAME', 'LAST NAME', 'SCHOOL', ...DERIVED]) {
      if (!columns.has(header)) throw new Error(sheet.getName() + ': missing header ' + header + '.');
    }
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const idColumn = columns.get(ID_HEADER);
    const ids = lastRow > 1 ? sheet.getRange(2, idColumn, lastRow - 1, 1).getValues() : [];
    const rows = new Map();
    const rowKeys = new Map();
    ids.forEach((entry, i) => {
      const key = studentKey(entry[0]);
      if (!key) return;
      if (rows.has(key)) throw new Error(sheet.getName() + ': duplicate Student Number ' + key + '.');
      rows.set(key, i + 2);
      rowKeys.set(i + 2, key);
    });
    return {sheet, headers, columns, idColumn, lastRow, rows, rowKeys};
  }

  function indexes(ss) {
    const result = TABS.map(name => readIndex(ss.getSheetByName(name)));
    const master = result[0];
    if (!master.rows.size) throw new Error('Master has no student records.');
    for (const index of result.slice(1)) {
      if (index.columns.size !== master.columns.size ||
          master.headers.some(header => !index.columns.has(header))) {
        throw new Error(index.sheet.getName() + ': headers differ from Master. Order may differ, names must match.');
      }
      if (index.rows.size !== master.rows.size ||
          [...master.rows.keys()].some(key => !index.rows.has(key))) {
        throw new Error(index.sheet.getName() + ': student list differs from Master. Match the lists before syncing.');
      }
    }
    return result;
  }

  function withLock(action) {
    return HR10CFG.withLock(action, 'Roster sync is busy. Please retry the edit.');
  }

  function derivedFormula(index, row, header) {
    const cell = name => letter(index.columns.get(name)) + row;
    if (header === 'Name') {
      return '=CONCATENATE(' + cell('FIRST NAME') + '," ",' + cell('LAST NAME') + ')';
    }
    return '=LEFT(' + cell('FIRST NAME') + ',1)&" "&' + cell('LAST NAME') +
      '&", "&' + cell(ID_HEADER) + '&", "&' + cell('SCHOOL');
  }

  // derivedKeys: students whose Name / Subject Line formulas should be rewritten.
  function operations(all, source, changes, derivedKeys, bulk) {
    const result = [];
    for (const target of all) {
      const perRow = new Map();
      const add = (key, header, value, formula) => {
        const row = target.rows.get(key);
        const column = target.columns.get(header);
        if (!row || !column) throw new Error('A target student or field is missing.');
        if (!perRow.has(row)) perRow.set(row, []);
        perRow.get(row).push({column, value, formula});
      };
      if (target.sheet.getName() !== source.sheet.getName()) {
        changes.forEach(change => add(change.key, change.header, change.value, false));
      }
      // Calculated fields use each tab's own layout; never paste a formula
      // from a different column order or flatten an existing calculation.
      for (const key of derivedKeys) {
        const row = target.rows.get(key);
        DERIVED.forEach(header => add(key, header, derivedFormula(target, row, header), true));
      }
      for (const [row, cells] of perRow) {
        cells.sort((a, b) => a.column - b.column);
        let offset = 0;
        while (offset < cells.length) {
          let end = offset + 1;
          while (end < cells.length && cells[end].column === cells[end - 1].column + 1) end++;
          const block = cells.slice(offset, end);
          const range = target.sheet.getRange(row, block[0].column, 1, block.length);
          // Raw values plus a formula mask; apply() protects text using each cell's format.
          const op = {range, values: [block.map(cell => cell.value)], formulas: [block.map(cell => cell.formula)]};
          if (!bulk) {
            const beforeValues = range.getValues()[0], beforeFormulas = range.getFormulas()[0];
            op.before = [beforeValues.map((value, i) => beforeFormulas[i] || value)];
            op.beforeFormulas = [beforeFormulas.map(Boolean)];
          }
          result.push(op);
          offset = end;
        }
      }
    }
    return bulk ? batchOperations(result) : result;
  }

  // Join adjacent rows with the same affected columns; never include a gap.
  function batchOperations(raw) {
    const groups = new Map(), result = [];
    raw.forEach(op => {
      const key = JSON.stringify([op.range.getSheet().getName(), op.range.getColumn(), op.range.getNumColumns()]);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(op);
    });
    groups.forEach(group => {
      group.sort((a, b) => a.range.getRow() - b.range.getRow());
      for (let i = 0; i < group.length;) {
        const first = group[i], values = first.values.slice(), masks = first.formulas.slice();
        let end = first.range.getRow() + values.length;
        i++;
        while (i < group.length && group[i].range.getRow() === end) {
          values.push(...group[i].values); masks.push(...group[i].formulas);
          end += group[i].values.length; i++;
        }
        const range = first.range.getSheet().getRange(first.range.getRow(), first.range.getColumn(), values.length, first.range.getNumColumns());
        const before = range.getValues(), formulas = range.getFormulas();
        result.push({range, values, formulas: masks,
          before: before.map((row, r) => row.map((v, c) => formulas[r][c] || v)),
          beforeFormulas: formulas.map(row => row.map(Boolean))});
      }
    });
    return result;
  }

  function apply(ops) {
    // Formats are read once per block before writing. Rollback formats are read
    // only on failure: setValues never sets Plain text, so they still match.
    ops.forEach(op => { op.ready = HR10CFG.prepareValues(op.range, op.values, op.formulas); });
    const attempted = [];
    try {
      ops.forEach(op => {
        attempted.push(op);
        op.range.setValues(op.ready);
      });
      SpreadsheetApp.flush();
    } catch (error) {
      const rollbackFailures = [];
      attempted.reverse().forEach(op => {
        try { op.range.setValues(HR10CFG.prepareValues(op.range, op.before, op.beforeFormulas)); }
        catch (rollbackError) { rollbackFailures.push(String(rollbackError)); }
      });
      try { SpreadsheetApp.flush(); } catch (flushError) { rollbackFailures.push(String(flushError)); }
      throw new Error('Sync failed: ' + error.message +
        (rollbackFailures.length ? '. Some writes could not be restored; review the affected sheets.' :
          '. Script writes were restored. Your original edit remains; retry after fixing the error.'));
    }
  }

  function sync(range) {
    if (!range) throw new Error('Provide the edited range; do not run the edit handler manually.');
    const ss = workbook(range.getSheet().getParent());
    if (!TABS.includes(range.getSheet().getName())) return {cells: 0};
    return withLock(() => {
      const all = indexes(ss);
      const source = all.find(index => index.sheet.getName() === range.getSheet().getName());
      if (range.getRow() === 1) {
        throw new Error('Header edits are not synchronized. Keep matching headers in all four tabs.');
      }
      if (range.getColumn() <= source.idColumn && range.getLastColumn() >= source.idColumn) {
        throw new Error('Student Numbers are sync keys. Paste editable fields without the Student Number column.');
      }
      const values = range.getValues();
      const formulas = range.getFormulas();
      const changes = [];
      const touchedKeys = new Set();
      for (let r = 0; r < values.length; r++) {
        const key = source.rowKeys.get(range.getRow() + r);
        if (!key) {
          if (values[r].some(value => value !== '' && value != null)) {
            throw new Error('Row ' + (range.getRow() + r) + ' has no Student Number. Add students to all four tabs first.');
          }
          continue;
        }
        touchedKeys.add(key);
        for (let c = 0; c < values[r].length; c++) {
          const header = source.headers[range.getColumn() + c - 1];
          if (!header) {
            if (values[r][c] !== '') throw new Error('The edited column has no matching header.');
            continue;
          }
          if (DERIVED.includes(header)) continue;
          if (formulas[r][c]) throw new Error('Use values in editable fields. Name and Subject Line calculate automatically.');
          changes.push({key, header, value: values[r][c]});
        }
      }
      apply(operations(all, source, changes, derivedKeysFor(changes)));
      return {cells: changes.length, students: touchedKeys.size};
    });
  }

  function refresh() {
    const ss = workbook();
    return withLock(() => {
      const all = indexes(ss), source = all[0];
      const rows = source.sheet.getRange(2, 1, source.lastRow - 1, source.headers.length).getValues();
      const formulas = source.sheet.getRange(2, 1, source.lastRow - 1, source.headers.length).getFormulas();
      const changes = [], keys = new Set(source.rows.keys());
      for (const [key, row] of source.rows) source.headers.forEach((header, c) => {
        if (header === ID_HEADER || DERIVED.includes(header)) return;
        if (formulas[row - 2][c]) throw new Error('Master has a formula in an editable field: ' + header + '.');
        changes.push({key, header, value: rows[row - 2][c]});
      });
      apply(operations(all, source, changes, keys));
      ss.toast('Master values synchronized to SCC, ECC, and STAR.', 'Roster sync', 6);
      return {cells: changes.length, students: keys.size};
    });
  }

  // Planner reads current Master values while holding the same lock as edits.
  // Return [{studentNumber, header, value}]; all four sheets change together.
  function updateFields(planner) {
    if (typeof planner !== 'function') throw new Error('Provide a read-only update planner.');
    const ss = workbook();
    return withLock(() => {
      const all = indexes(ss), master = all[0];
      const packet = planner(master);
      const planned = Array.isArray(packet) ? packet : packet && packet.fields;
      const related = Array.isArray(packet) ? [] : packet && packet.relatedWrites || [];
      if (!Array.isArray(planned) || !Array.isArray(related)) throw new Error('Update planner must return an array or {fields, relatedWrites}.');
      const seen = new Set(), keys = new Set();
      const changes = planned.map(item => {
        const key = studentKey(item.studentNumber), header = normalizeHeader(item.header);
        if (!master.rows.has(key)) throw new Error('Student Number is not in Master: ' + key);
        if (!master.columns.has(header) || header === ID_HEADER || DERIVED.includes(header)) {
          throw new Error('Cannot update this roster field: ' + header);
        }
        const token = JSON.stringify([key, header]);
        if (seen.has(token)) throw new Error('Duplicate field in update: ' + header);
        seen.add(token); keys.add(key);
        const value = item.value;
        if (value == null || !['string', 'number', 'boolean'].includes(typeof value) && !(value instanceof Date)) {
          throw new Error('Use a value or an empty string for ' + header);
        }
        if (typeof value === 'number' && !Number.isFinite(value) || value instanceof Date && isNaN(value.getTime())) {
          throw new Error('Invalid value for ' + header);
        }
        return {key, header, value};
      });
      const bulk = !Array.isArray(packet) && packet.bulk === true;
      const ops = changes.length ? operations(all, {sheet: {getName: () => ''}}, changes, derivedKeysFor(changes), bulk) : [];
      const occupied = new Set();
      let relatedCells = 0;
      related.forEach(item => {
        const sheet = ss.getSheetByName(item.sheet);
        const values = item.values;
        if (!sheet || !Array.isArray(values) || !values.length || !Array.isArray(values[0]) || !values[0].length ||
            !Number.isInteger(item.row) || !Number.isInteger(item.column)) throw new Error('Invalid related write.');
        const height = values.length, width = values[0].length;
        if (!(item.sheet === 'Contacts' && item.row >= 2 && item.column >= 1 ||
              item.sheet === 'Info' && item.row === 3 && item.column === 8 && height === 1 && width === 1)) {
          throw new Error('Related writes are restricted to Contacts data and Info!H3.');
        }
        if (item.row + height - 1 > sheet.getMaxRows() || item.column + width - 1 > sheet.getMaxColumns()) {
          throw new Error('Add enough empty rows/columns to ' + item.sheet + ' before refreshing.');
        }
        values.forEach((row, r) => {
          if (!Array.isArray(row) || row.length !== width) throw new Error('Related write has uneven rows.');
          row.forEach((value, c) => {
            if (value == null || !['string', 'number', 'boolean'].includes(typeof value) && !(value instanceof Date) ||
                typeof value === 'number' && !Number.isFinite(value) || value instanceof Date && isNaN(value.getTime())) {
              throw new Error('Invalid value in related write.');
            }
            const token = JSON.stringify([item.sheet, item.row + r, item.column + c]);
            if (occupied.has(token)) throw new Error('Overlapping related writes.');
            occupied.add(token);
          });
        });
        const range = sheet.getRange(item.row, item.column, height, width);
        const before = range.getValues(), formulas = range.getFormulas();
        if (formulas.some(row => row.some(Boolean))) throw new Error('Related write would overwrite a formula in ' + item.sheet + '.');
        ops.push({range, values, before, formulas: null, beforeFormulas: null});
        relatedCells += height * width;
      });
      const dryRun = !Array.isArray(packet) && packet.dryRun === true;
      if (ops.length && !dryRun) apply(ops);
      return {cells: changes.length, students: keys.size, relatedCells, dryRun};
    });
  }

  function matchingTriggers(ss) {
    return ScriptApp.getProjectTriggers().filter(trigger =>
      trigger.getHandlerFunction() === HANDLER && trigger.getTriggerSourceId() === ss.getId() &&
      trigger.getEventType() === ScriptApp.EventType.ON_EDIT);
  }

  function install() {
    const ss = workbook();
    return withLock(() => {
      const all = indexes(ss);
      const triggers = matchingTriggers(ss);
      if (!triggers.length) ScriptApp.newTrigger(HANDLER).forSpreadsheet(ss).onEdit().create();
      // Only this handler's duplicate triggers owned by the current installer.
      triggers.slice(1).forEach(trigger => ScriptApp.deleteTrigger(trigger));
      const result = {installed: true, students: all[0].rows.size, tabs: TABS.slice()};
      console.log(JSON.stringify(result));
      ss.toast('Automatic roster sync is installed. Test one edit and one clear.', 'Roster sync', 10);
      return result;
    });
  }

  function check() {
    const ss = workbook();
    return withLock(() => {
      const all = indexes(ss);
      const count = matchingTriggers(ss).length;
      const result = {installedForCurrentUser: count === 1, triggerCount: count,
        students: all[0].rows.size, tabs: TABS.slice()};
      console.log(JSON.stringify(result));
      ss.toast(count === 1 ? 'Sync trigger and roster structure are valid.' :
        'This account has ' + count + ' sync triggers. Have the installer run install10RosterSync.', 'Roster sync', 10);
      return result;
    });
  }

  function uninstall() {
    const ss = workbook();
    matchingTriggers(ss).forEach(trigger => ScriptApp.deleteTrigger(trigger));
    ss.toast('This account\'s roster sync trigger was removed.', 'Roster sync', 6);
  }

  function edit(event) {
    if (!event || !event.range) throw new Error('Do not run handle10RosterEdit manually; run install10RosterSync.');
    try { return sync(event.range); }
    catch (error) {
      console.error('Roster sync: ' + error.message);
      try { event.range.getSheet().getParent().toast(error.message, 'Roster sync needs attention', 15); }
      catch (toastError) { console.error('Unable to show sync notice.'); }
      throw error; // Visible in Executions and normal trigger failure notices.
    }
  }

  return Object.freeze({install, check, uninstall, edit, sync, refresh, updateFields});
})();

function install10RosterSync() { return HR10SYNC.install(); }
function handle10RosterEdit(event) { return HR10SYNC.edit(event); }
function sync10RosterRange(range) { return HR10SYNC.sync(range); }
function check10RosterSync() { return HR10SYNC.check(); }
function refresh10RosterFromMaster() { return HR10SYNC.refresh(); }
function uninstall10RosterSync() { return HR10SYNC.uninstall(); }

function update10RosterFields(planner) { return HR10SYNC.updateFields(planner); }
