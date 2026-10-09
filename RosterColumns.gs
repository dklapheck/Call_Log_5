/** Add roster columns v1.1 (extends each tab's filter over new columns). Requires Config.gs; menu in Menus.gs.
 * Update > Call Tools > Add columns to all four roster tabs.
 * Asks for one or more new headers (separated by commas) and the existing
 * header to insert them after, then inserts the same columns in Master, SCC,
 * ECC and STAR in one locked step. Each tab keeps its own column order: the new
 * columns go after that header wherever it sits in that tab. If any tab fails,
 * the columns already added are removed again. New cells start blank, without
 * the neighbouring column's dropdowns. Student data is not changed.
 * Sheets only grows a filter for columns inserted inside it, so when new columns
 * land past a tab's filter the filter is rebuilt to cover them, keeping every
 * column's filter conditions. Update > Call Tools > Extend roster filters to all
 * columns does the same for columns added earlier.
 */
const HR10COLUMNS = (() => {
  const TABS = HR10CFG.ROSTER_TABS;
  const RESERVED = ['Student Number', 'Name', 'Subject Line'];

  /** "A, B ,C" -> ['A', 'B', 'C'], validated against Master's headers. */
  function parse(input, existing) {
    const names = HR10CFG.text(input).split(',').map(HR10CFG.normalize).filter(Boolean);
    if (!names.length) throw new Error('Enter at least one column name. Separate several with commas.');
    const taken = new Set(existing.map(h => h.toLowerCase()));
    const seen = new Set();
    names.forEach(name => {
      const key = name.toLowerCase();
      if (RESERVED.some(r => r.toLowerCase() === key) || taken.has(key)) throw new Error('A column named "' + name + '" already exists.');
      if (seen.has(key)) throw new Error('"' + name + '" is listed twice.');
      seen.add(key);
    });
    return names;
  }

  /** headersByTab: {tab: normalized row-1 headers}. after: '' = at the end.
   * Returns {tab: 1-based column to insert after}. Headers must already match Master. */
  function positions(headersByTab, after) {
    const master = headersByTab.Master.filter(Boolean);
    const result = {};
    TABS.forEach(tab => {
      const headers = headersByTab[tab];
      if (headers.some(h => !h)) throw new Error(tab + ' has a blank column header. Fix it first.');
      if (headers.length !== master.length || master.some(h => !headers.includes(h))) {
        throw new Error(tab + ': headers differ from Master. Run Check roster sync and fix that first.');
      }
      const target = HR10CFG.normalize(after);
      if (!target) { result[tab] = headers.length; return; }
      const i = headers.indexOf(target);
      if (i < 0) throw new Error('No column named "' + target + '". Use an existing header exactly, or leave it blank.');
      result[tab] = i + 1;
    });
    return result;
  }

  /** Rebuilds a tab's basic filter so it reaches lastColumn, keeping each
   * column's conditions. Returns true if it changed. No filter: nothing to do. */
  function extendFilter(sheet, lastColumn) {
    const filter = sheet.getFilter();
    if (!filter) return false;
    const range = filter.getRange();
    const first = range.getColumn(), end = range.getLastColumn();
    if (end >= lastColumn) return false;
    const criteria = [];
    for (let c = first; c <= end; c++) {
      const value = filter.getColumnFilterCriteria(c);
      if (value) criteria.push([c, value]);
    }
    filter.remove();
    const rebuilt = sheet.getRange(range.getRow(), first, range.getNumRows(), lastColumn - first + 1).createFilter();
    criteria.forEach(([c, value]) => rebuilt.setColumnFilterCriteria(c, value));
    return true;
  }

  function headersOf(sheet) {
    const width = sheet.getLastColumn();
    return width ? sheet.getRange(1, 1, 1, width).getValues()[0].map(HR10CFG.normalize) : [];
  }

  function insert(ss, names, after) {
    return HR10CFG.withLock(() => {
      const sheets = {}, headersByTab = {};
      TABS.forEach(tab => {
        sheets[tab] = ss.getSheetByName(tab);
        if (!sheets[tab]) throw new Error(tab + ' tab is missing.');
        headersByTab[tab] = headersOf(sheets[tab]);
      });
      parse(names.join(','), headersByTab.Master); // re-check under the lock
      const where = positions(headersByTab, after);
      const done = [];
      try {
        TABS.forEach(tab => {
          const sheet = sheets[tab], column = where[tab];
          sheet.insertColumnsAfter(column, names.length);
          done.push(tab);
          const header = sheet.getRange(1, column + 1, 1, names.length);
          header.setValues(HR10CFG.prepareValues(header, [names]));
          if (sheet.getMaxRows() > 1) {
            sheet.getRange(2, column + 1, sheet.getMaxRows() - 1, names.length).clearContent().clearDataValidations();
          }
        });
        SpreadsheetApp.flush();
        TABS.forEach(tab => {
          const h = headersOf(sheets[tab]);
          if (names.some(n => !h.includes(n))) throw new Error(tab + ': new headers were not saved.');
        });
      } catch (error) {
        done.forEach(tab => {
          try { sheets[tab].deleteColumns(where[tab] + 1, names.length); }
          catch (e) { error.message += ' Could not undo ' + tab + ': ' + e.message; }
        });
        throw new Error('No columns were added (' + error.message + ')');
      }
      // Columns are in; a filter problem is reported but never undoes them.
      const filters = {extended: [], failed: []};
      TABS.forEach(tab => {
        try { if (extendFilter(sheets[tab], where[tab] + names.length)) filters.extended.push(tab); }
        catch (e) { filters.failed.push(tab + ' (' + e.message + ')'); }
      });
      return {where, filters};
    }, 'Roster tools are busy; no columns were added. Please retry.');
  }

  function run() {
    const ss = HR10CFG.workbook(), ui = SpreadsheetApp.getUi();
    const first = ui.prompt('Add roster columns', 'New column names, separated by commas.\nExample: Attendance Notes, Tutoring', ui.ButtonSet.OK_CANCEL);
    if (first.getSelectedButton() !== ui.Button.OK) return;
    const names = parse(first.getResponseText(), headersOf(ss.getSheetByName('Master')));
    const second = ui.prompt('Where?', 'Insert after which existing column? Type its header exactly (example: STAR Notes).\nLeave blank to add at the end of each tab.', ui.ButtonSet.OK_CANCEL);
    if (second.getSelectedButton() !== ui.Button.OK) return;
    const after = HR10CFG.normalize(second.getResponseText());
    const ok = ui.alert('Add ' + names.length + ' column' + (names.length > 1 ? 's' : '') + '?',
      names.join(', ') + '\n\nwill be added ' + (after ? 'after "' + after + '"' : 'at the end') +
      ' in Master, SCC, ECC and STAR.', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
    let result;
    try {
      result = insert(ss, names, after);
    } catch (e) {
      HR10CFG.log('ERROR', 'AddRosterColumns', '', e.message, HR10CFG.errorDetails(e));
      throw e;
    }
    HR10CFG.log('SUCCESS', 'AddRosterColumns', '', 'Added ' + names.join(', ') + (after ? ' after ' + after : ' at end'), '');
    const f = result.filters;
    ss.toast(names.join(', ') + ' added to all four roster tabs.' +
      (f.extended.length ? ' Filter extended on ' + f.extended.join(', ') + '.' : '') +
      (f.failed.length ? ' Could not extend the filter on ' + f.failed.join(', ') + '; use Extend roster filters.' : ''),
      'Roster columns', 10);
  }

  /** Extends each roster tab's filter to its last header column. */
  function fixFilters() {
    const ss = HR10CFG.workbook();
    const changed = HR10CFG.withLock(() => TABS.filter(tab => {
      const sheet = ss.getSheetByName(tab);
      if (!sheet) throw new Error(tab + ' tab is missing.');
      return extendFilter(sheet, headersOf(sheet).length);
    }), 'Roster tools are busy. Please retry.');
    HR10CFG.log('SUCCESS', 'ExtendRosterFilters', '', 'Extended: ' + (changed.join(', ') || 'none'), '');
    ss.toast(changed.length ? 'Filter now covers every column on ' + changed.join(', ') + '.' :
      'Filters already cover every column (or a tab has no filter).', 'Roster filters', 8);
    return changed;
  }

  return Object.freeze({run, parse, positions, insert, extendFilter, fixFilters});
})();

function add10RosterColumns() { return HR10COLUMNS.run(); }
function extend10RosterFilters() { return HR10COLUMNS.fixFilters(); }
