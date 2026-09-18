/*
5Roster ORN — optional Apps Script actions

Open the working Google Sheet, choose Extensions > Apps Script, paste this text
into a new script file, and save. Run installRosterMenu once and authorize it.
The Roster Tools menu then appears when you reopen the spreadsheet.

Existing Engageli functions and onOpen code can remain in the project. This
script uses a separate installable open trigger, so it does not replace onOpen.

The WIG menu action saves current Dash!B22:G22 as values in the matching
Tuesday row of WIG. To schedule it automatically, run
enableTuesdayWigSnapshots once. The scheduled run occurs on Tuesdays between
4:00 and 5:00 PM in the script time zone. Confirm that project time zone is
America/Los_Angeles before enabling it. No past row is overwritten.

PowerSchool: Call Entry!P2 is a copy-ready draft. Review and paste it into the
school system yourself. This script does not send or post information there.
*/

function installRosterMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fn = 'showRosterMenu';
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === fn) ScriptApp.deleteTrigger(trigger);
  }
  ScriptApp.newTrigger(fn).forSpreadsheet(ss).onOpen().create();
  showRosterMenu();
}

function showRosterMenu() {
  SpreadsheetApp.getUi().createMenu('Roster Tools')
    .addItem('Save SCC call note', 'saveSccCallEntry')
    .addItem('Save ECC entry', 'saveEccEntry')
    .addItem('Save this week’s WIG snapshot', 'saveWigSnapshot')
    .addToUi();
}

function saveSccCallEntry() {
  withRosterLock_(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const entry = requiredSheet_(ss, 'Call Entry');
    const scc = requiredSheet_(ss, 'SCC');
    const id = String(entry.getRange('O2').getDisplayValue()).trim();
    const note = String(entry.getRange('D1').getDisplayValue()).trim();
    const todo = String(entry.getRange('I1').getDisplayValue()).trim();
    if (!id || !note) throw new Error('Choose a student and finish a call note first.');
    if (/\bERROR:|Choose Yes or No/i.test(note)) {
      throw new Error('Review the call-entry responses before saving the note.');
    }
    const row = rowForId_(scc, 2, id);
    if (!row) throw new Error('This Student Number is not in SCC: ' + id);
    const notesCell = scc.getRange(row, 20); // T: SCC Notes
    const todosCell = scc.getRange(row, 21); // U: SCC To Do
    const oldNotes = String(notesCell.getValue() || '');
    if (oldNotes.includes(note)) throw new Error('This exact note is already in SCC.');
    notesCell.setValue(joinEntries_(oldNotes, note));
    if (todo) {
      const oldTodos = String(todosCell.getValue() || '');
      if (!oldTodos.includes(todo)) todosCell.setValue(joinEntries_(oldTodos, todo));
    }
    SpreadsheetApp.getUi().alert('SCC note saved for Student Number ' + id + '. Review the SCC completion status separately.');
  });
}

function saveEccEntry() {
  withRosterLock_(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const form = requiredSheet_(ss, 'ECC Entry');
    const ecc = requiredSheet_(ss, 'ECC');
    const scc = requiredSheet_(ss, 'SCC');
    const id = String(form.getRange('B2').getDisplayValue()).trim();
    const date = form.getRange('B3').getValue();
    const result = String(form.getRange('B4').getDisplayValue()).trim();
    const note = String(form.getRange('B5').getValue() || '').trim();
    const todo = String(form.getRange('B6').getValue() || '').trim();
    const link = String(form.getRange('B7').getValue() || '').trim();
    if (!id || !(date instanceof Date) || !['Conversation', 'Attempt'].includes(result) || !note) {
      throw new Error('Enter a Student Number, date, result, and note in ECC Entry.');
    }
    const sccRow = rowForId_(scc, 2, id);
    if (!sccRow) throw new Error('This Student Number is not in SCC: ' + id);
    let eccRow = rowForId_(ecc, 2, id);
    if (!eccRow) {
      eccRow = Math.max(ecc.getLastRow() + 1, 2);
      if (eccRow > ecc.getMaxRows()) ecc.insertRowsAfter(ecc.getMaxRows(), 1);
      // Identity fields retain the layout of the existing ECC worksheet.
      const src = scc.getRange(sccRow, 1, 1, 19).getValues()[0];
      ecc.getRange(eccRow, 1).setValue(src[0]);  // Preferred Name
      ecc.getRange(eccRow, 2).setValue(src[1]);  // Student Number
      ecc.getRange(eccRow, 3).setValue(scc.getRange(sccRow, 20).getValue());
      ecc.getRange(eccRow, 9).setValue(src[17]); // Small Group
      ecc.getRange(eccRow, 10).setValue(src[18]); // Student Email
      ecc.getRange(eccRow, 11, 1, 9).setValues([[src[2], src[3], src[4], src[5], src[6], src[7], src[8], src[9], src[10]]]);
      ecc.getRange(eccRow, 20).setValue(src[11]); // Parent Language
      ecc.getRange(eccRow, 21).setFormula('=CONCATENATE(L' + eccRow + '," ",K' + eccRow + ')');
      ecc.getRange(eccRow, 22).setFormula('=LEFT(L' + eccRow + ',1)&" "&K' + eccRow + '&", "&B' + eccRow + '&", "&O' + eccRow);
    }
    const dateLabel = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'M/d/yyyy');
    const text = note + (todo ? '\nTo do: ' + todo : '');
    const oldRecent = String(ecc.getRange(eccRow, 5).getValue() || '');
    const oldHistory = String(ecc.getRange(eccRow, 6).getValue() || '');
    const rendered = '-- ' + dateLabel + ': ' + text;
    if (result === 'Conversation') {
      if (oldRecent === text || oldHistory.includes(rendered)) {
        throw new Error('This ECC note appears to have been saved already.');
      }
      if (oldRecent) {
        const previousDate = ecc.getRange(eccRow, 4).getValue();
        const previousLabel = previousDate instanceof Date
          ? Utilities.formatDate(previousDate, ss.getSpreadsheetTimeZone(), 'M/d/yyyy')
          : 'Earlier';
        ecc.getRange(eccRow, 6).setValue(joinEntries_(oldHistory, '-- ' + previousLabel + ': ' + oldRecent));
      }
      ecc.getRange(eccRow, 4).setValue(date);
      ecc.getRange(eccRow, 5).setValue(text);
    } else {
      const first = ecc.getRange(eccRow, 7);
      const second = ecc.getRange(eccRow, 8);
      if (String(first.getValue() || '').includes(rendered) || String(second.getValue() || '').includes(rendered) || oldHistory.includes(rendered)) {
        throw new Error('This attempt appears to have been saved already.');
      }
      if (!first.getValue()) first.setValue(rendered);
      else if (!second.getValue()) second.setValue(rendered);
      else ecc.getRange(eccRow, 6).setValue(joinEntries_(oldHistory, rendered));
    }
    if (link) ecc.getRange(eccRow, 23).setValue(link); // W: OneNote link
    SpreadsheetApp.getUi().alert('ECC entry saved for Student Number ' + id + '.');
  });
}

function saveWigSnapshot() {
  try {
    const row = snapshotWig_();
    SpreadsheetApp.getUi().alert('WIG columns B:G saved as values on row ' + row + '.');
  } catch (error) {
    SpreadsheetApp.getUi().alert(error.message);
  }
}

function scheduledWigSnapshot() {
  snapshotWig_(); // Errors appear in Apps Script execution history.
}

function enableTuesdayWigSnapshots() {
  const fn = 'scheduledWigSnapshot';
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === fn) ScriptApp.deleteTrigger(trigger);
  }
  ScriptApp.newTrigger(fn).timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(16).create();
}

function snapshotWig_() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dash = requiredSheet_(ss, 'Dash');
    const wig = requiredSheet_(ss, 'WIG');
    const date = dash.getRange('A22').getValue();
    if (!(date instanceof Date)) throw new Error('Dash A22 has no valid week date.');
    const key = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    const dates = wig.getRange(2, 1, Math.max(wig.getLastRow() - 1, 1), 1).getValues();
    const match = dates.findIndex(row => row[0] instanceof Date &&
      Utilities.formatDate(row[0], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') === key);
    if (match < 0) throw new Error('Add the date ' + key + ' to WIG column A first.');
    const rowNumber = match + 2;
    const target = wig.getRange(rowNumber, 2, 1, 6);
    if (target.getValues()[0].some(value => value !== '')) {
      throw new Error('WIG row ' + rowNumber + ' already has a snapshot. It was not changed.');
    }
    const values = dash.getRange('B22:G22').getValues()[0];
    if (values.slice(0, 2).some(value => typeof value !== 'number') ||
        values.slice(4).some(value => typeof value !== 'number')) {
      throw new Error('Review the live WIG values on Dash row 22 before saving.');
    }
    target.setValues([values]);
    target.setNote('Snapshot saved ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'M/d/yyyy h:mm a') + '.');
    return rowNumber;
  } finally {
    lock.releaseLock();
  }
}

function withRosterLock_(work) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try { work(); }
  catch (error) { SpreadsheetApp.getUi().alert(error.message); }
  finally { lock.releaseLock(); }
}

function requiredSheet_(ss, title) {
  const sheet = ss.getSheetByName(title);
  if (!sheet) throw new Error('Missing sheet: ' + title);
  return sheet;
}

function rowForId_(sheet, column, id) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const values = sheet.getRange(2, column, last - 1, 1).getDisplayValues();
  const index = values.findIndex(row => String(row[0]).trim() === String(id).trim());
  return index < 0 ? 0 : index + 2;
}

function joinEntries_(oldText, newText) {
  return oldText.trim() ? oldText.trim() + '\n\n' + newText : newText;
}
