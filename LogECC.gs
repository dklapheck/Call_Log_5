/**
 * 2Roster ORN - ECC Tools (one-click PowerSchool handoff)
 *
 * Required named ranges:
 *   ECC_Dates
 *   ECC_Notes
 *   Old_ECC_Dates_and_Notes
 *   ECC_Student_Number
 *
 * IMPORTANT:
 * - Keep your single combined onOpen() in Menus.gs.
 * - Do NOT add another onOpen() to this file.
 * - The Chrome extension must be v3 or later so it can detect the
 *   ECC_HANDOFF_V1 toast on this specific Google Sheet.
 */

const ECC_HANDOFF_PREFIX_ = 'ECC_HANDOFF_V1:';


function archiveAndOpenPowerSchoolECC() {
  const data = getCurrentECCRowData_();

  if (!data) return;

  const archived = archiveECCData_(data, false);

  if (!archived) return;

  const payload = {
    v: 1,
    studentNumber: data.studentNumber,
    date: data.date,
    note: data.note
  };

  const encoded = Utilities
    .base64EncodeWebSafe(
      JSON.stringify(payload),
      Utilities.Charset.UTF_8
    )
    .replace(/=+$/g, '');

  // Make sure the archive write is committed before handoff.
  SpreadsheetApp.flush();

  // No modal dialog. The extension running ONLY on 2Roster ORN watches
  // for this short-lived toast and opens PowerSchool in a new tab.
  data.ss.toast(
    ECC_HANDOFF_PREFIX_ + encoded,
    'ECC Tools',
    10
  );
}


function archiveCurrentECCNote() {
  const data = getCurrentECCRowData_();

  if (!data) return;

  archiveECCData_(data, true);
}


function getCurrentECCRowData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== 'ECC') {
    ui.alert('Please select a student row on the ECC tab first.');
    return null;
  }

  const activeRow = sheet.getActiveCell().getRow();

  if (activeRow <= 1) {
    ui.alert('Please select a student row, not the header row.');
    return null;
  }

  const dateRange = ss.getRangeByName('ECC_Dates');
  const noteRange = ss.getRangeByName('ECC_Notes');
  const oldRange = ss.getRangeByName('Old_ECC_Dates_and_Notes');
  const studentRange = ss.getRangeByName('ECC_Student_Number');

  const missing = [];

  if (!dateRange) missing.push('ECC_Dates');
  if (!noteRange) missing.push('ECC_Notes');
  if (!oldRange) missing.push('Old_ECC_Dates_and_Notes');
  if (!studentRange) missing.push('ECC_Student_Number');

  if (missing.length) {
    ui.alert(
      'These required named ranges were not found:\n\n' +
      missing.join('\n')
    );
    return null;
  }

  const dateCell = getCellForSheetRow_(dateRange, activeRow);
  const noteCell = getCellForSheetRow_(noteRange, activeRow);
  const oldCell = getCellForSheetRow_(oldRange, activeRow);
  const studentCell = getCellForSheetRow_(studentRange, activeRow);

  if (!dateCell || !noteCell || !oldCell || !studentCell) {
    ui.alert(
      'The selected row is outside one or more ECC named ranges.'
    );
    return null;
  }

  const date = dateCell.getDisplayValue().trim();
  const note = noteCell.getDisplayValue().trim();
  const oldText = oldCell.getDisplayValue();
  const studentNumber = studentCell.getDisplayValue().trim();

  if (!studentNumber) {
    ui.alert('No student number was found for the selected row.');
    return null;
  }

  if (!date) {
    ui.alert(
      'No ECC Date is entered for student ' + studentNumber + '.'
    );
    return null;
  }

  if (!note) {
    ui.alert(
      'No ECC Note is entered for student ' + studentNumber + '.'
    );
    return null;
  }

  return {
    ss,
    sheet,
    activeRow,
    date,
    note,
    oldText,
    oldCell,
    studentNumber
  };
}


function getCellForSheetRow_(namedRange, sheetRow) {
  const firstRow = namedRange.getRow();
  const offset = sheetRow - firstRow;

  if (offset < 0 || offset >= namedRange.getNumRows()) {
    return null;
  }

  return namedRange.getCell(offset + 1, 1);
}


function archiveECCData_(data, showToast) {
  const newEntry = '-- ' + data.date + ': ' + data.note;
  const existingText = String(data.oldText || '');

  const entries = existingText
    .split('\n')
    .map(value => value.trim());

  if (entries.includes(newEntry)) {
    if (showToast) {
      data.ss.toast(
        'This ECC date/note is already in the archive.',
        'ECC Tools',
        4
      );
    }

    // A duplicate should NOT block PowerSchool launch.
    return true;
  }

  const updatedText = existingText.trim()
    ? existingText.replace(/\s+$/g, '') + '\n' + newEntry
    : newEntry;

  data.oldCell.setValue(updatedText);
  data.oldCell.setWrap(true);

  if (showToast) {
    data.ss.toast(
      'ECC note archived for student ' + data.studentNumber + '.',
      'ECC Tools',
      4
    );
  }

  return true;
}
