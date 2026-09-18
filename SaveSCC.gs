function saveSccToRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const callSheet = getRequiredSheet_(
    ss,
    SCC_CONFIG.sheets.callEntry
  );

  const rosterSheet = getRequiredSheet_(
    ss,
    SCC_CONFIG.sheets.roster
  );

  // --------------------------------------------------
  // GET SELECTED STUDENT
  // --------------------------------------------------

  const studentId = callSheet
    .getRange(SCC_CONFIG.callEntry.studentId)
    .getValue();

  if (!studentId) {
    ss.toast(
      'Select a student before saving.',
      'SCC Not Saved',
      4
    );
    return;
  }

  const rosterRow = findStudentRosterRow_(
    rosterSheet,
    studentId
  );

  if (!rosterRow) {
    ss.toast(
      'Student ID ' +
        studentId +
        ' was not found on the Roster sheet.',
      'SCC Not Saved',
      5
    );
    return;
  }

  const studentName = getStudentName_(
    rosterSheet,
    rosterRow
  );

  // --------------------------------------------------
  // CHECK WHETHER CONTACT WAS SUCCESSFUL
  // --------------------------------------------------

  const contactRow = findQuestionRow_(
    callSheet,
    'Successful contact?'
  );

  const successfulContact = callSheet
    .getRange(contactRow, 4) // Yes checkbox in column D
    .getValue() === true;

  const unsuccessfulContact = callSheet
    .getRange(contactRow, 6) // No checkbox in column F
    .getValue() === true;

  if (successfulContact && unsuccessfulContact) {
    ss.toast(
      'Choose either Yes or No for Successful contact?',
      'SCC Not Saved',
      5
    );
    return;
  }

  if (!successfulContact && !unsuccessfulContact) {
    ss.toast(
      'Answer Successful contact? before saving.',
      'SCC Not Saved',
      5
    );
    return;
  }

  // --------------------------------------------------
  // GET GENERATED CONTACT NOTE
  // --------------------------------------------------

  const sccNote = callSheet
    .getRange(SCC_CONFIG.callEntry.note)
    .getDisplayValue()
    .trim();

  if (!sccNote) {
    ss.toast(
      'There is no contact note to save for ' +
        studentName +
        '.',
      'SCC Not Saved',
      4
    );
    return;
  }

  // --------------------------------------------------
  // UNSUCCESSFUL CONTACT
  // --------------------------------------------------

  if (unsuccessfulContact) {
    saveFailedSccAttempt_(
      ss,
      rosterSheet,
      rosterRow,
      studentName,
      sccNote
    );

    return;
  }

  // --------------------------------------------------
  // SUCCESSFUL CONTACT
  // GET SCC TO DO
  // --------------------------------------------------

  const sccToDo = callSheet
    .getRange(SCC_CONFIG.callEntry.toDo)
    .getDisplayValue()
    .trim();

  // --------------------------------------------------
  // FIND SCC NOTES AND SCC TO DO COLUMNS
  // --------------------------------------------------

  let notesColumn;
  let toDoColumn;

  try {
    notesColumn = findHeaderColumn_(
      rosterSheet,
      SCC_CONFIG.roster.notesHeader
    );

    toDoColumn = findHeaderColumn_(
      rosterSheet,
      SCC_CONFIG.roster.toDoHeader
    );
  } catch (error) {
    ss.toast(
      'SCC was not saved because the Roster save columns could not be verified. ' +
        error.message,
      'SCC Not Saved',
      8
    );
    return;
  }

  // --------------------------------------------------
  // SAVE SUCCESSFUL SCC FIRST
  // --------------------------------------------------

  rosterSheet
    .getRange(rosterRow, notesColumn)
    .setValue(sccNote);

  rosterSheet
    .getRange(rosterRow, toDoColumn)
    .setValue(sccToDo);

  SpreadsheetApp.flush();

  // --------------------------------------------------
  // FIND SCC COMPLETION COLUMN
  // --------------------------------------------------

  let completionColumn;

  try {
    completionColumn = findHeaderColumn_(
      rosterSheet,
      SCC_CONFIG.roster.completionHeader
    );
  } catch (error) {
    ss.toast(
      'SCC saved for ' +
        studentName +
        ', but not marked Completed because the "' +
        SCC_CONFIG.roster.completionHeader +
        '" column could not be found.',
      'SCC Saved — Completion Not Updated',
      8
    );
    return;
  }

  // --------------------------------------------------
  // MARK COMPLETED
  // --------------------------------------------------

  rosterSheet
    .getRange(rosterRow, completionColumn)
    .setValue(
      SCC_CONFIG.roster.completedValue
    );

  SpreadsheetApp.flush();

  ss.toast(
    'SCC saved and marked Completed for ' +
      studentName +
      '.',
    'SCC Saved',
    4
  );
}