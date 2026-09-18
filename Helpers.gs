function getRequiredSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      'Required sheet not found: ' + sheetName
    );
  }

  return sheet;
}


function findHeaderColumn_(sheet, headerText) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    throw new Error(
      'No columns were found on sheet: ' +
      sheet.getName()
    );
  }

  const headers = sheet
    .getRange(
      SCC_CONFIG.roster.headerRow,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0];

  const matches = [];

  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === headerText) {
      matches.push(i + 1);
    }
  }

  if (matches.length === 0) {
    throw new Error(
      'Could not find the Roster header "' +
      headerText +
      '".'
    );
  }

  if (matches.length > 1) {
    throw new Error(
      'More than one Roster column is labeled "' +
      headerText +
      '".'
    );
  }

  return matches[0];
}


function findStudentRosterRow_(rosterSheet, studentId) {
  const studentIdColumn = findHeaderColumn_(
    rosterSheet,
    SCC_CONFIG.roster.studentIdHeader
  );

  const lastRow = rosterSheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const studentIds = rosterSheet
    .getRange(
      2,
      studentIdColumn,
      lastRow - 1,
      1
    )
    .getValues();

  for (let i = 0; i < studentIds.length; i++) {
    if (
      String(studentIds[i][0]).trim() ===
      String(studentId).trim()
    ) {
      return i + 2;
    }
  }

  return null;
}


function getStudentName_(rosterSheet, rosterRow) {
  const firstNameColumn = findHeaderColumn_(
    rosterSheet,
    SCC_CONFIG.roster.firstNameHeader
  );

  const lastNameColumn = findHeaderColumn_(
    rosterSheet,
    SCC_CONFIG.roster.lastNameHeader
  );

  const firstName = rosterSheet
    .getRange(rosterRow, firstNameColumn)
    .getDisplayValue()
    .trim();

  const lastName = rosterSheet
    .getRange(rosterRow, lastNameColumn)
    .getDisplayValue()
    .trim();

  return (firstName + ' ' + lastName).trim();
}

function findQuestionRow_(sheet, questionText) {
  const match = sheet
    .getRange('B7:B50')
    .createTextFinder(questionText)
    .matchEntireCell(true)
    .findNext();

  if (!match) {
    throw new Error(
      'Could not find this Call Entry question:\n' +
      questionText
    );
  }

  return match.getRow();
}

// Save unsuccessfull contact attempts.
function saveFailedSccAttempt_(
  ss,
  rosterSheet,
  rosterRow,
  studentName,
  contactNote
) {
  const attemptHeaders = [
    'Attempt 1',
    'Attempt 2',
    'Attempt 3',
    'Attempt 4',
    'Attempt 5'
  ];

  const completionStatuses = [
    'First Attempt',
    'Second Attempt',
    'Third Attempt',
    'Fourth Attempt',
    'Fifth Attempt'
  ];

  let attemptColumn = null;
  let attemptNumber = null;

  // Find the first empty attempt cell for this student.
  for (let i = 0; i < attemptHeaders.length; i++) {
    const column = findHeaderColumn_(
      rosterSheet,
      attemptHeaders[i]
    );

    const existingValue = rosterSheet
      .getRange(rosterRow, column)
      .getDisplayValue()
      .trim();

    if (!existingValue) {
      attemptColumn = column;
      attemptNumber = i;
      break;
    }
  }

  if (attemptColumn === null) {
    ss.toast(
      'No empty Attempt 1–5 column is available for ' +
        studentName +
        '. The contact note was not saved.',
      'Attempt Not Saved',
      8
    );
    return;
  }

  // Save the unsuccessful contact note.
  rosterSheet
    .getRange(rosterRow, attemptColumn)
    .setValue(contactNote);

  SpreadsheetApp.flush();

  // Update SCC Completion to the corresponding attempt status.
  let completionColumn;

  try {
    completionColumn = findHeaderColumn_(
      rosterSheet,
      SCC_CONFIG.roster.completionHeader
    );
  } catch (error) {
    ss.toast(
      'Attempt ' +
        (attemptNumber + 1) +
        ' saved for ' +
        studentName +
        ', but SCC Completion was not updated.',
      'Attempt Saved',
      8
    );
    return;
  }

  rosterSheet
    .getRange(rosterRow, completionColumn)
    .setValue(
      completionStatuses[attemptNumber]
    );

  SpreadsheetApp.flush();

  ss.toast(
    'Attempt ' +
      (attemptNumber + 1) +
      ' saved for ' +
      studentName +
      '.',
    'Contact Attempt Saved',
    4
  );
}
