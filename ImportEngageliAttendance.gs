/**
 * ENGAGELI ATTENDANCE IMPORTER
 * Copies Engageli "Presence" durations into the Roster sheet.
 * Students are matched by the leading student number in the email:
 *   11101615_k12@example.com -> 11101615
 */

const ENGAGELI_CONFIG = {
  sheets: {
    roster: 'Roster',
    data: 'Engageli Data',
    settings: 'Instructions and Settings'
  },

  headers: {
    rosterStudentId: 'Student Number',
    rosterLastName: 'LAST NAME',
    rosterFirstName: 'FIRST NAME',
    rosterName: 'Name',
    dataName: 'name',
    dataEmail: 'email',
    dataRole: 'role'
  },

  settings: {
    sectionTitle: 'Engageli Attendance Settings',
    teacherCheckboxLabel: 'Import teacher row at bottom of Roster',
    defaultTeacherCheckboxValue: true,
    firstPreferredRow: 15
  },

  source: {
    presenceSuffix: ' Presence',
    studentRoles: ['student'],
    teacherRoles: ['prof', 'teacher', 'instructor']
  },

  teacher: {
    rosterMarker: 'TEACHER',
    defaultName: 'Teacher Benchmark'
  },

  colors: {
    full: '#D9EAD3',        // 90% or more
    most: '#E2F0D9',        // 75%–89%
    half: '#FFF2CC',        // 50%–74%
    brief: '#FCE5CD',       // 25%–49%
    veryBrief: '#F4CCCC',   // Below 25%
    unavailable: '#D9D9D9',
    teacher: '#CFE2F3',
    header: '#1F4E78',
    headerText: '#FFFFFF'
  }
};


/**
 * See Menus.gs for the onOpen() function,
 * that creates one menue button.
 */

/**
 * Creates or repairs the checkbox and color key.
 */
function setupEngageliAttendanceSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    engEnsureSettings_(ss);

    ss.toast(
      'The teacher-row checkbox and attendance color key are ready.',
      'Engageli Settings Ready',
      5
    );
  } catch (error) {
    ss.toast(
      error.message,
      'Settings Not Created',
      10
    );

    throw error;
  }
}


/**
 * Main attendance-import command.
 */
function updateEngageliAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const settings = engEnsureSettings_(ss);

    const rosterSheet = engGetRequiredSheet_(
      ss,
      ENGAGELI_CONFIG.sheets.roster
    );

    const dataSheet = engGetRequiredSheet_(
      ss,
      ENGAGELI_CONFIG.sheets.data
    );

    const result = engImportAttendance_(
      rosterSheet,
      dataSheet,
      settings.importTeacherRow
    );

    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert(
      'Engageli Attendance Updated',
      engBuildSummary_(result),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    ss.toast(
      error.message,
      'Attendance Not Updated',
      10
    );

    throw error;
  }
}


/**
 * Creates or repairs the settings area.
 */
function engEnsureSettings_(ss) {
  const sheet = engGetRequiredSheet_(
    ss,
    ENGAGELI_CONFIG.sheets.settings
  );

  const config = ENGAGELI_CONFIG.settings;
  const lastRow = Math.max(sheet.getLastRow(), 1);

  const existingLabel = sheet
    .getRange(1, 1, lastRow, 1)
    .createTextFinder(config.teacherCheckboxLabel)
    .matchEntireCell(true)
    .findNext();

  let settingRow;

  if (existingLabel) {
    settingRow = existingLabel.getRow();
  } else {
    const proposedSectionRow = Math.max(
      sheet.getLastRow() + 2,
      config.firstPreferredRow
    );

    settingRow = proposedSectionRow + 1;

    engEnsureSheetRows_(
      sheet,
      proposedSectionRow + 9
    );

    sheet
      .getRange(proposedSectionRow, 1, 1, 2)
      .setBackground(
        ENGAGELI_CONFIG.colors.header
      )
      .setFontColor(
        ENGAGELI_CONFIG.colors.headerText
      )
      .setFontWeight('bold');

    sheet
      .getRange(proposedSectionRow, 1)
      .setValue(config.sectionTitle);

    sheet
      .getRange(settingRow, 1)
      .setValue(config.teacherCheckboxLabel);
  }

  engEnsureSheetRows_(
    sheet,
    settingRow + 8
  );

  const checkboxCell = sheet.getRange(
    settingRow,
    2
  );

  const priorCheckboxValue =
    checkboxCell.getValue();

  checkboxCell.insertCheckboxes();

  if (
    priorCheckboxValue === true ||
    priorCheckboxValue === false
  ) {
    checkboxCell.setValue(
      priorCheckboxValue
    );
  } else {
    checkboxCell.setValue(
      config.defaultTeacherCheckboxValue
    );
  }

  checkboxCell.setNote(
    'Checked: add or update a blue teacher benchmark row at the bottom ' +
      'of the Roster. Unchecked: remove the script-created teacher row. ' +
      'Teacher time is still used for student color coding.'
  );

  sheet
    .getRange(settingRow, 1, 1, 2)
    .setWrap(true);

  const legendValues = [
    [
      'Color key: student time compared with teacher time',
      'Color'
    ],
    ['90% or more', ''],
    ['75%–89%', ''],
    ['50%–74%', ''],
    ['25%–49%', ''],
    ['Below 25%', ''],
    ['Teacher time unavailable', ''],
    ['Teacher benchmark row', '']
  ];

  const legendStartRow = settingRow + 1;

  const legendRange = sheet.getRange(
    legendStartRow,
    1,
    legendValues.length,
    2
  );

  legendRange
    .setValues(legendValues)
    .setWrap(true);

  sheet
    .getRange(legendStartRow, 1, 1, 2)
    .setFontWeight('bold')
    .setBackground('#EAF2F8');

  const legendColors = [
    ENGAGELI_CONFIG.colors.full,
    ENGAGELI_CONFIG.colors.most,
    ENGAGELI_CONFIG.colors.half,
    ENGAGELI_CONFIG.colors.brief,
    ENGAGELI_CONFIG.colors.veryBrief,
    ENGAGELI_CONFIG.colors.unavailable,
    ENGAGELI_CONFIG.colors.teacher
  ];

  for (
    let i = 0;
    i < legendColors.length;
    i++
  ) {
    sheet
      .getRange(
        legendStartRow + 1 + i,
        2
      )
      .setBackground(legendColors[i]);
  }

  return {
    importTeacherRow:
      checkboxCell.getValue() === true
  };
}


/**
 * Matches students and imports attendance.
 */
function engImportAttendance_(
  rosterSheet,
  dataSheet,
  importTeacherRow
) {
  const studentIdColumn =
    engFindRequiredHeaderColumn_(
      rosterSheet,
      ENGAGELI_CONFIG.headers
        .rosterStudentId
    );

  /*
   * Remove a previously created teacher row.
   * The script recreates it at the bottom if
   * the checkbox is selected.
   */
  engRemoveTeacherRows_(
    rosterSheet,
    studentIdColumn
  );

  const rosterInfo = engBuildRosterMap_(
    rosterSheet,
    studentIdColumn
  );

  const sourceInfo =
    engReadEngageliData_(dataSheet);

  const sessionColumns =
    engEnsureRosterSessionColumns_(
      rosterSheet,
      sourceInfo.sessions
    );

  const unmatchedStudentIds = new Set();
  const matchedRosterRows = new Set();

  sourceInfo.students.forEach(
    function(studentRecord, studentId) {
      const rosterRow =
        rosterInfo.studentRows.get(studentId);

      if (!rosterRow) {
        unmatchedStudentIds.add(studentId);
        return;
      }

      matchedRosterRows.add(rosterRow);
    }
  );

  const rosterLastRow = Math.max(
    rosterSheet.getLastRow(),
    2
  );

  const attendanceRowCount =
    rosterLastRow - 1;

  for (
    let s = 0;
    s < sourceInfo.sessions.length;
    s++
  ) {
    const session =
      sourceInfo.sessions[s];

    const rosterColumn =
      sessionColumns.get(session.label);

    const targetRange =
      rosterSheet.getRange(
        2,
        rosterColumn,
        attendanceRowCount,
        1
      );

    const values =
      targetRange.getValues();

    const notes =
      targetRange.getNotes();

    const backgrounds =
      targetRange.getBackgrounds();

    const teacherDuration =
      sourceInfo.teacherDurations[s];

    sourceInfo.students.forEach(
      function(studentRecord, studentId) {
        const rosterRow =
          rosterInfo.studentRows.get(
            studentId
          );

        if (!rosterRow) {
          return;
        }

        const duration =
          studentRecord.durations[s];

        if (duration === null) {
          return;
        }

        const arrayRow =
          rosterRow - 2;

        values[arrayRow][0] = duration;

        backgrounds[arrayRow][0] =
          engGetAttendanceColor_(
            duration,
            teacherDuration
          );

        notes[arrayRow][0] =
          engBuildAttendanceNote_(
            duration,
            teacherDuration
          );
      }
    );

    targetRange
      .setValues(values)
      .setNotes(notes)
      .setBackgrounds(backgrounds)
      .setNumberFormat('[h]:mm');
  }

  let teacherRowAdded = false;

  if (
    importTeacherRow &&
    sourceInfo.teacherRowsFound > 0
  ) {
    engAddTeacherRow_(
      rosterSheet,
      studentIdColumn,
      sourceInfo.teacherName,
      sourceInfo.sessions,
      sourceInfo.teacherDurations,
      sessionColumns
    );

    teacherRowAdded = true;
  }

  return {
    sessionsImported:
      sourceInfo.sessions.length,

    uniqueStudentsInEngageli:
      sourceInfo.students.size,

    studentsMatched:
      matchedRosterRows.size,

    unmatchedStudentIds:
      Array.from(
        unmatchedStudentIds
      ).sort(),

    invalidStudentEmails:
      sourceInfo.invalidStudentEmails,

    skippedNonAttendanceRoles:
      sourceInfo.skippedNonAttendanceRoles,

    duplicateStudentRowsCombined:
      sourceInfo.duplicateStudentRowsCombined,

    teacherRowsFound:
      sourceInfo.teacherRowsFound,

    teacherRowAdded:
      teacherRowAdded,

    importTeacherRowRequested:
      importTeacherRow
  };
}


/**
 * Reads Engageli Data and aggregates
 * duplicate student or teacher rows.
 */
function engReadEngageliData_(dataSheet) {
  const lastRow =
    dataSheet.getLastRow();

  const lastColumn =
    dataSheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {
    throw new Error(
      'The Engageli Data sheet does not contain attendance records.'
    );
  }

  const headers = dataSheet
    .getRange(
      1,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0]
    .map(function(header) {
      return String(header).trim();
    });

  const emailColumn =
    engFindHeaderIndexInArray_(
      headers,
      ENGAGELI_CONFIG.headers.dataEmail
    );

  const roleColumn =
    engFindHeaderIndexInArray_(
      headers,
      ENGAGELI_CONFIG.headers.dataRole
    );

  const nameColumn =
    engFindOptionalHeaderIndexInArray_(
      headers,
      ENGAGELI_CONFIG.headers.dataName
    );

  const sessions = [];
  const seenSessionLabels = new Set();

  const suffix =
    ENGAGELI_CONFIG.source
      .presenceSuffix;

  for (
    let column = 0;
    column < headers.length;
    column++
  ) {
    const header = headers[column];

    if (
      !header
        .toLowerCase()
        .endsWith(
          suffix.toLowerCase()
        )
    ) {
      continue;
    }

    const sessionLabel = header
      .slice(
        0,
        header.length - suffix.length
      )
      .trim();

    if (!sessionLabel) {
      throw new Error(
        'A Presence column on Engageli Data has no session name.'
      );
    }

    if (
      seenSessionLabels.has(
        sessionLabel
      )
    ) {
      throw new Error(
        'More than one Engageli Presence column represents the session "' +
          sessionLabel +
          '".'
      );
    }

    seenSessionLabels.add(
      sessionLabel
    );

    sessions.push({
      label: sessionLabel,
      sourceColumnIndex: column
    });
  }

  if (sessions.length === 0) {
    throw new Error(
      'No headers ending in "Presence" were found on Engageli Data.'
    );
  }

  const rows = dataSheet
    .getRange(
      2,
      1,
      lastRow - 1,
      lastColumn
    )
    .getValues();

  const students = new Map();

  const teacherDurations =
    sessions.map(function() {
      return null;
    });

  let teacherName =
    ENGAGELI_CONFIG.teacher.defaultName;

  let teacherRowsFound = 0;
  let invalidStudentEmails = 0;
  let skippedNonAttendanceRoles = 0;
  let duplicateStudentRowsCombined = 0;

  for (
    let r = 0;
    r < rows.length;
    r++
  ) {
    const row = rows[r];

    const role = String(
      row[roleColumn] || ''
    )
      .trim()
      .toLowerCase();

    const isStudent =
      ENGAGELI_CONFIG.source
        .studentRoles
        .indexOf(role) !== -1;

    const isTeacher =
      ENGAGELI_CONFIG.source
        .teacherRoles
        .indexOf(role) !== -1;

    if (
      !isStudent &&
      !isTeacher
    ) {
      skippedNonAttendanceRoles++;
      continue;
    }

    const durations =
      sessions.map(
        function(session) {
          return engParseDuration_(
            row[
              session.sourceColumnIndex
            ]
          );
        }
      );

    if (isTeacher) {
      teacherRowsFound++;

      if (
        nameColumn !== null &&
        teacherName ===
          ENGAGELI_CONFIG.teacher
            .defaultName
      ) {
        const possibleName =
          engCleanEngageliName_(
            row[nameColumn]
          );

        if (possibleName) {
          teacherName =
            possibleName;
        }
      }

      for (
        let s = 0;
        s < durations.length;
        s++
      ) {
        teacherDurations[s] =
          engMaxDuration_(
            teacherDurations[s],
            durations[s]
          );
      }

      continue;
    }

    const studentId =
      engExtractStudentIdFromEmail_(
        row[emailColumn]
      );

    if (!studentId) {
      invalidStudentEmails++;
      continue;
    }

    if (
      !students.has(studentId)
    ) {
      students.set(studentId, {
        durations: durations
      });

      continue;
    }

    duplicateStudentRowsCombined++;

    const existing =
      students.get(studentId);

    for (
      let s = 0;
      s < durations.length;
      s++
    ) {
      existing.durations[s] =
        engMaxDuration_(
          existing.durations[s],
          durations[s]
        );
    }
  }

  return {
    sessions: sessions,
    students: students,
    teacherDurations:
      teacherDurations,
    teacherName: teacherName,
    teacherRowsFound:
      teacherRowsFound,
    invalidStudentEmails:
      invalidStudentEmails,
    skippedNonAttendanceRoles:
      skippedNonAttendanceRoles,
    duplicateStudentRowsCombined:
      duplicateStudentRowsCombined
  };
}


/**
 * Builds a Student Number -> Roster row map.
 */
function engBuildRosterMap_(
  rosterSheet,
  studentIdColumn
) {
  const lastRow =
    rosterSheet.getLastRow();

  if (lastRow < 2) {
    throw new Error(
      'The Roster sheet does not contain any student rows.'
    );
  }

  const idValues =
    rosterSheet
      .getRange(
        2,
        studentIdColumn,
        lastRow - 1,
        1
      )
      .getValues();

  const studentRows = new Map();
  const duplicateIds = [];

  for (
    let i = 0;
    i < idValues.length;
    i++
  ) {
    const normalizedId =
      engNormalizeRosterId_(
        idValues[i][0]
      );

    if (!normalizedId) {
      continue;
    }

    if (
      studentRows.has(
        normalizedId
      )
    ) {
      duplicateIds.push(
        normalizedId
      );

      continue;
    }

    studentRows.set(
      normalizedId,
      i + 2
    );
  }

  if (duplicateIds.length > 0) {
    throw new Error(
      'Duplicate Student Numbers were found on the Roster: ' +
        Array.from(
          new Set(duplicateIds)
        )
          .slice(0, 10)
          .join(', ')
    );
  }

  return {
    studentRows: studentRows
  };
}


/**
 * Finds existing session columns or
 * appends missing session columns.
 */
function engEnsureRosterSessionColumns_(
  rosterSheet,
  sessions
) {
  let lastColumn =
    rosterSheet.getLastColumn();

  const headers = rosterSheet
    .getRange(
      1,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0]
    .map(function(header) {
      return String(header).trim();
    });

  const sessionColumns =
    new Map();

  for (
    let s = 0;
    s < sessions.length;
    s++
  ) {
    const sessionLabel =
      sessions[s].label;

    const matches = [];

    for (
      let c = 0;
      c < headers.length;
      c++
    ) {
      if (
        headers[c] ===
        sessionLabel
      ) {
        matches.push(c + 1);
      }
    }

    if (matches.length > 1) {
      throw new Error(
        'The Roster contains more than one column labeled "' +
          sessionLabel +
          '".'
      );
    }

    if (matches.length === 1) {
      sessionColumns.set(
        sessionLabel,
        matches[0]
      );

      continue;
    }

    const newColumn =
      lastColumn + 1;

    engEnsureSheetColumns_(
      rosterSheet,
      newColumn
    );

    const formatSourceColumn =
      Math.max(1, lastColumn);

    rosterSheet
      .getRange(
        1,
        formatSourceColumn
      )
      .copyTo(
        rosterSheet.getRange(
          1,
          newColumn
        ),
        SpreadsheetApp
          .CopyPasteType
          .PASTE_FORMAT,
        false
      );

    rosterSheet
      .getRange(1, newColumn)
      .setValue(sessionLabel)
      .setWrap(true)
      .setNote(
        'Engageli Presence duration imported by the Engageli Attendance script.'
      );

    rosterSheet.setColumnWidth(
      newColumn,
      110
    );

    lastColumn = newColumn;
    headers.push(sessionLabel);

    sessionColumns.set(
      sessionLabel,
      newColumn
    );
  }

  return sessionColumns;
}


/**
 * Adds the optional teacher benchmark
 * row at the bottom of the Roster.
 */
function engAddTeacherRow_(
  rosterSheet,
  studentIdColumn,
  teacherName,
  sessions,
  teacherDurations,
  sessionColumns
) {
  const teacherRow =
    rosterSheet.getLastRow() + 1;

  const lastColumn =
    rosterSheet.getLastColumn();

  engEnsureSheetRows_(
    rosterSheet,
    teacherRow
  );

  if (teacherRow > 2) {
    rosterSheet
      .getRange(
        teacherRow - 1,
        1,
        1,
        lastColumn
      )
      .copyTo(
        rosterSheet.getRange(
          teacherRow,
          1,
          1,
          lastColumn
        ),
        SpreadsheetApp
          .CopyPasteType
          .PASTE_FORMAT,
        false
      );
  }

  const teacherRange =
    rosterSheet.getRange(
      teacherRow,
      1,
      1,
      lastColumn
    );

  teacherRange
    .clearContent()
    .clearNote()
    .setBackground(
      ENGAGELI_CONFIG.colors.teacher
    )
    .setFontWeight('bold');

  rosterSheet
    .getRange(
      teacherRow,
      studentIdColumn
    )
    .setValue(
      ENGAGELI_CONFIG.teacher
        .rosterMarker
    )
    .setNote(
      'This row was created by the Engageli Attendance script and is used as the session-time benchmark.'
    );

  const lastNameColumn =
    engFindOptionalHeaderColumn_(
      rosterSheet,
      ENGAGELI_CONFIG.headers
        .rosterLastName
    );

  const firstNameColumn =
    engFindOptionalHeaderColumn_(
      rosterSheet,
      ENGAGELI_CONFIG.headers
        .rosterFirstName
    );

  const nameColumn =
    engFindOptionalHeaderColumn_(
      rosterSheet,
      ENGAGELI_CONFIG.headers
        .rosterName
    );

  const nameParts =
    engSplitTeacherName_(
      teacherName
    );

  if (lastNameColumn !== null) {
    rosterSheet
      .getRange(
        teacherRow,
        lastNameColumn
      )
      .setValue(
        nameParts.lastName
      );
  }

  if (firstNameColumn !== null) {
    rosterSheet
      .getRange(
        teacherRow,
        firstNameColumn
      )
      .setValue(
        nameParts.firstName
      );
  }

  if (nameColumn !== null) {
    rosterSheet
      .getRange(
        teacherRow,
        nameColumn
      )
      .setValue(teacherName);
  }

  for (
    let s = 0;
    s < sessions.length;
    s++
  ) {
    const column =
      sessionColumns.get(
        sessions[s].label
      );

    const duration =
      teacherDurations[s];

    const cell =
      rosterSheet.getRange(
        teacherRow,
        column
      );

    cell.setNumberFormat(
      '[h]:mm'
    );

    if (duration !== null) {
      cell
        .setValue(duration)
        .setBackground(
          ENGAGELI_CONFIG
            .colors.teacher
        )
        .setNote(
          'Teacher benchmark: ' +
            engFormatDuration_(
              duration
            )
        );
    }
  }
}


/**
 * Deletes only rows carrying
 * the exact TEACHER marker.
 */
function engRemoveTeacherRows_(
  rosterSheet,
  studentIdColumn
) {
  const lastRow =
    rosterSheet.getLastRow();

  if (lastRow < 2) {
    return 0;
  }

  const values = rosterSheet
    .getRange(
      2,
      studentIdColumn,
      lastRow - 1,
      1
    )
    .getDisplayValues();

  const rowsToDelete = [];

  for (
    let i = 0;
    i < values.length;
    i++
  ) {
    if (
      String(values[i][0])
        .trim()
        .toUpperCase() ===
      ENGAGELI_CONFIG.teacher
        .rosterMarker
    ) {
      rowsToDelete.push(i + 2);
    }
  }

  for (
    let i =
      rowsToDelete.length - 1;
    i >= 0;
    i--
  ) {
    rosterSheet.deleteRow(
      rowsToDelete[i]
    );
  }

  return rowsToDelete.length;
}


/**
 * Selects a color based on the
 * student's percentage of teacher time.
 */
function engGetAttendanceColor_(
  studentDuration,
  teacherDuration
) {
  if (
    teacherDuration === null ||
    !isFinite(teacherDuration) ||
    teacherDuration <= 0
  ) {
    return ENGAGELI_CONFIG.colors
      .unavailable;
  }

  const ratio = Math.max(
    0,
    studentDuration /
      teacherDuration
  );

  if (ratio >= 0.9) {
    return ENGAGELI_CONFIG.colors.full;
  }

  if (ratio >= 0.75) {
    return ENGAGELI_CONFIG.colors.most;
  }

  if (ratio >= 0.5) {
    return ENGAGELI_CONFIG.colors.half;
  }

  if (ratio >= 0.25) {
    return ENGAGELI_CONFIG.colors.brief;
  }

  return ENGAGELI_CONFIG.colors
    .veryBrief;
}


/**
 * Creates the hover note for
 * each attendance cell.
 */
function engBuildAttendanceNote_(
  studentDuration,
  teacherDuration
) {
  if (
    teacherDuration === null ||
    !isFinite(teacherDuration) ||
    teacherDuration <= 0
  ) {
    return (
      'Student time: ' +
      engFormatDuration_(
        studentDuration
      ) +
      '\nTeacher session time is unavailable, so no percentage could be calculated.'
    );
  }

  const ratio = Math.max(
    0,
    studentDuration /
      teacherDuration
  );

  const displayedPercent =
    Math.round(ratio * 100);

  return (
    displayedPercent +
    '% of teacher session time' +
    '\nStudent: ' +
    engFormatDuration_(
      studentDuration
    ) +
    '\nTeacher: ' +
    engFormatDuration_(
      teacherDuration
    )
  );
}


/**
 * Builds the completion report.
 */
function engBuildSummary_(result) {
  const lines = [
    'Sessions imported: ' +
      result.sessionsImported,

    'Students matched: ' +
      result.studentsMatched +
      ' of ' +
      result.uniqueStudentsInEngageli,

    'Teacher records found: ' +
      result.teacherRowsFound
  ];

  if (
    result.importTeacherRowRequested
  ) {
    lines.push(
      result.teacherRowAdded
        ? 'Teacher benchmark row: added at the bottom of the Roster'
        : 'Teacher benchmark row: not added because no teacher record was found'
    );
  } else {
    lines.push(
      'Teacher benchmark row: turned off in Settings'
    );
  }

  if (
    result.duplicateStudentRowsCombined >
    0
  ) {
    lines.push(
      'Duplicate Engageli student rows combined: ' +
        result
          .duplicateStudentRowsCombined
    );
  }

  if (
    result.invalidStudentEmails > 0
  ) {
    lines.push(
      'Student rows with unusable email addresses: ' +
        result.invalidStudentEmails
    );
  }

  if (
    result.skippedNonAttendanceRoles >
    0
  ) {
    lines.push(
      'Rows with other roles skipped: ' +
        result
          .skippedNonAttendanceRoles
    );
  }

  if (
    result.unmatchedStudentIds.length >
    0
  ) {
    const preview =
      result.unmatchedStudentIds.slice(
        0,
        12
      );

    let unmatchedLine =
      'Student numbers not found on the Roster: ' +
      preview.join(', ');

    if (
      result.unmatchedStudentIds
        .length > preview.length
    ) {
      unmatchedLine +=
        ' and ' +
        (
          result.unmatchedStudentIds
            .length - preview.length
        ) +
        ' more';
    }

    lines.push(unmatchedLine);
  }

  return lines.join('\n');
}


/**
 * Extracts the leading student number
 * from an Engageli email.
 */
function engExtractStudentIdFromEmail_(
  emailValue
) {
  const email = String(
    emailValue || ''
  )
    .trim()
    .toLowerCase();

  const atIndex =
    email.indexOf('@');

  if (atIndex <= 0) {
    return null;
  }

  const localPart =
    email.slice(0, atIndex);

  const match =
    localPart.match(/^(\d+)/);

  return match ? match[1] : null;
}


/**
 * Normalizes Student Number values
 * read from the Roster.
 */
function engNormalizeRosterId_(value) {
  if (
    typeof value === 'number' &&
    isFinite(value)
  ) {
    if (
      Math.floor(value) === value
    ) {
      return String(value);
    }

    return String(value).trim();
  }

  return String(value || '').trim();
}


/**
 * Converts a duration into the numeric
 * day fraction used by Google Sheets.
 */
function engParseDuration_(value) {
  if (
    value === '' ||
    value === null ||
    typeof value === 'undefined'
  ) {
    return null;
  }

  if (typeof value === 'number') {
    return (
      isFinite(value) &&
      value >= 0
    )
      ? value
      : null;
  }

  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {
    return (
      value.getHours() / 24 +
      value.getMinutes() / 1440 +
      value.getSeconds() / 86400
    );
  }

  const text =
    String(value).trim();

  if (!text) {
    return null;
  }

  if (
    /^\d+(\.\d+)?$/.test(text)
  ) {
    const numericValue =
      Number(text);

    return (
      isFinite(numericValue) &&
      numericValue >= 0
    )
      ? numericValue
      : null;
  }

  const timeMatch = text.match(
    /^(\d+):(\d{1,2})(?::(\d{1,2}))?$/
  );

  if (!timeMatch) {
    return null;
  }

  const hours =
    Number(timeMatch[1]);

  const minutes =
    Number(timeMatch[2]);

  const seconds =
    Number(timeMatch[3] || 0);

  if (
    minutes > 59 ||
    seconds > 59
  ) {
    return null;
  }

  return (
    hours * 3600 +
    minutes * 60 +
    seconds
  ) / 86400;
}


/**
 * Returns the greater of two durations
 * while correctly handling blank values.
 */
function engMaxDuration_(
  first,
  second
) {
  if (first === null) {
    return second;
  }

  if (second === null) {
    return first;
  }

  return Math.max(
    first,
    second
  );
}


/**
 * Formats a duration as hours:minutes
 * for cell notes and messages.
 */
function engFormatDuration_(duration) {
  if (
    duration === null ||
    !isFinite(duration) ||
    duration < 0
  ) {
    return 'Unavailable';
  }

  const totalMinutes =
    Math.round(
      duration * 24 * 60
    );

  const hours =
    Math.floor(
      totalMinutes / 60
    );

  const minutes =
    totalMinutes % 60;

  return (
    hours +
    ':' +
    String(minutes).padStart(
      2,
      '0'
    )
  );
}


/**
 * Removes the Engageli role suffix
 * from the teacher's name.
 */
function engCleanEngageliName_(value) {
  return String(value || '')
    .replace(
      /\s*\((?:TE|PROF|INSTRUCTOR)\)\s*$/i,
      ''
    )
    .trim();
}


/**
 * Splits the teacher's name into the
 * Roster first- and last-name columns.
 */
function engSplitTeacherName_(
  teacherName
) {
  const cleanedName =
    String(teacherName || '')
      .trim();

  const parts =
    cleanedName
      .split(/\s+/)
      .filter(function(part) {
        return part !== '';
      });

  if (parts.length === 0) {
    return {
      firstName: 'Teacher',
      lastName: 'Benchmark'
    };
  }

  if (parts.length === 1) {
    return {
      firstName: parts[0],
      lastName: 'Teacher'
    };
  }

  return {
    firstName:
      parts
        .slice(0, -1)
        .join(' '),

    lastName:
      parts[parts.length - 1]
  };
}


/**
 * Gets a required sheet.
 */
function engGetRequiredSheet_(
  ss,
  sheetName
) {
  const sheet =
    ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      'Required sheet not found: ' +
        sheetName
    );
  }

  return sheet;
}


/**
 * Finds a required Roster header.
 */
function engFindRequiredHeaderColumn_(
  sheet,
  headerText
) {
  const column =
    engFindOptionalHeaderColumn_(
      sheet,
      headerText
    );

  if (column === null) {
    throw new Error(
      'Could not find the header "' +
        headerText +
        '" on the ' +
        sheet.getName() +
        ' sheet.'
    );
  }

  return column;
}


/**
 * Finds an optional header.
 */
function engFindOptionalHeaderColumn_(
  sheet,
  headerText
) {
  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {
    return null;
  }

  const headers = sheet
    .getRange(
      1,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0];

  const matches = [];

  for (
    let i = 0;
    i < headers.length;
    i++
  ) {
    if (
      String(headers[i]).trim() ===
      headerText
    ) {
      matches.push(i + 1);
    }
  }

  if (matches.length > 1) {
    throw new Error(
      'More than one column is labeled "' +
        headerText +
        '" on the ' +
        sheet.getName() +
        ' sheet.'
    );
  }

  return matches.length === 1
    ? matches[0]
    : null;
}


/**
 * Finds a required header in an array.
 */
function engFindHeaderIndexInArray_(
  headers,
  headerText
) {
  const index =
    engFindOptionalHeaderIndexInArray_(
      headers,
      headerText
    );

  if (index === null) {
    throw new Error(
      'Could not find the Engageli Data header "' +
        headerText +
        '".'
    );
  }

  return index;
}


/**
 * Finds an optional header in an array.
 */
function engFindOptionalHeaderIndexInArray_(
  headers,
  headerText
) {
  const matches = [];

  const expected =
    headerText.toLowerCase();

  for (
    let i = 0;
    i < headers.length;
    i++
  ) {
    if (
      String(headers[i])
        .trim()
        .toLowerCase() ===
      expected
    ) {
      matches.push(i);
    }
  }

  if (matches.length > 1) {
    throw new Error(
      'More than one Engageli Data column is labeled "' +
        headerText +
        '".'
    );
  }

  return matches.length === 1
    ? matches[0]
    : null;
}


/**
 * Adds rows if the sheet is too short.
 */
function engEnsureSheetRows_(
  sheet,
  requiredLastRow
) {
  const missingRows =
    requiredLastRow -
    sheet.getMaxRows();

  if (missingRows > 0) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      missingRows
    );
  }
}


/**
 * Adds columns if the sheet is too narrow.
 */
function engEnsureSheetColumns_(
  sheet,
  requiredLastColumn
) {
  const missingColumns =
    requiredLastColumn -
    sheet.getMaxColumns();

  if (missingColumns > 0) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      missingColumns
    );
  }
}