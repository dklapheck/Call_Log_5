function resetCallEntry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = getRequiredSheet_(
    ss,
    SCC_CONFIG.sheets.callEntry
  );

  // Uncheck every checkbox in the SCC form.
  // Non-checkbox cells in this range are ignored.
  sheet
    .getRange(
      SCC_CONFIG.callEntry.formRange
    )
    .uncheck();

  // Clear all manually entered SCC information
  // and the selected student.
  sheet
    .getRangeList([
      SCC_CONFIG.callEntry.studentSelector,
      ...SCC_CONFIG.callEntry.resetRanges
    ])
    .clearContent();

  // Return cursor to student selection.
  sheet
    .getRange(
      SCC_CONFIG.callEntry.studentSelector
    )
    .activate();

  ss.toast(
    'Call Entry has been reset.',
    'Reset',
    3
  );
}