function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Teacher Tools')
    .addItem(
      'Archive & Open PowerSchool ECC Log',
      'archiveAndOpenPowerSchoolECC'
    )
    .addItem(
      'Archive Current ECC Note Only',
      'archiveCurrentECCNote'
    )
    .addSeparator()
    .addItem(
      'Update Roster Attendance',
      'updateEngageliAttendance'
    )
    .addItem(
      'Set Up or Repair Attendance Settings',
      'setupEngageliAttendanceSettings'
    )
    .addToUi();
}