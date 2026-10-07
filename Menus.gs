/** All 10RosterHR menus (v2.0). Requires Config.gs.
 * onOpen is a simple trigger: it runs for every editor who opens the file, with
 * no install step and no per-person trigger. It only builds menus.
 * This must be the only onOpen function in the project.
 *
 * ONE-TIME MOVE FROM v1: each person who ever ran an install10... function
 * should run install10RosterTools once. It removes that person's old
 * menu-building open triggers (which would otherwise add duplicate menus).
 * The roster edit trigger and the Tuesday WIG schedule are not touched.
 */
function onOpen() { build10Menus_(); }

function build10Menus_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Call Tools')
    .addItem('Open student demographics', 'openStudentDemographics')
    .addSeparator()
    .addItem('Save / archive Call to rosters', 'saveSccToRoster')
    .addItem('Log SCC success / failed attempt in PowerSchool', 'logSccInPowerSchool')
    .addItem('Save Call & open PowerSchool', 'archiveSccAndOpenPowerSchool')
    .addItem('Reset Call form', 'resetCallEntry')
    .addSeparator()
    .addItem('Check roster sync', 'check10RosterSync')
    .addToUi();
  ui.createMenu('ECC Tools')
    .addItem('Archive current conversation', 'archiveCurrentECCNote')
    .addItem('Archive conversation & open PowerSchool', 'logCurrentEccRowAndOpenPowerSchool')
    .addSeparator()
    .addItem('Archive student contact attempt', 'archiveCurrentECCAttempt')
    .addItem('Archive attempt & open PowerSchool', 'logCurrentEccAttemptAndOpenPowerSchool')
    .addSeparator()
    .addItem('Refresh ECC Date colors', 'applyEccDateColors')
    .addToUi();
  ui.createMenu('Student refresh')
    .addItem('Preview source changes', 'preview10StudentRefresh')
    .addItem('Refresh existing students and contacts', 'refreshStudentData')
    .addToUi();
  ui.createMenu('Grade Reports')
    .addItem('Preview grade report photos', 'preview10GradeReportPhotos')
    .addItem('Import grade report photos', 'import10GradeReportPhotos')
    .addToUi();
  ui.createMenu('WIG Tools')
    .addItem('Preview this week\u2019s snapshot', 'previewWigSnapshot')
    .addItem('Save this week\u2019s snapshot', 'saveWigSnapshot')
    .addItem('Reformat student lists', 'recolorWigStudentRepeats')
    .addSeparator()
    .addItem('Enable Tuesday 4 PM snapshots', 'enableTuesdayWigSnapshots')
    .addItem('Disable my Tuesday snapshots', 'disableTuesdayWigSnapshots')
    .addToUi();
}

/** Handlers the old per-tool open triggers call. They do nothing now, so an old
 * trigger that has not been removed yet cannot add duplicate menus. */
const LEGACY_MENU_HANDLERS_ = ['show10CallToolsMenu', 'show10EccToolsMenu', 'show10StudentRefreshMenu', 'show10WigToolsMenu'];
function show10CallToolsMenu() {}
function show10EccToolsMenu() {}
function show10StudentRefreshMenu() {}
function show10WigToolsMenu() {}

/** Removes this account's old menu open triggers. Other people's triggers can
 * only be removed by them (Apps Script only lists your own). */
function remove10LegacyMenuTriggers() {
  const ss = HR10CFG.workbook();
  const old = ScriptApp.getProjectTriggers().filter(t =>
    LEGACY_MENU_HANDLERS_.includes(t.getHandlerFunction()) &&
    t.getEventType() === ScriptApp.EventType.ON_OPEN && t.getTriggerSourceId() === ss.getId());
  old.forEach(t => ScriptApp.deleteTrigger(t));
  return old.length;
}

/** One install for every tool: validates the rosters, Call named ranges and
 * PowerSchool settings, removes old menu triggers, then shows the menus.
 * Changes no student data. Safe to run again. */
function install10RosterTools() {
  const ss = HR10CALL.workbook();
  update10RosterFields(master => { HR10CALL.validate(master); return []; });
  HR10ECC.validate();
  if (!ss.getSheetByName('Call') || !ss.getSheetByName('Info')) throw new Error('Call and Info tabs must exist.');
  if (typeof HR10ECCPS === 'undefined') throw new Error('Add EccPowerSchool before installing.');
  ['SCC Success', 'SCC Attempt', 'ECC Conversation', 'ECC Attempt'].forEach(w => HR10ECCPS.settings(ss, w));
  HR10CALL.setupNames();
  const removed = remove10LegacyMenuTriggers();
  build10Menus_();
  ss.toast('Menus ready' + (removed ? '; removed ' + removed + ' old menu triggers' : '') +
    '. Student data was not changed.', 'Roster tools', 8);
}
// Old install names now all run the single install.
function install10CallTools() { return install10RosterTools(); }
function install10EccTools() { return install10RosterTools(); }
function install10StudentRefresh() { return install10RosterTools(); }
function install10WigTools() { return install10RosterTools(); }
