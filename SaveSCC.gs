/** Save/Archive Call into all four rosters using the single shared save helper.
 * Version 1.2: Call cells are found through named ranges, so inserting rows or
 * columns on Call moves them automatically. Run setup10CallNamedRanges once
 * (install10RosterTools does it for you); it names the current cells only if a
 * name is missing and never moves an existing name.
 * CallStatus holds both validation messages and normal contact text.
 * Confirmation occurs outside the document lock; the confirmed note and form
 * are rechecked inside the lock before writing. Requires RosterSync 1.1.
 */
const HR10CALL = (() => {
  const TABS = ['Master', 'SCC', 'ECC', 'STAR'];
  // Named range -> address used only when setup creates a missing name.
  const NAMES = {
    studentNumber: ['CallStudentNumber', 'O2'], student: ['CallStudent', 'C2'],
    note: ['CallNote', 'D1'], toDo: ['CallToDo', 'I1'], status: ['CallStatus', 'N7'],
    form: ['CallForm', 'A7:O50'], checkboxes: ['CallCheckboxes', 'A7:I50'],
    yes: ['CallSuccessYes', null], no: ['CallSuccessNo', null]
  };
  const SUCCESS_QUESTION = 'Successful contact?';
  const RESET_PREFIX = 'CallReset_';
  const RESET_DEFAULTS = ['C2', 'H14', 'H15', 'H17:I22', 'C20', 'C22', 'C32:E32', 'C36', 'C37', 'C39', 'H42', 'C47:H47'];
  const ATTEMPTS = [1,2,3,4,5].map(n => 'Attempt ' + n + ' LG call').concat('Attempt 6+ LG call');
  const SEP = '\n\n-- Additional LG call --\n';
  const SUCCESS_FIELDS = ['SCC Notes', 'SCC To Do', 'SCC Completion'];
  const FAIL_FIELDS = ATTEMPTS.concat('SCC Completion');
  const trim = v => HR10CFG.text(v);
  const header = v => HR10CFG.normalize(v);
  function workbook() {
    const ss = HR10CFG.workbook();
    if (typeof update10RosterFields !== 'function') throw new Error('Install RosterSync 1.3 first.');
    return ss;
  }
  function named(ss, key) {
    const name = NAMES[key][0], range = ss.getRangeByName(name);
    if (!range || range.getSheet().getName() !== 'Call') {
      throw new Error('Named range ' + name + ' is missing or not on Call. Run setup10CallNamedRanges once.');
    }
    return range;
  }
  function rowHasSuccessQuestion(range) {
    const sheet = range.getSheet();
    return sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
      .some(v => header(v) === SUCCESS_QUESTION);
  }
  /** Creates only missing names, at the layout the scripts were written for. */
  function setupNames() {
    const ss = HR10CFG.workbook(), call = sheet(ss, 'Call'), created = [];
    const missing = name => { const r = ss.getRangeByName(name); return !r || r.getSheet().getName() !== 'Call'; };
    Object.keys(NAMES).forEach(key => {
      const [name, address] = NAMES[key];
      if (address && missing(name)) { ss.setNamedRange(name, call.getRange(address)); created.push(name); }
    });
    if (missing(NAMES.yes[0]) || missing(NAMES.no[0])) {
      const questions = call.getRange('B7:B50').getDisplayValues();
      const rows = questions.map((r, i) => header(r[0]) === SUCCESS_QUESTION ? i + 7 : 0).filter(Boolean);
      if (rows.length !== 1) throw new Error('Expected one ' + SUCCESS_QUESTION + ' question in Call!B7:B50 to name its Yes/No boxes.');
      if (missing(NAMES.yes[0])) { ss.setNamedRange(NAMES.yes[0], call.getRange(rows[0], 4)); created.push(NAMES.yes[0]); }
      if (missing(NAMES.no[0])) { ss.setNamedRange(NAMES.no[0], call.getRange(rows[0], 6)); created.push(NAMES.no[0]); }
    }
    const hasReset = ss.getNamedRanges().some(n => n.getName().startsWith(RESET_PREFIX));
    if (!hasReset) RESET_DEFAULTS.forEach((address, i) => {
      const name = RESET_PREFIX + String(i + 1).padStart(2, '0');
      ss.setNamedRange(name, call.getRange(address)); created.push(name);
    });
    ss.toast(created.length ? 'Created ' + created.length + ' Call named ranges.' : 'Call named ranges already exist; nothing changed.', 'Call Tools', 8);
    return created;
  }
  function sheet(ss, name) {
    const result = ss.getSheetByName(name);
    if (!result) throw new Error('Missing tab: ' + name);
    return result;
  }
  function key(value) {
    const result = trim(value);
    if (!/^\d+$/.test(result)) throw new Error('Select a student with a valid Student Number.');
    return result;
  }
  function cellText(cell) {
    const text = cell.getDisplayValue();
    if (/^#(?:REF!|N\/A|VALUE!|ERROR!|DIV\/0!|NAME\?|NUM!)/.test(text)) {
      throw new Error('Fix the Call formula error before saving: ' + text);
    }
    return trim(text);
  }
  function callData(ss) {
    const studentNumber = key(named(ss, 'studentNumber').getValue());
    if (!cellText(named(ss, 'student'))) throw new Error('Select a student on Call first.');
    const yesCell = named(ss, 'yes'), noCell = named(ss, 'no');
    if (!rowHasSuccessQuestion(yesCell) || !rowHasSuccessQuestion(noCell)) {
      throw new Error('CallSuccessYes/CallSuccessNo no longer sit on the ' + SUCCESS_QUESTION + ' row. Fix them in Data > Named ranges.');
    }
    const yes = yesCell.getValue() === true;
    const no = noCell.getValue() === true;
    if (yes === no) throw new Error('Choose either Yes or No for Successful contact?.');
    const form = named(ss, 'form').getDisplayValues().flat();
    const invalid = form.find(v => /^#(?:REF!|N\/A|VALUE!|ERROR!|DIV\/0!|NAME\?|NUM!)/.test(String(v)) || /ERROR:/.test(String(v)));
    if (invalid) throw new Error('Fix the Call response before saving: ' + invalid);
    const operational = cellText(named(ss, 'status'));
    // CallStatus also returns the normal dated contact narrative. Only these
    // explicit validation messages block a save; ordinary notes are valid.
    const blockers = ['Choose Yes or No for successful contact.',
      'Select at least one participant.', 'Select at least one contact method.'];
    if (blockers.includes(operational)) throw new Error(operational);
    const note = cellText(named(ss, 'note')), toDo = cellText(named(ss, 'toDo'));
    if (!note || /ERROR:/.test(note + '\n' + toDo)) throw new Error('Enter a valid contact note on Call.');
    return {studentNumber, note, toDo, successful: yes, workflow: yes ? 'SCC Success' : 'SCC Attempt'};
  }
  function read(master, studentNumber, field) {
    const row = master.rows.get(studentNumber), col = master.columns.get(field);
    if (!row || !col) throw new Error('Master is missing the student or field: ' + field);
    const cell = master.sheet.getRange(row,col);
    if (cell.getFormulas()[0][0]) throw new Error('Use a value in the editable field: ' + field);
    return cell.getValues()[0][0];
  }
  // Read related editable fields on every copy before planning a save. This
  // catches a pending edit rather than overwriting it with an older Master value.
  function consistent(ss, master, studentNumber, fields) {
    const central = Object.fromEntries(fields.map(f => [f, read(master,studentNumber,f)]));
    for (const tab of TABS.slice(1)) {
      const local = sheet(ss,tab), headers = HR10CFG.headerColumns(local).headers;
      const idCol = headers.indexOf('Student Number') + 1;
      const ids = local.getRange(2,idCol,local.getLastRow()-1,1).getValues();
      const row = ids.findIndex(r => trim(r[0]) === studentNumber) + 2;
      if (row < 2) throw new Error('Student is missing on ' + tab);
      const range = local.getRange(row,1,1,headers.length), values = range.getValues()[0], formulas = range.getFormulas()[0];
      for (const field of fields) {
        const i = headers.indexOf(field), v = values[i], c = central[field];
        const same = v instanceof Date && c instanceof Date ? v.getTime() === c.getTime() : v === c;
        if (i < 0 || formulas[i] || !same) throw new Error(tab + ' does not yet match Master for ' + field + '. Wait for sync and try again.');
      }
    }
    return central;
  }
  function overflowEntries(value) { return trim(value) ? String(value).split(SEP).map(trim) : []; }
  function attemptPlan(values, note) {
    const saved = ATTEMPTS.slice(0,5).map(f => trim(values[f]));
    const existing = saved.indexOf(note);
    if (existing >= 0) return {number: existing+1, header: ATTEMPTS[existing], duplicate: true};
    const overflow = overflowEntries(values[ATTEMPTS[5]]);
    if (overflow.includes(note)) return {number: 6, header: ATTEMPTS[5], duplicate: true};
    const empty = saved.indexOf('');
    if (empty >= 0) return {number: empty+1, header: ATTEMPTS[empty], value: note, duplicate: false};
    const prior = String(values[ATTEMPTS[5]] || '').trimEnd();
    return {number: 6, header: ATTEMPTS[5], value: prior ? prior + SEP + note : note, duplicate: false};
  }
  function log(level, action, studentNumber, message, details) {
    HR10CFG.log(level,action,studentNumber,message,details);
  }
  function report(ss, action, studentNumber, error, prefix) {
    const message = (prefix || '') + (error.message || String(error));
    log('ERROR',action,studentNumber,message,HR10CFG.errorDetails(error));
    ss.toast(message,action + ' needs attention',12);
    throw error;
  }
  function save() {
    const ss = workbook(); let entry, result;
    try {
      entry = callData(ss);
      let before;
      // Snapshot under the shared lock, then release it before opening any alert.
      update10RosterFields(master => {
        before = consistent(ss,master,entry.studentNumber,entry.successful ? SUCCESS_FIELDS : FAIL_FIELDS);
        return [];
      });
      const existing = trim(before['SCC Notes']);
      if (entry.successful && existing && existing !== entry.note) {
        const ui = SpreadsheetApp.getUi();
        const answer = ui.alert('Replace saved SCC note?',
          'Student ' + entry.studentNumber + ' already has a different SCC note:\n\n' + existing +
          '\n\nReplace it with the note on Call?',ui.ButtonSet.YES_NO);
        if (answer !== ui.Button.YES) {
          log('INFO','SCC Archive',entry.studentNumber,'Saved SCC note kept; replacement cancelled.','');
          ss.toast('The saved SCC note was kept.','SCC Not Saved',6);
          return {saved:false,cancelled:true};
        }
      }
      update10RosterFields(master => {
        const current = callData(ss);
        if (JSON.stringify(current) !== JSON.stringify(entry)) throw new Error('Call changed while this action was running. Review it and save again.');
        const values = consistent(ss,master,entry.studentNumber,entry.successful ? SUCCESS_FIELDS : FAIL_FIELDS);
        if (entry.successful) {
          // Recheck all reviewed fields; UI alerts suspend execution and release locks.
          if (SUCCESS_FIELDS.some(f => values[f] !== before[f])) throw new Error('The saved SCC changed while you were reviewing it. Review the new note and save again.');
          const planned = {'SCC Notes':entry.note,'SCC To Do':entry.toDo,'SCC Completion':'Completed'};
          const changes = SUCCESS_FIELDS.filter(f => planned[f] !== values[f]).map(f => ({studentNumber:entry.studentNumber,header:f,value:planned[f]}));
          result = {saved:true,duplicate:changes.length===0,studentNumber:entry.studentNumber,note:entry.note,workflow:entry.workflow,attemptNumber:null};
          return changes;
        }
        const attempt = attemptPlan(values,entry.note);
        result = {saved:true,duplicate:attempt.duplicate,studentNumber:entry.studentNumber,note:entry.note,workflow:entry.workflow,attemptNumber:attempt.number};
        if (attempt.duplicate) return [];
        const changes = [{studentNumber:entry.studentNumber,header:attempt.header,value:attempt.value}];
        // Keep an already completed SCC completed if a later contact fails.
        if (trim(values['SCC Completion']) !== 'Completed') {
          const statuses = ['First Attempt','Second Attempt','Third Attempt','Fourth Attempt','Fifth Attempt','Sixth+ Attempt'];
          changes.push({studentNumber:entry.studentNumber,header:'SCC Completion',value:statuses[attempt.number-1]});
        }
        return changes;
      });
      log('INFO','SCC Archive',entry.studentNumber,result.duplicate ? 'Entry already saved; no duplicate added.' : 'Entry saved to all four rosters.',entry.workflow);
      ss.toast(result.duplicate ? 'This SCC entry was already saved.' :
        'SCC ' + (entry.successful ? 'saved and marked Completed' : 'attempt ' + (result.attemptNumber===6 ? '6+' : result.attemptNumber) + ' saved') + ' on all four rosters.','SCC Saved',7);
      return result;
    } catch(error) { return report(ss,'SCC Archive',entry && entry.studentNumber,error); }
  }
  function selected(ss) {
    const active = SpreadsheetApp.getActiveRange();
    if (!active) throw new Error('Select a student first.');
    if (active.getSheet().getParent().getId() !== ss.getId()) throw new Error('Wrong workbook selection.');
    if (active.getSheet().getName() === 'Call') return {call:callData(ss)};
    if (!TABS.includes(active.getSheet().getName()) || active.getNumRows()!==1 || active.getNumColumns()!==1 || active.getRow()<=1) {
      throw new Error('On a roster, select one SCC Notes or Attempt LG call cell for a student.');
    }
    const local = active.getSheet(), headers = HR10CFG.headerColumns(local).headers;
    const selectedHeader = headers[active.getColumn()-1];
    if (selectedHeader !== 'SCC Notes' && !ATTEMPTS.includes(selectedHeader)) throw new Error('Select SCC Notes or an Attempt LG call cell.');
    return {studentNumber:HR10CFG.studentOnRow(local, active.getRow()), selectedHeader};
  }
  function prepareLog(ss) {
    const selection = selected(ss); let result;
    update10RosterFields(master => {
      if (selection.call) {
        const entry = callData(ss);
        if (JSON.stringify(entry)!==JSON.stringify(selection.call)) throw new Error('Call changed. Review it and try again.');
        const values = consistent(ss,master,entry.studentNumber,entry.successful ? SUCCESS_FIELDS : FAIL_FIELDS);
        result = {studentNumber:entry.studentNumber,note:entry.note,workflow:entry.workflow,attemptNumber:entry.successful ? null : attemptPlan(values,entry.note).number};
      } else {
        const values = consistent(ss,master,selection.studentNumber,[selection.selectedHeader]);
        let note = trim(values[selection.selectedHeader]);
        if (!note) throw new Error(selection.selectedHeader + ' is blank for this student.');
        if (selection.selectedHeader === ATTEMPTS[5]) {
          const entries = overflowEntries(note);
          if (entries.length>1) throw new Error('Attempt 6+ contains several entries. Use Call to log one specific attempt.');
          note = entries[0];
        }
        const success = selection.selectedHeader==='SCC Notes';
        result = {studentNumber:selection.studentNumber,note,workflow:success ? 'SCC Success':'SCC Attempt',attemptNumber:success ? null : ATTEMPTS.indexOf(selection.selectedHeader)+1};
      }
      return [];
    });
    return result;
  }
  function send(result) {
    const ss = workbook();
    if (typeof HR10ECCPS === 'undefined') throw new Error('Add EccPowerSchool.gs before using PowerSchool handoffs.');
    const settings = HR10ECCPS.settings(ss,result.workflow);
    const date = result.note.match(/\b(\d{1,2}\/\d{1,2}\/\d{4})\b/);
    // Blank date preserves the original handoff contract: review/set it manually.
    // Opening Log Text goes to PowerSchool only; the roster note is unchanged.
    return HR10ECCPS.handoff('scc',{studentNumber:result.studentNumber,note:HR10ECCPS.withOpening(settings,result.note),
      outcome:result.workflow,attemptNumber:result.attemptNumber,date:date ? date[1] : ''},settings);
  }
  function logOnly() {
    const ss = workbook(); let result;
    try {
      result = prepareLog(ss); send(result); return result;
    } catch(error) { return report(ss,'SCC Handoff',result && result.studentNumber,error); }
  }
  function saveAndLog() {
    const ss = workbook(), result = save();
    if (!result || !result.saved) return result;
    try { send(result); return result; }
    catch(error) { return report(ss,'SCC Archive & Handoff',result.studentNumber,error,'SCC entry is saved; PowerSchool did not open. '); }
  }
  function demographics() {
    const ss = workbook(); let studentNumber='';
    try {
      const active = SpreadsheetApp.getActiveRange();
      if (!active || active.getSheet().getParent().getId()!==ss.getId()) throw new Error('Select a student first.');
      if (active.getSheet().getName()==='Call') studentNumber=key(named(ss,'studentNumber').getValue());
      else if (TABS.includes(active.getSheet().getName()) && active.getNumRows()===1 && active.getRow()>1) {
        studentNumber=HR10CFG.studentOnRow(active.getSheet(), active.getRow());
      } else if (active.getNumRows()===1 && active.getNumColumns()===1) studentNumber=key(active.getValue());
      else throw new Error('Select one student row on a roster, or one cell containing a Student Number.');
      update10RosterFields(master => {
        if (!master.rows.has(studentNumber)) throw new Error('This Student Number is not in Master.');
        return [];
      });
      if (typeof HR10ECCPS==='undefined') throw new Error('Add EccPowerSchool.gs first.');
      return HR10ECCPS.handoff('demographics',{studentNumber});
    } catch(error) { return report(ss,'Demographics Handoff',studentNumber,error); }
  }
  function reset() {
    const ss=workbook();
    const targets=ss.getNamedRanges().filter(n=>n.getName().startsWith(RESET_PREFIX)).map(n=>n.getRange());
    if (!targets.length) throw new Error('No '+RESET_PREFIX+' named ranges found. Run setup10CallNamedRanges once.');
    // Check everything first so a moved name can never erase part of the form's formulas.
    targets.forEach(range=>{
      if (range.getSheet().getName()!=='Call') throw new Error('A '+RESET_PREFIX+' named range is not on Call: '+range.getA1Notation());
      if (range.getFormulas().some(row=>row.some(Boolean))) {
        throw new Error('Reset stopped: Call!'+range.getA1Notation()+' contains a formula. Fix that '+RESET_PREFIX+' named range. Nothing was cleared.');
      }
    });
    named(ss,'checkboxes').uncheck();
    targets.forEach(range=>range.clearContent());
    named(ss,'student').activate();
    ss.toast('Call form reset.','Reset',4);
  }
  function validate(master) {
    SUCCESS_FIELDS.concat(ATTEMPTS).forEach(f => {if (!master.columns.has(f)) throw new Error('Missing SCC header: '+f);});
  }
  return Object.freeze({workbook,save,logOnly,saveAndLog,demographics,reset,validate,setupNames});
})();
function saveSccToRoster() { return HR10CALL.save(); }
function archiveCurrentSCCNote() { return saveSccToRoster(); }
function logSccInPowerSchool() { return HR10CALL.logOnly(); }
// Preserve the original button's log-only behavior.
function saveSccAndOpenPowerSchool() { return logSccInPowerSchool(); }
function archiveSccAndOpenPowerSchool() { return HR10CALL.saveAndLog(); }
function resetCallEntry() { return HR10CALL.reset(); }
function setup10CallNamedRanges() { return HR10CALL.setupNames(); }
