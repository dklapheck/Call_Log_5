// Colors the ECC Date column on the ECC tab by how long ago the date was.
// Requires Config.gs. Run applyEccDateColors() from ECC Tools or the editor.
// The rules stay in the sheet and recalculate every day on their own (TODAY()).
//
//   Empty ECC Date (student row)  -> maroon, white text
//   21+ days ago                  -> red
//   17-20 days ago                -> deep gold
//   14-16 days ago                -> gold
//   10-13 days ago                -> light yellow
//   7-9 days ago                  -> pale yellow
//   0-6 days ago (or in future)   -> green
//
// Re-running removes every rule this script made before, wherever it now sits
// (including on a column ECC Date has since moved away from), then adds a fresh
// set on the current ECC Date column. Rules you made yourself are untouched.

function applyEccDateColors() {
  const ss = HR10CFG.workbook();
  const sheet = ss.getSheetByName('ECC');
  if (!sheet) throw new Error('ECC tab not found.');

  const dateCol = HR10CFG.findColumn(sheet, 'ECC Date');
  const studentCol = HR10CFG.findColumn(sheet, 'Student Number');
  const d = '$' + HR10CFG.columnLetter(dateCol) + '2';
  const s = '$' + HR10CFG.columnLetter(studentCol) + '2';

  const lastRow = sheet.getMaxRows();
  const target = sheet.getRange(2, dateCol, lastRow - 1, 1);

  // Our rules look like =AND(ISNUMBER($AD2),...) or =AND($AD2="",$B2<>"") and
  // apply to the same single column the formula checks. Sheets keeps the two in
  // step when columns are inserted or moved, so this finds the old set anywhere.
  function madeByThisScript(rule) {
    const condition = rule.getBooleanCondition();
    if (!condition || condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return false;
    const formula = String(condition.getCriteriaValues()[0] || '').replace(/\s+/g, '');
    const match = formula.match(/^=AND\((?:ISNUMBER\(\$([A-Z]+)\d+\)|\$([A-Z]+)\d+="",\$[A-Z]+\d+<>"")/i);
    if (!match) return false;
    const column = (match[1] || match[2]).toUpperCase();
    return rule.getRanges().every(function(range) {
      return range.getSheet().getName() === 'ECC' && range.getNumColumns() === 1 &&
        HR10CFG.columnLetter(range.getColumn()) === column;
    });
  }
  const all = sheet.getConditionalFormatRules();
  const kept = all.filter(function(rule) { return !madeByThisScript(rule); });

  function rule(formula, background, fontColor) {
    const builder = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(background)
      .setRanges([target]);
    if (fontColor) builder.setFontColor(fontColor);
    return builder.build();
  }

  const age = 'TODAY()-' + d;
  const isDate = 'ISNUMBER(' + d + ')';
  const rules = [
    rule('=AND(' + d + '="",' + s + '<>"")', '#7B1113', '#FFFFFF'),   // maroon: empty
    rule('=AND(' + isDate + ',' + age + '>=21)', '#E06666'),          // red: past 20 days
    rule('=AND(' + isDate + ',' + age + '>=17,' + age + '<=20)', '#F1C232'),
    rule('=AND(' + isDate + ',' + age + '>=14,' + age + '<=16)', '#FFD966'),
    rule('=AND(' + isDate + ',' + age + '>=10,' + age + '<=13)', '#FFE599'),
    rule('=AND(' + isDate + ',' + age + '>=7,' + age + '<=9)', '#FFF2CC'),
    rule('=AND(' + isDate + ',' + age + '<=6)', '#B6D7A8')            // green: recent
  ];

  sheet.setConditionalFormatRules(kept.concat(rules));
  const removed = all.length - kept.length;
  ss.toast('ECC Date colors applied' + (removed ? '; replaced ' + removed + ' earlier rules.' : '.'), 'ECC Tools', 5);
}
