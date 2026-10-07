/** Grade report photos v1.1 (thumbnail fallback). Requires Config.gs; menus in Menus.gs.
 * @NotOnlyCurrentDoc
 * Reads private Drive blobs; embeds over-grid images anchored in the chosen
 * weekly column. No public URLs or sharing changes are needed.
 */
const HR10GRADES = (() => {
  const FIELDS = ['Preferred Name', 'Student Number', 'LAST NAME', 'FIRST NAME', 'GRADE', 'START DATE'];
  const PREFIX = 'HR10 grade photo|';
  const MIME_TYPES = ['image/png', 'image/jpeg', 'image/gif'];
  const MAX_BYTES = 2 * 1024 * 1024;
  const WIDTH = 420, HEIGHT = 240, PADDING = 6;

  // Exact token sets allow first/last-name order changes and multipart names.
  // Preferred and legal names are separate aliases; no fuzzy name matching.
  function nameKey(value) {
    return HR10CFG.text(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().split(/\s+/).sort().join(' ');
  }
  function photoName(filename) {
    const match = filename.match(/^(.+?)(\d+)_question_\d+_\d+_/i);
    if (match) return nameKey(match[1]); // Canvas ID here is NOT Student Number.
    return nameKey(filename.replace(/\.[^.]+$/, ''));
  }
  function columnNumber(letter) {
    const value = HR10CFG.text(letter).toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(value)) throw new Error('Enter a column letter, for example Z.');
    const column = [...value].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
    if (column < 7 || column > 18278) throw new Error('Choose a weekly photo column after A:F.');
    return column;
  }
  function folderId(value) {
    const input = HR10CFG.text(value);
    const match = input.match(/^https:\/\/drive\.google\.com\/drive\/(?:u\/\d+\/)?folders\/([\w-]+)(?:[/?#].*)?$/);
    if (match) return match[1];
    if (/^[\w-]{10,}$/.test(input)) return input;
    throw new Error('Paste the grade report Drive folder URL or folder ID.');
  }
  function records(values, label) {
    const seen = new Set();
    return values.map((row, i) => {
      if (row.every(HR10CFG.blank)) return null;
      const id = HR10CFG.studentId(row[1]);
      if (!id) throw new Error(label + ' row ' + (i + 2) + ' has no Student Number.');
      if (seen.has(id)) throw new Error(label + ': duplicate Student Number ' + id + '.');
      seen.add(id);
      return {id, row: i + 2, values: row, name: HR10CFG.text(row[0]) || row[3] + ' ' + row[2]};
    }).filter(Boolean);
  }
  function matchFiles(students, master, files) {
    const aliases = new Map(), masterById = new Map(master.map(s => [s.id, s]));
    students.forEach(s => {
      const names = new Set();
      [s, masterById.get(s.id)].filter(Boolean).forEach(item => {
        names.add(nameKey(item.values[0]));
        names.add(nameKey(item.values[3] + ' ' + item.values[2]));
      });
      names.forEach(key => {
        if (!key) return;
        if (!aliases.has(key)) aliases.set(key, new Set());
        aliases.get(key).add(s);
      });
    });
    const candidates = new Map(), issues = [];
    files.forEach(file => {
      if (!MIME_TYPES.includes(file.mime)) { issues.push(file.name + ': unsupported image type; skipped.'); return; }
      const matches = [...(aliases.get(photoName(file.name)) || [])];
      if (matches.length !== 1) {
        issues.push(file.name + ': ' + (matches.length ? 'ambiguous name' : 'no matching student') + '; skipped.'); return;
      }
      const student = matches[0];
      if (!candidates.has(student.id)) candidates.set(student.id, []);
      candidates.get(student.id).push({student, file});
    });
    const matches = [];
    candidates.forEach(items => {
      if (items.length === 1) matches.push(items[0]);
      else issues.push(items[0].student.name + ': multiple photos found; all skipped. Keep one photo per student in the folder.');
    });
    return {matches, issues, missing: students.filter(s => !matches.some(m => m.student.id === s.id))};
  }
  function readPlan(ss, options) {
    const sheet = ss.getSheetByName('Grades');
    if (!sheet) throw new Error('Grades tab is missing.');
    const headers = sheet.getRange(1, 1, 1, 6).getDisplayValues()[0].map(HR10CFG.normalize);
    if (FIELDS.some((h, i) => headers[i] !== h)) throw new Error('Grades A:F headers must be: ' + FIELDS.join(', ') + '.');
    if (options.column > sheet.getMaxColumns()) throw new Error('Add the destination column to Grades first.');
    const header = sheet.getRange(1, options.column).getDisplayValue();
    if (!HR10CFG.text(header)) throw new Error('Label the weekly destination column before importing.');
    const source = HR10CFG.table(ss.getSheetByName('Master'), FIELDS);
    const master = records(source.rows.map(row => FIELDS.map(h => row[source.columns.get(h)])), 'Master');
    const height = sheet.getLastRow();
    const data = height > 1 ? sheet.getRange(2, 1, height - 1, sheet.getLastColumn()).getValues() : [];
    const formulas = height > 1 ? sheet.getRange(2, 1, height - 1, sheet.getLastColumn()).getFormulas() : [];
    const seed = !data.some(row => row.some(v => !HR10CFG.blank(v))) && !formulas.some(row => row.some(Boolean));
    if (seed && sheet.getImages().length) throw new Error('Grades has images but no student rows. Restore the student rows first.');
    const students = seed ? master.map((s, i) => ({...s, row: i + 2})) : records(data.map(row => row.slice(0, 6)), 'Grades');
    if (!students.length) throw new Error('No students found to import.');
    const folder = DriveApp.getFolderById(options.folder);
    const iterator = folder.getFiles(), files = [];
    while (iterator.hasNext()) {
      const file = iterator.next();
      if (!file.isTrashed()) files.push({id: file.getId(), name: file.getName(), mime: file.getMimeType(), size: file.getSize(),
        updated: file.getLastUpdated().getTime(), driveFile: file});
    }
    const plan = matchFiles(students, master, files);
    return {...plan, sheet, seed, students, header, options};
  }
  function ownedImages(sheet, student, column) {
    const title = PREFIX + student.id;
    return sheet.getImages().filter(image => image.getAltTextTitle() === title && image.getAnchorCell().getColumn() === column);
  }
  function thumbnailBlob(file) {
    const blob = file.driveFile.getThumbnail();
    if (!blob) throw new Error('Drive has no thumbnail for this photo. Resize a copy and retry.');
    if (blob.getBytes().length > MAX_BYTES) throw new Error('Drive thumbnail also exceeds 2 MB. Resize a copy and retry.');
    return blob;
  }
  function insertPhoto(sheet, file, blob, column, row, thumbnail) {
    try {
      return {image: sheet.insertImage(blob, column, row, PADDING, PADDING), thumbnail};
    } catch (error) {
      // Retry only insertion size/pixel-limit failures, not permission or
      // unrelated service errors. The old image remains until the new one fits.
      if (thumbnail || !/blob.*too large|maximum.*(?:blob size|number of pixels)|image.*too large/i.test(error.message)) throw error;
      const preview = thumbnailBlob(file);
      try { return {image: sheet.insertImage(preview, column, row, PADDING, PADDING), thumbnail: true}; }
      catch (previewError) { throw new Error('Original exceeds image limits; thumbnail insertion failed: ' + previewError.message); }
    }
  }
  function applyPlan(plan) {
    const {sheet, options} = plan;
    const result = {inserted: 0, unchanged: 0, thumbnails: [], issues: plan.issues.slice(), missing: plan.missing.map(s => s.name)};
    if (!plan.matches.length) return result;
    // Fetch every blob before changing the sheet. Authorization/download errors
    // cannot leave a newly seeded roster without any photos.
    plan.matches.forEach(m => {
      m.thumbnail = m.file.size > MAX_BYTES;
      m.blob = m.thumbnail ? thumbnailBlob(m.file) : m.file.driveFile.getBlob();
    });
    if (plan.seed) {
      if (plan.students.length + 1 > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), plan.students.length + 1 - sheet.getMaxRows());
      const range = sheet.getRange(2, 1, plan.students.length, 6);
      range.setValues(HR10CFG.prepareValues(range, plan.students.map(s => s.values)));
    }
    sheet.setColumnWidth(options.column, WIDTH);
    plan.matches.forEach(({student, file, blob, thumbnail}) => {
      const old = ownedImages(sheet, student, options.column);
      const description = 'Drive file: ' + file.id + '; updated: ' + file.updated;
      const previewDescription = description + '; preview: thumbnail';
      const current = old.length === 1 && [description, previewDescription].includes(old[0].getAltTextDescription()) ? old[0] : null;
      let image;
      try {
        const inserted = current ? {image: current, thumbnail: current.getAltTextDescription() === previewDescription} :
          insertPhoto(sheet, file, blob, options.column, student.row, thumbnail);
        image = inserted.image;
        image.setAltTextTitle(PREFIX + student.id).setAltTextDescription(inserted.thumbnail ? previewDescription : description);
        const scale = Math.min((WIDTH - 2 * PADDING) / image.getInherentWidth(), (HEIGHT - 2 * PADDING) / image.getInherentHeight(), 1);
        image.setWidth(Math.max(1, Math.round(image.getInherentWidth() * scale)))
          .setHeight(Math.max(1, Math.round(image.getInherentHeight() * scale)))
          .setAnchorCell(sheet.getRange(student.row, options.column)).setAnchorCellXOffset(PADDING).setAnchorCellYOffset(PADDING);
        sheet.setRowHeight(student.row, Math.max(HEIGHT, sheet.getRowHeight(student.row)));
        if (!current) old.forEach(previous => previous.remove());
        if (current) result.unchanged++; else result.inserted++;
        if (inserted.thumbnail) result.thumbnails.push(student.name);
      } catch (error) {
        if (image && !current) image.remove(); // retain the previous photo on failure
        result.issues.push(student.name + ': image not imported: ' + error.message);
        result.missing.push(student.name);
      }
    });
    SpreadsheetApp.flush();
    return result;
  }
  function promptOptions(ss) {
    const ui = SpreadsheetApp.getUi(), properties = PropertiesService.getDocumentProperties();
    const saved = properties.getProperty('HR10_GRADE_PHOTO_FOLDER') || '';
    const folder = ui.prompt('Grade report photos', 'Paste this week\u2019s Drive folder URL or ID.' +
      (saved ? '\nLeave blank to reuse the last folder.' : ''), ui.ButtonSet.OK_CANCEL);
    if (folder.getSelectedButton() !== ui.Button.OK) return null;
    const target = ui.prompt('Weekly photo column', 'Enter the destination column letter (this week: Z). Leave blank for Z.', ui.ButtonSet.OK_CANCEL);
    if (target.getSelectedButton() !== ui.Button.OK) return null;
    return {folder: folderId(folder.getResponseText() || saved), column: columnNumber(target.getResponseText() || 'Z')};
  }
  function show(plan, result, preview) {
    const esc = HR10CFG.escapeHtml;
    const summary = preview ? plan.matches.length + ' matched photos; ' + (plan.seed ? plan.students.length + ' student rows will be copied from Master.' : 'existing student rows retained.') :
      result.inserted + ' photos imported; ' + result.unchanged + ' already present. ' + result.thumbnails.length + ' thumbnail previews.';
    const issues = preview ? plan.issues : result.issues;
    const html = '<div style="font:14px Arial"><p>' + esc(summary) + '</p><p>Destination: Grades column ' +
      esc(HR10CFG.columnLetter(plan.options.column)) + ' (' + esc(plan.header) + ')</p><table border="1" cellpadding="5" style="border-collapse:collapse">' +
      '<tr><th>Student</th><th>Cell</th><th>Original photo</th></tr>' + plan.matches.map(m => '<tr><td>' + esc(m.student.name) +
        '</td><td>' + esc(HR10CFG.columnLetter(plan.options.column) + m.student.row) + '</td><td><a target="_blank" href="' +
        esc(m.file.driveFile.getUrl()) + '">Open report</a></td></tr>').join('') + '</table><p>Students without an imported photo: ' +
      esc((preview ? plan.missing.map(s => s.name) : result.missing).join(', ') || 'None') + '</p>' +
      (preview || !result.thumbnails.length ? '' : '<p>Thumbnail previews (use Open report for full-resolution text): ' +
        esc(result.thumbnails.join(', ')) + '</p>') + '<ul>' + issues.map(issue => '<li>' + esc(issue) + '</li>').join('') + '</ul></div>';
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(850).setHeight(600), preview ? 'Grade photo preview — no changes' : 'Grade photos imported');
  }
  function run(preview) {
    const ss = HR10CFG.workbook(), options = promptOptions(ss);
    if (!options) return;
    let plan, result;
    HR10CFG.withLock(() => {
      plan = readPlan(ss, options);
      if (!preview) {
        result = applyPlan(plan);
        PropertiesService.getDocumentProperties().setProperty('HR10_GRADE_PHOTO_FOLDER', options.folder);
        HR10CFG.log(result.issues.length ? 'WARN' : 'SUCCESS', 'GradeReportPhotos', '',
          result.inserted + ' imported; ' + result.unchanged + ' already present.', result);
      }
    });
    show(plan, result, preview);
    return result || {matched: plan.matches.length, issues: plan.issues};
  }
  return Object.freeze({run, nameKey, photoName, columnNumber, folderId, records, matchFiles, applyPlan});
})();

function preview10GradeReportPhotos() { return HR10GRADES.run(true); }
function import10GradeReportPhotos() { return HR10GRADES.run(false); }
