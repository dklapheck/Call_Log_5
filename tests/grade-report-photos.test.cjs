const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const context = vm.createContext({console, SpreadsheetApp: {flush() {}}});
for (const file of ['Config.gs', 'GradeReportPhotos.gs']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', file), 'utf8'), context, {filename: file});
}
const api = vm.runInContext('HR10GRADES', context);
const student = (id, preferred, last, first, row = 2) => ({id, row, name: preferred, values: [preferred, id, last, first, '10th Grade', '']});
const file = (name, id = 'file-1', extra = {}) => ({name, id, mime: 'image/png', size: 80000, updated: 1, driveFile: {getBlob: () => ({id})}, ...extra});

test('Canvas IDs are not roster IDs; preferred, legal and multipart names match', () => {
  const students = [student('1001', 'Sam Perez', 'Perez', 'Sequoia'),
    student('1002', 'Mario Rodriguez Gomez', 'Rodriguez Gomez', 'Mario', 3),
    student('1003', 'RAYden King', 'King', 'Reidyn', 4)];
  const result = api.matchFiles(students, students, [file('perez_sam555_question_123_789_photo.png'),
    file('gomez_mario_rodriguez777_question_123_888_photo.png', 'f2'),
    file('king_reidyn999_question_123_999_photo.png', 'f3')]);
  assert.equal(result.matches.length, 3);
  assert.equal(result.issues.length, 0);
  assert.equal(result.missing.length, 0);
});

test('ambiguous aliases and duplicate submissions never choose an arbitrary photo', () => {
  const a = student('1', 'Alex Gray', 'Gray', 'Alex');
  const b = student('2', 'Alex Gray', 'Gray', 'Alex', 3);
  let result = api.matchFiles([a, b], [a, b], [file('gray_alex99_question_1_2_photo.png')]);
  assert.equal(result.matches.length, 0);
  assert.match(result.issues[0], /ambiguous/);
  result = api.matchFiles([a], [a], [file('gray_alex99_question_1_2_photo.png'), file('gray_alex99_question_1_3_photo.png', 'f2')]);
  assert.equal(result.matches.length, 0);
  assert.match(result.issues[0], /multiple photos/);
});

test('rejects unmatched/unsupported files but allows oversized originals for thumbnail fallback', () => {
  const a = student('1', 'Alex Gray', 'Gray', 'Alex');
  const result = api.matchFiles([a], [a], [file('unknown_person99_question_1_2_photo.png'),
    file('gray_alex99_question_1_2_photo.heic', 'f2', {mime: 'image/heic'}),
    file('gray_alex99_question_1_2_photo.png', 'f3', {size: 2097153})]);
  assert.equal(result.matches.length, 1);
  assert.equal(result.issues.length, 2);
});

test('destination and roster validation reject unsafe inputs', () => {
  assert.equal(api.columnNumber('z'), 26);
  assert.equal(api.columnNumber('AA'), 27);
  for (const value of ['B', 'Z2', 'AAAA', '=Z']) assert.throws(() => api.columnNumber(value));
  assert.equal(api.folderId('https://drive.google.com/drive/folders/example-folder-id?usp=drive_link'), 'example-folder-id');
  assert.throws(() => api.folderId('https://example.com/folder'));
  assert.throws(() => api.records([['Alex', '1'], ['Other', '1']], 'Grades'), /duplicate/);
  assert.throws(() => api.records([['Alex', '']], 'Grades'), /no Student Number/);
});

function sheetMock() {
  const sheet = {images: [], rowHeight: 21, seeded: null, failInsert: false,
    getImages() { return this.images.filter(image => !image.removed); },
    getRange(row, column) { return {getColumn: () => column, row, setValues: values => {sheet.seeded = values;}, getNumberFormats: () => [['@', '@', '@', '@', '@', '@']]}; },
    getMaxRows: () => 1000, setColumnWidth() {}, getRowHeight() { return this.rowHeight; },
    setRowHeight(row, height) { this.rowHeight = height; },
    insertImage(blob, column, row) {
      if (this.failInsert) throw new Error('image service failed');
      if (blob.tooLarge) throw new Error('The blob was too large. The maximum blob size is 2MB. The maximum number of pixels is 1 million.');
      const image = {anchor: this.getRange(row, column), title: '', description: '', removed: false,
        getAltTextTitle() { return this.title; }, getAltTextDescription() { return this.description; },
        getAnchorCell() { return this.anchor; }, setAnchorCell(cell) { this.anchor = cell; return this; },
        setAltTextTitle(value) { this.title = value; return this; },
        setAltTextDescription(value) { this.description = value; return this; },
        getInherentWidth: () => 800, getInherentHeight: () => 400,
        setWidth(value) { this.width = value; return this; }, setHeight(value) { this.height = value; return this; },
        setAnchorCellXOffset() { return this; }, setAnchorCellYOffset() { return this; },
        remove() { this.removed = true; }};
      this.images.push(image); return image;
    }};
  return sheet;
}
function plan(sheet, imageFile = file('gray_alex99_question_1_2_photo.png')) {
  const a = student('1', 'Alex Gray', 'Gray', 'Alex');
  return {sheet, seed: true, students: [a], matches: [{student: a, file: imageFile}], issues: [], missing: [], options: {column: 26}};
}

test('private blob import seeds roster, preserves aspect ratio and is idempotent', () => {
  const sheet = sheetMock(), p = plan(sheet);
  assert.equal(api.applyPlan(p).inserted, 1);
  assert.equal(sheet.seeded[0][1], '1');
  assert.equal(sheet.images[0].anchor.getColumn(), 26);
  assert.equal(sheet.images[0].width / sheet.images[0].height, 2);
  p.seed = false;
  assert.equal(api.applyPlan(p).unchanged, 1);
  assert.equal(sheet.getImages().length, 1);
  p.matches[0].file = file('gray_alex99_question_1_2_photo.png', 'file-1', {updated: 2});
  assert.equal(api.applyPlan(p).inserted, 1);
  assert.equal(sheet.getImages().length, 1);
});

test('failed replacement retains prior photo; other weeks and manual images survive', () => {
  const sheet = sheetMock(), p = plan(sheet);
  api.applyPlan(p); p.seed = false;
  const old = sheet.getImages()[0];
  const otherWeek = sheet.insertImage({}, 25, 2).setAltTextTitle(old.title);
  const manual = sheet.insertImage({}, 26, 3).setAltTextTitle('Manual photo');
  p.matches[0].file = file('gray_alex99_question_1_3_photo.png', 'new-file');
  sheet.failInsert = true;
  const result = api.applyPlan(p);
  assert.equal(result.inserted, 0);
  assert.equal(result.issues.length, 1);
  assert.equal(result.missing.length, 1);
  assert.equal(old.removed, false);
  sheet.failInsert = false;
  assert.equal(api.applyPlan(p).inserted, 1);
  assert.equal(old.removed, true);
  assert.equal(otherWeek.removed, false);
  assert.equal(manual.removed, false);
});

test('a Drive download failure causes no roster or image writes', () => {
  const sheet = sheetMock(), p = plan(sheet, file('gray_alex99_question_1_2_photo.png', 'f1',
    {driveFile: {getBlob() { throw new Error('no access'); }}}));
  assert.throws(() => api.applyPlan(p), /no access/);
  assert.equal(sheet.seeded, null);
  assert.equal(sheet.getImages().length, 0);
});

test('pixel-limit failures retry with private thumbnail and keep the result on reruns', () => {
  const sheet = sheetMock();
  let thumbnailCalls = 0;
  const p = plan(sheet, file('gray_alex99_question_1_2_photo.png', 'f1', {driveFile: {
    getBlob: () => ({tooLarge: true}),
    getThumbnail: () => { thumbnailCalls++; return {getBytes: () => [1, 2, 3]}; }
  }}));
  let result = api.applyPlan(p);
  assert.equal(result.inserted, 1);
  assert.equal(result.thumbnails.length, 1);
  assert.equal(result.issues.length, 0);
  assert.equal(thumbnailCalls, 1);
  assert.match(sheet.getImages()[0].description, /preview: thumbnail/);
  p.seed = false;
  result = api.applyPlan(p);
  assert.equal(result.unchanged, 1);
  assert.equal(thumbnailCalls, 1);
  assert.equal(sheet.getImages().length, 1);
});

test('originals above 2 MB use thumbnails without downloading the original', () => {
  const sheet = sheetMock(), p = plan(sheet, file('gray_alex99_question_1_2_photo.png', 'f1', {
    size: 2097153, driveFile: {
      getBlob() { throw new Error('Should not download large original'); },
      getThumbnail: () => ({getBytes: () => [1, 2, 3]})
    }
  }));
  assert.equal(api.applyPlan(p).thumbnails.length, 1);
});

test('missing thumbnails preserve the previous image and report the failure', () => {
  const sheet = sheetMock(), p = plan(sheet);
  api.applyPlan(p); p.seed = false;
  const previous = sheet.getImages()[0];
  p.matches[0].file = file('gray_alex99_question_1_3_photo.png', 'new-file', {driveFile: {
    getBlob: () => ({tooLarge: true}), getThumbnail: () => null
  }});
  const result = api.applyPlan(p);
  assert.equal(result.inserted, 0);
  assert.equal(previous.removed, false);
  assert.match(result.issues[0], /no thumbnail/);
});

test('unrelated insertion errors do not attempt thumbnail fallback', () => {
  const sheet = sheetMock(); sheet.failInsert = true;
  const p = plan(sheet, file('gray_alex99_question_1_2_photo.png', 'f1', {driveFile: {
    getBlob: () => ({}), getThumbnail() { throw new Error('Should not request a thumbnail'); }
  }}));
  const result = api.applyPlan(p);
  assert.match(result.issues[0], /image service failed/);
});

test('unusable thumbnails report errors and never replace a prior image', () => {
  for (const thumbnail of [{getBytes: () => ({length: 2097153})}, {tooLarge: true, getBytes: () => [1]}]) {
    const sheet = sheetMock(), p = plan(sheet);
    api.applyPlan(p); p.seed = false;
    const previous = sheet.getImages()[0];
    p.matches[0].file = file('gray_alex99_question_1_3_photo.png', 'new-file', {driveFile: {
      getBlob: () => ({tooLarge: true}), getThumbnail: () => thumbnail
    }});
    const result = api.applyPlan(p);
    assert.equal(result.inserted, 0);
    assert.equal(result.issues.length, 1);
    assert.equal(previous.removed, false);
    assert.equal(result.pendingResize, 1);
  }
});

function pngHeader(width, height) {
  const bytes = Buffer.alloc(24);
  Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]).copy(bytes);
  bytes.write('IHDR', 12);
  bytes.writeUInt32BE(width, 16); bytes.writeUInt32BE(height, 20);
  return [...bytes];
}

test('resized PNG validation checks true dimensions rather than display size', () => {
  assert.equal(api.pngDimensions(pngHeader(900, 900)).width, 900);
  assert.equal(api.pngDimensions(pngHeader(450, 900)).height, 900);
  for (const [width, height] of [[901, 800], [500, 2000], [0, 800]]) {
    assert.throws(() => api.pngDimensions(pngHeader(width, height)), /pixel dimensions/);
  }
  assert.throws(() => api.pngDimensions([1, 2, 3]), /PNG/);
  // Apps Script's getBytes returns signed bytes; the parser accepts them.
  assert.equal(api.pngDimensions(pngHeader(800, 600).map(x => x > 127 ? x - 256 : x)).width, 800);
});

test('resized previews are reused without duplicating images on the next import', () => {
  const sheet = sheetMock(), p = plan(sheet);
  api.applyPlan(p); p.seed = false;
  sheet.getImages()[0].description += '; preview: resized';
  const result = api.applyPlan(p);
  assert.equal(result.unchanged, 1);
  assert.equal(result.resized.length, 1);
  assert.equal(sheet.getImages().length, 1);
});

function resizeServerMock() {
  const sheet = sheetMock();
  sheet.getLastRow = () => 4;
  const getRange = sheet.getRange.bind(sheet);
  sheet.getRange = (row, column) => row === 1 ? {getDisplayValue: () => 'Week 7'} :
    row === 2 && column === 2 ? {getValues: () => [[''], [''], ['1']]} : getRange(row, column);
  const state = {job: JSON.stringify({id: 'job', expires: Date.now() + 60000, column: 26, header: 'Week 7',
    files: [{studentId: '1', fileId: 'f1', updated: 1000}]}), updated: 1000};
  context.PropertiesService = {getDocumentProperties: () => ({getProperty: () => state.job,
    setProperty(key, value) {state.job = value;}, deleteProperty() {state.job = null;}})};
  context.Utilities = {base64Decode: value => [...Buffer.from(value, 'base64')], newBlob: bytes => ({bytes})};
  context.SpreadsheetApp.getActiveSpreadsheet = () => ({getId: () => '1flssTRbZ74t1FFnoKlamhiSURcM-Byq5wraQHuI5TDw', getSheetByName: () => sheet});
  context.LockService = {getDocumentLock: () => ({tryLock: () => true, releaseLock() {}})};
  context.DriveApp = {getFileById: () => ({getLastUpdated: () => new Date(state.updated)})};
  context.console = {log() {}, error() {}};
  const payload = {jobId: 'job', studentId: '1', imageBase64: Buffer.from(pngHeader(900, 600)).toString('base64')};
  return {sheet, state, payload};
}

test('resize callback finds moved student rows and completes its authorized session', () => {
  const {sheet, state, payload} = resizeServerMock();
  const dimensions = api.saveResized(payload);
  assert.equal(dimensions.width, 900);
  assert.equal(sheet.getImages()[0].anchor.row, 4);
  assert.match(sheet.getImages()[0].description, /preview: resized/);
  assert.equal(state.job, null);
  assert.throws(() => api.saveResized(payload), /expired/);
  assert.equal(sheet.getImages().length, 1);
});

test('resize callback rejects tampered students, oversized dimensions and stale source revisions', () => {
  const {sheet, state, payload} = resizeServerMock();
  assert.throws(() => api.saveResized({...payload, studentId: 'other'}), /pending resize/);
  assert.throws(() => api.saveResized({...payload, imageBase64: Buffer.from(pngHeader(1512, 956)).toString('base64')}), /pixel dimensions/);
  state.updated++;
  assert.throws(() => api.saveResized(payload), /original photo changed/);
  assert.equal(sheet.getImages().length, 0);
});
