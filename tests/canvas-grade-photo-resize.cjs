// Optional integration test: requires @napi-rs/canvas. Pass PNG file paths or
// exercise a synthetic large screenshot. It executes the generated dialog's
// actual client-side resize/save script using a native Canvas implementation.
// It does not exercise Google's live HTML-service or Sheets service.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const {createCanvas, Image} = require('@napi-rs/canvas');

(async () => {
  const inputs = process.argv.slice(2).map(file => fs.readFileSync(file));
  if (!inputs.length) {
    const canvas = createCanvas(2400, 1600), ctx = canvas.getContext('2d');
    ctx.fillStyle = 'white'; ctx.fillRect(0, 0, 2400, 1600);
    ctx.fillStyle = 'black'; ctx.font = '48px sans-serif';
    for (let y = 80; y < 1600; y += 80) ctx.fillText('Grade report resizing test', 50, y);
    inputs.push(canvas.toBuffer('image/png'));
  }
  let job;
  const server = vm.createContext({console,
    Utilities: {getUuid: () => 'test-job', base64Encode: bytes => Buffer.from(bytes).toString('base64')},
    PropertiesService: {getDocumentProperties: () => ({setProperty(key, value) {job = JSON.parse(value);}})}
  });
  for (const file of ['Config.gs', 'GradeReportPhotos.gs']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', file), 'utf8'), server);
  }
  const api = vm.runInContext('HR10GRADES', server);
  const plan = {header: 'Week 7', options: {column: 26}, resizeQueue: inputs.map((input, i) => ({
    student: {id: String(i + 1), name: 'Photo ' + (i + 1)}, file: {id: 'f' + i, updated: 1},
    blob: {getContentType: () => 'image/png', getBytes: () => [...input]}
  }))};
  const html = api.resizeHtml(plan), saved = [];
  const status = {textContent: ''}, results = {appendChild() {}};
  const runner = {withSuccessHandler(handler) {this.success = handler; return this;},
    withFailureHandler(handler) {this.failure = handler; return this;},
    save10ResizedGradePhoto(payload) {
      try {
        assert.equal(payload.jobId, job.id);
        assert.ok(job.files.some(file => file.studentId === payload.studentId));
        const bytes = Buffer.from(payload.imageBase64, 'base64'), dimensions = api.pngDimensions([...bytes]);
        assert.ok(bytes.length <= 2097152);
        saved.push({...dimensions, bytes: bytes.length}); this.success(dimensions);
      } catch (error) {this.failure(error);}
    }};
  class BrowserImage extends Image {
    get naturalWidth() {return this.width;}
    get naturalHeight() {return this.height;}
  }
  const client = vm.createContext({Image: BrowserImage, atob: value => Buffer.from(value, 'base64').toString('binary'),
    google: {script: {run: runner}},
    document: {getElementById: id => id === 'resize-status' ? status : results,
      createElement: tag => tag === 'canvas' ? createCanvas(1, 1) : {textContent: ''}}
  });
  await vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], client);
  assert.equal(saved.length, inputs.length);
  assert.ok(status.textContent.includes('0 failed'));
  saved.forEach((output, i) => console.log(JSON.stringify({photo: i + 1,
    originalWidth: inputs[i].readUInt32BE(16), originalHeight: inputs[i].readUInt32BE(20),
    resizedWidth: output.width, resizedHeight: output.height,
    resizedPixels: output.width * output.height, resizedBytes: output.bytes})));
  console.log('Actual generated resize/save script processed all inputs successfully using native Canvas.');
})().catch(error => {console.error(error); process.exitCode = 1;});
