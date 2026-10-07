/** Shared SCC, ECC and demographics PowerSchool handoff. Requires Config.gs.
 * Settings come from the PowerSchoolSettings table on Info, found by its
 * "Workflow" header wherever it sits, so adding columns or moving the table
 * needs no code change. "Opening Log Text" (optional, per workflow) is put at
 * the start of the PowerSchool note only; roster history never includes it.
 * PowerSchool cannot save line breaks, so every note sent there is put on one
 * line with " | " between lines. Roster fields keep their line breaks. */
const HR10ECCPS = (() => {
// The settings table: the header row holding the single "Workflow" cell, the
// unbroken run of headers on either side of it, and every row below it until
// the Workflow column is blank.
function settingsTable_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const hits = [];
  values.forEach(function(row, r) {
    row.forEach(function(value, c) { if (HR10CFG.normalize(value) === 'Workflow') hits.push([r, c]); });
  });
  if (hits.length !== 1) {
    throw new Error('Info needs exactly one "Workflow" header for the PowerSchoolSettings table; found ' + hits.length + '.');
  }
  const r0 = hits[0][0], headerRow = values[r0].map(function(value) { return HR10CFG.normalize(value); });
  let first = hits[0][1], last = hits[0][1];
  while (first > 0 && headerRow[first - 1]) first--;
  while (last + 1 < headerRow.length && headerRow[last + 1]) last++;
  const headers = headerRow.slice(first, last + 1);
  headers.forEach(function(header, i) {
    if (headers.indexOf(header) !== i) throw new Error('PowerSchoolSettings has two "' + header + '" columns.');
  });
  const workflowColumn = headers.indexOf('Workflow'), rows = [];
  for (let r = r0 + 1; r < values.length && HR10CFG.text(values[r][first + workflowColumn]); r++) {
    rows.push(values[r].slice(first, last + 1));
  }
  return {headers: headers, rows: rows, workflowColumn: workflowColumn};
}

// Editable PowerSchool contact choices in the PowerSchoolSettings table on Info.
// Read on every handoff, so the next call uses the latest saved cells.
function readSettings(ss, workflow) {
  const sheet = ss.getSheetByName('Info');
  if (!sheet) throw new Error('Info tab is missing.');

  const table = settingsTable_(sheet);
  const headers = table.headers, rows = table.rows, workflowColumn = table.workflowColumn;
  const matchingRows = rows.filter(function(values) {
    return String(values[workflowColumn] || '').trim() === workflow;
  });
  if (matchingRows.length !== 1) throw new Error('Expected one PowerSchool settings row for ' + workflow);
  const row = matchingRows[0];
  if (!row) throw new Error('PowerSchool settings row is missing: ' + workflow);

  function valueFor_(header) {
    const index = headers.indexOf(header);
    return index === -1 ? '' : HR10CFG.text(row[index]);
  }

  const typeValue = valueFor_('Log Type value');
  const subtypeValue = valueFor_('Subtype value');
  if (Boolean(typeValue) !== Boolean(subtypeValue)) {
    throw new Error('Fill both Log Type value and Subtype value for ' + workflow +
      ' in Info, or leave both blank for manual selection.');
  }

  let extraDropdowns = [];
  const extraText = valueFor_('Additional dropdowns (JSON)');
  if (extraText) {
    try { extraDropdowns = JSON.parse(extraText); }
    catch (_) { throw new Error('Additional dropdowns for ' + workflow + ' must be valid JSON.'); }
    if (!Array.isArray(extraDropdowns) || extraDropdowns.length > 8 ||
        !extraDropdowns.every(function(entry) {
          return entry && typeof entry.name === 'string' &&
            /^[A-Za-z0-9_:-]{1,80}$/.test(entry.name) &&
            !/student|pupil|person|parent|guardian|contact|teacher|staff|school|section|course|email|phone|address|frn|date|month|day|year|calendar|time/i.test(entry.name) &&
            typeof entry.value === 'string' && entry.value.length > 0 && entry.value.length <= 120;
        })) {
      throw new Error('Additional dropdowns for ' + workflow +
        ' must be a JSON array of up to 8 {"name":"...","value":"..."} pairs.');
    }
  }

  const dateFields = [
    valueFor_('Date & Time field'),
    valueFor_('Incident Date field'),
    valueFor_('Action Date field')
  ].filter(Boolean);
  if (!dateFields.length) {
    const legacyDateField = valueFor_('Log date field');
    if (legacyDateField) dateFields.push(legacyDateField);
  }
  if (dateFields.length > 6 || dateFields.some(function(name) {
    return !/^[A-Za-z0-9_$:=.-]{1,120}$/.test(name);
  })) {
    throw new Error('A log date field for ' + workflow + ' is not a safe PowerSchool control name.');
  }

  let tagMap = {};
  const tagMapText = valueFor_('Attempt tags (JSON)');
  if (tagMapText) {
    try { tagMap = JSON.parse(tagMapText); }
    catch (_) { throw new Error('Attempt tags for ' + workflow + ' must be valid JSON.'); }
    if (!tagMap || Array.isArray(tagMap) || typeof tagMap !== 'object' ||
        !Object.keys(tagMap).every(function(key) {
          return /^[1-6]$/.test(key) && typeof tagMap[key] === 'string' &&
            tagMap[key].length > 0 && tagMap[key].length <= 120 && !/[\t\r\n]/.test(tagMap[key]);
        })) {
      throw new Error('Attempt tags for ' + workflow +
        ' must map attempt numbers 1–6 to one-line PowerSchool labels.');
    }
  }

  const tagLabel = valueFor_('Tag label');
  if (tagLabel.length > 120 || /[\t\r\n]/.test(tagLabel)) {
    throw new Error('Tag label for ' + workflow + ' must be one line and no more than 120 characters.');
  }

  const openingText = valueFor_('Opening Log Text');
  if (openingText.length > 500 || /\t/.test(openingText)) {
    throw new Error('Opening Log Text for ' + workflow + ' must be 500 characters or fewer, with no tabs.');
  }

  return {
    workflow: workflow,
    openingText: openingText,
    typeValue: typeValue,
    subtypeValue: subtypeValue,
    extraDropdowns: extraDropdowns,
    dateFields: dateFields,
    dateField: dateFields[0] || '',
    tagMap: tagMap,
    tagLabel: tagLabel
  };
}

const POWERSCHOOL_HANDOFF_HOST_ = 'https://californiak12.powerschool.com';
const POWERSCHOOL_HANDOFF_PATH_ = '/teachers/home.html';
const POWERSCHOOL_HANDOFF_TYPES_ = Object.freeze({
  scc: true,
  ecc: true,
  demographics: true
});

function encodePowerSchoolHandoffPayload_(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    throw new Error('The PowerSchool handoff payload is invalid.');
  }
  return Utilities.base64EncodeWebSafe(
    JSON.stringify(payload),
    Utilities.Charset.UTF_8
  ).replace(/=+$/g, '');
}

function buildPowerSchoolHandoffUrl_(type, encodedPayload) {
  const handoffType = String(type || '').toLowerCase();
  if (!POWERSCHOOL_HANDOFF_TYPES_[handoffType]) {
    throw new Error('Unsupported PowerSchool handoff type.');
  }
  if (!/^[A-Za-z0-9_-]+$/.test(String(encodedPayload || ''))) {
    throw new Error('The PowerSchool handoff data is invalid.');
  }
  return POWERSCHOOL_HANDOFF_HOST_ + POWERSCHOOL_HANDOFF_PATH_ +
    '#' + handoffType + '=' + encodedPayload;
}

function escapePowerSchoolDialogAttribute_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function buildPowerSchoolHandoffDialogHtml_(url) {
  const allowedUrl = new RegExp(
    '^https://californiak12\\.powerschool\\.com/teachers/home\\.html#' +
    '(scc|ecc|demographics)=[A-Za-z0-9_-]+$'
  );
  if (!allowedUrl.test(String(url))) {
    throw new Error('The PowerSchool handoff URL is invalid.');
  }
  const safeUrl = escapePowerSchoolDialogAttribute_(url);
  return '<!doctype html><html><head><base target="_blank">' +
    '<meta charset="utf-8"><style>' +
    'body{font:14px Arial,sans-serif;color:#202124;margin:20px;line-height:1.45}' +
    '.actions{display:flex;gap:10px;align-items:center;margin-top:18px}' +
    '.open{background:#174ea6;color:#fff;text-decoration:none;padding:10px 16px;' +
    'border-radius:4px;font-weight:700}.close{padding:9px 14px;cursor:pointer}' +
    '</style></head><body>' +
    '<p>Open PowerSchool to prepare this entry. Review every field, then submit it manually.</p>' +
    '<p>This tool will never click <strong>Submit</strong> for you.</p>' +
    '<div class="actions"><a id="open-powerschool" class="open" href="' + safeUrl +
    '" target="_blank" rel="noopener noreferrer" onclick="return openOnce(event)">' +
    'Open PowerSchool</a>' +
    '<button class="close" type="button" onclick="google.script.host.close()">Cancel</button></div>' +
    '<script>let opened=false;function openOnce(event){if(opened){event.preventDefault();return false;}' +
    'opened=true;const link=event.currentTarget;link.textContent="Opening PowerSchool…";' +
    'link.setAttribute("aria-disabled","true");link.style.pointerEvents="none";' +
    'setTimeout(function(){google.script.host.close();},500);return true;}</script>' +
    '</body></html>';
}

function showPowerSchoolHandoffDialog_(type, payload, title) {
  const encodedPayload = encodePowerSchoolHandoffPayload_(payload);
  const url = buildPowerSchoolHandoffUrl_(type, encodedPayload);
  const output = HtmlService.createHtmlOutput(
    buildPowerSchoolHandoffDialogHtml_(url)
  ).setWidth(430).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(output, title || 'Open PowerSchool');
  return url;
}

  /** PowerSchool note: the workflow's Opening Log Text (if any), then the note. */
  function withOpening(settings, note) {
    const opening = settings && settings.openingText ? settings.openingText : '';
    const body = HR10CFG.text(note);
    return opening && body ? opening + '\n' + body : opening || body;
  }
  /** One line for PowerSchool: lines joined with " | ", blank lines dropped. */
  function oneLine(note) {
    return String(note == null ? '' : note).split(/\r\n|\r|\n/)
      .map(function(line) { return line.replace(/\s+/g, ' ').trim(); })
      .filter(Boolean).join(' | ');
  }
  function handoff(type, payload, settings) {
    HR10CFG.workbook();
    const title = type === 'demographics' ? 'Open PowerSchool Demographics' :
      'Open PowerSchool ' + type.toUpperCase() + ' Log';
    const requestId = Utilities.getUuid();
    const data = Object.assign({}, payload, {v: 1, requestId});
    if (typeof data.note === 'string') data.note = oneLine(data.note);
    if (settings) data.settings = settings;
    try {
      const url = showPowerSchoolHandoffDialog_(type, data, title);
      HR10CFG.log('INFO', type.toUpperCase() + ' Handoff', data.studentNumber,
        'PowerSchool handoff dialog opened; submission is manual.', 'Request ID: ' + requestId);
      return url;
    } catch (error) {
      HR10CFG.log('ERROR', type.toUpperCase() + ' Handoff', data.studentNumber,
        'Could not open PowerSchool handoff.', HR10CFG.errorDetails(error));
      throw error;
    }
  }
  function open(payload, settings) {
    return handoff('ecc', {
      studentNumber: payload.studentNumber,
      date: payload.date, note: payload.note, outcome: payload.outcome
    }, settings);
  }
  return Object.freeze({settings: readSettings, open, handoff, withOpening, oneLine});
})();
