/**
 * Parcel Doot - Priority Service
 * Google Apps Script backend
 *
 * SETUP:
 * 1. Create a new Google Apps Script project.
 * 2. Paste this file into Code.gs and create index.html using the companion file.
 * 3. Run setupSheets() once from the editor and authorize access.
 * 4. In Project Settings > Script properties, set ADMIN_PASSWORD to a strong password.
 * 5. Deploy > New deployment > Web app; Execute as: Me; choose access appropriate for your users.
 */

const CONFIG_SHEET = 'FormConfig';
const RESPONSES_SHEET = 'Responses';
const SPREADSHEET_ID_PROPERTY = 'SPREADSHEET_ID';
const ADMIN_PASSWORD_PROPERTY = 'ADMIN_PASSWORD';

function doGet() {
  setupSheetsIfNeeded_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Parcel Doot - Priority Service')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Run once manually. Creates a spreadsheet with configuration and response tabs.
 * If SPREADSHEET_ID is already stored, uses that spreadsheet.
 */
function setupSheets() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const savedId = props.getProperty(SPREADSHEET_ID_PROPERTY);

  if (savedId) {
    ss = SpreadsheetApp.openById(savedId);
  } else {
    ss = SpreadsheetApp.create('Parcel Doot - Priority Service Responses');
    props.setProperty(SPREADSHEET_ID_PROPERTY, ss.getId());
  }

  let config = ss.getSheetByName(CONFIG_SHEET);
  if (!config) config = ss.insertSheet(CONFIG_SHEET);
  let responses = ss.getSheetByName(RESPONSES_SHEET);
  if (!responses) responses = ss.insertSheet(RESPONSES_SHEET);

  if (config.getLastRow() === 0) {
    const rows = [
      ['FieldID', 'Section', 'Label', 'Type', 'Required', 'Options', 'Hidden'],
      ['customer_name', 'Customer Details', 'Name of Customer / Organisation', 'text', true, '', false],
      ['mobile_number', 'Customer Details', 'Mobile Number', 'text', true, '', false],
      ['email', 'Customer Details', 'Email Address', 'email', false, '', false],
      ['booking_type', 'Booking Details', 'Type of Booking', 'select', true, 'Bulk Booking,International Booking,Domestic Parcel,Other', false],
      ['organisation', 'Booking Details', 'Company / Organisation', 'text', false, '', false],
      ['monthly_volume', 'Booking Details', 'Approximate Monthly Volume (in pieces)', 'number', false, '', false],
      ['requirement', 'Booking Details', 'Brief Requirement', 'text', false, '', false]
    ];
    config.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }

  if (responses.getLastRow() === 0) {
    responses.appendRow(['Timestamp', 'Submission ID', 'Form Data JSON']);
  }

  [config, responses].forEach(sheet => {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .setFontWeight('bold')
      .setBackground('#d32f2f')
      .setFontColor('#ffffff');
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });

  if (!props.getProperty(ADMIN_PASSWORD_PROPERTY)) {
    // Change this immediately in Script Properties after first setup.
    props.setProperty(ADMIN_PASSWORD_PROPERTY, 'CHANGE_THIS_ADMIN_PASSWORD');
  }

  return {
    spreadsheetUrl: ss.getUrl(),
    message: 'Sheets created/verified. Set ADMIN_PASSWORD in Script Properties before sharing.'
  };
}

function setupSheetsIfNeeded_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(SPREADSHEET_ID_PROPERTY);
  if (!id) {
    setupSheets();
    return;
  }
  const ss = SpreadsheetApp.openById(id);
  if (!ss.getSheetByName(CONFIG_SHEET) || !ss.getSheetByName(RESPONSES_SHEET)) {
    setupSheets();
  }
}

function getSpreadsheet_() {
  setupSheetsIfNeeded_();
  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY);
  return SpreadsheetApp.openById(id);
}

function getFormConfig() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG_SHEET);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift().map(String);
  return values.filter(row => row.some(cell => cell !== '')).map(row => {
    const item = {};
    headers.forEach((header, i) => item[header] = row[i]);
    item.Required = item.Required === true || String(item.Required).toLowerCase() === 'true';
    item.Hidden = item.Hidden === true || String(item.Hidden).toLowerCase() === 'true';
    item.FieldID = String(item.FieldID || '');
    item.Section = String(item.Section || 'General');
    item.Label = String(item.Label || item.FieldID);
    item.Type = String(item.Type || 'text').toLowerCase();
    item.Options = String(item.Options || '');
    return item;
  });
}

function submitForm(data) {
  if (!data || typeof data !== 'object') throw new Error('No form data received.');
  const config = getFormConfig();
  const output = {};
  config.forEach(field => {
    if (field.Hidden) return;
    const id = field.FieldID;
    if (!id) return;
    let value = data[id];
    if (value === undefined || value === null) value = '';
    value = String(value).trim();
    if (field.Required && !value) {
      throw new Error('Please complete the required field: ' + field.Label);
    }
    // Store only configured fields; ignore unexpected client-supplied keys.
    output[field.Label] = value;
  });

  const sheet = getSpreadsheet_().getSheetByName(RESPONSES_SHEET);
  const submissionId = Utilities.getUuid();
  sheet.appendRow([new Date(), submissionId, JSON.stringify(output)]);
  return { success: true, submissionId: submissionId };
}

function verifyAdmin(password) {
  const saved = PropertiesService.getScriptProperties().getProperty(ADMIN_PASSWORD_PROPERTY);
  if (!saved || saved === 'CHANGE_THIS_ADMIN_PASSWORD') return false;
  return constantTimeEquals_(String(password || ''), saved);
}

function constantTimeEquals_(a, b) {
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i++) result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return result === 0;
}

function updateConfig(config) {
  if (!Array.isArray(config)) throw new Error('Invalid form configuration.');
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG_SHEET);
  const headers = ['FieldID', 'Section', 'Label', 'Type', 'Required', 'Options', 'Hidden'];
  const allowedTypes = ['text', 'number', 'email', 'select', 'checkbox'];
  const rows = [headers];
  const ids = new Set();

  config.forEach(field => {
    const id = String(field.FieldID || '').trim();
    const label = String(field.Label || '').trim();
    const section = String(field.Section || 'General').trim();
    const type = String(field.Type || 'text').toLowerCase();
    if (!id || !label) throw new Error('Every field must have a Field ID and label.');
    if (!/^[a-zA-Z][a-zA-Z0-9_-]{0,59}$/.test(id)) {
      throw new Error('Invalid Field ID: ' + id + '. Use letters, numbers, underscore or hyphen; begin with a letter.');
    }
    if (ids.has(id)) throw new Error('Duplicate Field ID: ' + id);
    ids.add(id);
    if (!allowedTypes.includes(type)) throw new Error('Unsupported field type: ' + type);
    rows.push([
      id, section, label, type,
      field.Required === true || String(field.Required).toLowerCase() === 'true',
      String(field.Options || ''),
      field.Hidden === true || String(field.Hidden).toLowerCase() === 'true'
    ]);
  });

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d32f2f').setFontColor('#ffffff');
  sheet.autoResizeColumns(1, headers.length);
  return 'Form configuration saved successfully.';
}

/** Utility: open the linked response spreadsheet from the Apps Script editor. */
function getResponseSpreadsheetUrl() {
  return getSpreadsheet_().getUrl();
}
