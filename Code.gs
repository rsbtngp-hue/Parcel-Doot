/**
 * Parcel Doot — Google Sheets API for the GitHub Pages frontend.
 * Replace the existing Apps Script project code with this file, run setup(),
 * then redeploy as a Web app (Execute as: Me; access: Anyone as permitted).
 *
 * Sheets created automatically: Form_Config, Inquiries
 * API calls are POSTed as form-urlencoded "payload" and answered with a
 * postMessage response from the returned HTML iframe. This avoids exposing
 * customer data in query strings and lets the static GitHub Pages site receive
 * an acknowledgement without requiring browser CORS response access.
 */
const SHEET_CONFIG = 'Form_Config';
const SHEET_INQUIRIES = 'Inquiries';
const DEFAULT_ADMIN_PASSWORD = 'change-me-now-2026';
const DEFAULT_FORM_FIELDS = [
  {id:'fullName',label:'Full name',type:'text',required:true,placeholder:'Your name',width:'half'},
  {id:'mobile',label:'Mobile number',type:'tel',required:true,placeholder:'Contact number',width:'half'},
  {id:'email',label:'Email address',type:'email',required:false,placeholder:'name@example.com',width:'half'},
  {id:'trackingNumber',label:'Article / tracking number',type:'text',required:false,placeholder:'If available',width:'half'},
  {id:'serviceType',label:'Type of inquiry',type:'select',required:true,options:['Parcel status','Priority service','Booking / charges','Complaint','Other'],width:'half'},
  {id:'message',label:'Inquiry details',type:'textarea',required:true,placeholder:'Describe your request',width:'full'}
];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfg = ss.getSheetByName(SHEET_CONFIG);
  if (!cfg) cfg = ss.insertSheet(SHEET_CONFIG);
  if (cfg.getLastRow() === 0) {
    cfg.appendRow(['key','value']);
    cfg.appendRow(['fields', JSON.stringify(DEFAULT_FORM_FIELDS)]);
  }
  let data = ss.getSheetByName(SHEET_INQUIRIES);
  if (!data) data = ss.insertSheet(SHEET_INQUIRIES);
  if (data.getLastRow() === 0) data.appendRow(['Submitted At','Reference ID','Full Name','Mobile','Email','Tracking Number','Inquiry Type','Message','All Form Data JSON']);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PARCEL_DOOT_ADMIN_PASSWORD')) props.setProperty('PARCEL_DOOT_ADMIN_PASSWORD', DEFAULT_ADMIN_PASSWORD);
  Logger.log('Setup complete. Change the default admin password using setAdminPassword().');
}

function doPost(e) {
  let response;
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw) throw new Error('Missing request payload.');
    const body = JSON.parse(raw);
    response = routeRequest_(body || {});
  } catch (err) {
    response = {ok:false,error:String(err && err.message || err)};
  }
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><body><script>parent.postMessage(' +
    JSON.stringify(Object.assign({parcelDoot:true}, response)).replace(/</g,'\\u003c') +
    ', "*");</script></body></html>'
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function routeRequest_(body) {
  switch (String(body.action || '')) {
    case 'getFormConfig': return {ok:true,fields:getFields_()};
    case 'submitForm': return saveInquiry_(body.data || {});
    case 'verifyAdmin': return {ok:true,authorized:verifyPassword_(body.password)};
    case 'updateConfig':
      requireAdmin_(body.password);
      saveFields_(body.fields);
      return {ok:true};
    default: throw new Error('Unsupported action.');
  }
}

function getFields_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  if (!sh || sh.getLastRow() < 2) return DEFAULT_FORM_FIELDS;
  const values = sh.getRange(1,1,sh.getLastRow(),2).getValues();
  for (let i=1;i<values.length;i++) if (values[i][0] === 'fields') {
    try { const fields=JSON.parse(values[i][1]); if (Array.isArray(fields) && fields.length) return fields; } catch (_) {}
  }
  return DEFAULT_FORM_FIELDS;
}

function saveFields_(fields) {
  if (!Array.isArray(fields) || !fields.length || fields.length > 60) throw new Error('Invalid field configuration.');
  const ids = new Set();
  fields.forEach(f => {
    if (!f || !/^[A-Za-z][A-Za-z0-9_-]{0,60}$/.test(String(f.id||''))) throw new Error('Invalid field id.');
    if (ids.has(f.id)) throw new Error('Field IDs must be unique.');
    ids.add(f.id);
    if (!String(f.label||'').trim()) throw new Error('Every field needs a label.');
    if (!['text','email','tel','number','date','select','textarea'].includes(f.type)) throw new Error('Unsupported field type: '+f.type);
    if (f.options && (!Array.isArray(f.options) || f.options.length > 100)) throw new Error('Invalid select options.');
  });
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  const values=sh.getRange(1,1,Math.max(sh.getLastRow(),1),2).getValues();
  let row=values.findIndex((r,i)=>i>0 && r[0]==='fields')+1;
  if (row < 2) row=sh.getLastRow()+1;
  sh.getRange(row,1,1,2).setValues([['fields',JSON.stringify(fields)]]);
}

function saveInquiry_(data) {
  const fields=getFields_();
  const safe={};
  fields.filter(f=>f.visible!==false).forEach(f=>{
    const value=String(data[f.id] == null ? '' : data[f.id]).trim().slice(0,5000);
    if (f.required && !value) throw new Error('Please complete: '+f.label);
    safe[f.id]=value;
  });
  const ref='PD-'+Utilities.getUuid().slice(0,8).toUpperCase();
  const byId={};fields.forEach(f=>byId[f.id]=f);
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INQUIRIES);
  if (!sh) throw new Error('Inquiries sheet is missing. Run setup() in Apps Script.');
  sh.appendRow([
    new Date(),ref,
    safe.fullName||'',safe.mobile||'',safe.email||'',safe.trackingNumber||'',
    safe.serviceType||'',safe.message||'',JSON.stringify(safe)
  ]);
  return {ok:true,reference:ref};
}

function verifyPassword_(password) {
  const expected=PropertiesService.getScriptProperties().getProperty('PARCEL_DOOT_ADMIN_PASSWORD') || DEFAULT_ADMIN_PASSWORD;
  return constantTimeEqual_(String(password||''), expected);
}
function requireAdmin_(password) {
  if (!verifyPassword_(password)) throw new Error('Administrator verification failed.');
}
function constantTimeEqual_(a,b) {
  if (a.length !== b.length) return false;
  let out=0;for(let i=0;i<a.length;i++) out|=a.charCodeAt(i)^b.charCodeAt(i);
  return out===0;
}

/** Run manually from Apps Script after setup to set a strong unique password. */
function setAdminPassword() {
  const password='REPLACE-WITH-A-STRONG-UNIQUE-PASSWORD';
  if (password === 'REPLACE-WITH-A-STRONG-UNIQUE-PASSWORD' || password.length < 14) {
    throw new Error('Edit setAdminPassword() with a strong unique password (14+ characters), then run it again.');
  }
  PropertiesService.getScriptProperties().setProperty('PARCEL_DOOT_ADMIN_PASSWORD', password);
}
