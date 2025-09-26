const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

// Stub services required at load time
global.PropertiesService = {
  getScriptProperties: () => ({
    getProperty: () => 'demo-token'
  })
};

// Load backend script into the current context
vm.runInThisContext(fs.readFileSync('./backend/Code.gs', 'utf8'));

// Stub remaining Apps Script services
const sheetData = [['Ejecutivo', 'Cliente']]; // No Trip column
const sheet = {
  getDataRange: () => ({
    getValues: () => sheetData
  })
};

global.SpreadsheetApp = {
  openById: () => ({
    getSheetByName: () => sheet,
    getSpreadsheetTimeZone: () => 'UTC'
  })
};

global.ContentService = {
  createTextOutput: () => ({
    content: '',
    headers: {},
    setContent(c) { this.content = c; return this; },
    setMimeType() { return this; },
    setHeader(name, value) { this.headers[name] = value; return this; }
  }),
  MimeType: { JSON: 'application/json', TEXT: 'text/plain' }
};

// Ensure missing postData returns 400 and authorization is checked
let authCalled = false;
const originalIsAuthorized = isAuthorized;
global.isAuthorized = function(e) {
  authCalled = true;
  return originalIsAuthorized(e);
};

const noPostDataResult = doPost({ headers: { Authorization: 'Bearer demo-token' }, parameter: {} });
const noPostDataPayload = JSON.parse(noPostDataResult.content);
assert.strictEqual(noPostDataPayload.error, 'Missing postData');
assert.ok(!('status' in noPostDataPayload), 'Status should not be present in payload');
assert.strictEqual(
  noPostDataResult.headers['X-Http-Status-Code-Override'],
  '400'
);
assert.ok(authCalled, 'isAuthorized should be called');
console.log('Missing postData test passed.');

// Execute doPost with missing Trip column
const e = {
  postData: {},
  parameter: {
    action: 'update',
    originalTrip: '123'
  },
  headers: { Authorization: 'Bearer demo-token' }
};

const result = doPost(e);
const payload = JSON.parse(result.content);
assert.strictEqual(payload.error, 'Trip column not found');
assert.strictEqual(result.headers['X-Http-Status-Code-Override'], '500');
console.log('Trip column missing test passed.');