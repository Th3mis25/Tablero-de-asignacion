const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

global.PropertiesService = {
  getScriptProperties: () => ({
    getProperty: () => 'demo-token'
  })
};

global.ContentService = {
  createTextOutput: () => ({
    content: '',
    setContent(c){ this.content = c; return this; },
    setMimeType(){ return this; },
    setHeader(){ return this; }
  }),
  MimeType: { JSON: 'application/json', TEXT: 'text/plain' }
};

let sheetData;
const sheet = {
  getDataRange: () => ({ getValues: () => sheetData }),
  appendRow: row => { sheetData.push(row); },
  getRange: (r,c) => ({ setValue: val => { sheetData[r-1][c-1] = val; } })
};

global.SpreadsheetApp = {
  openById: () => ({
    getSheetByName: () => sheet,
    getSpreadsheetTimeZone: () => 'UTC'
  })
};

vm.runInThisContext(fs.readFileSync('./backend/Code.gs','utf8'));

function callAdd(trip){
  const e = {
    postData: {},
    parameter: {
      action: 'add',
      trip,
      ejecutivo: 'E1',
      estatus: 'Live',
      cliente: 'C1',
      citaCarga: '2024-01-01T00:00:00'
    },
    headers: { Authorization: 'Bearer demo-token' }
  };
  const res = doPost(e);
  return JSON.parse(res.content);
}

sheetData = [['Trip','Ejecutivo'], ['225100','X']];
let payload = callAdd('225100');
assert.strictEqual(payload.error, 'Trip already exists');

sheetData = [['Trip','Ejecutivo']];
payload = callAdd('22A0');
assert.strictEqual(payload.error, 'Invalid trip');

sheetData = [['Trip','Ejecutivo']];
payload = callAdd('200000');
assert.strictEqual(payload.error, 'Trip must be >= 225000');

console.log('Trip validation tests passed.');
