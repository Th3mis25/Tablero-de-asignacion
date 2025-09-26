const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const headers = ['Ejecutivo', 'Trip', 'Caja', 'Referencia', 'Cliente', 'Destino', 'Estatus', 'Segmento', 'TR-MX', 'TR-USA', 'Cita carga', 'Llegada carga', 'Cita entrega', 'Llegada entrega', 'Comentarios', 'Docs', 'Tracking'];

let storedValues;
let appendedValues;

const sheet = {
  getDataRange: () => ({ getValues: () => storedValues }),
  getLastRow: () => storedValues.length,
  getRange: () => ({
    setValues: (vals) => {
      appendedValues = vals.map(row => row.slice());
      for (let i = 0; i < vals.length; i++) {
        storedValues.push(vals[i].slice());
      }
    }
  })
};

const ss = {
  getSheetByName: () => sheet,
  getSpreadsheetTimeZone: () => 'UTC'
};

const sandbox = {
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: () => 'demo-token'
    })
  },
  SpreadsheetApp: { openById: () => ss },
  ContentService: {
    MimeType: { JSON: 'application/json' },
    createTextOutput: () => ({
      content: '',
      headers: {},
      setContent(value) { this.content = value; return this; },
      setMimeType() { return this; },
      setHeader(name, value) { this.headers[name] = value; return this; }
    })
  },
  Utilities: { parseDate: (val) => val }
};

vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(__dirname + '/backend/Code.gs', 'utf8'), sandbox);

function callBulkAdd(rows) {
  const event = {
    postData: {},
    parameter: {
      action: 'bulkAdd',
      token: 'demo-token',
      rows: JSON.stringify(rows)
    },
    headers: { Authorization: 'Bearer demo-token' }
  };
  appendedValues = [];
  const response = sandbox.doPost(event);
  return JSON.parse(response.content);
}

function buildRow(trip, ejecutivo) {
  const row = new Array(headers.length).fill('');
  row[0] = ejecutivo;
  row[1] = trip;
  return row;
}

storedValues = [headers.slice(), buildRow('300', 'Ana')];

const result = callBulkAdd([
  { 'Trip': '301', 'Ejecutivo': 'Luis' },
  { 'Trip': '300', 'Ejecutivo': 'Carla' },
  { 'Trip': '301', 'Ejecutivo': 'Miguel' }
]);

assert.strictEqual(result.success, true, 'bulkAdd should succeed');
assert.strictEqual(result.inserted, 1, 'Only one row should be inserted');
assert.deepStrictEqual(result.duplicates, ['300', '301'], 'Duplicates should include existing and repeated trips');
assert.strictEqual(appendedValues.length, 1, 'Only one row should be appended');
const tripIndex = headers.indexOf('Trip');
assert.strictEqual(appendedValues[0][tripIndex], '301', 'Trip 301 should be inserted');

console.log('Bulk add duplicates test passed.');
