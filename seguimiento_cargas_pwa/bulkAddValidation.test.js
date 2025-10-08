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

storedValues = [headers.slice()];

const result = callBulkAdd([
  { 'Trip': '', 'Ejecutivo': 'Luis' },
  { 'Trip': 'ABC', 'Ejecutivo': 'Ana' },
  { 'Trip': '224999', 'Ejecutivo': 'Eva' },
  { 'Trip': '225123', 'Ejecutivo': '' },
  { 'Trip': '225124', 'Ejecutivo': 'Zoe' }
]);

assert.strictEqual(result.success, true, 'bulkAdd should succeed even with invalid rows');
assert.strictEqual(result.inserted, 2, 'Two valid rows should be inserted when only Trip is required');
assert.deepStrictEqual(result.duplicates, [], 'No duplicates should be reported');
assert.deepStrictEqual(
  result.invalidRows,
  [
    'Fila 2: Trip vacío.',
    'Fila 3: Trip inválido.',
    'Fila 4: Trip menor a 225000.'
  ],
  'Invalid rows should include detailed issues'
);
assert.strictEqual(appendedValues.length, 2, 'Two rows should be appended for valid data');
const tripIndex = headers.indexOf('Trip');
const ejecutivoIndex = headers.indexOf('Ejecutivo');
assert.strictEqual(appendedValues[0][tripIndex], '225123', 'Trip 225123 should be stored for rows with optional Ejecutivo');
assert.strictEqual(appendedValues[0][ejecutivoIndex], '', 'Ejecutivo can remain blank in bulk uploads');
assert.strictEqual(appendedValues[1][tripIndex], '225124', 'Trip should be trimmed and stored as string');

console.log('Bulk add validation test passed.');
