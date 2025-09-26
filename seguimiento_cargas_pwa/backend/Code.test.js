const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const headers = ['Ejecutivo', 'Trip', 'Caja', 'Referencia', 'Cliente', 'Destino', 'Estatus', 'Segmento', 'TR-MX', 'TR-USA', 'Cita carga', 'Llegada carga', 'Cita entrega', 'Llegada entrega', 'Comentarios', 'Docs', 'Tracking'];

const appendedRows = [];

const sheet = {
  getDataRange: () => ({ getValues: () => [headers] }),
  appendRow: (row) => appendedRows.push(row)
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
      setContent() { return this; },
      setMimeType() { return this; },
      setHeader() { return this; }
    })
  },
  Utilities: { parseDate: (val) => val }
};

vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(__dirname + '/Code.gs', 'utf8'), sandbox);

const e = {
  postData: {},
  parameter: {
    token: 'demo-token',
    action: 'add',
    ejecutivo: 'Maria',
    trip: '225001',
    caja: 'CJ1',
    referencia: 'R1',
    cliente: 'C1',
    destino: 'D1',
    estatus: 'E1',
    segmento: 'S1',
    trmx: 'MX1',
    trusa: 'US1',
    citaCarga: '2024-08-01T12:00:00',
    llegadaCarga: '2024-08-02T01:30:00',
    citaEntrega: '2024-08-03T09:15:00',
    llegadaEntrega: '2024-08-04T17:45:00',
    comentarios: 'Sin observaciones',
    docs: 'DOCS',
    tracking: 'TRK-1'
  }
};

sandbox.doPost(e);

const row = appendedRows[0];
assert.strictEqual(row[0], 'Maria');
assert.strictEqual(row[1], '225001');
assert.strictEqual(row[2], 'CJ1');
assert.strictEqual(row[3], 'R1');
assert.strictEqual(row[4], 'C1');
assert.strictEqual(row[5], 'D1');
assert.strictEqual(row[6], 'E1');
assert.strictEqual(row[7], 'S1');
assert.strictEqual(row[8], 'MX1');
assert.strictEqual(row[9], 'US1');
assert.strictEqual(row[10], '2024-08-01T12:00:00');
assert.strictEqual(row[11], '2024-08-02T01:30:00');
assert.strictEqual(row[12], '2024-08-03T09:15:00');
assert.strictEqual(row[13], '2024-08-04T17:45:00');
assert.strictEqual(row[14], 'Sin observaciones');
assert.strictEqual(row[15], 'DOCS');
assert.strictEqual(row[16], 'TRK-1');
console.log('Code.gs tests passed.');
