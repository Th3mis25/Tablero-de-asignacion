const assert = require('assert');
const { prepareBulkRows } = require('./app.js');

const preparation = prepareBulkRows([
  {
    Trip: '225500',
    Ejecutivo: 'Ana',
    'TR-MX': 'TMX-001',
    'Comentarios': 'Listo para envío',
    'Llegada carga': '2024-06-01 08:30'
  }
]);

assert.ok(Array.isArray(preparation.rows), 'prepareBulkRows should return rows array');
assert.strictEqual(preparation.rows.length, 1, 'One row should be returned for valid input');
assert.deepStrictEqual(preparation.issues, [], 'No issues expected for valid payload');
const row = preparation.rows[0];
assert.strictEqual(row['TR-MX'], 'TMX-001', 'TR-MX value should be preserved in payload');
assert.strictEqual(row['Comentarios'], 'Listo para envío', 'Comentarios value should be preserved in payload');
assert.strictEqual(
  row['Llegada carga'],
  '2024-06-01T08:30:00',
  'Llegada carga should be converted to ISO format for the payload'
);

console.log('Bulk upload payload test passed.');
