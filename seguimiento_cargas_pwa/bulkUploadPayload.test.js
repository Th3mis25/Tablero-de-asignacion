const assert = require('assert');
const { prepareBulkRows } = require('./app.js');

(function testPrepareBulkRowsWithExcelArray() {
  const excelRows = [
    ['Trip', 'Ejecutivo', 'TR-MX', 'Comentarios', 'Llegada carga'],
    ['225500', 'Ana', 'TMX-001', 'Listo para envío', '2024-06-01 08:30'],
    ['', 'Carlos', '', '', '']
  ];

  const preparation = prepareBulkRows(excelRows);
  assert.ok(Array.isArray(preparation.rows), 'prepareBulkRows should return rows array');
  assert.strictEqual(preparation.rows.length, 1, 'One row should be returned for valid input');
  assert.deepStrictEqual(
    preparation.issues,
    ['Fila 3: Trip vacío.'],
    'Issues should include row number from original Excel data'
  );

  const row = preparation.rows[0];
  assert.strictEqual(row['Trip'], '225500', 'Trip should be normalized from Excel header');
  assert.strictEqual(row['Ejecutivo'], 'Ana', 'Ejecutivo should be normalized from Excel header');
  assert.strictEqual(row['TR-MX'], 'TMX-001', 'TR-MX value should be preserved in payload');
  assert.strictEqual(row['Comentarios'], 'Listo para envío', 'Comentarios value should be preserved in payload');
  assert.strictEqual(
    row['Llegada carga'],
    '2024-06-01T08:30:00',
    'Llegada carga should be converted to ISO format for the payload'
  );
})();

(function testPrepareBulkRowsWithObjectInput() {
  const preparation = prepareBulkRows([
    {
      Trip: '225501',
      Ejecutivo: 'Luis',
      'Cita entrega': '2024-06-02',
      'Tracking': 'TRK-002'
    }
  ]);

  assert.strictEqual(preparation.rows.length, 1, 'Object input should still be supported');
  assert.deepStrictEqual(preparation.issues, [], 'No issues expected for valid object input');
  assert.strictEqual(
    preparation.rows[0]['Cita entrega'],
    '2024-06-02T00:00:00',
    'Date values from object input should be normalized'
  );
})();

console.log('Bulk upload payload test passed.');
