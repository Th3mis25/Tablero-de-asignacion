const assert = require('assert');
const { formatHeaderLabel } = require('./app.js');

assert.strictEqual(formatHeaderLabel('TR-MX'), 'TR-MX');
assert.strictEqual(formatHeaderLabel('TR-USA'), 'TR-USA');
assert.strictEqual(formatHeaderLabel('cliente'), 'Cliente');
assert.strictEqual(formatHeaderLabel(' cita carga '), 'Cita carga');

console.log('formatHeaderLabel tests passed.');
