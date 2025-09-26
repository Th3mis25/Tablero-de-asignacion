const assert = require('assert');
const { fmtDate, DEFAULT_LOCALE } = require('./app.js');

const isoSample = '2024-06-09T15:30:00';

assert.strictEqual(fmtDate(isoSample, 'en-US'), '06/09/2024 15:30');
assert.strictEqual(fmtDate(isoSample, DEFAULT_LOCALE), '09/06/2024 15:30');
assert.strictEqual(fmtDate(isoSample, 'de-DE'), '09.06.2024 15:30');
assert.strictEqual(fmtDate(isoSample), '09/06/2024 15:30');

const customSample = '09/06/2024 15:30:00';

assert.strictEqual(fmtDate(customSample, 'en-US'), '06/09/2024 15:30');
assert.strictEqual(fmtDate(customSample, DEFAULT_LOCALE), '09/06/2024 15:30');
assert.strictEqual(fmtDate(customSample, 'de-DE'), '09.06.2024 15:30');

// Ensure custom format keeps local time without offset
const diffSample = '30/08/2025 08:00:00';
assert.strictEqual(fmtDate(diffSample, 'en-US'), '08/30/2025 08:00');
assert.strictEqual(fmtDate(diffSample, DEFAULT_LOCALE), '30/08/2025 08:00');
assert.strictEqual(fmtDate(diffSample, 'de-DE'), '30.08.2025 08:00');

// ISO string with explicit Z should keep the correct time
const isoZSample = '2024-06-09T09:00:00Z';
assert.strictEqual(fmtDate(isoZSample, 'en-US'), '06/09/2024 09:00');
assert.strictEqual(fmtDate(isoZSample, DEFAULT_LOCALE), '09/06/2024 09:00');
assert.strictEqual(fmtDate(isoZSample, 'de-DE'), '09.06.2024 09:00');

// Short d/m/yyyy without leading zeros
const shortSample = '1/9/2025 12:00';
assert.strictEqual(fmtDate(shortSample, 'en-US'), '09/01/2025 12:00');
assert.strictEqual(fmtDate(shortSample, DEFAULT_LOCALE), '01/09/2025 12:00');
assert.strictEqual(fmtDate(shortSample, 'de-DE'), '01.09.2025 12:00');

// Soportar formato con AM/PM
const ampmSample = '09/06/2024 03:30 PM';
assert.strictEqual(fmtDate(ampmSample, 'en-US'), '06/09/2024 15:30');
assert.strictEqual(fmtDate(ampmSample, DEFAULT_LOCALE), '09/06/2024 15:30');
assert.strictEqual(fmtDate(ampmSample, 'de-DE'), '09.06.2024 15:30');

console.log('All fmtDate tests passed.');
