const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const appScriptPath = path.join(__dirname, 'app.js');
const appScriptContent = fs.readFileSync(appScriptPath, 'utf8');

const allowedMatch = appScriptContent.match(
  /const BULK_ALLOWED_EXTENSIONS = new Set\((\[[^\]]*\])\)/
);

assert.ok(
  allowedMatch,
  'No se encontró la constante BULK_ALLOWED_EXTENSIONS en app.js.',
);

const allowedExtensions = Array.from(
  vm.runInNewContext(allowedMatch[1]),
  (ext) => String(ext).trim(),
).filter(Boolean);

assert.ok(
  !allowedExtensions.includes('zip'),
  'La carga masiva no debe aceptar archivos .zip.',
);

const rejectedMatch = appScriptContent.match(
  /const BULK_REJECTED_EXTENSIONS = new Map\((\[[\s\S]*?\])\)/
);

assert.ok(
  rejectedMatch,
  'No se encontró la constante BULK_REJECTED_EXTENSIONS en app.js.',
);

const rejectedPairs = vm.runInNewContext(rejectedMatch[1]);
const rejectedMap = new Map(
  rejectedPairs.map((pair) => [String(pair[0]), String(pair[1])]),
);

assert.ok(
  rejectedMap.has('zip'),
  'El archivo .zip debe rechazarse explícitamente en la carga masiva.',
);

assert.ok(
  /\.xlsx/.test(rejectedMap.get('zip')),
  'El mensaje de rechazo para .zip debe indicar el uso de un archivo .xlsx.',
);

console.log('La validación de extensiones para carga masiva es correcta.');
