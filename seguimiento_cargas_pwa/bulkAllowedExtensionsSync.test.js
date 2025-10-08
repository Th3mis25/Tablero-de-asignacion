const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const htmlPath = path.join(__dirname, 'index.html');
const appScriptPath = path.join(__dirname, 'app.js');

const htmlContent = fs.readFileSync(htmlPath, 'utf8');
const inputMatch = htmlContent.match(/<input[\s\S]*?data-bulk-upload-input[\s\S]*?>/);

assert.ok(inputMatch, 'No se encontró el input de carga masiva.');

const acceptMatch = inputMatch[0].match(/accept="([^"]+)"/);

assert.ok(acceptMatch, 'No se encontró el atributo accept del input de carga masiva.');

const acceptedExtensionsFromHtml = acceptMatch[1]
  .split(',')
  .map((ext) => ext.trim().replace(/^[.]/, ''))
  .filter(Boolean)
  .sort();

const appScriptContent = fs.readFileSync(appScriptPath, 'utf8');
const bulkExtensionsMatch = appScriptContent.match(
  /const BULK_ALLOWED_EXTENSIONS = new Set\((\[[^\]]*\])\)/,
);

assert.ok(
  bulkExtensionsMatch,
  'No se encontró la constante BULK_ALLOWED_EXTENSIONS en app.js.',
);

const extensionsArrayLiteral = bulkExtensionsMatch[1];
const allowedExtensionsFromScript = Array.from(
  vm.runInNewContext(extensionsArrayLiteral),
  (ext) => String(ext).trim(),
)
  .filter(Boolean)
  .sort();

assert.deepStrictEqual(
  acceptedExtensionsFromHtml,
  allowedExtensionsFromScript,
  'Las extensiones permitidas del input y de la validación no coinciden.',
);

console.log('Las extensiones permitidas están sincronizadas entre la interfaz y la lógica.');
