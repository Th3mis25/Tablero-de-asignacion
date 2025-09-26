#!/usr/bin/env node
/**
 * Fails if any JS file uses toLocaleDateString/toLocaleTimeString directly.
 * The fmtDate helper in app.js is whitelisted.
 */
const fs = require('fs');
const path = require('path');

const FORBIDDEN = ['toLocaleDateString', 'toLocaleTimeString'];

function walk(dir, acc = []) {
  for (const dirent of fs.readdirSync(dir, { withFileTypes: true })) {
    if (dirent.name === 'node_modules' || dirent.name === '.git') continue;
    const res = path.join(dir, dirent.name);
    if (dirent.isDirectory()) {
      walk(res, acc);
    } else if (dirent.isFile() && res.endsWith('.js') && res !== path.join(__dirname, 'check-locale.js')) {
      acc.push(res);
    }
  }
  return acc;
}

function findFmtDateRanges(content) {
  const lines = content.split(/\r?\n/);
  const ranges = [];
  for (let i = 0; i < lines.length; i++) {
    if (/function\s+fmtDate\s*\(/.test(lines[i])) {
      let depth = 0;
      for (let j = i; j < lines.length; j++) {
        const line = lines[j];
        depth += (line.match(/{/g) || []).length;
        depth -= (line.match(/}/g) || []).length;
        if (j > i && depth === 0) {
          ranges.push([i, j]);
          break;
        }
      }
    }
  }
  return ranges;
}

function lineInRanges(line, ranges) {
  return ranges.some(([start, end]) => line >= start && line <= end);
}

let violations = [];

for (const file of walk(path.resolve(__dirname, '..'))) {
  const content = fs.readFileSync(file, 'utf8');
  const lines = content.split(/\r?\n/);
  const ranges = file.endsWith('app.js') ? findFmtDateRanges(content) : [];

  lines.forEach((line, idx) => {
    for (const m of FORBIDDEN) {
      if (line.includes(m) && !lineInRanges(idx, ranges)) {
        violations.push(`${path.relative(process.cwd(), file)}:${idx + 1} -> ${m}`);
      }
    }
  });
}

if (violations.length) {
  console.error('Direct use of locale formatting methods detected:');
  violations.forEach(v => console.error('  ' + v));
  process.exit(1);
} else {
  console.log('No forbidden locale formatting methods found.');
}
