const assert = require('assert');
const { prepareBulkRows, extractSheetRows } = require('./app.js');

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

(function testPrepareBulkRowsWithInlineStringCells() {
  class MiniTextNode {
    constructor(text) {
      this.text = text;
      this.nodeType = 3;
      this.children = [];
      this.parentNode = null;
    }

    getElementsByTagName() {
      return [];
    }

    getAttribute() {
      return null;
    }

    getAttributeNS() {
      return null;
    }

    get textContent() {
      return this.text;
    }
  }

  class MiniElement {
    constructor(tagName, attributes, nodeType) {
      this.tagName = tagName;
      this.attributes = attributes || {};
      this.children = [];
      this.parentNode = null;
      this.nodeType = nodeType != null ? nodeType : 1;
    }

    appendChild(child) {
      if (!child) {
        return;
      }
      this.children.push(child);
      child.parentNode = this;
    }

    getElementsByTagName(tagName) {
      const normalized = String(tagName || '').toLowerCase();
      const results = [];

      function traverse(node) {
        if (!node || !node.children) {
          return;
        }
        for (let i = 0; i < node.children.length; i++) {
          const child = node.children[i];
          if (!child) {
            continue;
          }
          if (child.nodeType === 1) {
            if (child.tagName && child.tagName.toLowerCase() === normalized) {
              results.push(child);
            }
            traverse(child);
          }
        }
      }

      traverse(this);
      return results;
    }

    getAttribute(name) {
      if (this.nodeType !== 1) {
        return null;
      }
      return Object.prototype.hasOwnProperty.call(this.attributes, name) ? this.attributes[name] : null;
    }

    getAttributeNS(namespace, localName) {
      if (this.nodeType !== 1) {
        return null;
      }
      if (namespace === 'http://www.w3.org/XML/1998/namespace') {
        const key = `xml:${localName}`;
        if (Object.prototype.hasOwnProperty.call(this.attributes, key)) {
          return this.attributes[key];
        }
      }
      return null;
    }

    get textContent() {
      let text = '';
      for (let i = 0; i < this.children.length; i++) {
        const child = this.children[i];
        if (!child) {
          continue;
        }
        text += child.nodeType === 3 ? child.text : child.textContent;
      }
      return text;
    }
  }

  class MiniDocument extends MiniElement {
    constructor() {
      super('#document', {}, 9);
    }
  }

  class MiniDOMParser {
    parseFromString(xmlText) {
      const doc = new MiniDocument();
      const stack = [doc];
      const tagRegex = /<[^>]+>/g;
      let lastIndex = 0;

      function appendTextContent(text) {
        if (text == null || text === '') {
          return;
        }
        const parent = stack[stack.length - 1];
        if (parent && parent.nodeType === 9 && /^\s*$/.test(text)) {
          return;
        }
        parent.appendChild(new MiniTextNode(text));
      }

      let match;
      while ((match = tagRegex.exec(xmlText))) {
        const textSegment = xmlText.slice(lastIndex, match.index);
        appendTextContent(textSegment);
        const tag = match[0];
        lastIndex = tagRegex.lastIndex;

        if (tag.startsWith('<?') || tag.startsWith('<!')) {
          continue;
        }

        if (tag.startsWith('</')) {
          if (stack.length > 1) {
            stack.pop();
          }
          continue;
        }

        const selfClosing = tag.endsWith('/>');
        const inner = tag.slice(1, tag.length - (selfClosing ? 2 : 1)).trim();
        if (!inner) {
          continue;
        }

        const spaceIndex = inner.search(/\s/);
        const tagName = spaceIndex === -1 ? inner : inner.slice(0, spaceIndex);
        const attrString = spaceIndex === -1 ? '' : inner.slice(spaceIndex + 1);
        const attributes = {};
        const attrRegex = /([^\s=]+)(?:\s*=\s*("([^"]*)"|'([^']*)'|([^\s"'>=]+)))?/g;
        let attrMatch;

        while ((attrMatch = attrRegex.exec(attrString))) {
          const name = attrMatch[1];
          if (!name) {
            continue;
          }
          const value =
            attrMatch[3] != null
              ? attrMatch[3]
              : attrMatch[4] != null
              ? attrMatch[4]
              : attrMatch[5] != null
              ? attrMatch[5]
              : '';
          attributes[name] = value;
        }

        const element = new MiniElement(tagName, attributes);
        const parent = stack[stack.length - 1];
        parent.appendChild(element);
        if (!selfClosing) {
          stack.push(element);
        }
      }

      const tailText = xmlText.slice(lastIndex);
      appendTextContent(tailText);

      return doc;
    }
  }

  const previousDomParser = global.DOMParser;
  global.DOMParser = MiniDOMParser;

  try {
    const sheetXml =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData>' +
      '<row r="1">' +
      '<c r="A1" t="inlineStr"><is><t>Trip</t></is></c>' +
      '<c r="B1" t="inlineStr"><is><t>Comentarios</t></is></c>' +
      '</row>' +
      '<row r="2">' +
      '<c r="A2"><v>225503</v></c>' +
      '<c r="B2" t="inlineStr"><is><r><t>Inline</t></r><r><t xml:space="preserve"> listo</t></r></is></c>' +
      '</row>' +
      '</sheetData>' +
      '</worksheet>';

    const rawRows = extractSheetRows(sheetXml, [], { cellXfs: [], numFmtMap: {} });
    const preparation = prepareBulkRows(rawRows);

    assert.strictEqual(preparation.rows.length, 1, 'Inline string rows should be processed');
    assert.deepStrictEqual(preparation.issues, [], 'Inline string rows should not trigger issues');
    assert.strictEqual(
      preparation.rows[0]['Comentarios'],
      'Inline listo',
      'Inline string cell text should be retained through prepareBulkRows'
    );
  } finally {
    global.DOMParser = previousDomParser;
  }
})();

console.log('Bulk upload payload test passed.');
