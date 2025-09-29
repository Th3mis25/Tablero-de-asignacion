const SHEET_NAME = 'Tabla_1';
const AUTH_TOKEN = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
const SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms';

function secureCompare(a, b) {
  if (typeof a !== 'string') {
    a = a == null ? '' : String(a);
  }
  if (typeof b !== 'string') {
    b = b == null ? '' : String(b);
  }
  var len = Math.max(a.length, b.length);
  var mismatch = a.length ^ b.length;
  for (var i = 0; i < len; i++) {
    var aCode = i < a.length ? a.charCodeAt(i) : 0;
    var bCode = i < b.length ? b.charCodeAt(i) : 0;
    mismatch |= aCode ^ bCode;
  }
  return mismatch === 0;
}

function isAuthorized(e) {
  if (!AUTH_TOKEN) {
    Logger.log('API token is not configured. Please set the API_TOKEN script property.');
    return false;
  }
  var headerToken = '';
  if (e && e.headers) {
    var authHeader = e.headers.Authorization || e.headers.authorization;
    if (authHeader && authHeader.indexOf('Bearer ') === 0) {
      headerToken = authHeader.substring(7).trim();
    }
  }
  var paramToken = e.parameter && e.parameter.token;
  return secureCompare(headerToken, AUTH_TOKEN) || secureCompare(paramToken, AUTH_TOKEN);
}

function normalizeOrigin(origin) {
  if (!origin) return '';
  var str = String(origin).trim();
  if (!str) return '';
  var match = str.match(/^(https?:\/\/[^/]+)/i);
  return match ? match[1] : '';
}

function getRequestOrigin(e) {
  if (!e || !e.headers) return '';
  var headers = e.headers;
  var originHeader = headers.Origin || headers.origin || '';
  var origin = normalizeOrigin(originHeader);
  if (origin) return origin;
  var refererHeader = headers.Referer || headers.referer || '';
  return normalizeOrigin(refererHeader);
}

function createJsonOutput(payload, status, origin) {
  var allowedOrigin = normalizeOrigin(origin) || '*';
  var output = ContentService.createTextOutput()
    .setContent(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', allowedOrigin)
    .setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization')
    .setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');

  if (allowedOrigin !== '*') {
    output.setHeader('Vary', 'Origin');
  }

  if (typeof status === 'number') {
    output.setHeader('X-Http-Status-Code-Override', String(status));
  }

  return output;
}

function parseDateSafe(value, timeZone) {
  if (value == null || value === '') {
    return '';
  }

  if (typeof Utilities !== 'undefined' && Utilities.parseDate) {
    try {
      return Utilities.parseDate(value, timeZone, "yyyy-MM-dd'T'HH:mm:ss");
    } catch (err) {
      try {
        return Utilities.parseDate(value, timeZone, "yyyy-MM-dd HH:mm:ss");
      } catch (ignored) {
        // Fall through to generic parsing.
      }
    }
  }

  var parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }

  return value;
}

function formatDateSafe(value, timeZone) {
  if (!(value instanceof Date)) {
    return value;
  }

  if (typeof Utilities !== 'undefined' && Utilities.formatDate) {
    return Utilities.formatDate(value, timeZone, "yyyy-MM-dd'T'HH:mm:ss");
  }

  var iso = value.toISOString();
  return iso ? iso.substring(0, 19) : value;
}

function doPost(e) {
  var origin = getRequestOrigin(e);

  if (!isAuthorized(e)) {
    return createJsonOutput({ error: 'Unauthorized' }, 401, origin);
  }

  if (!e.postData) {
    return createJsonOutput({ error: 'Missing postData' }, 400, origin);
  }

  try {
    var p = e.parameter;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet ' + SHEET_NAME + ' not found');
    var timeZone = ss.getSpreadsheetTimeZone();
    if (p.action === 'add') {
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var headerMap = {};
      for (var i = 0; i < headers.length; i++) {
        headerMap[String(headers[i]).trim().toLowerCase()] = i;
      }
      var tripIdx = headerMap['trip'];
      if (tripIdx == null) throw new Error('Trip column not found');
      var trip = String(p.trip || '').trim();
      if (!/^\d+$/.test(trip)) throw new Error('Invalid trip');
      if (Number(trip) < 225000) throw new Error('Trip must be >= 225000');
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][tripIdx]) === trip) throw new Error('Trip already exists');
      }
      var citaCargaDate = parseDateSafe(p.citaCarga, timeZone);
      var llegadaCargaDate = parseDateSafe(p.llegadaCarga, timeZone);
      var citaEntregaDate = parseDateSafe(p.citaEntrega, timeZone);
      var llegadaEntregaDate = parseDateSafe(p.llegadaEntrega, timeZone);
      var ejecutivo = (p.ejecutivo || p.Ejecutivo || '').trim();
      if (!ejecutivo) throw new Error('Missing ejecutivo');
      var row = new Array(headers.length).fill('');
      var map = {
        'Ejecutivo': ejecutivo,
        'Trip': trip,
        'Caja': p.caja || '',
        'Referencia': p.referencia || '',
        'Cliente': p.cliente || '',
        'Destino': p.destino || '',
        'Estatus': p.estatus || '',
        'Segmento': p.segmento || '',
        'TR-MX': p.trmx || '',
        'TR-USA': p.trusa || '',
        'Cita carga': citaCargaDate,
        'Llegada carga': llegadaCargaDate,
        'Cita entrega': citaEntregaDate,
        'Llegada entrega': llegadaEntregaDate,
        'Comentarios': p.comentarios || '',
        'Docs': p.docs || '',
        'Tracking': p.tracking || ''
      };
      for (var h in map) {
        var idx = headerMap[h.toLowerCase()];
        if (idx > -1) {
          row[idx] = map[h];
        }
      }
      sheet.appendRow(row);
      return createJsonOutput({ success: true }, 200, origin);
    } else if (p.action === 'bulkAdd') {
      var rows;
      try {
        rows = JSON.parse(p.rows || '[]');
      } catch(err) {
        throw new Error('Invalid rows data');
      }
      if (!Array.isArray(rows)) throw new Error('Invalid rows data');
      var data = sheet.getDataRange().getValues();
      var headers = data[0] || [];
      var headerMap = {};
      for (var i = 0; i < headers.length; i++) {
        headerMap[String(headers[i]).trim().toLowerCase()] = i;
      }
      var tripIdx = headerMap['trip'];
      if (tripIdx == null) throw new Error('Trip column not found');
      var ejecutivoIdx = headerMap['ejecutivo'];
      if (ejecutivoIdx == null) throw new Error('Ejecutivo column not found');
      var existingTrips = {};
      for (var r = 1; r < data.length; r++) {
        var existingTripValue = data[r][tripIdx];
        if (existingTripValue != null && existingTripValue !== '') {
          var existingTripKey = String(existingTripValue).trim();
          if (existingTripKey) {
            existingTrips[existingTripKey] = true;
          }
        }
      }
      var duplicates = [];
      var duplicateMap = {};
      var values = [];
      var invalidRows = [];
      for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        var rowObj = rows[rowIndex];
        var arr = new Array(headers.length).fill('');
        for (var key in rowObj) {
          var normalizedKey = String(key).trim();
          var idx = headerMap[normalizedKey.toLowerCase()];
          if (idx > -1) {
            var val = rowObj[key];
            if (normalizedKey === 'Cita carga' || normalizedKey === 'Llegada carga' ||
                normalizedKey === 'Cita entrega' || normalizedKey === 'Llegada entrega') {
              val = parseDateSafe(val, timeZone);
            }
            arr[idx] = val;
          }
        }
        var tripValue = arr[tripIdx];
        var tripKey = tripValue == null ? '' : String(tripValue).trim();
        var ejecutivoValue = arr[ejecutivoIdx];
        var ejecutivoKey = ejecutivoValue == null ? '' : String(ejecutivoValue).trim();
        var rowIssues = [];
        if (!tripKey) {
          rowIssues.push('Trip vacío');
        } else if (!/^\d+$/.test(tripKey)) {
          rowIssues.push('Trip inválido');
        } else if (Number(tripKey) < 225000) {
          rowIssues.push('Trip menor a 225000');
        }
        if (!ejecutivoKey) {
          rowIssues.push('Ejecutivo vacío');
        }
        if (rowIssues.length > 0) {
          invalidRows.push('Fila ' + (rowIndex + 2) + ': ' + rowIssues.join(', ') + '.');
          continue;
        }
        arr[tripIdx] = tripKey;
        arr[ejecutivoIdx] = ejecutivoKey;
        if (tripKey && existingTrips[tripKey]) {
          if (!duplicateMap[tripKey]) {
            duplicates.push(tripKey);
            duplicateMap[tripKey] = true;
          }
          continue;
        }
        if (tripKey) {
          existingTrips[tripKey] = true;
        }
        values.push(arr);
      }
      if(values.length){
        sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
      }
      return createJsonOutput({
        success: true,
        inserted: values.length,
        duplicates: duplicates,
        invalidRows: invalidRows
      }, 200, origin);
    } else if (p.action === 'update') {
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      // build a case-insensitive header map to avoid issues with extra spaces
      var headerMap = {};
      for (var i = 0; i < headers.length; i++) {
        headerMap[String(headers[i]).trim().toLowerCase()] = i;
      }
      var tripIdx = headerMap['trip'];
      if (tripIdx == null) throw new Error('Trip column not found');
      var rowIndex = -1;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][tripIdx]) === String(p.originalTrip)) {
          rowIndex = i;
          break;
        }
      }
      if (rowIndex === -1) throw new Error('Trip not found');
      var trip = String(p.trip || '').trim();
      if (!/^\d+$/.test(trip)) throw new Error('Invalid trip');
      if (Number(trip) < 225000) throw new Error('Trip must be >= 225000');
      for (var i = 1; i < data.length; i++) {
        if (i !== rowIndex && String(data[i][tripIdx]) === trip) {
          throw new Error('Trip already exists');
        }
      }
      // Parse dates using the sheet timezone to keep the submitted local time without adding offsets
      var ejecutivo = (p.ejecutivo || p.Ejecutivo || '').trim();
      if (!ejecutivo) throw new Error('Missing ejecutivo');
      var citaCarga = parseDateSafe(p.citaCarga, timeZone);
      var llegadaCarga = parseDateSafe(p.llegadaCarga, timeZone);
      var citaEntrega = parseDateSafe(p.citaEntrega, timeZone);
      var llegadaEntrega = parseDateSafe(p.llegadaEntrega, timeZone);
      var map = {
        'Ejecutivo': ejecutivo,
        'Trip': trip,
        'Caja': p.caja || '',
        'Referencia': p.referencia || '',
        'Cliente': p.cliente || '',
        'Destino': p.destino || '',
        'Estatus': p.estatus || '',
        'Segmento': p.segmento || '',
        'TR-MX': p.trmx || '',
        'TR-USA': p.trusa || '',
        'Cita carga': citaCarga,
        'Llegada carga': llegadaCarga,
        'Cita entrega': citaEntrega,
        'Llegada entrega': llegadaEntrega,
        'Comentarios': p.comentarios || '',
        'Docs': p.docs || '',
        'Tracking': p.tracking || ''
      };
      for (var h in map) {
        var idx = headerMap[h.toLowerCase()];
        if (idx > -1) {
          sheet.getRange(rowIndex + 1, idx + 1).setValue(map[h]);
        }
      }
      return createJsonOutput({ success: true }, 200, origin);
    } else {
      return createJsonOutput({ error: 'Unsupported action' }, 400, origin);
    }
  } catch (err) {
    var status = (
      err.message === 'Trip not found' ||
      err.message === 'Missing ejecutivo' ||
      err.message === 'Invalid trip' ||
      err.message === 'Trip must be >= 225000' ||
      err.message === 'Trip already exists'
    ) ? 400 : 500;
    return createJsonOutput({ error: err.message }, status, origin);
  }
}

function doGet(e) {
  var origin = getRequestOrigin(e);

  if (!isAuthorized(e)) {
    return createJsonOutput({ error: 'Unauthorized' }, 401, origin);
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet ' + SHEET_NAME + ' not found');
    var timeZone = ss.getSpreadsheetTimeZone();
    var data = sheet.getDataRange().getValues();
    var formattedData = data.map(function(row) {
      return row.map(function(cell) {
        return cell instanceof Date
          ? formatDateSafe(cell, timeZone)
          : cell;
      });
    });
    return createJsonOutput({ data: formattedData }, 200, origin);
  } catch (err) {
    return createJsonOutput({ error: err.message }, 500, origin);
  }
}

function doOptions(e) {
  var origin = getRequestOrigin(e);
  var allowedOrigin = normalizeOrigin(origin) || '*';
  var output = ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader('Access-Control-Allow-Origin', allowedOrigin)
    .setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization')
    .setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS')
    .setHeader('Access-Control-Max-Age', '3600');

  if (allowedOrigin !== '*') {
    output.setHeader('Vary', 'Origin');
  }

  return output;
}
