const SHEET_NAME = 'Packages';
const LOG_SHEET_NAME = 'InstallLog';
const PACKAGE_HEADERS = ['Name', 'Url', 'Description', 'Arguments', 'Version', 'Category', 'Enabled'];
const LOG_HEADERS = ['Timestamp', 'Package', 'Status', 'Message', 'User', 'Machine', 'Arguments', 'Url', 'ClientVersion', 'Raw'];

/**
 * Entry point for HTTP GET requests.
 * Supports ?action=packages to return the current package catalog.
 */
function doGet(e) {
  return withErrorHandling(() => handleGet(e));
}

/**
 * Entry point for HTTP POST requests.
 * Supports action=log to persist install telemetry.
 */
function doPost(e) {
  return withErrorHandling(() => handlePost(e));
}

function handleGet(e) {
  const params = normalizeParameters(e);
  authorize(params.key);

  switch (params.action) {
    case 'packages': {
      const packages = getPackageCatalog();
      return jsonResponse({
        ok: true,
        packages,
        generatedAt: new Date().toISOString(),
      });
    }
    default:
      throw createHttpError(400, 'Unsupported action: ' + params.action);
  }
}

function handlePost(e) {
  const params = normalizeParameters(e);
  authorize(params.key);

  switch (params.action) {
    case 'log': {
      const payload = parseJsonBody(e);
      validateLogPayload(payload);
      const receipt = recordInstallEvent(payload);
      return jsonResponse({
        ok: true,
        receipt,
      });
    }
    default:
      throw createHttpError(400, 'Unsupported action: ' + params.action);
  }
}

function withErrorHandling(fn) {
  try {
    return fn();
  } catch (err) {
    return handleError(err);
  }
}

function handleError(err) {
  const status = err && err.statusCode ? err.statusCode : 500;
  const message = err && err.publicMessage ? err.publicMessage : (err && err.message) || 'Unexpected error';
  return jsonResponse({
    ok: false,
    error: message,
  }, status);
}

function normalizeParameters(e) {
  const params = (e && e.parameter) || {};
  const action = (params.action || params.Action || '').toString().trim().toLowerCase();
  const key = (params.key || params.Key || '').toString().trim();
  return {
    action: action || 'packages',
    key,
  };
}

function authorize(key) {
  const expected = (PropertiesService.getScriptProperties().getProperty('PACKAGE_API_KEY') || '').trim();
  if (!expected) {
    return;
  }
  if (!key || key !== expected) {
    throw createHttpError(401, 'Invalid or missing API key');
  }
}

function createHttpError(statusCode, publicMessage) {
  const error = new Error(publicMessage);
  error.statusCode = statusCode;
  error.publicMessage = publicMessage;
  return error;
}

function jsonResponse(payload, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(payload));
  output.setMimeType(ContentService.MimeType.JSON);
  if (statusCode) {
    output.setStatusCode(statusCode);
  }
  return output;
}

function parseJsonBody(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw createHttpError(400, 'Missing request body');
  }
  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    throw createHttpError(400, 'Body must be valid JSON');
  }
}

function validateLogPayload(payload) {
  if (!payload || typeof payload !== 'object') {
    throw createHttpError(400, 'Payload must be a JSON object');
  }
  if (!payload.packageName) {
    throw createHttpError(400, 'packageName is required');
  }
  if (!payload.status) {
    throw createHttpError(400, 'status is required');
  }
}

function getPackageCatalog() {
  const sheet = ensureSheet(SHEET_NAME, PACKAGE_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const range = sheet.getRange(1, 1, lastRow, PACKAGE_HEADERS.length);
  const values = range.getValues();
  const headers = values.shift();
  const index = buildHeaderIndex(headers);

  return values
    .map(row => mapPackageRow(row, index))
    .filter(pkg => pkg && pkg.name && pkg.url);
}

function mapPackageRow(row, index) {
  if (!row) {
    return null;
  }
  const enabled = getBoolean(row[index.enabled], true);
  if (!enabled) {
    return null;
  }
  return {
    name: getField(row, index.name),
    url: getField(row, index.url),
    description: getField(row, index.description),
    arguments: getField(row, index.arguments),
    version: getField(row, index.version),
    category: getField(row, index.category),
  };
}

function getField(row, idx) {
  if (typeof idx !== 'number' || idx < 0) {
    return '';
  }
  const value = row[idx];
  return value == null ? '' : value.toString();
}

function getBoolean(value, fallback) {
  if (value == null || value === '') {
    return fallback;
  }
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  const text = value.toString().trim().toLowerCase();
  if (!text) {
    return fallback;
  }
  return ['true', 'yes', 'y', '1'].indexOf(text) !== -1;
}

function buildHeaderIndex(headers) {
  const index = {
    name: -1,
    url: -1,
    description: -1,
    arguments: -1,
    version: -1,
    category: -1,
    enabled: -1,
  };

  headers.forEach(function (header, position) {
    if (!header) {
      return;
    }
    switch (header.toString().trim().toLowerCase()) {
      case 'name':
        index.name = position;
        break;
      case 'url':
        index.url = position;
        break;
      case 'description':
        index.description = position;
        break;
      case 'arguments':
        index.arguments = position;
        break;
      case 'version':
        index.version = position;
        break;
      case 'category':
        index.category = position;
        break;
      case 'enabled':
        index.enabled = position;
        break;
    }
  });

  return index;
}

function recordInstallEvent(payload) {
  const sheet = ensureSheet(LOG_SHEET_NAME, LOG_HEADERS);
  const timestamp = new Date();
  const row = [
    timestamp,
    payload.packageName || '',
    payload.status || '',
    payload.message || '',
    payload.user || '',
    payload.machine || '',
    payload.arguments || '',
    payload.url || '',
    payload.clientVersion || '',
    JSON.stringify(payload),
  ];
  sheet.appendRow(row);
  return {
    row: sheet.getLastRow(),
    timestamp: timestamp.toISOString(),
  };
}

function ensureSheet(name, headers) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  if (headers && headers.length > 0) {
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}
