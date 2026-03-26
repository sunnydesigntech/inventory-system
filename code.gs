/**
 * D&T QR Inventory System
 * Standalone single-file Google Apps Script web app.
 *
 * Required Script Properties:
 * - SPREADSHEET_ID
 * - WEB_APP_BASE_URL (recommended; app can still read inventory without it)
 * - INVENTORY_SHEET_NAME (optional)
 */

const CONFIG = Object.freeze({
  APP_TITLE: 'D&T QR Inventory',
  HEADER_ROW: 1,
  DEFAULT_SHEET_NAME: 'Inventory',
  SPREADSHEET_ID_PROPERTY: 'SPREADSHEET_ID',
  WEB_APP_URL_PROPERTY: 'WEB_APP_BASE_URL',
  INVENTORY_SHEET_NAME_PROPERTY: 'INVENTORY_SHEET_NAME',
  STATUS_OPTIONS: ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'],
  HAZARD_CATEGORIES: ['chemicals', 'chemical'],
  QUICKCHART_QR_BASE: 'https://quickchart.io/qr?size=220&text=',
  DEBUG_PANEL: false,
  ALIASES: {
    itemId: ['item id', 'itemid', 'id'],
    itemName: ['item name', 'name'],
    room: ['room'],
    location: ['specific location', 'location', 'specificlocation'],
    qty: ['qty', 'quantity'],
    category: ['category'],
    status: ['status'],
    qrLink: [
      'qr code link (auto-generated)',
      'qr code link',
      'auto-generated qr link',
      'qr link',
      'qr url'
    ],
    qrImage: ['qr code image', 'qr image'],
    unit: ['unit', 'uom'],
    remarks: ['remarks', 'remark', 'notes', 'note'],
    locationCode: ['location code', 'locationcode', 'location id'],
    storageId: ['storage id', 'storageid', 'storage_id'],
    storageLabel: ['storage label', 'storagelabel', 'storage_label']
  },
  REQUIRED_COLUMNS: ['itemId', 'itemName', 'room', 'location', 'qty', 'category', 'status'],
  OPTIONAL_COLUMNS: ['qrLink', 'qrImage', 'unit', 'remarks', 'locationCode', 'storageId', 'storageLabel']
});

function doGet(e) {
  const params = getRequestParams_(e);
  let bootstrap;

  try {
    bootstrap = buildBootstrapData_(params);
  } catch (err) {
    bootstrap = buildErrorBootstrap_(params, err);
  }

  return HtmlService
    .createHtmlOutput(buildPageHtml_(params, bootstrap))
    .setTitle(CONFIG.APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('D&T Inventory')
    .addItem('Refresh QR Links', 'refreshQrLinks')
    .addItem('Refresh QR Images (Optional)', 'refreshQrImages')
    .addSeparator()
    .addItem('Open Web App', 'openWebApp_')
    .addItem('Set App Config', 'promptSetAppConfig_')
    .addItem('Set WEB_APP_BASE_URL', 'promptSetWebAppBaseUrl_')
    .addSeparator()
    .addItem('Config Status / Diagnostics', 'getConfigStatus')
    .addToUi();
}

function openWebApp_() {
  const ui = SpreadsheetApp.getUi();
  const url = getWebAppBaseUrl_({ silent: true });
  if (!url) {
    ui.alert('WEB_APP_BASE_URL is not configured yet.');
    return;
  }
  ui.alert('Open this URL in your browser:\n\n' + url);
}

function promptSetAppConfig_() {
  const ui = SpreadsheetApp.getUi();

  const spreadsheetId = ui.prompt(
    'Set SPREADSHEET_ID',
    'Paste the Google Sheet ID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (spreadsheetId.getSelectedButton() !== ui.Button.OK) return;

  const webAppUrl = ui.prompt(
    'Set WEB_APP_BASE_URL',
    'Paste the deployed /exec URL. Leave blank to clear it.',
    ui.ButtonSet.OK_CANCEL
  );
  if (webAppUrl.getSelectedButton() !== ui.Button.OK) return;

  const sheetName = ui.prompt(
    'Set INVENTORY_SHEET_NAME',
    'Paste the inventory sheet tab name. Leave blank to use the default fallback.',
    ui.ButtonSet.OK_CANCEL
  );
  if (sheetName.getSelectedButton() !== ui.Button.OK) return;

  setAppConfig(
    spreadsheetId.getResponseText(),
    webAppUrl.getResponseText(),
    sheetName.getResponseText()
  );

  ui.alert('App configuration saved.');
}

function promptSetWebAppBaseUrl_() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Set WEB_APP_BASE_URL',
    'Paste your deployed /exec web app URL:',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  setWebAppBaseUrl(result.getResponseText());
  ui.alert('WEB_APP_BASE_URL saved.');
}

function getRequestParams_(e) {
  return {
    room: cleanString_(e && e.parameter && e.parameter.room),
    loc: cleanString_(e && e.parameter && e.parameter.loc),
    mode: normalizeMode_(e && e.parameter && e.parameter.mode)
  };
}

function buildBootstrapData_(params) {
  const appConfig = getAppConfig_();
  const warnings = [];
  const webAppBaseUrl = getWebAppBaseUrl_({ silent: true });
  const diagnostics = getDiagnostics_();

  if (!webAppBaseUrl) {
    warnings.push('WEB_APP_BASE_URL is not configured. QR links and location shortcuts are disabled until this is set.');
  }

  const hasLocation = !!(params.room && params.loc);
  if (hasLocation) {
    const locationResult = getInventoryData({ room: params.room, loc: params.loc });
    return {
      pageType: 'location',
      appTitle: CONFIG.APP_TITLE,
      room: locationResult.room,
      loc: locationResult.loc,
      mode: params.mode,
      rows: locationResult.rows,
      locations: [],
      message: locationResult.message || '',
      error: '',
      warnings: warnings,
      config: appConfig,
      webAppBaseUrl: webAppBaseUrl,
      diagnostics: diagnostics
    };
  }

  const locationDirectory = getAllLocations();
  return {
    pageType: 'landing',
    appTitle: CONFIG.APP_TITLE,
    room: '',
    loc: '',
    mode: 'view',
    rows: [],
    locations: locationDirectory.locations,
    message: locationDirectory.locations.length ? '' : 'No inventory locations found yet.',
    error: '',
    warnings: warnings,
    config: appConfig,
    webAppBaseUrl: webAppBaseUrl,
    diagnostics: diagnostics
  };
}

function buildErrorBootstrap_(params, err) {
  const message = err && err.message ? err.message : 'Unexpected application error.';
  return {
    pageType: 'error',
    appTitle: CONFIG.APP_TITLE,
    room: params.room || '',
    loc: params.loc || '',
    mode: params.mode || 'view',
    rows: [],
    locations: [],
    message: '',
    error: 'Configuration error: ' + message,
    warnings: [],
    config: getAppConfigSafe_(),
    webAppBaseUrl: '',
    diagnostics: getDiagnosticsSafe_()
  };
}

function getInventoryData(params) {
  const room = cleanString_(params && params.room);
  const loc = cleanString_(params && params.loc);

  if (!room || !loc) {
    return {
      success: true,
      rows: [],
      room: room,
      loc: loc,
      message: 'Please scan a valid QR code for a storage location.'
    };
  }

  const rows = getInventoryRowsForLocation_(room, loc);
  return {
    success: true,
    rows: rows,
    room: room,
    loc: loc,
    message: rows.length ? '' : 'No inventory items have been entered for this storage yet.'
  };
}

function saveInventoryUpdates(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid save payload.');
  }

  const room = cleanString_(payload.room);
  const loc = cleanString_(payload.loc);
  if (!room || !loc) {
    throw new Error('Room and location are required to save updates.');
  }

  if (!Array.isArray(payload.updates) || !payload.updates.length) {
    throw new Error('No updates provided.');
  }

  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    throw new Error('The inventory sheet is empty.');
  }

  const map = getColumnMap_(values[0], { requireQrLink: false });
  const roomNeedle = room.toLowerCase();
  const locNeedle = loc.toLowerCase();
  const lastRow = values.length;
  const seen = {};

  payload.updates.forEach(function (update) {
    const rowNum = Number(update.sheetRow);
    if (!Number.isInteger(rowNum) || rowNum <= CONFIG.HEADER_ROW || rowNum > lastRow) {
      throw new Error('Invalid row number: ' + update.sheetRow);
    }
    if (seen[rowNum]) {
      throw new Error('Duplicate row in payload: ' + rowNum);
    }
    seen[rowNum] = true;

    const row = values[rowNum - 1];
    if (!rowMatchesRoomLoc_(row, map, roomNeedle, locNeedle)) {
      throw new Error('Row ' + rowNum + ' does not belong to the selected room/location.');
    }

    if (update.qty === '' || update.qty == null) {
      throw new Error('Quantity is required for row ' + rowNum + '.');
    }

    const qty = Number(update.qty);
    if (!Number.isFinite(qty) || qty < 0) {
      throw new Error('Quantity must be a non-negative number for row ' + rowNum + '.');
    }

    const status = normalizeStatus_(update.status);
    if (CONFIG.STATUS_OPTIONS.indexOf(status) === -1) {
      throw new Error('Invalid status for row ' + rowNum + ': ' + status);
    }

    sheet.getRange(rowNum, map.qty + 1).setValue(qty);
    sheet.getRange(rowNum, map.status + 1).setValue(status);
  });

  const refreshedRows = getInventoryRowsForLocation_(room, loc);
  return {
    success: true,
    updatedCount: payload.updates.length,
    room: room,
    loc: loc,
    rows: refreshedRows,
    html: renderInventoryHtml_(refreshedRows, 'tech'),
    timestamp: new Date().toISOString()
  };
}

function getAllLocations() {
  return { success: true, locations: getAllLocations_() };
}

function getAllLocations_() {
  const dataset = getInventoryDataset_();
  const map = dataset.map;
  const values = dataset.values;
  const seen = {};
  const locations = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const room = cleanString_(row[map.room]);
    const loc = cleanString_(row[map.location]);
    if (!room || !loc) continue;

    const storageId = getOptionalValue_(row, map.storageId);
    const storageLabel = getOptionalValue_(row, map.storageLabel);
    const locationCode = getOptionalValue_(row, map.locationCode);
    const identity = buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode);
    const dedupeKey = identity.key;

    if (seen[dedupeKey]) continue;
    seen[dedupeKey] = true;

    locations.push({
      room: room,
      loc: loc,
      displayLoc: storageLabel || loc,
      locationCode: locationCode,
      storageId: storageId,
      storageLabel: storageLabel,
      routeLoc: identity.routeLoc,
      canonicalKey: identity.key,
      searchText: [room, loc, storageLabel, storageId, locationCode].filter(Boolean).join(' ').toLowerCase(),
      sortKey: [room, storageId || '', storageLabel || '', loc].join(' | ').toLowerCase()
    });
  }

  locations.sort(function (a, b) {
    return a.sortKey.localeCompare(b.sortKey);
  });

  return locations;
}

function getInventoryRowsForLocation_(room, loc) {
  const selectedRoom = cleanString_(room);
  const selectedLoc = cleanString_(loc);
  if (!selectedRoom || !selectedLoc) return [];

  const dataset = getInventoryDataset_();
  const map = dataset.map;
  const values = dataset.values;
  const roomNeedle = selectedRoom.toLowerCase();
  const locNeedle = selectedLoc.toLowerCase();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const sourceRow = values[i];
    const roomValue = cleanString_(sourceRow[map.room]);
    if (!roomValue) continue;
    if (!rowMatchesRoomLoc_(sourceRow, map, roomNeedle, locNeedle)) continue;

    rows.push(buildInventoryRowView_(
      sourceRow,
      map,
      i + 1,
      cleanString_(sourceRow[map.category]),
      roomValue,
      cleanString_(sourceRow[map.location])
    ));
  }

  rows.sort(function (a, b) {
    return a.itemName.localeCompare(b.itemName);
  });

  return rows;
}

function rowMatchesRoomLoc_(rowValues, map, roomNeedle, locNeedle) {
  const rowRoom = cleanString_(rowValues[map.room]).toLowerCase();
  if (rowRoom !== roomNeedle) return false;

  const rowLoc = cleanString_(rowValues[map.location]).toLowerCase();
  if (rowLoc === locNeedle) return true;

  const rowStorageId = getOptionalValue_(rowValues, map.storageId).toLowerCase();
  if (rowStorageId && rowStorageId === locNeedle) return true;

  const rowStorageLabel = getOptionalValue_(rowValues, map.storageLabel).toLowerCase();
  if (rowStorageLabel && rowStorageLabel === locNeedle) return true;

  const rowLocationCode = getOptionalValue_(rowValues, map.locationCode).toLowerCase();
  if (rowLocationCode && rowLocationCode === locNeedle) return true;

  return false;
}

function buildInventoryRowView_(sourceRow, map, sheetRow, category, roomVal, locVal) {
  const status = normalizeStatus_(sourceRow[map.status]);
  const storageId = getOptionalValue_(sourceRow, map.storageId);
  const storageLabel = getOptionalValue_(sourceRow, map.storageLabel);
  return {
    sheetRow: sheetRow,
    itemId: cleanString_(sourceRow[map.itemId]),
    itemName: cleanString_(sourceRow[map.itemName]),
    room: roomVal,
    specificLocation: locVal,
    qty: toNonNegativeNumber_(sourceRow[map.qty]),
    category: category,
    status: status,
    unit: getOptionalValue_(sourceRow, map.unit),
    remarks: getOptionalValue_(sourceRow, map.remarks),
    locationCode: getOptionalValue_(sourceRow, map.locationCode),
    storageId: storageId,
    storageLabel: storageLabel,
    isHazard: isHazardCategory_(category),
    statusClass: statusClassServer_(status),
    displayLocation: storageLabel || locVal,
    routeLoc: storageId || storageLabel || locVal
  };
}

function refreshQrLinks() {
  const sheet = getInventorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return;

  const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = getColumnMap_(header, { requireQrLink: true });
  const baseUrl = getWebAppBaseUrl_();
  const rows = sheet.getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, sheet.getLastColumn()).getValues();

  const output = rows.map(function (row) {
    const room = cleanString_(row[map.room]);
    const loc = cleanString_(row[map.location]);
    if (!room || !loc) return [''];

    const storageId = getOptionalValue_(row, map.storageId);
    const storageLabel = getOptionalValue_(row, map.storageLabel);
    const locationCode = getOptionalValue_(row, map.locationCode);
    const identity = buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode);
    return [buildLocationUrl_(baseUrl, room, identity.routeLoc)];
  });

  sheet.getRange(CONFIG.HEADER_ROW + 1, map.qrLink + 1, output.length, 1).setValues(output);
}

function refreshQrImages() {
  const sheet = getInventorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return;

  const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = getColumnMap_(header, { requireQrLink: true, requireQrImage: true });
  const qrLinkColA1 = columnLetter_(map.qrLink + 1);
  const qrImageCol = map.qrImage + 1;

  for (let row = CONFIG.HEADER_ROW + 1; row <= lastRow; row++) {
    const formula = '=IF(' + qrLinkColA1 + row + '="","",IMAGE("' + CONFIG.QUICKCHART_QR_BASE + '"&ENCODEURL(' + qrLinkColA1 + row + ')))';
    sheet.getRange(row, qrImageCol).setFormula(formula);
  }
}

function setAppConfig(spreadsheetId, webAppBaseUrl, inventorySheetName) {
  const props = PropertiesService.getScriptProperties();
  const nextSpreadsheetId = cleanString_(spreadsheetId);
  const nextWebAppBaseUrl = cleanString_(webAppBaseUrl);
  const nextSheetName = cleanString_(inventorySheetName);

  if (!nextSpreadsheetId) {
    throw new Error('Please provide a non-empty spreadsheet ID.');
  }

  props.setProperty(CONFIG.SPREADSHEET_ID_PROPERTY, nextSpreadsheetId);
  if (nextWebAppBaseUrl) props.setProperty(CONFIG.WEB_APP_URL_PROPERTY, nextWebAppBaseUrl);
  else props.deleteProperty(CONFIG.WEB_APP_URL_PROPERTY);

  if (nextSheetName) props.setProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY, nextSheetName);
  else props.deleteProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY);

  return getAppConfig_();
}

function setWebAppBaseUrl(url) {
  const value = cleanString_(url);
  if (!value) throw new Error('Please provide a non-empty URL.');
  PropertiesService.getScriptProperties().setProperty(CONFIG.WEB_APP_URL_PROPERTY, value);
}

function getSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(CONFIG.SPREADSHEET_ID_PROPERTY);
  if (!spreadsheetId || !spreadsheetId.trim()) {
    throw new Error('SPREADSHEET_ID is not set.');
  }
  return SpreadsheetApp.openById(spreadsheetId.trim());
}

function getInventorySheet_() {
  const ss = getSpreadsheet_();
  const configuredName = PropertiesService.getScriptProperties().getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY);
  const preferredNames = [];

  if (configuredName && configuredName.trim()) preferredNames.push(configuredName.trim());
  preferredNames.push(CONFIG.DEFAULT_SHEET_NAME);

  for (var i = 0; i < preferredNames.length; i++) {
    const sheet = ss.getSheetByName(preferredNames[i]);
    if (sheet) return sheet;
  }

  const first = ss.getSheets()[0];
  if (!first) throw new Error('No sheets found in the configured spreadsheet.');
  return first;
}

function getInventoryDataset_() {
  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    throw new Error('The inventory sheet is empty.');
  }
  return {
    sheet: sheet,
    values: values,
    map: getColumnMap_(values[0], { requireQrLink: false })
  };
}

function getColumnMap_(headerRow, options) {
  const opts = options || {};
  const headers = headerRow.map(normalizeHeader_);
  const map = {};

  Object.keys(CONFIG.ALIASES).forEach(function (key) {
    map[key] = findHeaderIndex_(headers, CONFIG.ALIASES[key]);
  });

  const required = CONFIG.REQUIRED_COLUMNS.slice();
  if (opts.requireQrLink) required.push('qrLink');
  if (opts.requireQrImage) required.push('qrImage');

  const missing = required.filter(function (key) {
    return map[key] === -1;
  });

  if (missing.length) {
    throw new Error('Required column missing: ' + missing.map(function (key) {
      return CONFIG.ALIASES[key][0];
    }).join(', '));
  }

  return map;
}

function normalizeHeader_(value) {
  return String(value || '')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .trim();
}

function findHeaderIndex_(normalizedHeaders, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const alias = normalizeHeader_(aliases[i]);
    const idx = normalizedHeaders.indexOf(alias);
    if (idx !== -1) return idx;
  }
  return -1;
}

function getWebAppBaseUrl_(options) {
  const opts = options || {};
  const propertyUrl = PropertiesService.getScriptProperties().getProperty(CONFIG.WEB_APP_URL_PROPERTY);
  if (propertyUrl && propertyUrl.trim()) return propertyUrl.trim();

  const deployedUrl = ScriptApp.getService().getUrl();
  if (deployedUrl) return deployedUrl;

  if (opts.silent) return '';
  throw new Error('WEB_APP_BASE_URL is not configured.');
}

function getAppConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: cleanString_(props.getProperty(CONFIG.SPREADSHEET_ID_PROPERTY)),
    inventorySheetName: cleanString_(props.getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY)),
    webAppBaseUrl: cleanString_(props.getProperty(CONFIG.WEB_APP_URL_PROPERTY))
  };
}

function getAppConfigSafe_() {
  try {
    return getAppConfig_();
  } catch (err) {
    return {
      spreadsheetId: '',
      inventorySheetName: '',
      webAppBaseUrl: ''
    };
  }
}

function getDiagnostics_() {
  const status = getConfigStatus();
  return status;
}

function getDiagnosticsSafe_() {
  try {
    return getDiagnostics_();
  } catch (err) {
    return { error: err.message || String(err) };
  }
}

function getConfigStatus() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = cleanString_(props.getProperty(CONFIG.SPREADSHEET_ID_PROPERTY));
  const webAppBaseUrl = cleanString_(props.getProperty(CONFIG.WEB_APP_URL_PROPERTY));
  const inventorySheetName = cleanString_(props.getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY));

  const status = {
    scriptProperties: {
      spreadsheetIdConfigured: !!spreadsheetId,
      webAppBaseUrlConfigured: !!webAppBaseUrl,
      inventorySheetNameConfigured: !!inventorySheetName,
      inventorySheetName: inventorySheetName || CONFIG.DEFAULT_SHEET_NAME
    }
  };

  try {
    const sheet = getInventorySheet_();
    const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const normalizedHeaders = header.map(normalizeHeader_);
    const map = {};

    Object.keys(CONFIG.ALIASES).forEach(function (key) {
      map[key] = findHeaderIndex_(normalizedHeaders, CONFIG.ALIASES[key]);
    });

    status.sheetInUse = sheet.getName();
    status.dataRows = Math.max(sheet.getLastRow() - CONFIG.HEADER_ROW, 0);
    status.requiredColumns = {};
    status.optionalColumns = {};

    CONFIG.REQUIRED_COLUMNS.forEach(function (key) {
      status.requiredColumns[key] = map[key] !== -1 ? 'found (col ' + (map[key] + 1) + ')' : 'MISSING';
    });

    CONFIG.OPTIONAL_COLUMNS.forEach(function (key) {
      status.optionalColumns[key] = map[key] !== -1 ? 'found (col ' + (map[key] + 1) + ')' : 'not present';
    });

    status.headerPreview = normalizedHeaders;
  } catch (err) {
    status.sheetError = err.message || String(err);
  }

  Logger.log(JSON.stringify(status, null, 2));
  return status;
}

function buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode) {
  const routeLoc = storageId || storageLabel || locationCode || loc;
  const key = [room, storageId || '', storageLabel || '', locationCode || '', loc].join('||').toLowerCase();
  return {
    routeLoc: routeLoc,
    key: key
  };
}

function buildLocationUrl_(baseUrl, room, loc) {
  return String(baseUrl || '').replace(/\?+$/, '') + '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc);
}

function normalizeMode_(value) {
  return cleanString_(value).toLowerCase() === 'tech' ? 'tech' : 'view';
}

function cleanString_(value) {
  return String(value == null ? '' : value).trim();
}

function getOptionalValue_(row, index) {
  if (typeof index !== 'number' || index < 0) return '';
  return cleanString_(row[index]);
}

function normalizeStatus_(value) {
  const raw = cleanString_(value);
  const match = CONFIG.STATUS_OPTIONS.filter(function (status) {
    return status.toLowerCase() === raw.toLowerCase();
  })[0];
  return match || 'Good';
}

function toNonNegativeNumber_(value) {
  const num = Number(value);
  return Number.isFinite(num) && num >= 0 ? num : 0;
}

function isHazardCategory_(category) {
  return CONFIG.HAZARD_CATEGORIES.indexOf(cleanString_(category).toLowerCase()) !== -1;
}

function statusClassServer_(status) {
  switch (cleanString_(status)) {
    case 'Good':
      return 'text-emerald-700 bg-emerald-100';
    case 'Low Stock':
      return 'text-amber-700 bg-amber-100';
    case 'Missing':
      return 'text-red-700 bg-red-100';
    case 'Needs Maintenance':
      return 'text-orange-800 bg-orange-100';
    default:
      return 'text-stone-700 bg-stone-100';
  }
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function columnLetter_(indexOneBased) {
  let n = indexOneBased;
  let result = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    n = Math.floor((n - mod) / 26);
  }
  return result;
}

function renderInitialItemsHtml_(bootstrap) {
  if (bootstrap.pageType === 'landing') {
    return renderLandingHtml_(bootstrap.locations || [], bootstrap.webAppBaseUrl);
  }
  if (bootstrap.pageType === 'error') {
    return renderErrorStateHtml_(bootstrap.error);
  }
  return renderInventoryHtml_(bootstrap.rows || [], bootstrap.mode);
}

function renderErrorStateHtml_(message) {
  return '<div class="p-5"><div class="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm text-red-800">' + escapeHtml_(message) + '</div></div>';
}

function renderLandingHtml_(locations, webAppBaseUrl) {
  let html = '';
  html += '<div class="border-b border-stone-200 bg-stone-50 px-4 py-4 sm:px-5">';
  html += '<h2 class="text-base font-semibold text-stone-900">Browse storage locations</h2>';
  html += '<p class="mt-1 text-sm text-stone-600">Scan a QR code or search by room, location name, storage ID, or code.</p>';
  html += '<div class="mt-4"><label class="sr-only" for="locationSearch">Search room, location, storage ID, or code</label><input id="locationSearch" type="search" placeholder="Search room, location, storage ID, or code…" class="w-full rounded-xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-900 shadow-sm outline-none transition focus:border-sky-500 focus:ring-2 focus:ring-sky-200" /></div>';
  html += '</div>';

  if (!locations.length) {
    html += '<p class="p-5 text-sm text-stone-500">No locations found in the inventory sheet.</p>';
    return html;
  }

  html += '<div id="locationResults">';
  let currentRoom = '';
  locations.forEach(function (entry) {
    if (entry.room !== currentRoom) {
      currentRoom = entry.room;
      html += '<div class="room-group-header border-y border-stone-200 bg-stone-100 px-4 py-2 text-[11px] font-semibold uppercase tracking-[0.18em] text-stone-600" data-room-group="' + escapeHtml_(entry.room) + '">Room ' + escapeHtml_(entry.room) + '</div>';
    }

    const viewUrl = webAppBaseUrl ? buildLocationUrl_(webAppBaseUrl, entry.room, entry.routeLoc) : '#';
    const techUrl = webAppBaseUrl ? (viewUrl + '&mode=tech') : '#';
    const displayName = entry.displayLoc || entry.loc;

    html += '<article class="location-card border-b border-stone-200 px-4 py-4 last:border-b-0" data-search="' + escapeHtml_(entry.searchText) + '" data-room="' + escapeHtml_(entry.room) + '">';
    html += '<div class="flex items-start justify-between gap-3">';
    html += '<div>';
    html += '<p class="text-base font-semibold text-stone-900">' + escapeHtml_(displayName) + '</p>';
    if (entry.displayLoc !== entry.loc) {
      html += '<p class="mt-0.5 text-xs text-stone-400 italic">Location: ' + escapeHtml_(entry.loc) + '</p>';
    }
    html += '<p class="mt-1 text-xs text-stone-500">Room: ' + escapeHtml_(entry.room) + '</p>';
    if (entry.storageId || entry.locationCode) {
      html += '<div class="mt-2 flex flex-wrap gap-1.5">';
      if (entry.storageId) {
        html += '<span class="inline-flex rounded-full bg-teal-100 px-2.5 py-1 text-[11px] font-semibold text-teal-800">ID: ' + escapeHtml_(entry.storageId) + '</span>';
      }
      if (entry.locationCode) {
        html += '<span class="inline-flex rounded-full bg-sky-100 px-2.5 py-1 text-[11px] font-medium text-sky-800">Code: ' + escapeHtml_(entry.locationCode) + '</span>';
      }
      html += '</div>';
    }
    html += '</div>';
    html += '<div class="flex flex-col gap-2 sm:flex-row">';
    html += '<a class="rounded-xl bg-stone-200 px-3 py-2 text-xs font-medium text-stone-800 transition hover:bg-stone-300 ' + (webAppBaseUrl ? '' : 'pointer-events-none opacity-50') + '" href="' + escapeHtml_(viewUrl) + '">Open View</a>';
    html += '<a class="rounded-xl bg-sky-700 px-3 py-2 text-xs font-medium text-white transition hover:bg-sky-800 ' + (webAppBaseUrl ? '' : 'pointer-events-none opacity-50') + '" href="' + escapeHtml_(techUrl) + '">Open Tech</a>';
    html += '</div></div></article>';
  });
  html += '<p id="locationNoResults" class="hidden p-5 text-sm text-stone-500">No matching locations found.</p>';
  html += '</div>';
  return html;
}

function renderInventoryHtml_(rows, mode) {
  if (!rows.length) {
    return '<div class="p-5"><p class="text-sm text-stone-500">No inventory items have been entered for this storage yet.</p></div>';
  }

  let html = '';
  rows.forEach(function (row) {
    const hazardBadge = row.isHazard
      ? '<span class="inline-flex rounded-full bg-red-200 px-2.5 py-1 text-[11px] font-semibold uppercase tracking-wide text-red-900">Hazard: Chemical</span>'
      : '';
    const hazardNote = row.isHazard
      ? '<p class="mt-3 rounded-xl border border-red-200 bg-red-100 px-3 py-2 text-xs text-red-900">Handle according to chemical storage and safety procedures.</p>'
      : '';

    const meta = [
      'ID: ' + escapeHtml_(row.itemId),
      'Category: ' + escapeHtml_(row.category)
    ];
    if (row.unit) meta.push('Unit: ' + escapeHtml_(row.unit));
    if (row.storageId) meta.push('Storage ID: ' + escapeHtml_(row.storageId));
    if (row.locationCode) meta.push('Code: ' + escapeHtml_(row.locationCode));

    html += '<article class="border-b border-stone-200 p-4 last:border-b-0 ' + (row.isHazard ? 'bg-red-50' : 'bg-white') + '">';
    html += '<div class="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">';
    html += '<div class="min-w-0 flex-1">';

    const storageBadge = row.storageLabel
      ? '<span class="ml-1 inline-flex rounded-full bg-teal-100 px-2 py-0.5 text-[10px] font-semibold text-teal-800">' + escapeHtml_(row.storageLabel) + '</span>'
      : '';

    html += '<div class="flex flex-wrap items-center gap-2"><h3 class="text-base font-semibold text-stone-900">' + escapeHtml_(row.itemName) + '</h3>' + hazardBadge + storageBadge + '</div>';
    html += '<p class="mt-1 text-xs text-stone-500">' + meta.join(' · ') + '</p>';

    if (row.remarks) {
      html += '<p class="mt-2 text-sm text-stone-700"><span class="font-medium text-stone-900">Remarks:</span> ' + escapeHtml_(row.remarks) + '</p>';
    }

    html += hazardNote;
    html += '</div>';

    if (mode === 'tech') {
      const opts = CONFIG.STATUS_OPTIONS.map(function (status) {
        return '<option value="' + escapeHtml_(status) + '" ' + (status === row.status ? 'selected' : '') + '>' + escapeHtml_(status) + '</option>';
      }).join('');

      html += '<div class="grid grid-cols-1 gap-3 sm:min-w-[240px]">';
      if (row.storageId) {
        html += '<p class="text-xs text-stone-500"><span class="font-medium text-stone-700">Storage ID:</span> ' + escapeHtml_(row.storageId) + '</p>';
      }
      if (row.unit) {
        html += '<p class="text-xs text-stone-500"><span class="font-medium text-stone-700">Unit:</span> ' + escapeHtml_(row.unit) + '</p>';
      }
      html += '<label class="block"><span class="text-xs font-medium uppercase tracking-wide text-stone-500">Qty</span><input type="number" min="0" step="1" class="qty-input mt-1 w-full rounded-xl border border-stone-300 px-3 py-2 text-sm text-stone-900 outline-none focus:border-sky-500 focus:ring-2 focus:ring-sky-200" data-row="' + escapeHtml_(String(row.sheetRow)) + '" value="' + escapeHtml_(String(row.qty)) + '" /></label>';
      html += '<label class="block"><span class="text-xs font-medium uppercase tracking-wide text-stone-500">Status</span><select class="status-input mt-1 w-full rounded-xl border border-stone-300 px-3 py-2 text-sm text-stone-900 outline-none focus:border-sky-500 focus:ring-2 focus:ring-sky-200" data-row="' + escapeHtml_(String(row.sheetRow)) + '">' + opts + '</select></label>';
      html += '<p class="inline-flex w-fit rounded-full px-2.5 py-1 text-xs font-medium ' + row.statusClass + '">' + escapeHtml_(row.status) + '</p>';
      html += '</div>';
    } else {
      html += '<div class="rounded-2xl bg-stone-100 px-4 py-3 text-right sm:min-w-[170px]">';
      html += '<p class="text-[11px] font-medium uppercase tracking-wide text-stone-500">Expected Qty</p>';
      html += '<p class="mt-1 text-3xl font-bold text-stone-900">' + escapeHtml_(String(row.qty)) + '</p>';
      if (row.unit) {
        html += '<p class="mt-1 text-xs text-stone-500">' + escapeHtml_(row.unit) + '</p>';
      }
      html += '<p class="mt-2 inline-flex rounded-full px-2.5 py-1 text-xs font-medium ' + row.statusClass + '">' + escapeHtml_(row.status) + '</p>';
      html += '</div>';
    }

    html += '</div></article>';
  });
  return html;
}

function buildPageHtml_(params, bootstrap) {
  const room = bootstrap.room || params.room || '';
  const loc = bootstrap.loc || params.loc || '';
  const mode = bootstrap.mode === 'tech' ? 'tech' : 'view';
  const initialHtml = renderInitialItemsHtml_(bootstrap);
  const isLocationPage = bootstrap.pageType === 'location';
  const modeBadgeLabel = bootstrap.pageType === 'landing'
    ? 'Landing'
    : (bootstrap.pageType === 'error' ? 'Configuration' : (mode === 'tech' ? 'Technician Mode' : 'View Mode'));

  const firstRow = (bootstrap.rows && bootstrap.rows.length > 0) ? bootstrap.rows[0] : null;
  const storageIdDisplay = firstRow && firstRow.storageId ? firstRow.storageId : '';
  const storageLabelDisplay = firstRow && firstRow.storageLabel ? firstRow.storageLabel : '';
  const diagnostics = bootstrap.diagnostics || {};

  const storageHeaderHtml = storageIdDisplay
    ? '<p class="mt-1.5 flex flex-wrap items-center gap-1.5"><span class="inline-flex rounded-full bg-teal-500/30 px-2.5 py-1 text-[11px] font-semibold text-teal-100">Storage ID: ' + escapeHtml_(storageIdDisplay) + '</span>' + (storageLabelDisplay ? '<span class="inline-flex rounded-full bg-white/15 px-2.5 py-1 text-[11px] font-medium text-slate-200">' + escapeHtml_(storageLabelDisplay) + '</span>' : '') + '</p>'
    : '';

  return '<!doctype html>' +
    '<html><head>' +
    '<meta charset="utf-8" />' +
    '<meta name="viewport" content="width=device-width,initial-scale=1" />' +
    '<title>' + escapeHtml_(CONFIG.APP_TITLE) + '</title>' +
    '<script src="https://cdn.tailwindcss.com"></script>' +
    '<script>tailwind.config={theme:{extend:{fontFamily:{sans:[\'Instrument Sans\',\'Segoe UI\',\'sans-serif\']},boxShadow:{panel:\'0 18px 45px rgba(15,23,42,0.10)\'}}}};</script>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com">' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
    '<link href="https://fonts.googleapis.com/css2?family=Instrument+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">' +
    '</head><body class="min-h-screen bg-[radial-gradient(circle_at_top,_rgba(14,165,233,0.14),_transparent_28%),linear-gradient(180deg,_#f8fafc_0%,_#f5f5f4_100%)] text-stone-900">' +
    '<main class="mx-auto max-w-5xl px-4 py-5 sm:px-6 sm:py-8">' +
    '<section class="overflow-hidden rounded-[28px] border border-white/70 bg-white/90 shadow-panel backdrop-blur">' +
    '<header class="border-b border-stone-200 bg-[linear-gradient(135deg,_#0f172a_0%,_#1e293b_55%,_#0f766e_100%)] px-5 py-5 text-white sm:px-6 sm:py-6">' +
    '<div class="flex flex-col gap-5 sm:flex-row sm:items-end sm:justify-between">' +
    '<div>' +
    '<p class="text-[11px] font-semibold uppercase tracking-[0.22em] text-sky-200">Internal Inventory</p>' +
    '<h1 class="mt-2 text-2xl font-semibold sm:text-3xl">' + escapeHtml_(CONFIG.APP_TITLE) + '</h1>' +
    '<p class="mt-2 max-w-2xl text-sm text-slate-200">Room: <span id="roomLabel" class="font-semibold text-white">' + (escapeHtml_(room) || '-') + '</span><span class="mx-2 text-slate-400">/</span>Location: <span id="locLabel" class="font-semibold text-white">' + (escapeHtml_(loc) || '-') + '</span></p>' +
    storageHeaderHtml +
    '</div>' +
    '<div class="flex flex-wrap gap-2">' +
    '<a id="viewModeLink" class="rounded-xl bg-white/15 px-3.5 py-2 text-xs font-medium text-white transition hover:bg-white/25" href="#">View Mode</a>' +
    '<a id="techModeLink" class="rounded-xl bg-sky-400 px-3.5 py-2 text-xs font-medium text-slate-950 transition hover:bg-sky-300" href="#">Technician Access</a>' +
    '</div></div></header>' +
    '<section class="px-4 pt-4 sm:px-6">' +
    '<section id="notice" class="hidden rounded-2xl px-4 py-3 text-sm"></section>' +
    '<section id="bridgeWarning" class="hidden mt-3 rounded-2xl bg-amber-100 px-4 py-3 text-sm text-amber-900"></section>' +
    '<section id="warningList" class="' + ((bootstrap.warnings && bootstrap.warnings.length) ? '' : 'hidden') + ' mt-3 space-y-2">' +
    (bootstrap.warnings || []).map(function (warning) {
      return '<div class="rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900">' + escapeHtml_(warning) + '</div>';
    }).join('') +
    '</section>' +
    '<section id="errorBanner" class="' + (bootstrap.error ? '' : 'hidden') + ' mt-3 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800">' + escapeHtml_(bootstrap.error || '') + '</section>' +
    '<section id="messageBanner" class="' + (bootstrap.message ? '' : 'hidden') + ' mt-3 rounded-2xl border border-sky-200 bg-sky-50 px-4 py-3 text-sm text-sky-900">' + escapeHtml_(bootstrap.message || '') + '</section>' +
    (CONFIG.DEBUG_PANEL ? '<section id="debugPanel" class="mt-3 rounded-2xl bg-slate-900 px-4 py-3 text-xs text-slate-100"></section>' : '') +
    '</section>' +
    '<section class="px-4 pb-4 pt-4 sm:px-6 sm:pb-6">' +
    '<div class="overflow-hidden rounded-[24px] border border-stone-200 bg-white">' +
    '<div class="flex items-center justify-between border-b border-stone-200 bg-stone-50 px-4 py-3">' +
    '<div><h2 class="text-sm font-semibold text-stone-900">' + (isLocationPage ? 'Inventory Items' : 'Storage Locations') + '</h2><p class="mt-0.5 text-xs text-stone-500">' + (isLocationPage ? 'Review expected stock and update technician records.' : 'Select a location to view or update inventory.') + '</p></div>' +
    '<span id="modeBadge" class="rounded-full px-2.5 py-1 text-xs font-medium">' + escapeHtml_(modeBadgeLabel) + '</span>' +
    '</div>' +
    '<div id="loading" class="hidden p-4 text-sm text-stone-500">Saving updates...</div>' +
    '<div id="items">' + initialHtml + '</div>' +
    '</div>' +
    '<footer class="mt-4 flex justify-end">' +
    '<button id="saveBtn" class="' + ((mode === 'tech' && isLocationPage) ? '' : 'hidden ') + 'rounded-xl bg-emerald-600 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-emerald-700 disabled:cursor-not-allowed disabled:opacity-50" type="button">Save Updates</button>' +
    '</footer>' +
    '</section>' +
    '</section>' +
    '</main>' +
    '<script>' +
    'const APP={room:' + JSON.stringify(room) + ',loc:' + JSON.stringify(loc) + ',mode:' + JSON.stringify(mode) + ',bootstrap:' + JSON.stringify(bootstrap) + ',diagnostics:' + JSON.stringify(diagnostics) + '};' +
    'document.addEventListener("DOMContentLoaded",init);' +
    'function init(){buildModeLinks();renderModeBadge();setBridgeWarning();bindLandingSearch();if(' + (CONFIG.DEBUG_PANEL ? 'true' : 'false') + '){renderDebug();}var saveBtn=document.getElementById("saveBtn");if(saveBtn){saveBtn.addEventListener("click",saveUpdates);}}' +
    'function buildModeLinks(){var view=document.getElementById("viewModeLink");var tech=document.getElementById("techModeLink");if(!APP.room||!APP.loc){view.classList.add("opacity-50","pointer-events-none");tech.classList.add("opacity-50","pointer-events-none");view.title="Choose a location below first";tech.title="Choose a location below first";view.href="#";tech.href="#";return;}var base=window.location.pathname+"?room="+encodeURIComponent(APP.room)+"&loc="+encodeURIComponent(APP.loc);view.href=base;tech.href=base+"&mode=tech";}' +
    'function renderModeBadge(){var badge=document.getElementById("modeBadge");var saveBtn=document.getElementById("saveBtn");if(APP.bootstrap.pageType==="landing"){badge.className="rounded-full bg-sky-100 px-2.5 py-1 text-xs font-medium text-sky-800";badge.textContent="Landing";if(saveBtn)saveBtn.classList.add("hidden");return;}if(APP.bootstrap.pageType==="error"){badge.className="rounded-full bg-red-100 px-2.5 py-1 text-xs font-medium text-red-800";badge.textContent="Configuration";if(saveBtn)saveBtn.classList.add("hidden");return;}if(APP.mode==="tech"){badge.className="rounded-full bg-sky-100 px-2.5 py-1 text-xs font-medium text-sky-800";badge.textContent="Technician Mode";if(saveBtn)saveBtn.classList.remove("hidden");}else{badge.className="rounded-full bg-stone-200 px-2.5 py-1 text-xs font-medium text-stone-700";badge.textContent="View Mode";if(saveBtn)saveBtn.classList.add("hidden");}}' +
    'function hasBridge(){return !!(window.google&&google.script&&google.script.run);}' +
    'function setBridgeWarning(){if(APP.mode!=="tech")return;if(hasBridge())return;var warn=document.getElementById("bridgeWarning");warn.textContent="Interactive save is unavailable in this context. Open the deployed /exec web app URL to use full functionality.";warn.classList.remove("hidden");}' +
    'function bindLandingSearch(){var input=document.getElementById("locationSearch");if(!input)return;input.addEventListener("input",filterLocations);}' +
    'function filterLocations(){var input=document.getElementById("locationSearch");var query=(input.value||"").trim().toLowerCase();var cards=Array.prototype.slice.call(document.querySelectorAll(".location-card"));var visibleCount=0;cards.forEach(function(card){var haystack=card.getAttribute("data-search")||"";var matches=!query||haystack.indexOf(query)!==-1;card.classList.toggle("hidden",!matches);if(matches)visibleCount+=1;});var headers=Array.prototype.slice.call(document.querySelectorAll(".room-group-header"));headers.forEach(function(header){var roomName=header.getAttribute("data-room-group");var roomCards=cards.filter(function(card){return card.getAttribute("data-room")===roomName&&!card.classList.contains("hidden");});header.classList.toggle("hidden",roomCards.length===0);});var empty=document.getElementById("locationNoResults");if(empty)empty.classList.toggle("hidden",visibleCount!==0);}' +
    'function renderDebug(){var el=document.getElementById("debugPanel");if(!el)return;var rowCount=APP.bootstrap&&APP.bootstrap.rows?APP.bootstrap.rows.length:0;var locCount=APP.bootstrap&&APP.bootstrap.locations?APP.bootstrap.locations.length:0;var urlConfigured=APP.bootstrap&&APP.bootstrap.webAppBaseUrl?"yes":"no";var sheetName=(APP.diagnostics&&APP.diagnostics.sheetInUse)||"-";el.innerHTML="mode="+esc(APP.mode)+" | room="+esc(APP.room||"-")+" | loc="+esc(APP.loc||"-")+" | bridge="+(hasBridge()?"yes":"no")+" | urlConfigured="+urlConfigured+" | sheet="+esc(sheetName)+" | rows="+rowCount+" | locations="+locCount;}' +
    'function saveUpdates(){if(!hasBridge()){showNotice("Interactive save is unavailable in this context. Open the deployed /exec web app URL.",true);return;}var saveBtn=document.getElementById("saveBtn");var loading=document.getElementById("loading");saveBtn.disabled=true;loading.classList.remove("hidden");var byRow={};var invalidQty=false;document.querySelectorAll(".qty-input").forEach(function(input){var row=input.getAttribute("data-row");var qty=Number(input.value);if(!Number.isFinite(qty)||qty<0)invalidQty=true;byRow[row]=byRow[row]||{sheetRow:Number(row)};byRow[row].qty=qty;});if(invalidQty){saveBtn.disabled=false;loading.classList.add("hidden");showNotice("Please enter valid non-negative quantities before saving.",true);return;}document.querySelectorAll(".status-input").forEach(function(select){var row=select.getAttribute("data-row");byRow[row]=byRow[row]||{sheetRow:Number(row)};byRow[row].status=select.value;});google.script.run.withSuccessHandler(function(res){saveBtn.disabled=false;loading.classList.add("hidden");APP.bootstrap.rows=res.rows||[];document.getElementById("items").innerHTML=res.html||"";showNotice("Saved "+res.updatedCount+" item(s) successfully.",false);}).withFailureHandler(function(err){saveBtn.disabled=false;loading.classList.add("hidden");showNotice((err&&err.message)||"Save failed. Please try again.",true);}).saveInventoryUpdates({room:APP.room,loc:APP.loc,updates:Object.keys(byRow).map(function(key){return byRow[key];})});}' +
    'function showNotice(message,isError){if(!message)return;var el=document.getElementById("notice");el.textContent=message;el.className="rounded-2xl px-4 py-3 text-sm mt-0";if(isError){el.classList.add("bg-red-100","text-red-800","border","border-red-200");}else{el.classList.add("bg-emerald-100","text-emerald-800","border","border-emerald-200");}el.scrollIntoView({behavior:"smooth",block:"nearest"});}' +
    'function esc(value){return String(value==null?"":value).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\"/g,"&quot;").replace(/\'/g,"&#39;");}' +
    '</script></body></html>';
}
