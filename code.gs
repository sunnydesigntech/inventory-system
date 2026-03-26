/**
 * D&T QR Inventory System - standalone single-file Google Apps Script app.
 */
const CONFIG = {
  APP_TITLE: 'D&T QR Inventory',
  HEADER_ROW: 1,
  SPREADSHEET_ID_PROPERTY: 'SPREADSHEET_ID',
  WEB_APP_URL_PROPERTY: 'WEB_APP_BASE_URL',
  INVENTORY_SHEET_NAME_PROPERTY: 'INVENTORY_SHEET_NAME',
  DEFAULT_SHEET_NAME: 'Inventory',
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
    qrLink: ['qr code link (auto-generated)', 'qr code link', 'auto-generated qr link', 'qr link', 'qr url'],
    qrImage: ['qr code image', 'qr image'],
    unit: ['unit'],
    remarks: ['remarks', 'remark', 'notes', 'note'],
    locationCode: ['location code', 'locationcode', 'location id'],
    storageId: ['storage id', 'storageid', 'storage_id'],
    storageLabel: ['storage label', 'storagelabel', 'storage_label']
  }
};

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

function getRequestParams_(e) {
  return {
    room: String((e && e.parameter && e.parameter.room) || '').trim(),
    loc: String((e && e.parameter && e.parameter.loc) || '').trim(),
    mode: String((e && e.parameter && e.parameter.mode) || 'view').trim().toLowerCase() === 'tech' ? 'tech' : 'view'
  };
}

function buildBootstrapData_(params) {
  const config = getAppConfig_();
  const warnings = [];
  const webAppBaseUrl = getWebAppBaseUrl_({ silent: true });
  if (!webAppBaseUrl) {
    warnings.push('WEB_APP_BASE_URL is not configured. Location links are disabled until this is set.');
  }

  const hasLocation = !!(params.room && params.loc);
  if (hasLocation) {
    const result = getInventoryData({ room: params.room, loc: params.loc });
    return {
      pageType: 'location',
      appTitle: CONFIG.APP_TITLE,
      room: result.room,
      loc: result.loc,
      mode: params.mode,
      rows: result.rows || [],
      locations: [],
      message: result.message || '',
      error: '',
      warnings: warnings,
      config: config,
      webAppBaseUrl: webAppBaseUrl
    };
  }

  const locResult = getAllLocations();
  return {
    pageType: 'landing',
    appTitle: CONFIG.APP_TITLE,
    room: '',
    loc: '',
    mode: 'view',
    rows: [],
    locations: locResult.locations || [],
    message: '',
    error: '',
    warnings: warnings,
    config: config,
    webAppBaseUrl: webAppBaseUrl
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
    config: {
      spreadsheetId: '',
      inventorySheetName: '',
      webAppBaseUrl: ''
    },
    webAppBaseUrl: ''
  };
}

function getInventoryData(params) {
  const room = String((params && params.room) || '').trim();
  const loc = String((params && params.loc) || '').trim();

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

function rowMatchesRoomLoc_(rowValues, map, roomNeedle, locNeedle) {
  const rowRoom = String(rowValues[map.room] || '').trim().toLowerCase();
  if (rowRoom !== roomNeedle) return false;
  const rowLoc = String(rowValues[map.location] || '').trim().toLowerCase();
  if (rowLoc === locNeedle) return true;
  const rowStorageId = getOptionalValue_(rowValues, map.storageId).toLowerCase();
  if (rowStorageId && rowStorageId === locNeedle) return true;
  const rowStorageLabel = getOptionalValue_(rowValues, map.storageLabel).toLowerCase();
  if (rowStorageLabel && rowStorageLabel === locNeedle) return true;
  const rowLocationCode = getOptionalValue_(rowValues, map.locationCode).toLowerCase();
  if (rowLocationCode && rowLocationCode === locNeedle) return true;
  return false;
}

function getInventoryRowsForLocation_(room, loc) {
  const data = getInventoryDataset_();
  const map = data.map;
  const values = data.values;
  const roomNeedle = room.toLowerCase();
  const locNeedle = loc.toLowerCase();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const sourceRow = values[i];
    const roomVal = String(sourceRow[map.room] || '').trim();
    if (!roomVal) continue;
    if (!rowMatchesRoomLoc_(sourceRow, map, roomNeedle, locNeedle)) continue;
    const locVal = String(sourceRow[map.location] || '').trim();
    const category = String(sourceRow[map.category] || '').trim();
    rows.push(buildInventoryRowView_(sourceRow, map, i + 1, category, roomVal, locVal));
  }

  return rows;
}

function buildInventoryRowView_(sourceRow, map, sheetRow, category, roomVal, locVal) {
  const status = normalizeStatus_(sourceRow[map.status]);
  return {
    sheetRow: sheetRow,
    itemId: String(sourceRow[map.itemId] || '').trim(),
    itemName: String(sourceRow[map.itemName] || '').trim(),
    room: roomVal,
    specificLocation: locVal,
    qty: toNonNegativeNumber_(sourceRow[map.qty]),
    category: category,
    status: status,
    unit: getOptionalValue_(sourceRow, map.unit),
    remarks: getOptionalValue_(sourceRow, map.remarks),
    locationCode: getOptionalValue_(sourceRow, map.locationCode),
    storageId: getOptionalValue_(sourceRow, map.storageId),
    storageLabel: getOptionalValue_(sourceRow, map.storageLabel),
    isHazard: isHazardCategory_(category),
    statusClass: statusClassServer_(status)
  };
}

function saveInventoryUpdates(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid save payload.');
  }

  const room = String(payload.room || '').trim();
  const loc = String(payload.loc || '').trim();
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
  const qtyCol = map.qty + 1;
  const statusCol = map.status + 1;
  const roomNeedle = room.toLowerCase();
  const locNeedle = loc.toLowerCase();
  const maxRow = values.length;
  const seenRows = {};

  payload.updates.forEach(function (update) {
    const rowNum = Number(update.sheetRow);
    if (!Number.isInteger(rowNum) || rowNum <= CONFIG.HEADER_ROW || rowNum > maxRow) {
      throw new Error('Invalid row number: ' + update.sheetRow);
    }
    if (seenRows[rowNum]) {
      throw new Error('Duplicate row in payload: ' + rowNum);
    }
    seenRows[rowNum] = true;

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

    sheet.getRange(rowNum, qtyCol).setValue(qty);
    sheet.getRange(rowNum, statusCol).setValue(status);
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
  const data = getInventoryDataset_();
  const map = data.map;
  const values = data.values;
  const seen = {};
  const locations = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const room = String(row[map.room] || '').trim();
    const loc = String(row[map.location] || '').trim();
    if (!room || !loc) continue;

    const locationCode = getOptionalValue_(row, map.locationCode);
    const storageId = getOptionalValue_(row, map.storageId);
    const storageLabel = getOptionalValue_(row, map.storageLabel);
    const key = room.toLowerCase() + '||' + loc.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    locations.push({
      room: room,
      loc: loc,
      locationCode: locationCode,
      storageId: storageId,
      storageLabel: storageLabel,
      searchText: [room, loc, locationCode, storageId, storageLabel].filter(Boolean).join(' ').toLowerCase()
    });
  }

  locations.sort(function (a, b) {
    if (a.room === b.room) return a.loc.localeCompare(b.loc);
    return a.room.localeCompare(b.room);
  });

  return locations;
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
    const room = String(row[map.room] || '').trim();
    const loc = String(row[map.location] || '').trim();
    if (!room || !loc) return [''];
    // Prefer Storage ID as the authoritative loc parameter if present
    const storageId = (map.storageId !== undefined && map.storageId !== -1)
      ? String(row[map.storageId] || '').trim()
      : '';
    const locParam = storageId || loc;
    return [buildLocationUrl_(baseUrl, room, locParam)];
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
  const nextSpreadsheetId = String(spreadsheetId || '').trim();
  const nextWebAppBaseUrl = String(webAppBaseUrl || '').trim();
  const nextSheetName = String(inventorySheetName || '').trim();

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
  const value = String(url || '').trim();
  if (!value) throw new Error('Please provide a non-empty URL.');
  PropertiesService.getScriptProperties().setProperty(CONFIG.WEB_APP_URL_PROPERTY, value);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('D&T Inventory')
    .addItem('Refresh QR Links', 'refreshQrLinks')
    .addItem('Open Web App', 'openWebApp_')
    .addItem('Refresh QR Images (Optional)', 'refreshQrImages')
    .addSeparator()
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

  const spreadsheetId = ui.prompt('Set SPREADSHEET_ID', 'Paste the Google Sheet ID:', ui.ButtonSet.OK_CANCEL);
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

  setAppConfig(spreadsheetId.getResponseText(), webAppUrl.getResponseText(), sheetName.getResponseText());
  ui.alert('App configuration saved.');
}

function promptSetWebAppBaseUrl_() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set WEB_APP_BASE_URL', 'Paste your deployed /exec web app URL:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  setWebAppBaseUrl(result.getResponseText());
  ui.alert('WEB_APP_BASE_URL saved.');
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

  for (let i = 0; i < preferredNames.length; i++) {
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
  const map = {
    itemId: findHeaderIndex_(headers, CONFIG.ALIASES.itemId),
    itemName: findHeaderIndex_(headers, CONFIG.ALIASES.itemName),
    room: findHeaderIndex_(headers, CONFIG.ALIASES.room),
    location: findHeaderIndex_(headers, CONFIG.ALIASES.location),
    qty: findHeaderIndex_(headers, CONFIG.ALIASES.qty),
    category: findHeaderIndex_(headers, CONFIG.ALIASES.category),
    status: findHeaderIndex_(headers, CONFIG.ALIASES.status),
    qrLink: findHeaderIndex_(headers, CONFIG.ALIASES.qrLink),
    qrImage: findHeaderIndex_(headers, CONFIG.ALIASES.qrImage),
    unit: findHeaderIndex_(headers, CONFIG.ALIASES.unit),
    remarks: findHeaderIndex_(headers, CONFIG.ALIASES.remarks),
    locationCode: findHeaderIndex_(headers, CONFIG.ALIASES.locationCode),
    storageId: findHeaderIndex_(headers, CONFIG.ALIASES.storageId),
    storageLabel: findHeaderIndex_(headers, CONFIG.ALIASES.storageLabel)
  };

  const required = ['itemId', 'itemName', 'room', 'location', 'qty', 'category', 'status'];
  if (opts.requireQrLink) required.push('qrLink');
  if (opts.requireQrImage) required.push('qrImage');

  const missing = required.filter(function (key) {
    return map[key] === -1;
  });
  if (missing.length) {
    const label = missing.map(function (key) {
      return CONFIG.ALIASES[key][0];
    }).join(', ');
    throw new Error('Required column missing: ' + label);
  }

  return map;
}

function normalizeHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function findHeaderIndex_(normalizedHeaders, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const idx = normalizedHeaders.indexOf(normalizeHeader_(aliases[i]));
    if (idx !== -1) return idx;
  }
  return -1;
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
    spreadsheetId: String(props.getProperty(CONFIG.SPREADSHEET_ID_PROPERTY) || '').trim(),
    inventorySheetName: String(props.getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY) || '').trim(),
    webAppBaseUrl: String(props.getProperty(CONFIG.WEB_APP_URL_PROPERTY) || '').trim()
  };
}

function getConfigStatus() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = String(props.getProperty(CONFIG.SPREADSHEET_ID_PROPERTY) || '').trim();
  const webAppBaseUrl = String(props.getProperty(CONFIG.WEB_APP_URL_PROPERTY) || '').trim();
  const inventorySheetName = String(props.getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY) || '').trim();

  const status = {
    SPREADSHEET_ID: spreadsheetId ? 'SET (' + spreadsheetId.substring(0, 8) + '...)' : 'NOT SET — required',
    WEB_APP_BASE_URL: webAppBaseUrl ? 'SET' : 'NOT SET — QR links disabled',
    INVENTORY_SHEET_NAME: inventorySheetName || '(using default: ' + CONFIG.DEFAULT_SHEET_NAME + ')'
  };

  try {
    const data = getInventoryDataset_();
    const map = data.map;
    status.sheetInUse = data.sheet.getName();
    status.dataRows = data.values.length - 1;

    const reqCols = ['itemId', 'itemName', 'room', 'location', 'qty', 'category', 'status'];
    const optCols = ['unit', 'remarks', 'locationCode', 'storageId', 'storageLabel', 'qrLink', 'qrImage'];

    status.requiredColumns = {};
    reqCols.forEach(function (k) {
      status.requiredColumns[k] = map[k] !== -1 ? 'found (col ' + (map[k] + 1) + ')' : 'MISSING — required';
    });

    status.optionalColumns = {};
    optCols.forEach(function (k) {
      status.optionalColumns[k] = map[k] !== -1 ? 'found (col ' + (map[k] + 1) + ')' : 'not present (optional)';
    });
  } catch (e) {
    status.sheetError = e.message;
  }

  Logger.log(JSON.stringify(status, null, 2));
  return status;
}

function toNonNegativeNumber_(value) {
  const num = Number(value);
  return Number.isFinite(num) && num >= 0 ? num : 0;
}

function normalizeStatus_(value) {
  const status = String(value || '').trim();
  return status || 'Good';
}

function getOptionalValue_(row, index) {
  if (typeof index !== 'number' || index < 0) return '';
  return String(row[index] || '').trim();
}

function isHazardCategory_(category) {
  return CONFIG.HAZARD_CATEGORIES.indexOf(String(category || '').trim().toLowerCase()) !== -1;
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function statusClassServer_(status) {
  switch (String(status || '')) {
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

function renderInitialItemsHtml_(bootstrap) {
  if (bootstrap.pageType === 'landing') {
    return renderLandingHtml_(bootstrap.locations || []);
  }
  if (bootstrap.pageType === 'error') {
    return renderErrorStateHtml_(bootstrap.error);
  }
  return renderInventoryHtml_(bootstrap.rows || [], bootstrap.mode);
}

function renderErrorStateHtml_(message) {
  return '<div class="p-5"><div class="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm text-red-800">' + escapeHtml_(message) + '</div></div>';
}

function renderLandingHtml_(locations) {
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

    const viewUrl = '?room=' + encodeURIComponent(entry.room) + '&loc=' + encodeURIComponent(entry.loc);
    const techUrl = viewUrl + '&mode=tech';
    const displayName = entry.storageLabel || entry.loc;
    html += '<article class="location-card border-b border-stone-200 px-4 py-4 last:border-b-0" data-search="' + escapeHtml_(entry.searchText) + '" data-room="' + escapeHtml_(entry.room) + '">';
    html += '<div class="flex items-start justify-between gap-3">';
    html += '<div>';
    html += '<p class="text-base font-semibold text-stone-900">' + escapeHtml_(displayName) + '</p>';
    if (entry.storageLabel && entry.storageLabel !== entry.loc) {
      html += '<p class="mt-0.5 text-xs text-stone-400 italic">Loc: ' + escapeHtml_(entry.loc) + '</p>';
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
    html += '<a class="rounded-xl bg-stone-200 px-3 py-2 text-xs font-medium text-stone-800 transition hover:bg-stone-300" href="' + escapeHtml_(viewUrl) + '">Open View</a>';
    html += '<a class="rounded-xl bg-sky-700 px-3 py-2 text-xs font-medium text-white transition hover:bg-sky-800" href="' + escapeHtml_(techUrl) + '">Open Tech</a>';
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
      html += '<div class="grid grid-cols-1 gap-3 sm:min-w-[220px]">';
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
      html += '<p class="mt-2 inline-flex rounded-full px-2.5 py-1 text-xs font-medium ' + row.statusClass + '">' + escapeHtml_(row.status) + '</p>';
      html += '</div>';
    }

    html += '</div></article>';
  });
  return html;
}

function buildLocationUrl_(baseUrl, room, loc) {
  return baseUrl + '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc);
}

function buildPageHtml_(params, bootstrap) {
  const room = bootstrap.room || params.room || '';
  const loc = bootstrap.loc || params.loc || '';
  const mode = bootstrap.mode === 'tech' ? 'tech' : 'view';
  const initialHtml = renderInitialItemsHtml_(bootstrap);
  const isLocationPage = bootstrap.pageType === 'location';
  const modeBadgeLabel = bootstrap.pageType === 'landing' ? 'Landing' : (mode === 'tech' ? 'Technician Mode' : 'View Mode');

  const firstRow = (bootstrap.rows && bootstrap.rows.length > 0) ? bootstrap.rows[0] : null;
  const storageIdDisplay = firstRow && firstRow.storageId ? firstRow.storageId : '';
  const storageLabelDisplay = firstRow && firstRow.storageLabel ? firstRow.storageLabel : '';
  const storageHeaderHtml = storageIdDisplay
    ? '<p class="mt-1.5 flex flex-wrap items-center gap-1.5"><span class="inline-flex rounded-full bg-teal-500/30 px-2.5 py-1 text-[11px] font-semibold text-teal-100">Storage ID: ' + escapeHtml_(storageIdDisplay) + '</span>' + (storageLabelDisplay ? '<span class="inline-flex rounded-full bg-white/15 px-2.5 py-1 text-[11px] font-medium text-slate-200">' + escapeHtml_(storageLabelDisplay) + '</span>' : '') + '</p>'
    : '';

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>${escapeHtml_(CONFIG.APP_TITLE)}</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {
      theme: {
        extend: {
          fontFamily: {
            sans: ['Instrument Sans', 'Segoe UI', 'sans-serif']
          },
          boxShadow: {
            panel: '0 18px 45px rgba(15, 23, 42, 0.10)'
          }
        }
      }
    };
  </script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Instrument+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
</head>
<body class="min-h-screen bg-[radial-gradient(circle_at_top,_rgba(14,165,233,0.14),_transparent_28%),linear-gradient(180deg,_#f8fafc_0%,_#f5f5f4_100%)] text-stone-900">
  <main class="mx-auto max-w-5xl px-4 py-5 sm:px-6 sm:py-8">
    <section class="overflow-hidden rounded-[28px] border border-white/70 bg-white/90 shadow-panel backdrop-blur">
      <header class="border-b border-stone-200 bg-[linear-gradient(135deg,_#0f172a_0%,_#1e293b_55%,_#0f766e_100%)] px-5 py-5 text-white sm:px-6 sm:py-6">
        <div class="flex flex-col gap-5 sm:flex-row sm:items-end sm:justify-between">
          <div>
            <p class="text-[11px] font-semibold uppercase tracking-[0.22em] text-sky-200">Internal Inventory</p>
            <h1 class="mt-2 text-2xl font-semibold sm:text-3xl">${escapeHtml_(CONFIG.APP_TITLE)}</h1>
            <p class="mt-2 max-w-2xl text-sm text-slate-200">
              Room: <span id="roomLabel" class="font-semibold text-white">${escapeHtml_(room) || '-'}</span>
              <span class="mx-2 text-slate-400">/</span>
              Location: <span id="locLabel" class="font-semibold text-white">${escapeHtml_(loc) || '-'}</span>
            </p>
            ${storageHeaderHtml}
          </div>
          <div class="flex flex-wrap gap-2">
            <a id="viewModeLink" class="rounded-xl bg-white/15 px-3.5 py-2 text-xs font-medium text-white transition hover:bg-white/25" href="#">View Mode</a>
            <a id="techModeLink" class="rounded-xl bg-sky-400 px-3.5 py-2 text-xs font-medium text-slate-950 transition hover:bg-sky-300" href="#">Technician Access</a>
          </div>
        </div>
      </header>

      <section class="px-4 pt-4 sm:px-6">
        <section id="notice" class="hidden rounded-2xl px-4 py-3 text-sm"></section>
        <section id="bridgeWarning" class="hidden mt-3 rounded-2xl bg-amber-100 px-4 py-3 text-sm text-amber-900"></section>
        <section id="warningList" class="${bootstrap.warnings && bootstrap.warnings.length ? '' : 'hidden'} mt-3 space-y-2">
          ${(bootstrap.warnings || []).map(function (warning) {
            return '<div class="rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900">' + escapeHtml_(warning) + '</div>';
          }).join('')}
        </section>
        <section id="errorBanner" class="${bootstrap.error ? '' : 'hidden'} mt-3 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800">${escapeHtml_(bootstrap.error || '')}</section>
        <section id="messageBanner" class="${bootstrap.message ? '' : 'hidden'} mt-3 rounded-2xl border border-sky-200 bg-sky-50 px-4 py-3 text-sm text-sky-900">${escapeHtml_(bootstrap.message || '')}</section>
        ${CONFIG.DEBUG_PANEL ? '<section id="debugPanel" class="mt-3 rounded-2xl bg-slate-900 px-4 py-3 text-xs text-slate-100"></section>' : ''}
      </section>

      <section class="px-4 pb-4 pt-4 sm:px-6 sm:pb-6">
        <div class="overflow-hidden rounded-[24px] border border-stone-200 bg-white">
          <div class="flex items-center justify-between border-b border-stone-200 bg-stone-50 px-4 py-3">
            <div>
              <h2 class="text-sm font-semibold text-stone-900">${isLocationPage ? 'Inventory Items' : 'Storage Locations'}</h2>
              <p class="mt-0.5 text-xs text-stone-500">${isLocationPage ? 'Review expected stock and update technician records.' : 'Select a location to view or update inventory.'}</p>
            </div>
            <span id="modeBadge" class="rounded-full px-2.5 py-1 text-xs font-medium">${escapeHtml_(modeBadgeLabel)}</span>
          </div>
          <div id="loading" class="hidden p-4 text-sm text-stone-500">Saving updates...</div>
          <div id="items">${initialHtml}</div>
        </div>

        <footer class="mt-4 flex justify-end">
          <button id="saveBtn" class="${mode === 'tech' && isLocationPage ? '' : 'hidden '}rounded-xl bg-emerald-600 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-emerald-700 disabled:cursor-not-allowed disabled:opacity-50" type="button">Save Updates</button>
        </footer>
      </section>
    </section>
  </main>

  <script>
    const APP = {
      room: ${JSON.stringify(room)},
      loc: ${JSON.stringify(loc)},
      mode: ${JSON.stringify(mode)},
      bootstrap: ${JSON.stringify(bootstrap)}
    };

    document.addEventListener('DOMContentLoaded', init);

    function init() {
      buildModeLinks();
      renderModeBadge();
      setBridgeWarning();
      bindLandingSearch();
      if (${CONFIG.DEBUG_PANEL ? 'true' : 'false'}) renderDebug();

      const saveBtn = document.getElementById('saveBtn');
      if (saveBtn) saveBtn.addEventListener('click', saveUpdates);
    }

    function buildModeLinks() {
      const view = document.getElementById('viewModeLink');
      const tech = document.getElementById('techModeLink');
      if (!APP.room || !APP.loc) {
        view.classList.add('opacity-50', 'pointer-events-none');
        tech.classList.add('opacity-50', 'pointer-events-none');
        view.setAttribute('title', 'Choose a location below first');
        tech.setAttribute('title', 'Choose a location below first');
        view.href = '#';
        tech.href = '#';
        return;
      }
      const base = window.location.pathname + '?room=' + encodeURIComponent(APP.room) + '&loc=' + encodeURIComponent(APP.loc);
      view.href = base;
      tech.href = base + '&mode=tech';
    }

    function renderModeBadge() {
      const badge = document.getElementById('modeBadge');
      const saveBtn = document.getElementById('saveBtn');
      if (APP.bootstrap.pageType === 'landing') {
        badge.className = 'rounded-full bg-sky-100 px-2.5 py-1 text-xs font-medium text-sky-800';
        badge.textContent = 'Landing';
        if (saveBtn) saveBtn.classList.add('hidden');
        return;
      }
      if (APP.bootstrap.pageType === 'error') {
        badge.className = 'rounded-full bg-red-100 px-2.5 py-1 text-xs font-medium text-red-800';
        badge.textContent = 'Configuration';
        if (saveBtn) saveBtn.classList.add('hidden');
        return;
      }
      if (APP.mode === 'tech') {
        badge.className = 'rounded-full bg-sky-100 px-2.5 py-1 text-xs font-medium text-sky-800';
        badge.textContent = 'Technician Mode';
        if (saveBtn) saveBtn.classList.remove('hidden');
      } else {
        badge.className = 'rounded-full bg-stone-200 px-2.5 py-1 text-xs font-medium text-stone-700';
        badge.textContent = 'View Mode';
        if (saveBtn) saveBtn.classList.add('hidden');
      }
    }

    function hasBridge() {
      return !!(window.google && google.script && google.script.run);
    }

    function setBridgeWarning() {
      if (APP.mode !== 'tech') return;
      if (hasBridge()) return;
      const warn = document.getElementById('bridgeWarning');
      warn.textContent = 'Interactive save is unavailable in this context. Open the deployed /exec web app URL to use full functionality.';
      warn.classList.remove('hidden');
    }

    function bindLandingSearch() {
      const input = document.getElementById('locationSearch');
      if (!input) return;
      input.addEventListener('input', filterLocations);
    }

    function filterLocations() {
      const input = document.getElementById('locationSearch');
      const query = (input.value || '').trim().toLowerCase();
      const cards = Array.prototype.slice.call(document.querySelectorAll('.location-card'));
      let visibleCount = 0;

      cards.forEach(function (card) {
        const haystack = card.getAttribute('data-search') || '';
        const matches = !query || haystack.indexOf(query) !== -1;
        card.classList.toggle('hidden', !matches);
        if (matches) visibleCount += 1;
      });

      const headers = Array.prototype.slice.call(document.querySelectorAll('.room-group-header'));
      headers.forEach(function (header) {
        const roomName = header.getAttribute('data-room-group');
        const roomCards = cards.filter(function (card) {
          return card.getAttribute('data-room') === roomName && !card.classList.contains('hidden');
        });
        header.classList.toggle('hidden', roomCards.length === 0);
      });

      const empty = document.getElementById('locationNoResults');
      if (empty) empty.classList.toggle('hidden', visibleCount !== 0);
    }

    function renderDebug() {
      const el = document.getElementById('debugPanel');
      if (!el) return;
      const rowCount = APP.bootstrap && APP.bootstrap.rows ? APP.bootstrap.rows.length : 0;
      const locCount = APP.bootstrap && APP.bootstrap.locations ? APP.bootstrap.locations.length : 0;
      el.innerHTML = 'mode=' + esc(APP.mode) + ' | room=' + esc(APP.room || '-') + ' | loc=' + esc(APP.loc || '-') + ' | bridge=' + (hasBridge() ? 'yes' : 'no') + ' | rows=' + rowCount + ' | locations=' + locCount;
    }

    function saveUpdates() {
      if (!hasBridge()) {
        showNotice('Interactive save is unavailable in this context. Open the deployed /exec web app URL.', true);
        return;
      }

      const saveBtn = document.getElementById('saveBtn');
      const loading = document.getElementById('loading');
      saveBtn.disabled = true;
      loading.classList.remove('hidden');

      const byRow = {};
      let invalidQty = false;
      document.querySelectorAll('.qty-input').forEach(function (input) {
        const row = input.getAttribute('data-row');
        const qty = Number(input.value);
        if (!Number.isFinite(qty) || qty < 0) invalidQty = true;
        byRow[row] = byRow[row] || { sheetRow: Number(row) };
        byRow[row].qty = qty;
      });

      if (invalidQty) {
        saveBtn.disabled = false;
        loading.classList.add('hidden');
        showNotice('Please enter valid non-negative quantities before saving.', true);
        return;
      }

      document.querySelectorAll('.status-input').forEach(function (select) {
        const row = select.getAttribute('data-row');
        byRow[row] = byRow[row] || { sheetRow: Number(row) };
        byRow[row].status = select.value;
      });

      google.script.run
        .withSuccessHandler(function (res) {
          saveBtn.disabled = false;
          loading.classList.add('hidden');
          APP.bootstrap.rows = res.rows || [];
          document.getElementById('items').innerHTML = res.html || '';
          showNotice('Saved ' + res.updatedCount + ' item(s) successfully.', false);
        })
        .withFailureHandler(function (err) {
          saveBtn.disabled = false;
          loading.classList.add('hidden');
          showNotice((err && err.message) || 'Save failed. Please try again.', true);
        })
        .saveInventoryUpdates({
          room: APP.room,
          loc: APP.loc,
          updates: Object.keys(byRow).map(function (key) {
            return byRow[key];
          })
        });
    }

    function showNotice(message, isError) {
      if (!message) return;
      const el = document.getElementById('notice');
      el.textContent = message;
      el.className = 'rounded-2xl px-4 py-3 text-sm mt-0';
      if (isError) {
        el.classList.add('bg-red-100', 'text-red-800', 'border', 'border-red-200');
      } else {
        el.classList.add('bg-emerald-100', 'text-emerald-800', 'border', 'border-emerald-200');
      }
      el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    function esc(value) {
      return String(value == null ? '' : value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }
  </script>
</body>
</html>`;
}
