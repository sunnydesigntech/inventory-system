/*************************************
 * D&T QR Inventory System (Standalone)
 * Single-file production-ready Apps Script web app
 *************************************/

const DEBUG = false;
const DEFAULT_SHEET_NAME = 'Inventory';
const REQUIRED_FIELDS = ['itemId', 'itemName', 'room', 'location', 'qty', 'category', 'status', 'qrLink'];
const OPTIONAL_FIELDS = ['unit', 'remarks', 'locationCode'];
const STATUS_OPTIONS = ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'];
const STATUS_CANONICAL_MAP = {
  'good': 'Good',
  'low stock': 'Low Stock',
  'missing': 'Missing',
  'needs maintenance': 'Needs Maintenance'
};

/** ---------------------------
 *  Public entry points
 *  --------------------------- */

function doGet(e) {
  let state;
  try {
    state = buildPageState_(e);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      buildShellHtml_(buildFatalState_(err))
    ).setTitle('D&T Inventory').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutput(buildShellHtml_(state))
    .setTitle('D&T Inventory')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function saveInventoryUpdates(payload) {
  try {
    validateSavePayload_(payload);

    const room = safeString_(payload.room);
    const location = safeString_(payload.location);
    const updates = payload.updates;

    const sheet = getInventorySheet_();
    const colMap = getColumnMap_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      throw new Error('Save failure: inventory sheet contains no data rows.');
    }

    const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const rowDataMap = {};
    values.forEach(function(row, i) {
      rowDataMap[i + 2] = row;
    });

    updates.forEach(function(update) {
      const rowNumber = Number(update.rowNumber);
      const qtyValue = normalizeQty_(update.qty);
      const statusValue = normalizeStatus_(update.status);

      if (!Number.isInteger(rowNumber) || rowNumber < 2 || rowNumber > lastRow) {
        throw new Error('Invalid row number: ' + update.rowNumber);
      }
      if (!Number.isFinite(qtyValue) || qtyValue < 0) {
        throw new Error('Invalid quantity at row ' + rowNumber + '. Must be non-negative numeric value.');
      }
      if (!statusValue || STATUS_OPTIONS.indexOf(statusValue) === -1) {
        throw new Error('Invalid status at row ' + rowNumber + '.');
      }

      const rowValues = rowDataMap[rowNumber];
      if (!rowValues) {
        throw new Error('Row not found in sheet: ' + rowNumber);
      }

      const rowRoom = safeString_(rowValues[colMap.room - 1]);
      const rowLocation = safeString_(rowValues[colMap.location - 1]);
      if (rowRoom !== room || rowLocation !== location) {
        throw new Error('Row ' + rowNumber + ' does not belong to selected room/location.');
      }

      sheet.getRange(rowNumber, colMap.qty).setValue(qtyValue);
      sheet.getRange(rowNumber, colMap.status).setValue(statusValue);
    });

    const updatedItems = getInventoryRowsForLocation_(room, location);

    return {
      ok: true,
      message: 'Saved ' + updates.length + ' item(s) successfully.',
      updatedCount: updates.length,
      items: updatedItems,
      renderedItemsHtml: renderItemCardsHtml_(updatedItems, true, colMap)
    };
  } catch (err) {
    return {
      ok: false,
      message: err && err.message ? err.message : String(err)
    };
  }
}

function refreshQrLinks() {
  const cfg = getConfigStatus_();
  if (!cfg.spreadsheetIdConfigured) {
    throw new Error('Configuration error: SPREADSHEET_ID is not set.');
  }
  if (!cfg.webAppBaseUrlConfigured) {
    throw new Error('Configuration error: WEB_APP_BASE_URL is not set.');
  }

  const baseUrl = getWebAppBaseUrl_();
  const sheet = getInventorySheet_();
  const colMap = getColumnMap_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { ok: true, updatedRows: 0, message: 'No data rows to update.' };
  }

  const width = Math.max(colMap.room, colMap.location, colMap.qrLink);
  const rows = sheet.getRange(2, 1, lastRow - 1, width).getValues();

  const linkValues = rows.map(function(row) {
    const room = safeString_(row[colMap.room - 1]);
    const loc = safeString_(row[colMap.location - 1]);
    if (!room || !loc) return [''];
    return [baseUrl + '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc)];
  });

  sheet.getRange(2, colMap.qrLink, linkValues.length, 1).setValues(linkValues);
  return {
    ok: true,
    updatedRows: linkValues.length,
    message: 'QR links refreshed for ' + linkValues.length + ' rows.'
  };
}

function setAppConfig(spreadsheetId, webAppBaseUrl, inventorySheetName) {
  if (!spreadsheetId || !webAppBaseUrl) {
    throw new Error('Both spreadsheetId and webAppBaseUrl are required.');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', safeString_(spreadsheetId));
  props.setProperty('WEB_APP_BASE_URL', safeString_(webAppBaseUrl));

  if (inventorySheetName) {
    props.setProperty('INVENTORY_SHEET_NAME', safeString_(inventorySheetName));
  } else {
    props.deleteProperty('INVENTORY_SHEET_NAME');
  }

  return {
    ok: true,
    config: getConfigStatus_()
  };
}

/** Optional admin utility for script editor. */
function getConfigStatus() {
  return getConfigStatus_();
}

/** Optional menu helper for admins (not required by web app flow). */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Admin')
    .addItem('Refresh QR Links', 'refreshQrLinks')
    .addItem('Show Config Status (Logs)', 'logConfigStatus_')
    .addToUi();
}

function logConfigStatus_() {
  Logger.log(JSON.stringify(getConfigStatus_(), null, 2));
}

/** ---------------------------
 *  Config helpers
 *  --------------------------- */

function getConfigStatus_() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = safeString_(props.getProperty('SPREADSHEET_ID'));
  const webAppBaseUrl = safeString_(props.getProperty('WEB_APP_BASE_URL'));
  const inventorySheetName = safeString_(props.getProperty('INVENTORY_SHEET_NAME'));

  return {
    spreadsheetIdConfigured: !!spreadsheetId,
    webAppBaseUrlConfigured: !!webAppBaseUrl,
    inventorySheetNameConfigured: !!inventorySheetName,
    spreadsheetId: spreadsheetId,
    webAppBaseUrl: webAppBaseUrl,
    inventorySheetName: inventorySheetName
  };
}

function getSpreadsheet_() {
  const spreadsheetId = safeString_(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
  if (!spreadsheetId) {
    throw new Error('Configuration error: SPREADSHEET_ID is not set.');
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

function getInventorySheet_() {
  const ss = getSpreadsheet_();
  const preferredName = safeString_(PropertiesService.getScriptProperties().getProperty('INVENTORY_SHEET_NAME'));

  let sheet = null;
  if (preferredName) {
    sheet = ss.getSheetByName(preferredName);
  }
  if (!sheet) {
    sheet = ss.getSheetByName(DEFAULT_SHEET_NAME);
  }
  if (!sheet) {
    const sheets = ss.getSheets();
    if (!sheets || !sheets.length) {
      throw new Error('Configuration error: no sheets found in spreadsheet.');
    }
    sheet = sheets[0];
  }
  return sheet;
}

function getWebAppBaseUrl_() {
  const raw = safeString_(PropertiesService.getScriptProperties().getProperty('WEB_APP_BASE_URL'));
  if (!raw) {
    throw new Error('Configuration error: WEB_APP_BASE_URL is not set.');
  }
  return raw.replace(/\/+$/, '');
}

/** ---------------------------
 *  Data helpers
 *  --------------------------- */

function getColumnMap_(sheet) {
  const targetSheet = sheet || getInventorySheet_();
  const lastCol = targetSheet.getLastColumn();
  if (lastCol < 1) {
    throw new Error('Configuration error: header row is missing.');
  }

  const headers = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);

  const aliases = {
    itemId: ['item id', 'itemid', 'id'],
    itemName: ['item name', 'name', 'item'],
    room: ['room'],
    location: ['specific location', 'location', 'loc'],
    qty: ['qty', 'quantity', 'stock', 'count'],
    category: ['category', 'type'],
    status: ['status', 'condition'],
    qrLink: ['qr code link (auto-generated)', 'qr code link', 'qr link', 'qrcode link'],
    unit: ['unit'],
    remarks: ['remarks', 'remark', 'notes', 'note'],
    locationCode: ['location code', 'loc code', 'locationcode']
  };

  const map = {};
  Object.keys(aliases).forEach(function(field) {
    map[field] = findHeaderIndex_(headers, aliases[field]);
  });

  const missing = REQUIRED_FIELDS.filter(function(field) {
    return map[field] <= 0;
  });
  if (missing.length) {
    throw new Error('Configuration error: missing required columns: ' + missing.join(', '));
  }

  OPTIONAL_FIELDS.forEach(function(field) {
    if (!map[field] || map[field] < 0) map[field] = -1;
  });

  return map;
}

function getAllLocations_() {
  const sheet = getInventorySheet_();
  const colMap = getColumnMap_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const width = Math.max(colMap.room, colMap.location);
  const rows = sheet.getRange(2, 1, lastRow - 1, width).getValues();

  const seen = {};
  const out = [];
  const baseUrl = getSafeBaseUrl_();

  rows.forEach(function(row) {
    const room = safeString_(row[colMap.room - 1]);
    const location = safeString_(row[colMap.location - 1]);
    if (!room || !location) return;

    const key = room + '||' + location;
    if (seen[key]) return;

    seen[key] = true;
    out.push({
      room: room,
      location: location,
      viewUrl: baseUrl ? buildLocationUrl_(room, location, false, baseUrl) : '',
      techUrl: baseUrl ? buildLocationUrl_(room, location, true, baseUrl) : ''
    });
  });

  out.sort(function(a, b) {
    if (a.room === b.room) return a.location.localeCompare(b.location);
    return a.room.localeCompare(b.room);
  });

  return out;
}

function getInventoryRowsForLocation_(room, loc) {
  const roomValue = safeString_(room);
  const locValue = safeString_(loc);
  const sheet = getInventorySheet_();
  const colMap = getColumnMap_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const items = [];

  rows.forEach(function(row, idx) {
    const rowRoom = safeString_(row[colMap.room - 1]);
    const rowLoc = safeString_(row[colMap.location - 1]);
    if (rowRoom !== roomValue || rowLoc !== locValue) return;

    items.push({
      rowNumber: idx + 2,
      itemId: safeString_(row[colMap.itemId - 1]),
      itemName: safeString_(row[colMap.itemName - 1]),
      room: rowRoom,
      location: rowLoc,
      qty: safeString_(row[colMap.qty - 1]),
      category: safeString_(row[colMap.category - 1]),
      status: safeString_(row[colMap.status - 1]),
      unit: getOptionalValue_(row, colMap.unit),
      remarks: getOptionalValue_(row, colMap.remarks),
      locationCode: getOptionalValue_(row, colMap.locationCode)
    });
  });

  return items;
}

/** ---------------------------
 *  Route/page state helpers
 *  --------------------------- */

function buildPageState_(e) {
  const params = (e && e.parameter) || {};
  const room = safeString_(params.room);
  const location = safeString_(params.loc);
  const mode = safeString_(params.mode).toLowerCase() === 'tech' ? 'tech' : 'view';
  const isLocationRoute = !!(room && location);

  const cfg = getConfigStatus_();
  const notices = [];

  if (!cfg.spreadsheetIdConfigured) {
    notices.push({ type: 'error', text: 'Configuration error: SPREADSHEET_ID is not set.' });
    return buildBaseState_({
      mode: 'landing',
      room: '',
      location: '',
      items: [],
      locations: [],
      notices: notices,
      config: cfg,
      columns: null
    });
  }

  if (!cfg.webAppBaseUrlConfigured) {
    notices.push({
      type: 'warn',
      text: 'WEB_APP_BASE_URL is not configured. Location links are disabled until this is set.'
    });
  }

  let columns = null;
  let locations = [];
  let items = [];

  try {
    const sheet = getInventorySheet_();
    columns = getColumnMap_(sheet);
    locations = getAllLocations_();

    if (sheet.getLastRow() < 2) {
      notices.push({ type: 'warn', text: 'Inventory sheet is currently empty.' });
    }

    if (isLocationRoute) {
      items = getInventoryRowsForLocation_(room, location);
      if (!items.length) {
        notices.push({ type: 'warn', text: 'No inventory records found for this location.' });
      }
    } else if (!locations.length) {
      notices.push({ type: 'info', text: 'No inventory locations found yet.' });
    }
  } catch (err) {
    notices.push({ type: 'error', text: err && err.message ? err.message : String(err) });
    return buildBaseState_({
      mode: isLocationRoute ? mode : 'landing',
      room: room,
      location: location,
      items: [],
      locations: [],
      notices: notices,
      config: cfg,
      columns: columns
    });
  }

  return buildBaseState_({
    mode: isLocationRoute ? mode : 'landing',
    room: room,
    location: location,
    items: items,
    locations: locations,
    notices: notices,
    config: cfg,
    columns: columns
  });
}

function buildBaseState_(input) {
  const baseUrl = getSafeBaseUrl_();
  return {
    appTitle: 'D&T QR Inventory System',
    mode: input.mode,
    room: input.room,
    location: input.location,
    isLanding: input.mode === 'landing',
    isTechMode: input.mode === 'tech',
    items: input.items || [],
    locations: input.locations || [],
    locationsByRoom: groupLocationsByRoom_(input.locations || []),
    notices: input.notices || [],
    statusOptions: STATUS_OPTIONS,
    config: input.config,
    columns: input.columns,
    baseUrl: baseUrl,
    hasBaseUrl: !!baseUrl,
    debug: DEBUG
  };
}

function buildFatalState_(err) {
  const msg = (err && err.message) ? err.message : String(err);
  return buildBaseState_({
    mode: 'landing',
    room: '',
    location: '',
    items: [],
    locations: [],
    notices: [{ type: 'error', text: 'Configuration error: ' + msg }],
    config: getConfigStatus_(),
    columns: null
  });
}

/** ---------------------------
 *  Rendering helpers (server)
 *  --------------------------- */

function buildShellHtml_(state) {
  const topControls = renderTopControls_(state);
  const noticesHtml = renderNotices_(state.notices);
  const mainContent = state.isLanding ? renderLanding_(state) : renderLocationPage_(state);
  const debugPanel = state.debug ? renderDebugPanel_(state) : '';

  const clientData = {
    mode: state.mode,
    room: state.room,
    location: state.location,
    isTechMode: state.isTechMode,
    statusOptions: state.statusOptions,
    columns: state.columns,
    debug: state.debug,
    hasBaseUrl: state.hasBaseUrl,
    initialItemCount: state.items.length,
    initialLocationCount: state.locations.length
  };

  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
    '  <meta charset="utf-8">',
    '  <meta name="viewport" content="width=device-width,initial-scale=1">',
    '  <title>D&T Inventory</title>',
    '  <style>', css_(), '</style>',
    '</head>',
    '<body>',
    '  <main class="container">',
    '    <header class="header">',
    '      <h1>' + escapeHtml_(state.appTitle) + '</h1>',
    '      <p class="sub">' + escapeHtml_(state.isLanding
          ? 'Browse inventory by room/location or scan a QR code.'
          : ('Room ' + state.room + ' · Location ' + state.location)) + '</p>',
    '    </header>',
    topControls,
    noticesHtml,
    mainContent,
    debugPanel,
    '  </main>',
    '  <script>window.__APP_DATA__=' + JSON.stringify(clientData) + ';</script>',
    '  <script>' + clientScript_() + '</script>',
    '</body>',
    '</html>'
  ].join('');
}

function renderTopControls_(state) {
  if (state.isLanding) {
    return [
      '<div class="top-actions disabled-state">',
      '  <button class="btn secondary" disabled title="Choose a location card first">View Mode</button>',
      '  <button class="btn" disabled title="Choose a location card first">Technician Access</button>',
      '</div>'
    ].join('');
  }

  const viewUrl = state.hasBaseUrl ? buildLocationUrl_(state.room, state.location, false, state.baseUrl) : '#';
  const techUrl = state.hasBaseUrl ? buildLocationUrl_(state.room, state.location, true, state.baseUrl) : '#';
  const homeUrl = state.hasBaseUrl ? state.baseUrl : '#';

  return [
    '<div class="top-actions">',
    '  <a class="btn ' + (state.isTechMode ? 'secondary' : '') + '" href="' + escapeHtml_(viewUrl) + '">View Mode</a>',
    '  <a class="btn ' + (state.isTechMode ? '' : 'secondary') + '" href="' + escapeHtml_(techUrl) + '">Technician Access</a>',
    '  <a class="btn ghost" href="' + escapeHtml_(homeUrl) + '">All Locations</a>',
    '</div>'
  ].join('');
}

function renderNotices_(notices) {
  if (!notices || !notices.length) return '';
  return '<section class="notices">' + notices.map(function(n) {
    return '<div class="notice ' + escapeHtml_(n.type || 'info') + '">' + escapeHtml_(n.text || '') + '</div>';
  }).join('') + '</section>';
}

function renderLanding_(state) {
  const rooms = Object.keys(state.locationsByRoom || {}).sort(function(a, b) {
    return a.localeCompare(b);
  });

  if (!rooms.length) {
    return '<section class="empty">No inventory locations found yet.</section>';
  }

  const groupsHtml = rooms.map(function(room) {
    const cardsHtml = state.locationsByRoom[room].map(function(loc) {
      const viewBtn = state.hasBaseUrl
        ? '<a class="btn secondary" href="' + escapeHtml_(loc.viewUrl) + '">Open View</a>'
        : '<button class="btn secondary" disabled>Open View</button>';
      const techBtn = state.hasBaseUrl
        ? '<a class="btn" href="' + escapeHtml_(loc.techUrl) + '">Open Tech</a>'
        : '<button class="btn" disabled>Open Tech</button>';

      return [
        '<article class="card">',
        '  <h3>' + escapeHtml_(loc.location) + '</h3>',
        '  <p class="muted">Room ' + escapeHtml_(loc.room) + '</p>',
        '  <div class="row-actions">',
        viewBtn,
        techBtn,
        '  </div>',
        '</article>'
      ].join('');
    }).join('');

    return [
      '<section class="room-group">',
      '  <h2>Room ' + escapeHtml_(room) + '</h2>',
      '  <div class="card-grid">', cardsHtml, '  </div>',
      '</section>'
    ].join('');
  }).join('');

  return '<section class="landing">' + groupsHtml + '</section>';
}

function renderLocationPage_(state) {
  const itemCardsHtml = renderItemCardsHtml_(state.items, state.isTechMode, state.columns);

  const saveSection = state.isTechMode
    ? [
        '<div class="save-panel">',
        '  <button id="saveBtn" class="btn">Save Changes</button>',
        '  <div id="saveMsg" class="save-msg"></div>',
        '</div>',
        '<p id="bridgeWarn" class="bridge-warning" style="display:none">',
        'Interactive save is unavailable in this context. Open the deployed /exec web app URL.',
        '</p>'
      ].join('')
    : '';

  return [
    '<section class="location-view">',
    '  <div id="itemList" class="item-list">', itemCardsHtml, '</div>',
    saveSection,
    '</section>'
  ].join('');
}

function renderItemCardsHtml_(items, isTechMode, colMap) {
  if (!items || !items.length) {
    return '<div class="empty">No inventory records found for this location.</div>';
  }

  const hasUnit = !!(colMap && colMap.unit > 0);
  const hasRemarks = !!(colMap && colMap.remarks > 0);
  const hasLocationCode = !!(colMap && colMap.locationCode > 0);

  return items.map(function(item) {
    const isChemical = safeString_(item.category).toLowerCase() === 'chemicals';
    const normalizedStatus = normalizeStatus_(item.status) || safeString_(item.status);

    const qtyField = isTechMode
      ? '<input class="qty-input" type="number" min="0" step="1" data-field="qty" data-row="' + item.rowNumber + '" value="' + escapeHtml_(item.qty) + '">'
      : '<span class="value">' + escapeHtml_(item.qty) + '</span>';

    const statusField = isTechMode
      ? ('<select class="status-select" data-field="status" data-row="' + item.rowNumber + '">' +
          STATUS_OPTIONS.map(function(opt) {
            const selected = (opt === normalizedStatus) ? ' selected' : '';
            return '<option value="' + escapeHtml_(opt) + '"' + selected + '>' + escapeHtml_(opt) + '</option>';
          }).join('') +
         '</select>')
      : '<span class="status-pill ' + statusClass_(normalizedStatus) + '">' + escapeHtml_(normalizedStatus || item.status) + '</span>';

    const extraRows = [
      (hasUnit && item.unit) ? '<div><span class="k">Unit</span><span class="v">' + escapeHtml_(item.unit) + '</span></div>' : '',
      (hasLocationCode && item.locationCode) ? '<div><span class="k">Location Code</span><span class="v">' + escapeHtml_(item.locationCode) + '</span></div>' : '',
      (hasRemarks && item.remarks) ? '<div class="remarks"><span class="k">Remarks</span><span class="v">' + escapeHtml_(item.remarks) + '</span></div>' : ''
    ].join('');

    return [
      '<article class="item-card ' + (isChemical ? 'hazard' : '') + '">',
      '  <div class="item-head">',
      '    <h3>' + escapeHtml_(item.itemName || item.itemId) + '</h3>',
      isChemical ? '    <span class="hazard-badge">Hazard</span>' : '',
      '  </div>',
      isChemical ? '  <p class="hazard-note">Chemical item — handle/store according to safety procedure.</p>' : '',
      '  <p class="muted">' + escapeHtml_(item.itemId) + ' · ' + escapeHtml_(item.category || 'Uncategorized') + '</p>',
      '  <div class="item-grid">',
      '    <label>Qty</label>', qtyField,
      '    <label>Status</label>', statusField,
      '  </div>',
      '  <div class="meta-grid">',
      extraRows,
      '  </div>',
      '</article>'
    ].join('');
  }).join('');
}

function renderDebugPanel_(state) {
  return [
    '<aside class="debug-panel">',
    '  <h4>Debug</h4>',
    '  <ul>',
    '    <li>mode: ' + escapeHtml_(state.mode) + '</li>',
    '    <li>room: ' + escapeHtml_(state.room || '-') + '</li>',
    '    <li>location: ' + escapeHtml_(state.location || '-') + '</li>',
    '    <li>items loaded: ' + state.items.length + '</li>',
    '    <li>locations loaded: ' + state.locations.length + '</li>',
    '    <li>web app URL configured: ' + (state.config.webAppBaseUrlConfigured ? 'true' : 'false') + '</li>',
    '    <li>google.script.run exists: <span id="dbgBridge">(checking...)</span></li>',
    '  </ul>',
    '</aside>'
  ].join('');
}

/** ---------------------------
 *  Client helpers
 *  --------------------------- */

function clientScript_() {
  return [
    '(function(){',
    '  var app = window.__APP_DATA__ || {};',
    '  var hasBridge = !!(window.google && google.script && google.script.run);',
    '  var dbgBridge = document.getElementById("dbgBridge");',
    '  if (dbgBridge) dbgBridge.textContent = hasBridge ? "true" : "false";',
    '  if (!app.isTechMode) return;',
    '  var saveBtn = document.getElementById("saveBtn");',
    '  var saveMsg = document.getElementById("saveMsg");',
    '  var itemList = document.getElementById("itemList");',
    '  var warn = document.getElementById("bridgeWarn");',
    '  if (!hasBridge) {',
    '    if (warn) warn.style.display = "block";',
    '    if (saveBtn) saveBtn.disabled = true;',
    '    return;',
    '  }',
    '  if (!saveBtn) return;',
    '  saveBtn.addEventListener("click", function(){',
    '    var qtyEls = Array.prototype.slice.call(document.querySelectorAll("input[data-field=qty]"));',
    '    var statusEls = Array.prototype.slice.call(document.querySelectorAll("select[data-field=status]"));',
    '    var byRow = {};',
    '    qtyEls.forEach(function(el){',
    '      var r = el.getAttribute("data-row");',
    '      byRow[r] = byRow[r] || { rowNumber: Number(r) };',
    '      byRow[r].qty = el.value;',
    '    });',
    '    statusEls.forEach(function(el){',
    '      var r = el.getAttribute("data-row");',
    '      byRow[r] = byRow[r] || { rowNumber: Number(r) };',
    '      byRow[r].status = el.value;',
    '    });',
    '    var updates = Object.keys(byRow).map(function(k){ return byRow[k]; });',
    '    if (!updates.length) {',
    '      saveMsg.textContent = "No updates to save.";',
    '      saveMsg.className = "save-msg";',
    '      return;',
    '    }',
    '    saveBtn.disabled = true;',
    '    saveMsg.textContent = "Saving...";',
    '    saveMsg.className = "save-msg";',
    '    google.script.run',
    '      .withSuccessHandler(function(res){',
    '        saveBtn.disabled = false;',
    '        if (!res || !res.ok) {',
    '          saveMsg.textContent = (res && res.message) ? res.message : "Save failed.";',
    '          saveMsg.className = "save-msg err";',
    '          return;',
    '        }',
    '        saveMsg.textContent = res.message || "Saved.";',
    '        saveMsg.className = "save-msg ok";',
    '        if (itemList && res.renderedItemsHtml) {',
    '          itemList.innerHTML = res.renderedItemsHtml;',
    '        }',
    '      })',
    '      .withFailureHandler(function(err){',
    '        saveBtn.disabled = false;',
    '        saveMsg.textContent = (err && err.message) ? err.message : "Save failed.";',
    '        saveMsg.className = "save-msg err";',
    '      })',
    '      .saveInventoryUpdates({',
    '        room: app.room,',
    '        location: app.location,',
    '        updates: updates',
    '      });',
    '  });',
    '})();'
  ].join('');
}

function css_() {
  return [
    ':root{--bg:#f8fafc;--surface:#ffffff;--text:#0f172a;--muted:#475569;--line:#e2e8f0;--brand:#1d4ed8;--good:#166534;--low:#c2410c;--missing:#b91c1c;--maint:#ea580c;}',
    '*{box-sizing:border-box;}',
    'body{margin:0;background:var(--bg);color:var(--text);font:16px/1.45 Arial,sans-serif;}',
    '.container{max-width:980px;margin:0 auto;padding:14px 14px 28px;}',
    '.header h1{margin:0;font-size:1.45rem;}',
    '.header .sub{margin:.4rem 0 0;color:var(--muted);}',
    '.top-actions{display:flex;gap:8px;flex-wrap:wrap;margin:14px 0;}',
    '.disabled-state .btn{opacity:.6;cursor:not-allowed;}',
    '.btn{display:inline-block;padding:10px 14px;border-radius:10px;background:var(--brand);color:#fff;border:0;text-decoration:none;font-weight:700;font-size:.95rem;}',
    '.btn.secondary{background:#334155;}',
    '.btn.ghost{background:#e2e8f0;color:#0f172a;}',
    '.btn[disabled]{pointer-events:none;}',
    '.notices{display:grid;gap:8px;margin:10px 0 16px;}',
    '.notice{border-radius:10px;padding:10px 12px;border:1px solid var(--line);background:#fff;}',
    '.notice.info{border-color:#bfdbfe;background:#eff6ff;color:#1e3a8a;}',
    '.notice.warn{border-color:#fed7aa;background:#fff7ed;color:#9a3412;}',
    '.notice.error{border-color:#fecaca;background:#fef2f2;color:#991b1b;}',
    '.room-group{margin:18px 0;}',
    '.room-group h2{margin:0 0 10px;font-size:1.05rem;}',
    '.card-grid{display:grid;grid-template-columns:1fr;gap:10px;}',
    '.card,.item-card,.empty{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:12px;}',
    '.card h3,.item-card h3{margin:0;font-size:1.04rem;}',
    '.muted{margin:.35rem 0;color:var(--muted);}',
    '.row-actions{display:flex;gap:8px;flex-wrap:wrap;}',
    '.item-list{display:grid;gap:10px;}',
    '.item-card.hazard{border-color:#f59e0b;background:#fffbeb;}',
    '.item-head{display:flex;justify-content:space-between;align-items:center;gap:8px;}',
    '.hazard-badge{padding:4px 8px;border-radius:999px;background:#b91c1c;color:#fff;font-size:.72rem;font-weight:700;}',
    '.hazard-note{margin:.45rem 0;color:#7c2d12;font-size:.88rem;}',
    '.item-grid{display:grid;grid-template-columns:90px 1fr;gap:8px 10px;align-items:center;margin:.45rem 0;}',
    '.qty-input,.status-select{width:100%;padding:8px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;font-size:1rem;}',
    '.value{font-weight:700;}',
    '.status-pill{display:inline-block;padding:5px 9px;border-radius:999px;font-size:.84rem;font-weight:700;}',
    '.status-good{background:#dcfce7;color:var(--good);}',
    '.status-low{background:#ffedd5;color:var(--low);}',
    '.status-missing{background:#fee2e2;color:var(--missing);}',
    '.status-maint{background:#ffedd5;color:var(--maint);}',
    '.status-default{background:#e2e8f0;color:#334155;}',
    '.meta-grid{display:grid;grid-template-columns:1fr;gap:6px;margin-top:6px;}',
    '.meta-grid .k{display:inline-block;min-width:105px;color:var(--muted);font-size:.9rem;}',
    '.meta-grid .v{font-size:.95rem;}',
    '.remarks{padding-top:2px;border-top:1px dashed #dbe2ea;}',
    '.save-panel{position:sticky;bottom:0;background:rgba(248,250,252,.95);backdrop-filter:blur(2px);padding-top:10px;margin-top:10px;display:flex;gap:10px;align-items:center;}',
    '.save-msg{font-weight:700;color:#334155;}',
    '.save-msg.ok{color:#166534;}',
    '.save-msg.err{color:#b91c1c;}',
    '.bridge-warning{margin-top:10px;background:#fff7ed;color:#9a3412;border:1px solid #fed7aa;border-radius:10px;padding:10px;}',
    '.debug-panel{margin-top:14px;background:#0f172a;color:#e2e8f0;border-radius:10px;padding:10px;font-family:monospace;font-size:.83rem;}',
    '.debug-panel h4{margin:0 0 6px;}',
    '.debug-panel ul{margin:0;padding-left:18px;}',
    '@media (min-width:760px){.container{padding:20px 20px 32px;}.header h1{font-size:1.8rem;}.card-grid{grid-template-columns:repeat(2,minmax(0,1fr));}}'
  ].join('');
}

/** ---------------------------
 *  Utility helpers
 *  --------------------------- */

function groupLocationsByRoom_(locations) {
  const grouped = {};
  (locations || []).forEach(function(loc) {
    if (!grouped[loc.room]) grouped[loc.room] = [];
    grouped[loc.room].push(loc);
  });
  return grouped;
}

function buildLocationUrl_(room, location, techMode, baseUrl) {
  const root = (baseUrl || getWebAppBaseUrl_()).replace(/\/+$/, '');
  let url = root + '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(location);
  if (techMode) url += '&mode=tech';
  return url;
}

function getSafeBaseUrl_() {
  try {
    return getWebAppBaseUrl_();
  } catch (err) {
    return '';
  }
}

function validateSavePayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid save payload.');
  }
  if (!safeString_(payload.room) || !safeString_(payload.location)) {
    throw new Error('Invalid save payload: room and location are required.');
  }
  if (!Array.isArray(payload.updates) || payload.updates.length === 0) {
    throw new Error('Invalid save payload: updates must be a non-empty array.');
  }
}

function normalizeHeader_(value) {
  return safeString_(value)
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function findHeaderIndex_(normalizedHeaders, allowedAliases) {
  for (let i = 0; i < normalizedHeaders.length; i++) {
    if (allowedAliases.indexOf(normalizedHeaders[i]) !== -1) {
      return i + 1;
    }
  }
  return -1;
}

function normalizeStatus_(status) {
  const key = safeString_(status).toLowerCase();
  return STATUS_CANONICAL_MAP[key] || '';
}

function normalizeQty_(qty) {
  const n = Number(qty);
  return Number.isFinite(n) ? n : NaN;
}

function statusClass_(status) {
  const normalized = normalizeStatus_(status) || safeString_(status).toLowerCase();
  if (normalized === 'good') return 'status-good';
  if (normalized === 'low stock') return 'status-low';
  if (normalized === 'missing') return 'status-missing';
  if (normalized === 'needs maintenance') return 'status-maint';
  return 'status-default';
}

function getOptionalValue_(row, colIndex) {
  if (!colIndex || colIndex < 1) return '';
  return safeString_(row[colIndex - 1]);
}

function safeString_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function escapeHtml_(value) {
  return safeString_(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
