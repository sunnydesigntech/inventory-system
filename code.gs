const DEBUG = false;

const APP = Object.freeze({
  title: 'D&T QR Inventory System',
  defaultInventorySheetName: 'Inventory',
  statuses: ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'],
  propertyKeys: {
    spreadsheetId: 'SPREADSHEET_ID',
    webAppBaseUrl: 'WEB_APP_BASE_URL',
    inventorySheetName: 'INVENTORY_SHEET_NAME'
  },
  headerAliases: {
    itemId: ['item id', 'itemid', 'id'],
    itemName: ['item name', 'itemname', 'name'],
    room: ['room'],
    location: ['specific location', 'location', 'specificlocation'],
    qty: ['qty', 'quantity'],
    category: ['category'],
    status: ['status'],
    qrLink: ['qr code link (auto-generated)', 'qr code link', 'qr link', 'qr url']
  }
});

function doGet(e) {
  try {
    const params = parseParams_(e);
    const page = buildPageState_(params);
    return HtmlService.createHtmlOutput(renderAppHtml_(page))
      .setTitle(APP.title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(renderConfigErrorHtml_(err))
      .setTitle(APP.title);
  }
}

function saveInventoryUpdates(payload) {
  try {
    validateSavePayloadShape_(payload);

    const room = cleanString_(payload.room);
    const location = cleanString_(payload.location);
    if (!room || !location) {
      throw new Error('Missing room/location context for save.');
    }

    const sheet = getInventorySheet_();
    const columnMap = getColumnMap_(sheet);
    const dataRows = getSheetDataRows_(sheet, columnMap);

    if (!dataRows.length) {
      throw new Error('The inventory sheet has no data rows to update.');
    }

    const lookup = buildRowLookup_(dataRows, columnMap);
    const allowedStatuses = APP.statuses.reduce(function (acc, value) {
      acc[value] = true;
      return acc;
    }, {});

    const updates = payload.items.map(function (entry) {
      const rowNumber = Number(entry.rowNumber);
      if (!rowNumber || !lookup[rowNumber]) {
        throw new Error('Invalid row number: ' + entry.rowNumber);
      }

      const source = lookup[rowNumber];
      if (source.room !== room || source.location !== location) {
        throw new Error('Row ' + rowNumber + ' is outside selected room/location.');
      }

      const qty = Number(entry.qty);
      if (!isFinite(qty) || qty < 0) {
        throw new Error('Invalid quantity for row ' + rowNumber + '. Quantity must be numeric and non-negative.');
      }

      const status = cleanString_(entry.status);
      if (!allowedStatuses[status]) {
        throw new Error('Invalid status for row ' + rowNumber + ': ' + status);
      }

      return { rowNumber: rowNumber, qty: qty, status: status };
    });

    updates.forEach(function (u) {
      sheet.getRange(u.rowNumber, columnMap.qty).setValue(u.qty);
      sheet.getRange(u.rowNumber, columnMap.status).setValue(u.status);
    });

    return {
      ok: true,
      message: 'Inventory updated successfully.',
      room: room,
      location: location,
      items: getInventoryRowsForLocation_(room, location)
    };
  } catch (err) {
    return {
      ok: false,
      message: err && err.message ? err.message : String(err)
    };
  }
}

function refreshQrLinks() {
  const baseUrl = getWebAppBaseUrl_();
  const sheet = getInventorySheet_();
  const columnMap = getColumnMap_(sheet);
  const dataRows = getSheetDataRows_(sheet, columnMap);

  if (!dataRows.length) {
    return { ok: true, updatedRows: 0, message: 'No data rows found.' };
  }

  const output = dataRows.map(function (row) {
    const room = cleanString_(row[columnMap.room - 1]);
    const location = cleanString_(row[columnMap.location - 1]);
    if (!room || !location) return [''];
    return [createAppUrl_(baseUrl, room, location, '')];
  });

  sheet.getRange(2, columnMap.qrLink, output.length, 1).setValues(output);

  const updatedRows = output.filter(function (row) { return !!row[0]; }).length;
  return {
    ok: true,
    updatedRows: updatedRows,
    message: 'QR links refreshed for ' + updatedRows + ' row(s).'
  };
}

function setAppConfig(spreadsheetId, webAppBaseUrl, inventorySheetName) {
  const id = cleanString_(spreadsheetId);
  const url = cleanString_(webAppBaseUrl);

  if (!id) {
    throw new Error('spreadsheetId is required.');
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty(APP.propertyKeys.spreadsheetId, id);

  if (url) {
    props.setProperty(APP.propertyKeys.webAppBaseUrl, stripQueryString_(url));
  }

  if (cleanString_(inventorySheetName)) {
    props.setProperty(APP.propertyKeys.inventorySheetName, cleanString_(inventorySheetName));
  }

  return {
    ok: true,
    message: 'Script Properties updated.',
    spreadsheetId: id,
    webAppBaseUrl: url ? stripQueryString_(url) : '',
    inventorySheetName: cleanString_(inventorySheetName) || APP.defaultInventorySheetName
  };
}

function getSpreadsheet_() {
  const spreadsheetId = cleanString_(PropertiesService.getScriptProperties().getProperty(APP.propertyKeys.spreadsheetId));
  if (!spreadsheetId) {
    throw new Error('Configuration error: SPREADSHEET_ID is not set in Script Properties.');
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

function getInventorySheet_() {
  const ss = getSpreadsheet_();
  const configuredName = cleanString_(PropertiesService.getScriptProperties().getProperty(APP.propertyKeys.inventorySheetName));
  const preferredName = configuredName || APP.defaultInventorySheetName;
  return ss.getSheetByName(preferredName) || ss.getSheets()[0];
}

function getColumnMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    throw new Error('Configuration error: inventory sheet is missing a header row.');
  }

  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);
  const map = {};

  Object.keys(APP.headerAliases).forEach(function (key) {
    const aliases = APP.headerAliases[key];
    const index = headerRow.findIndex(function (value) {
      return aliases.indexOf(value) !== -1;
    });
    if (index === -1) {
      throw new Error('Configuration error: missing required column for ' + key + '.');
    }
    map[key] = index + 1;
  });

  return map;
}

function getAllLocations_() {
  const sheet = getInventorySheet_();
  const columnMap = getColumnMap_(sheet);
  const dataRows = getSheetDataRows_(sheet, columnMap);
  const baseUrl = safeWebAppUrlForLinks_();

  const dedupe = {};
  dataRows.forEach(function (row) {
    const room = cleanString_(row[columnMap.room - 1]);
    const location = cleanString_(row[columnMap.location - 1]);
    if (!room || !location) return;

    const key = room + '||' + location;
    if (dedupe[key]) return;

    dedupe[key] = {
      room: room,
      location: location,
      viewUrl: createAppUrl_(baseUrl, room, location, ''),
      techUrl: createAppUrl_(baseUrl, room, location, 'tech')
    };
  });

  return Object.keys(dedupe)
    .map(function (key) { return dedupe[key]; })
    .sort(function (a, b) {
      return (a.room + '|' + a.location).localeCompare(b.room + '|' + b.location);
    });
}

function getInventoryRowsForLocation_(room, loc) {
  const selectedRoom = cleanString_(room);
  const selectedLocation = cleanString_(loc);
  if (!selectedRoom || !selectedLocation) return [];

  const sheet = getInventorySheet_();
  const columnMap = getColumnMap_(sheet);
  const dataRows = getSheetDataRows_(sheet, columnMap);

  return dataRows
    .map(function (row, index) {
      return {
        rowNumber: index + 2,
        itemId: row[columnMap.itemId - 1],
        itemName: row[columnMap.itemName - 1],
        room: cleanString_(row[columnMap.room - 1]),
        location: cleanString_(row[columnMap.location - 1]),
        qty: row[columnMap.qty - 1],
        category: cleanString_(row[columnMap.category - 1]),
        status: cleanString_(row[columnMap.status - 1])
      };
    })
    .filter(function (item) {
      return item.room === selectedRoom && item.location === selectedLocation;
    })
    .sort(function (a, b) {
      return String(a.itemName).localeCompare(String(b.itemName));
    });
}

function getWebAppBaseUrl_() {
  const configured = cleanString_(PropertiesService.getScriptProperties().getProperty(APP.propertyKeys.webAppBaseUrl));
  if (configured) return stripQueryString_(configured);

  const serviceUrl = cleanString_(ScriptApp.getService().getUrl());
  if (serviceUrl) return stripQueryString_(serviceUrl);

  throw new Error('Configuration error: WEB_APP_BASE_URL is not set in Script Properties.');
}

function parseParams_(e) {
  const p = (e && e.parameter) || {};
  return {
    room: cleanString_(p.room),
    location: cleanString_(p.loc),
    mode: cleanString_(p.mode).toLowerCase() === 'tech' ? 'tech' : 'view'
  };
}

function buildPageState_(params) {
  const hasSelection = !!(params.room && params.location);
  const page = {
    title: APP.title,
    mode: params.mode,
    room: params.room,
    location: params.location,
    hasSelection: hasSelection,
    notice: '',
    error: '',
    items: [],
    locations: [],
    locationsByRoom: [],
    appUrlConfigured: true
  };

  try {
    getWebAppBaseUrl_();
  } catch (urlErr) {
    page.appUrlConfigured = false;
  }

  if (hasSelection) {
    page.items = getInventoryRowsForLocation_(params.room, params.location);
    if (!page.items.length) {
      page.notice = 'No inventory records found for this location.';
    }
  } else {
    page.locations = getAllLocations_();
    page.locationsByRoom = groupLocationsByRoom_(page.locations);
    if (!page.locations.length) {
      page.notice = 'No locations are available yet. Add inventory rows to the sheet first.';
    }
  }

  return page;
}

function getSheetDataRows_(sheet, columnMap) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const width = Math.max.apply(null, Object.keys(columnMap).map(function (k) { return columnMap[k]; }));
  return sheet.getRange(2, 1, lastRow - 1, width).getValues();
}

function buildRowLookup_(rows, columnMap) {
  return rows.reduce(function (acc, row, index) {
    const rowNumber = index + 2;
    acc[rowNumber] = {
      room: cleanString_(row[columnMap.room - 1]),
      location: cleanString_(row[columnMap.location - 1])
    };
    return acc;
  }, {});
}

function validateSavePayloadShape_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid save payload.');
  }
  if (!Array.isArray(payload.items) || !payload.items.length) {
    throw new Error('Invalid save payload: no item updates were supplied.');
  }
}

function safeWebAppUrlForLinks_() {
  try {
    return getWebAppBaseUrl_();
  } catch (err) {
    return '';
  }
}

function groupLocationsByRoom_(locations) {
  const groups = {};
  locations.forEach(function (entry) {
    if (!groups[entry.room]) groups[entry.room] = [];
    groups[entry.room].push(entry);
  });

  return Object.keys(groups)
    .sort()
    .map(function (room) {
      return {
        room: room,
        locations: groups[room].sort(function (a, b) {
          return a.location.localeCompare(b.location);
        })
      };
    });
}

function normalizeHeader_(value) {
  return cleanString_(value).toLowerCase().replace(/\s+/g, ' ');
}

function cleanString_(value) {
  return String(value == null ? '' : value).trim();
}

function stripQueryString_(url) {
  return cleanString_(url).split('?')[0];
}

function createAppUrl_(baseUrl, room, location, mode) {
  const root = stripQueryString_(baseUrl);
  if (!root) return '#';

  const parts = [
    'room=' + encodeURIComponent(room),
    'loc=' + encodeURIComponent(location)
  ];
  if (mode) parts.push('mode=' + encodeURIComponent(mode));
  return root + '?' + parts.join('&');
}

function statusClass_(status) {
  const value = cleanString_(status);
  if (value === 'Good') return 'status-good';
  if (value === 'Low Stock') return 'status-low';
  if (value === 'Missing') return 'status-missing';
  if (value === 'Needs Maintenance') return 'status-maintenance';
  return '';
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function renderAppHtml_(page) {
  const baseUrl = safeWebAppUrlForLinks_();
  const viewHref = page.hasSelection ? createAppUrl_(baseUrl, page.room, page.location, '') : '#';
  const techHref = page.hasSelection ? createAppUrl_(baseUrl, page.room, page.location, 'tech') : '#';

  const bootstrap = JSON.stringify({
    mode: page.mode,
    room: page.room,
    location: page.location,
    hasSelection: page.hasSelection,
    statuses: APP.statuses,
    debug: DEBUG,
    itemCount: page.items.length,
    locationCount: page.locations.length
  }).replace(/</g, '\\u003c');

  const content = page.hasSelection ? renderLocationView_(page) : renderLandingView_(page);
  const debugPanel = DEBUG ? renderDebugPanel_(page) : '';

  return [
    '<!DOCTYPE html>',
    '<html><head>',
    '<meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<title>', escapeHtml_(APP.title), '</title>',
    '<style>', baseStyles_(), '</style>',
    '</head><body>',
    '<div class="app-shell">',
    '<header class="page-header">',
    '  <h1>', escapeHtml_(APP.title), '</h1>',
    '  <div class="subhead"><strong>Room:</strong> ', escapeHtml_(page.room || '-'), ' · <strong>Location:</strong> ', escapeHtml_(page.location || '-'), '</div>',
    '  <div class="top-actions">',
    '    <a class="button secondary ', page.hasSelection ? '' : 'disabled', '" href="', page.hasSelection ? escapeHtml_(viewHref) : '#', '">View Mode</a>',
    '    <a class="button primary ', page.hasSelection ? '' : 'disabled', '" href="', page.hasSelection ? escapeHtml_(techHref) : '#', '">Technician Access</a>',
    '  </div>',
    '</header>',
    content,
    debugPanel,
    '</div>',
    '<script>const BOOTSTRAP=', bootstrap, ';</script>',
    '<script>', clientScript_(), '</script>',
    '</body></html>'
  ].join('');
}

function renderLandingView_(page) {
  const groupsHtml = page.locationsByRoom.length
    ? page.locationsByRoom.map(function (group) {
        return [
          '<section class="room-group">',
          '<div class="room-heading">ROOM ', escapeHtml_(group.room), '</div>',
          group.locations.map(function (entry) {
            return [
              '<div class="location-card">',
              '<div class="location-name">', escapeHtml_(entry.location), '</div>',
              '<div class="location-meta">Room: ', escapeHtml_(entry.room), '</div>',
              '<div class="location-actions">',
              '<a class="button secondary" href="', escapeHtml_(entry.viewUrl), '">Open View</a>',
              '<a class="button primary" href="', escapeHtml_(entry.techUrl), '">Open Tech</a>',
              '</div>',
              '</div>'
            ].join('');
          }).join(''),
          '</section>'
        ].join('');
      }).join('')
    : '<div class="empty-state">' + escapeHtml_(page.notice || 'No locations found.') + '</div>';

  return [
    '<main class="panel">',
    '<div class="panel-head"><h2>Landing</h2><span class="mode-pill">Browse Locations</span></div>',
    '<div class="hero"><h3>Choose a location</h3><p>Scan a QR code or pick a room/location below.</p></div>',
    (!page.appUrlConfigured ? '<div class="notice error">Configuration error: WEB_APP_BASE_URL is not set. Links and QR generation will not work until configured.</div>' : ''),
    groupsHtml,
    '</main>'
  ].join('');
}

function renderLocationView_(page) {
  const modeLabel = page.mode === 'tech' ? 'Tech Mode' : 'View Mode';

  const items = page.items.length
    ? page.items.map(function (item) { return renderItemCard_(item, page.mode === 'tech'); }).join('')
    : '<div class="empty-state">' + escapeHtml_(page.notice || 'No inventory records found for this location.') + '</div>';

  const saveBar = page.mode === 'tech'
    ? '<div class="save-bar"><button id="saveButton" class="button primary">Save Changes</button><span id="saveMessage" class="save-message"></span></div>'
    : '';

  return [
    '<main class="panel">',
    '<div class="panel-head"><h2>Inventory Items</h2><span class="mode-pill">', escapeHtml_(modeLabel), '</span></div>',
    page.notice ? '<div class="notice info">' + escapeHtml_(page.notice) + '</div>' : '',
    saveBar,
    '<div id="inventoryList" class="inventory-list">', items, '</div>',
    '</main>'
  ].join('');
}

function renderItemCard_(item, techMode) {
  const isChemical = normalizeHeader_(item.category) === 'chemicals';
  const qtyValue = escapeHtml_(String(item.qty));
  const statusClass = statusClass_(item.status);

  return [
    '<article class="item-card ', isChemical ? 'hazard' : '', '" data-row="', escapeHtml_(String(item.rowNumber)), '">',
    '<div class="item-header">',
    '<div><div class="item-name">', escapeHtml_(item.itemName), '</div><div class="item-meta">', escapeHtml_(item.itemId), ' · ', escapeHtml_(item.category), '</div></div>',
    isChemical ? '<span class="hazard-badge">HAZARD</span>' : '',
    '</div>',
    '<div class="item-grid">',
    '<div><label>Quantity</label>',
    techMode
      ? '<input class="field-input qty-input" type="number" min="0" step="1" value="' + qtyValue + '">'
      : '<div class="field-value">' + qtyValue + '</div>',
    '</div>',
    '<div><label>Status</label>',
    techMode ? renderStatusSelect_(item.status) : '<div class="status-pill ' + statusClass + '">' + escapeHtml_(item.status) + '</div>',
    '</div>',
    '</div>',
    isChemical ? '<div class="hazard-note">Chemical item — follow storage and handling procedures.</div>' : '',
    '</article>'
  ].join('');
}

function renderStatusSelect_(selectedValue) {
  const options = APP.statuses.map(function (status) {
    const selected = status === selectedValue ? ' selected' : '';
    return '<option value="' + escapeHtml_(status) + '"' + selected + '>' + escapeHtml_(status) + '</option>';
  }).join('');
  return '<select class="field-select status-select">' + options + '</select>';
}

function renderDebugPanel_(page) {
  return [
    '<section class="panel debug-panel">',
    '<div class="panel-head"><h2>Debug</h2></div>',
    '<div class="debug-grid">',
    '<div><strong>mode:</strong> ', escapeHtml_(page.mode), '</div>',
    '<div><strong>room:</strong> ', escapeHtml_(page.room || '-'), '</div>',
    '<div><strong>location:</strong> ', escapeHtml_(page.location || '-'), '</div>',
    '<div><strong>items loaded:</strong> ', escapeHtml_(String(page.items.length)), '</div>',
    '<div><strong>locations loaded:</strong> ', escapeHtml_(String(page.locations.length)), '</div>',
    '<div><strong>google.script.run:</strong> <span id="bridgeStatus">(checking...)</span></div>',
    '</div>',
    '</section>'
  ].join('');
}

function renderConfigErrorHtml_(err) {
  const message = err && err.message ? err.message : String(err);
  return [
    '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">',
    '<title>', escapeHtml_(APP.title), '</title>',
    '<style>', baseStyles_(), '</style>',
    '</head><body><div class="app-shell">',
    '<header class="page-header"><h1>', escapeHtml_(APP.title), '</h1></header>',
    '<main class="panel">',
    '<div class="panel-head"><h2>Configuration error</h2></div>',
    '<div class="notice error">', escapeHtml_(message), '</div>',
    '<div class="empty-state">Set Script Properties: SPREADSHEET_ID and WEB_APP_BASE_URL (and optional INVENTORY_SHEET_NAME). Then redeploy if needed.</div>',
    '</main></div></body></html>'
  ].join('');
}

function baseStyles_() {
  return [
    ':root{--bg:#eef3f8;--card:#fff;--text:#14213d;--muted:#55627a;--line:#d6dee8;--primary:#4f46e5;--secondary:#e7edf4;--success:#e6f4ea;--success-text:#166534;--warn:#fff4e5;--warn-text:#b45309;--danger:#fde8e8;--danger-text:#b91c1c;--hazard:#fff1f2;--shadow:0 2px 10px rgba(16,24,40,.06);--radius:20px;}',
    '*{box-sizing:border-box}',
    'body{margin:0;font-family:Arial,Helvetica,sans-serif;background:var(--bg);color:var(--text)}',
    '.app-shell{max-width:1180px;margin:0 auto;padding:28px 20px 48px}',
    '.page-header h1{margin:0 0 8px;font-size:44px;line-height:1.1;letter-spacing:-0.02em}',
    '.subhead{color:var(--muted);font-size:18px;margin-bottom:16px}',
    '.top-actions{display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap}',
    '.button{display:inline-flex;align-items:center;justify-content:center;text-decoration:none;border:none;border-radius:12px;padding:12px 16px;font-size:15px;font-weight:700;cursor:pointer}',
    '.button.primary{background:var(--primary);color:#fff}',
    '.button.secondary{background:var(--secondary);color:#243147}',
    '.button.disabled{opacity:.45;pointer-events:none}',
    '.panel{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);overflow:hidden;box-shadow:var(--shadow)}',
    '.panel-head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:18px 20px;border-bottom:1px solid var(--line)}',
    '.panel-head h2{margin:0;font-size:24px}',
    '.mode-pill{padding:8px 12px;border-radius:999px;background:#dbeafe;color:#1d4ed8;font-size:13px;font-weight:700}',
    '.hero{background:#eaf2ff;padding:20px;border-bottom:1px solid var(--line)}',
    '.hero h3{margin:0 0 8px;font-size:24px}',
    '.hero p{margin:0;color:#334155;font-size:16px}',
    '.room-group + .room-group{border-top:1px solid var(--line)}',
    '.room-heading{padding:12px 20px;background:#f4f7fb;border-bottom:1px solid var(--line);font-weight:800;letter-spacing:.06em;color:#475569;font-size:13px}',
    '.location-card{padding:18px 20px;border-bottom:1px solid var(--line)}',
    '.location-card:last-child{border-bottom:none}',
    '.location-name{font-size:22px;font-weight:800;margin-bottom:6px}',
    '.location-meta{color:#64748b;font-size:15px;margin-bottom:14px}',
    '.location-actions{display:flex;gap:10px;flex-wrap:wrap}',
    '.inventory-list{padding:16px;display:grid;gap:12px}',
    '.item-card{border:1px solid var(--line);border-radius:16px;padding:14px;background:#fff}',
    '.item-card.hazard{background:var(--hazard);border-color:#fecdd3}',
    '.item-header{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;margin-bottom:10px}',
    '.item-name{font-size:20px;font-weight:800;margin-bottom:4px}',
    '.item-meta{color:#64748b;font-size:14px}',
    '.hazard-badge{background:#dc2626;color:#fff;border-radius:999px;padding:6px 10px;font-size:11px;font-weight:800;letter-spacing:.04em}',
    '.item-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}',
    'label{display:block;font-size:12px;font-weight:700;color:#64748b;margin-bottom:6px;text-transform:uppercase;letter-spacing:.04em}',
    '.field-input,.field-select,.field-value{width:100%;min-height:42px;border-radius:10px;border:1px solid var(--line);padding:10px 12px;font-size:15px;background:#fff}',
    '.field-value{display:flex;align-items:center;font-weight:700;color:#1f2937}',
    '.status-pill{display:inline-flex;align-items:center;justify-content:center;min-height:42px;border-radius:10px;padding:8px 12px;font-weight:800;border:1px solid transparent;width:100%}',
    '.status-good{background:var(--success);color:var(--success-text);border-color:#bbf7d0}',
    '.status-low{background:var(--warn);color:var(--warn-text);border-color:#fed7aa}',
    '.status-missing{background:var(--danger);color:var(--danger-text);border-color:#fecaca}',
    '.status-maintenance{background:#fff7ed;color:#c2410c;border-color:#fdba74}',
    '.hazard-note{margin-top:10px;color:#991b1b;font-size:13px;font-weight:700}',
    '.notice{margin:14px 16px 0;padding:12px 14px;border-radius:12px;font-weight:700}',
    '.notice.info{background:#eff6ff;color:#1d4ed8}',
    '.notice.error{background:var(--danger);color:var(--danger-text)}',
    '.save-bar{display:flex;align-items:center;gap:12px;padding:16px;border-bottom:1px solid var(--line);flex-wrap:wrap}',
    '.save-message{font-weight:700;color:#334155}',
    '.empty-state{padding:20px;color:#64748b;font-size:16px}',
    '.debug-panel{margin-top:16px}',
    '.debug-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;padding:16px 20px}',
    '@media (max-width:720px){.page-header h1{font-size:34px}.subhead{font-size:16px}.panel-head h2{font-size:20px}.item-grid{grid-template-columns:1fr}.button{width:100%}.location-actions .button{width:auto}}'
  ].join('');
}

function clientScript_() {
  return [
    '(function(){',
    'var bridgeAvailable=!!(window.google&&google.script&&google.script.run);',
    'if(BOOTSTRAP.debug){var statusNode=document.getElementById("bridgeStatus");if(statusNode){statusNode.textContent=bridgeAvailable?"available":"unavailable";}}',
    'if(!(BOOTSTRAP.hasSelection&&BOOTSTRAP.mode==="tech")){return;}',
    'var saveButton=document.getElementById("saveButton");',
    'var saveMessage=document.getElementById("saveMessage");',
    'if(!bridgeAvailable){if(saveMessage){saveMessage.textContent="Interactive save is unavailable in this context. Open the deployed /exec web app URL.";}if(saveButton){saveButton.disabled=true;saveButton.classList.add("disabled");}return;}',
    'if(!saveButton){return;}',
    'saveButton.addEventListener("click",function(){',
    'var cards=Array.prototype.slice.call(document.querySelectorAll(".item-card[data-row]"));',
    'var items=cards.map(function(card){',
    'var rowNumber=Number(card.getAttribute("data-row"));',
    'var qtyInput=card.querySelector(".qty-input");',
    'var statusSelect=card.querySelector(".status-select");',
    'return{rowNumber:rowNumber,qty:qtyInput?Number(qtyInput.value):0,status:statusSelect?statusSelect.value:""};',
    '});',
    'if(saveMessage){saveMessage.textContent="Saving...";}',
    'google.script.run.withSuccessHandler(function(result){',
    'if(!result||!result.ok){if(saveMessage){saveMessage.textContent=(result&&result.message)?result.message:"Save failed.";}return;}',
    'if(saveMessage){saveMessage.textContent=result.message||"Saved.";}',
    '}).withFailureHandler(function(error){',
    'if(saveMessage){saveMessage.textContent=(error&&error.message)?error.message:String(error);}',
    '}).saveInventoryUpdates({room:BOOTSTRAP.room,location:BOOTSTRAP.location,items:items});',
    '});',
    '})();'
  ].join('');
}
