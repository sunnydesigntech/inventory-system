const APP = Object.freeze({
  title: 'D&T Inventory',
  inventorySheetName: 'Inventory',
  statuses: ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'],
  debug: false,
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
    qrLink: ['qr code link (auto-generated)', 'qr link', 'qr url', 'qr code link']
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
    return HtmlService.createHtmlOutput(renderErrorHtml_(err))
      .setTitle(APP.title);
  }
}

/**
 * Save button handler from technician mode.
 * Payload format:
 * {
 *   room: "419A",
 *   location: "Chemical Cabinet",
 *   items: [{ rowNumber: 4, qty: 2, status: "Low Stock" }]
 * }
 */
function saveInventoryUpdates(payload) {
  try {
    if (!payload || !Array.isArray(payload.items) || payload.items.length === 0) {
      throw new Error('No update data was received.');
    }

    const room = cleanString_(payload.room);
    const location = cleanString_(payload.location);
    if (!room || !location) {
      throw new Error('Missing room or location context for save.');
    }

    const sheet = getInventorySheet_();
    const colMap = getColumnMap_(sheet);
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error('The inventory sheet does not contain any data rows.');
    }

    const width = Math.max(...Object.values(colMap));
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    const rowLookup = {};
    values.forEach(function (row, idx) {
      const absoluteRow = idx + 2;
      rowLookup[absoluteRow] = row;
    });

    const allowedStatuses = APP.statuses.reduce(function (acc, status) {
      acc[status] = true;
      return acc;
    }, {});

    const updates = payload.items.map(function (item) {
      const rowNumber = Number(item.rowNumber);
      if (!rowNumber || !rowLookup[rowNumber]) {
        throw new Error('Invalid row number: ' + item.rowNumber);
      }

      const sourceRow = rowLookup[rowNumber];
      const sourceRoom = cleanString_(sourceRow[colMap.room - 1]);
      const sourceLocation = cleanString_(sourceRow[colMap.location - 1]);

      if (sourceRoom !== room || sourceLocation !== location) {
        throw new Error('Row ' + rowNumber + ' does not belong to ' + room + ' / ' + location + '.');
      }

      const qty = Number(item.qty);
      if (!isFinite(qty) || qty < 0) {
        throw new Error('Invalid quantity for row ' + rowNumber + '.');
      }

      const status = cleanString_(item.status);
      if (!allowedStatuses[status]) {
        throw new Error('Invalid status for row ' + rowNumber + ': ' + status);
      }

      return {
        rowNumber: rowNumber,
        qty: qty,
        status: status
      };
    });

    updates.forEach(function (item) {
      sheet.getRange(item.rowNumber, colMap.qty).setValue(item.qty);
      sheet.getRange(item.rowNumber, colMap.status).setValue(item.status);
    });

    const refreshedItems = getInventoryRowsForLocation_(room, location);

    return {
      ok: true,
      message: 'Inventory updated successfully.',
      room: room,
      location: location,
      items: refreshedItems
    };
  } catch (err) {
    return {
      ok: false,
      message: err.message || String(err)
    };
  }
}

/**
 * Generates web app links back into the "QR Code Link (Auto-Generated)" column.
 */
function refreshQrLinks() {
  const baseUrl = getWebAppBaseUrl_();
  const sheet = getInventorySheet_();
  const colMap = getColumnMap_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { ok: true, updatedRows: 0, message: 'No data rows found.' };
  }

  const width = Math.max(...Object.values(colMap));
  const data = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const formulas = [];
  let updatedRows = 0;

  data.forEach(function (row) {
    const room = cleanString_(row[colMap.room - 1]);
    const location = cleanString_(row[colMap.location - 1]);

    if (!room || !location) {
      formulas.push(['']);
      return;
    }

    const url = createAppUrl_(baseUrl, room, location, '');
    formulas.push([url]);
    updatedRows += 1;
  });

  sheet.getRange(2, colMap.qrLink, formulas.length, 1).setValues(formulas);

  return {
    ok: true,
    updatedRows: updatedRows,
    message: 'QR links refreshed for ' + updatedRows + ' row(s).'
  };
}

/**
 * Optional helper to store both required properties in one run.
 */
function setAppConfig(spreadsheetId, webAppBaseUrl) {
  const props = PropertiesService.getScriptProperties();

  if (!spreadsheetId) {
    throw new Error('spreadsheetId is required.');
  }

  props.setProperty(APP.propertyKeys.spreadsheetId, spreadsheetId);

  if (webAppBaseUrl) {
    props.setProperty(APP.propertyKeys.webAppBaseUrl, stripQueryString_(webAppBaseUrl));
  }

  props.setProperty(APP.propertyKeys.inventorySheetName, APP.inventorySheetName);

  return {
    ok: true,
    spreadsheetId: spreadsheetId,
    webAppBaseUrl: webAppBaseUrl ? stripQueryString_(webAppBaseUrl) : '',
    message: 'Script Properties updated.'
  };
}

function buildPageState_(params) {
  const mode = params.mode === 'tech' ? 'tech' : 'view';
  const room = params.room || '';
  const location = params.location || '';
  const hasSelection = !!(room && location);

  const page = {
    title: APP.title,
    mode: mode,
    room: room,
    location: location,
    hasSelection: hasSelection,
    debug: APP.debug,
    error: '',
    notice: '',
    items: [],
    locations: [],
    locationsByRoom: []
  };

  if (hasSelection) {
    const items = getInventoryRowsForLocation_(room, location);
    page.items = items;
    if (items.length === 0) {
      page.notice = 'No inventory records found for this location.';
    }
  } else {
    const locations = getAllLocations_();
    page.locations = locations;
    page.locationsByRoom = groupLocationsByRoom_(locations);
    if (locations.length === 0) {
      page.notice = 'No locations are available yet. Add inventory rows to the sheet first.';
    }
  }

  return page;
}

function getAllLocations_() {
  const sheet = getInventorySheet_();
  const colMap = getColumnMap_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const width = Math.max(...Object.values(colMap));
  const data = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const baseUrl = safeWebAppUrlForLinks_();
  const dedup = {};

  data.forEach(function (row) {
    const room = cleanString_(row[colMap.room - 1]);
    const location = cleanString_(row[colMap.location - 1]);

    if (!room || !location) return;

    const key = room + '|||'+ location;
    if (!dedup[key]) {
      dedup[key] = {
        room: room,
        location: location,
        viewUrl: createAppUrl_(baseUrl, room, location, ''),
        techUrl: createAppUrl_(baseUrl, room, location, 'tech')
      };
    }
  });

  return Object.keys(dedup)
    .map(function (key) { return dedup[key]; })
    .sort(function (a, b) {
      return (a.room + '|' + a.location).localeCompare(b.room + '|' + b.location);
    });
}

function getInventoryRowsForLocation_(room, location) {
  const targetRoom = cleanString_(room);
  const targetLocation = cleanString_(location);

  if (!targetRoom || !targetLocation) return [];

  const sheet = getInventorySheet_();
  const colMap = getColumnMap_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const width = Math.max(...Object.values(colMap));
  const data = sheet.getRange(2, 1, lastRow - 1, width).getValues();

  return data
    .map(function (row, index) {
      return {
        rowNumber: index + 2,
        itemId: row[colMap.itemId - 1],
        itemName: row[colMap.itemName - 1],
        room: cleanString_(row[colMap.room - 1]),
        location: cleanString_(row[colMap.location - 1]),
        qty: row[colMap.qty - 1],
        category: cleanString_(row[colMap.category - 1]),
        status: cleanString_(row[colMap.status - 1])
      };
    })
    .filter(function (item) {
      return item.room === targetRoom && item.location === targetLocation;
    })
    .sort(function (a, b) {
      return String(a.itemName).localeCompare(String(b.itemName));
    });
}

function getInventorySheet_() {
  const ss = getSpreadsheet_();
  const preferredName = PropertiesService.getScriptProperties()
    .getProperty(APP.propertyKeys.inventorySheetName) || APP.inventorySheetName;

  return ss.getSheetByName(preferredName) || ss.getSheets()[0];
}

function getSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(APP.propertyKeys.spreadsheetId);

  if (!spreadsheetId) {
    throw new Error(
      'SPREADSHEET_ID is not set. Set it in Script Properties, or run setAppConfig(spreadsheetId, webAppBaseUrl).'
    );
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

function getWebAppBaseUrl_() {
  const props = PropertiesService.getScriptProperties();
  const configured = cleanString_(props.getProperty(APP.propertyKeys.webAppBaseUrl));

  if (configured) return stripQueryString_(configured);

  const scriptUrl = ScriptApp.getService().getUrl();
  if (scriptUrl) return stripQueryString_(scriptUrl);

  throw new Error(
    'WEB_APP_BASE_URL is not set. Deploy the web app, then save the /exec URL in Script Properties.'
  );
}

function safeWebAppUrlForLinks_() {
  try {
    return getWebAppBaseUrl_();
  } catch (err) {
    return '#';
  }
}

function getColumnMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    throw new Error('The inventory sheet has no header row.');
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);
  const map = {};

  Object.keys(APP.headerAliases).forEach(function (logicalKey) {
    const aliases = APP.headerAliases[logicalKey];
    const index = headers.findIndex(function (header) {
      return aliases.indexOf(header) !== -1;
    });

    if (index === -1) {
      throw new Error('Missing required column: ' + logicalKey);
    }

    map[logicalKey] = index + 1;
  });

  return map;
}

function parseParams_(e) {
  const p = (e && e.parameter) || {};
  return {
    room: cleanString_(p.room),
    location: cleanString_(p.loc),
    mode: cleanString_(p.mode).toLowerCase() === 'tech' ? 'tech' : 'view'
  };
}

function normalizeHeader_(value) {
  return cleanString_(value)
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanString_(value) {
  return String(value == null ? '' : value).trim();
}

function stripQueryString_(url) {
  return cleanString_(url).split('?')[0];
}

function createAppUrl_(baseUrl, room, location, mode) {
  if (!baseUrl || baseUrl === '#') return '#';

  const parts = [
    'room=' + encodeURIComponent(room),
    'loc=' + encodeURIComponent(location)
  ];

  if (mode) {
    parts.push('mode=' + encodeURIComponent(mode));
  }

  return stripQueryString_(baseUrl) + '?' + parts.join('&');
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

function renderAppHtml_(page) {
  const bootstrapJson = JSON.stringify({
    mode: page.mode,
    room: page.room,
    location: page.location,
    hasSelection: page.hasSelection,
    statuses: APP.statuses,
    items: page.items,
    locations: page.locations,
    debug: APP.debug
  }).replace(/</g, '\\u003c');

  const roomText = page.room ? escapeHtml_(page.room) : '-';
  const locationText = page.location ? escapeHtml_(page.location) : '-';

  const viewHref = page.hasSelection ? createAppUrl_(safeWebAppUrlForLinks_(), page.room, page.location, '') : '#';
  const techHref = page.hasSelection ? createAppUrl_(safeWebAppUrlForLinks_(), page.room, page.location, 'tech') : '#';

  const contentHtml = page.hasSelection
    ? renderLocationView_(page)
    : renderLandingView_(page);

  const debugHtml = page.debug ? renderDebugPanel_(page) : '';

  return [
    '<!DOCTYPE html>',
    '<html><head>',
    '<meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<title>', escapeHtml_(APP.title), '</title>',
    '<style>',
    baseStyles_(),
    '</style>',
    '</head><body>',
    '<div class="app-shell">',
    '  <header class="page-header">',
    '    <h1>', escapeHtml_(APP.title), '</h1>',
    '    <div class="subhead"><strong>Room:</strong> ', roomText, ' · <strong>Location:</strong> ', locationText, '</div>',
    '    <div class="top-actions">',
    '      <a class="button secondary ', page.hasSelection && page.mode !== 'view' ? '' : (!page.hasSelection ? 'disabled' : ''), '" href="', page.hasSelection ? escapeHtml_(viewHref) : '#', '">View Mode</a>',
    '      <a class="button primary ', page.hasSelection && page.mode === 'tech' ? '' : (!page.hasSelection ? 'disabled' : ''), '" href="', page.hasSelection ? escapeHtml_(techHref) : '#', '">Technician Access</a>',
    '    </div>',
    '  </header>',
    '  <main>',
         contentHtml,
         debugHtml,
    '  </main>',
    '</div>',
    '<script>',
    'const BOOTSTRAP = ', bootstrapJson, ';',
    clientScript_(),
    '</script>',
    '</body></html>'
  ].join('');
}

function renderLandingView_(page) {
  const groupHtml = page.locationsByRoom.length
    ? page.locationsByRoom.map(function (group) {
        return [
          '<section class="room-group">',
          '  <div class="room-heading">ROOM ', escapeHtml_(group.room), '</div>',
             group.locations.map(function (entry) {
               return [
                 '<div class="location-card">',
                 '  <div class="location-name">', escapeHtml_(entry.location), '</div>',
                 '  <div class="location-meta">Room: ', escapeHtml_(entry.room), '</div>',
                 '  <div class="location-actions">',
                 '    <a class="button secondary" href="', escapeHtml_(entry.viewUrl), '">Open View</a>',
                 '    <a class="button primary" href="', escapeHtml_(entry.techUrl), '">Open Tech</a>',
                 '  </div>',
                 '</div>'
               ].join('');
             }).join(''),
          '</section>'
        ].join('');
      }).join('')
    : '<div class="empty-state">' + escapeHtml_(page.notice || 'No locations found.') + '</div>';

  return [
    '<section class="panel">',
    '  <div class="panel-head">',
    '    <h2>Inventory Items</h2>',
    '    <span class="mode-pill">Landing</span>',
    '  </div>',
    '  <div class="hero">',
    '    <h3>Choose a location</h3>',
    '    <p>Scan a QR code or select a room/location below.</p>',
    '  </div>',
         page.notice && page.locationsByRoom.length ? '<div class="notice info">' + escapeHtml_(page.notice) + '</div>' : '',
         groupHtml,
    '</section>'
  ].join('');
}

function renderLocationView_(page) {
  const modeLabel = page.mode === 'tech' ? 'Tech Mode' : 'View Mode';
  const itemsHtml = page.items.length
    ? page.items.map(function (item) { return renderItemCard_(item, page.mode === 'tech'); }).join('')
    : '<div class="empty-state">' + escapeHtml_(page.notice || 'No inventory records found for this location.') + '</div>';

  const saveBar = page.mode === 'tech'
    ? [
        '<div class="save-bar">',
        '  <button id="saveButton" class="button primary">Save Changes</button>',
        '  <span id="saveMessage" class="save-message"></span>',
        '</div>'
      ].join('')
    : '';

  return [
    '<section class="panel">',
    '  <div class="panel-head">',
    '    <h2>Inventory Items</h2>',
    '    <span class="mode-pill">', escapeHtml_(modeLabel), '</span>',
    '  </div>',
         page.notice ? '<div class="notice info">' + escapeHtml_(page.notice) + '</div>' : '',
         saveBar,
    '  <div id="inventoryList" class="inventory-list">',
         itemsHtml,
    '  </div>',
    '</section>'
  ].join('');
}

function renderItemCard_(item, techMode) {
  const isChemical = normalizeHeader_(item.category) === 'chemicals';
  const statusClass = statusClass_(item.status);
  const qtyValue = escapeHtml_(String(item.qty));

  return [
    '<div class="item-card ', isChemical ? 'hazard' : '', '" data-row="', escapeHtml_(String(item.rowNumber)), '">',
    '  <div class="item-header">',
    '    <div>',
    '      <div class="item-name">', escapeHtml_(item.itemName), '</div>',
    '      <div class="item-meta">', escapeHtml_(item.itemId), ' · ', escapeHtml_(item.category), '</div>',
    '    </div>',
           isChemical ? '<span class="hazard-badge">HAZARD</span>' : '',
    '  </div>',
    '  <div class="item-grid">',
    '    <div>',
    '      <label>Quantity</label>',
           techMode
             ? '<input class="field-input qty-input" type="number" min="0" step="1" value="' + qtyValue + '">'
             : '<div class="field-value">' + qtyValue + '</div>',
    '    </div>',
    '    <div>',
    '      <label>Status</label>',
           techMode
             ? renderStatusSelect_(item.status)
             : '<div class="status-pill ' + statusClass + '">' + escapeHtml_(item.status) + '</div>',
    '    </div>',
    '  </div>',
         isChemical ? '<div class="hazard-note">Chemical item — check handling, storage, and safety procedures before use.</div>' : '',
    '</div>'
  ].join('');
}

function renderStatusSelect_(currentStatus) {
  const options = APP.statuses.map(function (status) {
    const selected = status === currentStatus ? ' selected' : '';
    return '<option value="' + escapeHtml_(status) + '"' + selected + '>' + escapeHtml_(status) + '</option>';
  }).join('');

  return '<select class="field-select status-select">' + options + '</select>';
}

function statusClass_(status) {
  switch (cleanString_(status)) {
    case 'Good':
      return 'status-good';
    case 'Low Stock':
      return 'status-low';
    case 'Missing':
      return 'status-missing';
    case 'Needs Maintenance':
      return 'status-maintenance';
    default:
      return '';
  }
}

function renderDebugPanel_(page) {
  return [
    '<section class="panel debug-panel">',
    '  <div class="panel-head"><h2>Debug</h2></div>',
    '  <div class="debug-grid">',
    '    <div><strong>Mode:</strong> ', escapeHtml_(page.mode), '</div>',
    '    <div><strong>Room:</strong> ', escapeHtml_(page.room || '-'), '</div>',
    '    <div><strong>Location:</strong> ', escapeHtml_(page.location || '-'), '</div>',
    '    <div><strong>Items:</strong> ', escapeHtml_(String(page.items.length)), '</div>',
    '    <div><strong>Locations:</strong> ', escapeHtml_(String(page.locations.length)), '</div>',
    '  </div>',
    '</section>'
  ].join('');
}

function renderErrorHtml_(err) {
  const message = err && err.message ? err.message : String(err);
  return [
    '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">',
    '<title>', escapeHtml_(APP.title), '</title><style>', baseStyles_(), '</style></head><body>',
    '<div class="app-shell"><header class="page-header"><h1>', escapeHtml_(APP.title), '</h1></header>',
    '<section class="panel"><div class="notice error"><strong>Configuration error:</strong> ', escapeHtml_(message), '</div>',
    '<div class="empty-state">Check Script Properties for SPREADSHEET_ID and WEB_APP_BASE_URL, and confirm the Inventory sheet headers are correct.</div>',
    '</section></div></body></html>'
  ].join('');
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function baseStyles_() {
  return `
    :root{
      --bg:#eef3f8;
      --card:#ffffff;
      --text:#14213d;
      --muted:#55627a;
      --line:#d6dee8;
      --primary:#4f46e5;
      --primary-soft:#c7d2fe;
      --secondary:#e7edf4;
      --success:#e6f4ea;
      --success-text:#166534;
      --warn:#fff4e5;
      --warn-text:#b45309;
      --danger:#fde8e8;
      --danger-text:#b91c1c;
      --hazard:#fff1f2;
      --shadow:0 2px 10px rgba(16,24,40,.06);
      --radius:20px;
    }
    *{box-sizing:border-box}
    body{
      margin:0;
      font-family:Arial,Helvetica,sans-serif;
      background:var(--bg);
      color:var(--text);
    }
    .app-shell{
      max-width:1180px;
      margin:0 auto;
      padding:28px 20px 48px;
    }
    .page-header h1{
      margin:0 0 8px;
      font-size:56px;
      line-height:1;
      letter-spacing:-0.03em;
    }
    .subhead{
      color:var(--muted);
      font-size:22px;
      margin-bottom:22px;
    }
    .top-actions{
      display:flex;
      gap:14px;
      margin-bottom:26px;
      flex-wrap:wrap;
    }
    .button{
      display:inline-flex;
      align-items:center;
      justify-content:center;
      text-decoration:none;
      border:none;
      border-radius:14px;
      padding:14px 20px;
      font-size:16px;
      font-weight:700;
      cursor:pointer;
      transition:.15s ease;
    }
    .button.primary{
      background:var(--primary);
      color:#fff;
    }
    .button.secondary{
      background:var(--secondary);
      color:#243147;
    }
    .button.disabled{
      opacity:.45;
      pointer-events:none;
    }
    .panel{
      background:var(--card);
      border:1px solid var(--line);
      border-radius:var(--radius);
      overflow:hidden;
      box-shadow:var(--shadow);
    }
    .panel + .panel{
      margin-top:18px;
    }
    .panel-head{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:12px;
      padding:20px 26px;
      border-bottom:1px solid var(--line);
      background:#fff;
    }
    .panel-head h2{
      margin:0;
      font-size:28px;
    }
    .mode-pill{
      padding:10px 16px;
      border-radius:999px;
      background:#dbeafe;
      color:#2563eb;
      font-size:14px;
      font-weight:700;
    }
    .hero{
      background:#eaf2ff;
      padding:26px;
      border-bottom:1px solid var(--line);
    }
    .hero h3{
      margin:0 0 10px;
      font-size:28px;
    }
    .hero p{
      margin:0;
      color:#334155;
      font-size:20px;
    }
    .room-group + .room-group{
      border-top:1px solid var(--line);
    }
    .room-heading{
      padding:14px 26px;
      background:#f4f7fb;
      border-bottom:1px solid var(--line);
      font-weight:800;
      letter-spacing:.06em;
      color:#475569;
      font-size:14px;
    }
    .location-card{
      padding:26px;
      border-bottom:1px solid var(--line);
      background:#fff;
    }
    .location-card:last-child{
      border-bottom:none;
    }
    .location-name{
      font-size:24px;
      font-weight:800;
      margin-bottom:8px;
    }
    .location-meta{
      color:#64748b;
      font-size:16px;
      margin-bottom:18px;
    }
    .location-actions{
      display:flex;
      gap:12px;
      flex-wrap:wrap;
    }
    .inventory-list{
      padding:20px;
      display:grid;
      gap:16px;
    }
    .item-card{
      border:1px solid var(--line);
      border-radius:18px;
      padding:18px;
      background:#fff;
    }
    .item-card.hazard{
      background:var(--hazard);
      border-color:#fecdd3;
    }
    .item-header{
      display:flex;
      align-items:flex-start;
      justify-content:space-between;
      gap:12px;
      margin-bottom:14px;
    }
    .item-name{
      font-size:22px;
      font-weight:800;
      margin-bottom:6px;
    }
    .item-meta{
      color:#64748b;
      font-size:15px;
    }
    .hazard-badge{
      background:#dc2626;
      color:#fff;
      border-radius:999px;
      padding:8px 12px;
      font-size:12px;
      font-weight:800;
      letter-spacing:.04em;
    }
    .item-grid{
      display:grid;
      grid-template-columns:repeat(2,minmax(0,1fr));
      gap:14px;
      margin-top:8px;
    }
    label{
      display:block;
      font-size:13px;
      font-weight:700;
      color:#64748b;
      margin-bottom:8px;
      text-transform:uppercase;
      letter-spacing:.04em;
    }
    .field-input,.field-select,.field-value{
      width:100%;
      min-height:48px;
      border-radius:12px;
      border:1px solid var(--line);
      padding:12px 14px;
      font-size:16px;
      background:#fff;
    }
    .field-value{
      display:flex;
      align-items:center;
      font-weight:700;
      color:#1f2937;
    }
    .status-pill{
      display:inline-flex;
      align-items:center;
      justify-content:center;
      min-height:48px;
      border-radius:12px;
      padding:10px 14px;
      font-weight:800;
      border:1px solid transparent;
      width:100%;
    }
    .status-good{
      background:var(--success);
      color:var(--success-text);
      border-color:#bbf7d0;
    }
    .status-low{
      background:var(--warn);
      color:var(--warn-text);
      border-color:#fed7aa;
    }
    .status-missing{
      background:var(--danger);
      color:var(--danger-text);
      border-color:#fecaca;
    }
    .status-maintenance{
      background:#fff7ed;
      color:#c2410c;
      border-color:#fdba74;
    }
    .hazard-note{
      margin-top:14px;
      color:#991b1b;
      font-size:14px;
      font-weight:700;
    }
    .notice{
      margin:18px 20px 0;
      padding:14px 16px;
      border-radius:14px;
      font-weight:700;
    }
    .notice.info{
      background:#eff6ff;
      color:#1d4ed8;
    }
    .notice.error{
      background:var(--danger);
      color:var(--danger-text);
    }
    .save-bar{
      display:flex;
      align-items:center;
      gap:14px;
      padding:20px;
      border-bottom:1px solid var(--line);
      flex-wrap:wrap;
    }
    .save-message{
      font-weight:700;
      color:#334155;
    }
    .empty-state{
      padding:26px;
      color:#64748b;
      font-size:18px;
    }
    .debug-panel{
      margin-top:18px;
    }
    .debug-grid{
      display:grid;
      grid-template-columns:repeat(2,minmax(0,1fr));
      gap:12px;
      padding:20px 26px;
      color:#334155;
    }
    @media (max-width: 720px){
      .page-header h1{font-size:42px}
      .subhead{font-size:18px}
      .panel-head h2{font-size:22px}
      .hero h3{font-size:24px}
      .hero p{font-size:18px}
      .item-grid{grid-template-columns:1fr}
      .location-name,.item-name{font-size:20px}
      .button{width:100%}
      .location-actions .button{width:auto}
    }
  `;
}

function clientScript_() {
  return `
    (function () {
      var bridgeAvailable = !!(window.google && google.script && google.script.run);

      if (BOOTSTRAP.hasSelection && BOOTSTRAP.mode === 'tech') {
        var saveButton = document.getElementById('saveButton');
        var saveMessage = document.getElementById('saveMessage');

        if (!bridgeAvailable) {
          if (saveMessage) {
            saveMessage.textContent = 'Interactive save is unavailable here. Open the deployed /exec web app URL.';
          }
          if (saveButton) {
            saveButton.disabled = true;
            saveButton.classList.add('disabled');
          }
          return;
        }

        if (saveButton) {
          saveButton.addEventListener('click', function () {
            try {
              var cards = Array.prototype.slice.call(document.querySelectorAll('.item-card[data-row]'));
              var items = cards.map(function (card) {
                var rowNumber = Number(card.getAttribute('data-row'));
                var qtyInput = card.querySelector('.qty-input');
                var statusSelect = card.querySelector('.status-select');

                return {
                  rowNumber: rowNumber,
                  qty: qtyInput ? Number(qtyInput.value) : 0,
                  status: statusSelect ? statusSelect.value : ''
                };
              });

              if (saveMessage) {
                saveMessage.textContent = 'Saving...';
              }

              google.script.run
                .withSuccessHandler(function (result) {
                  if (!result || !result.ok) {
                    if (saveMessage) {
                      saveMessage.textContent = (result && result.message) ? result.message : 'Save failed.';
                    }
                    return;
                  }

                  if (saveMessage) {
                    saveMessage.textContent = result.message || 'Saved.';
                  }
                })
                .withFailureHandler(function (error) {
                  if (saveMessage) {
                    saveMessage.textContent = (error && error.message) ? error.message : String(error);
                  }
                })
                .saveInventoryUpdates({
                  room: BOOTSTRAP.room,
                  location: BOOTSTRAP.location,
                  items: items
                });
            } catch (err) {
              if (saveMessage) {
                saveMessage.textContent = err && err.message ? err.message : String(err);
              }
            }
          });
        }
      }

      if (${APP.debug ? 'true' : 'false'}) {
        var panel = document.createElement('div');
        panel.style.marginTop = '16px';
        panel.style.fontSize = '12px';
        panel.style.color = '#64748b';
        panel.textContent = 'bridgeAvailable=' + bridgeAvailable;
        document.body.appendChild(panel);
      }
    })();
  `;
}
