var MATCH_BACKEND = {
  defaultStorageSpreadsheetId: '1uFP6fsP37jxplPzk3-M2ONsuXDxJ2P1b2Fr-gGMDeC4',
  historySheetName: 'MatchingHistory',
  queueSheetName: 'UnmatchedQueue',
  historyHeaders: [
    'runId',
    'savedAt',
    'sheetTitle',
    'previewOnly',
    'actorEmail',
    'actorName',
    'joinedCount',
    'leftCount',
    'attendingCount',
    'finalLeftCount',
    'missingCount',
    'currentUnmatchedCount',
    'resolvedPendingCount',
    'filesJson',
    'detailsJson'
  ],
  queueHeaders: [
    'queueKey',
    'status',
    'sheetTitle',
    'category',
    'name',
    'nameNormalized',
    'phone4',
    'label',
    'reason',
    'firstSeenAt',
    'lastSeenAt',
    'attemptCount',
    'openedRunId',
    'lastRunId',
    'resolvedAt',
    'resolvedRunId',
    'contextJson'
  ]
};

function doGet(e) {
  try {
    return respond_(handleRequest_(parseGetRequest_(e)));
  } catch (error) {
    return respondError_(error);
  }
}

function doPost(e) {
  try {
    return respond_(handleRequest_(parsePostRequest_(e)));
  } catch (error) {
    return respondError_(error);
  }
}

function handleRequest_(request) {
  verifyToken_(request.token);

  switch (request.action) {
    case 'health':
      return health_();
    case 'listSheetTitles':
      return listSheetTitles_(request.payload || {});
    case 'loadRosterRows':
      return loadRosterRows_(request.payload || {});
    case 'writeUpdates':
      return writeUpdates_(request.payload || {});
    case 'listPending':
      return listPending_(request.payload || {});
    case 'syncRun':
      return syncRun_(request.payload || {});
    default:
      throw new Error('Unsupported action: ' + request.action);
  }
}

function health_() {
  ensureStorageSheet_(MATCH_BACKEND.historySheetName, MATCH_BACKEND.historyHeaders);
  ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);

  return {
    enabled: true,
    ok: true
  };
}

function listSheetTitles_(payload) {
  var spreadsheet = openSpreadsheetById_(payload.spreadsheetId, 'attendance spreadsheet');
  return {
    enabled: true,
    titles: spreadsheet.getSheets().map(function(sheet) {
      return sheet.getName();
    })
  };
}

function loadRosterRows_(payload) {
  var spreadsheet = openSpreadsheetById_(payload.spreadsheetId, 'attendance spreadsheet');
  var sheetTitle = normalizeText_(payload.sheetTitle);
  var sheet = spreadsheet.getSheetByName(sheetTitle);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetTitle);
  }

  var startRow = Math.max(Number(payload.startRow || 2), 1);
  var columns = payload.columns || {};
  var targetRoleText = normalizeText_(payload.targetRoleText);
  var lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    return {
      enabled: true,
      rows: []
    };
  }

  var rowCount = lastRow - startRow + 1;
  var nameValues = readColumnValues_(sheet, startRow, columns.name, rowCount);
  var phoneValues = readColumnValues_(sheet, startRow, columns.phone, rowCount);
  var statusValues = readColumnValues_(sheet, startRow, columns.status, rowCount);
  var roleValues = readColumnValues_(sheet, startRow, columns.role, rowCount);
  var rows = [];

  for (var index = 0; index < rowCount; index += 1) {
    var name = normalizeText_(nameValues[index]);
    var phone = normalizeText_(phoneValues[index]);
    var statusText = normalizeText_(statusValues[index]);
    var roleText = normalizeText_(roleValues[index]);

    if (targetRoleText && roleText !== targetRoleText) continue;
    if (!name && !phone && !statusText) continue;

    rows.push({
      rowIndex: index,
      rowNumber: startRow + index,
      sheetTitle: sheetTitle,
      name: name,
      phone: phone,
      status: statusText,
      role: roleText,
      nameNormalized: normalizeName_(name)
    });
  }

  return {
    enabled: true,
    rows: rows
  };
}

function writeUpdates_(payload) {
  var spreadsheet = openSpreadsheetById_(payload.spreadsheetId, 'attendance spreadsheet');
  var updates = Array.isArray(payload.updates) ? payload.updates : [];
  var totalUpdatedCells = 0;

  updates.forEach(function(item) {
    if (!item || !item.range || !Array.isArray(item.values) || !item.values.length) return;

    var parsedRange = parseA1Range_(item.range);
    var sheet = spreadsheet.getSheetByName(parsedRange.sheetTitle);
    if (!sheet) {
      throw new Error('Sheet not found: ' + parsedRange.sheetTitle);
    }

    sheet.getRange(parsedRange.a1Notation).setValues(item.values);
    totalUpdatedCells += item.values.reduce(function(sum, row) {
      return sum + (Array.isArray(row) ? row.length : 0);
    }, 0);
  });

  return {
    enabled: true,
    totalUpdatedCells: totalUpdatedCells
  };
}

function listPending_(payload) {
  var queueSheet = ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);
  var rows = readSheetObjects_(queueSheet);
  var targetSheetTitle = normalizeText_(payload.sheetTitle);

  return {
    enabled: true,
    items: rows
      .filter(function(row) {
        if (String(row.status || '').trim() !== 'OPEN') return false;
        if (!targetSheetTitle) return true;
        return normalizeText_(row.sheetTitle) === targetSheetTitle;
      })
      .map(queueRowToItem_)
  };
}

function syncRun_(payload) {
  var nowIso = new Date().toISOString();
  var runId = Utilities.getUuid();
  var summary = payload.summary || {};
  var currentUnmatchedItems = Array.isArray(payload.currentUnmatchedItems) ? payload.currentUnmatchedItems : [];
  var resolvedPendingItems = Array.isArray(payload.resolvedPendingItems) ? payload.resolvedPendingItems : [];

  appendHistoryRow_(runId, nowIso, payload, summary, currentUnmatchedItems, resolvedPendingItems);

  var queueSheet = ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);
  var queueRows = readSheetObjects_(queueSheet);
  var queueIndex = buildQueueIndex_(queueRows);

  var resolvedCount = resolveQueueItems_(queueSheet, queueIndex, resolvedPendingItems, runId, nowIso);
  var openedCount = upsertQueueItems_(queueSheet, queueIndex, currentUnmatchedItems, runId, nowIso, payload.sheetTitle);
  var openCount = countOpenQueueItems_(queueIndex, payload.sheetTitle);

  return {
    enabled: true,
    synced: true,
    runId: runId,
    resolvedCount: resolvedCount,
    openedCount: openedCount,
    openCount: openCount
  };
}

function appendHistoryRow_(runId, nowIso, payload, summary, currentUnmatchedItems, resolvedPendingItems) {
  var historySheet = ensureStorageSheet_(MATCH_BACKEND.historySheetName, MATCH_BACKEND.historyHeaders);
  var details = {
    currentUnmatchedItems: currentUnmatchedItems,
    resolvedPendingItems: resolvedPendingItems
  };

  historySheet.appendRow([
    runId,
    nowIso,
    normalizeText_(payload.sheetTitle),
    payload.previewOnly ? 'TRUE' : 'FALSE',
    normalizeText_(payload.actorEmail),
    normalizeText_(payload.actorName),
    Number(summary.joinedCount || 0),
    Number(summary.leftCount || 0),
    Number(summary.attendingCount || 0),
    Number(summary.finalLeftCount || 0),
    Number(summary.missingCount || 0),
    Number(summary.currentUnmatchedCount || 0),
    Number(summary.resolvedPendingCount || 0),
    safeJson_(payload.files || []),
    safeJson_(details)
  ]);
}

function resolveQueueItems_(queueSheet, queueIndex, resolvedPendingItems, runId, nowIso) {
  var count = 0;

  resolvedPendingItems.forEach(function(rawItem) {
    var item = normalizeQueueItem_(rawItem);
    if (!item.queueKey) return;

    var existing = queueIndex[item.queueKey];
    if (!existing) return;

    if (String(existing.status || '').trim() !== 'RESOLVED') {
      count += 1;
    }

    existing.status = 'RESOLVED';
    existing.lastSeenAt = nowIso;
    existing.lastRunId = runId;
    existing.resolvedAt = nowIso;
    existing.resolvedRunId = runId;
    existing.contextJson = safeJson_(item);
    writeQueueRow_(queueSheet, existing);
  });

  return count;
}

function upsertQueueItems_(queueSheet, queueIndex, currentUnmatchedItems, runId, nowIso, defaultSheetTitle) {
  var count = 0;

  currentUnmatchedItems.forEach(function(rawItem) {
    var item = normalizeQueueItem_(rawItem, defaultSheetTitle);
    if (!item.queueKey) return;

    var existing = queueIndex[item.queueKey];
    if (existing) {
      existing.status = 'OPEN';
      existing.sheetTitle = item.sheetTitle || existing.sheetTitle;
      existing.category = item.category;
      existing.name = item.name;
      existing.nameNormalized = item.nameNormalized;
      existing.phone4 = item.phone4;
      existing.label = item.label;
      existing.reason = item.reason;
      existing.lastSeenAt = nowIso;
      existing.attemptCount = Number(existing.attemptCount || 0) + 1;
      existing.lastRunId = runId;
      existing.resolvedAt = '';
      existing.resolvedRunId = '';
      existing.contextJson = safeJson_(item);
      writeQueueRow_(queueSheet, existing);
      count += 1;
      return;
    }

    var newRow = {
      rowNumber: queueSheet.getLastRow() + 1,
      queueKey: item.queueKey,
      status: 'OPEN',
      sheetTitle: item.sheetTitle,
      category: item.category,
      name: item.name,
      nameNormalized: item.nameNormalized,
      phone4: item.phone4,
      label: item.label,
      reason: item.reason,
      firstSeenAt: nowIso,
      lastSeenAt: nowIso,
      attemptCount: 1,
      openedRunId: runId,
      lastRunId: runId,
      resolvedAt: '',
      resolvedRunId: '',
      contextJson: safeJson_(item)
    };

    queueSheet.appendRow(queueObjectToRow_(newRow));
    queueIndex[item.queueKey] = newRow;
    count += 1;
  });

  return count;
}

function countOpenQueueItems_(queueIndex, sheetTitle) {
  var targetSheetTitle = normalizeText_(sheetTitle);

  return Object.keys(queueIndex).reduce(function(total, key) {
    var item = queueIndex[key];
    if (!item || String(item.status || '').trim() !== 'OPEN') return total;
    if (targetSheetTitle && normalizeText_(item.sheetTitle) !== targetSheetTitle) return total;
    return total + 1;
  }, 0);
}

function buildQueueIndex_(rows) {
  return rows.reduce(function(map, row) {
    if (!row.queueKey) return map;
    map[row.queueKey] = row;
    return map;
  }, {});
}

function writeQueueRow_(sheet, item) {
  sheet
    .getRange(item.rowNumber, 1, 1, MATCH_BACKEND.queueHeaders.length)
    .setValues([queueObjectToRow_(item)]);
}

function queueObjectToRow_(item) {
  return [
    item.queueKey || '',
    item.status || 'OPEN',
    item.sheetTitle || '',
    item.category || '',
    item.name || '',
    item.nameNormalized || '',
    item.phone4 || '',
    item.label || '',
    item.reason || '',
    item.firstSeenAt || '',
    item.lastSeenAt || '',
    Number(item.attemptCount || 0),
    item.openedRunId || '',
    item.lastRunId || '',
    item.resolvedAt || '',
    item.resolvedRunId || '',
    item.contextJson || ''
  ];
}

function queueRowToItem_(row) {
  return {
    queueKey: row.queueKey,
    sheetTitle: row.sheetTitle,
    category: row.category,
    name: row.name,
    nameNormalized: row.nameNormalized,
    phone4: row.phone4,
    label: row.label,
    reason: row.reason
  };
}

function normalizeQueueItem_(rawItem, defaultSheetTitle) {
  var name = normalizeText_(rawItem && rawItem.name);
  var nameNormalized = normalizeName_(rawItem && (rawItem.nameNormalized || rawItem.name) || '');
  var phone4 = last4Digits_(rawItem && (rawItem.phone4 || rawItem.phone) || '');
  var category = normalizeText_(rawItem && rawItem.category) || 'outside-roster';
  var sheetTitle = normalizeText_(rawItem && rawItem.sheetTitle) || normalizeText_(defaultSheetTitle);

  return {
    queueKey: buildQueueKey_(sheetTitle, nameNormalized, phone4, category),
    sheetTitle: sheetTitle,
    category: category,
    name: name || nameNormalized,
    nameNormalized: nameNormalized,
    phone4: phone4,
    label: normalizeText_(rawItem && rawItem.label) || formatLabel_(name || nameNormalized, phone4),
    reason: normalizeText_(rawItem && rawItem.reason)
  };
}

function ensureStorageSheet_(sheetName, headers) {
  var spreadsheet = getStorageSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    var needsUpdate = headers.some(function(header, index) {
      return String(currentHeaders[index] || '').trim() !== header;
    });
    if (needsUpdate) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  return sheet;
}

function readSheetObjects_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn === 0) return [];

  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

  return values.map(function(row, index) {
    var item = { rowNumber: index + 2 };
    headers.forEach(function(header, headerIndex) {
      item[String(header || '').trim()] = row[headerIndex];
    });
    return item;
  });
}

function getStorageSpreadsheet_() {
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('MATCH_BACKEND_SPREADSHEET_ID');
  if (!spreadsheetId) {
    spreadsheetId = MATCH_BACKEND.defaultStorageSpreadsheetId;
  }

  return openSpreadsheetById_(spreadsheetId, 'storage spreadsheet');
}

function openSpreadsheetById_(spreadsheetId, label) {
  var normalizedId = normalizeText_(spreadsheetId);
  if (!normalizedId) {
    throw new Error('Missing spreadsheet ID for ' + normalizeText_(label || 'spreadsheet') + '.');
  }

  return SpreadsheetApp.openById(normalizedId);
}

function readColumnValues_(sheet, startRow, columnLetter, rowCount) {
  var columnNumber = columnLetterToNumber_(columnLetter);
  if (!columnNumber) {
    throw new Error('Invalid column: ' + columnLetter);
  }

  return sheet
    .getRange(startRow, columnNumber, rowCount, 1)
    .getDisplayValues()
    .map(function(row) {
      return row[0];
    });
}

function columnLetterToNumber_(value) {
  var text = normalizeText_(value).toUpperCase();
  if (!text) return 0;

  var total = 0;
  for (var index = 0; index < text.length; index += 1) {
    var code = text.charCodeAt(index);
    if (code < 65 || code > 90) {
      throw new Error('Invalid column letter: ' + value);
    }
    total = (total * 26) + (code - 64);
  }

  return total;
}

function parseA1Range_(value) {
  var raw = normalizeText_(value);
  var bangIndex = raw.lastIndexOf('!');
  if (bangIndex < 0) {
    throw new Error('Invalid range: ' + raw);
  }

  var sheetTitle = raw.slice(0, bangIndex);
  if (sheetTitle.charAt(0) === '\'' && sheetTitle.charAt(sheetTitle.length - 1) === '\'') {
    sheetTitle = sheetTitle.slice(1, -1).replace(/''/g, '\'');
  }

  return {
    sheetTitle: sheetTitle,
    a1Notation: raw.slice(bangIndex + 1)
  };
}

function verifyToken_(token) {
  var expected = PropertiesService.getScriptProperties().getProperty('MATCH_BACKEND_TOKEN');
  if (!expected) {
    throw new Error('MATCH_BACKEND_TOKEN script property is missing.');
  }
  if (normalizeText_(token) !== expected) {
    throw new Error('Invalid backend token.');
  }
}

function parseGetRequest_(e) {
  var parameter = (e && e.parameter) || {};
  return {
    token: parameter.token,
    action: normalizeText_(parameter.action) || 'health',
    payload: {
      sheetTitle: parameter.sheetTitle,
      spreadsheetId: parameter.spreadsheetId
    }
  };
}

function parsePostRequest_(e) {
  var payload = safeParseJson_((e && e.postData && e.postData.contents) || '') || {};
  return {
    token: payload.token,
    action: normalizeText_(payload.action) || 'health',
    payload: payload.payload || {}
  };
}

function respond_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function respondError_(error) {
  return respond_({
    enabled: false,
    error: error && error.message ? error.message : 'Apps Script backend error'
  });
}

function safeParseJson_(text) {
  try {
    return JSON.parse(text);
  } catch (error) {
    return null;
  }
}

function safeJson_(value) {
  try {
    return JSON.stringify(value);
  } catch (error) {
    return '[]';
  }
}

function normalizeText_(value) {
  return String(value || '').trim();
}

function normalizeName_(value) {
  return normalizeText_(value)
    .replace(/[\s]*[\(\[\{][^)\]\}]*[\)\]\}]\s*$/, '')
    .replace(/[/_\-\s]+$/, '')
    .replace(/\s+/g, ' ');
}

function digitsOnly_(value) {
  return String(value || '').replace(/\D+/g, '');
}

function last4Digits_(value) {
  return digitsOnly_(value).slice(-4);
}

function buildQueueKey_(sheetTitle, nameNormalized, phone4, category) {
  return [
    normalizeText_(sheetTitle),
    normalizeText_(nameNormalized),
    normalizeText_(phone4),
    normalizeText_(category)
  ].join('::');
}

function formatLabel_(name, phone4) {
  return phone4 ? normalizeText_(name) + '/' + phone4 : normalizeText_(name);
}
