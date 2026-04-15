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
    'resolutionType',
    'resolutionLabel',
    'resolutionTargetRow',
    'resolutionTargetName',
    'resolutionTargetPhone',
    'handledBy',
    'handledAt',
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
    case 'applyManualAction':
      return applyManualAction_(request.payload || {});
    case 'getStorageOverview':
      return getStorageOverview_();
    case 'getStoredSheetData':
      return getStoredSheetData_(request.payload || {});
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
  var startRow = Math.max(Number(payload.startRow || 2), 1);
  var columns = payload.columns || {};
  var targetRoleText = normalizeText_(payload.targetRoleText);
  var allowRoleFallback = payload.allowRoleFallback !== false;
  var lastRow;
  var rowCount;
  var nameValues;
  var phoneValues;
  var statusValues;
  var roleValues;
  var rows = [];
  var filteredRows;
  var availableRoles;

  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetTitle);
  }

  lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    return {
      enabled: true,
      rows: [],
      roleFilterApplied: false,
      fallbackUsed: false,
      matchedRoleCount: 0,
      availableRoles: []
    };
  }

  rowCount = lastRow - startRow + 1;
  nameValues = readColumnValues_(sheet, startRow, columns.name, rowCount);
  phoneValues = readColumnValues_(sheet, startRow, columns.phone, rowCount);
  statusValues = readColumnValues_(sheet, startRow, columns.status, rowCount);
  roleValues = readColumnValues_(sheet, startRow, columns.role, rowCount);

  for (var index = 0; index < rowCount; index += 1) {
    var name = normalizeText_(nameValues[index]);
    var phone = normalizeText_(phoneValues[index]);
    var statusText = normalizeText_(statusValues[index]);
    var roleText = normalizeText_(roleValues[index]);

    if (!name && !phone && !statusText && !roleText) continue;

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

  filteredRows = targetRoleText ? rows.filter(function(row) {
    return normalizeText_(row.role) === targetRoleText;
  }) : rows.slice();
  availableRoles = uniqueValues_(rows.map(function(row) {
    return row.role;
  }));

  if (targetRoleText && !filteredRows.length && allowRoleFallback) {
    return {
      enabled: true,
      rows: rows,
      roleFilterApplied: false,
      fallbackUsed: true,
      matchedRoleCount: 0,
      availableRoles: availableRoles
    };
  }

  return {
    enabled: true,
    rows: filteredRows,
    roleFilterApplied: !!targetRoleText,
    fallbackUsed: false,
    matchedRoleCount: filteredRows.length,
    availableRoles: availableRoles
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
  var openItems = [];
  var manualRules = [];

  rows.forEach(function(row) {
    if (targetSheetTitle && normalizeText_(row.sheetTitle) !== targetSheetTitle) return;

    if (String(row.status || '').trim() === 'OPEN') {
      openItems.push(queueRowToItem_(row));
      return;
    }

    if (isManualRuleStatus_(row.status)) {
      manualRules.push(queueRowToManualRule_(row));
    }
  });

  return {
    enabled: true,
    items: openItems,
    manualRules: manualRules
  };
}

function syncRun_(payload) {
  var nowIso = new Date().toISOString();
  var runId = Utilities.getUuid();
  var summary = payload.summary || {};
  var currentUnmatchedItems = Array.isArray(payload.currentUnmatchedItems) ? payload.currentUnmatchedItems : [];
  var resolvedPendingItems = Array.isArray(payload.resolvedPendingItems) ? payload.resolvedPendingItems : [];
  var queueSheet = ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);
  var queueRows = readSheetObjects_(queueSheet);
  var queueIndex = buildQueueIndex_(queueRows);
  var resolvedCount;
  var openedCount;
  var openCount;

  appendHistoryRow_(runId, nowIso, payload, summary, currentUnmatchedItems, resolvedPendingItems);

  resolvedCount = resolveQueueItems_(queueSheet, queueIndex, resolvedPendingItems, runId, nowIso);
  openedCount = upsertQueueItems_(queueSheet, queueIndex, currentUnmatchedItems, runId, nowIso, payload.sheetTitle);
  openCount = countOpenQueueItems_(queueIndex, payload.sheetTitle);

  return {
    enabled: true,
    synced: true,
    runId: runId,
    resolvedCount: resolvedCount,
    openedCount: openedCount,
    openCount: openCount
  };
}

function applyManualAction_(payload) {
  var queueSheet = ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);
  var queueRows = readSheetObjects_(queueSheet);
  var queueIndex = buildQueueIndex_(queueRows);
  var nowIso = new Date().toISOString();
  var item = normalizeQueueItem_(payload.item, payload.sheetTitle);
  var existing = queueIndex[item.queueKey];
  var targetRow = payload.targetRow || {};

  if (!item.queueKey) {
    throw new Error('Invalid queue item.');
  }

  if (!existing) {
    existing = createQueueRowObject_(item);
    existing.rowNumber = queueSheet.getLastRow() + 1;
    existing.firstSeenAt = nowIso;
    existing.lastSeenAt = nowIso;
    queueSheet.appendRow(queueObjectToRow_(existing));
    queueIndex[item.queueKey] = existing;
  }

  existing.sheetTitle = item.sheetTitle;
  existing.category = item.category;
  existing.name = item.name;
  existing.nameNormalized = item.nameNormalized;
  existing.phone4 = item.phone4;
  existing.label = item.label;
  existing.reason = item.reason;
  existing.lastSeenAt = nowIso;
  existing.contextJson = safeJson_(item);
  existing.handledBy = normalizeText_(payload.actorName || payload.actorEmail);
  existing.handledAt = nowIso;
  existing.resolvedAt = nowIso;

  if (normalizeText_(payload.actionType) === 'match-student') {
    if (!targetRow || !targetRow.rowNumber) {
      throw new Error('Target row is required for match-student.');
    }
    existing.status = 'MANUAL_MATCH';
    existing.resolutionType = 'MATCH_ROSTER';
    existing.resolutionLabel = formatLabel_(targetRow.name, last4Digits_(targetRow.phone));
    existing.resolutionTargetRow = Number(targetRow.rowNumber || 0);
    existing.resolutionTargetName = normalizeText_(targetRow.name);
    existing.resolutionTargetPhone = normalizeText_(targetRow.phone);
  } else if (normalizeText_(payload.actionType) === 'exclude-staff') {
    existing.status = 'EXCLUDED_STAFF';
    existing.resolutionType = 'EXCLUDE_STAFF';
    existing.resolutionLabel = '코칭스태프 제외';
    existing.resolutionTargetRow = '';
    existing.resolutionTargetName = '';
    existing.resolutionTargetPhone = '';
  } else {
    throw new Error('Unsupported manual action.');
  }

  writeQueueRow_(queueSheet, existing);

  return {
    enabled: true,
    queueItem: queueRowToEntry_(existing),
    manualRule: queueRowToManualRule_(existing)
  };
}

function getStorageOverview_() {
  var historyRows = readSheetObjects_(ensureStorageSheet_(MATCH_BACKEND.historySheetName, MATCH_BACKEND.historyHeaders));
  var queueRows = readSheetObjects_(ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders));
  var map = {};

  historyRows.forEach(function(row) {
    var key = normalizeText_(row.sheetTitle);
    if (!key) return;
    ensureOverviewItem_(map, key).runCount += 1;
    ensureOverviewItem_(map, key).lastSavedAt = maxIso_(ensureOverviewItem_(map, key).lastSavedAt, normalizeText_(row.savedAt));
  });

  queueRows.forEach(function(row) {
    var key = normalizeText_(row.sheetTitle);
    var item;
    if (!key) return;
    item = ensureOverviewItem_(map, key);

    if (String(row.status || '').trim() === 'OPEN') {
      item.openCount += 1;
    }
    if (isManualRuleStatus_(row.status)) {
      item.manualRuleCount += 1;
    }
  });

  return {
    enabled: true,
    sheets: objectValues_(map).sort(function(left, right) {
      return String(left.sheetTitle || '').localeCompare(String(right.sheetTitle || ''), 'ko');
    })
  };
}

function getStoredSheetData_(payload) {
  var sheetTitle = normalizeText_(payload.sheetTitle);
  var historyRows = readSheetObjects_(ensureStorageSheet_(MATCH_BACKEND.historySheetName, MATCH_BACKEND.historyHeaders));
  var queueRows = readSheetObjects_(ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders));
  var filteredHistory = historyRows.filter(function(row) {
    return normalizeText_(row.sheetTitle) === sheetTitle;
  });
  var filteredQueue = queueRows.filter(function(row) {
    return normalizeText_(row.sheetTitle) === sheetTitle;
  });

  filteredHistory.sort(function(left, right) {
    return normalizeText_(right.savedAt).localeCompare(normalizeText_(left.savedAt));
  });
  filteredQueue.sort(function(left, right) {
    return normalizeText_(right.lastSeenAt).localeCompare(normalizeText_(left.lastSeenAt));
  });

  return {
    enabled: true,
    summary: {
      sheetTitle: sheetTitle,
      runCount: filteredHistory.length,
      openCount: filteredQueue.filter(function(row) { return String(row.status || '').trim() === 'OPEN'; }).length,
      manualRuleCount: filteredQueue.filter(function(row) { return isManualRuleStatus_(row.status); }).length,
      lastSavedAt: filteredHistory.length ? normalizeText_(filteredHistory[0].savedAt) : ''
    },
    runs: filteredHistory.map(historyRowToSummary_),
    queueItems: filteredQueue.map(queueRowToEntry_)
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
    var existing = queueIndex[item.queueKey];
    if (!item.queueKey || !existing || isManualRuleStatus_(existing.status)) return;

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
    var existing;
    if (!item.queueKey) return;

    existing = queueIndex[item.queueKey];
    if (existing) {
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
      existing.contextJson = safeJson_(item);

      if (!isManualRuleStatus_(existing.status)) {
        existing.status = 'OPEN';
        existing.resolvedAt = '';
        existing.resolvedRunId = '';
        existing.resolutionType = '';
        existing.resolutionLabel = '';
        existing.resolutionTargetRow = '';
        existing.resolutionTargetName = '';
        existing.resolutionTargetPhone = '';
        existing.handledBy = '';
        existing.handledAt = '';
        count += 1;
      }

      writeQueueRow_(queueSheet, existing);
      return;
    }

    existing = createQueueRowObject_(item);
    existing.rowNumber = queueSheet.getLastRow() + 1;
    existing.status = 'OPEN';
    existing.firstSeenAt = nowIso;
    existing.lastSeenAt = nowIso;
    existing.attemptCount = 1;
    existing.openedRunId = runId;
    existing.lastRunId = runId;
    existing.contextJson = safeJson_(item);
    queueSheet.appendRow(queueObjectToRow_(existing));
    queueIndex[item.queueKey] = existing;
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

function ensureOverviewItem_(map, sheetTitle) {
  if (!map[sheetTitle]) {
    map[sheetTitle] = {
      sheetTitle: sheetTitle,
      runCount: 0,
      openCount: 0,
      manualRuleCount: 0,
      lastSavedAt: ''
    };
  }
  return map[sheetTitle];
}

function buildQueueIndex_(rows) {
  return rows.reduce(function(map, row) {
    if (!row.queueKey) return map;
    map[row.queueKey] = row;
    return map;
  }, {});
}

function createQueueRowObject_(item) {
  return {
    queueKey: item.queueKey,
    status: 'OPEN',
    sheetTitle: item.sheetTitle,
    category: item.category,
    name: item.name,
    nameNormalized: item.nameNormalized,
    phone4: item.phone4,
    label: item.label,
    reason: item.reason,
    firstSeenAt: '',
    lastSeenAt: '',
    attemptCount: 0,
    openedRunId: '',
    lastRunId: '',
    resolvedAt: '',
    resolvedRunId: '',
    resolutionType: '',
    resolutionLabel: '',
    resolutionTargetRow: '',
    resolutionTargetName: '',
    resolutionTargetPhone: '',
    handledBy: '',
    handledAt: '',
    contextJson: safeJson_(item)
  };
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
    item.resolutionType || '',
    item.resolutionLabel || '',
    item.resolutionTargetRow || '',
    item.resolutionTargetName || '',
    item.resolutionTargetPhone || '',
    item.handledBy || '',
    item.handledAt || '',
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

function queueRowToManualRule_(row) {
  return {
    queueKey: row.queueKey,
    sheetTitle: row.sheetTitle,
    status: row.status,
    actionType: String(row.status || '').trim() === 'MANUAL_MATCH' ? 'match-student' : 'exclude-staff',
    resolutionLabel: row.resolutionLabel,
    targetRowNumber: Number(row.resolutionTargetRow || 0),
    targetName: row.resolutionTargetName,
    targetPhone4: last4Digits_(row.resolutionTargetPhone),
    reason: row.reason
  };
}

function queueRowToEntry_(row) {
  return {
    queueKey: row.queueKey,
    sheetTitle: row.sheetTitle,
    status: row.status,
    category: row.category,
    label: row.label,
    reason: row.reason,
    attemptCount: Number(row.attemptCount || 0),
    firstSeenAt: normalizeText_(row.firstSeenAt),
    lastSeenAt: normalizeText_(row.lastSeenAt),
    resolutionLabel: normalizeText_(row.resolutionLabel),
    handledBy: normalizeText_(row.handledBy),
    handledAt: normalizeText_(row.handledAt)
  };
}

function historyRowToSummary_(row) {
  return {
    runId: row.runId,
    savedAt: normalizeText_(row.savedAt),
    previewOnly: String(row.previewOnly || '').trim() === 'TRUE',
    actorEmail: normalizeText_(row.actorEmail),
    actorName: normalizeText_(row.actorName),
    joinedCount: Number(row.joinedCount || 0),
    leftCount: Number(row.leftCount || 0),
    attendingCount: Number(row.attendingCount || 0),
    missingCount: Number(row.missingCount || 0),
    currentUnmatchedCount: Number(row.currentUnmatchedCount || 0)
  };
}

function isManualRuleStatus_(status) {
  var value = String(status || '').trim();
  return value === 'MANUAL_MATCH' || value === 'EXCLUDED_STAFF';
}

function ensureStorageSheet_(sheetName, headers) {
  var spreadsheet = getStorageSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var currentHeaders;
  var needsUpdate;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }

  currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  needsUpdate = headers.some(function(header, index) {
    return String(currentHeaders[index] || '').trim() !== header;
  });
  if (needsUpdate) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sheet;
}

function readSheetObjects_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = Math.max(sheet.getLastColumn(), MATCH_BACKEND.queueHeaders.length);
  var headers;
  var values;

  if (lastRow <= 1 || lastColumn === 0) return [];

  headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

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

  return sheet.getRange(startRow, columnNumber, rowCount, 1).getDisplayValues().map(function(row) {
    return row[0];
  });
}

function columnLetterToNumber_(value) {
  var text = normalizeText_(value).toUpperCase();
  var total = 0;

  if (!text) return 0;

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
  var sheetTitle;

  if (bangIndex < 0) {
    throw new Error('Invalid range: ' + raw);
  }

  sheetTitle = raw.slice(0, bangIndex);
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
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function respondError_(error) {
  return respond_({
    enabled: false,
    error: error && error.message ? error.message : 'Apps Script backend error'
  });
}

function normalizeQueueItem_(rawItem, defaultSheetTitle) {
  var name = normalizeText_(rawItem && rawItem.name);
  var nameNormalized = normalizeName_((rawItem && (rawItem.nameNormalized || rawItem.name)) || '');
  var phone4 = last4Digits_((rawItem && (rawItem.phone4 || rawItem.phone)) || '');
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

function uniqueValues_(values) {
  var seen = {};
  var result = [];
  (values || []).forEach(function(value) {
    var key = normalizeText_(value);
    if (!key || seen[key]) return;
    seen[key] = true;
    result.push(key);
  });
  return result;
}

function objectValues_(map) {
  return Object.keys(map || {}).map(function(key) {
    return map[key];
  });
}

function maxIso_(left, right) {
  return normalizeText_(left) > normalizeText_(right) ? normalizeText_(left) : normalizeText_(right);
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
