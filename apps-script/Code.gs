var MATCH_BACKEND = {
  defaultStorageSpreadsheetId: '1uFP6fsP37jxplPzk3-M2ONsuXDxJ2P1b2Fr-gGMDeC4',
  snapshotSheetName: 'LatestSheetState',
  queueSheetName: 'UnmatchedQueue',
  summarySheetName: 'SheetSummary',
  snapshotChunkSize: 40000,
  cacheTtlSeconds: {
    sheetTitles: 300,
    storageOverview: 45,
    storedSheetData: 20
  },
  snapshotHeaders: [
    'sheetTitle',
    'savedAt',
    'previewOnly',
    'actorEmail',
    'actorName',
    'partIndex',
    'partCount',
    'payloadChunk'
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
  ],
  summaryHeaders: [
    'sheetTitle',
    'snapshotReady',
    'openCount',
    'manualRuleCount',
    'lastSavedAt',
    'lastSnapshotAt'
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
    case 'saveSnapshot':
      return saveSnapshot_(request.payload || {});
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
  ensureStorageSheet_(MATCH_BACKEND.snapshotSheetName, MATCH_BACKEND.snapshotHeaders);
  ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);
  ensureStorageSheet_(MATCH_BACKEND.summarySheetName, MATCH_BACKEND.summaryHeaders);
  return {
    enabled: true,
    ok: true
  };
}

function listSheetTitles_(payload) {
  var cacheKey = buildCacheKey_('sheetTitles', payload.spreadsheetId);
  var cached = getCachedJson_(cacheKey);
  if (cached) {
    return cached;
  }
  var spreadsheet = openSpreadsheetById_(payload.spreadsheetId, 'attendance spreadsheet');
  var response = {
    enabled: true,
    titles: spreadsheet.getSheets().map(function(sheet) {
      return sheet.getName();
    })
  };
  putCachedJson_(cacheKey, response, MATCH_BACKEND.cacheTtlSeconds.sheetTitles);
  return response;
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
  var dataset;
  var nameColumn = columnLetterToNumber_(columns.name);
  var phoneColumn = columnLetterToNumber_(columns.phone);
  var statusColumn = columnLetterToNumber_(columns.status);
  var roleColumn = columnLetterToNumber_(columns.role);
  var minColumn;
  var maxColumn;
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
  minColumn = Math.min(nameColumn, phoneColumn, statusColumn, roleColumn);
  maxColumn = Math.max(nameColumn, phoneColumn, statusColumn, roleColumn);
  dataset = sheet.getRange(startRow, minColumn, rowCount, (maxColumn - minColumn) + 1).getDisplayValues();

  for (var index = 0; index < rowCount; index += 1) {
    var row = dataset[index] || [];
    var name = normalizeText_(row[nameColumn - minColumn]);
    var phone = normalizeText_(row[phoneColumn - minColumn]);
    var statusText = normalizeText_(row[statusColumn - minColumn]);
    var roleText = normalizeText_(row[roleColumn - minColumn]);

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
  var plan = buildWritePlan_(updates);
  var totalUpdatedCells = 0;
  var sheetCache = {};

  plan.batched.forEach(function(group) {
    var sheet = getSheetFromCache_(spreadsheet, sheetCache, group.sheetTitle);
    if (!sheet) {
      throw new Error('Sheet not found: ' + group.sheetTitle);
    }

    sheet.getRange(group.startRow, group.columnNumber, group.values.length, group.width).setValues(group.values);
    totalUpdatedCells += group.values.length * group.width;
  });

  plan.single.forEach(function(item) {
    if (!item || !item.range || !Array.isArray(item.values) || !item.values.length) return;

    var parsedRange = parseA1Range_(item.range);
    var sheet = getSheetFromCache_(spreadsheet, sheetCache, parsedRange.sheetTitle);
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
  var targetSheetTitle = normalizeText_(payload.sheetTitle);
  var rows = targetSheetTitle ? readSheetObjectsByMatch_(queueSheet, 3, targetSheetTitle) : readSheetObjects_(queueSheet);
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
  var summary = payload.summary || {};
  var currentUnmatchedItems = Array.isArray(payload.currentUnmatchedItems) ? payload.currentUnmatchedItems : [];
  var resolvedPendingItems = Array.isArray(payload.resolvedPendingItems) ? payload.resolvedPendingItems : [];
  var queueSheet = ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);
  var normalizedSheetTitle = normalizeText_(payload.sheetTitle);
  var queueState;
  var snapshotSavedAt;
  var summaryRow;
  var resolvedCount;
  var openedCount;
  var queueCounts;

  if (!normalizedSheetTitle) {
    throw new Error('Sheet title is required.');
  }

  queueState = readQueueStateForSheet_(queueSheet, normalizedSheetTitle);
  resolvedCount = resolveQueueItemsInState_(queueState, resolvedPendingItems, nowIso);
  openedCount = upsertQueueItemsInState_(queueState, currentUnmatchedItems, nowIso, payload.sheetTitle);
  flushQueueState_(queueSheet, queueState);
  queueCounts = buildQueueCounts_(queueState.rows);
  snapshotSavedAt = saveLatestSnapshot_(normalizedSheetTitle, payload.snapshot, {
    savedAt: normalizeText_(payload.executedAt) || nowIso,
    actorEmail: normalizeText_(payload.actorEmail),
    actorName: normalizeText_(payload.actorName),
    previewOnly: !!payload.previewOnly,
    summary: summary,
    metrics: {
      openCount: Number(queueCounts.openCount || 0),
      manualRuleCount: Number(queueCounts.manualRuleCount || 0)
    }
  });
  summaryRow = writeSummarySnapshotState_(normalizedSheetTitle, queueCounts, {
    lastSavedAt: nowIso,
    lastSnapshotAt: snapshotSavedAt
  });
  invalidateStorageCaches_(normalizedSheetTitle);

  return {
    enabled: true,
    synced: true,
    resolvedCount: resolvedCount,
    openedCount: openedCount,
    openCount: Number(summaryRow.openCount || 0),
    manualRuleCount: Number(summaryRow.manualRuleCount || 0),
    savedAt: snapshotSavedAt
  };
}

function saveSnapshot_(payload) {
  var sheetTitle = normalizeText_(payload.sheetTitle);
  var nowIso = normalizeText_(payload.savedAt) || new Date().toISOString();
  var queueSheet = ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);
  var queueCounts = buildQueueCounts_(readSheetObjectsByMatch_(queueSheet, 3, sheetTitle));
  var snapshotSavedAt;
  var summaryRow;

  if (!sheetTitle) {
    throw new Error('Sheet title is required.');
  }

  snapshotSavedAt = saveLatestSnapshot_(sheetTitle, payload.snapshot, {
    savedAt: nowIso,
    actorEmail: normalizeText_(payload.actorEmail),
    actorName: normalizeText_(payload.actorName),
    previewOnly: !!payload.previewOnly,
    metrics: {
      openCount: Number(queueCounts.openCount || 0),
      manualRuleCount: Number(queueCounts.manualRuleCount || 0)
    }
  });
  summaryRow = writeSummarySnapshotState_(sheetTitle, queueCounts, {
    lastSavedAt: nowIso,
    lastSnapshotAt: snapshotSavedAt
  });
  invalidateStorageCaches_(sheetTitle);

  return {
    enabled: true,
    saved: true,
    summary: summaryRowToOverview_(summaryRow)
  };
}

function applyManualAction_(payload) {
  var queueSheet = ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);
  var nowIso = new Date().toISOString();
  var item = normalizeQueueItem_(payload.item, payload.sheetTitle);
  var normalizedSheetTitle = normalizeText_(item.sheetTitle || payload.sheetTitle);
  var queueState = readQueueStateForSheet_(queueSheet, normalizedSheetTitle);
  var existing = queueState.byKey[item.queueKey];
  var targetRow = payload.targetRow || {};
  var summaryRow;

  if (!item.queueKey) {
    throw new Error('Invalid queue item.');
  }

  if (!existing) {
    existing = createQueueRowObject_(item);
    existing.firstSeenAt = nowIso;
    existing.lastSeenAt = nowIso;
    queueState.rows.push(existing);
    queueState.byKey[item.queueKey] = existing;
    queueState.newRows.push(existing);
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

  markQueueStateDirty_(queueState, existing);
  flushQueueState_(queueSheet, queueState);
  summaryRow = writeSummarySnapshotState_(normalizedSheetTitle, buildQueueCounts_(queueState.rows), {
    lastSavedAt: nowIso
  });
  invalidateStorageCaches_(normalizedSheetTitle);

  return {
    enabled: true,
    queueItem: queueRowToEntry_(existing),
    manualRule: queueRowToManualRule_(existing),
    summary: summaryRowToOverview_(summaryRow)
  };
}

function getStorageOverview_() {
  var cacheKey = buildCacheKey_('storageOverview', 'all');
  var cached = getCachedJson_(cacheKey);
  var summarySheet;
  var rows;
  var response;

  if (cached) {
    return cached;
  }

  summarySheet = ensureStorageSheet_(MATCH_BACKEND.summarySheetName, MATCH_BACKEND.summaryHeaders);
  if (summarySheet.getLastRow() <= 1) {
    rebuildSummarySheet_();
  }

  rows = readSheetObjects_(summarySheet);
  response = {
    enabled: true,
    sheets: rows
      .filter(function(row) { return normalizeText_(row.sheetTitle); })
      .map(summaryRowToOverview_)
      .sort(function(left, right) {
        return String(left.sheetTitle || '').localeCompare(String(right.sheetTitle || ''), 'ko');
      })
  };
  putCachedJson_(cacheKey, response, MATCH_BACKEND.cacheTtlSeconds.storageOverview);
  return response;
}

function getStoredSheetData_(payload) {
  var sheetTitle = normalizeText_(payload.sheetTitle);
  var cacheKey = buildCacheKey_('storedSheetData', sheetTitle);
  var cached = getCachedJson_(cacheKey);
  var queueSheet;
  var queueCounts;
  var summaryRow;
  var snapshot;
  var response;

  if (cached) {
    return cached;
  }

  if (!sheetTitle) {
    throw new Error('Sheet title is required.');
  }

  queueSheet = ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);
  queueCounts = buildQueueCounts_(readSheetObjectsByMatch_(queueSheet, 3, sheetTitle));
  summaryRow = findSummaryRowBySheetTitle_(ensureStorageSheet_(MATCH_BACKEND.summarySheetName, MATCH_BACKEND.summaryHeaders), sheetTitle);
  snapshot = loadLatestSnapshot_(sheetTitle);
  if (!summaryRow || Number(summaryRow.openCount || 0) !== queueCounts.openCount || Number(summaryRow.manualRuleCount || 0) !== queueCounts.manualRuleCount || !!summaryRowToOverview_(summaryRow).snapshotReady !== !!snapshot) {
    summaryRow = writeSummarySnapshotState_(sheetTitle, queueCounts, {
      lastSavedAt: normalizeText_(summaryRow && summaryRow.lastSavedAt) || normalizeText_(snapshot && snapshot.savedAt),
      lastSnapshotAt: snapshot ? normalizeText_(snapshot.savedAt) : normalizeText_(summaryRow && summaryRow.lastSnapshotAt)
    });
  }

  response = {
    enabled: true,
    summary: summaryRowToOverview_(summaryRow),
    snapshot: snapshot
  };
  putCachedJson_(cacheKey, response, MATCH_BACKEND.cacheTtlSeconds.storedSheetData);
  return response;
}

function resolveQueueItems_(queueSheet, resolvedPendingItems, nowIso) {
  var count = 0;

  resolvedPendingItems.forEach(function(rawItem) {
    var item = normalizeQueueItem_(rawItem);
    var existing = findQueueRowByQueueKey_(queueSheet, item.queueKey);
    if (!item.queueKey || !existing || isManualRuleStatus_(existing.status)) return;

    if (String(existing.status || '').trim() !== 'RESOLVED') {
      count += 1;
    }

    existing.status = 'RESOLVED';
    existing.lastSeenAt = nowIso;
    existing.resolvedAt = nowIso;
    existing.contextJson = safeJson_(item);
    writeQueueRow_(queueSheet, existing);
  });

  return count;
}

function upsertQueueItems_(queueSheet, currentUnmatchedItems, nowIso, defaultSheetTitle) {
  var count = 0;

  currentUnmatchedItems.forEach(function(rawItem) {
    var item = normalizeQueueItem_(rawItem, defaultSheetTitle);
    var existing;
    if (!item.queueKey) return;

    existing = findQueueRowByQueueKey_(queueSheet, item.queueKey);
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
      existing.contextJson = safeJson_(item);

      if (!isManualRuleStatus_(existing.status)) {
        existing.status = 'OPEN';
        existing.resolvedAt = '';
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
    existing.contextJson = safeJson_(item);
    queueSheet.appendRow(queueObjectToRow_(existing));
    count += 1;
  });

  return count;
}

function readQueueStateForSheet_(queueSheet, sheetTitle) {
  var normalizedSheetTitle = normalizeText_(sheetTitle);
  var rows = normalizedSheetTitle ? readSheetObjectsByMatch_(queueSheet, 3, normalizedSheetTitle) : [];
  return {
    sheetTitle: normalizedSheetTitle,
    rows: rows,
    byKey: buildQueueIndex_(rows),
    dirtyRowsByNumber: {},
    newRows: []
  };
}

function resolveQueueItemsInState_(queueState, resolvedPendingItems, nowIso) {
  var count = 0;

  resolvedPendingItems.forEach(function(rawItem) {
    var item = normalizeQueueItem_(rawItem, queueState.sheetTitle);
    var existing = queueState.byKey[item.queueKey];
    if (!item.queueKey || !existing || isManualRuleStatus_(existing.status)) return;

    if (String(existing.status || '').trim() !== 'RESOLVED') {
      count += 1;
    }

    existing.status = 'RESOLVED';
    existing.lastSeenAt = nowIso;
    existing.resolvedAt = nowIso;
    existing.contextJson = safeJson_(item);
    markQueueStateDirty_(queueState, existing);
  });

  return count;
}

function upsertQueueItemsInState_(queueState, currentUnmatchedItems, nowIso, defaultSheetTitle) {
  var count = 0;

  currentUnmatchedItems.forEach(function(rawItem) {
    var item = normalizeQueueItem_(rawItem, defaultSheetTitle || queueState.sheetTitle);
    var existing;
    if (!item.queueKey) return;

    existing = queueState.byKey[item.queueKey];
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
      existing.contextJson = safeJson_(item);

      if (!isManualRuleStatus_(existing.status)) {
        existing.status = 'OPEN';
        existing.resolvedAt = '';
        existing.resolutionType = '';
        existing.resolutionLabel = '';
        existing.resolutionTargetRow = '';
        existing.resolutionTargetName = '';
        existing.resolutionTargetPhone = '';
        existing.handledBy = '';
        existing.handledAt = '';
        count += 1;
      }

      markQueueStateDirty_(queueState, existing);
      return;
    }

    existing = createQueueRowObject_(item);
    existing.status = 'OPEN';
    existing.firstSeenAt = nowIso;
    existing.lastSeenAt = nowIso;
    existing.attemptCount = 1;
    existing.contextJson = safeJson_(item);
    queueState.rows.push(existing);
    queueState.byKey[item.queueKey] = existing;
    queueState.newRows.push(existing);
    count += 1;
  });

  return count;
}

function markQueueStateDirty_(queueState, row) {
  var rowNumber = Number(row && row.rowNumber || 0);
  if (!rowNumber) return;
  queueState.dirtyRowsByNumber[rowNumber] = row;
}

function flushQueueState_(queueSheet, queueState) {
  var dirtyRows = objectValues_(queueState.dirtyRowsByNumber);
  var startRow;

  if (dirtyRows.length) {
    writeRowsByObjects_(queueSheet, dirtyRows, queueObjectToRow_, MATCH_BACKEND.queueHeaders.length);
  }

  if (!queueState.newRows.length) {
    return;
  }

  startRow = queueSheet.getLastRow() + 1;
  queueState.newRows.forEach(function(row, index) {
    row.rowNumber = startRow + index;
  });
  queueSheet
    .getRange(startRow, 1, queueState.newRows.length, MATCH_BACKEND.queueHeaders.length)
    .setValues(queueState.newRows.map(queueObjectToRow_));
}

function buildQueueCounts_(queueRows) {
  return (queueRows || []).reduce(function(counts, row) {
    var status = String(row && row.status || '').trim();
    if (status === 'OPEN') counts.openCount += 1;
    if (isManualRuleStatus_(status)) counts.manualRuleCount += 1;
    return counts;
  }, {
    openCount: 0,
    manualRuleCount: 0
  });
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
      snapshotReady: false,
      openCount: 0,
      manualRuleCount: 0,
      lastSavedAt: '',
      lastSnapshotAt: ''
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

function buildWritePlan_(updates) {
  var grouped = {};
  var single = [];
  var batched = [];

  (updates || []).forEach(function(item) {
    var parsedRange;
    var cellInfo;
    var key;

    if (!item || !item.range || !Array.isArray(item.values) || item.values.length !== 1 || !Array.isArray(item.values[0]) || item.values[0].length !== 1) {
      single.push(item);
      return;
    }

    parsedRange = parseA1Range_(item.range);
    cellInfo = parseSingleCellA1_(parsedRange.a1Notation);
    if (!cellInfo) {
      single.push(item);
      return;
    }

    key = parsedRange.sheetTitle + '::' + cellInfo.columnNumber;
    if (!grouped[key]) {
      grouped[key] = [];
    }
    grouped[key].push({
      sheetTitle: parsedRange.sheetTitle,
      rowNumber: cellInfo.rowNumber,
      columnNumber: cellInfo.columnNumber,
      value: item.values[0][0]
    });
  });

  Object.keys(grouped).forEach(function(key) {
    var rows = grouped[key].sort(function(left, right) {
      return left.rowNumber - right.rowNumber;
    });
    var current = null;

    rows.forEach(function(item) {
      if (!current || item.rowNumber !== current.lastRow + 1) {
        if (current) {
          batched.push({
            sheetTitle: current.sheetTitle,
            startRow: current.startRow,
            columnNumber: current.columnNumber,
            width: 1,
            values: current.values
          });
        }
        current = {
          sheetTitle: item.sheetTitle,
          startRow: item.rowNumber,
          lastRow: item.rowNumber,
          columnNumber: item.columnNumber,
          values: [[item.value]]
        };
        return;
      }

      current.lastRow = item.rowNumber;
      current.values.push([item.value]);
    });

    if (current) {
      batched.push({
        sheetTitle: current.sheetTitle,
        startRow: current.startRow,
        columnNumber: current.columnNumber,
        width: 1,
        values: current.values
      });
    }
  });

  return {
    batched: batched,
    single: single
  };
}

function parseSingleCellA1_(a1Notation) {
  var match = normalizeText_(a1Notation).match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;

  return {
    columnNumber: columnLetterToNumber_(match[1]),
    rowNumber: Number(match[2] || 0)
  };
}

function getSheetFromCache_(spreadsheet, cache, sheetTitle) {
  if (!cache[sheetTitle]) {
    cache[sheetTitle] = spreadsheet.getSheetByName(sheetTitle);
  }
  return cache[sheetTitle];
}

function findQueueRowByQueueKey_(sheet, queueKey) {
  return findRowObjectByTextInColumn_(sheet, 1, queueKey);
}

function findSummaryRowBySheetTitle_(sheet, sheetTitle) {
  return findRowObjectByTextInColumn_(sheet, 1, sheetTitle);
}

function findRowObjectByTextInColumn_(sheet, columnNumber, targetText) {
  var lastRow = sheet.getLastRow();
  var match;
  if (lastRow <= 1 || !targetText) return null;

  match = sheet
    .getRange(2, columnNumber, lastRow - 1, 1)
    .createTextFinder(normalizeText_(targetText))
    .matchEntireCell(true)
    .findNext();

  return match ? readRowObject_(sheet, match.getRow()) : null;
}

function readSheetObjectsByMatch_(sheet, columnNumber, targetText) {
  var lastRow = sheet.getLastRow();
  var range;
  var matches;
  var rowNumbers;

  if (lastRow <= 1 || !targetText) return [];

  range = sheet.getRange(2, columnNumber, lastRow - 1, 1);
  matches = range.createTextFinder(normalizeText_(targetText)).matchEntireCell(true).findAll();
  rowNumbers = matches.map(function(foundRange) {
    return foundRange.getRow();
  });

  return readSheetObjectsByRowNumbers_(sheet, rowNumbers);
}

function readRowObject_(sheet, rowNumber) {
  var lastColumn = sheet.getLastColumn();
  var headers;
  var values;
  var item = { rowNumber: rowNumber };

  if (rowNumber <= 1 || lastColumn === 0) return item;

  headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  values = sheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
  headers.forEach(function(header, index) {
    item[String(header || '').trim()] = values[index];
  });
  return item;
}

function readSheetObjectsByRowNumbers_(sheet, rowNumbers) {
  var lastColumn = sheet.getLastColumn();
  var headers;
  var sortedRows;
  var groups = [];
  var current = null;
  var results = {};
  var output = [];

  if (!rowNumbers || !rowNumbers.length || lastColumn === 0) return [];

  headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  sortedRows = rowNumbers.slice().sort(function(left, right) {
    return left - right;
  });

  sortedRows.forEach(function(rowNumber) {
    if (!current || rowNumber !== current.end + 1) {
      if (current) groups.push(current);
      current = { start: rowNumber, end: rowNumber };
      return;
    }
    current.end = rowNumber;
  });
  if (current) groups.push(current);

  groups.forEach(function(group) {
    var values = sheet.getRange(group.start, 1, (group.end - group.start) + 1, lastColumn).getValues();
    values.forEach(function(row, index) {
      var rowNumber = group.start + index;
      var item = { rowNumber: rowNumber };
      headers.forEach(function(header, headerIndex) {
        item[String(header || '').trim()] = row[headerIndex];
      });
      results[rowNumber] = item;
    });
  });

  rowNumbers.forEach(function(rowNumber) {
    if (results[rowNumber]) {
      output.push(results[rowNumber]);
    }
  });

  return output;
}

function refreshSummaryForSheet_(sheetTitle, options) {
  var normalizedSheetTitle = normalizeText_(sheetTitle);
  var summarySheet = ensureStorageSheet_(MATCH_BACKEND.summarySheetName, MATCH_BACKEND.summaryHeaders);
  var queueSheet = ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders);
  var snapshotSheet = ensureStorageSheet_(MATCH_BACKEND.snapshotSheetName, MATCH_BACKEND.snapshotHeaders);
  var summaryRow = findSummaryRowBySheetTitle_(summarySheet, normalizedSheetTitle);
  var queueRows = normalizedSheetTitle ? readSheetObjectsByMatch_(queueSheet, 3, normalizedSheetTitle) : [];
  var snapshotMeta = normalizedSheetTitle ? getLatestSnapshotMeta_(snapshotSheet, normalizedSheetTitle) : null;
  var openCount = 0;
  var manualRuleCount = 0;
  var rowObject;

  queueRows.forEach(function(row) {
    var status = String(row.status || '').trim();
    if (status === 'OPEN') openCount += 1;
    if (isManualRuleStatus_(status)) manualRuleCount += 1;
  });

  rowObject = {
    rowNumber: summaryRow ? summaryRow.rowNumber : (summarySheet.getLastRow() + 1),
    sheetTitle: normalizedSheetTitle,
    snapshotReady: !!snapshotMeta,
    openCount: openCount,
    manualRuleCount: manualRuleCount,
    lastSavedAt: normalizeText_(options.lastSavedAt) || normalizeText_(summaryRow && summaryRow.lastSavedAt),
    lastSnapshotAt: normalizeText_(options.lastSnapshotAt) || normalizeText_(summaryRow && summaryRow.lastSnapshotAt)
  };

  if (!rowObject.lastSavedAt && snapshotMeta) {
    rowObject.lastSavedAt = normalizeText_(snapshotMeta.savedAt);
  }
  if (!rowObject.lastSnapshotAt && snapshotMeta) {
    rowObject.lastSnapshotAt = normalizeText_(snapshotMeta.savedAt);
  }

  if (summaryRow) {
    summarySheet.getRange(rowObject.rowNumber, 1, 1, MATCH_BACKEND.summaryHeaders.length).setValues([summaryObjectToRow_(rowObject)]);
  } else {
    summarySheet.appendRow(summaryObjectToRow_(rowObject));
  }

  return rowObject;
}

function rebuildSummarySheet_() {
  var snapshotRows = readSheetObjects_(ensureStorageSheet_(MATCH_BACKEND.snapshotSheetName, MATCH_BACKEND.snapshotHeaders));
  var queueRows = readSheetObjects_(ensureStorageSheet_(MATCH_BACKEND.queueSheetName, MATCH_BACKEND.queueHeaders));
  var summarySheet = ensureStorageSheet_(MATCH_BACKEND.summarySheetName, MATCH_BACKEND.summaryHeaders);
  var map = {};
  var rows;

  snapshotRows.forEach(function(row) {
    var key = normalizeText_(row.sheetTitle);
    var savedAt = normalizeText_(row.savedAt);
    if (!key) return;
    ensureOverviewItem_(map, key).snapshotReady = true;
    ensureOverviewItem_(map, key).lastSavedAt = maxIso_(ensureOverviewItem_(map, key).lastSavedAt, savedAt);
    ensureOverviewItem_(map, key).lastSnapshotAt = maxIso_(ensureOverviewItem_(map, key).lastSnapshotAt, savedAt);
  });

  queueRows.forEach(function(row) {
    var key = normalizeText_(row.sheetTitle);
    if (!key) return;
    if (String(row.status || '').trim() === 'OPEN') {
      ensureOverviewItem_(map, key).openCount += 1;
    }
    if (isManualRuleStatus_(row.status)) {
      ensureOverviewItem_(map, key).manualRuleCount += 1;
    }
  });

  rows = objectValues_(map).map(summaryObjectToRow_);
  summarySheet.clearContents();
  summarySheet.getRange(1, 1, 1, MATCH_BACKEND.summaryHeaders.length).setValues([MATCH_BACKEND.summaryHeaders]);
  if (rows.length) {
    summarySheet.getRange(2, 1, rows.length, MATCH_BACKEND.summaryHeaders.length).setValues(rows);
  }
}

function summaryRowToOverview_(row) {
  var snapshotReady = row && (row.snapshotReady === true || String(row.snapshotReady || '').trim().toUpperCase() === 'TRUE');
  return {
    sheetTitle: normalizeText_(row.sheetTitle),
    snapshotReady: !!snapshotReady,
    openCount: Number(row.openCount || 0),
    manualRuleCount: Number(row.manualRuleCount || 0),
    lastSavedAt: normalizeText_(row.lastSavedAt),
    lastSnapshotAt: normalizeText_(row.lastSnapshotAt)
  };
}

function summaryObjectToRow_(item) {
  return [
    item.sheetTitle || '',
    item.snapshotReady ? 'TRUE' : 'FALSE',
    Number(item.openCount || 0),
    Number(item.manualRuleCount || 0),
    item.lastSavedAt || '',
    item.lastSnapshotAt || ''
  ];
}

function saveLatestSnapshot_(sheetTitle, snapshot, options) {
  var normalizedSheetTitle = normalizeText_(sheetTitle);
  var snapshotSheet = ensureStorageSheet_(MATCH_BACKEND.snapshotSheetName, MATCH_BACKEND.snapshotHeaders);
  var savedAt = normalizeText_(options && options.savedAt) || new Date().toISOString();
  var existingRows;
  var raw = safeJson_(buildSnapshotEnvelope_(normalizedSheetTitle, snapshot, options, savedAt));
  var chunks;
  var rows;

  if (!normalizedSheetTitle) {
    throw new Error('Sheet title is required.');
  }
  if (!snapshot) {
    throw new Error('Snapshot payload is required.');
  }

  chunks = chunkText_(raw, MATCH_BACKEND.snapshotChunkSize);
  existingRows = readSheetObjectsByMatch_(snapshotSheet, 1, normalizedSheetTitle);
  rows = chunks.map(function(chunk, index) {
    return [
      normalizedSheetTitle,
      savedAt,
      options && options.previewOnly ? 'TRUE' : 'FALSE',
      normalizeText_(options && options.actorEmail),
      normalizeText_(options && options.actorName),
      index + 1,
      chunks.length,
      chunk
    ];
  });
  replaceMatchedRows_(snapshotSheet, existingRows, rows, MATCH_BACKEND.snapshotHeaders.length);

  return savedAt;
}

function buildSnapshotEnvelope_(sheetTitle, snapshot, options, savedAt) {
  var execution = withSnapshotMetrics_(snapshot.execution || snapshot, options && options.metrics);
  return {
    version: 1,
    sheetTitle: sheetTitle,
    savedAt: savedAt,
    previewOnly: !!(options && options.previewOnly),
    actorEmail: normalizeText_(options && options.actorEmail),
    actorName: normalizeText_(options && options.actorName),
    summary: options && options.summary ? options.summary : {},
    execution: execution
  };
}

function withSnapshotMetrics_(execution, metrics) {
  var cloned = safeParseJson_(safeJson_(execution || {})) || {};
  var historyState = cloned.historyState || {};

  if (metrics) {
    historyState.openCount = Number(metrics.openCount || 0);
    historyState.pendingCount = Number(metrics.openCount || 0);
    historyState.manualRuleCount = Number(metrics.manualRuleCount || 0);
  }
  historyState.synced = true;
  cloned.historyState = historyState;
  return cloned;
}

function loadLatestSnapshot_(sheetTitle) {
  var snapshotSheet = ensureStorageSheet_(MATCH_BACKEND.snapshotSheetName, MATCH_BACKEND.snapshotHeaders);
  var rows = readSheetObjectsByMatch_(snapshotSheet, 1, normalizeText_(sheetTitle));
  var latestSavedAt;
  var latestRows;
  var raw;

  if (!rows.length) {
    return null;
  }

  latestSavedAt = rows.reduce(function(latest, row) {
    return maxIso_(latest, normalizeText_(row.savedAt));
  }, '');
  latestRows = rows
    .filter(function(row) {
      return normalizeText_(row.savedAt) === latestSavedAt;
    })
    .sort(function(left, right) {
      return Number(left.partIndex || 0) - Number(right.partIndex || 0);
    });
  raw = latestRows.map(function(row) {
    return String(row.payloadChunk || '');
  }).join('');

  return safeParseJson_(raw);
}

function getLatestSnapshotMeta_(snapshotSheet, sheetTitle) {
  var rows = readSheetObjectsByMatch_(snapshotSheet, 1, normalizeText_(sheetTitle));
  var latestSavedAt;

  if (!rows.length) {
    return null;
  }

  latestSavedAt = rows.reduce(function(latest, row) {
    return maxIso_(latest, normalizeText_(row.savedAt));
  }, '');
  return {
    sheetTitle: normalizeText_(sheetTitle),
    savedAt: latestSavedAt
  };
}

function deleteSnapshotRows_(sheet, sheetTitle) {
  var rows = readSheetObjectsByMatch_(sheet, 1, normalizeText_(sheetTitle));
  deleteRowsDescending_(sheet, rows.map(function(row) {
    return Number(row.rowNumber || 0);
  }));
}

function replaceMatchedRows_(sheet, existingRows, newRows, columnCount) {
  var existingRowNumbers = (existingRows || []).map(function(row) {
    return Number(row && row.rowNumber || 0);
  }).filter(function(rowNumber) {
    return rowNumber > 1;
  }).sort(function(left, right) {
    return left - right;
  });
  var reusableCount = Math.min(existingRowNumbers.length, (newRows || []).length);

  if (reusableCount) {
    writeRowDataByNumbers_(sheet, existingRowNumbers.slice(0, reusableCount), newRows.slice(0, reusableCount), columnCount);
  }

  if (existingRowNumbers.length > reusableCount) {
    clearRowsByNumbers_(sheet, existingRowNumbers.slice(reusableCount), columnCount);
  }

  if ((newRows || []).length > reusableCount) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, newRows.length - reusableCount, columnCount)
      .setValues(newRows.slice(reusableCount));
  }
}

function writeSummarySnapshotState_(sheetTitle, queueCounts, options) {
  var normalizedSheetTitle = normalizeText_(sheetTitle);
  var summarySheet = ensureStorageSheet_(MATCH_BACKEND.summarySheetName, MATCH_BACKEND.summaryHeaders);
  var summaryRow = findSummaryRowBySheetTitle_(summarySheet, normalizedSheetTitle);
  var lastSnapshotAt = normalizeText_(options && options.lastSnapshotAt) || normalizeText_(summaryRow && summaryRow.lastSnapshotAt);
  var rowObject = {
    rowNumber: summaryRow ? summaryRow.rowNumber : (summarySheet.getLastRow() + 1),
    sheetTitle: normalizedSheetTitle,
    snapshotReady: !!lastSnapshotAt,
    openCount: Number(queueCounts && queueCounts.openCount || 0),
    manualRuleCount: Number(queueCounts && queueCounts.manualRuleCount || 0),
    lastSavedAt: normalizeText_(options && options.lastSavedAt) || normalizeText_(summaryRow && summaryRow.lastSavedAt),
    lastSnapshotAt: lastSnapshotAt
  };

  if (summaryRow) {
    summarySheet.getRange(rowObject.rowNumber, 1, 1, MATCH_BACKEND.summaryHeaders.length).setValues([summaryObjectToRow_(rowObject)]);
  } else {
    summarySheet.appendRow(summaryObjectToRow_(rowObject));
  }

  return rowObject;
}

function writeRowsByObjects_(sheet, rows, mapper, columnCount) {
  var sorted = (rows || []).slice().filter(function(row) {
    return Number(row && row.rowNumber || 0) > 1;
  }).sort(function(left, right) {
    return Number(left.rowNumber || 0) - Number(right.rowNumber || 0);
  });

  writeRowDataByNumbers_(sheet, sorted.map(function(row) {
    return Number(row.rowNumber || 0);
  }), sorted.map(function(row) {
    return mapper(row);
  }), columnCount);
}

function writeRowDataByNumbers_(sheet, rowNumbers, rowValues, columnCount) {
  var groups = [];
  var current = null;

  (rowNumbers || []).forEach(function(rowNumber, index) {
    var values = rowValues[index];
    if (!current || rowNumber !== current.lastRow + 1) {
      if (current) groups.push(current);
      current = {
        startRow: rowNumber,
        lastRow: rowNumber,
        values: [values]
      };
      return;
    }

    current.lastRow = rowNumber;
    current.values.push(values);
  });

  if (current) {
    groups.push(current);
  }

  groups.forEach(function(group) {
    sheet
      .getRange(group.startRow, 1, group.values.length, columnCount)
      .setValues(group.values);
  });
}

function clearRowsByNumbers_(sheet, rowNumbers, columnCount) {
  var blanks = (rowNumbers || []).map(function() {
    return buildBlankRow_(columnCount);
  });

  if (!blanks.length) {
    return;
  }

  writeRowDataByNumbers_(sheet, rowNumbers, blanks, columnCount);
}

function buildBlankRow_(columnCount) {
  var row = [];
  var index;

  for (index = 0; index < Number(columnCount || 0); index += 1) {
    row.push('');
  }

  return row;
}

function deleteRowsDescending_(sheet, rowNumbers) {
  rowNumbers
    .filter(function(rowNumber) { return Number(rowNumber || 0) > 1; })
    .sort(function(left, right) { return right - left; })
    .forEach(function(rowNumber) {
      sheet.deleteRow(rowNumber);
    });
}

function chunkText_(text, chunkSize) {
  var size = Math.max(Number(chunkSize || 0), 1000);
  var raw = String(text || '');
  var chunks = [];
  var index;

  if (!raw) {
    return ['{}'];
  }

  for (index = 0; index < raw.length; index += size) {
    chunks.push(raw.slice(index, index + size));
  }
  return chunks;
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
  var lastColumn = sheet.getLastColumn();
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

function getCachedJson_(key) {
  var cache = CacheService.getScriptCache();
  var raw = cache.get(key);
  return raw ? safeParseJson_(raw) : null;
}

function putCachedJson_(key, value, ttlSeconds) {
  var raw = safeJson_(value);
  if (raw.length > 90000) return;
  CacheService.getScriptCache().put(key, raw, Number(ttlSeconds || 30));
}

function invalidateStorageCaches_(sheetTitle) {
  var cache = CacheService.getScriptCache();
  cache.remove(buildCacheKey_('storageOverview', 'all'));
  if (sheetTitle) {
    cache.remove(buildCacheKey_('storedSheetData', sheetTitle));
  }
}

function buildCacheKey_(prefix, key) {
  return ['kakao-check', prefix, normalizeText_(key)].join('::');
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
