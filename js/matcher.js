(function(){
  var config = window.KakaoCheckConfig;
  var parser = window.KakaoCheckParser;
  var utils = window.KakaoCheckUtils;
  var NAME_ONLY_SENTINEL = parser.NAME_ONLY_SENTINEL;

  function buildResult(rosterRows, parsed, options){
    var pendingItems;
    var manualRules;
    var summary;
    var rosterIndex;
    var reportItems = [];
    var reportByRowNumber = new Map();
    var updates = [];
    var currentUnmatchedItems;
    var pendingResolution;
    var manualResolution;
    var report;
    var reportEntered;
    var reportLeft;
    var reportMissing;
    var extra;
    var leavers;

    options = options || {};
    pendingItems = Array.isArray(options.pendingItems) ? options.pendingItems : [];
    manualRules = Array.isArray(options.manualRules) ? options.manualRules : [];
    summary = parser.summarizeActiveState(parsed.events || []);
    rosterIndex = buildRosterIndex(Array.isArray(rosterRows) ? rosterRows : []);

    rosterIndex.rows.forEach(function(row){
      var item = buildRosterReportItem(row, summary, rosterIndex);
      reportItems.push(item);
      reportByRowNumber.set(item.rowNumber, item);
      if(item.status === '입장'){
        updates.push(createStatusUpdate(row));
      }
    });

    currentUnmatchedItems = buildCurrentUnmatchedItems(
      summary.activeDigitsByName,
      rosterIndex,
      String(options.sheetTitle || (rosterIndex.rows[0] && rosterIndex.rows[0].sheetTitle) || '').trim()
    );

    pendingResolution = resolvePendingItems(pendingItems, rosterIndex, reportByRowNumber);
    manualResolution = applyManualRules(currentUnmatchedItems, manualRules, rosterIndex, reportByRowNumber);
    currentUnmatchedItems = manualResolution.remainingItems;
    updates = dedupeUpdates(updates.concat(pendingResolution.updates, manualResolution.updates));

    report = reportItems.map(reportItemToRow);
    reportEntered = reportItems.filter(isEntered).map(reportItemToRow);
    reportLeft = reportItems.filter(isLeft).map(reportItemToRow);
    reportMissing = reportItems.filter(isMissing).map(reportItemToRow);
    extra = currentUnmatchedItems
      .filter(function(item){ return item.category === 'outside-roster'; })
      .map(function(item){ return [item.label]; })
      .sort(compareRow);
    leavers = buildLeavers(summary.leftFinalDigitsByName, summary.activeDigitsByName);

    return {
      report: report,
      reportItems: reportItems,
      reportEntered: reportEntered,
      reportLeft: reportLeft,
      reportMissing: reportMissing,
      missingRosterItems: reportItems.filter(isMissing).map(copyReportItem),
      extra: extra,
      leavers: leavers,
      currentUnmatchedItems: currentUnmatchedItems,
      currentUnmatchedRows: currentUnmatchedItems.map(currentUnmatchedItemToRow),
      pendingResolvedItems: pendingResolution.resolvedItems,
      pendingResolvedRows: pendingResolution.resolvedItems.map(pendingResolvedItemToRow),
      pendingRemainingItems: pendingResolution.remainingItems,
      manualResolvedItems: manualResolution.resolvedItems,
      manualResolvedRows: manualResolution.resolvedItems.map(manualResolvedItemToRow),
      excludedByRuleItems: manualResolution.excludedItems,
      excludedByRuleRows: manualResolution.excludedItems.map(excludedItemToRow),
      updates: updates,
      joinedCount: Number(parsed.joinedCount || 0),
      leftCount: Number(parsed.leftCount || 0),
      attendingCount: reportEntered.length,
      finalLeftCount: reportLeft.length,
      missingCount: reportMissing.length
    };
  }

  function buildRosterIndex(rosterRows){
    var byName = new Map();
    var byRowNumber = new Map();
    var byIdentity = new Map();
    var nameCount = new Map();
    var uniqueIdentityCountByName = new Map();

    rosterRows.forEach(function(row){
      var key = row && row.nameNormalized;
      var identityKey;
      if(!row) return;

      byRowNumber.set(Number(row.rowNumber), row);
      if(!key) return;

      if(!byName.has(key)){
        byName.set(key, []);
      }
      byName.get(key).push(row);
      nameCount.set(key, (nameCount.get(key) || 0) + 1);

      identityKey = getIdentityKey(row);
      if(!byIdentity.has(identityKey)){
        byIdentity.set(identityKey, []);
      }
      byIdentity.get(identityKey).push(row);
    });

    byName.forEach(function(rows, nameKey){
      var identityKeys = new Set(rows.map(getIdentityKey));
      uniqueIdentityCountByName.set(nameKey, identityKeys.size);
    });

    return {
      rows: rosterRows.slice(),
      byName: byName,
      byRowNumber: byRowNumber,
      byIdentity: byIdentity,
      nameCount: nameCount,
      uniqueIdentityCountByName: uniqueIdentityCountByName
    };
  }

  function buildRosterReportItem(row, activeSummary, rosterIndex){
    var nameKey = row.nameNormalized;
    var phone4 = utils.last4Digits(row.phone);
    var activeSet = activeSummary.activeDigitsByName.get(nameKey) || new Set();
    var leftSet = activeSummary.leftFinalDigitsByName.get(nameKey) || new Set();
    var duplicateName = (rosterIndex.nameCount.get(nameKey) || 0) > 1;
    var ambiguousDuplicate = duplicateName && (rosterIndex.uniqueIdentityCountByName.get(nameKey) || 0) > 1;
    var hasNameOnlyActive = activeSet.has(NAME_ONLY_SENTINEL);
    var hasNameOnlyLeft = leftSet.has(NAME_ONLY_SENTINEL);
    var hasExactActive = phone4 ? activeSet.has(phone4) : false;
    var hasExactLeft = phone4 ? leftSet.has(phone4) : false;
    var hasAnyActive = activeSet.size > 0;
    var hasAnyLeft = leftSet.size > 0;
    var status = '미입장';
    var matched = '';
    var notes = '';

    if(!nameKey){
      return {
        rowNumber: row.rowNumber,
        sheetTitle: row.sheetTitle,
        name: row.name,
        phone: row.phone,
        status: status,
        matched: matched,
        notes: notes
      };
    }

    if(ambiguousDuplicate){
      if(hasExactActive){
        status = '입장';
        matched = utils.formatPersonLabel(row.name, phone4);
      } else if(hasExactLeft){
        status = '퇴장';
        matched = utils.formatPersonLabel(row.name, phone4);
        notes = '최종 퇴장';
      } else if(hasNameOnlyActive && hasAnyActive){
        notes = '동명이인 / 번호 없음';
      } else if(hasAnyActive){
        notes = '동명이인 / 번호 불일치';
      } else if(hasNameOnlyLeft && hasAnyLeft){
        notes = '동명이인 / 퇴장 번호 없음';
      } else if(hasAnyLeft){
        notes = '동명이인 / 퇴장 번호 불일치';
      }
    } else {
      if(hasExactActive){
        status = '입장';
        matched = utils.formatPersonLabel(row.name, phone4);
      } else if(hasAnyActive){
        status = '입장';
        matched = hasNameOnlyActive
          ? String(row.name || '')
          : utils.formatPersonLabel(row.name, firstNonSentinel(activeSet));
        if(hasNameOnlyActive){
          notes = '번호 없음';
        }
      } else if(hasExactLeft){
        status = '퇴장';
        matched = utils.formatPersonLabel(row.name, phone4);
        notes = '최종 퇴장';
      } else if(hasAnyLeft){
        status = '퇴장';
        matched = hasNameOnlyLeft
          ? String(row.name || '')
          : utils.formatPersonLabel(row.name, firstNonSentinel(leftSet));
        notes = hasNameOnlyLeft ? '번호 없음 / 최종 퇴장' : '최종 퇴장';
      }
    }

    return {
      rowNumber: row.rowNumber,
      sheetTitle: row.sheetTitle,
      name: row.name,
      phone: row.phone,
      status: status,
      matched: matched,
      notes: notes
    };
  }

  function buildCurrentUnmatchedItems(activeDigitsByName, rosterIndex, fallbackSheetTitle){
    var items = [];

    activeDigitsByName.forEach(function(markerSet, nameKey){
      var candidates;
      var uniqueIdentityCount;
      if(!nameKey || !markerSet || !markerSet.size) return;

      candidates = rosterIndex.byName.get(nameKey) || [];
      Array.from(markerSet).forEach(function(marker){
        var exactMatches;
        if(!candidates.length){
          items.push(createQueueItem({
            sheetTitle: fallbackSheetTitle,
            category: 'outside-roster',
            name: nameKey,
            nameNormalized: nameKey,
            phone4: marker === NAME_ONLY_SENTINEL ? '' : marker,
            label: formatMarkerLabel(nameKey, marker),
            reason: '명단에서 찾지 못함'
          }));
          return;
        }

        uniqueIdentityCount = Number(rosterIndex.uniqueIdentityCountByName.get(nameKey) || 0);
        if(uniqueIdentityCount <= 1){
          return;
        }

        if(marker === NAME_ONLY_SENTINEL){
          items.push(createQueueItem({
            sheetTitle: candidates[0].sheetTitle,
            category: 'duplicate-name-no-number',
            name: candidates[0].name || nameKey,
            nameNormalized: nameKey,
            phone4: '',
            label: formatMarkerLabel(candidates[0].name || nameKey, marker),
            reason: '동명이인 / 번호 없음'
          }));
          return;
        }

        exactMatches = candidates.filter(function(row){
          return utils.last4Digits(row.phone) === marker;
        });

        if(exactMatches.length === 0){
          items.push(createQueueItem({
            sheetTitle: candidates[0].sheetTitle,
            category: 'duplicate-name-mismatch',
            name: candidates[0].name || nameKey,
            nameNormalized: nameKey,
            phone4: marker,
            label: formatMarkerLabel(candidates[0].name || nameKey, marker),
            reason: '동명이인 / 번호 불일치'
          }));
        } else if(countUniqueIdentityRows(exactMatches) > 1){
          items.push(createQueueItem({
            sheetTitle: candidates[0].sheetTitle,
            category: 'duplicate-name-collision',
            name: candidates[0].name || nameKey,
            nameNormalized: nameKey,
            phone4: marker,
            label: formatMarkerLabel(candidates[0].name || nameKey, marker),
            reason: '동명이인 / 번호 중복'
          }));
        }
      });
    });

    return dedupeQueueItems(items);
  }

  function resolvePendingItems(pendingItems, rosterIndex, reportByRowNumber){
    var updates = [];
    var resolvedItems = [];
    var remainingItems = [];

    pendingItems.forEach(function(rawItem){
      var item = normalizeQueueItem(rawItem);
      var resolution = resolvePendingItem(item, rosterIndex);
      var primaryRow;

      if(!resolution){
        remainingItems.push(item);
        return;
      }

      primaryRow = resolution.primaryRow;
      resolution.rows.forEach(function(targetRow){
        var reportItem = reportByRowNumber.get(Number(targetRow.rowNumber));
        if(reportItem && reportItem.status === '미입장'){
          reportItem.status = '입장';
          reportItem.matched = resolution.label;
          reportItem.notes = mergeNotes(reportItem.notes, '이전 미매칭 재매칭');
          updates.push(createStatusUpdate(targetRow));
        } else if(reportItem && reportItem.status === '입장'){
          reportItem.notes = mergeNotes(reportItem.notes, '이전 미매칭 재매칭');
        }
      });

      resolvedItems.push({
        queueKey: item.queueKey,
        sheetTitle: primaryRow.sheetTitle,
        rowNumber: formatRowNumbers(resolution.rows),
        name: primaryRow.name,
        phone: primaryRow.phone,
        label: resolution.label || item.label,
        reason: item.reason || '이전 미매칭 재매칭'
      });
    });

    return {
      resolvedItems: resolvedItems,
      remainingItems: remainingItems,
      updates: dedupeUpdates(updates)
    };
  }

  function resolvePendingItem(item, rosterIndex){
    var candidates;
    var uniqueIdentityCount;
    var rowPhone4;
    var exactMatches;

    if(!item.nameNormalized) return null;

    candidates = rosterIndex.byName.get(item.nameNormalized) || [];
    if(!candidates.length) return null;
    uniqueIdentityCount = Number(rosterIndex.uniqueIdentityCountByName.get(item.nameNormalized) || 0);

    if(uniqueIdentityCount <= 1){
      rowPhone4 = utils.last4Digits(candidates[0].phone);
      if(item.phone4 && rowPhone4 && rowPhone4 !== item.phone4){
        return null;
      }
      return buildTargetResolution(candidates, item.phone4 || rowPhone4);
    }

    if(!item.phone4) return null;

    exactMatches = candidates.filter(function(row){
      return utils.last4Digits(row.phone) === item.phone4;
    });
    if(!exactMatches.length || countUniqueIdentityRows(exactMatches) !== 1) return null;

    return buildTargetResolution(exactMatches, item.phone4);
  }

  function applyManualRules(currentUnmatchedItems, manualRules, rosterIndex, reportByRowNumber){
    var rulesByQueueKey = new Map();
    var remainingItems = [];
    var resolvedItems = [];
    var excludedItems = [];
    var updates = [];

    manualRules.forEach(function(rawRule){
      var rule = normalizeManualRule(rawRule);
      if(rule && rule.queueKey){
        rulesByQueueKey.set(rule.queueKey, rule);
      }
    });

    currentUnmatchedItems.forEach(function(item){
      var rule = rulesByQueueKey.get(item.queueKey);
      var targetRows;
      var primaryRow;
      if(!rule){
        remainingItems.push(item);
        return;
      }

      if(rule.actionType === 'exclude-staff'){
        excludedItems.push({
          queueKey: item.queueKey,
          label: item.label,
          reason: rule.reason || '코칭스태프로 제외'
        });
        return;
      }

      if(rule.actionType !== 'match-student'){
        remainingItems.push(item);
        return;
      }

      targetRows = findManualTargetRows(rule, rosterIndex);
      if(!targetRows.length){
        remainingItems.push(item);
        return;
      }

      primaryRow = targetRows[0];
      targetRows.forEach(function(targetRow){
        var reportItem = reportByRowNumber.get(Number(targetRow.rowNumber));
        if(reportItem && reportItem.status === '미입장'){
          reportItem.status = '입장';
          reportItem.matched = rule.resolutionLabel || utils.formatPersonLabel(primaryRow.name, primaryRow.phone);
          reportItem.notes = mergeNotes(reportItem.notes, '저장된 수동 매칭');
          updates.push(createStatusUpdate(targetRow));
        } else if(reportItem){
          reportItem.notes = mergeNotes(reportItem.notes, '저장된 수동 매칭');
        }
      });

      resolvedItems.push({
        queueKey: item.queueKey,
        rowNumber: formatRowNumbers(targetRows),
        name: primaryRow.name,
        phone: primaryRow.phone,
        label: item.label,
        reason: rule.reason || '저장된 수동 매칭'
      });
    });

    return {
      remainingItems: remainingItems,
      resolvedItems: resolvedItems,
      excludedItems: excludedItems,
      updates: dedupeUpdates(updates)
    };
  }

  function findManualTargetRows(rule, rosterIndex){
    var rowNumber = Number(rule.targetRowNumber || 0);
    var row;
    var candidates;
    var exact;
    var uniqueIdentityCount;

    if(rowNumber && rosterIndex.byRowNumber.has(rowNumber)){
      return getIdentityRows(rosterIndex.byRowNumber.get(rowNumber), rosterIndex);
    }

    if(!rule.targetNameNormalized) return [];

    candidates = rosterIndex.byName.get(rule.targetNameNormalized) || [];
    if(!candidates.length) return [];
    uniqueIdentityCount = Number(rosterIndex.uniqueIdentityCountByName.get(rule.targetNameNormalized) || 0);
    if(uniqueIdentityCount <= 1) return dedupeRows(candidates);

    if(rule.targetPhone4){
      exact = candidates.filter(function(candidate){
        return utils.last4Digits(candidate.phone) === rule.targetPhone4;
      });
      if(exact.length && countUniqueIdentityRows(exact) === 1){
        return dedupeRows(exact);
      }
    }

    return [];
  }

  function buildLeavers(leftFinalDigitsByName, activeDigitsByName){
    var leavers = [];

    leftFinalDigitsByName.forEach(function(markerSet, nameKey){
      var activeSet = activeDigitsByName.get(nameKey);
      markerSet.forEach(function(marker){
        if(activeSet && activeSet.has(marker)) return;
        leavers.push([formatMarkerLabel(nameKey, marker)]);
      });
    });

    leavers.sort(compareRow);
    return leavers;
  }

  function createQueueItem(data){
    var sheetTitle = String(data.sheetTitle || '').trim();
    var name = String(data.name || data.nameNormalized || '').trim();
    var nameNormalized = utils.normalizeName(data.nameNormalized || name);
    var phone4 = utils.last4Digits(data.phone4);
    var category = String(data.category || 'outside-roster').trim();

    return {
      queueKey: buildQueueKey(sheetTitle, nameNormalized, phone4, category),
      sheetTitle: sheetTitle,
      category: category,
      name: name,
      nameNormalized: nameNormalized,
      phone4: phone4,
      label: String(data.label || utils.formatPersonLabel(name || nameNormalized, phone4)).trim(),
      reason: String(data.reason || '').trim()
    };
  }

  function normalizeQueueItem(rawItem){
    if(!rawItem || typeof rawItem !== 'object'){
      return createQueueItem({});
    }

    return createQueueItem({
      sheetTitle: rawItem.sheetTitle,
      category: rawItem.category,
      name: rawItem.name,
      nameNormalized: rawItem.nameNormalized || rawItem.name,
      phone4: rawItem.phone4 || rawItem.phone,
      label: rawItem.label,
      reason: rawItem.reason
    });
  }

  function normalizeManualRule(rawRule){
    var status;
    var actionType;
    if(!rawRule || typeof rawRule !== 'object') return null;

    status = String(rawRule.status || '').trim();
    actionType = String(rawRule.actionType || '').trim();

    if(!actionType){
      if(status === 'MANUAL_MATCH'){
        actionType = 'match-student';
      } else if(status === 'EXCLUDED_STAFF'){
        actionType = 'exclude-staff';
      }
    }

    return {
      queueKey: String(rawRule.queueKey || '').trim(),
      actionType: actionType,
      reason: String(rawRule.reason || rawRule.resolutionLabel || '').trim(),
      resolutionLabel: String(rawRule.resolutionLabel || '').trim(),
      targetRowNumber: Number(rawRule.targetRowNumber || rawRule.resolutionTargetRow || 0),
      targetName: String(rawRule.targetName || rawRule.resolutionTargetName || '').trim(),
      targetNameNormalized: utils.normalizeName(rawRule.targetName || rawRule.resolutionTargetName || ''),
      targetPhone4: utils.last4Digits(rawRule.targetPhone4 || rawRule.resolutionTargetPhone || '')
    };
  }

  function dedupeQueueItems(items){
    var seen = new Set();
    return items.filter(function(item){
      if(!item.queueKey || seen.has(item.queueKey)) return false;
      seen.add(item.queueKey);
      return true;
    });
  }

  function dedupeUpdates(items){
    var seen = new Set();
    return items.filter(function(item){
      if(!item || !item.range || seen.has(item.range)) return false;
      seen.add(item.range);
      return true;
    });
  }

  function buildQueueKey(sheetTitle, nameNormalized, phone4, category){
    return [
      String(sheetTitle || '').trim(),
      String(nameNormalized || '').trim(),
      String(phone4 || '').trim(),
      String(category || '').trim()
    ].join('::');
  }

  function buildIdentityKey(nameNormalized, phone){
    return [
      String(nameNormalized || '').trim(),
      utils.last4Digits(phone)
    ].join('::');
  }

  function getIdentityKey(row){
    return buildIdentityKey(row && row.nameNormalized, row && row.phone);
  }

  function dedupeRows(rows){
    var seen = new Set();
    return (rows || []).filter(function(row){
      var rowNumber = Number(row && row.rowNumber || 0);
      if(!rowNumber || seen.has(rowNumber)) return false;
      seen.add(rowNumber);
      return true;
    }).sort(function(left, right){
      return Number(left.rowNumber || 0) - Number(right.rowNumber || 0);
    });
  }

  function countUniqueIdentityRows(rows){
    return new Set((rows || []).map(getIdentityKey)).size;
  }

  function getIdentityRows(row, rosterIndex){
    if(!row || !rosterIndex) return [];
    return dedupeRows(rosterIndex.byIdentity.get(getIdentityKey(row)) || [row]);
  }

  function buildTargetResolution(rows, phone4){
    var targetRows = dedupeRows(rows);
    var primaryRow = targetRows[0];
    if(!primaryRow) return null;

    return {
      rows: targetRows,
      primaryRow: primaryRow,
      row: primaryRow,
      label: utils.formatPersonLabel(primaryRow.name, phone4 || primaryRow.phone)
    };
  }

  function formatRowNumbers(rows){
    return dedupeRows(rows).map(function(row){
      return row.rowNumber;
    }).join(', ');
  }

  function createStatusUpdate(row){
    var sheetTitle = config.escapeSheetTitle(row.sheetTitle);
    return {
      rowNumber: row.rowNumber,
      range: sheetTitle + '!' + config.columns.status + row.rowNumber,
      values: [[config.statusText]]
    };
  }

  function createStatusUpdatesForRow(row, rosterRows){
    var rosterIndex = buildRosterIndex(Array.isArray(rosterRows) ? rosterRows : []);
    var targetRows = getIdentityRows(row, rosterIndex);
    return dedupeUpdates(targetRows.map(createStatusUpdate));
  }

  function reportItemToRow(item){
    return [item.rowNumber, item.name, item.phone, item.status, item.matched, item.notes];
  }

  function currentUnmatchedItemToRow(item){
    return [getCategoryLabel(item.category), item.label, item.reason];
  }

  function pendingResolvedItemToRow(item){
    return [item.rowNumber, item.name, item.phone, item.label, item.reason];
  }

  function manualResolvedItemToRow(item){
    return [item.rowNumber, item.name, item.phone, item.label, item.reason];
  }

  function excludedItemToRow(item){
    return [item.label, item.reason];
  }

  function formatMarkerLabel(name, marker){
    return marker && marker !== NAME_ONLY_SENTINEL
      ? utils.formatPersonLabel(name, marker)
      : String(name || '').trim();
  }

  function firstNonSentinel(markerSet){
    var values = Array.from(markerSet || []);
    var index;
    for(index = 0; index < values.length; index += 1){
      if(values[index] !== NAME_ONLY_SENTINEL){
        return values[index];
      }
    }
    return '';
  }

  function mergeNotes(existing, incoming){
    return Array.from(new Set([existing, incoming].map(function(value){
      return String(value || '').trim();
    }).filter(Boolean))).join(' / ');
  }

  function getCategoryLabel(category){
    if(category === 'outside-roster') return '명단 외';
    if(category === 'duplicate-name-no-number') return '동명이인 / 번호 없음';
    if(category === 'duplicate-name-mismatch') return '동명이인 / 번호 불일치';
    if(category === 'duplicate-name-collision') return '동명이인 / 번호 중복';
    return String(category || '').trim();
  }

  function copyReportItem(item){
    return {
      rowNumber: item.rowNumber,
      sheetTitle: item.sheetTitle,
      name: item.name,
      phone: item.phone,
      status: item.status,
      matched: item.matched,
      notes: item.notes
    };
  }

  function isEntered(item){
    return item.status === '입장';
  }

  function isLeft(item){
    return item.status === '퇴장';
  }

  function isMissing(item){
    return item.status === '미입장';
  }

  function compareRow(left, right){
    return String(left[0] || '').localeCompare(String(right[0] || ''), 'ko');
  }

  window.KakaoCheckMatcher = {
    buildQueueKey: buildQueueKey,
    buildResult: buildResult,
    createStatusUpdate: createStatusUpdate,
    createStatusUpdatesForRow: createStatusUpdatesForRow,
    normalizeManualRule: normalizeManualRule,
    normalizeQueueItem: normalizeQueueItem
  };
})();
