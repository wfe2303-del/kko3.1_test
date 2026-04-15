(function(){
  var utils = window.KakaoCheckUtils;
  var parser = window.KakaoCheckParser;
  var config = window.KakaoCheckConfig;
  var NAME_ONLY_SENTINEL = parser.NAME_ONLY_SENTINEL;

  function buildResult(rosterRows, parsed, options){
    options = options || {};

    var pendingItems = Array.isArray(options.pendingItems) ? options.pendingItems : [];
    var summary = parser.summarizeActiveState(parsed.events);
    var rosterIndex = buildRosterIndex(rosterRows);
    var reportItems = [];
    var reportByRowNumber = new Map();
    var updates = [];

    rosterRows.forEach(function(row){
      var item = buildRosterReportItem(row, summary, rosterIndex.nameCount);
      reportItems.push(item);
      reportByRowNumber.set(item.rowNumber, item);
      if(item.status === '입장'){
        updates.push(makeStatusUpdate(row));
      }
    });

    var currentUnmatchedItems = buildCurrentUnmatchedItems(
      summary.activeDigitsByName,
      rosterIndex,
      String(options.sheetTitle || (rosterRows[0] && rosterRows[0].sheetTitle) || '').trim()
    );
    var pendingResolution = resolvePendingItems(pendingItems, rosterIndex, reportByRowNumber);
    updates = dedupeUpdates(updates.concat(pendingResolution.updates));

    var report = reportItems.map(reportItemToRow);
    var reportEntered = reportItems.filter(isEntered).map(reportItemToRow);
    var reportLeft = reportItems.filter(isLeft).map(reportItemToRow);
    var reportMissing = reportItems.filter(isMissing).map(reportItemToRow);

    var extra = currentUnmatchedItems
      .filter(function(item){ return item.category === 'outside-roster'; })
      .map(function(item){ return [item.label]; })
      .sort(compareRow);

    var leavers = buildLeavers(summary.leftFinalDigitsByName, summary.activeDigitsByName);

    return {
      report: report,
      reportItems: reportItems,
      reportEntered: reportEntered,
      reportLeft: reportLeft,
      reportMissing: reportMissing,
      extra: extra,
      leavers: leavers,
      currentUnmatchedItems: currentUnmatchedItems,
      currentUnmatchedRows: currentUnmatchedItems.map(currentUnmatchedItemToRow),
      pendingResolvedItems: pendingResolution.resolvedItems,
      pendingResolvedRows: pendingResolution.resolvedItems.map(pendingResolvedItemToRow),
      pendingRemainingItems: pendingResolution.remainingItems,
      pendingRemainingRows: pendingResolution.remainingItems.map(currentUnmatchedItemToRow),
      updates: updates,
      joinedCount: parsed.joinedCount,
      leftCount: parsed.leftCount,
      attendingCount: reportEntered.length,
      finalLeftCount: reportLeft.length,
      missingCount: reportMissing.length
    };
  }

  function buildRosterIndex(rosterRows){
    var byName = new Map();
    var nameCount = new Map();

    rosterRows.forEach(function(row){
      if(!row || !row.nameNormalized) return;
      if(!byName.has(row.nameNormalized)) byName.set(row.nameNormalized, []);
      byName.get(row.nameNormalized).push(row);
      nameCount.set(row.nameNormalized, (nameCount.get(row.nameNormalized) || 0) + 1);
    });

    return {
      byName: byName,
      nameCount: nameCount
    };
  }

  function buildRosterReportItem(row, activeSummary, nameCount){
    var nameKey = row.nameNormalized;
    var phone4 = utils.last4Digits(row.phone);
    var activeSet = activeSummary.activeDigitsByName.get(nameKey) || new Set();
    var leftSet = activeSummary.leftFinalDigitsByName.get(nameKey) || new Set();
    var isDup = (nameCount.get(nameKey) || 0) > 1;
    var hasNameOnlyActive = activeSet.has(NAME_ONLY_SENTINEL);
    var hasNameOnlyLeft = leftSet.has(NAME_ONLY_SENTINEL);
    var hasExactActive = phone4 ? activeSet.has(phone4) : false;
    var hasExactLeft = phone4 ? leftSet.has(phone4) : false;
    var hasAnyActive = activeSet.size > 0;
    var hasAnyLeft = leftSet.size > 0;
    var status = '미입장';
    var notes = '';
    var matched = '';

    if(!nameKey){
      return {
        rowNumber: row.rowNumber,
        name: row.name,
        phone: row.phone,
        status: status,
        matched: matched,
        notes: notes,
        fromPending: false
      };
    }

    if(isDup){
      if(hasExactActive){
        status = '입장';
        matched = formatPersonLabel(row.name, phone4);
      } else if(hasExactLeft){
        status = '퇴장';
        matched = formatPersonLabel(row.name, phone4);
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
        matched = formatPersonLabel(row.name, phone4);
      } else if(hasAnyActive){
        status = '입장';
        matched = hasNameOnlyActive
          ? row.name
          : formatPersonLabel(row.name, firstNonSentinel(activeSet));
        if(hasNameOnlyActive) notes = '번호 없음';
      } else if(hasExactLeft){
        status = '퇴장';
        matched = formatPersonLabel(row.name, phone4);
        notes = '최종 퇴장';
      } else if(hasAnyLeft){
        status = '퇴장';
        matched = hasNameOnlyLeft
          ? row.name
          : formatPersonLabel(row.name, firstNonSentinel(leftSet));
        notes = hasNameOnlyLeft ? '번호 없음 / 최종 퇴장' : '최종 퇴장';
      }
    }

    return {
      rowNumber: row.rowNumber,
      name: row.name,
      phone: row.phone,
      status: status,
      matched: matched,
      notes: notes,
      fromPending: false
    };
  }

  function buildCurrentUnmatchedItems(activeDigitsByName, rosterIndex, fallbackSheetTitle){
    var items = [];

    activeDigitsByName.forEach(function(set, nameKey){
      if(!nameKey || !set || !set.size) return;
      var candidates = rosterIndex.byName.get(nameKey) || [];
      Array.from(set).forEach(function(marker){
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

        if(candidates.length <= 1) return;

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

        var exactMatches = candidates.filter(function(row){
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
        } else if(exactMatches.length > 1){
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

      if(!resolution){
        remainingItems.push(item);
        return;
      }

      var reportItem = reportByRowNumber.get(resolution.row.rowNumber);
      if(reportItem && reportItem.status === '미입장'){
        reportItem.status = '입장';
        reportItem.matched = resolution.label;
        reportItem.notes = mergeNotes(reportItem.notes, '이전 미매칭 재매칭');
        reportItem.fromPending = true;
        updates.push(makeStatusUpdate(resolution.row));
      } else if(reportItem && reportItem.status === '입장'){
        reportItem.notes = mergeNotes(reportItem.notes, '이전 미매칭 재매칭');
      }

      resolvedItems.push({
        queueKey: item.queueKey,
        sheetTitle: resolution.row.sheetTitle,
        rowNumber: resolution.row.rowNumber,
        name: resolution.row.name,
        phone: resolution.row.phone,
        label: item.label || resolution.label,
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
    if(!item.nameNormalized) return null;

    var candidates = rosterIndex.byName.get(item.nameNormalized) || [];
    if(!candidates.length) return null;

    if(candidates.length === 1){
      var only = candidates[0];
      var rowPhone4 = utils.last4Digits(only.phone);
      if(item.phone4 && rowPhone4 && rowPhone4 !== item.phone4) return null;
      return {
        row: only,
        label: formatPersonLabel(only.name, item.phone4 || rowPhone4)
      };
    }

    if(!item.phone4) return null;

    var exactMatches = candidates.filter(function(row){
      return utils.last4Digits(row.phone) === item.phone4;
    });
    if(exactMatches.length !== 1) return null;

    return {
      row: exactMatches[0],
      label: formatPersonLabel(exactMatches[0].name, item.phone4)
    };
  }

  function buildLeavers(leftFinalDigitsByName, activeDigitsByName){
    var leavers = [];

    leftFinalDigitsByName.forEach(function(set, nameKey){
      var activeSet = activeDigitsByName.get(nameKey);
      set.forEach(function(value){
        if(activeSet && activeSet.has(value)) return;
        leavers.push([formatMarkerLabel(nameKey, value)]);
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
      label: String(data.label || formatPersonLabel(name || nameNormalized, phone4)).trim(),
      reason: String(data.reason || '').trim()
    };
  }

  function normalizeQueueItem(rawItem){
    if(!rawItem || typeof rawItem !== 'object') return createQueueItem({});
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

  function makeStatusUpdate(row){
    return {
      rowNumber: row.rowNumber,
      range: row.sheetTitle + '!' + config.columns.status + row.rowNumber,
      values: [[config.statusText]]
    };
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

  function formatMarkerLabel(name, marker){
    if(marker && marker !== NAME_ONLY_SENTINEL){
      return formatPersonLabel(name, marker);
    }
    return String(name || '').trim();
  }

  function formatPersonLabel(name, phone4){
    var displayName = String(name || '').trim();
    var digits = utils.last4Digits(phone4);
    return digits ? displayName + '/' + digits : displayName;
  }

  function firstNonSentinel(set){
    var values = Array.from(set || []);
    for(var i = 0; i < values.length; i += 1){
      if(values[i] !== NAME_ONLY_SENTINEL) return values[i];
    }
    return '';
  }

  function mergeNotes(existing, incoming){
    var values = [existing, incoming]
      .map(function(value){ return String(value || '').trim(); })
      .filter(Boolean);
    return Array.from(new Set(values)).join(' / ');
  }

  function getCategoryLabel(category){
    if(category === 'outside-roster') return '명단 외';
    if(category === 'duplicate-name-no-number') return '동명이인 / 번호 없음';
    if(category === 'duplicate-name-mismatch') return '동명이인 / 번호 불일치';
    if(category === 'duplicate-name-collision') return '동명이인 / 번호 중복';
    return category || '';
  }

  function isEntered(item){ return item.status === '입장'; }
  function isLeft(item){ return item.status === '퇴장'; }
  function isMissing(item){ return item.status === '미입장'; }

  function compareRow(a, b){
    return String(a[0] || '').localeCompare(String(b[0] || ''), 'ko');
  }

  window.KakaoCheckMatcher = {
    buildResult: buildResult
  };
})();
