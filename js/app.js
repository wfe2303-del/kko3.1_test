(function(){
  var auth = window.KakaoCheckAuth;
  var history = window.KakaoCheckHistory;
  var matcher = window.KakaoCheckMatcher;
  var parser = window.KakaoCheckParser;
  var sheets = window.KakaoCheckSheets;
  var ui = window.KakaoCheckUI;
  var utils = window.KakaoCheckUtils;

  var state = {
    nextPanelId: 1,
    sheetTitles: [],
    panels: new Map(),
    lastSheetLoadError: '',
    backendHealthy: false
  };

  function bootstrap(){
    ui.setLoginReady(false);
    bindTopLevelEvents();

    auth.onChange(function(payload){
      ui.setAuthState(payload);
      ui.setAuthError('');
      if(payload && payload.accessToken){
        handleAuthenticatedState();
      } else {
        ui.setAppNotice('', '');
      }
    });

    auth.init().then(function(){
      ui.setLoginReady(true);
      ui.setAuthError('');
    }).catch(function(error){
      ui.setLoginReady(true);
      ui.setAuthError(toUserMessage(error));
    });

    addPanel();
  }

  function bindTopLevelEvents(){
    document.getElementById('loginForm').addEventListener('submit', async function(event){
      var username = String(document.getElementById('loginUsername').value || '').trim();
      var password = String(document.getElementById('loginPassword').value || '');

      event.preventDefault();
      try {
        ui.setAuthError('');
        ui.setLoginReady(false);
        await auth.login(username, password);
        document.getElementById('loginPassword').value = '';
      } catch (error) {
        ui.setAuthError(toUserMessage(error));
      } finally {
        ui.setLoginReady(true);
      }
    });

    document.getElementById('logoutBtn').addEventListener('click', function(){
      auth.revoke();
      ui.closeModal();
    });

    document.getElementById('refreshTabsBtn').addEventListener('click', function(){
      loadSheetTitles();
    });

    document.getElementById('backendDataBtn').addEventListener('click', function(){
      openStorageOverview();
    });

    document.getElementById('addPanelBtn').addEventListener('click', function(){
      addPanel();
    });
  }

  async function handleAuthenticatedState(){
    await loadSheetTitles();
  }

  async function loadSheetTitles(){
    ui.showLoading('시트 목록 불러오는 중', '접근 가능한 시트 탭을 확인하고 있습니다.');
    try {
      state.sheetTitles = await sheets.listSheetTitles();
      state.lastSheetLoadError = '';
      state.backendHealthy = true;

      if(!state.sheetTitles.length){
        ui.setAppNotice('불러올 수 있는 시트가 없습니다. 계정 권한 또는 Apps Script의 대상 스프레드시트 접근 권한을 확인해 주세요.', 'warn');
      } else if(state.backendHealthy){
        ui.setAppNotice('', '');
      }
    } catch (error) {
      state.sheetTitles = [];
      state.lastSheetLoadError = toUserMessage(error);
      state.backendHealthy = false;
      ui.setAppNotice('시트 목록을 불러오지 못했습니다. ' + state.lastSheetLoadError, 'bad');
    } finally {
      ui.hideLoading();
    }
  }

  function addPanel(){
    var panelId = state.nextPanelId;
    var panelEl = ui.createPanelElement(panelId);
    var panelState = {
      id: panelId,
      el: panelEl,
      files: [],
      selectedSheet: '',
      lastExecution: null
    };

    state.nextPanelId += 1;
    state.panels.set(panelId, panelState);
    bindPanelEvents(panelState);
    ui.renderSelectedSheet(panelEl, '');
    ui.renderFileSummary(panelEl, []);
    ui.setPanelStatus(panelEl, '', '');
    ui.setPanelActionState(panelEl, {
      manualEnabled: false,
      manualLabel: '미매칭 수동 처리',
      storageEnabled: false
    });

    document.getElementById('panelsRoot').appendChild(panelEl);
    refreshRemoveButtons();
  }

  function removePanel(panelId){
    var panelState = state.panels.get(panelId);
    if(!panelState || state.panels.size <= 1) return;

    panelState.el.remove();
    state.panels.delete(panelId);

    if(ui.getModalState() && ui.getModalState().panelId === panelId){
      ui.closeModal();
    }

    refreshRemoveButtons();
  }

  function refreshRemoveButtons(){
    state.panels.forEach(function(panelState){
      var button = utils.qs('.panel-remove-btn', panelState.el);
      if(button){
        button.disabled = state.panels.size <= 1;
      }
    });
  }

  function bindPanelEvents(panelState){
    var el = panelState.el;
    var fileInput = utils.qs('.js-file-input', el);
    var runBtn = utils.qs('.js-run-btn', el);
    var demoBtn = utils.qs('.js-demo-btn', el);
    var removeBtn = utils.qs('.panel-remove-btn', el);
    var sheetBtn = utils.qs('.js-open-sheet-modal-btn', el);
    var fileBtn = utils.qs('.js-open-file-modal-btn', el);
    var manualBtn = utils.qs('.js-manual-review-btn', el);
    var storageBtn = utils.qs('.js-storage-data-btn', el);

    sheetBtn.addEventListener('click', function(){
      openSheetPicker(panelState.id);
    });
    fileBtn.addEventListener('click', function(){
      openFileManager(panelState.id);
    });
    manualBtn.addEventListener('click', function(){
      openManualReview(panelState.id);
    });
    storageBtn.addEventListener('click', function(){
      openStoredSheetData(panelState.selectedSheet, panelState.id);
    });
    fileInput.addEventListener('change', function(event){
      var incoming = Array.from(event.target.files || []);
      if(!incoming.length) return;

      panelState.files = appendUniqueFiles(panelState.files, incoming);
      ui.renderFileSummary(el, panelState.files);
      ui.setPanelStatus(el, panelState.files.length + '개 파일 준비됨', '');
      event.target.value = '';

      if(ui.getModalState() && ui.getModalState().type === 'file-manager' && ui.getModalState().panelId === panelState.id){
        openFileManager(panelState.id);
      }
    });
    runBtn.addEventListener('click', function(){
      runPanel(panelState.id);
    });
    demoBtn.addEventListener('click', function(){
      runDemo(panelState.id);
    });
    removeBtn.addEventListener('click', function(){
      removePanel(panelState.id);
    });
  }

  function openSheetPicker(panelId){
    var panelState = state.panels.get(panelId);
    if(!panelState) return;

    ui.openSheetPickerModal({
      panelId: panelId,
      title: panelState.selectedSheet ? (panelState.selectedSheet + ' 시트 선택') : '시트 선택',
      titles: state.sheetTitles,
      selectedTitle: panelState.selectedSheet,
      emptyText: state.lastSheetLoadError || '검색 결과가 없습니다.',
      onSelect: function(title){
        panelState.selectedSheet = title;
        panelState.lastExecution = null;
        ui.renderSelectedSheet(panelState.el, title);
        ui.setPanelError(panelState.el, '');
        ui.renderResults(panelState.el, []);
        syncPanelButtons(panelState);
      }
    });
  }

  function openFileManager(panelId){
    var panelState = state.panels.get(panelId);
    if(!panelState) return;
    ui.openFileManagerModal(buildFileModalOptions(panelState));
  }

  function buildFileModalOptions(panelState){
    return {
      panelId: panelState.id,
      title: panelState.selectedSheet ? (panelState.selectedSheet + ' 로그 파일 관리') : '로그 파일 관리',
      files: panelState.files,
      onAddRequest: function(){
        var input = utils.qs('.js-file-input', panelState.el);
        if(input) input.click();
      },
      onRemoveIndex: function(index){
        panelState.files.splice(index, 1);
        ui.renderFileSummary(panelState.el, panelState.files);
        ui.setPanelStatus(panelState.el, panelState.files.length ? '파일 목록 수정됨' : '파일 없음', '');
      },
      onClearAll: function(){
        panelState.files = [];
        ui.renderFileSummary(panelState.el, panelState.files);
        ui.setPanelStatus(panelState.el, '파일 없음', '');
      },
      getFreshOptions: function(){
        return buildFileModalOptions(panelState);
      }
    };
  }

  function appendUniqueFiles(existingFiles, incomingFiles){
    var seen = new Set(existingFiles.map(function(file){
      return [file.name, file.size, file.lastModified].join('::');
    }));
    var merged = existingFiles.slice();

    incomingFiles.forEach(function(file){
      var key = [file.name, file.size, file.lastModified].join('::');
      if(seen.has(key)) return;
      seen.add(key);
      merged.push(file);
    });

    return merged;
  }

  function createHistoryState(){
    return {
      enabled: false,
      pendingCount: 0,
      manualRuleCount: 0,
      openedCount: 0,
      resolvedCount: 0,
      openCount: 0,
      synced: false,
      syncSkipped: false,
      syncError: '',
      message: ''
    };
  }

  function syncPanelButtons(panelState){
    var execution = panelState.lastExecution;
    var manualEnabled = !!(state.backendHealthy && execution && !execution.previewOnly && execution.result && execution.result.currentUnmatchedItems.length);
    var manualLabel = '미매칭 수동 처리';

    if(execution && execution.previewOnly){
      manualLabel = '미리보기 결과는 수동 처리 불가';
    } else if(execution && execution.result){
      manualLabel = '미매칭 수동 처리 (' + execution.result.currentUnmatchedItems.length + ')';
    }

    ui.setPanelActionState(panelState.el, {
      manualEnabled: manualEnabled,
      manualLabel: manualLabel,
      storageEnabled: !!(state.backendHealthy && panelState.selectedSheet)
    });
  }

  async function runPanel(panelId){
    var panelState = state.panels.get(panelId);
    var previewOnly;
    var chatText;
    var parsed;
    var rosterResponse;
    var pendingResponse;
    var historyState = createHistoryState();
    var execution;
    var result;
    var statusMessage;

    if(!panelState) return;
    previewOnly = utils.qs('.js-preview-only', panelState.el).checked;

    try {
      ui.showLoading('입장 체크 처리 중', '수강생 명단과 미매칭 데이터를 불러오고 있습니다.');
      ui.setPanelError(panelState.el, '');
      ui.setPanelStatus(panelState.el, '실행 중...', 'warn');

      if(!auth.getAccessToken()) throw new Error('먼저 로그인해 주세요.');
      if(!panelState.selectedSheet) throw new Error('시트를 먼저 선택해 주세요.');
      if(!panelState.files.length) throw new Error('카카오톡 입장 로그 파일을 하나 이상 추가해 주세요.');

      chatText = await parser.combineFiles(panelState.files);
      parsed = parser.parseChatText(chatText);
      var settled = await Promise.allSettled([
        sheets.loadRosterRows(panelState.selectedSheet),
        history.listPending(panelState.selectedSheet)
      ]);

      if(settled[0].status !== 'fulfilled'){
        throw settled[0].reason;
      }

      rosterResponse = settled[0].value;

      if(settled[1].status === 'fulfilled'){
        pendingResponse = settled[1].value;
        historyState.enabled = !!pendingResponse.enabled;
        historyState.pendingCount = Array.isArray(pendingResponse.items) ? pendingResponse.items.length : 0;
        historyState.manualRuleCount = Array.isArray(pendingResponse.manualRules) ? pendingResponse.manualRules.length : 0;
      } else {
        pendingResponse = { enabled: false, items: [], manualRules: [] };
        historyState.enabled = false;
        historyState.message = toUserMessage(settled[1].reason);
      }

      if(!rosterResponse.rows.length){
        throw new Error('선택한 시트에서 N열이 수강생인 명단을 찾지 못했습니다. Apps Script 접근 권한과 N열 값을 확인해 주세요.');
      }

      result = matcher.buildResult(rosterResponse.rows, parsed, {
        pendingItems: pendingResponse.items || [],
        manualRules: pendingResponse.manualRules || [],
        sheetTitle: panelState.selectedSheet
      });

      if(!previewOnly && result.updates.length){
        await sheets.writeUpdates(result.updates);
      }

      execution = createExecutionRecord({
        parsed: parsed,
        previewOnly: previewOnly,
        rosterRows: rosterResponse.rows || [],
        sheetMeta: rosterResponse.meta || {},
        pendingItems: Array.isArray(pendingResponse.items) ? pendingResponse.items.slice() : [],
        manualRules: Array.isArray(pendingResponse.manualRules) ? pendingResponse.manualRules.slice() : [],
        result: result,
        historyState: historyState
      });

      try {
        await syncHistoryRun(panelState, execution);
      } catch (error) {
        historyState.syncError = toUserMessage(error);
      }

      panelState.lastExecution = execution;

      renderRunResult(panelState);

      statusMessage = previewOnly ? '미리보기 완료' : '체크 완료';
      if(rosterResponse.meta && rosterResponse.meta.fallbackUsed){
        statusMessage += ' (역할값 fallback 적용)';
      }

      if(historyState.syncError){
        ui.setPanelStatus(panelState.el, statusMessage + ' / 저장 경고', 'warn');
        ui.setPanelError(panelState.el, historyState.syncError);
      } else {
        ui.setPanelStatus(panelState.el, statusMessage, previewOnly || (rosterResponse.meta && rosterResponse.meta.fallbackUsed) ? 'warn' : 'ok');
      }
    } catch (error) {
      panelState.lastExecution = null;
      ui.renderResults(panelState.el, []);
      ui.setPanelStatus(panelState.el, '실패', 'bad');
      ui.setPanelError(panelState.el, toUserMessage(error));
      syncPanelButtons(panelState);
      console.error(error);
    } finally {
      ui.hideLoading();
    }
  }

  function runDemo(panelId){
    var panelState = state.panels.get(panelId);
    var roster;
    var parsed;

    if(!panelState) return;

    roster = [
      { rowNumber: 2, sheetTitle: '데모', name: '김하늘', phone: '010-1234-5678', nameNormalized: utils.normalizeName('김하늘') },
      { rowNumber: 3, sheetTitle: '데모', name: '박시온', phone: '010-3456-7890', nameNormalized: utils.normalizeName('박시온') },
      { rowNumber: 4, sheetTitle: '데모', name: '이하린', phone: '010-1717-1771', nameNormalized: utils.normalizeName('이하린') },
      { rowNumber: 5, sheetTitle: '데모', name: '최도윤', phone: '010-1111-0000', nameNormalized: utils.normalizeName('최도윤') },
      { rowNumber: 6, sheetTitle: '데모', name: '김하늘', phone: '010-9999-9999', nameNormalized: utils.normalizeName('김하늘') }
    ];

    parsed = parser.parseChatText([
      '최도윤0000님이 입장하셨습니다.',
      '김하늘5678님이 입장하셨습니다.',
      '이하린1771님이 들어왔습니다.',
      '박시온님이 입장하셨습니다.',
      '진주5148님이 입장하셨습니다.',
      '모델스탭님이 입장하셨습니다.',
      '김하늘5678님이 나갔습니다.',
      '김하늘5678님이 입장하셨습니다.',
      '박시온님이 나갔습니다.'
    ].join('\n'));

    panelState.lastExecution = createExecutionRecord({
      parsed: parsed,
      previewOnly: true,
      rosterRows: roster,
      sheetMeta: {
        fallbackUsed: false
      },
      pendingItems: [],
      manualRules: [],
      result: matcher.buildResult(roster, parsed, {
        pendingItems: [],
        manualRules: [],
        sheetTitle: '데모'
      }),
      historyState: createHistoryState()
    });

    ui.setPanelError(panelState.el, '');
    renderRunResult(panelState);
    ui.setPanelStatus(panelState.el, '데모 완료', 'ok');
  }

  async function syncHistoryRun(panelState, execution){
    var response;
    var historyState = execution.historyState || createHistoryState();

    if(execution.previewOnly){
      historyState.syncSkipped = true;
      return;
    }

    response = await history.syncRun(buildStoragePayload(panelState, execution));
    if(response && response.enabled === false){
      historyState.enabled = false;
      historyState.message = response.message || '저장 실행 백엔드가 비활성화되어 있습니다.';
      historyState.syncSkipped = true;
      return;
    }

    historyState.enabled = true;
    historyState.synced = !!response.synced;
    historyState.syncError = '';
    historyState.openCount = Number(response.openCount || 0);
    historyState.openedCount = Number(response.openedCount || 0);
    historyState.resolvedCount = Number(response.resolvedCount || 0);
    historyState.manualRuleCount = Number(response.manualRuleCount || historyState.manualRuleCount || 0);
  }

  function createExecutionRecord(payload){
    return {
      parsed: payload.parsed,
      previewOnly: !!payload.previewOnly,
      rosterRows: Array.isArray(payload.rosterRows) ? payload.rosterRows : [],
      sheetMeta: payload.sheetMeta || {},
      pendingItems: Array.isArray(payload.pendingItems) ? payload.pendingItems : [],
      manualRules: Array.isArray(payload.manualRules) ? payload.manualRules : [],
      result: payload.result || matcher.buildResult([], { events: [] }, {}),
      historyState: payload.historyState || createHistoryState()
    };
  }

  function buildStoragePayload(panelState, execution){
    var user = auth.getCurrentUser() || {};
    var result = execution.result || {};
    return {
      sheetTitle: panelState.selectedSheet,
      previewOnly: !!execution.previewOnly,
      actorEmail: String(user.email || '').trim(),
      actorName: String(user.name || '').trim(),
      executedAt: new Date().toISOString(),
      summary: {
        joinedCount: result.joinedCount,
        leftCount: result.leftCount,
        attendingCount: result.attendingCount,
        finalLeftCount: result.finalLeftCount,
        missingCount: result.missingCount,
        currentUnmatchedCount: result.currentUnmatchedItems.length,
        resolvedPendingCount: result.pendingResolvedItems.length,
        manualResolvedCount: result.manualResolvedItems.length,
        excludedByRuleCount: result.excludedByRuleItems.length
      },
      currentUnmatchedItems: result.currentUnmatchedItems,
      resolvedPendingItems: result.pendingResolvedItems.concat(result.manualResolvedItems),
      snapshot: buildSnapshotPayload(panelState, execution)
    };
  }

  function buildSnapshotPayload(panelState, execution){
    var historyState = Object.assign(createHistoryState(), clonePlain(execution.historyState || {}));
    historyState.pendingCount = Array.isArray(execution.pendingItems) ? execution.pendingItems.length : 0;
    historyState.manualRuleCount = Array.isArray(execution.manualRules) ? execution.manualRules.length : 0;
    historyState.openCount = Math.max(
      Number(historyState.openCount || 0),
      historyState.pendingCount
    );

    return {
      version: 1,
      sheetTitle: panelState.selectedSheet,
      savedAt: new Date().toISOString(),
      execution: {
        previewOnly: !!execution.previewOnly,
        parsed: parser.compactParsed(execution.parsed || { events: [], joinedCount: 0, leftCount: 0 }),
        rosterRows: compactRosterRows(execution.rosterRows || []),
        sheetMeta: clonePlain(execution.sheetMeta || {}),
        pendingItems: clonePlain(execution.pendingItems || []),
        manualRules: clonePlain(execution.manualRules || []),
        historyState: historyState
      }
    };
  }

  function compactRosterRows(rows){
    return (Array.isArray(rows) ? rows : []).map(function(row){
      return {
        rowNumber: Number(row && row.rowNumber || 0),
        sheetTitle: String(row && row.sheetTitle || '').trim(),
        name: String(row && row.name || '').trim(),
        phone: String(row && row.phone || '').trim(),
        nameNormalized: utils.normalizeName(row && (row.nameNormalized || row.name) || '')
      };
    }).filter(function(row){
      return row.rowNumber && row.nameNormalized;
    });
  }

  function clonePlain(value){
    return JSON.parse(JSON.stringify(value == null ? null : value));
  }

  function restoreExecutionFromSnapshot(snapshot, fallbackSheetTitle){
    var container = snapshot && snapshot.execution ? snapshot.execution : snapshot;
    var parsed = clonePlain(container && container.parsed ? container.parsed : { events: [], joinedCount: 0, leftCount: 0 });
    var rosterRows = Array.isArray(container && container.rosterRows) ? clonePlain(container.rosterRows) : [];
    var pendingItems = Array.isArray(container && container.pendingItems) ? clonePlain(container.pendingItems) : [];
    var manualRules = Array.isArray(container && container.manualRules) ? clonePlain(container.manualRules) : [];
    var historyState = Object.assign(createHistoryState(), clonePlain(container && container.historyState ? container.historyState : {}));
    var sheetTitle = String((snapshot && snapshot.sheetTitle) || fallbackSheetTitle || '').trim();

    historyState.enabled = true;
    historyState.synced = true;
    historyState.pendingCount = pendingItems.length;
    historyState.manualRuleCount = manualRules.length;
    historyState.openCount = Math.max(Number(historyState.openCount || 0), pendingItems.length);

    return createExecutionRecord({
      parsed: parsed,
      previewOnly: !!(container && container.previewOnly),
      rosterRows: rosterRows,
      sheetMeta: container && container.sheetMeta ? container.sheetMeta : {},
      pendingItems: pendingItems,
      manualRules: manualRules,
      result: matcher.buildResult(rosterRows, parsed, {
        pendingItems: pendingItems,
        manualRules: manualRules,
        sheetTitle: sheetTitle
      }),
      historyState: historyState
    });
  }

  function buildResultSections(execution){
    var result = execution.result;
    var historyState = execution.historyState || createHistoryState();

    return [
      {
        title: '명단 기준 상태',
        headers: ['row', 'name', 'phone', 'status', 'matched', 'notes'],
        rows: result.report,
        groups: [
          { key: 'all', label: '전체', rows: result.report },
          { key: 'entered', label: '입장', rows: result.reportEntered },
          { key: 'left', label: '퇴장', rows: result.reportLeft },
          { key: 'missing', label: '미입장', rows: result.reportMissing }
        ]
      },
      {
        title: '현재 미매칭',
        headers: ['category', 'label', 'reason'],
        rows: result.currentUnmatchedRows
      },
      {
        title: '이전 미매칭 재매칭',
        headers: ['row', 'name', 'phone', 'label', 'reason'],
        rows: result.pendingResolvedRows
      },
      {
        title: '저장된 수동 매칭',
        headers: ['row', 'name', 'phone', 'label', 'reason'],
        rows: result.manualResolvedRows
      },
      {
        title: '저장된 코칭스태프 제외',
        headers: ['label', 'reason'],
        rows: result.excludedByRuleRows
      },
      {
        title: '명단 외 입장',
        headers: ['name/last4'],
        rows: result.extra
      },
      {
        title: '퇴장 로그(최종)',
        headers: ['name/last4'],
        rows: result.leavers
      },
      {
        title: execution.previewOnly ? '예상 반영 값' : '실제 반영 값',
        headers: ['range', 'value'],
        rows: result.updates.map(function(item){
          return [item.range, item.values && item.values[0] ? item.values[0][0] : ''];
        })
      },
      {
        summary: [
          { label: '입장 로그', value: result.joinedCount },
          { label: '퇴장 로그', value: result.leftCount },
          { label: '명단 입장', value: result.attendingCount },
          { label: '미입장', value: result.missingCount },
          { label: '현재 미매칭', value: result.currentUnmatchedItems.length },
          { label: '수동 규칙', value: historyState.manualRuleCount || 0 },
          { label: '열린 대기', value: historyState.openCount || 0 },
          { label: '저장 상태', value: buildSyncLabel(historyState, execution.previewOnly) }
        ]
      }
    ];
  }

  function renderRunResult(panelState){
    var execution = panelState.lastExecution;

    if(!execution){
      ui.renderResults(panelState.el, []);
      syncPanelButtons(panelState);
      return;
    }

    ui.renderResults(panelState.el, buildResultSections(execution));
    syncPanelButtons(panelState);
  }

  function buildSyncLabel(historyState, previewOnly){
    if(previewOnly) return '미리보기';
    if(historyState.syncError) return '저장 경고';
    if(historyState.synced) return '저장 완료';
    if(historyState.syncSkipped) return '저장 안 함';
    return historyState.enabled ? '대기 중' : (historyState.message || '백엔드 미설정');
  }

  function openManualReview(panelId){
    var panelState = state.panels.get(panelId);
    var execution;
    var result;
    var missingItems;

    if(!panelState || !panelState.lastExecution){
      ui.openCustomModal({
        panelId: panelId,
        title: '미매칭 수동 처리',
        subtitle: '먼저 실행 결과가 있어야 합니다.',
        render: function(body){
          ui.appendEmptyState(body, '이 패널에서 먼저 입장 체크를 실행해 주세요.');
        }
      });
      return;
    }

    execution = panelState.lastExecution;
    result = execution.result;
    missingItems = result.missingRosterItems || [];

    ui.openCustomModal({
      type: 'manual-review',
      panelId: panelId,
      title: '미매칭 수동 처리',
      subtitle: panelState.selectedSheet + ' · 현재 미매칭 ' + result.currentUnmatchedItems.length + '건',
      render: function(body){
        var toolbar = document.createElement('div');
        var list = document.createElement('div');
        var notice;
        var toggleAllBtn;
        var bulkExcludeBtn;
        var selectedQueueKeys = new Set();
        var checkboxEntries = [];

        function syncSelectionUi(){
          var selectedCount = selectedQueueKeys.size;
          var totalCount = result.currentUnmatchedItems.length;

          if(toggleAllBtn){
            toggleAllBtn.textContent = selectedCount && selectedCount === totalCount ? '전체 선택 해제' : '전체 선택';
          }
          if(bulkExcludeBtn){
            bulkExcludeBtn.disabled = execution.previewOnly || !selectedCount;
            bulkExcludeBtn.textContent = selectedCount ? ('선택 ' + selectedCount + '건 코칭스태프로 제외') : '선택 항목 코칭스태프로 제외';
          }
          checkboxEntries.forEach(function(entry){
            entry.checkbox.checked = selectedQueueKeys.has(entry.item.queueKey);
          });
        }

        if(execution.previewOnly){
          notice = document.createElement('div');
          notice.className = 'modal-note';
          notice.textContent = '미리보기 결과에서는 수동 처리 저장을 할 수 없습니다. 실제 실행 후 다시 시도해 주세요.';
          body.appendChild(notice);
        }

        if(!result.currentUnmatchedItems.length){
          ui.appendEmptyState(body, '현재 미매칭 항목이 없습니다.');
          return;
        }

        toolbar.className = 'panel-secondary-actions';
        toggleAllBtn = document.createElement('button');
        toggleAllBtn.type = 'button';
        toggleAllBtn.className = 'btn';
        toggleAllBtn.addEventListener('click', function(){
          if(selectedQueueKeys.size === result.currentUnmatchedItems.length){
            selectedQueueKeys.clear();
          } else {
            result.currentUnmatchedItems.forEach(function(item){
              selectedQueueKeys.add(item.queueKey);
            });
          }
          syncSelectionUi();
        });

        bulkExcludeBtn = document.createElement('button');
        bulkExcludeBtn.type = 'button';
        bulkExcludeBtn.className = 'btn';
        bulkExcludeBtn.addEventListener('click', function(){
          var items = result.currentUnmatchedItems.filter(function(item){
            return selectedQueueKeys.has(item.queueKey);
          });

          if(!items.length) return;
          if(window.confirm('선택한 ' + items.length + '건을 코칭스태프로 제외할까요?')){
            applyBulkExclude(panelState.id, items);
          }
        });

        toolbar.appendChild(toggleAllBtn);
        toolbar.appendChild(bulkExcludeBtn);
        body.appendChild(toolbar);

        list.className = 'manual-review-list';
        result.currentUnmatchedItems.forEach(function(item){
          var card = document.createElement('div');
          var selectionWrap = document.createElement('label');
          var checkbox = document.createElement('input');
          var checkboxText = document.createElement('span');
          var title = document.createElement('strong');
          var meta1 = document.createElement('div');
          var meta2 = document.createElement('div');
          var actions = document.createElement('div');
          var matchBtn = document.createElement('button');
          var excludeBtn = document.createElement('button');

          card.className = 'manual-card';
          selectionWrap.className = 'check-row';
          checkbox.type = 'checkbox';
          checkbox.addEventListener('change', function(){
            if(checkbox.checked){
              selectedQueueKeys.add(item.queueKey);
            } else {
              selectedQueueKeys.delete(item.queueKey);
            }
            syncSelectionUi();
          });
          checkboxText.textContent = '코칭스태프 일괄 제외 대상에 포함';
          title.textContent = item.label;
          meta1.className = 'manual-meta';
          meta2.className = 'manual-meta';
          actions.className = 'manual-card-actions';
          meta1.textContent = '분류: ' + getUnmatchedCategoryText(item.category);
          meta2.textContent = '사유: ' + item.reason;

          matchBtn.type = 'button';
          matchBtn.className = 'btn btn-primary';
          matchBtn.textContent = missingItems.length ? '미입장 수강생으로 입장 처리' : '선택 가능한 미입장 수강생 없음';
          matchBtn.disabled = execution.previewOnly || !missingItems.length;
          matchBtn.addEventListener('click', function(){
            openManualMatchPicker(panelState.id, item);
          });

          excludeBtn.type = 'button';
          excludeBtn.className = 'btn';
          excludeBtn.textContent = '코칭스태프로 제외';
          excludeBtn.disabled = execution.previewOnly;
          excludeBtn.addEventListener('click', function(){
            if(window.confirm("'" + item.label + "' 항목을 코칭스태프로 제외할까요?")){
              applyManualAction(panelState.id, item, 'exclude-staff', null);
            }
          });

          selectionWrap.appendChild(checkbox);
          selectionWrap.appendChild(checkboxText);
          checkboxEntries.push({ item: item, checkbox: checkbox });
          card.appendChild(selectionWrap);
          card.appendChild(title);
          card.appendChild(meta1);
          card.appendChild(meta2);
          actions.appendChild(matchBtn);
          actions.appendChild(excludeBtn);
          card.appendChild(actions);
          list.appendChild(card);
        });

        body.appendChild(list);
        syncSelectionUi();
      }
    });
  }

  function openManualMatchPicker(panelId, queueItem){
    var panelState = state.panels.get(panelId);
    var execution;
    var candidates;

    if(!panelState || !panelState.lastExecution) return;
    execution = panelState.lastExecution;
    candidates = (execution.result && execution.result.missingRosterItems) || [];

    ui.openCustomModal({
      type: 'manual-match-picker',
      panelId: panelId,
      title: '수동 매칭 대상 선택',
      subtitle: queueItem.label + ' · 미입장 수강생에서 선택',
      render: function(body){
        var list = document.createElement('div');

        if(!candidates.length){
          ui.appendEmptyState(body, '선택 가능한 미입장 수강생이 없습니다.');
          return;
        }

        list.className = 'modal-list';
        candidates.forEach(function(candidate){
          var button = document.createElement('button');
          var meta = document.createElement('div');

          button.type = 'button';
          button.className = 'modal-list-btn';
          button.textContent = candidate.rowNumber + '행 · ' + utils.formatPersonLabel(candidate.name, candidate.phone);
          meta.className = 'list-item-subtext';
          meta.textContent = '현재 상태: ' + candidate.status + (candidate.notes ? ' · ' + candidate.notes : '');

          button.appendChild(meta);
          button.addEventListener('click', function(){
            applyManualAction(panelId, queueItem, 'match-student', candidate);
          });
          list.appendChild(button);
        });

        body.appendChild(list);
      }
    });
  }

  async function applyManualAction(panelId, queueItem, actionType, targetRow){
    var panelState = state.panels.get(panelId);
    var execution;
    var response;
    var snapshotError = '';

    if(!panelState || !panelState.lastExecution) return;
    execution = panelState.lastExecution;

    try {
      ui.showLoading('수동 처리 저장 중', '미매칭 처리 결과를 저장하고 있습니다.');
      ui.setPanelError(panelState.el, '');
      ui.setPanelStatus(panelState.el, '수동 처리 저장 중...', 'warn');

      response = await history.applyManualAction({
        sheetTitle: panelState.selectedSheet,
        item: queueItem,
        actionType: actionType,
        targetRow: targetRow ? {
          rowNumber: targetRow.rowNumber,
          name: targetRow.name,
          phone: targetRow.phone
        } : null
      });

      if(response && response.enabled === false){
        throw new Error(response.message || '수동 처리를 저장하지 못했습니다.');
      }

      if(actionType === 'match-student' && targetRow){
        await sheets.writeUpdates(matcher.createStatusUpdatesForRow(targetRow, execution.rosterRows));
      }

      execution.manualRules = upsertManualRule(execution.manualRules, response.manualRule || {});
      execution.pendingItems = execution.pendingItems.filter(function(item){
        return item.queueKey !== queueItem.queueKey;
      });
      execution.result = matcher.buildResult(execution.rosterRows, execution.parsed, {
        pendingItems: execution.pendingItems,
        manualRules: execution.manualRules,
        sheetTitle: panelState.selectedSheet
      });
      if(execution.historyState){
        execution.historyState.pendingCount = Math.max(0, Number(execution.historyState.pendingCount || 0) - 1);
        execution.historyState.openCount = Math.max(0, Number(execution.historyState.openCount || 0) - 1);
        execution.historyState.manualRuleCount = execution.manualRules.length;
        execution.historyState.synced = true;
        execution.historyState.syncError = '';
      }

      try {
        await history.saveSnapshot({
          sheetTitle: panelState.selectedSheet,
          previewOnly: execution.previewOnly,
          snapshot: buildSnapshotPayload(panelState, execution)
        });
      } catch (error) {
        snapshotError = toUserMessage(error);
      }

      if(execution.historyState){
        execution.historyState.syncError = snapshotError;
      }

      renderRunResult(panelState);
      if(snapshotError){
        ui.setPanelStatus(panelState.el, actionType === 'match-student' ? '수동 매칭 저장 완료 / 저장 경고' : '코칭스태프 제외 저장 완료 / 저장 경고', 'warn');
        ui.setPanelError(panelState.el, snapshotError);
      } else {
        ui.setPanelStatus(panelState.el, actionType === 'match-student' ? '수동 매칭 저장 완료' : '코칭스태프 제외 저장 완료', 'ok');
      }
      openManualReview(panelId);
    } catch (error) {
      ui.setPanelStatus(panelState.el, '수동 처리 실패', 'bad');
      ui.setPanelError(panelState.el, toUserMessage(error));
    } finally {
      ui.hideLoading();
    }
  }

  async function applyBulkExclude(panelId, queueItems){
    var panelState = state.panels.get(panelId);
    var execution;
    var snapshotError = '';
    var selectedKeys;
    var responses = [];
    var index;

    if(!panelState || !panelState.lastExecution || !Array.isArray(queueItems) || !queueItems.length){
      return;
    }

    execution = panelState.lastExecution;
    selectedKeys = new Set(queueItems.map(function(item){
      return item.queueKey;
    }).filter(Boolean));

    try {
      ui.showLoading('코칭스태프 일괄 제외 저장 중', '선택한 미매칭 항목을 한 번에 제외하고 있습니다.');
      ui.setPanelError(panelState.el, '');
      ui.setPanelStatus(panelState.el, '일괄 제외 저장 중...', 'warn');

      for(index = 0; index < queueItems.length; index += 1){
        var response = await history.applyManualAction({
          sheetTitle: panelState.selectedSheet,
          item: queueItems[index],
          actionType: 'exclude-staff',
          targetRow: null
        });

        if(response && response.enabled === false){
          throw new Error(response.message || '일괄 제외를 저장하지 못했습니다.');
        }
        responses.push(response);
      }

      responses.forEach(function(response){
        execution.manualRules = upsertManualRule(execution.manualRules, response && response.manualRule ? response.manualRule : {});
      });
      execution.pendingItems = execution.pendingItems.filter(function(item){
        return !selectedKeys.has(item.queueKey);
      });
      execution.result = matcher.buildResult(execution.rosterRows, execution.parsed, {
        pendingItems: execution.pendingItems,
        manualRules: execution.manualRules,
        sheetTitle: panelState.selectedSheet
      });

      if(execution.historyState){
        execution.historyState.pendingCount = Math.max(0, Number(execution.historyState.pendingCount || 0) - responses.length);
        execution.historyState.openCount = Math.max(0, Number(execution.historyState.openCount || 0) - responses.length);
        execution.historyState.manualRuleCount = execution.manualRules.length;
        execution.historyState.synced = true;
        execution.historyState.syncError = '';
      }

      try {
        await history.saveSnapshot({
          sheetTitle: panelState.selectedSheet,
          previewOnly: execution.previewOnly,
          snapshot: buildSnapshotPayload(panelState, execution)
        });
      } catch (error) {
        snapshotError = toUserMessage(error);
      }

      if(execution.historyState){
        execution.historyState.syncError = snapshotError;
      }

      renderRunResult(panelState);
      if(snapshotError){
        ui.setPanelStatus(panelState.el, '코칭스태프 일괄 제외 완료 / 저장 경고', 'warn');
        ui.setPanelError(panelState.el, snapshotError);
      } else {
        ui.setPanelStatus(panelState.el, '코칭스태프 일괄 제외 완료', 'ok');
      }
      openManualReview(panelId);
    } catch (error) {
      ui.setPanelStatus(panelState.el, '일괄 제외 실패', 'bad');
      ui.setPanelError(panelState.el, toUserMessage(error));
    } finally {
      ui.hideLoading();
    }
  }

  function upsertManualRule(rules, manualRule){
    var normalized = matcher.normalizeManualRule(manualRule);
    var filtered;

    if(!normalized || !normalized.queueKey){
      return Array.isArray(rules) ? rules.slice() : [];
    }

    filtered = (Array.isArray(rules) ? rules : []).filter(function(rule){
      var current = matcher.normalizeManualRule(rule);
      return !current || current.queueKey !== normalized.queueKey;
    });
    filtered.push(manualRule);
    return filtered;
  }

  async function openStorageOverview(){
    var response;

    try {
      ui.showLoading('서버 저장 데이터 불러오는 중', '시트별 저장 현황을 가져오고 있습니다.');
      response = await history.getStorageOverview();
      ui.openCustomModal({
        type: 'storage-overview',
        title: '서버 저장 데이터',
        subtitle: '백엔드 시트에 저장된 시트별 현황',
        render: function(body){
          var list = document.createElement('div');

          if(!response.sheets || !response.sheets.length){
            ui.appendEmptyState(body, '저장된 시트 데이터가 없습니다.');
            return;
          }

          body.appendChild(ui.buildMetricGrid([
            { label: '저장 시트', value: response.sheets.length },
            { label: '전체 열린 미매칭', value: response.totalOpenCount || 0 },
            { label: '수동 규칙', value: response.totalManualRuleCount || 0 },
            { label: '최근 저장', value: response.lastSavedAt || '-' }
          ]));

          list.className = 'modal-list';
          response.sheets.forEach(function(item){
            var button = document.createElement('button');
            var meta = document.createElement('div');

            button.type = 'button';
            button.className = 'modal-list-btn';
            button.textContent = item.sheetTitle;
            meta.className = 'list-item-subtext';
            meta.textContent = '열린 미매칭 ' + item.openCount + '건 · 수동 규칙 ' + item.manualRuleCount + '건 · 최신 저장 ' + (item.lastSavedAt || '-');
            button.appendChild(meta);
            button.addEventListener('click', function(){
              openStoredSheetData(item.sheetTitle);
            });
            list.appendChild(button);
          });

          body.appendChild(list);
        }
      });
    } catch (error) {
      ui.setAppNotice('서버 저장 데이터를 불러오지 못했습니다. ' + toUserMessage(error), 'bad');
    } finally {
      ui.hideLoading();
    }
  }

  async function openStoredSheetData(sheetTitle, panelId){
    var data;
    var panelState;
    if(!sheetTitle){
      ui.setAppNotice('저장 데이터를 볼 시트를 먼저 선택해 주세요.', 'warn');
      return;
    }

    try {
      ui.showLoading(panelId ? '저장 데이터 불러오는 중' : '저장 데이터 상세 불러오는 중', panelId ? '저장된 최신 실행 상태를 패널에 복원하고 있습니다.' : '선택한 시트의 저장 상태를 읽고 있습니다.');
      data = await history.getStoredSheetData(sheetTitle);
      if(panelId){
        panelState = state.panels.get(panelId);
        if(!panelState) return;
        loadStoredSheetIntoPanel(panelState, sheetTitle, data);
        return;
      }

      ui.openCustomModal({
        type: 'stored-sheet-data',
        title: '서버 저장 데이터',
        subtitle: sheetTitle,
        render: function(body){
          body.appendChild(ui.buildMetricGrid([
            { label: '열린 미매칭', value: data.summary ? data.summary.openCount : 0 },
            { label: '수동 규칙', value: data.summary ? data.summary.manualRuleCount : 0 },
            { label: '저장 상태', value: data.summary && data.summary.snapshotReady ? '저장됨' : '없음' },
            { label: '최근 저장', value: data.summary ? (data.summary.lastSavedAt || '-') : '-' }
          ]));

          appendStoredExecutionSections(body, sheetTitle, data);
        }
      });
    } catch (error) {
      ui.setAppNotice('저장된 시트 데이터를 불러오지 못했습니다. ' + toUserMessage(error), 'bad');
    } finally {
      ui.hideLoading();
    }
  }

  function loadStoredSheetIntoPanel(panelState, sheetTitle, data){
    if(!data || !data.snapshot){
      throw new Error('이 시트에는 저장된 최신 실행 상태가 없습니다.');
    }

    panelState.selectedSheet = sheetTitle;
    panelState.files = [];
    panelState.lastExecution = restoreExecutionFromSnapshot(data.snapshot, sheetTitle);

    ui.renderSelectedSheet(panelState.el, sheetTitle);
    ui.renderFileSummary(panelState.el, []);
    ui.setPanelError(panelState.el, '');
    renderRunResult(panelState);
    ui.setPanelStatus(panelState.el, '저장 데이터 불러옴', 'ok');
    ui.closeModal();
  }

  function appendStoredExecutionSections(root, sheetTitle, data){
    var execution;

    if(!data || !data.snapshot){
      ui.appendEmptyState(root, '저장된 최신 실행 상태가 없습니다.');
      return;
    }

    execution = restoreExecutionFromSnapshot(data.snapshot, sheetTitle);
    buildResultSections(execution).forEach(function(section){
      if(section.summary){
        root.appendChild(ui.buildMetricGrid(section.summary));
        return;
      }

      appendStorageSection(root, section.title, section.headers || [], section.rows || []);
    });
  }

  function appendStorageSection(root, titleText, headers, rows){
    var section = document.createElement('section');
    var title = document.createElement('h4');

    section.className = 'modal-section';
    title.className = 'modal-section-title';
    title.textContent = titleText;
    section.appendChild(title);

    if(rows && rows.length){
      section.appendChild(ui.buildTableWrap(headers, rows));
    } else {
      ui.appendEmptyState(section, '저장된 데이터가 없습니다.');
    }

    root.appendChild(section);
  }

  function getUnmatchedCategoryText(category){
    if(category === 'outside-roster') return '명단 외';
    if(category === 'duplicate-name-no-number') return '동명이인 / 번호 없음';
    if(category === 'duplicate-name-mismatch') return '동명이인 / 번호 불일치';
    if(category === 'duplicate-name-collision') return '동명이인 / 번호 중복';
    return category || '';
  }

  function toUserMessage(error){
    var message = '';

    if(typeof error === 'string'){
      message = error;
    } else if(error && typeof error.message === 'string' && error.message.trim()){
      message = error.message;
    } else if(error && typeof error.error === 'string' && error.error.trim()){
      message = error.error;
    } else if(error && error.cause && typeof error.cause.message === 'string' && error.cause.message.trim()){
      message = error.cause.message;
    } else if(error && typeof error === 'object'){
      try {
        message = JSON.stringify(error);
      } catch (jsonError) {
        message = String(error || '');
      }
    } else {
      message = String(error || '');
    }

    if(!message || message === '{}' || message === '[object Object]'){
      message = '알 수 없는 오류가 발생했습니다.';
    }

    if(message.indexOf('Unsupported action:') === 0){
      return '현재 연결된 Apps Script 웹앱이 구버전입니다. 최신 apps-script/Code.gs를 붙여넣고 새 버전으로 다시 배포해 주세요. (' + message + ')';
    }
    if(message.indexOf('MATCH_BACKEND_TOKEN script property is missing.') >= 0){
      return 'Apps Script 프로젝트의 스크립트 속성에 MATCH_BACKEND_TOKEN 값이 없습니다. 현재 Vercel의 KAKAO_CHECK_APPS_SCRIPT_TOKEN 값과 같은 값을 넣고 Apps Script 웹앱을 다시 배포해 주세요.';
    }
    if(message.indexOf('Invalid backend token.') >= 0){
      return 'Vercel의 KAKAO_CHECK_APPS_SCRIPT_TOKEN 값과 Apps Script의 MATCH_BACKEND_TOKEN 값이 서로 다릅니다.';
    }
    if(message.indexOf('Sheet not found:') === 0){
      return '선택한 시트를 찾지 못했습니다. 시트명이 바뀌었는지 확인해 주세요. (' + message.replace('Sheet not found:', '').trim() + ')';
    }
    if(message.indexOf('You do not have access to this sheet.') >= 0){
      return '현재 계정으로는 이 시트에 접근할 수 없습니다.';
    }
    return message;
  }

  document.addEventListener('DOMContentLoaded', bootstrap);
})();
