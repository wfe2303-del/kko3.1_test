(function(){
  var auth = window.KakaoCheckAuth;
  var sheets = window.KakaoCheckSheets;
  var parser = window.KakaoCheckParser;
  var matcher = window.KakaoCheckMatcher;
  var history = window.KakaoCheckHistory;
  var ui = window.KakaoCheckUI;
  var utils = window.KakaoCheckUtils;

  var state = {
    nextPanelId: 1,
    sheetTitles: [],
    panels: new Map()
  };

  function bootstrap(){
    ui.setLoginReady(false);
    bindTopLevelEvents();
    auth.onChange(function(payload){
      ui.setAuthState(payload);
      ui.setAuthError('');
      if(payload && payload.accessToken){
        loadSheetTitles();
      }
    });

    auth.init().then(function(){
      ui.setLoginReady(true);
      ui.setAuthState({ accessToken: null, user: null });
      ui.setAuthError('');
    }).catch(function(error){
      ui.setLoginReady(false);
      ui.setAuthError(error.message);
    });

    addPanel();
  }

  function bindTopLevelEvents(){
    document.getElementById('loginBtn').addEventListener('click', async function(){
      try {
        ui.setAuthError('');
        ui.setLoginReady(false);
        await auth.login();
      } catch (error) {
        ui.setAuthError(error.message);
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

    document.getElementById('addPanelBtn').addEventListener('click', function(){
      addPanel();
    });
  }

  async function loadSheetTitles(){
    try {
      state.sheetTitles = await sheets.listSheetTitles();
    } catch (error) {
      ui.setAuthError(error.message);
    }
  }

  function addPanel(){
    var panelId = state.nextPanelId;
    state.nextPanelId += 1;

    var panelEl = ui.createPanelElement(panelId);
    var panelState = {
      id: panelId,
      el: panelEl,
      files: [],
      selectedSheet: ''
    };

    state.panels.set(panelId, panelState);
    bindPanelEvents(panelState);
    ui.renderSelectedSheet(panelEl, '');
    ui.renderFileSummary(panelEl, []);
    ui.setPanelStatus(panelEl, '', '');
    document.getElementById('panelsRoot').appendChild(panelEl);
    refreshRemoveButtons();
  }

  function removePanel(panelId){
    if(state.panels.size <= 1) return;
    var panelState = state.panels.get(panelId);
    if(!panelState) return;

    panelState.el.remove();
    state.panels.delete(panelId);

    if(ui.getModalState() && ui.getModalState().panelId === panelId){
      ui.closeModal();
    }

    refreshRemoveButtons();
  }

  function refreshRemoveButtons(){
    state.panels.forEach(function(panelState){
      var btn = utils.qs('.panel-remove-btn', panelState.el);
      if(!btn) return;
      btn.disabled = state.panels.size <= 1;
    });
  }

  function bindPanelEvents(panelState){
    var el = panelState.el;
    var fileInput = utils.qs('.js-file-input', el);
    var runBtn = utils.qs('.js-run-btn', el);
    var demoBtn = utils.qs('.js-demo-btn', el);
    var removeBtn = utils.qs('.panel-remove-btn', el);
    var openSheetBtn = utils.qs('.js-open-sheet-modal-btn', el);
    var openFileBtn = utils.qs('.js-open-file-modal-btn', el);

    openSheetBtn.addEventListener('click', function(){
      openSheetPicker(panelState.id);
    });

    openFileBtn.addEventListener('click', function(){
      openFileManager(panelState.id);
    });

    fileInput.addEventListener('change', function(event){
      var incoming = Array.from(event.target.files || []);
      if(!incoming.length) return;

      panelState.files = appendUniqueFiles(panelState.files, incoming);
      ui.renderFileSummary(el, panelState.files);
      ui.setPanelStatus(el, panelState.files.length + '개 파일 준비', '');
      event.target.value = '';

      var modal = ui.getModalState();
      if(modal && modal.type === 'file-manager' && modal.panelId === panelState.id){
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
      title: panelState.selectedSheet ? panelState.selectedSheet + ' 시트 선택' : '시트 선택',
      subtitle: '',
      titles: state.sheetTitles,
      selectedTitle: panelState.selectedSheet,
      onSelect: function(title){
        panelState.selectedSheet = title;
        ui.renderSelectedSheet(panelState.el, panelState.selectedSheet);
        ui.setPanelStatus(panelState.el, '시트 선택 완료', '');
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
      title: panelState.selectedSheet ? panelState.selectedSheet + ' 로그 파일 관리' : '로그 파일 관리',
      subtitle: '',
      files: panelState.files,
      onAddRequest: function(){
        var input = utils.qs('.js-file-input', panelState.el);
        if(input) input.click();
      },
      onRemoveIndex: function(index){
        panelState.files.splice(index, 1);
        ui.renderFileSummary(panelState.el, panelState.files);
        ui.setPanelStatus(panelState.el, panelState.files.length ? '파일 목록 수정됨' : '파일 비어 있음', '');
      },
      onClearAll: function(){
        panelState.files = [];
        ui.renderFileSummary(panelState.el, panelState.files);
        ui.setPanelStatus(panelState.el, '파일 비어 있음', '');
      },
      getFreshOptions: function(){
        return buildFileModalOptions(panelState);
      }
    };
  }

  function appendUniqueFiles(existingFiles, incomingFiles){
    var merged = existingFiles.slice();
    var seen = new Set(existingFiles.map(function(file){
      return [file.name, file.size, file.lastModified].join('::');
    }));

    incomingFiles.forEach(function(file){
      var key = [file.name, file.size, file.lastModified].join('::');
      if(seen.has(key)) return;
      seen.add(key);
      merged.push(file);
    });

    return merged;
  }

  async function loadPendingItems(sheetTitle, historyState){
    if(!history || typeof history.listPending !== 'function'){
      historyState.enabled = false;
      historyState.message = '매칭 기록 백엔드가 연결되지 않았습니다.';
      return [];
    }

    var response = await history.listPending(sheetTitle);
    if(!response.enabled){
      historyState.enabled = false;
      historyState.message = response.message || '매칭 기록 백엔드가 설정되지 않았습니다.';
      return [];
    }

    historyState.enabled = true;
    historyState.pendingLoaded = true;
    historyState.pendingCount = Array.isArray(response.items) ? response.items.length : 0;
    return Array.isArray(response.items) ? response.items : [];
  }

  function buildHistoryPayload(panelState, result, previewOnly){
    var user = auth.getCurrentUser() || {};

    return {
      sheetTitle: panelState.selectedSheet,
      previewOnly: !!previewOnly,
      actorEmail: String(user.email || '').trim(),
      actorName: String(user.name || '').trim(),
      executedAt: new Date().toISOString(),
      files: panelState.files.map(function(file){
        return {
          name: file.name,
          size: file.size,
          lastModified: file.lastModified || null
        };
      }),
      summary: {
        joinedCount: result.joinedCount,
        leftCount: result.leftCount,
        attendingCount: result.attendingCount,
        finalLeftCount: result.finalLeftCount,
        missingCount: result.missingCount,
        currentUnmatchedCount: result.currentUnmatchedItems.length,
        resolvedPendingCount: result.pendingResolvedItems.length
      },
      currentUnmatchedItems: result.currentUnmatchedItems,
      resolvedPendingItems: result.pendingResolvedItems
    };
  }

  async function syncHistoryRun(panelState, result, previewOnly, historyState){
    if(previewOnly){
      historyState.syncSkipped = true;
      return;
    }

    if(!history || typeof history.syncRun !== 'function'){
      historyState.syncSkipped = true;
      return;
    }

    var response = await history.syncRun(buildHistoryPayload(panelState, result, previewOnly));
    if(!response.enabled){
      historyState.enabled = false;
      historyState.message = response.message || '매칭 기록 백엔드가 설정되지 않았습니다.';
      historyState.syncSkipped = true;
      return;
    }

    historyState.enabled = true;
    historyState.synced = !!response.synced;
    historyState.runId = String(response.runId || '').trim();
    historyState.openCount = Number(response.openCount || 0);
    historyState.openedCount = Number(response.openedCount || 0);
    historyState.resolvedCount = Number(response.resolvedCount || 0);
  }

  function buildHistorySummarySection(historyState, previewOnly){
    var rows = [];

    rows.push(['백엔드', historyState.enabled ? '연결됨' : (historyState.message || '미설정')]);
    rows.push(['불러온 기존 미매칭', historyState.pendingCount || 0]);
    rows.push(['이번 실행 신규 미매칭', historyState.openedCount || 0]);
    rows.push(['이번 실행 해소 건수', historyState.resolvedCount || 0]);
    rows.push(['현재 대기 건수', historyState.openCount || 0]);

    if(previewOnly){
      rows.push(['저장 여부', '미리보기라 저장 안 함']);
    } else if(historyState.synced){
      rows.push(['저장 여부', '저장 완료']);
    } else if(historyState.syncSkipped){
      rows.push(['저장 여부', '저장 안 함']);
    } else {
      rows.push(['저장 여부', '저장 실패']);
    }

    if(historyState.syncError){
      rows.push(['오류', historyState.syncError]);
    }

    return {
      title: '매칭 기록 동기화',
      headers: ['항목', '값'],
      rows: rows
    };
  }

  async function runPanel(panelId){
    var panelState = state.panels.get(panelId);
    if(!panelState) return;

    var previewOnly = utils.qs('.js-preview-only', panelState.el).checked;
    var historyState = {
      enabled: false,
      pendingLoaded: false,
      pendingCount: 0,
      openedCount: 0,
      resolvedCount: 0,
      openCount: 0,
      synced: false,
      syncSkipped: false,
      syncError: '',
      message: ''
    };

    try {
      ui.setPanelError(panelState.el, '');
      ui.setPanelStatus(panelState.el, '실행 중...', 'warn');

      if(!auth.getAccessToken()) throw new Error('먼저 Google 로그인을 해주세요.');
      if(!panelState.selectedSheet) throw new Error('시트를 먼저 선택해주세요.');
      if(!panelState.files.length) throw new Error('카카오 대화 로그 파일을 하나 이상 추가해주세요.');

      var chatText = await parser.combineFiles(panelState.files);
      var parsed = parser.parseChatText(chatText);
      var roster = await sheets.loadRosterRows(panelState.selectedSheet);
      var pendingItems = [];

      try {
        pendingItems = await loadPendingItems(panelState.selectedSheet, historyState);
      } catch (historyError) {
        historyState.syncError = historyError.message || String(historyError);
      }

      var result = matcher.buildResult(roster, parsed, {
        pendingItems: pendingItems,
        sheetTitle: panelState.selectedSheet
      });

      if(!previewOnly && result.updates.length){
        await sheets.writeUpdates(result.updates);
      }

      try {
        await syncHistoryRun(panelState, result, previewOnly, historyState);
      } catch (syncError) {
        historyState.syncError = syncError.message || String(syncError);
      }

      renderRunResult(panelState, result, previewOnly, historyState);

      if(historyState.syncError){
        ui.setPanelStatus(
          panelState.el,
          previewOnly ? '미리보기 완료 (기록 조회 경고)' : '체크 완료 (기록 저장 경고)',
          'warn'
        );
        ui.setPanelError(panelState.el, historyState.syncError);
      } else {
        ui.setPanelStatus(panelState.el, previewOnly ? '미리보기 완료' : '체크 완료', 'ok');
      }
    } catch (error) {
      ui.setPanelStatus(panelState.el, '실패', 'bad');
      ui.setPanelError(panelState.el, error.message || String(error));
      console.error(error);
    }
  }

  function runDemo(panelId){
    var panelState = state.panels.get(panelId);
    if(!panelState) return;

    var roster = [
      { rowNumber: 2, sheetTitle: '데모', name: '김하늘', phone: '010-1234-5678', nameNormalized: utils.normalizeName('김하늘') },
      { rowNumber: 3, sheetTitle: '데모', name: '박시연', phone: '010-3456-7890', nameNormalized: utils.normalizeName('박시연') },
      { rowNumber: 4, sheetTitle: '데모', name: '이하린', phone: '010-1717-1771', nameNormalized: utils.normalizeName('이하린') },
      { rowNumber: 5, sheetTitle: '데모', name: '최도윤', phone: '010-1111-0000', nameNormalized: utils.normalizeName('최도윤') },
      { rowNumber: 6, sheetTitle: '데모', name: '김하늘', phone: '010-9999-9999', nameNormalized: utils.normalizeName('김하늘') }
    ];

    var parsed = parser.parseChatText([
      '최도윤/0000님이 입장했습니다.',
      '김하늘5678님이 입장했습니다.',
      '이하린1771님이 들어왔습니다',
      '박시연님이 입장했습니다.',
      '진주연2514님이 입장했습니다.',
      '모델실험테스트님이 입장했습니다.',
      '김하늘5678님이 나갔습니다.',
      '김하늘5678님이 입장했습니다',
      '박시연님이 나갔습니다'
    ].join('\n'));

    var result = matcher.buildResult(roster, parsed, {
      pendingItems: [],
      sheetTitle: '데모'
    });
    ui.setPanelError(panelState.el, '');
    renderRunResult(panelState, result, true, {
      enabled: false,
      pendingLoaded: false,
      pendingCount: 0,
      openedCount: 0,
      resolvedCount: 0,
      openCount: 0,
      synced: false,
      syncSkipped: true,
      syncError: '',
      message: '데모 모드에서는 저장하지 않습니다.'
    });
    ui.setPanelStatus(panelState.el, '데모 완료', 'ok');
  }

  function renderRunResult(panelState, result, previewOnly, historyState){
    var sections = [
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
        title: previewOnly ? '예상 반영 값' : '실제 반영 값',
        headers: ['range', 'value'],
        rows: result.updates.map(function(item){
          return [item.range, item.values && item.values[0] ? item.values[0][0] : ''];
        })
      },
      {
        title: '로그 요약',
        headers: ['입장 로그 수', '퇴장 로그 수', '명단 입장 수', '명단 퇴장 수'],
        rows: [[result.joinedCount, result.leftCount, result.attendingCount, result.finalLeftCount]]
      },
      buildHistorySummarySection(historyState || {}, previewOnly)
    ];

    ui.renderResults(panelState.el, sections);
  }

  document.addEventListener('DOMContentLoaded', bootstrap);
})();
