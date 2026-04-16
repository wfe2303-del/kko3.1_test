(function(){
  var config = window.KakaoCheckConfig;

  function buildDisabledResponse(message, extra){
    return Object.assign({
      enabled: false,
      message: String(message || '저장 실행 백엔드가 설정되지 않았습니다.')
    }, extra || {});
  }

  async function apiFetch(url, init){
    var response = await fetch(url, Object.assign({
      cache: 'no-store',
      credentials: 'same-origin',
      referrerPolicy: 'no-referrer',
      headers: {
        'Accept': 'application/json'
      }
    }, init || {}));
    var text = await response.text();
    var payload = text ? safeParseJson(text) : null;

    if(!response.ok){
      throw new Error(payload && payload.error ? payload.error : (text || '저장 실행 API 호출에 실패했습니다.'));
    }

    return payload || {};
  }

  function listPending(sheetTitle){
    if(!config.matchHistoryEnabled){
      return Promise.resolve(buildDisabledResponse('', { items: [], manualRules: [] }));
    }

    return apiFetch('/api/match-records?action=listPending&sheetTitle=' + encodeURIComponent(sheetTitle || ''))
      .then(function(payload){
        if(payload && payload.enabled === false){
          return buildDisabledResponse(payload.message, { items: [], manualRules: [] });
        }

        return {
          enabled: true,
          items: Array.isArray(payload.items) ? payload.items : [],
          manualRules: Array.isArray(payload.manualRules) ? payload.manualRules : []
        };
      });
  }

  function syncRun(payload){
    if(!config.matchHistoryEnabled){
      return Promise.resolve(buildDisabledResponse('', {
        synced: false,
        openCount: 0,
        openedCount: 0,
        resolvedCount: 0
      }));
    }

    return apiFetch('/api/match-records', {
      method: 'POST',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        action: 'syncRun',
        payload: payload || {}
      })
    }).then(function(payload){
      if(payload && payload.enabled === false){
        return buildDisabledResponse(payload.message, {
          synced: false,
          openCount: 0,
          openedCount: 0,
          resolvedCount: 0
        });
      }

      return Object.assign({
        enabled: true,
        synced: true,
        openCount: 0,
        openedCount: 0,
        resolvedCount: 0
      }, payload || {});
    });
  }

  function applyManualAction(payload){
    if(!config.matchHistoryEnabled){
      return Promise.resolve(buildDisabledResponse('', { items: [], manualRules: [] }));
    }

    return apiFetch('/api/match-records', {
      method: 'POST',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        action: 'applyManualAction',
        payload: payload || {}
      })
    });
  }

  function applyManualActionsBatch(payload){
    if(!config.matchHistoryEnabled){
      return Promise.resolve(buildDisabledResponse('', { items: [], manualRules: [] }));
    }

    return apiFetch('/api/match-records', {
      method: 'POST',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        action: 'applyManualActionsBatch',
        payload: payload || {}
      })
    });
  }

  function saveSnapshot(payload){
    return apiFetch('/api/match-records', {
      method: 'POST',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        action: 'saveSnapshot',
        payload: payload || {}
      })
    });
  }

  function getStorageOverview(){
    return apiFetch('/api/match-records?action=getStorageOverview');
  }

  function getStoredSheetData(sheetTitle){
    return apiFetch('/api/match-records?action=getStoredSheetData&sheetTitle=' + encodeURIComponent(sheetTitle || ''));
  }

  function health(){
    return apiFetch('/api/match-records?action=health');
  }

  function safeParseJson(text){
    try {
      return JSON.parse(text);
    } catch (error) {
      return null;
    }
  }

  window.KakaoCheckHistory = {
    applyManualAction: applyManualAction,
    applyManualActionsBatch: applyManualActionsBatch,
    getStorageOverview: getStorageOverview,
    getStoredSheetData: getStoredSheetData,
    health: health,
    listPending: listPending,
    saveSnapshot: saveSnapshot,
    syncRun: syncRun
  };
})();
