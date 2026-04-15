(function(){
  var config = window.KakaoCheckConfig;

  function buildDisabledResponse(message, extra){
    return Object.assign({
      enabled: false,
      message: String(message || '매칭 기록 백엔드가 설정되지 않았습니다.')
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
      var message = payload && payload.error ? payload.error : (text || '매칭 기록 API 호출에 실패했습니다.');
      throw new Error(message);
    }

    return payload || {};
  }

  function safeParseJson(text){
    try {
      return JSON.parse(text);
    } catch (error) {
      return null;
    }
  }

  function listPending(sheetTitle){
    if(!config.matchHistoryEnabled){
      return Promise.resolve(buildDisabledResponse('', { items: [] }));
    }

    var url = '/api/match-records?action=listPending&sheetTitle=' + encodeURIComponent(sheetTitle || '');
    return apiFetch(url).then(function(payload){
      if(payload && payload.enabled === false){
        return buildDisabledResponse(payload.message, { items: [] });
      }
      return {
        enabled: true,
        items: Array.isArray(payload.items) ? payload.items : []
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
    }).then(function(response){
      if(response && response.enabled === false){
        return buildDisabledResponse(response.message, {
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
      }, response || {});
    });
  }

  window.KakaoCheckHistory = {
    listPending: listPending,
    syncRun: syncRun
  };
})();
