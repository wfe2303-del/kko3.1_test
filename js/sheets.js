(function(){
  var auth = window.KakaoCheckAuth;

  async function apiFetch(url, init){
    auth.requireToken();

    var options = init || {};
    var headers = Object.assign({
      'Accept': 'application/json'
    }, options.headers || {});

    var response = await fetch(url, Object.assign({}, options, {
      headers: headers,
      cache: 'no-store',
      credentials: 'same-origin',
      referrerPolicy: 'no-referrer'
    }));

    var text = await response.text();
    var data = safeParseJson(text);

    if(!response.ok){
      throw new Error(data && data.error ? data.error : (text || '시트 요청에 실패했습니다.'));
    }

    return data || {};
  }

  function listSheetTitles(){
    return apiFetch('/api/sheets?action=listSheetTitles').then(function(data){
      return Array.isArray(data.titles) ? data.titles : [];
    });
  }

  function loadRosterRows(sheetTitle){
    var url = '/api/sheets?action=loadRosterRows&sheetTitle=' + encodeURIComponent(sheetTitle || '');
    return apiFetch(url).then(function(data){
      return Array.isArray(data.rows) ? data.rows : [];
    });
  }

  function writeUpdates(updates){
    return apiFetch('/api/sheets', {
      method: 'POST',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        action: 'writeUpdates',
        updates: Array.isArray(updates) ? updates : []
      })
    });
  }

  function safeParseJson(text){
    try {
      return JSON.parse(text);
    } catch (error) {
      return null;
    }
  }

  window.KakaoCheckSheets = {
    listSheetTitles: listSheetTitles,
    loadRosterRows: loadRosterRows,
    writeUpdates: writeUpdates
  };
})();
