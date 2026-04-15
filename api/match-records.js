module.exports = async function handler(req, res) {
  var action = req.method === 'GET'
    ? String((req.query && req.query.action) || 'health').trim()
    : String((readBody(req).action || 'health')).trim();

  var payload = req.method === 'GET'
    ? {
        sheetTitle: String((req.query && req.query.sheetTitle) || '').trim()
      }
    : (readBody(req).payload || {});

  var appsScriptUrl = String(process.env.KAKAO_CHECK_APPS_SCRIPT_URL || '').trim();
  var appsScriptToken = String(process.env.KAKAO_CHECK_APPS_SCRIPT_TOKEN || '').trim();
  var configured = !!(appsScriptUrl && appsScriptToken);

  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Pragma', 'no-cache');

  if(req.method !== 'GET' && req.method !== 'POST'){
    res.statusCode = 405;
    res.end(JSON.stringify({ error: 'Method not allowed' }));
    return;
  }

  if(!configured){
    res.statusCode = 200;
    res.end(JSON.stringify(buildDisabledResponse(action)));
    return;
  }

  try {
    var response = await fetch(appsScriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify({
        token: appsScriptToken,
        action: action,
        payload: payload
      })
    });

    var text = await response.text();
    var data = safeParseJson(text);

    if(!response.ok){
      res.statusCode = response.status;
      res.end(JSON.stringify({
        error: data && data.error ? data.error : (text || 'Apps Script backend request failed')
      }));
      return;
    }

    res.statusCode = 200;
    res.end(JSON.stringify(data || { enabled: true }));
  } catch (error) {
    res.statusCode = 502;
    res.end(JSON.stringify({
      error: error && error.message ? error.message : 'Apps Script backend request failed'
    }));
  }
};

function readBody(req){
  if(!req || !req.body) return {};
  if(typeof req.body === 'object') return req.body;
  if(typeof req.body !== 'string') return {};
  return safeParseJson(req.body) || {};
}

function safeParseJson(text){
  try {
    return JSON.parse(text);
  } catch (error) {
    return null;
  }
}

function buildDisabledResponse(action){
  if(action === 'listPending'){
    return {
      enabled: false,
      message: '매칭 기록 백엔드가 설정되지 않았습니다.',
      items: []
    };
  }

  if(action === 'syncRun'){
    return {
      enabled: false,
      message: '매칭 기록 백엔드가 설정되지 않았습니다.',
      synced: false,
      openCount: 0,
      openedCount: 0,
      resolvedCount: 0
    };
  }

  return {
    enabled: false,
    message: '매칭 기록 백엔드가 설정되지 않았습니다.'
  };
}
