var auth = require('./_lib/auth');
var appsScript = require('./_lib/apps-script');

module.exports = async function handler(req, res) {
  setJsonHeaders(res);

  var session = auth.getSession(req);
  if(!session.configured){
    res.statusCode = 503;
    res.end(JSON.stringify({ error: 'Login is not configured. Set KAKAO_CHECK_SESSION_SECRET and KAKAO_CHECK_USERS_JSON.' }));
    return;
  }

  if(!session.user){
    res.statusCode = 401;
    res.end(JSON.stringify({ error: 'Login required.' }));
    return;
  }

  if(!appsScript.isConfigured()){
    res.statusCode = 200;
    res.end(JSON.stringify(buildDisabledResponse(getAction(req))));
    return;
  }

  try {
    var action = getAction(req);

    if(req.method !== 'GET' && req.method !== 'POST'){
      res.statusCode = 405;
      res.end(JSON.stringify({ error: 'Method not allowed' }));
      return;
    }

    if(action === 'listPending'){
      var payload = {
        sheetTitle: String((req.query && req.query.sheetTitle) || '').trim()
      };

      if(payload.sheetTitle && !auth.canAccessSheet(session.user, payload.sheetTitle)){
        res.statusCode = 403;
        res.end(JSON.stringify({ error: 'You do not have access to this sheet.' }));
        return;
      }

      var listResponse = await appsScript.callAppsScript('listPending', payload);
      if(listResponse && Array.isArray(listResponse.items)){
        listResponse.items = listResponse.items.filter(function(item){
          return auth.canAccessSheet(session.user, item && item.sheetTitle);
        });
      }
      res.statusCode = 200;
      res.end(JSON.stringify(listResponse || { enabled: true, items: [] }));
      return;
    }

    if(action === 'syncRun'){
      if(!session.user.canWrite){
        res.statusCode = 403;
        res.end(JSON.stringify({ error: 'This account cannot save attendance records.' }));
        return;
      }

      var body = readBody(req);
      var syncPayload = body.payload || {};
      var sheetTitle = String(syncPayload.sheetTitle || '').trim();

      if(sheetTitle && !auth.canAccessSheet(session.user, sheetTitle)){
        res.statusCode = 403;
        res.end(JSON.stringify({ error: 'You do not have access to this sheet.' }));
        return;
      }

      syncPayload.actorEmail = session.user.username;
      syncPayload.actorName = session.user.displayName;

      var syncResponse = await appsScript.callAppsScript('syncRun', syncPayload);
      res.statusCode = 200;
      res.end(JSON.stringify(syncResponse || { enabled: true, synced: true }));
      return;
    }

    if(action === 'health'){
      var healthResponse = await appsScript.callAppsScript('health', {});
      res.statusCode = 200;
      res.end(JSON.stringify(healthResponse || { enabled: true, ok: true }));
      return;
    }

    res.statusCode = 400;
    res.end(JSON.stringify({ error: 'Unsupported action.' }));
  } catch (error) {
    res.statusCode = 502;
    res.end(JSON.stringify({
      error: error && error.message ? error.message : 'Apps Script backend request failed.'
    }));
  }
};

function getAction(req){
  if(req.method === 'GET'){
    return String((req.query && req.query.action) || 'health').trim();
  }
  return String((readBody(req).action || 'health')).trim();
}

function readBody(req){
  if(!req || !req.body) return {};
  if(typeof req.body === 'object') return req.body;
  if(typeof req.body !== 'string') return {};

  try {
    return JSON.parse(req.body);
  } catch (error) {
    return {};
  }
}

function buildDisabledResponse(action){
  if(action === 'listPending'){
    return {
      enabled: false,
      message: 'Match history backend is not configured.',
      items: []
    };
  }

  if(action === 'syncRun'){
    return {
      enabled: false,
      message: 'Match history backend is not configured.',
      synced: false,
      openCount: 0,
      openedCount: 0,
      resolvedCount: 0
    };
  }

  return {
    enabled: false,
    message: 'Match history backend is not configured.'
  };
}

function setJsonHeaders(res){
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Pragma', 'no-cache');
}
