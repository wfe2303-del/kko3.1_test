var auth = require('./_lib/auth');
var appsScript = require('./_lib/apps-script');

module.exports = async function handler(req, res) {
  var session = auth.getSession(req);
  var action = getAction(req);

  setJsonHeaders(res);

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
    res.end(JSON.stringify(buildDisabledResponse(action)));
    return;
  }

  try {
    if(req.method !== 'GET' && req.method !== 'POST'){
      res.statusCode = 405;
      res.end(JSON.stringify({ error: 'Method not allowed.' }));
      return;
    }

    if(action === 'listPending'){
      await handleListPending(req, res, session.user);
      return;
    }

    if(action === 'syncRun'){
      await handleSyncRun(req, res, session.user);
      return;
    }

    if(action === 'saveSnapshot'){
      await handleSaveSnapshot(req, res, session.user);
      return;
    }

    if(action === 'applyManualAction'){
      await handleApplyManualAction(req, res, session.user);
      return;
    }

    if(action === 'applyManualActionsBatch'){
      await handleApplyManualActionsBatch(req, res, session.user);
      return;
    }

    if(action === 'getStorageOverview'){
      await handleStorageOverview(res, session.user);
      return;
    }

    if(action === 'getStoredSheetData'){
      await handleStoredSheetData(req, res, session.user);
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

async function handleListPending(req, res, user){
  var sheetTitle = String((req.query && req.query.sheetTitle) || '').trim();
  var response;

  if(sheetTitle && !auth.canAccessSheet(user, sheetTitle)){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'You do not have access to this sheet.' }));
    return;
  }

  response = await appsScript.callAppsScript('listPending', { sheetTitle: sheetTitle });
  response.items = filterSheetEntries(response.items, user);
  response.manualRules = filterSheetEntries(response.manualRules, user);
  res.statusCode = 200;
  res.end(JSON.stringify(response || { enabled: true, items: [], manualRules: [] }));
}

async function handleSyncRun(req, res, user){
  var body = readBody(req);
  var payload = body.payload || {};
  var sheetTitle = String(payload.sheetTitle || '').trim();
  var response;

  if(!user.canWrite){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'This account cannot save attendance records.' }));
    return;
  }

  if(sheetTitle && !auth.canAccessSheet(user, sheetTitle)){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'You do not have access to this sheet.' }));
    return;
  }

  payload.actorEmail = user.username;
  payload.actorName = user.displayName;

  response = await appsScript.callAppsScript('syncRun', payload);
  res.statusCode = 200;
  res.end(JSON.stringify(response || { enabled: true, synced: true }));
}

async function handleApplyManualAction(req, res, user){
  var body = readBody(req);
  var payload = body.payload || {};
  var sheetTitle = String(payload.sheetTitle || '').trim();
  var response;

  if(!user.canWrite){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'This account cannot save attendance records.' }));
    return;
  }

  if(sheetTitle && !auth.canAccessSheet(user, sheetTitle)){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'You do not have access to this sheet.' }));
    return;
  }

  payload.actorEmail = user.username;
  payload.actorName = user.displayName;
  response = await appsScript.callAppsScript('applyManualAction', payload);
  res.statusCode = 200;
  res.end(JSON.stringify(response || { enabled: true }));
}

async function handleApplyManualActionsBatch(req, res, user){
  var body = readBody(req);
  var payload = body.payload || {};
  var sheetTitle = String(payload.sheetTitle || '').trim();
  var items = Array.isArray(payload.items) ? payload.items : [];
  var response;

  if(!user.canWrite){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'This account cannot save attendance records.' }));
    return;
  }

  if(sheetTitle && !auth.canAccessSheet(user, sheetTitle)){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'You do not have access to this sheet.' }));
    return;
  }

  if(items.some(function(item){
    var itemSheetTitle = String(item && item.sheetTitle || sheetTitle).trim();
    return itemSheetTitle && !auth.canAccessSheet(user, itemSheetTitle);
  })){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'You do not have access to one or more selected items.' }));
    return;
  }

  payload.actorEmail = user.username;
  payload.actorName = user.displayName;
  response = await appsScript.callAppsScript('applyManualActionsBatch', payload);
  res.statusCode = 200;
  res.end(JSON.stringify(response || { enabled: true, items: [], manualRules: [] }));
}

async function handleSaveSnapshot(req, res, user){
  var body = readBody(req);
  var payload = body.payload || {};
  var sheetTitle = String(payload.sheetTitle || '').trim();
  var response;

  if(!user.canWrite){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'This account cannot save attendance records.' }));
    return;
  }

  if(sheetTitle && !auth.canAccessSheet(user, sheetTitle)){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'You do not have access to this sheet.' }));
    return;
  }

  payload.actorEmail = user.username;
  payload.actorName = user.displayName;
  response = await appsScript.callAppsScript('saveSnapshot', payload);
  res.statusCode = 200;
  res.end(JSON.stringify(response || { enabled: true, saved: true }));
}

async function handleStorageOverview(res, user){
  var response = await appsScript.callAppsScript('getStorageOverview', {});
  var sheets = Array.isArray(response.sheets) ? response.sheets.filter(function(item){
    return auth.canAccessSheet(user, item && item.sheetTitle);
  }) : [];

  res.statusCode = 200;
  res.end(JSON.stringify({
    enabled: true,
    sheets: sheets,
    totalOpenCount: sheets.reduce(sumField('openCount'), 0),
    totalManualRuleCount: sheets.reduce(sumField('manualRuleCount'), 0),
    lastSavedAt: getLatestValue(sheets, 'lastSavedAt')
  }));
}

async function handleStoredSheetData(req, res, user){
  var sheetTitle = String((req.query && req.query.sheetTitle) || '').trim();
  var response;

  if(!sheetTitle || !auth.canAccessSheet(user, sheetTitle)){
    res.statusCode = 403;
    res.end(JSON.stringify({ error: 'You do not have access to this sheet.' }));
    return;
  }

  response = await appsScript.callAppsScript('getStoredSheetData', { sheetTitle: sheetTitle });
  res.statusCode = 200;
  res.end(JSON.stringify(response || { enabled: true, summary: {}, snapshot: null }));
}

function filterSheetEntries(items, user){
  return (Array.isArray(items) ? items : []).filter(function(item){
    return auth.canAccessSheet(user, item && item.sheetTitle);
  });
}

function sumField(field){
  return function(total, item){
    return total + Number(item && item[field] || 0);
  };
}

function getLatestValue(items, field){
  return (items || []).reduce(function(latest, item){
    var value = String(item && item[field] || '').trim();
    if(!value) return latest;
    return !latest || value > latest ? value : latest;
  }, '');
}

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
      message: 'Stored execution backend is not configured.',
      items: [],
      manualRules: []
    };
  }

  if(action === 'syncRun'){
    return {
      enabled: false,
      message: 'Stored execution backend is not configured.',
      synced: false,
      openCount: 0,
      openedCount: 0,
      resolvedCount: 0
    };
  }

  if(action === 'saveSnapshot'){
    return {
      enabled: false,
      message: 'Stored execution backend is not configured.',
      saved: false
    };
  }

  if(action === 'getStorageOverview' || action === 'getStoredSheetData'){
    return {
      enabled: false,
      message: 'Stored execution backend is not configured.',
      sheets: [],
      summary: {},
      snapshot: null
    };
  }

  if(action === 'applyManualActionsBatch'){
    return {
      enabled: false,
      message: 'Stored execution backend is not configured.',
      items: [],
      manualRules: []
    };
  }

  return {
    enabled: false,
    message: 'Stored execution backend is not configured.'
  };
}

function setJsonHeaders(res){
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Pragma', 'no-cache');
}
