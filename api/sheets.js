var auth = require('./_lib/auth');
var appsScript = require('./_lib/apps-script');

module.exports = async function handler(req, res) {
  setJsonHeaders(res);

  var session = auth.getSession(req);
  var spreadsheetId = appsScript.getAttendanceSpreadsheetId();

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
    res.statusCode = 503;
    res.end(JSON.stringify({ error: 'Apps Script backend is not configured.' }));
    return;
  }

  if(!spreadsheetId){
    res.statusCode = 503;
    res.end(JSON.stringify({ error: 'KAKAO_CHECK_SPREADSHEET_ID is missing.' }));
    return;
  }

  try {
    if(req.method === 'GET'){
      await handleGet(req, res, session.user, spreadsheetId);
      return;
    }

    if(req.method === 'POST'){
      await handlePost(req, res, session.user, spreadsheetId);
      return;
    }

    res.statusCode = 405;
    res.end(JSON.stringify({ error: 'Method not allowed.' }));
  } catch (error) {
    res.statusCode = Number(error && error.statusCode) || 500;
    res.end(JSON.stringify({ error: error && error.message ? error.message : 'Sheet request failed.' }));
  }
};

async function handleGet(req, res, user, spreadsheetId){
  var action = String((req.query && req.query.action) || '').trim();

  if(action === 'listSheetTitles'){
    var payload = await appsScript.callAppsScript('listSheetTitles', {
      spreadsheetId: spreadsheetId
    });
    var titles = auth.filterAllowedSheets(user, payload.titles || []);
    res.statusCode = 200;
    res.end(JSON.stringify({ titles: titles }));
    return;
  }

  if(action === 'loadRosterRows'){
    var sheetTitle = String((req.query && req.query.sheetTitle) || '').trim();
    if(!auth.canAccessSheet(user, sheetTitle)){
      throw createError(403, 'You do not have access to this sheet.');
    }

    var roster = await appsScript.callAppsScript('loadRosterRows', {
      spreadsheetId: spreadsheetId,
      sheetTitle: sheetTitle,
      startRow: 2,
      targetRoleText: '수강생',
      allowRoleFallback: false,
      columns: {
        name: 'C',
        phone: 'D',
        status: 'M',
        role: 'N'
      }
    });

    res.statusCode = 200;
    res.end(JSON.stringify({
      rows: roster.rows || [],
      meta: {
        fallbackUsed: !!roster.fallbackUsed,
        roleFilterApplied: !!roster.roleFilterApplied,
        matchedRoleCount: Number(roster.matchedRoleCount || 0),
        availableRoles: Array.isArray(roster.availableRoles) ? roster.availableRoles : []
      }
    }));
    return;
  }

  throw createError(400, 'Unsupported action.');
}

async function handlePost(req, res, user, spreadsheetId){
  var body = readBody(req);
  var action = String(body.action || '').trim();
  var payload;
  var updates;

  if(action !== 'writeUpdates'){
    throw createError(400, 'Unsupported action.');
  }

  if(!user.canWrite){
    throw createError(403, 'This account cannot write attendance updates.');
  }

  updates = Array.isArray(body.updates) ? body.updates : [];
  updates.forEach(function(item){
    var sheetTitle = extractSheetTitle(item && item.range);
    if(!auth.canAccessSheet(user, sheetTitle)){
      throw createError(403, 'You do not have access to this sheet.');
    }
  });

  payload = await appsScript.callAppsScript('writeUpdates', {
    spreadsheetId: spreadsheetId,
    updates: updates
  });

  res.statusCode = 200;
  res.end(JSON.stringify(payload || { totalUpdatedCells: 0 }));
}

function extractSheetTitle(range){
  var value = String(range || '').trim();
  var bangIndex = value.lastIndexOf('!');
  var sheetTitle;

  if(bangIndex < 0) return '';
  sheetTitle = value.slice(0, bangIndex);
  if(sheetTitle[0] === '\'' && sheetTitle[sheetTitle.length - 1] === '\''){
    sheetTitle = sheetTitle.slice(1, -1).replace(/''/g, '\'');
  }
  return sheetTitle;
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

function createError(statusCode, message){
  var error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function setJsonHeaders(res){
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Pragma', 'no-cache');
}
