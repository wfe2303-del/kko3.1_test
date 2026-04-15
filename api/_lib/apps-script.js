function isConfigured(){
  return !!(
    String(process.env.KAKAO_CHECK_APPS_SCRIPT_URL || '').trim() &&
    String(process.env.KAKAO_CHECK_APPS_SCRIPT_TOKEN || '').trim()
  );
}

function getAttendanceSpreadsheetId(){
  return String(process.env.KAKAO_CHECK_SPREADSHEET_ID || '').trim();
}

async function callAppsScript(action, payload){
  var appsScriptUrl = String(process.env.KAKAO_CHECK_APPS_SCRIPT_URL || '').trim();
  var appsScriptToken = String(process.env.KAKAO_CHECK_APPS_SCRIPT_TOKEN || '').trim();

  if(!appsScriptUrl || !appsScriptToken){
    throw new Error('Apps Script backend is not configured.');
  }

  var response = await fetch(appsScriptUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    body: JSON.stringify({
      token: appsScriptToken,
      action: action,
      payload: payload || {}
    })
  });

  var text = await response.text();
  var data = safeParseJson(text);

  if(!response.ok){
    throw new Error(data && data.error ? data.error : (text || 'Apps Script backend request failed.'));
  }

  if(data && data.error){
    throw new Error(data.error);
  }

  return data || {};
}

function safeParseJson(text){
  try {
    return JSON.parse(text);
  } catch (error) {
    return null;
  }
}

module.exports = {
  callAppsScript: callAppsScript,
  getAttendanceSpreadsheetId: getAttendanceSpreadsheetId,
  isConfigured: isConfigured
};
