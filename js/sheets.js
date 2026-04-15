(function(){
  var config = window.KakaoCheckConfig;
  var auth = window.KakaoCheckAuth;
  var utils = window.KakaoCheckUtils;

  async function apiFetch(url, init){
    var token = auth.requireToken();
    var options = init || {};
    var headers = Object.assign({}, options.headers || {}, {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json'
    });
    var response = await fetch(url, Object.assign({}, options, {
      headers: headers,
      cache: 'no-store',
      credentials: 'omit',
      referrerPolicy: 'no-referrer'
    }));
    if(!response.ok) throw new Error(await response.text());
    if(response.status === 204) return null;
    return response.json();
  }

  async function listSheetTitles(){
    var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(config.spreadsheetId) + '?fields=sheets(properties(title))';
    var data = await apiFetch(url);
    return (data.sheets || []).map(function(sheet){ return sheet.properties.title; });
  }

  async function loadRosterRows(sheetTitle){
    var ranges = [
      sheetTitle + '!' + config.columns.name + config.startRow + ':' + config.columns.phone,
      sheetTitle + '!' + config.columns.status + config.startRow + ':' + config.columns.status,
      sheetTitle + '!' + config.columns.role + config.startRow + ':' + config.columns.role
    ];
    var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(config.spreadsheetId) + '/values:batchGet?ranges=' +
      ranges.map(function(range){ return encodeURIComponent(range); }).join('&ranges=');
    var data = await apiFetch(url);
    var valueRanges = data.valueRanges || [];
    var mainRows = valueRanges[0] && valueRanges[0].values ? valueRanges[0].values : [];
    var statusRows = valueRanges[1] && valueRanges[1].values ? valueRanges[1].values : [];
    var roleRows = valueRanges[2] && valueRanges[2].values ? valueRanges[2].values : [];
    var maxLength = Math.max(mainRows.length, statusRows.length, roleRows.length);
    var roster = [];

    for(var i = 0; i < maxLength; i += 1){
      var main = mainRows[i] || [];
      var status = statusRows[i] || [];
      var role = roleRows[i] || [];
      var name = String(main[0] || '').trim();
      var phone = String(main[1] || '').trim();
      var statusText = String(status[0] || '').trim();
      var roleText = String(role[0] || '').trim();
      if(roleText !== config.targetRoleText) continue;
      if(!name && !phone && !statusText) continue;
      roster.push({
        rowIndex: i,
        rowNumber: config.startRow + i,
        sheetTitle: sheetTitle,
        name: name,
        phone: phone,
        status: statusText,
        role: roleText,
        nameNormalized: utils.normalizeName(name)
      });
    }
    return roster;
  }

  async function writeUpdates(updates){
    if(!updates.length) return { totalUpdatedCells: 0 };
    var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(config.spreadsheetId) + '/values:batchUpdate';
    var valueRanges = updates.map(function(item){
      return {
        range: item.range,
        values: item.values
      };
    });
    return apiFetch(url, {
      method: 'POST',
      body: JSON.stringify({
        valueInputOption: 'USER_ENTERED',
        data: valueRanges
      })
    });
  }

  window.KakaoCheckSheets = {
    listSheetTitles: listSheetTitles,
    loadRosterRows: loadRosterRows,
    writeUpdates: writeUpdates
  };
})();
