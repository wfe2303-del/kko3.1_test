(function(){
  var runtime = window.__KakaoCheckRuntimeConfig__ || {};
  var runtimeReady = window.__KakaoCheckRuntimeConfigReady || Promise.resolve(runtime);

  function uniqueStrings(values){
    return Array.from(new Set((values || []).map(function(item){
      return String(item || '').trim();
    }).filter(Boolean)));
  }

  function normalizeOrigin(value){
    return String(value || '').trim().replace(/\/$/, '');
  }

  function getRuntimeArray(value){
    return Array.isArray(value) ? value : [];
  }

  function getMergedConfig(){
    return {
      spreadsheetId: String(runtime.spreadsheetId || '').trim(),
      clientId: String(runtime.clientId || '').trim(),
      googleLoginHint: String(runtime.googleLoginHint || '').trim(),
      googleHostedDomainHint: String(runtime.googleHostedDomainHint || '').trim(),
      matchHistoryEnabled: !!runtime.matchHistoryEnabled,
      allowedOrigins: uniqueStrings(getRuntimeArray(runtime.allowedOrigins) || [window.location.origin]).map(normalizeOrigin),
      allowedEmailDomains: uniqueStrings(getRuntimeArray(runtime.allowedEmailDomains)).map(function(item){ return item.toLowerCase(); }),
      allowedEmails: uniqueStrings(getRuntimeArray(runtime.allowedEmails)).map(function(item){ return item.toLowerCase(); })
    };
  }

  function validateConfig(config){
    var problems = [];
    if(!config.spreadsheetId) problems.push('spreadsheetId 누락');
    if(!config.clientId) problems.push('clientId 누락');

    var spreadsheetIdOk = /^[A-Za-z0-9-_]{20,}$/.test(config.spreadsheetId || '');
    if(config.spreadsheetId && !spreadsheetIdOk) problems.push('spreadsheetId 형식 오류');

    var clientIdOk = /^[0-9]+-[a-z0-9]+\.apps\.googleusercontent\.com$/i.test(config.clientId || '');
    if(config.clientId && !clientIdOk) problems.push('clientId 형식 오류');

    var currentOrigin = normalizeOrigin(window.location.origin);
    var allowedOrigins = (config.allowedOrigins || []).map(normalizeOrigin);
    if(!allowedOrigins.length){
      problems.push('allowedOrigins 비어 있음');
    } else if(allowedOrigins.indexOf(currentOrigin) < 0){
      problems.push('현재 origin 미허용: ' + currentOrigin);
    }

    if(problems.length){
      throw new Error(
        '보안 설정이 완료되지 않았습니다. ' + problems.join(', ') + '. ' +
        'js/config.runtime.example.js 를 복사해 js/config.runtime.js 를 만든 뒤 실제 운영값만 로컬에 넣어주세요.'
      );
    }
  }

  var config = {
    startRow: 2,
    columns: Object.freeze({
      name: 'C',
      phone: 'D',
      status: 'M',
      role: 'N'
    }),
    targetRoleText: '수강생',
    statusText: '입장',
    scopes: [
      'openid',
      'email',
      'profile',
      'https://www.googleapis.com/auth/spreadsheets'
    ].join(' '),
    assertRuntimeReady: function(){
      return Promise.resolve(runtimeReady).then(function(){
        var merged = getMergedConfig();
        validateConfig(merged);
        return true;
      });
    }
  };

  Object.defineProperties(config, {
    spreadsheetId: { get: function(){ return getMergedConfig().spreadsheetId; } },
    clientId: { get: function(){ return getMergedConfig().clientId; } },
    googleLoginHint: { get: function(){ return getMergedConfig().googleLoginHint; } },
    googleHostedDomainHint: { get: function(){ return getMergedConfig().googleHostedDomainHint; } },
    matchHistoryEnabled: { get: function(){ return getMergedConfig().matchHistoryEnabled; } },
    allowedOrigins: { get: function(){ return Object.freeze(getMergedConfig().allowedOrigins.slice()); } },
    allowedEmailDomains: { get: function(){ return Object.freeze(getMergedConfig().allowedEmailDomains.slice()); } },
    allowedEmails: { get: function(){ return Object.freeze(getMergedConfig().allowedEmails.slice()); } }
  });

  Object.freeze(config);
  window.KakaoCheckConfig = config;
})();
