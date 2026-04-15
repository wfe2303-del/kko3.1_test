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
    var allowedOrigins = uniqueStrings(getRuntimeArray(runtime.allowedOrigins));
    if(!allowedOrigins.length && window.location && window.location.origin){
      allowedOrigins = [window.location.origin];
    }

    return {
      matchHistoryEnabled: !!runtime.matchHistoryEnabled,
      allowedOrigins: allowedOrigins.map(normalizeOrigin)
    };
  }

  function validateConfig(config){
    var problems = [];
    var currentOrigin = normalizeOrigin(window.location.origin);
    var allowedOrigins = (config.allowedOrigins || []).map(normalizeOrigin);

    if(!allowedOrigins.length){
      problems.push('allowedOrigins missing');
    } else if(!isAllowedOrigin(currentOrigin, allowedOrigins)){
      problems.push('current origin is not allowed: ' + currentOrigin);
    }

    if(problems.length){
      throw new Error('Runtime config is invalid. ' + problems.join(', '));
    }
  }

  function isAllowedOrigin(origin, patterns){
    var normalizedOrigin = normalizeOrigin(origin);
    return (patterns || []).some(function(pattern){
      return matchesOriginPattern(normalizeOrigin(pattern), normalizedOrigin);
    });
  }

  function matchesOriginPattern(pattern, origin){
    if(!pattern) return false;
    if(pattern.indexOf('*') < 0) return pattern === origin;

    var escaped = escapeRegex(pattern).replace(/\\\*/g, '[^/]+');
    return new RegExp('^' + escaped + '$').test(origin);
  }

  function escapeRegex(value){
    return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
    assertRuntimeReady: function(){
      return Promise.resolve(runtimeReady).then(function(){
        validateConfig(getMergedConfig());
        return true;
      });
    },
    isAllowedOrigin: function(origin){
      return isAllowedOrigin(origin, getMergedConfig().allowedOrigins);
    }
  };

  Object.defineProperties(config, {
    matchHistoryEnabled: { get: function(){ return getMergedConfig().matchHistoryEnabled; } },
    allowedOrigins: { get: function(){ return Object.freeze(getMergedConfig().allowedOrigins.slice()); } }
  });

  Object.freeze(config);
  window.KakaoCheckConfig = config;
})();
