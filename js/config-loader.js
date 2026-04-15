(function(){
  var runtime = window.__KakaoCheckRuntimeConfig__ = window.__KakaoCheckRuntimeConfig__ || {};

  function mergeRuntime(payload){
    if(!payload || typeof payload !== 'object') return;
    Object.keys(payload).forEach(function(key){
      if(payload[key] === undefined || payload[key] === null) return;
      runtime[key] = payload[key];
    });
  }

  function loadOptionalScript(src){
    return new Promise(function(resolve){
      var script = document.createElement('script');
      script.src = src;
      script.async = false;
      script.onload = function(){ resolve(true); };
      script.onerror = function(){ resolve(false); };
      document.head.appendChild(script);
    });
  }

  function loadRemoteRuntimeConfig(){
    return fetch('/api/runtime-config', {
      method: 'GET',
      cache: 'no-store',
      credentials: 'omit',
      referrerPolicy: 'no-referrer',
      headers: { 'Accept': 'application/json' }
    }).then(function(res){
      if(!res.ok) return null;
      return res.json();
    }).then(function(payload){
      mergeRuntime(payload);
      return payload;
    }).catch(function(){
      return null;
    });
  }

  window.__KakaoCheckRuntimeConfigReady = loadOptionalScript('./js/config.runtime.local.js')
    .then(loadRemoteRuntimeConfig)
    .then(function(){ return runtime; });
})();
