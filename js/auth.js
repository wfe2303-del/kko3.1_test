(function(){
  var config = window.KakaoCheckConfig;
  var currentUser = null;
  var listeners = [];
  var initPromise = null;
  var initialized = false;

  function notify(){
    var payload = {
      accessToken: currentUser ? 'session' : null,
      user: currentUser
    };

    listeners.forEach(function(listener){
      try {
        listener(payload);
      } catch (error) {
        console.error(error);
      }
    });
  }

  function onChange(listener){
    listeners.push(listener);
  }

  function assertOriginAllowed(){
    var currentOrigin = String(window.location.origin || '').replace(/\/$/, '');
    if(config.isAllowedOrigin(currentOrigin)) return;
    throw new Error('허용되지 않은 배포 주소입니다: ' + currentOrigin);
  }

  function init(){
    if(initPromise) return initPromise;

    initPromise = Promise.resolve()
      .then(function(){
        return config.assertRuntimeReady();
      })
      .then(function(){
        assertOriginAllowed();
        return fetch('/api/session', {
          method: 'GET',
          cache: 'no-store',
          credentials: 'same-origin',
          referrerPolicy: 'no-referrer',
          headers: {
            'Accept': 'application/json'
          }
        });
      })
      .then(async function(response){
        var text = await response.text();
        var data = safeParseJson(text);

        if(!response.ok){
          throw new Error(data && data.error ? data.error : (text || '로그인 상태를 확인하지 못했습니다.'));
        }

        currentUser = data && data.authenticated ? normalizeUser(data.user) : null;
        initialized = true;
        notify();
        return true;
      });

    return initPromise;
  }

  async function login(username, password){
    var response = await fetch('/api/session', {
      method: 'POST',
      cache: 'no-store',
      credentials: 'same-origin',
      referrerPolicy: 'no-referrer',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: String(username || '').trim(),
        password: String(password || '')
      })
    });

    var text = await response.text();
    var data = safeParseJson(text);

    if(!response.ok){
      throw new Error(data && data.error ? data.error : (text || '로그인에 실패했습니다.'));
    }

    currentUser = normalizeUser(data.user);
    notify();
    return currentUser;
  }

  async function revoke(){
    await fetch('/api/session', {
      method: 'DELETE',
      cache: 'no-store',
      credentials: 'same-origin',
      referrerPolicy: 'no-referrer',
      headers: {
        'Accept': 'application/json'
      }
    }).catch(function(){});

    currentUser = null;
    notify();
  }

  function requireToken(){
    if(!currentUser) throw new Error('먼저 로그인해주세요.');
    return 'session';
  }

  function getAccessToken(){
    return currentUser ? 'session' : null;
  }

  function getCurrentUser(){
    return currentUser;
  }

  function isInitialized(){
    return initialized;
  }

  function normalizeUser(user){
    if(!user || typeof user !== 'object') return null;
    return {
      username: String(user.username || user.email || '').trim(),
      displayName: String(user.displayName || user.name || user.email || '').trim(),
      email: String(user.email || user.username || '').trim(),
      name: String(user.name || user.displayName || user.username || '').trim(),
      allowedSheets: Array.isArray(user.allowedSheets) ? user.allowedSheets.slice() : [],
      canWrite: !!user.canWrite
    };
  }

  function safeParseJson(text){
    try {
      return JSON.parse(text);
    } catch (error) {
      return null;
    }
  }

  window.KakaoCheckAuth = {
    init: init,
    login: login,
    revoke: revoke,
    onChange: onChange,
    requireToken: requireToken,
    getAccessToken: getAccessToken,
    getCurrentUser: getCurrentUser,
    isInitialized: isInitialized
  };
})();
