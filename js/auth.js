(function(){
  var config = window.KakaoCheckConfig;
  var accessToken = null;
  var tokenClient = null;
  var currentUser = null;
  var listeners = [];
  var initPromise = null;
  var initialized = false;
  var tokenExpiryTimer = null;
  var tokenExpiresAt = null;

  function notify(){
    listeners.forEach(function(listener){
      try {
        listener({ accessToken: accessToken, user: currentUser, tokenExpiresAt: tokenExpiresAt });
      } catch (error) {
        console.error(error);
      }
    });
  }

  function onChange(listener){
    listeners.push(listener);
  }

  function waitForGoogleIdentity(timeoutMs, intervalMs){
    return new Promise(function(resolve, reject){
      var startedAt = Date.now();
      function check(){
        if(window.google && google.accounts && google.accounts.oauth2){
          resolve();
          return;
        }
        if(Date.now() - startedAt >= timeoutMs){
          reject(new Error('Google 로그인 스크립트를 불러오지 못했습니다. 새로고침 후 다시 시도하세요.'));
          return;
        }
        setTimeout(check, intervalMs);
      }
      check();
    });
  }

  function clearTokenExpiryTimer(){
    if(tokenExpiryTimer){
      clearTimeout(tokenExpiryTimer);
      tokenExpiryTimer = null;
    }
  }

  function scheduleTokenExpiry(expiresInSeconds){
    clearTokenExpiryTimer();
    tokenExpiresAt = expiresInSeconds ? Date.now() + (Number(expiresInSeconds) * 1000) : null;
    if(!tokenExpiresAt) return;
    var delay = Math.max((tokenExpiresAt - Date.now()) - 5000, 0);
    tokenExpiryTimer = setTimeout(function(){
      revoke(false);
      alert('로그인 세션이 만료되어 자동 로그아웃되었습니다. 다시 로그인해주세요.');
    }, delay);
  }

  function assertOriginAllowed(){
    var currentOrigin = String(window.location.origin || '').replace(/\/$/, '');
    if(config.allowedOrigins.indexOf(currentOrigin) >= 0) return;
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
        return waitForGoogleIdentity(15000, 100);
      })
      .then(function(){
      tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: config.clientId,
        scope: config.scopes,
        prompt: '',
        login_hint: config.googleLoginHint || undefined,
        hd: config.googleHostedDomainHint || undefined,
        callback: handleTokenResponse,
        error_callback: function(err){
          console.error(err);
          alert('OAuth 오류: ' + (err && (err.error || err.type || JSON.stringify(err))));
        }
      });
      initialized = true;
      return true;
    });
    return initPromise;
  }

  async function handleTokenResponse(response){
    try {
      accessToken = response && response.access_token ? response.access_token : null;
      scheduleTokenExpiry(response && response.expires_in);
      currentUser = accessToken ? await fetchUserInfo() : null;
      validateUser(currentUser);
      notify();
    } catch (error) {
      accessToken = null;
      currentUser = null;
      tokenExpiresAt = null;
      clearTokenExpiryTimer();
      notify();
      alert(error.message);
    }
  }

  async function fetchUserInfo(){
    var res = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      method: 'GET',
      cache: 'no-store',
      credentials: 'omit',
      referrerPolicy: 'no-referrer',
      headers: { Authorization: 'Bearer ' + accessToken }
    });
    if(!res.ok) throw new Error('로그인 사용자 정보를 불러오지 못했습니다.');
    return res.json();
  }

  function validateUser(user){
    var email = String(user && user.email || '').toLowerCase();
    var domain = email.split('@')[1] || '';
    if(config.allowedEmails.length && config.allowedEmails.indexOf(email) >= 0) return;
    if(config.allowedEmailDomains.indexOf(domain) >= 0) return;
    revoke();
    throw new Error('허용되지 않은 계정입니다. 허용된 회사 계정으로 로그인하세요.');
  }

  async function login(){
    await init();
    if(!tokenClient) throw new Error('로그인 준비가 아직 끝나지 않았습니다.');
    tokenClient.requestAccessToken({ prompt: 'consent', login_hint: config.googleLoginHint || undefined });
  }

  function revoke(notifyListeners){
    if(notifyListeners === undefined) notifyListeners = true;
    if(accessToken && window.google && google.accounts && google.accounts.oauth2){
      google.accounts.oauth2.revoke(accessToken, function(){});
    }
    accessToken = null;
    currentUser = null;
    tokenExpiresAt = null;
    clearTokenExpiryTimer();
    if(notifyListeners) notify();
  }

  function softLogoutOnPageHide(){
    accessToken = null;
    currentUser = null;
    tokenExpiresAt = null;
    clearTokenExpiryTimer();
  }

  function requireToken(){
    if(!accessToken) throw new Error('먼저 Google 로그인하세요.');
    return accessToken;
  }

  function getAccessToken(){ return accessToken; }
  function getCurrentUser(){ return currentUser; }
  function isInitialized(){ return initialized; }

  window.addEventListener('pagehide', softLogoutOnPageHide);

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
