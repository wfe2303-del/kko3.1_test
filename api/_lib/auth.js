var crypto = require('crypto');

var COOKIE_NAME = 'kakao_check_session';
var SESSION_TTL_SECONDS = 60 * 60 * 12;

function getAuthConfig(){
  var users = loadUsers();
  var sessionSecret = String(process.env.KAKAO_CHECK_SESSION_SECRET || '').trim();

  return {
    configured: !!(sessionSecret && users.length),
    sessionSecret: sessionSecret,
    users: users
  };
}

function loadUsers(){
  var raw = String(process.env.KAKAO_CHECK_USERS_JSON || '').trim();
  if(!raw) return [];

  var parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    return [];
  }

  if(!Array.isArray(parsed)) return [];

  return parsed
    .map(normalizeUser)
    .filter(function(user){ return !!user; });
}

function normalizeUser(user){
  if(!user || typeof user !== 'object') return null;

  var username = String(user.username || '').trim();
  if(!username) return null;

  var allowedSheets = normalizeAllowedSheets(user.allowedSheets);

  return {
    username: username,
    displayName: String(user.displayName || user.name || username).trim(),
    password: String(user.password || ''),
    passwordSha256: String(user.passwordSha256 || '').trim().toLowerCase(),
    allowedSheets: allowedSheets,
    canWrite: user.canWrite !== false
  };
}

function normalizeAllowedSheets(value){
  if(value === '*' || value == null) return ['*'];
  if(Array.isArray(value)){
    var normalized = value
      .map(function(item){ return String(item || '').trim(); })
      .filter(Boolean);
    return normalized.length ? normalized : ['*'];
  }
  return [String(value).trim()].filter(Boolean);
}

function verifyPassword(user, password){
  var incoming = String(password || '');
  if(!incoming) return false;

  if(user.passwordSha256){
    return safeEqual(user.passwordSha256, sha256Hex(incoming));
  }

  return safeEqual(String(user.password || ''), incoming);
}

function sha256Hex(value){
  return crypto.createHash('sha256').update(String(value || ''), 'utf8').digest('hex');
}

function safeEqual(left, right){
  var leftBuffer = Buffer.from(String(left || ''), 'utf8');
  var rightBuffer = Buffer.from(String(right || ''), 'utf8');
  if(leftBuffer.length !== rightBuffer.length) return false;
  return crypto.timingSafeEqual(leftBuffer, rightBuffer);
}

function findUserByUsername(username){
  var normalized = String(username || '').trim();
  if(!normalized) return null;

  return getAuthConfig().users.find(function(user){
    return user.username === normalized;
  }) || null;
}

function buildUserProfile(user){
  if(!user) return null;

  return {
    username: user.username,
    displayName: user.displayName,
    email: user.username,
    name: user.displayName,
    allowedSheets: user.allowedSheets.slice(),
    canWrite: !!user.canWrite
  };
}

function signSessionPayload(payload, secret){
  return crypto.createHmac('sha256', secret).update(payload, 'utf8').digest('base64url');
}

function createSessionToken(username){
  var config = getAuthConfig();
  if(!config.sessionSecret) throw new Error('KAKAO_CHECK_SESSION_SECRET is missing.');

  var payload = JSON.stringify({
    username: String(username || '').trim(),
    exp: Date.now() + (SESSION_TTL_SECONDS * 1000)
  });

  var encoded = Buffer.from(payload, 'utf8').toString('base64url');
  var signature = signSessionPayload(encoded, config.sessionSecret);
  return encoded + '.' + signature;
}

function verifySessionToken(token){
  var config = getAuthConfig();
  if(!config.sessionSecret || !token) return null;

  var parts = String(token || '').split('.');
  if(parts.length !== 2) return null;

  var encoded = parts[0];
  var signature = parts[1];
  var expected = signSessionPayload(encoded, config.sessionSecret);
  if(!safeEqual(signature, expected)) return null;

  var payload;
  try {
    payload = JSON.parse(Buffer.from(encoded, 'base64url').toString('utf8'));
  } catch (error) {
    return null;
  }

  if(!payload || !payload.username || !payload.exp) return null;
  if(Number(payload.exp) < Date.now()) return null;

  return {
    username: String(payload.username).trim()
  };
}

function parseCookies(req){
  var raw = String((req && req.headers && req.headers.cookie) || '');
  if(!raw) return {};

  return raw.split(';').reduce(function(acc, part){
    var index = part.indexOf('=');
    if(index < 0) return acc;
    var key = part.slice(0, index).trim();
    var value = part.slice(index + 1).trim();
    acc[key] = decodeURIComponent(value);
    return acc;
  }, {});
}

function getSession(req){
  var config = getAuthConfig();
  if(!config.configured) return { configured: false, user: null };

  var cookies = parseCookies(req);
  var token = cookies[COOKIE_NAME];
  var session = verifySessionToken(token);
  if(!session) return { configured: true, user: null };

  var user = findUserByUsername(session.username);
  if(!user) return { configured: true, user: null };

  return {
    configured: true,
    user: buildUserProfile(user)
  };
}

function setSessionCookie(res, token, req){
  appendCookie(res, buildCookieValue(token, SESSION_TTL_SECONDS, req));
}

function clearSessionCookie(res, req){
  appendCookie(res, buildCookieValue('', 0, req));
}

function buildCookieValue(value, maxAge, req){
  var parts = [
    COOKIE_NAME + '=' + encodeURIComponent(String(value || '')),
    'Path=/',
    'HttpOnly',
    'SameSite=Lax',
    'Max-Age=' + Number(maxAge || 0)
  ];

  if(isSecureRequest(req)){
    parts.push('Secure');
  }

  return parts.join('; ');
}

function appendCookie(res, cookieValue){
  var existing = res.getHeader('Set-Cookie');
  if(!existing){
    res.setHeader('Set-Cookie', cookieValue);
    return;
  }
  if(Array.isArray(existing)){
    res.setHeader('Set-Cookie', existing.concat(cookieValue));
    return;
  }
  res.setHeader('Set-Cookie', [existing, cookieValue]);
}

function isSecureRequest(req){
  var forwardedProto = String((req && req.headers && req.headers['x-forwarded-proto']) || '').trim().toLowerCase();
  return forwardedProto === 'https';
}

function canAccessSheet(user, sheetTitle){
  if(!user) return false;
  if(!sheetTitle) return false;
  if(!Array.isArray(user.allowedSheets) || !user.allowedSheets.length) return false;
  if(user.allowedSheets.indexOf('*') >= 0) return true;

  var target = normalizeSheetName(sheetTitle);
  return user.allowedSheets.some(function(item){
    return normalizeSheetName(item) === target;
  });
}

function filterAllowedSheets(user, sheetTitles){
  return (sheetTitles || []).filter(function(title){
    return canAccessSheet(user, title);
  });
}

function normalizeSheetName(value){
  return String(value || '').trim().toLowerCase();
}

module.exports = {
  COOKIE_NAME: COOKIE_NAME,
  buildUserProfile: buildUserProfile,
  canAccessSheet: canAccessSheet,
  clearSessionCookie: clearSessionCookie,
  createSessionToken: createSessionToken,
  filterAllowedSheets: filterAllowedSheets,
  findUserByUsername: findUserByUsername,
  getAuthConfig: getAuthConfig,
  getSession: getSession,
  setSessionCookie: setSessionCookie,
  verifyPassword: verifyPassword
};
