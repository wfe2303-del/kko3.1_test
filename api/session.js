var auth = require('./_lib/auth');

module.exports = async function handler(req, res) {
  setJsonHeaders(res);

  if(req.method === 'GET'){
    return handleGet(req, res);
  }

  if(req.method === 'POST'){
    return handlePost(req, res);
  }

  if(req.method === 'DELETE'){
    return handleDelete(req, res);
  }

  res.statusCode = 405;
  res.end(JSON.stringify({ error: 'Method not allowed' }));
};

function handleGet(req, res){
  var config = auth.getAuthConfig();
  if(!config.configured){
    res.statusCode = 503;
    res.end(JSON.stringify({ error: 'Login is not configured. Set KAKAO_CHECK_SESSION_SECRET and KAKAO_CHECK_USERS_JSON.' }));
    return;
  }

  var session = auth.getSession(req);
  res.statusCode = 200;
  res.end(JSON.stringify({
    authenticated: !!session.user,
    user: session.user || null
  }));
}

function handlePost(req, res){
  var config = auth.getAuthConfig();
  if(!config.configured){
    res.statusCode = 503;
    res.end(JSON.stringify({ error: 'Login is not configured. Set KAKAO_CHECK_SESSION_SECRET and KAKAO_CHECK_USERS_JSON.' }));
    return;
  }

  var body = readBody(req);
  var username = String(body.username || '').trim();
  var password = String(body.password || '');

  if(!username || !password){
    res.statusCode = 400;
    res.end(JSON.stringify({ error: 'Username and password are required.' }));
    return;
  }

  var user = auth.findUserByUsername(username);
  if(!user || !auth.verifyPassword(user, password)){
    res.statusCode = 401;
    res.end(JSON.stringify({ error: 'Invalid username or password.' }));
    return;
  }

  var token = auth.createSessionToken(user.username);
  auth.setSessionCookie(res, token, req);

  res.statusCode = 200;
  res.end(JSON.stringify({
    authenticated: true,
    user: auth.buildUserProfile(user)
  }));
}

function handleDelete(req, res){
  auth.clearSessionCookie(res, req);
  res.statusCode = 200;
  res.end(JSON.stringify({ ok: true }));
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

function setJsonHeaders(res){
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Pragma', 'no-cache');
}
