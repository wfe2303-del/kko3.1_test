module.exports = function handler(req, res) {
  function parseList(value) {
    return String(value || '')
      .split(',')
      .map(function(item){ return item.trim(); })
      .filter(Boolean);
  }

  var payload = {
    matchHistoryEnabled: Boolean(
      String(process.env.KAKAO_CHECK_APPS_SCRIPT_URL || '').trim() &&
      String(process.env.KAKAO_CHECK_APPS_SCRIPT_TOKEN || '').trim()
    ),
    allowedOrigins: parseList(process.env.KAKAO_CHECK_ALLOWED_ORIGINS)
  };

  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Pragma', 'no-cache');
  res.status(200).json(payload);
};
