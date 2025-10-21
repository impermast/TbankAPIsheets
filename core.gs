/**
 * core.gs
 * Токен и HTTP-вызовы (без привязки к UI).
 */

// ============================== TOKEN ===============================
function setTinkoffToken(token) {
  if (!token || typeof token !== 'string') throw new Error('Пустой или неверный токен');
  PropertiesService.getUserProperties().setProperty('TINKOFF_TOKEN', token.trim());
}
function getTinkoffToken() {
  var t = PropertiesService.getUserProperties().getProperty('TINKOFF_TOKEN');
  if (!t) throw new Error('Не найден TINKOFF_TOKEN (проверьте: Настройки → Задать токен)');
  return t;
}

// ============================== HTTP ================================
function tinkoffFetch(methodPath, body, opt) {
  var url = 'https://invest-public-api.tinkoff.ru/rest/' + methodPath;
  var options = {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + getTinkoffToken() },
    payload: JSON.stringify(body || {})
  };
  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  var text = resp.getContentText();

  if (code === 429 || code >= 500) {
    Utilities.sleep((opt && opt.retrySleepMs) || 400);
    var resp2 = UrlFetchApp.fetch(url, options);
    code = resp2.getResponseCode();
    text = resp2.getContentText();
  }
  if (code === 404 && opt && opt.allow404) return null;
  if (code < 200 || code >= 300) throw new Error('Tinkoff API error ' + code + ': ' + text);
  return JSON.parse(text);
}
function tinkoffFetchRaw_(methodPath, body) {
  var url = 'https://invest-public-api.tinkoff.ru/rest/' + methodPath;
  var options = {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + getTinkoffToken() },
    payload: JSON.stringify(body || {})
  };
  var resp = UrlFetchApp.fetch(url, options);
  return { code: resp.getResponseCode(), text: resp.getContentText() };
}
