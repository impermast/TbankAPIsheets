/**
 * utils_converters.gs
 * Утилиты для преобразований чисел, дат и кодов API → удобные значения для таблиц.
 */





// ============================== CACHE ===============================
function cacheGet_(key){ return CacheService.getScriptCache().get(key); }
function cachePut_(key, val, ttlSec){ CacheService.getScriptCache().put(key, val, ttlSec || 1800); } // 30 мин

// ======================= CONVERTERS / HELPERS ======================
function qToNumber(q) {
  if (!q) return null;
  var units = Number(q.units || 0);
  var nano  = Number(q.nano  || 0);
  return units + nano / 1e9;
}
function moneyToNumber(m) {
  if (!m) return null;
  if (m.units != null || m.nano != null) return qToNumber(m);
  if (m.value != null) return Number(m.value);
  return null;
}
function tsToIso(ts) {
  if (!ts) return '';
  if (typeof ts === 'string') return ts;
  if (typeof ts === 'object' && ts.seconds != null) {
    var ms = Number(ts.seconds) * 1000 + Math.round(Number(ts.nanos || 0) / 1e6);
    return new Date(ms).toISOString();
  }
  return '';
}
function dateOrEmpty_(iso) {
  if (!iso) return '';
  try { return Utilities.formatDate(new Date(iso), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
  catch (e) { return ''; }
}

/**
 * Единый конвертер: если цена похожа на % от номинала → перевести в валюту.
 * Эвристика: если известен номинал и price <= max(200, nominal*0.25) → считаем процентом.
 */
function bondPricePctToCurrency_(price, nominal) {
  if (price == null || isNaN(price)) return null;
  var p = Number(price);
  var n = Number(nominal);
  if (!isNaN(n) && n > 0) {
    var threshold = Math.max(200, n * 0.25);
    if (p <= threshold) return Math.round((p * n / 100) * 100) / 100;
  }
  return Math.round(p * 100) / 100;
}

function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function getOrCreateInputSheet_() { return getOrCreateSheet_('Input'); }
function readInputFigis_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Input');
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 1).getValues().flat();
  var seen = {};
  return vals.map(function (v) { return String(v || '').trim(); })
             .filter(function (v) { if (!v || seen[v]) return false; seen[v] = 1; return true; });
}

function htmlEscape_(s){ s=(s==null)?'':String(s); return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }



/** Округление до 2 знаков */
function round2_(x) {
  if (x == null || isNaN(x)) return null;
  return Math.round(Number(x) * 100) / 100;
}

/** Тип купона (enum/string → человекочитаемый) */
function mapCouponType_(v) {
  var key = (v == null) ? '' : String(v);
  var map = {
    '0':'Неопределенное','1':'Постоянный','2':'Плавающий','3':'Дисконт','4':'Ипотечный','5':'Фиксированный','6':'Переменный','7':'Прочее',
    'COUPON_TYPE_UNSPECIFIED':'Неопределенное','COUPON_TYPE_CONSTANT':'Постоянный','COUPON_TYPE_FLOATING':'Плавающий','COUPON_TYPE_DISCOUNT':'Дисконт',
    'COUPON_TYPE_MORTGAGE':'Ипотечный','COUPON_TYPE_FIX':'Фиксированный','COUPON_TYPE_VARIABLE':'Переменный','COUPON_TYPE_OTHER':'Прочее'
  };
  return map[key] || '';
}

/** Уровень риска (enum/string → человекочитаемый) */
function mapRiskLevel_(v) {
  var key = (v == null) ? '' : String(v);
  var map = {
    '0':'Высокий','1':'Средний','2':'Низкий',
    'RISK_LEVEL_HIGH':'Высокий','RISK_LEVEL_MODERATE':'Средний','RISK_LEVEL_LOW':'Низкий'
  };
  return map[key] || '';
}
