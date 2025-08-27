/** tinkoff_api.gs — токен, HTTP, кэш, конвертеры и общие утилиты */

// ===== 1) TOKEN & HTTP =====
function setTinkoffToken(token) {
  PropertiesService.getUserProperties().setProperty('TINKOFF_TOKEN', token);
}
function getTinkoffToken() {
  var t = PropertiesService.getUserProperties().getProperty('TINKOFF_TOKEN');
  if (!t) throw new Error('Не найден TINKOFF_TOKEN в Пользовательских Свойствах.');
  return t;
}
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

// ===== 1a) Cache helpers =====
function cacheGet_(key){ return CacheService.getScriptCache().get(key); }
function cachePut_(key, val, ttlSec){ CacheService.getScriptCache().put(key, val, ttlSec || 1800); } // 30 мин

// ===== 2) Converters / Mappers / Sheet utils =====
function qToNumber(q) { if (!q) return null; return Number(q.units || 0) + Number(q.nano || 0) / 1e9; }
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
function round2_(x) { if (x == null || isNaN(x)) return null; return Math.round(Number(x) * 100) / 100; }

function mapCouponType_(v) {
  var key = (v == null) ? '' : String(v);
  var map = {
    '0':'Неопределенное','1':'Постоянный','2':'Плавающий','3':'Дисконт','4':'Ипотечный','5':'Фиксированный','6':'Переменный','7':'Прочее',
    'COUPON_TYPE_UNSPECIFIED':'Неопределенное','COUPON_TYPE_CONSTANT':'Постоянный','COUPON_TYPE_FLOATING':'Плавающий','COUPON_TYPE_DISCOUNT':'Дисконт',
    'COUPON_TYPE_MORTGAGE':'Ипотечный','COUPON_TYPE_FIX':'Фиксированный','COUPON_TYPE_VARIABLE':'Переменный','COUPON_TYPE_OTHER':'Прочее'
  };
  return map[key] || '';
}
function mapRiskLevel_(v) {
  var key = (v == null) ? '' : String(v);
  var map = {
    '0':'Высокий','1':'Средний','2':'Низкий',
    'RISK_LEVEL_HIGH':'Высокий','RISK_LEVEL_MODERATE':'Средний','RISK_LEVEL_LOW':'Низкий'
  };
  return map[key] || '';
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

/** Единая конверсия «% от номинала → валюта» */
function bondPricePctToCurrency_(price, nominal) {
  if (price == null || isNaN(price)) return null;
  var p = Number(price);
  var n = Number(nominal);
  if (!isNaN(n) && n > 0) {
    var threshold = Math.max(200, n * 0.25); // 25% от номинала, но не ниже 200
    if (p <= threshold) return Math.round((p * n / 100) * 100) / 100;
  }
  return Math.round(p * 100) / 100;
}

// ===== 3) API WRAPPERS =====
function callInstrumentsBondByFigi_(figi) {
  var ck = 'bondBy:' + figi;
  var cached = cacheGet_(ck);
  if (cached) return JSON.parse(cached);
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.InstrumentsService/BondBy',
                       { idType:'INSTRUMENT_ID_TYPE_FIGI', id:figi }, {allow404:true});
  var res = d ? d.instrument || null : null;
  if (res) cachePut_(ck, JSON.stringify(res), 3600); // 1 час
  return res;
}
function callInstrumentsShareByFigi_(figi) {
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.InstrumentsService/ShareBy',
                       { idType:'INSTRUMENT_ID_TYPE_FIGI', id:figi }, {allow404:true});
  return d ? d.instrument || null : null;
}
function callInstrumentsEtfByFigi_(figi) {
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.InstrumentsService/EtfBy',
                       { idType:'INSTRUMENT_ID_TYPE_FIGI', id:figi }, {allow404:true});
  return d ? d.instrument || null : null;
}
function callInstrumentsGetBondCoupons_(figi, fromIso, toIso) {
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.InstrumentsService/GetBondCoupons',
                       { figi:figi, from:fromIso, to:toIso }, {allow404:true});
  if (!d) return [];
  return d.coupons || d.events || d.bondCoupons || [];
}
function callInstrumentsGetBondCouponsCached_(figi, fromIso, toIso){
  var ck = 'coupons:' + figi + ':' + (fromIso||'').slice(0,10) + ':' + (toIso||'').slice(0,10);
  var c = cacheGet_(ck);
  if (c) return JSON.parse(c);
  var arr = callInstrumentsGetBondCoupons_(figi, fromIso, toIso) || [];
  cachePut_(ck, JSON.stringify(arr), 3*3600); // 3 часа
  return arr;
}

function callMarketLastPrices_(figis) {
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.MarketDataService/GetLastPrices',
                       { instrumentId: figis });
  var arr = (d && d.lastPrices) || [];
  return arr.map(function (x) {
    return {
      figi: x.figi || x.instrumentFigi || '',
      lastPrice: qToNumber(x.price || x.lastPrice),
      time: tsToIso(x.time || x.lastPriceTime)
    };
  }).filter(function (x) { return x.figi; });
}
function callMarketAccruedInterestsToday_(figi) {
  var now = new Date();
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.InstrumentsService/GetAccruedInterests',
                       { figi:figi, from: now.toISOString(), to: new Date(now.getTime()+24*3600*1000).toISOString() },
                       {allow404:true});
  if (!d) return null;
  var rec = (d.accruedInterests || [])[0];
  if (!rec) return null;
  return moneyToNumber(rec.value) ?? qToNumber(rec.accruedInterest) ?? moneyToNumber(rec.accruedValue) ?? null;
}

function callUsersGetAccounts_() {
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.UsersService/GetAccounts', {});
  var a = d.accounts || [];
  return a.map(function (x) { return { accountId: x.id || x.accountId, name: x.name || '' }; })
          .filter(function (x) { return x.accountId; });
}
function callPortfolioGetPortfolio_(accountId) {
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.OperationsService/GetPortfolio',
                       { accountId: accountId }, {allow404:true});
  if (!d) return [];
  var p = d.positions || [];
  return p.map(function (x) {
    return {
      figi: x.figi || x.instrumentFigi || '',
      quantity: qToNumber(x.quantity || x.balance),
      avg: moneyToNumber(x.averagePositionPrice || x.averagePositionPriceFifo || x.averagePositionPriceNoNkd),
      avg_fifo: moneyToNumber(x.averagePositionPriceFifo || x.averagePositionPrice || x.averagePositionPriceNoNkd)
    };
  });
}
function callPortfolioGetPositions_(accountId) {
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.OperationsService/GetPositions',
                       { accountId: accountId }, {allow404:true});
  if (!d) return [];
  var out = [];
  if (Array.isArray(d.securities)) {
    d.securities.forEach(function (s) {
      out.push({
        figi: s.figi || s.instrumentFigi || '',
        quantity: qToNumber(s.quantity) ?? (s.balance != null ? Number(s.balance) : null),
        avg: moneyToNumber(s.averagePositionPrice || s.averagePositionPriceFifo || s.averagePositionPriceNoNkd),
        avg_fifo: moneyToNumber(s.averagePositionPriceFifo || s.averagePositionPrice || s.averagePositionPriceNoNkd)
      });
    });
  }
  if (Array.isArray(d.positions)) {
    d.positions.forEach(function (p) {
      out.push({
        figi: p.figi || p.instrumentFigi || '',
        quantity: qToNumber(p.quantity) ?? (p.balance != null ? Number(p.balance) : null),
        avg: moneyToNumber(p.averagePositionPrice || p.averagePositionPriceFifo || p.averagePositionPriceNoNkd),
        avg_fifo: moneyToNumber(p.averagePositionPriceFifo || p.averagePositionPrice || p.averagePositionPriceNoNkd)
      });
    });
  }
  return out;
}
