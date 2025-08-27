/**
 * tinkoff_api.gs
 * Ядро: токен, HTTP, API-обёртки Tinkoff Invest v2
 */

// ============================== TOKEN ===============================
function setTinkoffToken(token) {
  if (!token || typeof token !== 'string') {
    throw new Error('Пустой или неверный токен');
  }
  // можно выбрать ScriptProperties (общий) или UserProperties (персональный)
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



// =========================== API WRAPPERS ===========================
// Instruments (с кэшем)

function callInstrumentsBondByFigi_(figi) {
  var ck = 'bondBy:' + figi;
  var cached = cacheGet_(ck);
  if (cached) return JSON.parse(cached);
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.InstrumentsService/BondBy',
                       { idType:'INSTRUMENT_ID_TYPE_FIGI', id:figi }, {allow404:true});
  var res = d ? d.instrument || null : null;
  if (res) cachePut_(ck, JSON.stringify(res), 3600);
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
function callInstrumentsOptionByFigi_(figi) {
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.InstrumentsService/OptionBy',
                       { idType:'INSTRUMENT_ID_TYPE_FIGI', id:figi }, {allow404:true});
  return d ? (d.instrument || d.option || null) : null;
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
  cachePut_(ck, JSON.stringify(arr), 3*3600);
  return arr;
}

// MarketData
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

// Users & Portfolio
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
