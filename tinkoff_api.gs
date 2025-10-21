/**
 * tinkoff_api.gs
 * Обёртки Tinkoff Invest v2 по сервисам API (без UI).
 * Зависимости: cacheGet_, cachePut_, qToNumber, moneyToNumber, tsToIso (utils).
 */

// =========================== Instruments ============================
function callInstrumentsBondByFigi_(figi) {
  var ck = 'bondBy:' + figi;
  var c = cacheGet_(ck); if (c) return JSON.parse(c);
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
  var ck = 'etfBy:' + figi;
  var c = cacheGet_(ck); if (c) return JSON.parse(c);
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.InstrumentsService/EtfBy',
                       { idType:'INSTRUMENT_ID_TYPE_FIGI', id:figi }, {allow404:true});
  var inst = d ? (d.instrument || d.etf || null) : null;
  if (!inst) return null;

  // Нормализация ETF
  if (inst.realExchange && !inst.exchange) inst.exchange = inst.realExchange;
  if (inst.lotSize != null && inst.lot == null) inst.lot = inst.lotSize;
  if (inst.company == null && inst.provider != null) inst.company = inst.provider;
  try {
    if (inst.expenseRatio == null && inst.totalExpense != null) {
      var er = moneyToNumber(inst.totalExpense);
      if (er != null) inst.expenseRatio = er;
    }
  } catch(e){}
  if (inst.blockedTcaFlag == null) {
    if (inst.isBlocked != null) inst.blockedTcaFlag = !!inst.isBlocked;
  } else {
    inst.blockedTcaFlag = !!inst.blockedTcaFlag;
  }
  if (!inst.currency) inst.currency = (inst.buyCurrency || inst.sellCurrency || '') || inst.currency;

  cachePut_(ck, JSON.stringify(inst), 6*3600);
  return inst;
}


/** ========== OPTIONS (совместимо с GetOptionBy/GetOptions и старыми OptionBy/OptionsBy) ========== */

/** Опцион по FIGI (если передали UID по ошибке — попробуем как UID). */
function callInstrumentsOptionByFigi_(id){
  var ck = 'optByFigi:v2:' + id;
  var c  = cacheGet_(ck); if (c) return JSON.parse(c);

  var inst = null;

  // Новый метод
  try{
    var d = tinkoffFetch(
      'tinkoff.public.invest.api.contract.v1.InstrumentsService/GetOptionBy',
      { idType:'INSTRUMENT_ID_TYPE_FIGI', id:String(id) },
      { allow404:true }
    );
    inst = d ? (d.instrument || d.option || null) : null;
  }catch(_){}

  // Фоллбек: вдруг это UID
  if (!inst){
    try{
      var d2 = tinkoffFetch(
        'tinkoff.public.invest.api.contract.v1.InstrumentsService/GetOptionBy',
        { idType:'INSTRUMENT_ID_TYPE_UID', id:String(id) },
        { allow404:true }
      );
      inst = d2 ? (d2.instrument || d2.option || null) : null;
    }catch(_){}
  }

  // Легаси фоллбек (старый метод)
  if (!inst){
    try{
      var d3 = tinkoffFetch(
        'tinkoff.public.invest.api.contract.v1.InstrumentsService/OptionBy',
        { idType:'INSTRUMENT_ID_TYPE_FIGI', id:String(id) },
        { allow404:true }
      );
      inst = d3 ? (d3.instrument || d3.option || null) : null;
    }catch(_){}
  }

  if (inst) cachePut_(ck, JSON.stringify(inst), 3*3600);
  return inst;
}

/** Опцион по UID. */
function callInstrumentsOptionByUid_(uid){
  var ck = 'optByUid:v2:' + uid;
  var c  = cacheGet_(ck); if (c) return JSON.parse(c);

  var inst = null;

  // Новый метод
  try{
    var d = tinkoffFetch(
      'tinkoff.public.invest.api.contract.v1.InstrumentsService/GetOptionBy',
      { idType:'INSTRUMENT_ID_TYPE_UID', id:String(uid) },
      { allow404:true }
    );
    inst = d ? (d.instrument || d.option || null) : null;
  }catch(_){}

  // Легаси фоллбек
  if (!inst){
    try{
      var d2 = tinkoffFetch(
        'tinkoff.public.invest.api.contract.v1.InstrumentsService/OptionBy',
        { idType:'INSTRUMENT_ID_TYPE_UID', id:String(uid) },
        { allow404:true }
      );
      inst = d2 ? (d2.instrument || d2.option || null) : null;
    }catch(_){}
  }

  if (inst) cachePut_(ck, JSON.stringify(inst), 3*3600);
  return inst;
}

/**
 * Все опционы по базовому активу.
 * ref: строка (basicAssetUid) ИЛИ объект { basicAssetUid, basicAssetPositionUid }.
 */
function callInstrumentsOptionsBy_(ref){
  var key = (typeof ref === 'string') ? ref : JSON.stringify(ref||{});
  var ck  = 'optionsBy:v2:' + key;
  var c   = cacheGet_(ck); if (c) return JSON.parse(c);

  // Собираем запрос для нового метода
  var req = {};
  if (typeof ref === 'string') {
    req.basicAssetUid = ref;
  } else if (ref && typeof ref === 'object') {
    if (ref.basicAssetUid) req.basicAssetUid = String(ref.basicAssetUid);
    if (ref.basicAssetPositionUid) req.basicAssetPositionUid = String(ref.basicAssetPositionUid);
  }

  var list = [];

  // Новый метод
  try{
    var d = tinkoffFetch(
      'tinkoff.public.invest.api.contract.v1.InstrumentsService/GetOptions',
      req,
      { allow404:true }
    );
    list = d ? (d.options || d.instruments || []) : [];
  }catch(_){}

  // Легаси фоллбек (в старом был только basicAssetUid)
  if (!list || !list.length){
    try{
      var d2 = tinkoffFetch(
        'tinkoff.public.invest.api.contract.v1.InstrumentsService/OptionsBy',
        { basicAssetUid: (req.basicAssetUid || key) },
        { allow404:true }
      );
      list = d2 ? (d2.instruments || d2.options || []) : [];
    }catch(_){}
  }

  cachePut_(ck, JSON.stringify(list||[]), 3*3600);
  return list || [];
}



/** Asset по UID (ETF-поля: focusType, totalExpense, tracking error и т. п.) */
function callInstrumentsGetAssetByUid_(assetUid){
  var ck = 'assetByUid:v2:' + assetUid;
  var c = cacheGet_(ck); if (c) return JSON.parse(c);
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.InstrumentsService/GetAssetBy',
                       { idType: 'ASSET_ID_TYPE_UID', id: assetUid }, { allow404:true });
  var asset = d ? (d.asset || null) : null;
  if (asset) cachePut_(ck, JSON.stringify(asset), 6*3600);
  return asset;
}
/** Купоны облигации (сырой/кэш) */
function callInstrumentsGetBondCoupons_(figi, fromIso, toIso) {
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.InstrumentsService/GetBondCoupons',
                       { figi:figi, from:fromIso, to:toIso }, {allow404:true});
  if (!d) return [];
  return d.coupons || d.events || d.bondCoupons || [];
}
function callInstrumentsGetBondCouponsCached_(figi, fromIso, toIso){
  var ck = 'coupons:' + figi + ':' + (fromIso||'').slice(0,10) + ':' + (toIso||'').slice(0,10);
  var c = cacheGet_(ck); if (c) return JSON.parse(c);
  var arr = callInstrumentsGetBondCoupons_(figi, fromIso, toIso) || [];
  cachePut_(ck, JSON.stringify(arr), 3*3600);
  return arr;
}

// ============================ MarketData ============================
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
/** Trading status */
function callMarketGetTradingStatus_(figi){
  var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.MarketDataService/GetTradingStatus',
                       { instrumentId: figi }, { allow404:true });
  if (!d) return null;
  return {
    tradingStatus: d.tradingStatus || d.status || d.securityTradingStatus || d.instrumentTradingStatus || '',
    time: tsToIso(d.time || d.lastTime || d.lastPriceTime || '')
  };
}

// ======================== Users / Portfolio ========================
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
  var d = tinkoffFetch(
    'tinkoff.public.invest.api.contract.v1.OperationsService/GetPositions',
    { accountId: accountId }, { allow404: true }
  );
  if (!d) return [];

  var out = [];
  function pushPos(p) {
    if (!p) return;
    out.push({
      figi: p.figi || p.instrumentFigi || '',
      instrumentUid: p.instrumentUid || p.uid || '',
      quantity: qToNumber(p.quantity) ?? (p.balance != null ? Number(p.balance) : null),
      avg: moneyToNumber(p.averagePositionPrice || p.averagePositionPriceFifo || p.averagePositionPriceNoNkd),
      avg_fifo: moneyToNumber(p.averagePositionPriceFifo || p.averagePositionPrice || p.averagePositionPriceNoNkd)
    });
  }

  (d.securities   || []).forEach(pushPos);
  (d.positions    || []).forEach(pushPos);
  (d.options      || []).forEach(pushPos);    // <-- опционы
  (d.futures      || []).forEach(pushPos);
  (d.derivatives  || []).forEach(pushPos);    // на всякий случай

  return out;
}
