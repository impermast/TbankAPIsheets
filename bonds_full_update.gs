/**
 * bonds_full_update.gs
 * Полное обновление листа «Bonds» с кэшированием купонов/карточек,
 * уникальными заголовками и фиксами формул.
 */

// ===== УНИКАЛЬНЫЕ заголовки: два риска — ручной и «подтягиваемый» из API =====
const BONDS_HEADERS = [
  // Блок 1 — приоритетные
  '📌 Риск (ручн.)',          // числовой, ручной скоринг 0..N
  'Комментарий ИИ',
  'Комментарий пользователя',
  'Тип инструмента',
  'FIGI',
  'Название',
  'Кол-во',
  'Средняя цена',
  'Текущая цена',
  'Купон, %',
  'Размер купона',
  'купон/год',
  'Тип купона (desc)',
  'Сектор',
  'Риск (уровень TCS)',       // текстовый, из API (mapRiskLevel_)

  // Блок 2 — расчётные
  'Инвестировано',
  'Рыночная стоимость',
  'P/L (руб)',
  'P/L (%)',
  'Доходность купонная годовая (прибл.)',

  // Блок 3 — прочие/ID/служебные
  'Номинал',
  'НКД',
  'Дата погашения',
  'Следующий купон',
  'Валюта',
  'Тикер',
  'Emitent / ISIN',
  'Лот',
  'Биржевой класс',
  'Класс риска TCS',
  'Время цены (PriceTime)'
];

function updateBondsFull() {
  var figis = readInputFigis_();
  if (!figis.length) {
    SpreadsheetApp.getActive().toast('Input пуст: нет FIGI', 'Bonds • Full', 5);
    return;
  }

  var bondsInfo  = fetchBondsInfo_(figis);        // профиль облигации (BondBy, кэш)
  var mdMap      = fetchBondsMarketData_(figis);  // цены/время/НКД
  var couponsMap = fetchBondsNextCoupons_(figis); // купонные события (кэш)
  var portfolio  = safeFetchAllPortfolios_();     // qty/avg по FIGI на аккаунт

  var rows = [];
  figis.forEach(function (figi) {
    var bi  = bondsInfo[figi]  || {};
    var md  = mdMap[figi]      || {};
    var cpn = couponsMap[figi] || {};

    // Приоритет купонных значений: события > карточка
    if (cpn.couponValue != null) bi.couponValue = cpn.couponValue;
    if (cpn.couponsPerYearFromEvent != null) bi.couponsPerYear = cpn.couponsPerYearFromEvent;
    if (!bi.nextCouponDate && cpn.nextCouponDate) bi.nextCouponDate = cpn.nextCouponDate;

    var couponTypeDesc = cpn.couponType ? mapCouponType_(cpn.couponType) : (bi.couponTypeDesc || '');

    // Достроить ставку/купон при недостающих полях
    if ((bi.couponRate == null || bi.couponRate === '') &&
        (bi.couponValue != null && bi.couponsPerYear != null && bi.nominal)) {
      bi.couponRate = (Number(bi.couponValue) * Number(bi.couponsPerYear) / Number(bi.nominal)) * 100;
    }
    if ((bi.couponValue == null || bi.couponValue === '') &&
        (bi.couponRate != null && bi.nominal != null && bi.couponsPerYear != null && Number(bi.couponsPerYear) > 0)) {
      bi.couponValue = Number(bi.nominal) * Number(bi.couponRate) / 100 / Number(bi.couponsPerYear);
    }

    var lastPriceAdj = bondPricePctToCurrency_(md.lastPrice, bi.nominal);

    var posArr = portfolio[figi] || [null]; // одна строка на аккаунт (если есть позиции)
    posArr.forEach(function (pos) {
      var qty    = pos ? pos.qty : '';
      var avgRaw = pos ? (pos.avg != null ? pos.avg : (pos.avg_fifo != null ? pos.avg_fifo : null)) : null;
      var avgAdj = bondPricePctToCurrency_(avgRaw, bi.nominal);

      rows.push([
        '',                                      // 📌 Риск (ручн.)
        '',                                      // Комментарий ИИ
        '',                                      // Комментарий пользователя
        'Облигация',                             // Тип инструмента
        figi,                                    // FIGI
        bi.name || '',                           // Название
        qty,                                     // Кол-во
        (avgAdj != null ? avgAdj : ''),          // Средняя цена
        (lastPriceAdj != null ? lastPriceAdj : ''), // Текущая цена
        (bi.couponRate != null ? bi.couponRate : ''),   // Купон, %
        (bi.couponValue != null ? bi.couponValue : ''), // Размер купона
        (bi.couponsPerYear != null ? bi.couponsPerYear : ''), // купон/год
        (couponTypeDesc || ''),                  // Тип купона (desc)
        (bi.sector || ''),                       // Сектор
        (bi.riskLevelDesc || ''),                // Риск (уровень TCS)

        '', '', '', '', '',                      // расчётные — формулы ниже

        (bi.nominal != null ? bi.nominal : ''),  // Номинал
        (md.accrued != null ? md.accrued : ''),  // НКД
        (bi.maturityDate || ''),                 // Дата погашения
        (bi.nextCouponDate || ''),               // Следующий купон
        (bi.currency || ''),                     // Валюта
        (bi.ticker || ''),                       // Тикер
        (bi.isin || ''),                         // ISIN
        (bi.lot || ''),                          // Лот
        (bi.classCode || ''),                    // Биржевой класс
        (bi.riskClass || ''),                    // Класс риска TCS (raw)
        (md.lastTime || '')                      // Время цены
      ]);
    });
  });

  var sh = getOrCreateSheet_('Bonds');
  sh.clear();
  sh.getRange(1, 1, 1, BONDS_HEADERS.length).setValues([BONDS_HEADERS]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, BONDS_HEADERS.length).setValues(rows);
    applyFormulas_(sh, 2, rows.length);
  }
  sh.autoResizeColumns(1, BONDS_HEADERS.length);
  sh.setFrozenRows(1);
}

/** ===================== FETCHERS (используют API-обёртки) ==================== */
function fetchBondsInfo_(figis) {
  var out = {};
  figis.forEach(function (figi) {
    var bi = callInstrumentsBondByFigi_(figi);
    if (!bi) return;

    var nominal        = (bi.nominal != null) ? (qToNumber(bi.nominal) ?? moneyToNumber(bi.nominal)) : null;
    var couponRate     = (bi.couponRate != null) ? qToNumber(bi.couponRate) : null;
    var couponValue    = (bi.couponValue != null) ? moneyToNumber(bi.couponValue) :
                         (bi.couponNominal != null ? moneyToNumber(bi.couponNominal) : null);
    var couponsPerYear = (bi.couponQuantityPerYear != null) ? Number(bi.couponQuantityPerYear) : null;
    var maturityDate   = dateOrEmpty_(tsToIso(bi.maturityDate));
    var nextCouponDate = dateOrEmpty_(tsToIso(bi.nextCouponDate));

    out[figi] = {
      ticker: bi.ticker || '',
      name: bi.name || bi.placementName || '',
      currency: bi.currency || bi.currencyCode || '',
      nominal: nominal,
      couponRate: couponRate,
      couponValue: couponValue,
      couponsPerYear: couponsPerYear,
      couponTypeDesc: mapCouponType_(bi.couponType),
      maturityDate: maturityDate,
      nextCouponDate: nextCouponDate,
      isin: bi.isin || '',
      lot: bi.lot || 1,
      classCode: bi.classCode || '',
      sector: bi.sector || '',
      riskLevelDesc: mapRiskLevel_(bi.riskLevel),
      riskClass: bi.riskLevel || ''
    };
  });
  return out;
}

function fetchBondsNextCoupons_(figis) {
  var now = new Date();
  var fromIso = new Date(now.getTime() - 30*24*3600*1000).toISOString();
  var toIso   = new Date(now.getTime() + 3*365*24*3600*1000).toISOString();
  var map = {};
  figis.forEach(function (figi) {
    try {
      var arr = callInstrumentsGetBondCouponsCached_(figi, fromIso, toIso) || [];
      var future = null, lastPast = null;

      arr.forEach(function (c) {
        var dtIso = tsToIso(c.couponDate || c.coupon_date || c.couponDateLt || c.date);
        if (!dtIso) return;
        var dt = new Date(dtIso);
        var val = moneyToNumber(c.payOneBond || c.pay_one_bond || c.couponValue || c.value);
        var typ = c.couponType || c.coupon_type;
        var per = (c.couponPeriod != null ? Number(c.couponPeriod) :
                  (c.coupon_period != null ? Number(c.coupon_period) : null));
        var rec = { dtIso: dtIso, value: (val != null ? Number(val) : null), type: typ, period: per };
        if (dt.getTime() >= now.getTime()) {
          if (!future || new Date(future.dtIso) > dt) future = rec;
        } else {
          if (!lastPast || new Date(lastPast.dtIso) < dt) lastPast = rec;
        }
      });

      var chosen = future || lastPast;
      if (chosen) {
        var cps = (chosen.period && chosen.period > 0) ? Math.max(1, Math.round(365 / chosen.period)) : null;
        map[figi] = {
          nextCouponDate: dateOrEmpty_(chosen.dtIso),
          couponValue: chosen.value,
          couponType: chosen.type,
          couponPeriodDays: chosen.period,
          couponsPerYearFromEvent: cps
        };
      }
    } catch (e) { /* пропускаем конкретный FIGI */ }
  });
  return map;
}

function fetchBondsMarketData_(figis) {
  var out = {};
  var last = callMarketLastPrices_(figis);
  last.forEach(function (x) {
    out[x.figi] = out[x.figi] || {};
    out[x.figi].lastPrice = (x.lastPrice != null) ? Number(x.lastPrice) : null;
    out[x.figi].lastTime  = x.time || '';
  });
  figis.forEach(function (f) {
    var aci = callMarketAccruedInterestsToday_(f);
    out[f] = out[f] || {};
    out[f].accrued = (aci != null) ? Number(aci) : null;
  });
  return out;
}

function safeFetchAllPortfolios_() {
  try {
    var accounts = callUsersGetAccounts_();
    var map = {};
    accounts.forEach(function (a) {
      var byPortfolio = callPortfolioGetPortfolio_(a.accountId);
      var byPositions = callPortfolioGetPositions_(a.accountId);
      var combined = byPortfolio.concat(byPositions);

      combined.forEach(function (p) {
        if (!p.figi) return;
        var qty = (p.quantity != null) ? Number(p.quantity) : null;
        var avg = (p.avg != null) ? Number(p.avg) :
                  ((p.avg_fifo != null) ? Number(p.avg_fifo) : null);
        map[p.figi] = map[p.figi] || [];
        var exist = map[p.figi].find(function (x) { return x.accountId === a.accountId; });
        if (exist) {
          if (exist.qty == null && qty != null) exist.qty = qty;
          if (exist.avg == null && avg != null) exist.avg = avg;
        } else {
          map[p.figi].push({
            accountId: a.accountId,
            accountName: a.name || '',
            qty: qty,
            avg: avg,
            avg_fifo: (p.avg_fifo != null ? Number(p.avg_fifo) : null)
          });
        }
      });
    });
    return map;
  } catch (e) { return {}; }
}

/** =================== Формулы (локаль-aware, ROUND до 2 знаков) =================== */
function applyFormulas_(sh, startRow, numRows) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  function idx(title) { return headers.indexOf(title) + 1; }
  var locale = SpreadsheetApp.getActive().getSpreadsheetLocale()||'';
  var SEP = (/^ru|^uk|^pl|^de|^fr|^it/i.test(locale)) ? ';' : ',';

  var cQty   = idx('Кол-во');
  var cAvg   = idx('Средняя цена');
  var cPrice = idx('Текущая цена');

  var cInv   = idx('Инвестировано');
  var cMkt   = idx('Рыночная стоимость');
  var cPL    = idx('P/L (руб)');
  var cPLPct = idx('P/L (%)');

  var cCupVal = idx('Размер купона');
  var cCupPY  = idx('купон/год');
  var cYield  = idx('Доходность купонная годовая (прибл.)');

  function d(from, to){ return from - to; }
  var R2 = function(expr){ return 'ROUND(' + expr + SEP + '2)'; };

  // Инвестировано = Qty * Avg
  sh.getRange(startRow, cInv, numRows, 1)
    .setFormulaR1C1('=IF(OR(LEN(RC['+d(cQty,cInv)+'])=0'+SEP+'LEN(RC['+d(cAvg,cInv)+'])=0)'+SEP+'""'+SEP+ R2('RC['+d(cQty,cInv)+']*RC['+d(cAvg,cInv)+']') +')');

  // Рыночная = Qty * Price
  sh.getRange(startRow, cMkt, numRows, 1)
    .setFormulaR1C1('=IF(OR(LEN(RC['+d(cQty,cMkt)+'])=0'+SEP+'LEN(RC['+d(cPrice,cMkt)+'])=0)'+SEP+'""'+SEP+ R2('RC['+d(cQty,cMkt)+']*RC['+d(cPrice,cMkt)+']') +')');

  // P/L (руб) = Рыночная - Инвестировано
  sh.getRange(startRow, cPL, numRows, 1)
    .setFormulaR1C1('=IF(OR(LEN(RC['+d(cMkt,cPL)+'])=0'+SEP+'LEN(RC['+d(cInv,cPL)+'])=0)'+SEP+'""'+SEP+ R2('RC['+d(cMkt,cPL)+']-RC['+d(cInv,cPL)+']') +')');

  // P/L (%) = P/L / Инвестировано * 100
  sh.getRange(startRow, cPLPct, numRows, 1)
    .setFormulaR1C1('=IF(OR(LEN(RC['+d(cInv,cPLPct)+'])=0'+SEP+'RC['+d(cInv,cPLPct)+']=0'+SEP+'LEN(RC['+d(cPL,cPLPct)+'])=0)'+SEP+'""'+SEP+ R2('(RC['+d(cPL,cPLPct)+']/RC['+d(cInv,cPLPct)+'])*100') +')');

  // Доходность купонная годовая (прибл.) = (РазмерКупона * купон/год) / Цена * 100
  // (ИСПРАВЛЕНО: LEN вместо ЛЕН)
  sh.getRange(startRow, cYield, numRows, 1)
    .setFormulaR1C1('=IF(OR(LEN(RC['+d(cCupVal,cYield)+'])=0'+SEP+'LEN(RC['+d(cCupPY,cYield)+'])=0'+SEP+'LEN(RC['+d(cPrice,cYield)+'])=0'+SEP+'RC['+d(cPrice,cYield)+']=0)'+SEP+'""'+SEP+ R2('(RC['+d(cCupVal,cYield)+']*RC['+d(cCupPY,cYield)+'])/RC['+d(cPrice,cYield)+']*100') +')');
}

/**
 * Обновляет на листе «Bonds» только рыночные поля: Текущая цена/НКД/Время цены.
 * Использует унифицированную конверсию bondPricePctToCurrency_().
 */
function updateBondPricesOnly() {
  var figis = readInputFigis_();
  if (!figis.length) {
    SpreadsheetApp.getActive().toast('Input пуст: нет FIGI', 'Bonds • Prices', 5);
    return;
  }

  var mdBy = {};
  (callMarketLastPrices_(figis) || []).forEach(function (x) {
    mdBy[x.figi] = {
      lastPrice: (x.lastPrice != null) ? Number(x.lastPrice) : null,
      lastTime:  x.time || ''
    };
  });

  figis.forEach(function (f) {
    var aci = callMarketAccruedInterestsToday_(f);
    mdBy[f] = mdBy[f] || {};
    mdBy[f].accrued = (aci != null) ? Number(aci) : null;
  });

  // Получим номиналы (кэш ускорит)
  var nominalBy = {};
  figis.forEach(function (f) {
    try {
      var card = callInstrumentsBondByFigi_(f);
      if (card && card.nominal != null) {
        var n = (qToNumber(card.nominal) != null) ? qToNumber(card.nominal) : moneyToNumber(card.nominal);
        if (n != null) nominalBy[f] = Number(n);
      }
    } catch (e) { /* skip */ }
  });

  var sh = SpreadsheetApp.getActive().getSheetByName('Bonds');
  if (!sh) {
    SpreadsheetApp.getActive().toast('Лист Bonds ещё не создан (сначала полное обновление).', 'Bonds • Prices', 7);
    return;
  }
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  function idx(name){ return headers.indexOf(name) + 1; }

  var colFIGI  = idx('FIGI');
  var colPrice = idx('Текущая цена');
  var colACI   = idx('НКД');
  var colTime  = idx('Время цены (PriceTime)');
  if (!(colFIGI && colPrice && colACI && colTime)) {
    SpreadsheetApp.getActive().toast('Не найдены нужные колонки FIGI/Текущая/НКД/Время', 'Bonds • Prices', 7);
    return;
  }

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var figiVals = sh.getRange(2, colFIGI, lastRow - 1, 1).getValues().flat();

  var priceArr = [], aciArr = [], timeArr = [];

  figiVals.forEach(function (f) {
    var md = mdBy[f] || {};
    var raw = md.lastPrice;
    var nominal = nominalBy[f];
    var priceRub = (raw != null) ? bondPricePctToCurrency_(raw, nominal) : null;

    console.log('FIGI=' + f + ' nominal=' + (nominal||'') + ' raw=' + (raw||'') + ' priceRub=' + (priceRub||''));

    priceArr.push([priceRub != null ? priceRub : '']);
    aciArr.push([md.accrued != null ? md.accrued : '']);
    timeArr.push([md.lastTime || '']);
  });

  sh.getRange(2, colPrice, priceArr.length, 1).setValues(priceArr);
  sh.getRange(2, colACI,   aciArr.length,   1).setValues(aciArr);
  sh.getRange(2, colTime,  timeArr.length,  1).setValues(timeArr);

  SpreadsheetApp.getActive().toast('Цены и НКД обновлены', 'Bonds • Prices', 5);
}

// ====== Dashboard с LockService (новое) ======
function buildBondsDashboard(){
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActive().toast('Другая операция обновляет Dashboard. Повторите позже.', 'Dashboard', 5);
    return;
  }
  try {
    var ss = SpreadsheetApp.getActive();
    var src = ss.getSheetByName('Bonds');
    if(!src){ SpreadsheetApp.getActive().toast('Лист Bonds не найден', 'Dashboard', 5); return; }

    var dst = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');
    dst.clear();

    // --- Индексы колонок по заголовкам ---
    var hdr = src.getRange(1,1,1,src.getLastColumn()).getValues()[0];
    function idx(name){ var i = hdr.indexOf(name); return (i>=0)? (i+1) : 0; }

    var cName    = idx('Название');
    var cFIGI    = idx('FIGI');
    var cRiskNum = idx('📌 Риск (ручн.)');      // ручной риск
    var cSector  = idx('Сектор');
    var cQty     = idx('Кол-во');
    var cPrice   = idx('Текущая цена');
    var cNominal = idx('Номинал');
    var cCupPY   = idx('купон/год');
    var cCupVal  = idx('Размер купона');
    var cMaturity= idx('Дата погашения');
    var cNextCp  = idx('Следующий купон');

    var cMkt     = idx('Рыночная стоимость');
    var cInv     = idx('Инвестировано');
    var cPL      = idx('P/L (руб)');
    var cPLPct   = idx('P/L (%)');

    if(!(cRiskNum && cSector && cMkt && cInv && cPL && cPLPct && cPrice)){
      SpreadsheetApp.getActive().toast('Не хватает обязательных колонок (📌 Риск (ручн.)/Сектор/Рыночная/…)', 'Dashboard', 7);
      return;
    }

    var lastRow = src.getLastRow();
    if(lastRow < 2){ SpreadsheetApp.getActive().toast('Нет данных для сводки', 'Dashboard', 5); return; }

    var rows = src.getRange(2,1,lastRow-1,src.getLastColumn()).getValues();

    // --- Агрегации ---
    var invested = 0, market = 0, plRub = 0;
    var sectors = {};                       // {sector: sumMarket}
    var riskCnt = { 'Низкий':0, 'Средний':0, 'Высокий':0 }; // 0–1, 2–4, 5+
    var wCouponNum = 0, wCouponDen = 0;

    // YTM: средневзвешенное по рыночной стоимости (приближённо)
    var wYtmNum = 0, wYtmDen = 0;
    var scatterRiskYield = [['Риск','YTM (%)','Тултип']];

    // Купоны по месяцам (горизонт 6 месяцев)
    var monthAgg = {};
    var today = new Date();
    var horizon = addMonths_(today, 6);
    var strip   = function(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); };
    var r2      = function(x){ return Math.round(Number(x||0)*100)/100; };

    rows.forEach(function(r){
      function val(ci){ return (ci? r[ci-1] : ''); }

      // Риск: бинирование по ручному числу
      var riskNum = Number(val(cRiskNum));
      if(!isNaN(riskNum)){
        if (riskNum <= 1)      riskCnt['Низкий']++;
        else if (riskNum <= 4) riskCnt['Средний']++;
        else                   riskCnt['Высокий']++;
      }

      var sec = String(val(cSector) || 'other');
      var mkt = Number(val(cMkt)) || 0;
      var inv = Number(val(cInv)) || 0;
      var pl  = Number(val(cPL))  || 0;

      invested += inv; market += mkt; plRub += pl;
      sectors[sec] = (sectors[sec]||0) + mkt;

      var coupPctAltIdx = idx('Купон, %');
      var coupPctAlt = Number(coupPctAltIdx ? val(coupPctAltIdx) : NaN);
      var coupPctCalc = Number(val(cCupPY) ? (((Number(val(cCupVal))||0) * Number(val(cCupPY))) / (Number(val(cPrice))||0) * 100) : NaN);
      var coupUse = !isNaN(coupPctAlt) ? coupPctAlt : coupPctCalc;
      if(!isNaN(coupUse) && mkt>0){ wCouponNum += (coupUse/100)*mkt; wCouponDen += mkt; }

      // --- YTM (приближённо)
      var price    = Number(val(cPrice));
      var nominal  = Number(val(cNominal));
      var cupPY    = Number(val(cCupPY));
      var cupVal   = Number(val(cCupVal));
      var matStr   = val(cMaturity);
      var name     = String(val(cName) || '');
      var figi     = String(val(cFIGI) || '');
      var yearsToMat = null;

      if(matStr){
        var d = (matStr instanceof Date)? matStr : (isNaN(Date.parse(matStr))? null : new Date(Date.parse(matStr)));
        if(d){ yearsToMat = Math.max(0.25, (strip(d) - strip(today)) / (365*24*3600*1000)); }
      }

      if(!(nominal>0)) nominal = 1000;
      var ytmPct = null;
      if(price>0 && nominal>0 && cupPY>0 && cupVal>=0 && yearsToMat){
        var C = cupVal * cupPY;
        ytmPct = ((C + (nominal - price) / yearsToMat) / ((nominal + price) / 2)) * 100;
        if(isFinite(ytmPct)){
          if(mkt>0){ wYtmNum += (ytmPct/100)*mkt; wYtmDen += mkt; }
          if(!isNaN(riskNum)){
            var tip = name + '\n' + figi + '\nYTM: ' + r2(ytmPct) + '%';
            scatterRiskYield.push([riskNum, r2(ytmPct), tip]);
          }
        } else {
          ytmPct = null;
        }
      }

      // --- купоны по месяцам (6 мес)
      var qty = Number(val(cQty)) || 0;
      var nextStr = val(cNextCp);
      if(qty>0 && cupPY>0 && cupVal>0 && nextStr){
        var first = (nextStr instanceof Date)? nextStr : (isNaN(Date.parse(nextStr))? null : new Date(Date.parse(nextStr)));
        if(first){
          var periodDays = Math.max(15, Math.round(365 / cupPY));
          var d2 = strip(first);
          while(d2 <= horizon){
            if(d2 >= strip(today)){
              var key = d2.getFullYear() + '-' + ('0'+(d2.getMonth()+1)).slice(-2);
              monthAgg[key] = (monthAgg[key]||0) + (cupVal * qty);
            }
            d2 = addDays_(d2, periodDays);
          }
        }
      }
    });

    var plPctTotal = invested>0 ? (plRub/invested*100) : 0;
    var wCouponPct = wCouponDen>0 ? (wCouponNum/wCouponDen*100) : 0;
    var wYtmPct    = wYtmDen>0    ? (wYtmNum/wYtmDen*100)       : 0;

    // --- KPI блок ---
    var kpi = [
      ['Показатель','Значение'],
      ['Инвестировано', round2_(invested)],
      ['Рыночная стоимость', round2_(market)],
      ['P/L (руб)', round2_(plRub)],
      ['P/L (%)', round2_(plPctTotal)],
      ['Средневзв. купонная доходность (%)', round2_(wCouponPct)],
      ['Средневзв. YTM (%)', round2_(wYtmPct)],
      ['Купоны в 30 дней (шт)', Object.keys(monthAgg).filter(function(k){
        var y = Number(k.slice(0,4)), m = Number(k.slice(5,7))-1;
        var d = new Date(y,m,1); return d <= addMonths_(today,1);
      }).length],
      ['Купоны в 30 дней (₽)', (function(){
        var s=0; Object.keys(monthAgg).forEach(function(k){
          var y = Number(k.slice(0,4)), m = Number(k.slice(5,7))-1;
          var d = new Date(y,m,1); if(d <= addMonths_(today,1)) s += monthAgg[k];
        }); return round2_(s);
      })()]
    ];
    dst.getRange(1,1,kpi.length,2).setValues(kpi);
    dst.getRange(1,1,1,2).setFontWeight('bold');

    // --- Таблица по рискам ---
    var riskTable = [
      ['Категория риска','Кол-во'],
      ['Низкий',  riskCnt['Низкий']],
      ['Средний', riskCnt['Средний']],
      ['Высокий', riskCnt['Высокий']]
    ];
    dst.getRange(1,4,riskTable.length,riskTable[0].length).setValues(riskTable).setFontWeight('bold');

    // --- Таблица по секторам ---
    var secArr = Object.keys(sectors).sort().map(function(s){ return [s, round2_(sectors[s])]; });
    dst.getRange(1,7,1,2).setValues([['Сектор','Рыночная стоимость']]).setFontWeight('bold');
    if(secArr.length) dst.getRange(2,7,secArr.length,2).setValues(secArr);

    // --- Таблица по годам погашения ---
    var cMaturity2 = cMaturity;
    var byYear = {};
    if(cMaturity2){
      rows.forEach(function(r){
        var mkt = Number(r[cMkt-1])||0;
        var v = r[cMaturity2-1];
        if(v){
          var d = (v instanceof Date)? v : (isNaN(Date.parse(v))? null : new Date(Date.parse(v)));
          if(d){ var y = d.getFullYear(); byYear[y] = (byYear[y]||0) + mkt; }
        }
      });
    }
    var years = Object.keys(byYear).sort();
    dst.getRange(1,10,1,2).setValues([['Год погашения','Рыночная стоимость']]).setFontWeight('bold');
    if(years.length){
      var yrArr = years.map(function(y){ return [Number(y), round2_(byYear[y])]; });
      dst.getRange(2,10,yrArr.length,2).setValues(yrArr);
    }

    // --- Таблица купонов по месяцам ---
    var monthsSorted = Object.keys(monthAgg).sort();
    dst.getRange(1,13,1,2).setValues([['Месяц','Купоны (₽)']]).setFontWeight('bold');
    if(monthsSorted.length){
      var monArr = monthsSorted.map(function(k){ return [k, round2_(monthAgg[k])]; });
      dst.getRange(2,13,monArr.length,2).setValues(monArr);
    }

    // Очистить старые диаграммы
    dst.getCharts().forEach(function(ch){ dst.removeChart(ch); });

    // Палитры
    var paletteMain = ['#4F46E5','#22C55E','#EAB308','#EF4444','#06B6D4','#A855F7','#F59E0B','#94A3B8','#10B981','#3B82F6'];
    var paletteRisk = ['#10B981','#EAB308','#EF4444'];

    // Диаграмма: Риски
    var riskRange = dst.getRange(1,4,4,2);
    var riskChart = dst.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(riskRange)
      .setPosition(1, 16, 0, 0)
      .setOption('title','Распределение по рискам (шт.)')
      .setOption('legend', { position: 'none' })
      .setOption('colors', [paletteRisk[0]])
      .build();
    dst.insertChart(riskChart);

    // Диаграмма: Сектора
    if(secArr.length){
      var secRange = dst.getRange(1,7,Math.max(2,secArr.length+1),2);
      var pie = dst.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(secRange)
        .setPosition(20, 1, 0, 0)
        .setOption('title','Структура по секторам (рыночная стоимость)')
        .setOption('legend', { position: 'right' })
        .setOption('pieSliceText', 'percentage')
        .setOption('colors', paletteMain.slice(0, Math.max(3, secArr.length)))
        .build();
      dst.insertChart(pie);
    }

    // Диаграмма: Погашения по годам
    if(years.length){
      var yrRange = dst.getRange(1,10,Math.max(2,years.length+1),2);
      var matChart = dst.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(yrRange)
        .setPosition(20, 7, 0, 0)
        .setOption('title','Сроки погашения (рыночная стоимость)')
        .setOption('legend', { position: 'none' })
        .setOption('colors', ['#3B82F6'])
        .build();
      dst.insertChart(matChart);
    }

    // Диаграмма: Купоны по месяцам
    if(monthsSorted.length){
      var monRange = dst.getRange(1,13,Math.max(2,monthsSorted.length+1),2);
      var monChart = dst.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(monRange)
        .setPosition(20, 13, 0, 0)
        .setOption('title','График купонных выплат (6 месяцев)')
        .setOption('legend', { position: 'none' })
        .setOption('colors', ['#22C55E'])
        .build();
      dst.insertChart(monChart);
    }

    // Диаграмма: YTM vs Купонная доходность
    var cmpStartRow = Math.max(22, 20 + Math.max(years.length, monthsSorted.length) + 2);
    var cmpData = [
      ['Метрика','Значение'],
      ['Средневзв. YTM (%)', round2_(wYtmPct)],
      ['Купонная доходность (%)', round2_(wCouponPct)]
    ];
    dst.getRange(cmpStartRow, 1, cmpData.length, 2).setValues(cmpData).setFontWeight('bold');
    var cmpRange = dst.getRange(cmpStartRow, 1, cmpData.length, 2);
    var cmpChart = dst.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(cmpRange)
      .setPosition(cmpStartRow, 4, 0, 0)
      .setOption('title','YTM vs Купонная доходность (средневзв., %)')
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#4F46E5'])
      .build();
    dst.insertChart(cmpChart);

    // Диаграмма: Риск vs YTM (scatter)
    if(scatterRiskYield.length > 1){
      dst.getRange(cmpStartRow, 7, scatterRiskYield.length, 3).setValues(scatterRiskYield).setFontWeight('bold');
      var scRange = dst.getRange(cmpStartRow, 7, scatterRiskYield.length, 3);
      var scChart = dst.newChart()
        .setChartType(Charts.ChartType.SCATTER)
        .addRange(scRange)
        .setPosition(cmpStartRow, 11, 0, 0)
        .setOption('title','Риск vs Доходность к погашению (YTM)')
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', { title: 'Риск (баллы)' })
        .setOption('vAxis', { title: 'YTM (%)' })
        .setOption('series', { 0: { pointSize: 5 } })
        .build();
      dst.insertChart(scChart);
    }

    dst.autoResizeColumns(1, 20);
    SpreadsheetApp.getActive().toast('Dashboard обновлён: кэш, фикс формул и лок-синхронизация', 'Dashboard', 4);
  } finally {
    lock.releaseLock();
  }
}

/** helpers */
function addDays_(d, n){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()+n); }
function addMonths_(d, n){ return new Date(d.getFullYear(), d.getMonth()+n, 1); }
