/**
 * bonds_update.gs
 * Лист «Bonds»: заголовки, полное обновление, формулы (включая Риск по Rules), обновление только цен.
 */

const BONDS_SHEET = 'Bonds';

// ===== УНИКАЛЬНЫЕ заголовки  =====
const BONDS_HEADERS = [
  // Блок 1 — приоритетные
  'Риск (ручн.)',            // числовой скоринг по Rules
  'Комментарий ИИ',
  'Комментарий пользователя',
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
  'Риск (уровень TCS)',

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
  'Время цены (PriceTime)',
  'Тип инструмента'          // перенесено в конец; везде «Облигация»
];

function updateBondsFull() {
  var figis = readInputFigisByType_('bond');
  if (!figis.length) { showSnack_('Input пуст: нет FIGI в колонке Облигации','Bonds',2500); return; }

  var bondsInfo  = fetchBondsInfo_(figis);
  var mdMap      = fetchBondsMarketData_(figis);
  var couponsMap = fetchBondsNextCoupons_(figis);
  var portfolio  = safeFetchAllPortfolios_();

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

    // Достроить ставку/купон
    if ((bi.couponRate == null || bi.couponRate === '') &&
        (bi.couponValue != null && bi.couponsPerYear != null && bi.nominal)) {
      bi.couponRate = (Number(bi.couponValue) * Number(bi.couponsPerYear) / Number(bi.nominal)) * 100;
    }
    if ((bi.couponValue == null || bi.couponValue === '') &&
        (bi.couponRate != null && bi.nominal != null && bi.couponsPerYear != null && Number(bi.couponsPerYear) > 0)) {
      bi.couponValue = Number(bi.nominal) * Number(bi.couponRate) / 100 / Number(bi.couponsPerYear);
    }

    var lastPriceAdj = bondPricePctToCurrency_(md.lastPrice, bi.nominal);

    var posArr = portfolio[figi] || [null];
    posArr.forEach(function (pos) {
      var qty    = pos ? pos.qty : '';
      var avgRaw = pos ? (pos.avg != null ? pos.avg : (pos.avg_fifo != null ? pos.avg_fifo : null)) : null;
      var avgAdj = bondPricePctToCurrency_(avgRaw, bi.nominal);

      rows.push([
        '',                                      // Риск (ручн.) — формула поставится ниже
        '',                                      // Комментарий ИИ
        '',                                      // Комментарий пользователя
        figi,                                    // FIGI
        bi.name || '',                           // Название
        qty,                                     // Кол-во
        (avgAdj != null ? avgAdj : ''),          // Средняя цена
        (lastPriceAdj != null ? lastPriceAdj : ''), // Текущая цена
        (bi.couponRate != null ? bi.couponRate : ''),   // Купон, %
        (bi.couponValue != null ? bi.couponValue : ''), // Размер купона
        (bi.couponsPerYear != null ? bi.couponsPerYear : ''), // купон/год
        (couponTypeDesc || ''),                  // Тип купона
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
        (md.lastTime || ''),                     // Время цены
        'Облигация'                              // Тип инструмента (служебный)
      ]);
    });
  });

  var sh = getOrCreateSheet_(BONDS_SHEET);
  sh.clear();
  sh.getRange(1, 1, 1, BONDS_HEADERS.length).setValues([BONDS_HEADERS]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, BONDS_HEADERS.length).setValues(rows);
    applyFormulas_(sh, 2, rows.length);        // все формулы, включая Риск (ручн.)
  }
  sh.autoResizeColumns(1, BONDS_HEADERS.length);
  sh.setFrozenRows(1);
  showSnack_('Обновление листа завершено','Bonds • Full',2000);
}

/** ===================== FETCHERS ==================== */
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
    } catch (e) { /* skip */ }
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

// ===== Вспомогательное: номер колонки → буквы A1 =====
function colToA1_(col){
  var s = '';
  while(col > 0){
    var m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = (col - m - 1) / 26 >> 0;
  }
  return s;
}

// =================== Формулы (локаль-aware, ROUND; вкл. Риск по Rules) ===================
function applyFormulas_(sh, startRow, numRows) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  function idx(title) { return headers.indexOf(title) + 1; }
  var locale = SpreadsheetApp.getActive().getSpreadsheetLocale()||'';
  var SEP = (/^ru|^uk|^pl|^de|^fr|^it/i.test(locale)) ? ';' : ',';

  var cRisk  = idx('Риск (ручн.)');
  var cFIGI  = idx('FIGI');
  var cName  = idx('Название');
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

  var cSector = idx('Сектор');
  var cMat    = idx('Дата погашения');

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
  sh.getRange(startRow, cYield, numRows, 1)
    .setFormulaR1C1('=IF(OR(LEN(RC['+d(cCupVal,cYield)+'])=0'+SEP+'LEN(RC['+d(cCupPY,cYield)+'])=0'+SEP+'LEN(RC['+d(cPrice,cYield)+'])=0'+SEP+'RC['+d(cPrice,cYield)+']=0)'+SEP+'""'+SEP+ R2('(RC['+d(cCupVal,cYield)+']*RC['+d(cCupPY,cYield)+'])/RC['+d(cPrice,cYield)+']*100') +')');


  
  // === Риск (ручн.) — накопительный: база TCS (0/2/4) + отрасль + просадка + дюрация ===
  var cRiskTcs = idx('Риск (уровень TCS)');   // текст: Низкий/Средний/Высокий
  var colSectorA1 = colToA1_(cSector);
  var colPLPctA1  = colToA1_(cPLPct);
  var colMatA1    = colToA1_(cMat);
  var colTcsA1    = colToA1_(cRiskTcs);

  var formulas = [];
  for (var r = startRow; r < startRow + numRows; r++) {
    var aSector = colSectorA1 + r;  // Сектор
    var aPLPct  = colPLPctA1  + r;  // P/L (%)
    var aMat    = colMatA1    + r;  // Дата погашения
    var aTcs    = colTcsA1    + r;  // Риск (уровень TCS)

    var baseTcs =
      'ЕСЛИ(' + aTcs + '="Низкий";0;ЕСЛИ(' + aTcs + '="Средний";2;ЕСЛИ(' + aTcs + '="Высокий";4;0)))';

    var f =
      '=ЕСЛИОШИБКА(' +
        baseTcs +
        '+ВПР(' + aSector + ';RISK_SECTORS;2;ЛОЖЬ)' +
        '+ЕСЛИ(ДЛСТР(' + aPLPct + ')=0;0;ЕСЛИ(' + aPLPct + '<=RISK_DD;1;0))' +
        '+ЕСЛИ(ДЛСТР(' + aMat + ')=0;0;ЕСЛИ(((' + aMat + '-СЕГОДНЯ())/365)<=RISK_DUR;1;0))' +
      ';0)';

    formulas.push([f]);
  }
sh.getRange(startRow, cRisk, numRows, 1).setFormulas(formulas);

}

/**
 * Обновляет на листе «Bonds» только рыночные поля и время.
 * Использует унифицированную конверсию.
 */
function updateBondPricesOnly() {
  var figis = readInputFigisByType_('bond');
  if (!figis.length) { showSnack_('Input пуст: нет FIGI','Bonds • Prices',2500); return; }

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

  var nominalBy = {};
  figis.forEach(function (f) {
    try {
      var card = callInstrumentsBondByFigi_(f);
      if (card && card.nominal != null) {
        var n = (qToNumber(card.nominal) != null) ? qToNumber(card.nominal) : moneyToNumber(card.nominal);
        if (n != null) nominalBy[f] = Number(n);
      }
    } catch (e) {}
  });

  var sh = SpreadsheetApp.getActive().getSheetByName(BONDS_SHEET);
  if (!sh) { showSnack_('Лист Bonds ещё не создан (сначала полное обновление).','Bonds • Prices',3000); return; }
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  function idx(name){ return headers.indexOf(name) + 1; }

  var colFIGI  = idx('FIGI');
  var colPrice = idx('Текущая цена');
  var colACI   = idx('НКД');
  var colTime  = idx('Время цены (PriceTime)');
  if (!(colFIGI && colPrice && colACI && colTime)) { showSnack_('Не найдены FIGI/Текущая/НКД/Время','Bonds • Prices',3000); return; }

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var figiVals = sh.getRange(2, colFIGI, lastRow - 1, 1).getValues().flat();

  var priceArr = [], aciArr = [], timeArr = [];
  figiVals.forEach(function (f) {
    var md = mdBy[f] || {};
    var raw = md.lastPrice;
    var nominal = nominalBy[f];
    var priceRub = (raw != null) ? bondPricePctToCurrency_(raw, nominal) : null;

    priceArr.push([priceRub != null ? priceRub : '']);
    aciArr.push([md.accrued != null ? md.accrued : '']);
    timeArr.push([md.lastTime || '']);
  });

  sh.getRange(2, colPrice, priceArr.length, 1).setValues(priceArr);
  sh.getRange(2, colACI,   aciArr.length,   1).setValues(aciArr);
  sh.getRange(2, colTime,  timeArr.length,  1).setValues(timeArr);

  showSnack_('Цены и НКД обновлены','Bonds • Prices',2000);
}
