/**
 * update_shares.gs
 * Лист «Shares»: полное обновление и обновление только рыночных данных.
 */

const SHARES_SHEET = 'Shares';

const SHARES_HEADERS = [
  // Блок 1 — приоритетные
  'ИИ',
  'Заметка',

  // Блок 2 — идентификация
  'FIGI',
  'Тикер',
  'Название',

  // Блок 3 — базовые числа
  'Кол-во',
  'Средняя цена',
  'Текущая цена',

  // Блок 4 — расчётные
  'Инвестировано',
  'Рыночная стоимость',
  'P/L (руб)',
  'P/L (%)',

  // Блок 5 — атрибуты и флаги
  'Сектор',
  'Страна риска',
  'Валюта',
  'Биржа',
  'Тип акции',
  'Шорт доступен',
  'Заблокирован (TCS)',

  // Блок 6 — служебные
  'Asset UID',
  'Время цены',

  // Блок 7 — long / фундаментал
  'Капитализация',
  'Выручка TTM',
  'EBITDA TTM',
  'Чистая прибыль TTM',
  'EPS TTM',
  'P/E TTM',
  'P/S TTM',
  'P/B TTM',
  'EV/EBITDA',
  'Debt/Equity',
  'NetDebt/EBITDA',
  'Free Float',
  'Beta',
  'Shares Outstanding',
  '52w High',
  '52w Low',
  'ROE',
  'ROA',
  'ROIC'
];

// ===== Публичные действия =====

function updateSharesFull() {
  setStatus_('Shares • полное обновление…');

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHARES_SHEET) || ss.insertSheet(SHARES_SHEET);
  sh.clear();
  sh.getRange(1, 1, 1, SHARES_HEADERS.length).setValues([SHARES_HEADERS]);
  sh.setFrozenRows(1);

  var pfMap = fetchSharesPortfolioMap_();
  var portfolioFigis = Object.keys(pfMap || {});
  if (!portfolioFigis.length) {
    formatSharesSheet_(sh);
    showSnack_('Нет текущих позиций в портфеле', 'Shares', 2500);
    return;
  }

  var shareFigis = [];
  portfolioFigis.forEach(function(figi) {
    try {
      if (callInstrumentsShareByFigi_(figi)) shareFigis.push(figi);
    } catch (e) {}
  });

  if (!shareFigis.length) {
    formatSharesSheet_(sh);
    showSnack_('В текущем портфеле не найдено акций', 'Shares', 2500);
    return;
  }

  var infoBy = fetchSharesInfo_(shareFigis);
  var mdBy = fetchSharesMarketData_(shareFigis);

  var assetByUid = {};
  var assetUids = [];
  var seenAssetUids = {};

  shareFigis.forEach(function(figi) {
    var info = infoBy[figi] || {};
    var assetUid = info.assetUid || '';
    if (!assetUid || seenAssetUids[assetUid]) return;

    seenAssetUids[assetUid] = true;
    assetUids.push(assetUid);

    if (typeof callInstrumentsGetAssetByUid_ === 'function') {
      try {
        assetByUid[assetUid] = callInstrumentsGetAssetByUid_(assetUid) || null;
      } catch (e) {
        assetByUid[assetUid] = null;
      }
    } else {
      assetByUid[assetUid] = null;
    }
  });

  var fundamentalsByUid = fetchSharesFundamentalsByAssetUid_(assetUids);

  var rows = [];
  shareFigis.forEach(function(figi) {
    var pf = pfMap[figi] || [];
    var pos = summarizeSharePosition_(pf);

    var info = infoBy[figi] || {};
    var md = mdBy[figi] || {};
    var assetUid = info.assetUid || '';
    var asset = assetUid ? (assetByUid[assetUid] || null) : null;
    var brand = asset && asset.brand ? asset.brand : null;
    var f = assetUid ? (fundamentalsByUid[assetUid] || {}) : {};

    rows.push([
      '',                                                      // ИИ
      '',                                                      // Заметка

      figi,                                                    // FIGI
      info.ticker || '',                                       // Тикер
      firstNonEmpty_(info.name, brand && brand.name, ''),      // Название

      (pos.qty != null ? pos.qty : ''),                        // Кол-во
      (pos.avg != null ? pos.avg : ''),                        // Средняя цена
      (md.lastPrice != null ? md.lastPrice : ''),              // Текущая цена

      '', '', '', '',                                          // расчётные — формулы ниже

      firstNonEmpty_(info.sector, brand && brand.sector, ''),  // Сектор
      firstNonEmpty_(info.countryOfRisk, asset && asset.countryOfRisk, ''),
      firstNonEmpty_(info.currency, asset && asset.currency, ''),
      firstNonEmpty_(info.exchange, info.classCode, ''),
      info.shareType || '',
      boolToRu_(info.shortEnabled),
      boolToRu_(info.blockedTcs),

      assetUid || '',
      md.lastTime || '',

      (f.marketCap != null ? f.marketCap : ''),
      (f.revenueTtm != null ? f.revenueTtm : ''),
      (f.ebitdaTtm != null ? f.ebitdaTtm : ''),
      (f.netIncomeTtm != null ? f.netIncomeTtm : ''),
      (f.epsTtm != null ? f.epsTtm : ''),
      (f.peRatioTtm != null ? f.peRatioTtm : ''),
      (f.priceToSalesTtm != null ? f.priceToSalesTtm : ''),
      (f.priceToBookTtm != null ? f.priceToBookTtm : ''),
      (f.evToEbitda != null ? f.evToEbitda : ''),
      (f.debtToEquity != null ? f.debtToEquity : ''),
      (f.netDebtToEbitda != null ? f.netDebtToEbitda : ''),
      (f.freeFloat != null ? f.freeFloat : ''),
      (f.beta != null ? f.beta : ''),
      (f.sharesOutstanding != null ? f.sharesOutstanding : ''),
      (f.high52w != null ? f.high52w : ''),
      (f.low52w != null ? f.low52w : ''),
      (f.roe != null ? f.roe : ''),
      (f.roa != null ? f.roa : ''),
      (f.roic != null ? f.roic : '')
    ]);
  });

  if (rows.length) {
    var bad = -1;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].length !== SHARES_HEADERS.length) {
        bad = i;
        break;
      }
    }
    if (bad !== -1) {
      throw new Error('Ширина строки №' + (bad + 2) + ' = ' + rows[bad].length + ' != ' + SHARES_HEADERS.length + ' (SHARES_HEADERS)');
    }

    sh.getRange(2, 1, rows.length, SHARES_HEADERS.length).setValues(rows);
    applySharesFormulas_(sh, 2, rows.length);
  }

  formatSharesSheet_(sh);
  showSnack_('Готово: Shares обновлён (' + rows.length + ' строк)', 'Shares', 2500);
}

function updateSharePricesOnly() {
  setStatus_('Shares • только цены…');

  var sh = SpreadsheetApp.getActive().getSheetByName(SHARES_SHEET);
  if (!sh) {
    showSnack_('Лист Shares не найден', 'Shares', 2000);
    return;
  }

  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    showSnack_('Нет данных для обновления цен', 'Shares', 2000);
    return;
  }

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var colFigi = headers.indexOf('FIGI') + 1;
  var colPrice = headers.indexOf('Текущая цена') + 1;
  var colTime = headers.indexOf('Время цены') + 1;

  if (!(colFigi && colPrice && colTime)) {
    showSnack_('Не найдены колонки FIGI / Текущая цена / Время цены', 'Shares', 2500);
    return;
  }

  var figiVals = sh.getRange(2, colFigi, lastRow - 1, 1).getValues().flat();
  var figis = [];
  var seen = {};

  figiVals.forEach(function(figi) {
    figi = String(figi || '').trim();
    if (!figi || seen[figi]) return;
    seen[figi] = true;
    figis.push(figi);
  });

  if (!figis.length) {
    showSnack_('FIGI не найдены', 'Shares', 2000);
    return;
  }

  var mdBy = fetchSharesMarketData_(figis);
  var priceArr = [];
  var timeArr = [];

  figiVals.forEach(function(figi) {
    figi = String(figi || '').trim();
    var md = figi ? (mdBy[figi] || {}) : {};
    priceArr.push([md.lastPrice != null ? md.lastPrice : '']);
    timeArr.push([md.lastTime || '']);
  });

  sh.getRange(2, colPrice, priceArr.length, 1).setValues(priceArr);
  sh.getRange(2, colTime, timeArr.length, 1).setValues(timeArr);

  applySharesFormulas_(sh, 2, lastRow - 1);
  formatSharesSheet_(sh);

  showSnack_('Цены обновлены и формулы пересчитаны', 'Shares • Prices', 2000);
}

// ===== Формулы =====

function applySharesFormulas_(sh, startRow, numRows) {
  if (!numRows || numRows < 1) return;

  var idx = function(name) { return SHARES_HEADERS.indexOf(name) + 1; };

  var loc = (SpreadsheetApp.getActive().getSpreadsheetLocale() || '').toLowerCase();
  var SEP = /^(bg|cs|da|de|el|es|et|fi|fr|hr|hu|it|lt|lv|nl|pl|pt|ro|ru|sk|sl|sr|sv|tr)/.test(loc) ? ';' : ',';

  var cQty = idx('Кол-во');
  var cAvg = idx('Средняя цена');
  var cPrice = idx('Текущая цена');
  var cInv = idx('Инвестировано');
  var cMkt = idx('Рыночная стоимость');
  var cPL = idx('P/L (руб)');
  var cPLPct = idx('P/L (%)');

  function d(from, to) { return from - to; }
  function r2(expr) { return 'ROUND(' + expr + SEP + '2)'; }

  sh.getRange(startRow, cInv, numRows, 1).setFormulaR1C1(
    '=IF(OR(LEN(RC[' + d(cQty, cInv) + '])=0' + SEP + 'LEN(RC[' + d(cAvg, cInv) + '])=0)' + SEP +
    '""' + SEP +
    r2('RC[' + d(cQty, cInv) + ']*RC[' + d(cAvg, cInv) + ']') + ')'
  );

  sh.getRange(startRow, cMkt, numRows, 1).setFormulaR1C1(
    '=IF(OR(LEN(RC[' + d(cQty, cMkt) + '])=0' + SEP + 'LEN(RC[' + d(cPrice, cMkt) + '])=0)' + SEP +
    '""' + SEP +
    r2('RC[' + d(cQty, cMkt) + ']*RC[' + d(cPrice, cMkt) + ']') + ')'
  );

  sh.getRange(startRow, cPL, numRows, 1).setFormulaR1C1(
    '=IF(OR(LEN(RC[' + d(cMkt, cPL) + '])=0' + SEP + 'LEN(RC[' + d(cInv, cPL) + '])=0)' + SEP +
    '""' + SEP +
    r2('RC[' + d(cMkt, cPL) + ']-RC[' + d(cInv, cPL) + ']') + ')'
  );

  sh.getRange(startRow, cPLPct, numRows, 1).setFormulaR1C1(
    '=IFERROR(RC[' + d(cPL, cPLPct) + ']/RC[' + d(cInv, cPLPct) + ']' + SEP + '0)'
  );
}

// ===== Форматирование =====

function formatSharesSheet_(sh) {
  sh.setFrozenRows(1);

  var lastRow = sh.getLastRow();
  var totalCols = SHARES_HEADERS.length;
  if (totalCols > 0) {
    try { sh.autoResizeColumns(1, totalCols); } catch (e) {}
  }
  if (lastRow < 2) return;

  var idx = function(name) { return SHARES_HEADERS.indexOf(name) + 1; };

  sh.getRange(2, idx('Кол-во'), lastRow - 1, 1).setNumberFormat('0.####');

  sh.getRange(2, idx('Средняя цена'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('Текущая цена'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('Инвестировано'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('Рыночная стоимость'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('P/L (руб)'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('P/L (%)'), lastRow - 1, 1).setNumberFormat('0.00%');
  sh.getRange(2, idx('Время цены'), lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  sh.getRange(2, idx('Капитализация'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('Выручка TTM'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('EBITDA TTM'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('Чистая прибыль TTM'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('EPS TTM'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('P/E TTM'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('P/S TTM'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('P/B TTM'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('EV/EBITDA'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('Debt/Equity'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('NetDebt/EBITDA'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('Free Float'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('Beta'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('Shares Outstanding'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('52w High'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('52w Low'), lastRow - 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, idx('ROE'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('ROA'), lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, idx('ROIC'), lastRow - 1, 1).setNumberFormat('0.00');
}

// ===== Helpers: API → данные =====

function fetchSharesInfo_(figis) {
  var out = {};

  (figis || []).forEach(function(figi) {
    try {
      var inst = callInstrumentsShareByFigi_(figi);
      if (!inst) return;

      out[figi] = {
        ticker: inst.ticker || '',
        name: inst.name || '',
        sector: inst.sector || '',
        countryOfRisk: inst.countryOfRisk || inst.countryOfRiskName || '',
        currency: inst.currency || '',
        exchange: inst.exchange || inst.realExchange || inst.classCode || '',
        classCode: inst.classCode || '',
        shareType: inst.shareType || '',
        shortEnabled: (inst.shortEnabledFlag === true || inst.shortEnabled === true) ? true :
                      ((inst.shortEnabledFlag === false || inst.shortEnabled === false) ? false : null),
        blockedTcs: (inst.blockedTcaFlag === true || inst.blockedTcs === true) ? true :
                    ((inst.blockedTcaFlag === false || inst.blockedTcs === false) ? false : null),
        assetUid: inst.assetUid || inst.asset_uid || ''
      };
    } catch (e) {}
  });

  return out;
}

function fetchSharesMarketData_(figis) {
  var out = {};
  var batches = splitIntoBatches_(figis || [], 200);

  batches.forEach(function(batch) {
    try {
      var last = callMarketLastPrices_(batch) || [];
      last.forEach(function(x) {
        if (!x || !x.figi) return;
        out[x.figi] = {
          lastPrice: normalizeNum_(firstNonNull_(x.lastPrice, x.price)),
          lastTime: normalizeTimeValue_(firstNonNull_(x.time, x.lastTime))
        };
      });
    } catch (e) {}
  });

  return out;
}

function fetchSharesFundamentalsByFigi_(figis) {
  var figiToAssetUid = {};
  var assetUids = [];
  var seen = {};

  (figis || []).forEach(function(figi) {
    try {
      var inst = callInstrumentsShareByFigi_(figi);
      if (!inst) return;

      var assetUid = inst.assetUid || inst.asset_uid || '';
      if (!assetUid) return;

      figiToAssetUid[figi] = assetUid;
      if (!seen[assetUid]) {
        seen[assetUid] = true;
        assetUids.push(assetUid);
      }
    } catch (e) {}
  });

  var byUid = fetchSharesFundamentalsByAssetUid_(assetUids);
  var out = {};

  Object.keys(figiToAssetUid).forEach(function(figi) {
    out[figi] = byUid[figiToAssetUid[figi]] || {};
  });

  return out;
}

function fetchSharesFundamentalsByAssetUid_(assetUids) {
  var out = {};
  if (!assetUids || !assetUids.length) return out;
  if (typeof callInstrumentsGetAssetFundamentals_ !== 'function') return out;

  var batches = splitIntoBatches_(assetUids, 100);
  batches.forEach(function(batch) {
    try {
      var raw = callInstrumentsGetAssetFundamentals_(batch) || {};

      if (Array.isArray(raw)) {
        raw.forEach(function(item) {
          var assetUid = item && (item.assetUid || item.asset_uid);
          if (!assetUid) return;
          out[assetUid] = normalizeShareFundamentals_(item);
        });
      } else {
        Object.keys(raw).forEach(function(assetUid) {
          out[assetUid] = normalizeShareFundamentals_(raw[assetUid]);
        });
      }
    } catch (e) {}
  });

  return out;
}

// ===== Portfolio helpers =====

function fetchSharesPortfolioMap_() {
  if (typeof safeFetchAllPortfolios_ === 'function') {
    try {
      return safeFetchAllPortfolios_() || {};
    } catch (e) {}
  }

  var map = {};
  if (typeof callUsersGetAccounts_ !== 'function') return map;
  if (typeof callPortfolioGetPortfolio_ !== 'function') return map;
  if (typeof callPortfolioGetPositions_ !== 'function') return map;

  try {
    var accounts = callUsersGetAccounts_() || [];
    accounts.forEach(function(acc) {
      var byPortfolio = [];
      var byPositions = [];

      try { byPortfolio = callPortfolioGetPortfolio_(acc.accountId) || []; } catch (e1) {}
      try { byPositions = callPortfolioGetPositions_(acc.accountId) || []; } catch (e2) {}

      var combined = byPortfolio.concat(byPositions);
      combined.forEach(function(p) {
        if (!p || !p.figi) return;

        var qty = normalizeNum_(firstNonNull_(p.qty, p.quantity));
        var avg = normalizeNum_(firstNonNull_(p.avg, p.avg_fifo));

        map[p.figi] = map[p.figi] || [];
        var exist = null;

        for (var i = 0; i < map[p.figi].length; i++) {
          if (map[p.figi][i].accountId === acc.accountId) {
            exist = map[p.figi][i];
            break;
          }
        }

        if (exist) {
          if (exist.qty == null && qty != null) exist.qty = qty;
          if (exist.avg == null && avg != null) exist.avg = avg;
          if (exist.avg_fifo == null && p.avg_fifo != null) exist.avg_fifo = normalizeNum_(p.avg_fifo);
        } else {
          map[p.figi].push({
            accountId: acc.accountId,
            accountName: acc.name || '',
            qty: qty,
            avg: normalizeNum_(p.avg),
            avg_fifo: normalizeNum_(p.avg_fifo)
          });
        }
      });
    });
  } catch (e) {
    return {};
  }

  return map;
}

function summarizeSharePosition_(positions) {
  var qtySum = 0;
  var costSum = 0;
  var hasQty = false;
  var hasWeighted = false;
  var fallbackAvg = null;

  (positions || []).forEach(function(p) {
    if (!p) return;

    var qty = normalizeNum_(firstNonNull_(p.qty, p.quantity));
    var avg = normalizeNum_(firstNonNull_(p.avg, p.avg_fifo));

    if (fallbackAvg == null && avg != null) fallbackAvg = avg;

    if (qty != null) {
      qtySum += qty;
      hasQty = true;
    }

    if (qty != null && avg != null) {
      costSum += qty * avg;
      hasWeighted = true;
    }
  });

  return {
    qty: hasQty ? qtySum : null,
    avg: (hasWeighted && qtySum) ? (costSum / qtySum) : fallbackAvg
  };
}

// ===== Нормализация =====

function normalizeShareFundamentals_(src) {
  src = src || {};
  return {
    assetUid: firstNonEmpty_(src.assetUid, src.asset_uid, ''),
    marketCap: normalizeNum_(firstNonNull_(src.marketCap, src.market_cap)),
    revenueTtm: normalizeNum_(firstNonNull_(src.revenueTtm, src.revenue_ttm)),
    ebitdaTtm: normalizeNum_(firstNonNull_(src.ebitdaTtm, src.ebitda_ttm)),
    netIncomeTtm: normalizeNum_(firstNonNull_(src.netIncomeTtm, src.net_income_ttm)),
    epsTtm: normalizeNum_(firstNonNull_(src.epsTtm, src.eps_ttm)),
    peRatioTtm: normalizeNum_(firstNonNull_(src.peRatioTtm, src.pe_ratio_ttm)),
    priceToSalesTtm: normalizeNum_(firstNonNull_(src.priceToSalesTtm, src.price_to_sales_ttm)),
    priceToBookTtm: normalizeNum_(firstNonNull_(src.priceToBookTtm, src.price_to_book_ttm)),
    evToEbitda: normalizeNum_(firstNonNull_(src.evToEbitda, src.ev_to_ebitda)),
    debtToEquity: normalizeNum_(firstNonNull_(src.debtToEquity, src.debt_to_equity)),
    netDebtToEbitda: normalizeNum_(firstNonNull_(src.netDebtToEbitda, src.net_debt_to_ebitda)),
    freeFloat: normalizeNum_(firstNonNull_(src.freeFloat, src.free_float)),
    beta: normalizeNum_(firstNonNull_(src.beta)),
    sharesOutstanding: normalizeNum_(firstNonNull_(src.sharesOutstanding, src.shares_outstanding)),
    high52w: normalizeNum_(firstNonNull_(src.high52w, src.high_52w)),
    low52w: normalizeNum_(firstNonNull_(src.low52w, src.low_52w)),
    roe: normalizeNum_(firstNonNull_(src.roe, src.roeTtm, src.roe_ttm)),
    roa: normalizeNum_(firstNonNull_(src.roa, src.roaTtm, src.roa_ttm)),
    roic: normalizeNum_(firstNonNull_(src.roic, src.roicTtm, src.roic_ttm))
  };
}

function normalizeTimeValue_(v) {
  if (v == null || v === '') return '';
  if (typeof tsToIso === 'function') {
    var iso = tsToIso(v);
    if (iso) return iso;
  }
  if (Object.prototype.toString.call(v) === '[object Date]') {
    try { return v.toISOString(); } catch (e) {}
  }
  return String(v);
}

function splitIntoBatches_(arr, batchSize) {
  var out = [];
  arr = arr || [];
  batchSize = Math.max(1, Number(batchSize) || 100);

  for (var i = 0; i < arr.length; i += batchSize) {
    out.push(arr.slice(i, i + batchSize));
  }
  return out;
}

function firstNonNull_() {
  for (var i = 0; i < arguments.length; i++) {
    if (arguments[i] != null) return arguments[i];
  }
  return null;
}

function firstNonEmpty_() {
  for (var i = 0; i < arguments.length; i++) {
    var v = arguments[i];
    if (v != null && String(v) !== '') return v;
  }
  return '';
}

function boolToRu_(v) {
  if (v === true) return 'Да';
  if (v === false) return 'Нет';
  return '';
}

function normalizeNum_(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;

  if (typeof moneyToNumber === 'function') {
    var m = moneyToNumber(v);
    if (m != null && !isNaN(m)) return Number(m);
  }

  if (typeof qToNumber === 'function') {
    var q = qToNumber(v);
    if (q != null && !isNaN(q)) return Number(q);
  }

  var n = Number(v);
  return isNaN(n) ? null : n;
}
