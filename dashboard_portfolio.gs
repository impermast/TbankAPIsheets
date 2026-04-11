/**
 * dashboard_portfolio.gs
 *
 * Единый владелец portfolio dashboard.
 *
 * Что строит:
 * - общий portfolio dashboard на листе Dashboard
 *
 * Какие листы читает:
 * - Bonds
 * - Shares
 * - Funds
 * - Options (только для общих totals по классам активов)
 *
 * Какие регионы пишет:
 * - general: общие итоги портфеля
 * - shares: акции
 * - bonds: облигации
 * - funds: фонды
 *
 * Публичные функции:
 * - buildPortfolioDashboard()
 * - renderGeneralDashboardRegion_()
 * - renderSharesDashboardRegion_()
 * - renderBondsDashboardRegion_()
 * - renderFundsDashboardRegion_()
 */


/* ========================================================================
 * Layout
 * ======================================================================== */

/**
 * Карта областей Dashboard.
 *
 * Сетка:
 * - general = A:J   (1..10)
 * - gap     = K
 * - shares  = L:U   (12..21)
 * - gap     = V
 * - bonds   = W:AF  (23..32)
 * - gap     = AG
 * - funds   = AH:AQ (34..43)
 *
 * По одной пустой разделительной колонке между регионами.
 */
var PORTFOLIO_DASHBOARD_LAYOUT = {
  sheetName: 'Dashboard',
  bounds: {
    maxRows: 115,
    maxCols: 43
  },
  history: {
    // legacy-compatible block for dashboard_export.gs
    r1: 12,
    r2: 28,
    c1: 1,
    w: 3
  },
  regions: {
    general: {
      key: 'general',
      title: 'Портфель',
      startCol: 1,
      width: 10,
      startRow: 1,
      maxRows: 115
    },
    shares: {
      key: 'shares',
      title: 'Акции',
      startCol: 12,
      width: 10,
      startRow: 1,
      maxRows: 105
    },
    bonds: {
      key: 'bonds',
      title: 'Облигации',
      startCol: 23,
      width: 10,
      startRow: 1,
      maxRows: 115
    },
    funds: {
      key: 'funds',
      title: 'Фонды',
      startCol: 34,
      width: 10,
      startRow: 1,
      maxRows: 80
    }
  }
};

function dashboardGetLayout_() {
  return PORTFOLIO_DASHBOARD_LAYOUT;
}


/* ========================================================================
 * Public orchestration
 * ======================================================================== */

function buildPortfolioDashboard() {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    if (typeof showSnack_ === 'function') {
      showSnack_('Другая операция обновляет Dashboard. Повторите позже.', 'Dashboard', 3000);
    }
    return;
  }

  try {
    var layout = dashboardGetLayout_();
    var sh = dashboardEnsureSheet_(layout.sheetName);
    var historySnapshot = portfolioHistorySnapshot_(sh, layout.history);

    dashboardResetSheet_(sh, layout);

    renderGeneralDashboardRegion_(sh, layout.regions.general, {
      layout: layout,
      historySnapshot: historySnapshot
    });
    renderSharesDashboardRegion_(sh, layout.regions.shares);
    renderBondsDashboardRegion_(sh, layout.regions.bonds);
    renderFundsDashboardRegion_(sh, layout.regions.funds);

    dashboardAutoResizeRegions_(sh, layout);
    SpreadsheetApp.flush();

    if (typeof showSnack_ === 'function') {
      showSnack_('Dashboard обновлён', 'Dashboard', 2000);
    }
  } finally {
    lock.releaseLock();
  }
}


/* ========================================================================
 * General region
 * ======================================================================== */

function renderGeneralDashboardRegion_(sh, region, ctx) {
  var layout = (ctx && ctx.layout) || dashboardGetLayout_();
  region = dashboardNormalizeRegion_(region, layout.regions.general);

  dashboardClearRegion_(sh, region);
  dashboardWriteRegionTitle_(sh, region, region.title || 'Портфель');

  var data = _buildGeneralPortfolioData_();
  var bondCompare = _buildGeneralBondCompareData_();

  dashboardWriteTwoColBlock_(
    sh,
    region.startRow + 2,
    region.startCol,
    'Портфель — итого',
    [
      ['Инвестировано', data.totals.invested],
      ['Рыночная стоимость', data.totals.market],
      ['P/L (руб)', data.totals.plRub],
      ['P/L (%)', data.totals.plPct]
    ],
    ['#,##0.00', '#,##0.00', '#,##0.00', '0.00%']
  );

  dashboardWriteTableBlock_(
    sh,
    region.startRow + 2,
    region.startCol + 3,
    'Классы активов',
    ['Класс', 'Бумаг', 'Инвестировано', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)', 'Доля портфеля'],
    data.classRows,
    {
      formatsByCol: {
        3: '#,##0.00',
        4: '#,##0.00',
        5: '#,##0.00',
        6: '0.00%',
        7: '0.00%'
      }
    }
  );

  var historyBlock = portfolioUpdateHistory_(
    ctx && ctx.historySnapshot ? ctx.historySnapshot : null,
    data.totals.invested,
    data.totals.market,
    layout.history
  );
  portfolioWriteHistoryBlock_(sh, layout.history, historyBlock);

  if (bondCompare.hasData) {
    var cmpInfo = dashboardWriteTableBlock_(
      sh,
      12,
      4,
      'YTM vs купон',
      ['Метрика', 'Значение'],
      [
        ['Средневзв. YTM (%)', bondCompare.weightedYtmPct],
        ['Купонная доходность (%)', bondCompare.weightedCouponPct]
      ],
      {
        formatsByCol: {
          2: '0.00'
        }
      }
    );

    if (cmpInfo && cmpInfo.dataRowsCount > 0) {
      dashboardBuildChartFromRanges_(sh, {
        chartType: Charts.ChartType.COLUMN,
        title: 'YTM vs Купонная доходность (средневзв., %)',
        anchorRow: 50,
        anchorCol: 4,
        ranges: [cmpInfo.tableRange],
        width: 380,
        height: 220,
        options: {
          legend: { position: 'none' }
        }
      });
    }

    if (bondCompare.scatterRows.length > 1) {
      var scatterInfo = dashboardWriteTableBlock_(
        sh,
        12,
        7,
        'Риск vs YTM',
        ['Риск', 'YTM (%)', 'Тултип'],
        bondCompare.scatterRows.slice(1),
        {
          formatsByCol: {
            1: '0.00',
            2: '0.00'
          }
        }
      );

      if (scatterInfo && scatterInfo.dataRowsCount > 0) {
        dashboardBuildChartFromRanges_(sh, {
          chartType: Charts.ChartType.SCATTER,
          title: 'Риск vs Доходность к погашению (YTM)',
          anchorRow: 72,
          anchorCol: 4,
          ranges: [scatterInfo.tableRange],
          width: 380,
          height: 220,
          options: {
            legend: { position: 'none' },
            hAxis: { title: 'Риск (баллы)' },
            vAxis: { title: 'YTM (%)' },
            series: { 0: { pointSize: 5 } }
          }
        });
      }
    }
  }

  var historyRange = sh.getRange(
    layout.history.r1,
    layout.history.c1,
    layout.history.r2 - layout.history.r1 + 1,
    layout.history.w
  );

  if (_dashboardChartRangeHasData_(historyRange)) {
    dashboardBuildChartFromRanges_(sh, {
      chartType: Charts.ChartType.LINE,
      title: 'История портфеля: Инвестировано vs Стоимость',
      anchorRow: 30,
      anchorCol: 1,
      ranges: [historyRange],
      width: 400,
      height: 220,
      options: {
        legend: { position: 'bottom' },
        hAxis: { title: 'Дата', format: 'dd.MM' },
        vAxis: { title: '₽' },
        pointSize: 5,
        series: {
          0: { labelInLegend: 'Инвестировано', pointSize: 5 },
          1: { labelInLegend: 'Стоимость', pointSize: 5 }
        }
      }
    });
  }
}

function _buildGeneralPortfolioData_() {
  var byClass = {
    bonds: _readSheetTotalsByHeaders_('Bonds'),
    funds: _readSheetTotalsByHeaders_('Funds'),
    shares: _readSheetTotalsByHeaders_('Shares'),
    options: _readSheetTotalsByHeaders_('Options')
  };

  var totals = {
    invested:
      (byClass.bonds.invested || 0) +
      (byClass.funds.invested || 0) +
      (byClass.shares.invested || 0) +
      (byClass.options.invested || 0),

    market:
      (byClass.bonds.market || 0) +
      (byClass.funds.market || 0) +
      (byClass.shares.market || 0) +
      (byClass.options.market || 0),

    plRub:
      (byClass.bonds.plRub || 0) +
      (byClass.funds.plRub || 0) +
      (byClass.shares.plRub || 0) +
      (byClass.options.plRub || 0),

    count:
      (byClass.bonds.count || 0) +
      (byClass.funds.count || 0) +
      (byClass.shares.count || 0) +
      (byClass.options.count || 0),

    plPct: null
  };

  totals.plPct = totals.invested ? totals.plRub / totals.invested : null;

  return {
    byClass: byClass,
    totals: totals,
    classRows: [
      _assetClassRow_('Облигации', byClass.bonds, totals.market),
      _assetClassRow_('Фонды', byClass.funds, totals.market),
      _assetClassRow_('Акции', byClass.shares, totals.market),
      _assetClassRow_('Опционы', byClass.options, totals.market)
    ]
  };
}

function _assetClassRow_(title, totals, totalMarket) {
  var t = totals || {};
  var invested = _toNumberSafe_(t.invested);
  var market = _toNumberSafe_(t.market);
  var plRub = _toNumberSafe_(t.plRub);
  var plPct = t.plPct != null ? t.plPct : (invested ? plRub / invested : null);
  var share = totalMarket ? market / totalMarket : null;

  return [
    title,
    _toNumberSafe_(t.count),
    invested,
    market,
    plRub,
    plPct,
    share
  ];
}

function portfolioHistorySnapshot_(sh, historyCfg) {
  try {
    if (!sh) return null;
    if (sh.getLastRow() < historyCfg.r1) return null;
    return sh.getRange(
      historyCfg.r1,
      historyCfg.c1,
      historyCfg.r2 - historyCfg.r1 + 1,
      historyCfg.w
    ).getValues();
  } catch (e) {
    return null;
  }
}

function portfolioUpdateHistory_(snapshot, investedRub, marketRub, historyCfg) {
  var entries = [];
  var i;
  var d;
  var inv;
  var mkt;
  var N = historyCfg.r2 - historyCfg.r1; // строк данных без header
  var now = new Date();

  if (snapshot && snapshot.length) {
    for (i = 1; i < snapshot.length; i++) {
      d = snapshot[i][0];
      inv = snapshot[i][1];
      mkt = snapshot[i][2];

      if (d === '' || d == null) continue;

      d = (d instanceof Date) ? d : new Date(d);
      if (!isFinite(d.getTime())) continue;

      inv = Number(inv);
      mkt = Number(mkt);

      entries.push({
        dt: d,
        invested: isFinite(inv) ? inv : '',
        market: isFinite(mkt) ? mkt : ''
      });
    }
  }

  entries.push({
    dt: now,
    invested: Number(investedRub) || 0,
    market: Number(marketRub) || 0
  });

  if (entries.length > N) {
    entries = entries.slice(entries.length - N);
  }

  while (entries.length < N) {
    entries.unshift(null);
  }

  var out = [['Дата', 'Инвестировано', 'Стоимость']];
  for (i = 0; i < entries.length; i++) {
    if (!entries[i]) {
      out.push(['', '', '']);
    } else {
      out.push([
        entries[i].dt,
        entries[i].invested,
        entries[i].market
      ]);
    }
  }

  return out;
}

function portfolioWriteHistoryBlock_(sh, historyCfg, block) {
  sh.getRange(
    historyCfg.r1,
    historyCfg.c1,
    historyCfg.r2 - historyCfg.r1 + 1,
    historyCfg.w
  ).setValues(block);

  sh.getRange(historyCfg.r1, historyCfg.c1, 1, historyCfg.w).setFontWeight('bold');
  sh.getRange(historyCfg.r1 + 1, historyCfg.c1, historyCfg.r2 - historyCfg.r1, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange(historyCfg.r1 + 1, historyCfg.c1 + 1, historyCfg.r2 - historyCfg.r1, 2)
    .setNumberFormat('#,##0.00');
}


/* ========================================================================
 * Shares: data
 * ======================================================================== */

function loadSharesDashboardData_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Shares');

  if (!sh) return { header: [], rows: [] };
  if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) return { header: [], rows: [] };

  return {
    header: sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0],
    rows: sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
  };
}

function mapSharesDashboardColumns_(hdr) {
  function idx(name) {
    var i = hdr.indexOf(name);
    return i >= 0 ? (i + 1) : 0;
  }

  return {
    name: idx('Название'),
    ticker: idx('Тикер'),
    market: idx('Рыночная стоимость'),
    invested: idx('Инвестировано'),
    plRub: idx('P/L (руб)'),
    plPct: idx('P/L (%)'),
    sector: idx('Сектор'),
    country: idx('Страна риска'),
    pe: idx('P/E TTM'),
    pb: idx('P/B TTM'),
    evEbitda: idx('EV/EBITDA'),
    roe: idx('ROE'),
    beta: idx('Beta')
  };
}

function validateSharesDashboardColumns_(c) {
  return !!(
    c &&
    c.name &&
    c.ticker &&
    c.market &&
    c.invested &&
    c.plRub &&
    c.plPct &&
    c.sector &&
    c.country &&
    c.pe &&
    c.pb &&
    c.evEbitda &&
    c.roe &&
    c.beta
  );
}

function computeSharesDashboardData_(rows, c) {
  function val(r, ci) {
    return ci ? r[ci - 1] : '';
  }

  function toNumber_(x) {
    if (x === null || x === '' || typeof x === 'undefined') return NaN;
    if (typeof x === 'number') return isFinite(x) ? x : NaN;
    if (x instanceof Date) return NaN;

    var s = String(x)
      .replace(/\u00A0/g, ' ')
      .replace(/\s+/g, '')
      .replace(/%/g, '')
      .replace(',', '.')
      .trim();

    if (!s) return NaN;

    var n = Number(s);
    return isFinite(n) ? n : NaN;
  }

  function round2_(x) {
    return Math.round(Number(x || 0) * 100) / 100;
  }

  function round4_(x) {
    return Math.round(Number(x || 0) * 10000) / 10000;
  }

  function safeName_(x) {
    var s = String(x == null ? '' : x).trim();
    return s || '—';
  }

  function pushAgg_(map, name, market) {
    map[name] = (map[name] || 0) + market;
  }

  var invested = 0;
  var market = 0;
  var plRub = 0;
  var count = 0;

  var items = [];
  var sectorsMap = {};
  var countriesMap = {};

  var peArr = [];
  var pbArr = [];
  var evEbitdaArr = [];
  var roeArr = [];
  var betaArr = [];

  rows.forEach(function (r) {
    var name = safeName_(val(r, c.name));
    var ticker = safeName_(val(r, c.ticker));

    var marketVal = toNumber_(val(r, c.market));
    var investedVal = toNumber_(val(r, c.invested));
    var plRubVal = toNumber_(val(r, c.plRub));
    var plPctVal = toNumber_(val(r, c.plPct));

    var sector = safeName_(val(r, c.sector));
    var country = safeName_(val(r, c.country));

    var peVal = toNumber_(val(r, c.pe));
    var pbVal = toNumber_(val(r, c.pb));
    var evEbitdaVal = toNumber_(val(r, c.evEbitda));
    var roeVal = toNumber_(val(r, c.roe));
    var betaVal = toNumber_(val(r, c.beta));

    var hasAnyData =
      (name !== '—') ||
      (ticker !== '—') ||
      isFinite(marketVal) ||
      isFinite(investedVal) ||
      isFinite(plRubVal) ||
      isFinite(plPctVal);

    if (!hasAnyData) return;

    count += 1;

    if (isFinite(investedVal)) invested += investedVal;
    if (isFinite(marketVal)) market += marketVal;
    if (isFinite(plRubVal)) plRub += plRubVal;

    pushAgg_(sectorsMap, sector || '—', isFinite(marketVal) ? marketVal : 0);
    pushAgg_(countriesMap, country || '—', isFinite(marketVal) ? marketVal : 0);

    if (isFinite(peVal)) peArr.push(peVal);
    if (isFinite(pbVal)) pbArr.push(pbVal);
    if (isFinite(evEbitdaVal)) evEbitdaArr.push(evEbitdaVal);
    if (isFinite(roeVal)) roeArr.push(roeVal);
    if (isFinite(betaVal)) betaArr.push(betaVal);

    items.push({
      name: name,
      ticker: ticker,
      market: isFinite(marketVal) ? marketVal : 0,
      plRub: isFinite(plRubVal) ? plRubVal : 0,
      plPct: isFinite(plPctVal) ? plPctVal : null
    });
  });

  var totalMarket = market;
  var plPct = invested > 0 ? (plRub / invested) : null;

  var topPositions = _sortDescByField_(items, 'market').slice(0, 5);

  var plItems = items.filter(function (x) {
    return isFinite(x.plRub);
  });

  var bestPL = _sortDescByField_(plItems, 'plRub').slice(0, 5);
  var worstPL = _sortAscByField_(plItems, 'plRub').slice(0, 5);

  var sectors = Object.keys(sectorsMap).map(function (k) {
    var m = sectorsMap[k] || 0;
    return {
      name: k || '—',
      market: round2_(m),
      sharePct: totalMarket > 0 ? round4_(m / totalMarket) : 0
    };
  });
  sectors = _sortDescByField_(sectors, 'market');

  var countries = Object.keys(countriesMap).map(function (k) {
    var m = countriesMap[k] || 0;
    return {
      name: k || '—',
      market: round2_(m),
      sharePct: totalMarket > 0 ? round4_(m / totalMarket) : 0
    };
  });
  countries = _sortDescByField_(countries, 'market');

  topPositions = topPositions.map(function (x) {
    return {
      name: x.name,
      ticker: x.ticker,
      market: round2_(x.market),
      plRub: round2_(x.plRub),
      plPct: x.plPct === null ? null : round4_(x.plPct)
    };
  });

  bestPL = bestPL.map(function (x) {
    return {
      name: x.name,
      ticker: x.ticker,
      market: round2_(x.market),
      plRub: round2_(x.plRub),
      plPct: x.plPct === null ? null : round4_(x.plPct)
    };
  });

  worstPL = worstPL.map(function (x) {
    return {
      name: x.name,
      ticker: x.ticker,
      market: round2_(x.market),
      plRub: round2_(x.plRub),
      plPct: x.plPct === null ? null : round4_(x.plPct)
    };
  });

  return {
    count: count,
    invested: round2_(invested),
    market: round2_(market),
    plRub: round2_(plRub),
    plPct: plPct === null ? null : round4_(plPct),
    topPositions: topPositions,
    bestPL: bestPL,
    worstPL: worstPL,
    sectors: sectors,
    countries: countries,
    valuation: {
      peMedian: _medianNumber_(peArr),
      pbMedian: _medianNumber_(pbArr),
      evEbitdaMedian: _medianNumber_(evEbitdaArr),
      roeMedian: _medianNumber_(roeArr),
      betaMedian: _medianNumber_(betaArr)
    }
  };
}

function _readSharesDashboardViewModel_() {
  var raw = loadSharesDashboardData_();
  var header = (raw && raw.header) || [];
  var rows = (raw && raw.rows) || [];

  if (!header.length || !rows.length) return null;

  var c = mapSharesDashboardColumns_(header);
  if (validateSharesDashboardColumns_(c)) {
    var data = computeSharesDashboardData_(rows, c);
    if (data && data.count) return data;
  }

  return _buildSharesFallbackData_(header, rows);
}

function _buildSharesFallbackData_(header, rows) {
  function round2_(x) {
    return Math.round(Number(x || 0) * 100) / 100;
  }

  function round4_(x) {
    return Math.round(Number(x || 0) * 10000) / 10000;
  }

  var c = _mapGenericSharesColumns_(header);
  var invested = 0;
  var market = 0;
  var plRub = 0;
  var count = 0;

  var items = [];
  var sectorsMap = {};
  var countriesMap = {};

  var peArr = [];
  var pbArr = [];
  var evEbitdaArr = [];
  var roeArr = [];
  var betaArr = [];

  rows.forEach(function (r) {
    if (!_rowHasAnyValue_(r)) return;

    var name = _safeTextByIndex_(r, c.name, '—');
    var ticker = _safeTextByIndex_(r, c.ticker, '—');
    var marketVal = _toNumberSafe_(_valueByIndex_(r, c.market));
    var investedVal = _toNumberSafe_(_valueByIndex_(r, c.invested));
    var plRubVal = _toNumberSafe_(_valueByIndex_(r, c.plRub));
    var plPctVal = _valueByIndex_(r, c.plPct);

    plPctVal = (plPctVal === '' || plPctVal == null) ? null : _toNumberSafe_(plPctVal);
    if (plPctVal == null) plPctVal = investedVal ? plRubVal / investedVal : null;

    count += 1;
    invested += investedVal;
    market += marketVal;
    plRub += plRubVal;

    items.push({
      name: name,
      ticker: ticker,
      market: marketVal,
      plRub: plRubVal,
      plPct: plPctVal
    });

    _pushMarketAgg_(sectorsMap, _safeTextByIndex_(r, c.sector, '—'), marketVal);
    _pushMarketAgg_(countriesMap, _safeTextByIndex_(r, c.country, '—'), marketVal);

    _pushNumberIfFinite_(peArr, _valueByIndex_(r, c.pe));
    _pushNumberIfFinite_(pbArr, _valueByIndex_(r, c.pb));
    _pushNumberIfFinite_(evEbitdaArr, _valueByIndex_(r, c.evEbitda));
    _pushNumberIfFinite_(roeArr, _valueByIndex_(r, c.roe));
    _pushNumberIfFinite_(betaArr, _valueByIndex_(r, c.beta));
  });

  var totalMarket = market;
  var sectors = _buildMarketShareRows_(sectorsMap, totalMarket);
  var countries = _buildMarketShareRows_(countriesMap, totalMarket);
  var topPositions = _sortDescByField_(items, 'market').slice(0, 5);
  var bestPL = _sortDescByField_(items, 'plRub').slice(0, 5);
  var worstPL = _sortAscByField_(items, 'plRub').slice(0, 5);

  function normalizeItems(arr) {
    return (arr || []).map(function (x) {
      return {
        name: x.name,
        ticker: x.ticker,
        market: round2_(x.market),
        plRub: round2_(x.plRub),
        plPct: x.plPct == null ? null : round4_(x.plPct)
      };
    });
  }

  return {
    count: count,
    invested: round2_(invested),
    market: round2_(market),
    plRub: round2_(plRub),
    plPct: invested ? round4_(plRub / invested) : null,
    topPositions: normalizeItems(topPositions),
    bestPL: normalizeItems(bestPL),
    worstPL: normalizeItems(worstPL),
    sectors: sectors,
    countries: countries,
    valuation: {
      peMedian: _medianNumber_(peArr),
      pbMedian: _medianNumber_(pbArr),
      evEbitdaMedian: _medianNumber_(evEbitdaArr),
      roeMedian: _medianNumber_(roeArr),
      betaMedian: _medianNumber_(betaArr)
    }
  };
}

function _mapGenericSharesColumns_(header) {
  return {
    name: _findHeaderIndexByAliases_(header, ['Название', 'Наименование', 'Инструмент', 'Компания']),
    ticker: _findHeaderIndexByAliases_(header, ['Тикер', 'Ticker']),
    invested: _findHeaderIndexByAliases_(header, ['Инвестировано']),
    market: _findHeaderIndexByAliases_(header, ['Рыночная стоимость', 'Стоимость', 'Рыночная стоимость, ₽']),
    plRub: _findHeaderIndexByAliases_(header, ['P/L (руб)', 'P/L, руб', 'P/L RUB']),
    plPct: _findHeaderIndexByAliases_(header, ['P/L (%)', 'P/L %']),
    sector: _findHeaderIndexByAliases_(header, ['Сектор']),
    country: _findHeaderIndexByAliases_(header, ['Страна', 'Country', 'Страна риска']),
    pe: _findHeaderIndexByAliases_(header, ['P/E', 'P/E (TTM)', 'PE', 'PE TTM']),
    pb: _findHeaderIndexByAliases_(header, ['P/B', 'PB', 'P/B TTM']),
    evEbitda: _findHeaderIndexByAliases_(header, ['EV/EBITDA', 'evToEbitda']),
    roe: _findHeaderIndexByAliases_(header, ['ROE', 'roe']),
    beta: _findHeaderIndexByAliases_(header, ['Beta', 'beta'])
  };
}

function _medianNumber_(arr) {
  if (!arr || !arr.length) return null;

  var nums = arr
    .map(function (x) { return typeof x === 'number' ? x : Number(x); })
    .filter(function (x) { return isFinite(x); })
    .sort(function (a, b) { return a - b; });

  if (!nums.length) return null;

  var mid = Math.floor(nums.length / 2);
  if (nums.length % 2) return nums[mid];
  return (nums[mid - 1] + nums[mid]) / 2;
}

function _sortDescByField_(arr, field) {
  return (arr || []).slice().sort(function (a, b) {
    var av = Number(a && a[field]);
    var bv = Number(b && b[field]);
    var aOk = isFinite(av);
    var bOk = isFinite(bv);

    if (!aOk && !bOk) return 0;
    if (!aOk) return 1;
    if (!bOk) return -1;
    return bv - av;
  });
}

function _sortAscByField_(arr, field) {
  return (arr || []).slice().sort(function (a, b) {
    var av = Number(a && a[field]);
    var bv = Number(b && b[field]);
    var aOk = isFinite(av);
    var bOk = isFinite(bv);

    if (!aOk && !bOk) return 0;
    if (!aOk) return 1;
    if (!bOk) return -1;
    return av - bv;
  });
}


/* ========================================================================
 * Shares: render
 * ======================================================================== */

function renderSharesDashboardRegion_(sh, region) {
  region = dashboardNormalizeRegion_(region, dashboardGetLayout_().regions.shares);

  dashboardClearRegion_(sh, region);
  dashboardWriteRegionTitle_(sh, region, region.title || 'Акции');

  var data = _readSharesDashboardViewModel_();
  if (!data || !data.count) {
    dashboardWritePlaceholder_(sh, region, 'Нет данных по акциям');
    return;
  }

  dashboardWriteTwoColBlock_(
    sh,
    _regionRow_(region, 3),
    region.startCol,
    'Акции — итого',
    [
      ['Инвестировано', data.invested],
      ['Рыночная стоимость', data.market],
      ['P/L (руб)', data.plRub],
      ['P/L (%)', data.plPct]
    ],
    ['#,##0.00', '#,##0.00', '#,##0.00', '0.00%']
  );

  var topPositionsRows = (data.topPositions || []).slice(0, 5).map(function (item) {
    return [item.name || '—', item.ticker || '—', item.market, item.plRub, item.plPct];
  });
  if (!topPositionsRows.length) topPositionsRows = [['Нет данных', '', '', '', '']];

  var topPositionsInfo = dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 10),
    region.startCol,
    'Акции — топ-5 по позиции',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    topPositionsRows,
    {
      formatsByCol: {
        3: '#,##0.00',
        4: '#,##0.00',
        5: '0.00%'
      }
    }
  );

  var bestPLRows = (data.bestPL || []).slice(0, 5).map(function (item) {
    return [item.name || '—', item.ticker || '—', item.market, item.plRub, item.plPct];
  });
  if (!bestPLRows.length) bestPLRows = [['Нет данных', '', '', '', '']];

  dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 19),
    region.startCol,
    'Акции — топ-5 по P/L',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    bestPLRows,
    {
      formatsByCol: {
        3: '#,##0.00',
        4: '#,##0.00',
        5: '0.00%'
      }
    }
  );

  var worstPLRows = (data.worstPL || []).slice(0, 5).map(function (item) {
    return [item.name || '—', item.ticker || '—', item.market, item.plRub, item.plPct];
  });
  if (!worstPLRows.length) worstPLRows = [['Нет данных', '', '', '', '']];

  dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 28),
    region.startCol,
    'Акции — топ-5 убыток',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    worstPLRows,
    {
      formatsByCol: {
        3: '#,##0.00',
        4: '#,##0.00',
        5: '0.00%'
      }
    }
  );

  var sectorsRows = (data.sectors || []).slice(0, 5).map(function (item) {
    return [item.name || '—', item.market, item.sharePct];
  });
  var hasSectorsRows = sectorsRows.length > 0;
  if (!sectorsRows.length) sectorsRows = [['Нет данных', '', '']];

  var sectorsInfo = dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 37),
    region.startCol,
    'Акции — по секторам',
    ['Сектор', 'Рыночная стоимость', 'Доля акций'],
    sectorsRows,
    {
      formatsByCol: {
        2: '#,##0.00',
        3: '0.00%'
      }
    }
  );

  var countriesRows = (data.countries || []).slice(0, 5).map(function (item) {
    return [item.name || '—', item.market, item.sharePct];
  });
  if (!countriesRows.length) countriesRows = [['Нет данных', '', '']];

  dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 47),
    region.startCol,
    'Акции — по странам',
    ['Страна', 'Рыночная стоимость', 'Доля акций'],
    countriesRows,
    {
      formatsByCol: {
        2: '#,##0.00',
        3: '0.00%'
      }
    }
  );

  dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 57),
    region.startCol,
    'Акции — valuation snapshot',
    ['Метрика', 'Значение'],
    _sharesValuationRowsFromData_(data),
    {}
  );

  if (hasSectorsRows) {
    dashboardBuildChartFromRanges_(sh, {
      chartType: Charts.ChartType.PIE,
      title: 'Акции — структура по секторам',
      anchorRow: _regionRow_(region, 66),
      anchorCol: region.startCol,
      ranges: [
        sh.getRange(
          sectorsInfo.headerRow,
          region.startCol,
          1 + sectorsInfo.dataRowsCount,
          2
        )
      ],
      width: 400,
      height: 220,
      options: {
        legend: { position: 'right' },
        pieSliceText: 'value'
      }
    });
  }

  if ((data.topPositions || []).length) {
    dashboardBuildChartFromRanges_(sh, {
      chartType: Charts.ChartType.BAR,
      title: 'Акции — топ-5 по позиции',
      anchorRow: _regionRow_(region, 86),
      anchorCol: region.startCol,
      ranges: [
        sh.getRange(topPositionsInfo.headerRow, region.startCol, 1 + topPositionsInfo.dataRowsCount, 1),
        sh.getRange(topPositionsInfo.headerRow, region.startCol + 2, 1 + topPositionsInfo.dataRowsCount, 1)
      ],
      width: 400,
      height: 220,
      options: {
        legend: { position: 'none' },
        hAxis: { format: '#,##0.00' },
        bars: 'horizontal'
      }
    });
  }
}

function _sharesValuationRowsFromData_(data) {
  var v = (data && data.valuation) || {};

  return [
    ['Median P/E', v.peMedian != null ? v.peMedian : ''],
    ['Median P/B', v.pbMedian != null ? v.pbMedian : ''],
    ['Median EV/EBITDA', v.evEbitdaMedian != null ? v.evEbitdaMedian : ''],
    ['Median ROE', v.roeMedian != null ? v.roeMedian : ''],
    ['Median Beta', v.betaMedian != null ? v.betaMedian : '']
  ];
}


/* ========================================================================
 * Bonds: data + calculations
 * ======================================================================== */

function _loadBondsRegionData_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Bonds');

  if (!sh) return { header: [], rows: [] };
  if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) return { header: [], rows: [] };

  return {
    header: sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0],
    rows: sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
  };
}

function _mapBondsRegionColumns_(hdr) {
  function idx() {
    var i;
    for (i = 0; i < arguments.length; i++) {
      var k = hdr.indexOf(arguments[i]);
      if (k >= 0) return k + 1;
    }
    return 0;
  }

  return {
    name: idx('Название'),
    figi: idx('FIGI'),
    riskNum: idx('Риск (ручн.)'),
    sector: idx('Сектор'),
    qty: idx('Кол-во'),
    price: idx('Текущая цена'),
    nominal: idx('Номинал'),
    couponPerYear: idx('купон/год'),
    couponValue: idx('Размер купона'),
    couponPct: idx('Купон, %'),
    maturity: idx('Дата погашения'),
    nextCoupon: idx('Следующий купон'),
    market: idx('Рыночная стоимость'),
    invested: idx('Инвестировано'),
    plRub: idx('P/L (руб)'),
    plPct: idx('P/L (%)')
  };
}

function _computeBondsTotals_(rows, c) {
  var invested = 0;
  var market = 0;
  var plRub = 0;

  rows.forEach(function (r) {
    var inv = _bondNumber_(r, c.invested);
    var mkt = _bondNumber_(r, c.market);
    var pl = _bondNumber_(r, c.plRub);

    if (isFinite(inv)) invested += inv;
    if (isFinite(mkt)) market += mkt;
    if (isFinite(pl)) plRub += pl;
  });

  return {
    invested: _bondRound2_(invested),
    market: _bondRound2_(market),
    plRub: _bondRound2_(plRub),
    plPct: invested > 0 ? _bondRound4_(plRub / invested) : null
  };
}

function _computeBondsTopYtm_(rows, c, now) {
  var out = [];

  rows.forEach(function (r) {
    var name = _bondText_(r, c.name);
    var figi = _bondText_(r, c.figi);
    var price = _bondNumber_(r, c.price);
    var nominal = _bondNumber_(r, c.nominal);
    var couponPerYear = _bondNumber_(r, c.couponPerYear);
    var couponValue = _bondNumber_(r, c.couponValue);
    var maturity = _bondDate_(r, c.maturity);

    if (!name || !figi) return;
    if (!isFinite(price) || price <= 0) return;
    if (!isFinite(nominal) || nominal <= 0) return;
    if (!isFinite(couponPerYear) || couponPerYear <= 0) return;
    if (!isFinite(couponValue) || couponValue < 0) return;
    if (!maturity) return;

    var yearsToMatRaw = (_bondStripDate_(maturity) - _bondStripDate_(now)) / (365 * 24 * 3600 * 1000);
    if (!isFinite(yearsToMatRaw) || yearsToMatRaw < 0) return;

    var yearsToMat = Math.max(0.25, yearsToMatRaw);
    var annualCoupon = couponValue * couponPerYear;
    var ytm = (annualCoupon + (nominal - price) / yearsToMat) / ((nominal + price) / 2);

    if (!isFinite(ytm)) return;

    out.push({
      name: name,
      figi: figi,
      ytm: ytm,
      years: yearsToMat
    });
  });

  out.sort(function (a, b) {
    return (b.ytm || -1e9) - (a.ytm || -1e9);
  });

  return out.slice(0, 5).map(function (x) {
    return {
      name: x.name,
      figi: x.figi,
      ytm: _bondRound4_(x.ytm),
      years: _bondRound2_(x.years)
    };
  });
}

function _computeBondsTopPl_(rows, c, asc) {
  var out = [];

  rows.forEach(function (r) {
    var name = _bondText_(r, c.name);
    var invested = _bondNumber_(r, c.invested);
    var plRub = _bondNumber_(r, c.plRub);

    if (!name) return;
    if (!isFinite(invested) || invested <= 0) return;
    if (!isFinite(plRub)) return;

    out.push({
      name: name,
      plRub: plRub,
      plPct: plRub / invested
    });
  });

  out.sort(function (a, b) {
    return asc ? (a.plRub - b.plRub) : (b.plRub - a.plRub);
  });

  return out.slice(0, 5).map(function (x) {
    return {
      name: x.name,
      plRub: _bondRound2_(x.plRub),
      plPct: _bondRound4_(x.plPct)
    };
  });
}

function _computeBondsSectorSummary_(rows, c) {
  var map = {};

  rows.forEach(function (r) {
    var sector = _bondText_(r, c.sector) || '—';
    var market = _bondNumber_(r, c.market);
    var plRub = _bondNumber_(r, c.plRub);

    if (!map[sector]) {
      map[sector] = { sector: sector, market: 0, plRub: 0 };
    }

    if (isFinite(market)) map[sector].market += market;
    if (isFinite(plRub)) map[sector].plRub += plRub;
  });

  var arr = Object.keys(map).map(function (k) {
    return {
      sector: map[k].sector,
      market: _bondRound2_(map[k].market),
      plRub: _bondRound2_(map[k].plRub)
    };
  });

  arr.sort(function (a, b) {
    return (b.market || 0) - (a.market || 0);
  });

  return arr.slice(0, 8);
}

function _computeBondsMaturitySummary_(rows, c, now) {
  var map = {};

  rows.forEach(function (r) {
    var dt = _bondDate_(r, c.maturity);
    var market = _bondNumber_(r, c.market);

    if (!dt) return;
    if (!isFinite(market)) return;
    if (_bondStripDate_(dt) < _bondStripDate_(now)) return;

    var year = dt.getFullYear();
    map[year] = (map[year] || 0) + market;
  });

  var arr = Object.keys(map).map(function (y) {
    return {
      year: Number(y),
      market: _bondRound2_(map[y])
    };
  });

  arr.sort(function (a, b) {
    return a.year - b.year;
  });

  return arr.slice(0, 7);
}

function _computeBondsCoupon6mSummary_(rows, c, now) {
  var start = _bondStripDate_(now);
  var horizon = _bondAddMonths_(start, 6);

  var months = [];
  var monthMap = {};
  var i;

  for (i = 0; i < 6; i++) {
    var d = new Date(start.getFullYear(), start.getMonth() + i, 1);
    var key = d.getFullYear() + '-' + _bondPad2_(d.getMonth() + 1);
    months.push(key);
    monthMap[key] = 0;
  }

  rows.forEach(function (r) {
    var qty = _bondNumber_(r, c.qty);
    var couponPerYear = _bondNumber_(r, c.couponPerYear);
    var couponValue = _bondNumber_(r, c.couponValue);
    var nextCoupon = _bondDate_(r, c.nextCoupon);

    if (!isFinite(qty) || qty <= 0) return;
    if (!isFinite(couponPerYear) || couponPerYear <= 0) return;
    if (!isFinite(couponValue) || couponValue <= 0) return;
    if (!nextCoupon) return;

    var periodDays = Math.max(15, Math.round(365 / couponPerYear));
    var d2 = _bondStripDate_(nextCoupon);

    while (d2 <= horizon) {
      if (d2 >= start) {
        var key = d2.getFullYear() + '-' + _bondPad2_(d2.getMonth() + 1);
        if (monthMap.hasOwnProperty(key)) {
          monthMap[key] += couponValue * qty;
        }
      }
      d2 = _bondAddDays_(d2, periodDays);
    }
  });

  return months.map(function (key) {
    return {
      month: key,
      amount: _bondRound2_(monthMap[key] || 0)
    };
  });
}

function _buildGeneralBondCompareData_() {
  var data = _loadBondsRegionData_();
  if (!data || !data.rows || !data.rows.length) {
    return {
      hasData: false,
      weightedYtmPct: 0,
      weightedCouponPct: 0,
      scatterRows: [['Риск', 'YTM (%)', 'Тултип']]
    };
  }

  var c = _mapBondsRegionColumns_(data.header);
  var now = new Date();

  var wCouponNum = 0;
  var wCouponDen = 0;
  var wYtmNum = 0;
  var wYtmDen = 0;
  var scatterRows = [['Риск', 'YTM (%)', 'Тултип']];

  data.rows.forEach(function (r) {
    var riskNum = _bondNumber_(r, c.riskNum);
    var name = _bondText_(r, c.name);
    var figi = _bondText_(r, c.figi);
    var market = _bondNumber_(r, c.market);
    var price = _bondNumber_(r, c.price);
    var nominal = _bondNumber_(r, c.nominal);
    var couponPerYear = _bondNumber_(r, c.couponPerYear);
    var couponValue = _bondNumber_(r, c.couponValue);
    var couponPct = _bondNumber_(r, c.couponPct);
    var maturity = _bondDate_(r, c.maturity);

    if (!isFinite(couponPct)) {
      if (isFinite(price) && price > 0 && isFinite(couponPerYear) && couponPerYear > 0 && isFinite(couponValue) && couponValue >= 0) {
        couponPct = (couponValue * couponPerYear) / price * 100;
      }
    }

    if (isFinite(couponPct) && isFinite(market) && market > 0) {
      wCouponNum += couponPct * market;
      wCouponDen += market;
    }

    if (isFinite(price) && price > 0 &&
        isFinite(nominal) && nominal > 0 &&
        isFinite(couponPerYear) && couponPerYear > 0 &&
        isFinite(couponValue) && couponValue >= 0 &&
        maturity) {
      var yearsToMatRaw = (_bondStripDate_(maturity) - _bondStripDate_(now)) / (365 * 24 * 3600 * 1000);
      if (isFinite(yearsToMatRaw) && yearsToMatRaw >= 0) {
        var yearsToMat = Math.max(0.25, yearsToMatRaw);
        var annualCoupon = couponValue * couponPerYear;
        var ytmPct = ((annualCoupon + (nominal - price) / yearsToMat) / ((nominal + price) / 2)) * 100;

        if (isFinite(ytmPct)) {
          if (isFinite(market) && market > 0) {
            wYtmNum += ytmPct * market;
            wYtmDen += market;
          }

          if (isFinite(riskNum)) {
            scatterRows.push([
              riskNum,
              _bondRound2_(ytmPct),
              (name || '—') + '\n' + (figi || '') + '\nYTM: ' + _bondRound2_(ytmPct) + '%'
            ]);
          }
        }
      }
    }
  });

  return {
    hasData: wCouponDen > 0 || wYtmDen > 0 || scatterRows.length > 1,
    weightedYtmPct: wYtmDen > 0 ? _bondRound2_(wYtmNum / wYtmDen) : 0,
    weightedCouponPct: wCouponDen > 0 ? _bondRound2_(wCouponNum / wCouponDen) : 0,
    scatterRows: scatterRows
  };
}


/* ========================================================================
 * Bonds: render
 * ======================================================================== */

function renderBondsDashboardRegion_(sh, region) {
  region = dashboardNormalizeRegion_(region, dashboardGetLayout_().regions.bonds);

  dashboardClearRegion_(sh, region);
  dashboardWriteRegionTitle_(sh, region, region.title || 'Облигации');

  var data = _loadBondsRegionData_();
  if (!data.rows.length) {
    dashboardWritePlaceholder_(sh, region, 'Нет данных по облигациям');
    return;
  }

  var c = _mapBondsRegionColumns_(data.header);
  var now = new Date();

  var totals = _computeBondsTotals_(data.rows, c);
  var topYtm = _computeBondsTopYtm_(data.rows, c, now);
  var topPl = _computeBondsTopPl_(data.rows, c, false);
  var worstPl = _computeBondsTopPl_(data.rows, c, true);
  var sectors = _computeBondsSectorSummary_(data.rows, c);
  var maturity = _computeBondsMaturitySummary_(data.rows, c, now);
  var coupon6m = _computeBondsCoupon6mSummary_(data.rows, c, now);

  dashboardWriteTwoColBlock_(
    sh,
    _regionRow_(region, 3),
    region.startCol,
    'Облигации — итого',
    [
      ['Инвестировано', totals.invested],
      ['Рыночная стоимость', totals.market],
      ['P/L (руб)', totals.plRub],
      ['P/L (%)', totals.plPct]
    ],
    ['#,##0.00', '#,##0.00', '#,##0.00', '0.00%']
  );

  dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 10),
    region.startCol,
    'Облигации — top YTM',
    ['Бумага', 'YTM (%)', 'До погашения (лет)', 'FIGI'],
    topYtm.map(function (x) {
      return [x.name, x.ytm, x.years, x.figi];
    }),
    {
      formatsByCol: {
        2: '0.00%',
        3: '0.00'
      }
    }
  );

  dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 19),
    region.startCol,
    'Облигации — top P/L',
    ['Бумага', 'P/L (руб)', 'P/L (%)'],
    topPl.map(function (x) {
      return [x.name, x.plRub, x.plPct];
    }),
    {
      formatsByCol: {
        2: '#,##0.00',
        3: '0.00%'
      }
    }
  );

  dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 28),
    region.startCol,
    'Облигации — worst P/L',
    ['Бумага', 'P/L (руб)', 'P/L (%)'],
    worstPl.map(function (x) {
      return [x.name, x.plRub, x.plPct];
    }),
    {
      formatsByCol: {
        2: '#,##0.00',
        3: '0.00%'
      }
    }
  );

  dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 37),
    region.startCol,
    'Облигации — по секторам',
    ['Сектор', 'Рыночная стоимость', 'P/L (руб)'],
    sectors.map(function (x) {
      return [x.sector, x.market, x.plRub];
    }),
    {
      formatsByCol: {
        2: '#,##0.00',
        3: '#,##0.00'
      }
    }
  );

  var maturityInfo = dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 47),
    region.startCol,
    'Облигации — по сроку',
    ['Год', 'Рыночная стоимость'],
    maturity.map(function (x) {
      return [x.year, x.market];
    }),
    {
      formatsByCol: {
        2: '#,##0.00'
      }
    }
  );

  var couponInfo = dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 56),
    region.startCol,
    'Облигации — купоны 6м',
    ['Месяц', 'Выплаты'],
    coupon6m.map(function (x) {
      return [x.month, x.amount];
    }),
    {
      formatsByCol: {
        2: '#,##0.00'
      }
    }
  );

  if (maturityInfo && maturityInfo.dataRowsCount > 0) {
    dashboardBuildChartFromRanges_(sh, {
      chartType: Charts.ChartType.COLUMN,
      title: 'Облигации — сроки погашения',
      anchorRow: _regionRow_(region, 66),
      anchorCol: region.startCol,
      ranges: [maturityInfo.tableRange],
      width: 420,
      height: 220,
      options: {
        legend: { position: 'none' }
      }
    });
  }

  if (couponInfo && couponInfo.dataRowsCount > 0) {
    dashboardBuildChartFromRanges_(sh, {
      chartType: Charts.ChartType.LINE,
      title: 'Облигации — купоны 6м',
      anchorRow: _regionRow_(region, 86),
      anchorCol: region.startCol,
      ranges: [couponInfo.tableRange],
      width: 420,
      height: 220,
      options: {
        legend: { position: 'none' }
      }
    });
  }
}


/* ========================================================================
 * Funds: data + calculations
 * ======================================================================== */

function _loadFundsRegionData_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Funds');

  if (!sh) return { header: [], rows: [] };
  if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) return { header: [], rows: [] };

  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  rows = rows.filter(function (r) {
    return _rowHasAnyValue_(r);
  });

  return { header: header, rows: rows };
}

function _mapFundsRegionColumns_(header) {
  return {
    name: _findHeaderIndexByAliases_(header, ['Название', 'Наименование', 'Инструмент', 'Компания']),
    ticker: _findHeaderIndexByAliases_(header, ['Тикер', 'Ticker']),
    sector: _findHeaderIndexByAliases_(header, ['Сектор', 'Категория', 'Тема', 'Фокус']),
    currency: _findHeaderIndexByAliases_(header, ['Валюта', 'Currency']),
    invested: _findHeaderIndexByAliases_(header, ['Инвестировано']),
    market: _findHeaderIndexByAliases_(header, ['Рыночная стоимость', 'Стоимость', 'Рыночная стоимость, ₽']),
    plRub: _findHeaderIndexByAliases_(header, ['P/L (руб)', 'P/L, руб', 'P/L RUB']),
    plPct: _findHeaderIndexByAliases_(header, ['P/L (%)', 'P/L %'])
  };
}

function _computeFundsTotals_(rows, cols) {
  var invested = 0;
  var market = 0;
  var plRub = 0;

  rows.forEach(function (row) {
    var inv = _toNumberSafe_(_valueByIndex_(row, cols.invested));
    var mkt = _toNumberSafe_(_valueByIndex_(row, cols.market));
    var pl = cols.plRub >= 0 ? _toNumberSafe_(_valueByIndex_(row, cols.plRub)) : (mkt - inv);

    invested += inv;
    market += mkt;
    plRub += pl;
  });

  return {
    invested: invested,
    market: market,
    plRub: plRub,
    plPct: invested !== 0 ? (plRub / invested) : 0
  };
}

function _computeFundsTopByMarket_(rows, cols, limit) {
  return rows
    .map(function (row) {
      var invested = _toNumberSafe_(_valueByIndex_(row, cols.invested));
      var market = _toNumberSafe_(_valueByIndex_(row, cols.market));
      var plRub = cols.plRub >= 0 ? _toNumberSafe_(_valueByIndex_(row, cols.plRub)) : (market - invested);

      return {
        name: _safeTextByIndex_(row, cols.name, '—'),
        ticker: _safeTextByIndex_(row, cols.ticker, '—'),
        invested: invested,
        market: market,
        plRub: plRub,
        plPct: invested !== 0 ? (plRub / invested) : 0
      };
    })
    .filter(function (x) {
      return x.name !== '—' || x.market !== 0 || x.plRub !== 0;
    })
    .sort(function (a, b) {
      return b.market - a.market;
    })
    .slice(0, limit || 5);
}

function _computeFundsTopByPl_(rows, cols, limit) {
  return rows
    .map(function (row) {
      var invested = _toNumberSafe_(_valueByIndex_(row, cols.invested));
      var market = _toNumberSafe_(_valueByIndex_(row, cols.market));
      var plRub = cols.plRub >= 0 ? _toNumberSafe_(_valueByIndex_(row, cols.plRub)) : (market - invested);

      return {
        name: _safeTextByIndex_(row, cols.name, '—'),
        ticker: _safeTextByIndex_(row, cols.ticker, '—'),
        invested: invested,
        market: market,
        plRub: plRub,
        plPct: invested !== 0 ? (plRub / invested) : 0
      };
    })
    .filter(function (x) {
      return x.name !== '—' || x.market !== 0 || x.plRub !== 0;
    })
    .sort(function (a, b) {
      return b.plRub - a.plRub;
    })
    .slice(0, limit || 5);
}

function _computeFundsSectorSummary_(rows, cols) {
  var totalMarket = 0;
  var map = {};

  rows.forEach(function (row) {
    var sector = _safeTextByIndex_(row, cols.sector, '—');
    var market = _toNumberSafe_(_valueByIndex_(row, cols.market));

    totalMarket += market;
    map[sector] = (map[sector] || 0) + market;
  });

  return Object.keys(map)
    .map(function (sector) {
      var market = map[sector];
      return {
        sector: sector,
        market: market,
        share: totalMarket !== 0 ? (market / totalMarket) : 0
      };
    })
    .sort(function (a, b) {
      return b.market - a.market;
    });
}

function _computeFundsCurrencySummary_(rows, cols) {
  var totalMarket = 0;
  var map = {};

  rows.forEach(function (row) {
    var currency = _safeTextByIndex_(row, cols.currency, '—');
    var market = _toNumberSafe_(_valueByIndex_(row, cols.market));

    totalMarket += market;
    map[currency] = (map[currency] || 0) + market;
  });

  return Object.keys(map)
    .map(function (currency) {
      var market = map[currency];
      return {
        currency: currency,
        market: market,
        share: totalMarket !== 0 ? (market / totalMarket) : 0
      };
    })
    .sort(function (a, b) {
      return b.market - a.market;
    });
}


/* ========================================================================
 * Funds: render
 * ======================================================================== */

function renderFundsDashboardRegion_(sh, region) {
  region = dashboardNormalizeRegion_(region, dashboardGetLayout_().regions.funds);

  dashboardClearRegion_(sh, region);
  dashboardWriteRegionTitle_(sh, region, region.title || 'Фонды');

  var data = _loadFundsRegionData_();
  if (!data || !data.rows || !data.rows.length) {
    dashboardWritePlaceholder_(sh, region, 'Нет данных по фондам');
    return;
  }

  var cols = _mapFundsRegionColumns_(data.header);
  var totals = _computeFundsTotals_(data.rows, cols);
  var topByMarket = _computeFundsTopByMarket_(data.rows, cols, 5);
  var topByPl = _computeFundsTopByPl_(data.rows, cols, 5);
  var bySector = _computeFundsSectorSummary_(data.rows, cols);
  var byCurrency = _computeFundsCurrencySummary_(data.rows, cols);

  dashboardWriteTwoColBlock_(
    sh,
    _regionRow_(region, 3),
    region.startCol,
    'Фонды — итого',
    [
      ['Инвестировано', totals.invested],
      ['Рыночная стоимость', totals.market],
      ['P/L (руб)', totals.plRub],
      ['P/L (%)', totals.plPct]
    ],
    ['#,##0.00', '#,##0.00', '#,##0.00', '0.00%']
  );

  dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 10),
    region.startCol,
    'Фонды — top-5 по позиции',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    topByMarket.map(function (x) {
      return [x.name, x.ticker, x.market, x.plRub, x.plPct];
    }),
    {
      formatsByCol: {
        3: '#,##0.00',
        4: '#,##0.00',
        5: '0.00%'
      }
    }
  );

  dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 19),
    region.startCol,
    'Фонды — top-5 по P/L',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    topByPl.map(function (x) {
      return [x.name, x.ticker, x.market, x.plRub, x.plPct];
    }),
    {
      formatsByCol: {
        3: '#,##0.00',
        4: '#,##0.00',
        5: '0.00%'
      }
    }
  );

  var sectorInfo = dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 28),
    region.startCol,
    'Фонды — по секторам',
    ['Сектор', 'Рыночная стоимость', 'Доля фондов'],
    bySector.map(function (x) {
      return [x.sector, x.market, x.share];
    }),
    {
      formatsByCol: {
        2: '#,##0.00',
        3: '0.00%'
      }
    }
  );

  var currencyInfo = dashboardWriteTableBlock_(
    sh,
    _regionRow_(region, 38),
    region.startCol,
    'Фонды — по валютам',
    ['Валюта', 'Рыночная стоимость', 'Доля фондов'],
    byCurrency.map(function (x) {
      return [x.currency, x.market, x.share];
    }),
    {
      formatsByCol: {
        2: '#,##0.00',
        3: '0.00%'
      }
    }
  );

  if (sectorInfo && sectorInfo.dataRowsCount > 0) {
    dashboardBuildChartFromRanges_(sh, {
      chartType: Charts.ChartType.PIE,
      title: 'Фонды — структура',
      anchorRow: _regionRow_(region, 48),
      anchorCol: region.startCol,
      ranges: [
        sh.getRange(sectorInfo.headerRow, region.startCol, 1 + sectorInfo.dataRowsCount, 2)
      ],
      width: 420,
      height: 220,
      options: {
        legend: { position: 'right' }
      }
    });
  } else if (currencyInfo && currencyInfo.dataRowsCount > 0) {
    dashboardBuildChartFromRanges_(sh, {
      chartType: Charts.ChartType.PIE,
      title: 'Фонды — структура',
      anchorRow: _regionRow_(region, 48),
      anchorCol: region.startCol,
      ranges: [
        sh.getRange(currencyInfo.headerRow, region.startCol, 1 + currencyInfo.dataRowsCount, 2)
      ],
      width: 420,
      height: 220,
      options: {
        legend: { position: 'right' }
      }
    });
  }
}


/* ========================================================================
 * Common dashboard render helpers
 * ======================================================================== */

function dashboardEnsureSheet_(sheetName) {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function dashboardNormalizeRegion_(region, defaults) {
  region = region || {};
  defaults = defaults || {};

  return {
    key: region.key || defaults.key || '',
    title: region.title || defaults.title || '',
    startCol: Number(region.startCol) || Number(defaults.startCol) || 1,
    width: Number(region.width) || Number(defaults.width) || 1,
    startRow: Number(region.startRow) || Number(defaults.startRow) || 1,
    maxRows: Number(region.maxRows) || Number(defaults.maxRows) || 50
  };
}

function dashboardResetSheet_(sh, layout) {
  var maxRows = Math.max(sh.getLastRow(), layout.bounds.maxRows);
  var maxCols = Math.max(sh.getLastColumn(), layout.bounds.maxCols);
  var rng;

  if (maxRows < 1 || maxCols < 1) return;

  rng = sh.getRange(1, 1, maxRows, maxCols);

  try { rng.breakApart(); } catch (e1) {}
  rng.clearContent();
  rng.clearFormat();
  try { rng.clearDataValidations(); } catch (e2) {}
  try { rng.clearNote(); } catch (e3) {}

  var charts = sh.getCharts() || [];
  var i;
  for (i = charts.length - 1; i >= 0; i--) {
    sh.removeChart(charts[i]);
  }
}

function dashboardClearRegion_(sh, region) {
  var rng = sh.getRange(region.startRow, region.startCol, region.maxRows, region.width);

  try { rng.breakApart(); } catch (e1) {}
  rng.clearContent();
  rng.clearFormat();
  try { rng.clearDataValidations(); } catch (e2) {}
  try { rng.clearNote(); } catch (e3) {}

  dashboardRemoveChartsInRegion_(sh, region);
}

function dashboardRemoveChartsInRegion_(sh, region) {
  var rowMin = region.startRow;
  var rowMax = region.startRow + region.maxRows - 1;
  var colMin = region.startCol;
  var colMax = region.startCol + region.width - 1;

  var charts = sh.getCharts() || [];
  var i;

  for (i = charts.length - 1; i >= 0; i--) {
    try {
      var info = charts[i].getContainerInfo();
      var r = info.getAnchorRow();
      var c = info.getAnchorColumn();

      if (r >= rowMin && r <= rowMax && c >= colMin && c <= colMax) {
        sh.removeChart(charts[i]);
      }
    } catch (e) {}
  }
}

function dashboardWriteRegionTitle_(sh, region, title) {
  var rng = sh.getRange(region.startRow, region.startCol, 1, region.width);

  try { rng.breakApart(); } catch (e) {}
  rng.merge();
  rng
    .setValue(title)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#E5E7EB');
}

function dashboardWritePlaceholder_(sh, region, text) {
  dashboardWriteRegionTitle_(sh, region, region.title || region.key || 'Раздел');

  var rng = sh.getRange(region.startRow + 2, region.startCol, 1, region.width);
  try { rng.breakApart(); } catch (e) {}
  rng.merge();
  rng
    .setValue(text)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
}

function dashboardWriteTwoColBlock_(sh, startRow, startCol, title, rows, valueFormats) {
  sh.getRange(startRow, startCol, 1, 2)
    .setValues([[title, '']])
    .setFontWeight('bold');

  if (!rows || !rows.length) {
    return {
      titleRow: startRow,
      dataStartRow: startRow + 1,
      dataRowsCount: 0
    };
  }

  sh.getRange(startRow + 1, startCol, rows.length, 2).setValues(rows);
  sh.getRange(startRow + 1, startCol, rows.length, 1).setFontWeight('bold');
  sh.getRange(startRow + 1, startCol + 1, rows.length, 1).setHorizontalAlignment('right');

  var i;
  for (i = 0; i < (valueFormats || []).length; i++) {
    if (!valueFormats[i]) continue;
    sh.getRange(startRow + 1 + i, startCol + 1, 1, 1).setNumberFormat(valueFormats[i]);
  }

  return {
    titleRow: startRow,
    dataStartRow: startRow + 1,
    dataRowsCount: rows.length
  };
}

function dashboardWriteTableBlock_(sh, startRow, startCol, title, headers, rows, opts) {
  opts = opts || {};
  rows = rows || [];

  var width = headers.length;
  var titleRow = [];
  var i;
  for (i = 0; i < width; i++) {
    titleRow.push(i === 0 ? title : '');
  }

  sh.getRange(startRow, startCol, 1, width)
    .setValues([titleRow])
    .setFontWeight('bold');

  sh.getRange(startRow + 1, startCol, 1, width)
    .setValues([headers])
    .setFontWeight('bold');

  if (rows.length) {
    sh.getRange(startRow + 2, startCol, rows.length, width).setValues(rows);
  }

  var formatsByCol = opts.formatsByCol || {};
  Object.keys(formatsByCol).forEach(function (key) {
    var colIndex = Number(key);
    if (!isFinite(colIndex) || colIndex < 1 || colIndex > width) return;
    if (!rows.length) return;

    sh.getRange(startRow + 2, startCol + colIndex - 1, rows.length, 1)
      .setNumberFormat(formatsByCol[key]);
  });

  return {
    titleRow: startRow,
    headerRow: startRow + 1,
    dataStartRow: startRow + 2,
    dataRowsCount: rows.length,
    tableRange: sh.getRange(startRow + 1, startCol, 1 + rows.length, width)
  };
}

function dashboardBuildChartFromRanges_(sh, cfg) {
  if (!cfg || !cfg.ranges || !cfg.ranges.length) return null;

  var builder = sh.newChart()
    .setChartType(cfg.chartType || Charts.ChartType.COLUMN)
    .setPosition(cfg.anchorRow || 1, cfg.anchorCol || 1, 0, 0)
    .setOption('title', cfg.title || '')
    .setOption('width', Math.min(Number(cfg.width) || 400, 420))
    .setOption('height', Math.min(Number(cfg.height) || 220, 220));

  var i;
  for (i = 0; i < cfg.ranges.length; i++) {
    builder.addRange(cfg.ranges[i]);
  }

  var options = cfg.options || {};
  for (var k in options) {
    if (!options.hasOwnProperty(k)) continue;
    builder.setOption(k, options[k]);
  }

  var chart = builder.build();
  sh.insertChart(chart);
  return chart;
}

function dashboardAutoResizeRegions_(sh, layout) {
  var regions = layout.regions;
  var keys = Object.keys(regions);
  var i;

  for (i = 0; i < keys.length; i++) {
    sh.autoResizeColumns(regions[keys[i]].startCol, regions[keys[i]].width);
  }

  try { sh.setColumnWidth(11, 22); } catch (e1) {}
  try { sh.setColumnWidth(22, 22); } catch (e2) {}
  try { sh.setColumnWidth(33, 22); } catch (e3) {}
}

function _dashboardChartRangeHasData_(range) {
  try {
    if (!range) return false;
    if (range.getNumRows() < 2 || range.getNumColumns() < 2) return false;

    var vals = range.getDisplayValues();
    var hasHeader = false;
    var hasData = false;
    var r;
    var c;

    for (c = 0; c < vals[0].length; c++) {
      if (String(vals[0][c] || '').trim()) {
        hasHeader = true;
        break;
      }
    }

    for (r = 1; r < vals.length; r++) {
      for (c = 0; c < vals[r].length; c++) {
        if (String(vals[r][c] || '').trim()) {
          hasData = true;
          break;
        }
      }
      if (hasData) break;
    }

    return hasHeader && hasData;
  } catch (e) {
    return false;
  }
}

function _regionRow_(region, relativeRow) {
  return region.startRow + relativeRow - 1;
}


/* ========================================================================
 * Primitive helpers
 * ======================================================================== */

function _readSheetTotalsByHeaders_(sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheetName);
  var out = { invested: 0, market: 0, plRub: 0, plPct: null, count: 0 };
  var data;
  var header;
  var rows;
  var idxInvested;
  var idxMarket;
  var idxPlRub;
  var i;

  if (!sh) return out;
  if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) return out;

  data = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  if (!data || !data.length) return out;

  header = data[0] || [];
  rows = data.slice(1);

  idxInvested = _findHeaderIndexByAliases_(header, ['Инвестировано']);
  idxMarket = _findHeaderIndexByAliases_(header, ['Рыночная стоимость']);
  idxPlRub = _findHeaderIndexByAliases_(header, ['P/L (руб)', 'P/L, руб', 'P/L RUB']);

  for (i = 0; i < rows.length; i++) {
    if (!_rowHasAnyValue_(rows[i])) continue;

    out.count += 1;
    if (idxInvested >= 0) out.invested += _toNumberSafe_(rows[i][idxInvested]);
    if (idxMarket >= 0) out.market += _toNumberSafe_(rows[i][idxMarket]);
    if (idxPlRub >= 0) out.plRub += _toNumberSafe_(rows[i][idxPlRub]);
  }

  out.plPct = out.invested ? out.plRub / out.invested : null;
  return out;
}

function _findHeaderIndexByAliases_(header, aliases) {
  var normalizedHeader = [];
  var i;
  var j;

  for (i = 0; i < header.length; i++) {
    normalizedHeader.push(_normalizeHeaderText_(header[i]));
  }

  for (j = 0; j < aliases.length; j++) {
    var alias = _normalizeHeaderText_(aliases[j]);
    for (i = 0; i < normalizedHeader.length; i++) {
      if (normalizedHeader[i] === alias) return i;
    }
  }

  for (j = 0; j < aliases.length; j++) {
    alias = _normalizeHeaderText_(aliases[j]);
    for (i = 0; i < normalizedHeader.length; i++) {
      if (normalizedHeader[i] && alias && normalizedHeader[i].indexOf(alias) >= 0) return i;
    }
  }

  return -1;
}

function _normalizeHeaderText_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[ё]/g, 'е')
    .trim();
}

function _toNumberSafe_(value) {
  var s;
  var n;

  if (value == null || value === '') return 0;
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  if (typeof value === 'boolean') return value ? 1 : 0;

  s = String(value).replace(/\s+/g, '').replace(',', '.');
  n = Number(s);

  return isNaN(n) ? 0 : n;
}

function _rowHasAnyValue_(row) {
  var i;
  if (!row) return false;

  for (i = 0; i < row.length; i++) {
    if (row[i] !== '' && row[i] != null) return true;
  }
  return false;
}

function _valueByIndex_(row, idx) {
  if (idx == null || idx < 0 || idx >= row.length) return '';
  return row[idx];
}

function _safeTextByIndex_(row, idx, fallback) {
  var s = String(_valueByIndex_(row, idx) == null ? '' : _valueByIndex_(row, idx)).trim();
  return s || (fallback || '');
}

function _pushNumberIfFinite_(arr, value) {
  var n = Number(value);
  if (isFinite(n)) arr.push(n);
}

function _pushMarketAgg_(map, key, market) {
  key = String(key == null || key === '' ? '—' : key);
  map[key] = (map[key] || 0) + (Number(market) || 0);
}

function _buildMarketShareRows_(map, totalMarket) {
  return Object.keys(map)
    .map(function (key) {
      return {
        name: key,
        market: Math.round(Number(map[key] || 0) * 100) / 100,
        sharePct: totalMarket ? (Number(map[key] || 0) / totalMarket) : 0
      };
    })
    .sort(function (a, b) {
      return b.market - a.market;
    });
}


/* ---------- Bonds primitive helpers ---------- */

function _bondValue_(r, ci) {
  return ci ? r[ci - 1] : '';
}

function _bondText_(r, ci) {
  var v = _bondValue_(r, ci);
  return String(v == null ? '' : v).trim();
}

function _bondNumber_(r, ci) {
  var x = _bondValue_(r, ci);

  if (x === null || x === '' || typeof x === 'undefined') return NaN;
  if (typeof x === 'number') return isFinite(x) ? x : NaN;
  if (x instanceof Date) return NaN;

  var s = String(x)
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, '')
    .replace(/%/g, '')
    .replace(',', '.')
    .trim();

  if (!s) return NaN;

  var n = Number(s);
  return isFinite(n) ? n : NaN;
}

function _bondDate_(r, ci) {
  var v = _bondValue_(r, ci);
  if (!v) return null;

  if (v instanceof Date) {
    return isFinite(v.getTime()) ? v : null;
  }

  if (typeof v === 'string') {
    var s = v.trim();

    var m1 = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m1) {
      var d1 = new Date(Number(m1[3]), Number(m1[2]) - 1, Number(m1[1]));
      return isFinite(d1.getTime()) ? d1 : null;
    }

    var m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m2) {
      var d2 = new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));
      return isFinite(d2.getTime()) ? d2 : null;
    }

    var parsed = new Date(s);
    return isFinite(parsed.getTime()) ? parsed : null;
  }

  return null;
}

function _bondStripDate_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function _bondAddDays_(d, days) {
  var x = new Date(d.getTime());
  x.setDate(x.getDate() + days);
  return x;
}

function _bondAddMonths_(d, months) {
  return new Date(d.getFullYear(), d.getMonth() + months, d.getDate());
}

function _bondPad2_(n) {
  return ('0' + n).slice(-2);
}

function _bondRound2_(x) {
  return Math.round(Number(x || 0) * 100) / 100;
}

function _bondRound4_(x) {
  return Math.round(Number(x || 0) * 10000) / 10000;
}
