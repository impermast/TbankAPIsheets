/**
 * dashboard_shares_helpers.gs
 * Helper-функции для чтения и расчётов по листу Shares.
 * Без рендера, без форматирования, только чтение и агрегации.
 */

function loadSharesDashboardData_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Shares');
  if (!sh) return { header: [], rows: [] };

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { header: [], rows: [] };

  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return { header: header, rows: rows };
}

function mapSharesDashboardColumns_(hdr) {
  function idx(name) {
    var i = hdr.indexOf(name);
    return i >= 0 ? (i + 1) : 0;
  }

  return {
    name:      idx('Название'),
    ticker:    idx('Тикер'),
    market:    idx('Рыночная стоимость'),
    invested:  idx('Инвестировано'),
    plRub:     idx('P/L (руб)'),
    plPct:     idx('P/L (%)'),
    sector:    idx('Сектор'),
    country:   idx('Страна риска'),
    pe:        idx('P/E TTM'),
    pb:        idx('P/B TTM'),
    evEbitda:  idx('EV/EBITDA'),
    roe:       idx('ROE'),
    beta:      idx('Beta')
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
      plPct: x.plPct === null ? null : round2_(x.plPct)
    };
  });

  bestPL = bestPL.map(function (x) {
    return {
      name: x.name,
      ticker: x.ticker,
      market: round2_(x.market),
      plRub: round2_(x.plRub),
      plPct: x.plPct === null ? null : round2_(x.plPct)
    };
  });

  worstPL = worstPL.map(function (x) {
    return {
      name: x.name,
      ticker: x.ticker,
      market: round2_(x.market),
      plRub: round2_(x.plRub),
      plPct: x.plPct === null ? null : round2_(x.plPct)
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

function _medianNumber_(arr) {
  if (!arr || !arr.length) return null;

  var nums = arr
    .map(function (x) {
      return typeof x === 'number' ? x : Number(x);
    })
    .filter(function (x) {
      return isFinite(x);
    })
    .sort(function (a, b) {
      return a - b;
    });

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
