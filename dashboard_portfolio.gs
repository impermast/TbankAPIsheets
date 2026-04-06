**
 * dashboard_portfolio.gs
 */


"function buildPortfolioDashboard() {
  var ss = SpreadsheetApp.getActive();
  var sh;
  var startCol = 20;
  var totalRow = 1;
  var classRow = 8;
  var sharesRow = 17;
  var byClass;
  var totals;

  if (typeof buildBondsDashboard === 'function') {
    try { buildBondsDashboard(); } catch (e) {}
  }

  sh = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');

  byClass = {
    bonds: _readSheetTotalsByHeaders_('Bonds'),
    funds: _readSheetTotalsByHeaders_('Funds'),
    shares: _readSheetTotalsByHeaders_('Shares'),
    options: _readSheetTotalsByHeaders_('Options')
  };

  totals = {
    invested: (byClass.bonds.invested || 0) + (byClass.funds.invested || 0) + (byClass.shares.invested || 0) + (byClass.options.invested || 0),
    market: (byClass.bonds.market || 0) + (byClass.funds.market || 0) + (byClass.shares.market || 0) + (byClass.options.market || 0),
    plRub: (byClass.bonds.plRub || 0) + (byClass.funds.plRub || 0) + (byClass.shares.plRub || 0) + (byClass.options.plRub || 0),
    plPct: null,
    count: (byClass.bonds.count || 0) + (byClass.funds.count || 0) + (byClass.shares.count || 0) + (byClass.options.count || 0)
  };
  totals.plPct = totals.invested ? totals.plRub / totals.invested : null;

  _writePortfolioTotalBlock_(sh, totalRow, startCol, totals);
  _writeAssetClassSummaryBlock_(sh, classRow, startCol, byClass, totals.market);
  _writeSharesBlocks_(sh, sharesRow, startCol, {
    byClass: byClass,
    totalMarket: totals.market,
    sharesTotals: byClass.shares
  });

  sh.autoResizeColumns(startCol, 8);
}

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

function _writePortfolioTotalBlock_(sh, row, col, totals) {
  var values = [
    ['Портфель — итого', ''],
    ['Инвестировано', totals && totals.invested != null ? totals.invested : 0],
    ['Рыночная стоимость', totals && totals.market != null ? totals.market : 0],
    ['P/L (руб)', totals && totals.plRub != null ? totals.plRub : 0],
    ['P/L (%)', totals && totals.plPct != null ? totals.plPct : '']
  ];

  sh.getRange(row, col, values.length, 2).setValues(values);
  sh.getRange(row, col, 1, 2).setFontWeight('bold');
  sh.getRange(row + 1, col + 1, 3, 1).setNumberFormat('#,##0.00');
  sh.getRange(row + 4, col + 1, 1, 1).setNumberFormat('0.00%');
}

function _writeAssetClassSummaryBlock_(sh, row, col, byClass, totalMarket) {
  var titleWidth = 7;
  var header = ['Класс', 'Бумаг', 'Инвестировано', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)', 'Доля портфеля'];
  var rows = [
    _assetClassRow_('Облигации', byClass && byClass.bonds, totalMarket),
    _assetClassRow_('Фонды', byClass && byClass.funds, totalMarket),
    _assetClassRow_('Акции', byClass && byClass.shares, totalMarket),
    _assetClassRow_('Опционы', byClass && byClass.options, totalMarket)
  ];

  sh.getRange(row, col, 1, titleWidth).setValues([['Классы активов', '', '', '', '', '', '']]);
  sh.getRange(row, col, 1, titleWidth).setFontWeight('bold');
  sh.getRange(row + 1, col, 1, header.length).setValues([header]);
  sh.getRange(row + 1, col, 1, header.length).setFontWeight('bold');
  sh.getRange(row + 2, col, rows.length, header.length).setValues(rows);
  sh.getRange(row + 2, col + 2, rows.length, 3).setNumberFormat('#,##0.00');
  sh.getRange(row + 2, col + 5, rows.length, 2).setNumberFormat('0.00%');
}

function _writeSharesBlocks_(sh, row, col, data) {
  var ss = SpreadsheetApp.getActive();
  var sharesSheet = ss.getSheetByName('Shares');
  var sharesData;
  var rows;
  var currentRow = row;

  if (!sharesSheet || sharesSheet.getLastRow() < 2 || sharesSheet.getLastColumn() < 1) {
    _writeTextBlock_(sh, currentRow, col, 'Акции', 'Нет данных по акциям', 4);
    return currentRow + 4;
  }

  if (!(typeof loadSharesDashboardData_ === 'function' &&
        typeof mapSharesDashboardColumns_ === 'function' &&
        typeof validateSharesDashboardColumns_ === 'function' &&
        typeof computeSharesDashboardData_ === 'function')) {
    _writeTextBlock_(sh, currentRow, col, 'Акции', 'Shares helpers not found', 4);
    return currentRow + 4;
  }

  sharesData = _prepareSharesDashboardBlocksData_(sharesSheet, data || {});
  if (!sharesData || !sharesData.hasData) {
    _writeTextBlock_(sh, currentRow, col, 'Акции', 'Нет данных по акциям', 4);
    return currentRow + 4;
  }

  rows = [
    ['Инвестировано', sharesData.totals.invested],
    ['Рыночная стоимость', sharesData.totals.market],
    ['P/L (руб)', sharesData.totals.plRub],
    ['P/L (%)', sharesData.totals.plPct]
  ];
  _writeLabelValueBlock_(sh, currentRow, col, 'Акции — итого', rows, [1, 2, 3], [4]);
  currentRow += rows.length + 3;

  _writeDataTableBlock_(
    sh,
    currentRow,
    col,
    'Акции — топ-5 по позиции',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    _mapTopPositionRows_(sharesData.topPosition),
    [2, 3],
    [4]
  );
  currentRow += _blockHeight_(sharesData.topPosition, 5) + 2;

  _writeDataTableBlock_(
    sh,
    currentRow,
    col,
    'Акции — топ-5 по P/L',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    _mapTopPlRows_(sharesData.topProfit),
    [2, 3],
    [4]
  );
  currentRow += _blockHeight_(sharesData.topProfit, 5) + 2;

  _writeDataTableBlock_(
    sh,
    currentRow,
    col,
    'Акции — топ-5 убыток',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    _mapTopPlRows_(sharesData.topLoss),
    [2, 3],
    [4]
  );
  currentRow += _blockHeight_(sharesData.topLoss, 5) + 2;

  _writeDataTableBlock_(
    sh,
    currentRow,
    col,
    'Акции — по секторам',
    ['Сектор', 'Бумаг', 'Инвестировано', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)', 'Доля акций'],
    _mapGroupRows_(sharesData.sectors),
    [2, 3, 4],
    [5, 6]
  );
  currentRow += _blockHeight_(sharesData.sectors, 7) + 2;

  _writeDataTableBlock_(
    sh,
    currentRow,
    col,
    'Акции — по странам',
    ['Страна', 'Бумаг', 'Инвестировано', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)', 'Доля акций'],
    _mapGroupRows_(sharesData.countries),
    [2, 3, 4],
    [5, 6]
  );
  currentRow += _blockHeight_(sharesData.countries, 7) + 2;

  _writeDataTableBlock_(
    sh,
    currentRow,
    col,
    'Акции — valuation snapshot',
    ['Метрика', 'Значение'],
    sharesData.valuation && sharesData.valuation.length ? sharesData.valuation : [['Нет данных', '']],
    [],
    []
  );

  return currentRow;
}

function _assetClassRow_(title, totals, totalMarket) {
  var t = totals || {};
  var share = totalMarket ? (_toNumberSafe_(t.market) / totalMarket) : null;
  return [
    title,
    _toNumberSafe_(t.count),
    _toNumberSafe_(t.invested),
    _toNumberSafe_(t.market),
    _toNumberSafe_(t.plRub),
    t.plPct != null ? t.plPct : (_toNumberSafe_(t.invested) ? _toNumberSafe_(t.plRub) / _toNumberSafe_(t.invested) : null),
    share
  ];
}

function _writeTextBlock_(sh, row, col, title, text, width) {
  var w = width || 2;
  var titleRow = [];
  var i;

  for (i = 0; i < w; i++) titleRow.push(i === 0 ? title : '');
  sh.getRange(row, col, 1, w).setValues([titleRow]);
  sh.getRange(row, col, 1, w).setFontWeight('bold');
  sh.getRange(row + 1, col, 1, w).setValues([[text, '', '', ''].slice(0, w)]);
}

function _writeLabelValueBlock_(sh, row, col, title, rows, moneyRowIndexes, pctRowIndexes) {
  var values = [[title, '']].concat(rows || []);
  var i;

  sh.getRange(row, col, values.length, 2).setValues(values);
  sh.getRange(row, col, 1, 2).setFontWeight('bold');

  moneyRowIndexes = moneyRowIndexes || [];
  for (i = 0; i < moneyRowIndexes.length; i++) {
    sh.getRange(row + moneyRowIndexes[i], col + 1, 1, 1).setNumberFormat('#,##0.00');
  }

  pctRowIndexes = pctRowIndexes || [];
  for (i = 0; i < pctRowIndexes.length; i++) {
    sh.getRange(row + pctRowIndexes[i], col + 1, 1, 1).setNumberFormat('0.00%');
  }
}

function _writeDataTableBlock_(sh, row, col, title, header, rows, moneyCols, pctCols) {
  var width = header.length;
  var titleRow = [];
  var body = rows && rows.length ? rows : [_emptyDataRow_(width)];
  var i;

  for (i = 0; i < width; i++) titleRow.push(i === 0 ? title : '');

  sh.getRange(row, col, 1, width).setValues([titleRow]);
  sh.getRange(row, col, 1, width).setFontWeight('bold');
  sh.getRange(row + 1, col, 1, width).setValues([header]);
  sh.getRange(row + 1, col, 1, width).setFontWeight('bold');
  sh.getRange(row + 2, col, body.length, width).setValues(body);

  moneyCols = moneyCols || [];
  for (i = 0; i < moneyCols.length; i++) {
    sh.getRange(row + 2, col + moneyCols[i], body.length, 1).setNumberFormat('#,##0.00');
  }

  pctCols = pctCols || [];
  for (i = 0; i < pctCols.length; i++) {
    sh.getRange(row + 2, col + pctCols[i], body.length, 1).setNumberFormat('0.00%');
  }
}

function _blockHeight_(rows, width) {
  var len = rows && rows.length ? rows.length : 1;
  return 2 + len;
}

function _emptyDataRow_(width) {
  var row = [];
  var i;
  for (i = 0; i < width; i++) row.push(i === 0 ? 'Нет данных' : '');
  return row;
}

function _prepareSharesDashboardBlocksData_(sharesSheet, data) {
  var loaded;
  var mapped;
  var valid;
  var computed;
  var normalizedLoaded;
  var fallback;

  fallback = _buildSharesFallbackData_(sharesSheet, data || {});

  loaded = _tryCallVariants_(loadSharesDashboardData_, [
    [sharesSheet],
    [],
    [sharesSheet.getName()],
    [SpreadsheetApp.getActive()]
  ]);
  normalizedLoaded = _normalizeLoadedSheetData_(loaded, sharesSheet);

  mapped = _tryCallVariants_(mapSharesDashboardColumns_, [
    [normalizedLoaded.header],
    [normalizedLoaded],
    [normalizedLoaded.header || []]
  ]);

  valid = _tryCallVariants_(validateSharesDashboardColumns_, [
    [mapped],
    [normalizedLoaded.header, mapped],
    [normalizedLoaded, mapped]
  ]);

  if (valid === false) return fallback;

  computed = _tryCallVariants_(computeSharesDashboardData_, [
    [normalizedLoaded.rows, mapped],
    [normalizedLoaded, mapped],
    [normalizedLoaded.rows, mapped, normalizedLoaded],
    [normalizedLoaded],
    [sharesSheet, normalizedLoaded, mapped]
  ]);

  return _mergeSharesComputedWithFallback_(computed, normalizedLoaded, fallback, data || {});
}

function _buildSharesFallbackData_(sharesSheet, data) {
  var loaded = _normalizeLoadedSheetData_(null, sharesSheet);
  var cols = _mapGenericSharesColumns_(loaded.header);
  var totals = { invested: 0, market: 0, plRub: 0, plPct: null, count: 0 };
  var positions = [];
  var sectors = {};
  var countries = {};
  var valuationAgg = {
    pe: { sum: 0, w: 0 },
    ps: { sum: 0, w: 0 },
    pb: { sum: 0, w: 0 },
    ev: { sum: 0, w: 0 },
    roe: { sum: 0, w: 0 },
    debt: { sum: 0, w: 0 },
    beta: { sum: 0, w: 0 }
  };
  var i;
  var r;
  var name;
  var ticker;
  var invested;
  var market;
  var plRub;
  var plPct;
  var sector;
  var country;
  var weight;

  for (i = 0; i < loaded.rows.length; i++) {
    r = loaded.rows[i];
    if (!_rowHasAnyValue_(r)) continue;

    ticker = _firstNonEmpty_([
      _valueByCol_(r, cols.ticker),
      ''
    ]);
    name = _firstNonEmpty_([
      _valueByCol_(r, cols.name),
      ticker,
      '—'
    ]);
    invested = _toNumberSafe_(_valueByCol_(r, cols.invested));
    market = _toNumberSafe_(_valueByCol_(r, cols.market));
    plRub = _toNumberSafe_(_valueByCol_(r, cols.plRub));
    plPct = _valueByCol_(r, cols.plPct);
    plPct = plPct === '' || plPct == null ? null : _toNumberSafe_(plPct);
    if (plPct == null) plPct = invested ? plRub / invested : null;
    sector = _firstNonEmpty_([_valueByCol_(r, cols.sector), '—']);
    country = _firstNonEmpty_([_valueByCol_(r, cols.country), '—']);

    totals.count += 1;
    totals.invested += invested;
    totals.market += market;
    totals.plRub += plRub;

    positions.push({
      name: String(name),
      ticker: String(ticker),
      invested: invested,
      market: market,
      plRub: plRub,
      plPct: plPct,
      share: null
    });

    _accumulateGroup_(sectors, sector, invested, market, plRub);
    _accumulateGroup_(countries, country, invested, market, plRub);

    weight = market > 0 ? market : 0;
    _accumulateWeightedMetric_(valuationAgg.pe, _valueByCol_(r, cols.pe), weight, false);
    _accumulateWeightedMetric_(valuationAgg.ps, _valueByCol_(r, cols.ps), weight, false);
    _accumulateWeightedMetric_(valuationAgg.pb, _valueByCol_(r, cols.pb), weight, false);
    _accumulateWeightedMetric_(valuationAgg.ev, _valueByCol_(r, cols.ev), weight, false);
    _accumulateWeightedMetric_(valuationAgg.roe, _valueByCol_(r, cols.roe), weight, false);
    _accumulateWeightedMetric_(valuationAgg.debt, _valueByCol_(r, cols.debt), weight, false);
    _accumulateWeightedMetric_(valuationAgg.beta, _valueByCol_(r, cols.beta), weight, false);
  }

  totals.plPct = totals.invested ? totals.plRub / totals.invested : null;
  _finalizePositionShares_(positions, totals.market);

  return {
    hasData: totals.count > 0,
    totals: totals,
    topPosition: _takeTop_(positions, 'market', 5, true),
    topProfit: _takeTop_(positions, 'plRub', 5, true),
    topLoss: _takeTop_(positions, 'plRub', 5, false),
    sectors: _finalizeGroups_(sectors, totals.market),
    countries: _finalizeGroups_(countries, totals.market),
    valuation: _buildValuationRowsFromAgg_(valuationAgg)
  };
}

function _mergeSharesComputedWithFallback_(computed, loaded, fallback, data) {
  var totals = _extractSharesTotals_(computed, fallback.totals);
  var topPosition = _extractSharesArray_(computed, [
    'topPosition', 'topPositions', 'topByPosition', 'top5ByPosition', 'largestPositions'
  ]);
  var topProfit = _extractSharesArray_(computed, [
    'topProfit', 'topPl', 'topByPl', 'bestPl', 'bestPL', 'top5Pl'
  ]);
  var topLoss = _extractSharesArray_(computed, [
    'topLoss', 'worstPl', 'worstPL', 'worstByPl', 'top5Loss'
  ]);
  var sectors = _extractSharesArray_(computed, [
    'sectors', 'bySector', 'sectorSummary', 'sectorRows'
  ]);
  var countries = _extractSharesArray_(computed, [
    'countries', 'byCountry', 'countrySummary', 'countryRows'
  ]);
  var valuation = _extractValuationRows_(computed);

  return {
    hasData: totals.count > 0,
    totals: totals,
    topPosition: topPosition && topPosition.length ? _normalizeTopPositionArray_(topPosition, totals.market) : fallback.topPosition,
    topProfit: topProfit && topProfit.length ? _normalizeTopPlArray_(topProfit) : fallback.topProfit,
    topLoss: topLoss && topLoss.length ? _normalizeTopPlArray_(topLoss) : fallback.topLoss,
    sectors: sectors && sectors.length ? _normalizeGroupArray_(sectors, totals.market) : fallback.sectors,
    countries: countries && countries.length ? _normalizeGroupArray_(countries, totals.market) : fallback.countries,
    valuation: valuation && valuation.length ? valuation : fallback.valuation
  };
}

function _normalizeLoadedSheetData_(loaded, sheet) {
  var data;
  var out = { header: [], rows: [] };

  if (loaded && loaded.header && loaded.rows) {
    out.header = loaded.header || [];
    out.rows = loaded.rows || [];
    return out;
  }

  if (loaded && loaded.headers && loaded.rows) {
    out.header = loaded.headers || [];
    out.rows = loaded.rows || [];
    return out;
  }

  if (_isArray_(loaded) && loaded.length && _isArray_(loaded[0])) {
    out.header = loaded[0] || [];
    out.rows = loaded.slice(1);
    return out;
  }

  data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  out.header = data && data.length ? data[0] : [];
  out.rows = data && data.length > 1 ? data.slice(1) : [];
  return out;
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
    pe: _findHeaderIndexByAliases_(header, ['P/E', 'P/E (TTM)', 'PE', 'PE TTM', 'peRatioTtm']),
    ps: _findHeaderIndexByAliases_(header, ['P/S', 'P/S (TTM)', 'PS', 'PS TTM', 'priceToSalesTtm']),
    pb: _findHeaderIndexByAliases_(header, ['P/B', 'PB', 'priceToBookTtm']),
    ev: _findHeaderIndexByAliases_(header, ['EV/EBITDA', 'evToEbitda']),
    roe: _findHeaderIndexByAliases_(header, ['ROE', 'roe']),
    debt: _findHeaderIndexByAliases_(header, ['Debt/Equity', 'debtToEquity']),
    beta: _findHeaderIndexByAliases_(header, ['Beta', 'beta'])
  };
}

function _findHeaderIndexByAliases_(header, aliases) {
  var normalized = [];
  var i;
  var j;
  var h;
  var a;

  for (i = 0; i < header.length; i++) normalized.push(_normalizeHeaderText_(header[i]));

  for (j = 0; j < aliases.length; j++) {
    a = _normalizeHeaderText_(aliases[j]);
    for (i = 0; i < normalized.length; i++) {
      h = normalized[i];
      if (h === a) return i;
    }
  }

  for (j = 0; j < aliases.length; j++) {
    a = _normalizeHeaderText_(aliases[j]);
    for (i = 0; i < normalized.length; i++) {
      h = normalized[i];
      if (h && a && h.indexOf(a) >= 0) return i;
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

function _isArray_(x) {
  return Object.prototype.toString.call(x) === '[object Array]';
}

function _valueByCol_(row, idx) {
  if (idx == null || idx < 0) return '';
  return row[idx];
}

function _firstNonEmpty_(arr) {
  var i;
  for (i = 0; i < arr.length; i++) {
    if (arr[i] !== '' && arr[i] != null) return arr[i];
  }
  return '';
}

function _accumulateGroup_(bucket, key, invested, market, plRub) {
  var k = String(key == null || key === '' ? '—' : key);
  if (!bucket[k]) bucket[k] = { name: k, count: 0, invested: 0, market: 0, plRub: 0, plPct: null, share: null };
  bucket[k].count += 1;
  bucket[k].invested += invested || 0;
  bucket[k].market += market || 0;
  bucket[k].plRub += plRub || 0;
}

function _finalizeGroups_(bucket, totalMarket) {
  var out = [];
  var k;
  for (k in bucket) {
    if (!bucket.hasOwnProperty(k)) continue;
    bucket[k].plPct = bucket[k].invested ? bucket[k].plRub / bucket[k].invested : null;
    bucket[k].share = totalMarket ? bucket[k].market / totalMarket : null;
    out.push(bucket[k]);
  }
  out.sort(function (a, b) { return (b.market || 0) - (a.market || 0); });
  return out;
}

function _finalizePositionShares_(positions, totalMarket) {
  var i;
  for (i = 0; i < positions.length; i++) {
    positions[i].share = totalMarket ? (positions[i].market || 0) / totalMarket : null;
  }
}

function _takeTop_(arr, field, limit, desc) {
  var out = (arr || []).slice();
  out.sort(function (a, b) {
    var av = _toNumberSafe_(a[field]);
    var bv = _toNumberSafe_(b[field]);
    return desc ? (bv - av) : (av - bv);
  });
  return out.slice(0, limit || 5);
}

function _accumulateWeightedMetric_(agg, value, weight, allowNegative) {
  var n = _toNumberSafe_(value);
  if (!weight || isNaN(n)) return;
  if (!allowNegative && n <= 0) return;
  agg.sum += n * weight;
  agg.w += weight;
}

function _buildValuationRowsFromAgg_(agg) {
  return [
    ['P/E (TTM)', agg.pe.w ? agg.pe.sum / agg.pe.w : 'n/a'],
    ['P/S (TTM)', agg.ps.w ? agg.ps.sum / agg.ps.w : 'n/a'],
    ['P/B', agg.pb.w ? agg.pb.sum / agg.pb.w : 'n/a'],
    ['EV/EBITDA', agg.ev.w ? agg.ev.sum / agg.ev.w : 'n/a'],
    ['ROE', agg.roe.w ? agg.roe.sum / agg.roe.w : 'n/a'],
    ['Debt/Equity', agg.debt.w ? agg.debt.sum / agg.debt.w : 'n/a'],
    ['Beta', agg.beta.w ? agg.beta.sum / agg.beta.w : 'n/a']
  ];
}

function _tryCallVariants_(fn, variants) {
  var i;
  if (typeof fn !== 'function') return null;
  for (i = 0; i < variants.length; i++) {
    try {
      return fn.apply(null, variants[i]);
    } catch (e) {}
  }
  return null;
}

function _extractSharesTotals_(computed, fallbackTotals) {
  var t = _pickFirstExisting_(computed, ['totals', 'total', 'summary', 'kpi', 'shareTotals']);
  var invested;
  var market;
  var plRub;
  var count;
  var plPct;

  if (!t) return fallbackTotals;

  invested = _pickNumber_(t, ['invested', 'totalInvested', 'sumInvested', 'investedRub']);
  market = _pickNumber_(t, ['market', 'marketValue', 'totalMarket', 'marketRub']);
  plRub = _pickNumber_(t, ['plRub', 'pl', 'profitRub', 'pnlRub']);
  count = _pickNumber_(t, ['count', 'items', 'positionsCount', 'paperCount']);
  plPct = _pickNullableNumber_(t, ['plPct', 'profitPct', 'pnlPct']);

  if (!count) count = fallbackTotals.count;
  if (plPct == null) plPct = invested ? plRub / invested : null;

  return {
    invested: invested,
    market: market,
    plRub: plRub,
    plPct: plPct,
    count: count
  };
}

function _extractSharesArray_(computed, keys) {
  var i;
  var arr;
  if (!computed) return null;
  for (i = 0; i < keys.length; i++) {
    arr = computed[keys[i]];
    if (_isArray_(arr)) return arr;
  }
  return null;
}

function _extractValuationRows_(computed) {
  var v = _pickFirstExisting_(computed, ['valuation', 'valuationSnapshot', 'snapshot', 'multiples', 'valuationRows']);
  var rows = [];
  var i;
  var keys;
  var key;

  if (!v) return null;

  if (_isArray_(v)) {
    for (i = 0; i < v.length; i++) {
      if (_isArray_(v[i])) {
        rows.push([v[i][0], v[i][1]]);
      } else if (typeof v[i] === 'object' && v[i]) {
        rows.push([_firstNonEmpty_([v[i].metric, v[i].name, 'Метрика']), _firstNonEmpty_([v[i].value, v[i].val, ''])]);
      }
    }
    return rows;
  }

  keys = ['pe', 'peTtm', 'ps', 'psTtm', 'pb', 'evToEbitda', 'roe', 'debtToEquity', 'beta'];
  for (i = 0; i < keys.length; i++) {
    key = keys[i];
    if (v[key] != null) rows.push([key, v[key]]);
  }
  return rows.length ? rows : null;
}

function _pickFirstExisting_(obj, keys) {
  var i;
  if (!obj) return null;
  for (i = 0; i < keys.length; i++) {
    if (obj[keys[i]] != null) return obj[keys[i]];
  }
  return null;
}

function _pickNumber_(obj, keys) {
  var i;
  for (i = 0; i < keys.length; i++) {
    if (obj && obj[keys[i]] != null && obj[keys[i]] !== '') return _toNumberSafe_(obj[keys[i]]);
  }
  return 0;
}

function _pickNullableNumber_(obj, keys) {
  var i;
  for (i = 0; i < keys.length; i++) {
    if (obj && obj[keys[i]] != null && obj[keys[i]] !== '') return _toNumberSafe_(obj[keys[i]]);
  }
  return null;
}

function _normalizeTopPositionArray_(arr, totalMarket) {
  var out = [];
  var i;
  var x;
  var market;
  var plRub;
  for (i = 0; i < arr.length; i++) {
    x = arr[i] || {};
    market = _pickNumber_(x, ['market', 'marketValue', 'position', 'value']);
    plRub = _pickNumber_(x, ['plRub', 'pl', 'profitRub', 'pnlRub']);
    out.push({
      name: String(_firstNonEmpty_([x.name, x.title, x.companyName, x.instrument, x.ticker, x.symbol, '—'])),
      ticker: String(_firstNonEmpty_([x.ticker, x.symbol, x.code, ''])),
      market: market,
      plRub: plRub,
      share: x.share != null ? _toNumberSafe_(x.share) : (totalMarket ? market / totalMarket : null),
      plPct: _pickNullableNumber_(x, ['plPct', 'profitPct', 'pnlPct'])
    });
  }
  out.sort(function (a, b) { return (b.market || 0) - (a.market || 0); });
  return out.slice(0, 5);
}

function _normalizeTopPlArray_(arr) {
  var out = [];
  var i;
  var x;
  for (i = 0; i < arr.length; i++) {
    x = arr[i] || {};
    out.push({
      name: String(_firstNonEmpty_([x.name, x.title, x.companyName, x.instrument, x.ticker, x.symbol, '—'])),
      ticker: String(_firstNonEmpty_([x.ticker, x.symbol, x.code, ''])),
      market: _pickNumber_(x, ['market', 'marketValue', 'position', 'value']),
      plRub: _pickNumber_(x, ['plRub', 'pl', 'profitRub', 'pnlRub', 'value']),
      plPct: _pickNullableNumber_(x, ['plPct', 'profitPct', 'pnlPct'])
    });
  }
  return out.slice(0, 5);
}

function _normalizeGroupArray_(arr, totalMarket) {
  var out = [];
  var i;
  var x;
  var market;
  var invested;
  var plRub;
  for (i = 0; i < arr.length; i++) {
    x = arr[i] || {};
    market = _pickNumber_(x, ['market', 'marketValue', 'value']);
    invested = _pickNumber_(x, ['invested', 'totalInvested']);
    plRub = _pickNumber_(x, ['plRub', 'pl', 'profitRub', 'pnlRub']);
    out.push({
      name: String(_firstNonEmpty_([x.name, x.sec, x.sector, x.country, x.title, '—'])),
      count: _pickNumber_(x, ['count', 'items', 'papers']),
      invested: invested,
      market: market,
      plRub: plRub,
      plPct: x.plPct != null ? _toNumberSafe_(x.plPct) : (invested ? plRub / invested : null),
      share: x.share != null ? _toNumberSafe_(x.share) : (totalMarket ? market / totalMarket : null)
    });
  }
  out.sort(function (a, b) { return (b.market || 0) - (a.market || 0); });
  return out;
}

function _mapTopPositionRows_(arr) {
  var rows = [];
  var i;
  if (!arr || !arr.length) return [['Нет данных', '', '', '', '']];
  for (i = 0; i < arr.length; i++) {
    rows.push([
      arr[i].name,
      arr[i].ticker || '',
      arr[i].market,
      arr[i].plRub,
      arr[i].plPct
    ]);
  }
  return rows;
}

function _mapTopPlRows_(arr) {
  var rows = [];
  var i;
  if (!arr || !arr.length) return [['Нет данных', '', '', '', '']];
  for (i = 0; i < arr.length; i++) {
    rows.push([
      arr[i].name,
      arr[i].ticker || '',
      arr[i].market,
      arr[i].plRub,
      arr[i].plPct
    ]);
  }
  return rows;
}

function _mapGroupRows_(arr) {
  var rows = [];
  var i;
  if (!arr || !arr.length) return [['Нет данных', '', '', '', '', '', '']];
  for (i = 0; i < arr.length; i++) {
    rows.push([
      arr[i].name,
      arr[i].count,
      arr[i].invested,
      arr[i].market,
      arr[i].plRub,
      arr[i].plPct,
      arr[i].share
    ]);
  }
  return rows;
}"
