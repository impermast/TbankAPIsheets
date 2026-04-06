/**
 * dashboard_funds_region.gs
 */

function renderFundsDashboardRegion_(sh, region) {
  region = _normalizeFundsRegion_(region);
  _clearFundsRegion_(sh, region);
  _removeChartsInFundsRegion_(sh, region);

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

  _writeFundsRegionTitle_(sh, region, 'Фонды');

  _writeFundsTwoColBlock_(sh, region, 3, 'Фонды — итого', [
    { label: 'Инвестировано',        value: totals.invested, format: '#,##0.00' },
    { label: 'Рыночная стоимость',   value: totals.market,   format: '#,##0.00' },
    { label: 'P/L (руб)',            value: totals.plRub,    format: '#,##0.00' },
    { label: 'P/L (%)',              value: totals.plPct,    format: '0.00%' }
  ]);

  var topByMarketInfo = _writeFundsTableBlock_(
    sh,
    region,
    10,
    'Фонды — top-5 по позиции',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    topByMarket.map(function(x) {
      return [x.name, x.ticker, x.market, x.plRub, x.plPct];
    }),
    ['@', '@', '#,##0.00', '#,##0.00', '0.00%']
  );

  var topByPlInfo = _writeFundsTableBlock_(
    sh,
    region,
    19,
    'Фонды — top-5 по P/L',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    topByPl.map(function(x) {
      return [x.name, x.ticker, x.market, x.plRub, x.plPct];
    }),
    ['@', '@', '#,##0.00', '#,##0.00', '0.00%']
  );

  var sectorInfo = _writeFundsTableBlock_(
    sh,
    region,
    28,
    'Фонды — по секторам',
    ['Сектор', 'Рыночная стоимость', 'Доля фондов'],
    bySector.map(function(x) {
      return [x.sector, x.market, x.share];
    }),
    ['@', '#,##0.00', '0.00%']
  );

  var currencyInfo = _writeFundsTableBlock_(
    sh,
    region,
    38,
    'Фонды — по валютам',
    ['Валюта', 'Рыночная стоимость', 'Доля фондов'],
    byCurrency.map(function(x) {
      return [x.currency, x.market, x.share];
    }),
    ['@', '#,##0.00', '0.00%']
  );

  var chartSource = null;
  if (sectorInfo && sectorInfo.dataRowCount > 0) {
    chartSource = {
      startRow: sectorInfo.dataStartRow,
      endRow: sectorInfo.dataEndRow,
      startCol: region.startCol,
      widthCols: 2
    };
  } else if (currencyInfo && currencyInfo.dataRowCount > 0) {
    chartSource = {
      startRow: currencyInfo.dataStartRow,
      endRow: currencyInfo.dataEndRow,
      startCol: region.startCol,
      widthCols: 2
    };
  }

  if (chartSource) {
    _buildFundsChartFromRange_(sh, region, {
      title: 'Фонды — структура',
      chartType: Charts.ChartType.PIE,
      anchorRow: 48,
      anchorCol: region.startCol,
      dataRange: sh.getRange(
        chartSource.startRow,
        chartSource.startCol,
        chartSource.endRow - chartSource.startRow + 1,
        chartSource.widthCols
      )
    });
  }
}

function _normalizeFundsRegion_(region) {
  region = region || {};
  return {
    startCol: Number(region.startCol) || 34,
    width: Number(region.width) || 10,
    startRow: Number(region.startRow) || 1,
    maxRows: Number(region.maxRows) || 80
  };
}

function _removeChartsInFundsRegion_(sh, region) {
  var rowMin = region.startRow;
  var rowMax = region.startRow + region.maxRows - 1;
  var colMin = region.startCol;
  var colMax = region.startCol + region.width - 1;

  var charts = sh.getCharts();
  for (var i = charts.length - 1; i >= 0; i--) {
    var info = charts[i].getContainerInfo();
    var r = info.getAnchorRow();
    var c = info.getAnchorColumn();
    if (r >= rowMin && r <= rowMax && c >= colMin && c <= colMax) {
      sh.removeChart(charts[i]);
    }
  }
}

function _clearFundsRegion_(sh, region) {
  var rng = sh.getRange(region.startRow, region.startCol, region.maxRows, region.width);
  rng.clearContent();
  rng.clearFormat();
  rng.clearNote();
  rng.clearDataValidations();
}

function _writeFundsRegionTitle_(sh, region, title) {
  if (typeof dashboardWriteRegionTitle_ === 'function') {
    dashboardWriteRegionTitle_(sh, region, title);
    return;
  }

  var rng = sh.getRange(region.startRow, region.startCol, 1, region.width);
  rng.merge();
  rng
    .setValue(title)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
}

function _loadFundsRegionData_() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName('Funds');
  if (!src) return { sheet: null, header: [], rows: [] };

  var lastRow = src.getLastRow();
  var lastCol = src.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { sheet: src, header: [], rows: [] };

  var header = src.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = src.getRange(2, 1, lastRow - 1, lastCol).getValues();

  rows = rows.filter(function(r) {
    if (typeof _rowHasAnyValue_ === 'function') return _rowHasAnyValue_(r);
    for (var i = 0; i < r.length; i++) {
      if (String(r[i] == null ? '' : r[i]).trim() !== '') return true;
    }
    return false;
  });

  return { sheet: src, header: header, rows: rows };
}

function _mapFundsRegionColumns_(header) {
  var aliasMap = {
    name: ['Название', 'Наименование', 'Инструмент', 'Компания'],
    ticker: ['Тикер', 'Ticker'],
    sector: ['Сектор', 'Категория', 'Тема', 'Фокус'],
    currency: ['Валюта', 'Currency'],
    invested: ['Инвестировано'],
    market: ['Рыночная стоимость', 'Стоимость', 'Рыночная стоимость, ₽'],
    plRub: ['P/L (руб)', 'P/L, руб', 'P/L RUB'],
    plPct: ['P/L (%)', 'P/L %']
  };

  function normalizeText(v) {
    if (typeof _normalizeHeaderText_ === 'function') return _normalizeHeaderText_(v);
    return String(v == null ? '' : v)
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .replace(/[ё]/g, 'е')
      .trim();
  }

  var normalizedHeader = header.map(normalizeText);

  function localFind(aliases) {
    var normAliases = aliases.map(normalizeText);
    for (var i = 0; i < normalizedHeader.length; i++) {
      if (normAliases.indexOf(normalizedHeader[i]) >= 0) return i;
    }
    return -1;
  }

  function findIndex(aliases) {
    var idx = -1;

    if (typeof _findHeaderIndexByAliases_ === 'function') {
      try {
        idx = _findHeaderIndexByAliases_(header, aliases);
      } catch (e1) {
        try {
          idx = _findHeaderIndexByAliases_(aliases, header);
        } catch (e2) {
          idx = -1;
        }
      }

      if (typeof idx === 'number' && isFinite(idx)) {
        var normAliases = aliases.map(normalizeText);

        if (idx >= 0 && idx < header.length && normAliases.indexOf(normalizeText(header[idx])) >= 0) {
          return idx;
        }
        if (idx > 0 && idx - 1 < header.length && normAliases.indexOf(normalizeText(header[idx - 1])) >= 0) {
          return idx - 1;
        }
      }
    }

    return localFind(aliases);
  }

  return {
    name: findIndex(aliasMap.name),
    ticker: findIndex(aliasMap.ticker),
    sector: findIndex(aliasMap.sector),
    currency: findIndex(aliasMap.currency),
    invested: findIndex(aliasMap.invested),
    market: findIndex(aliasMap.market),
    plRub: findIndex(aliasMap.plRub),
    plPct: findIndex(aliasMap.plPct)
  };
}

function _computeFundsTotals_(rows, cols) {
  function num(v) {
    if (typeof _toNumberSafe_ === 'function') return _toNumberSafe_(v);
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v == null ? '' : v).replace(/\s/g, '').replace(',', '.');
    var n = Number(s);
    return isFinite(n) ? n : 0;
  }

  function cell(row, idx) {
    return (idx >= 0 && idx < row.length) ? row[idx] : '';
  }

  var invested = 0;
  var market = 0;
  var plRub = 0;

  rows.forEach(function(row) {
    var inv = num(cell(row, cols.invested));
    var mkt = num(cell(row, cols.market));
    var pl = (cols.plRub >= 0) ? num(cell(row, cols.plRub)) : (mkt - inv);

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
  function num(v) {
    if (typeof _toNumberSafe_ === 'function') return _toNumberSafe_(v);
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v == null ? '' : v).replace(/\s/g, '').replace(',', '.');
    var n = Number(s);
    return isFinite(n) ? n : 0;
  }

  function txt(v) {
    return String(v == null ? '' : v).trim();
  }

  function cell(row, idx) {
    return (idx >= 0 && idx < row.length) ? row[idx] : '';
  }

  return rows.map(function(row) {
    var name = txt(cell(row, cols.name)) || '—';
    var ticker = txt(cell(row, cols.ticker)) || '—';
    var invested = num(cell(row, cols.invested));
    var market = num(cell(row, cols.market));
    var plRub = (cols.plRub >= 0) ? num(cell(row, cols.plRub)) : (market - invested);
    var plPct = invested !== 0 ? (plRub / invested) : 0;

    return {
      name: name,
      ticker: ticker,
      invested: invested,
      market: market,
      plRub: plRub,
      plPct: plPct
    };
  }).filter(function(x) {
    return x.name !== '—' || x.market !== 0 || x.plRub !== 0;
  }).sort(function(a, b) {
    return b.market - a.market;
  }).slice(0, limit || 5);
}

function _computeFundsTopByPl_(rows, cols, limit) {
  function num(v) {
    if (typeof _toNumberSafe_ === 'function') return _toNumberSafe_(v);
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v == null ? '' : v).replace(/\s/g, '').replace(',', '.');
    var n = Number(s);
    return isFinite(n) ? n : 0;
  }

  function txt(v) {
    return String(v == null ? '' : v).trim();
  }

  function cell(row, idx) {
    return (idx >= 0 && idx < row.length) ? row[idx] : '';
  }

  return rows.map(function(row) {
    var name = txt(cell(row, cols.name)) || '—';
    var ticker = txt(cell(row, cols.ticker)) || '—';
    var invested = num(cell(row, cols.invested));
    var market = num(cell(row, cols.market));
    var plRub = (cols.plRub >= 0) ? num(cell(row, cols.plRub)) : (market - invested);
    var plPct = invested !== 0 ? (plRub / invested) : 0;

    return {
      name: name,
      ticker: ticker,
      invested: invested,
      market: market,
      plRub: plRub,
      plPct: plPct
    };
  }).filter(function(x) {
    return x.name !== '—' || x.market !== 0 || x.plRub !== 0;
  }).sort(function(a, b) {
    return b.plRub - a.plRub;
  }).slice(0, limit || 5);
}

function _computeFundsSectorSummary_(rows, cols) {
  function num(v) {
    if (typeof _toNumberSafe_ === 'function') return _toNumberSafe_(v);
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v == null ? '' : v).replace(/\s/g, '').replace(',', '.');
    var n = Number(s);
    return isFinite(n) ? n : 0;
  }

  function txt(v) {
    return String(v == null ? '' : v).trim();
  }

  function cell(row, idx) {
    return (idx >= 0 && idx < row.length) ? row[idx] : '';
  }

  var totalMarket = 0;
  var map = {};

  rows.forEach(function(row) {
    var sector = txt(cell(row, cols.sector)) || '—';
    var market = num(cell(row, cols.market));

    totalMarket += market;
    map[sector] = (map[sector] || 0) + market;
  });

  return Object.keys(map).map(function(sector) {
    var market = map[sector];
    return {
      sector: sector,
      market: market,
      share: totalMarket !== 0 ? (market / totalMarket) : 0
    };
  }).sort(function(a, b) {
    return b.market - a.market;
  });
}

function _computeFundsCurrencySummary_(rows, cols) {
  function num(v) {
    if (typeof _toNumberSafe_ === 'function') return _toNumberSafe_(v);
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v == null ? '' : v).replace(/\s/g, '').replace(',', '.');
    var n = Number(s);
    return isFinite(n) ? n : 0;
  }

  function txt(v) {
    return String(v == null ? '' : v).trim();
  }

  function cell(row, idx) {
    return (idx >= 0 && idx < row.length) ? row[idx] : '';
  }

  var totalMarket = 0;
  var map = {};

  rows.forEach(function(row) {
    var currency = txt(cell(row, cols.currency)) || '—';
    var market = num(cell(row, cols.market));

    totalMarket += market;
    map[currency] = (map[currency] || 0) + market;
  });

  return Object.keys(map).map(function(currency) {
    var market = map[currency];
    return {
      currency: currency,
      market: market,
      share: totalMarket !== 0 ? (market / totalMarket) : 0
    };
  }).sort(function(a, b) {
    return b.market - a.market;
  });
}

function _writeFundsTwoColBlock_(sh, region, startRow, title, rows) {
  var col = region.startCol;
  var titleRange = sh.getRange(startRow, col, 1, 2);
  titleRange.merge();
  titleRange
    .setValue(title)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  if (!rows || !rows.length) {
    return { titleRow: startRow, dataStartRow: startRow + 1, dataEndRow: startRow, dataRowCount: 0 };
  }

  var values = rows.map(function(r) { return [r.label, r.value]; });
  var body = sh.getRange(startRow + 1, col, values.length, 2);
  body.setValues(values);

  sh.getRange(startRow + 1, col, values.length, 1).setFontWeight('bold');
  sh.getRange(startRow + 1, col, values.length, 2).setHorizontalAlignment('right');

  rows.forEach(function(r, i) {
    if (r.format) sh.getRange(startRow + 1 + i, col + 1).setNumberFormat(r.format);
  });

  return {
    titleRow: startRow,
    dataStartRow: startRow + 1,
    dataEndRow: startRow + rows.length,
    dataRowCount: rows.length
  };
}

function _writeFundsTableBlock_(sh, region, startRow, title, headers, rows, formats) {
  var col = region.startCol;
  var width = headers.length;

  var titleRange = sh.getRange(startRow, col, 1, width);
  titleRange.merge();
  titleRange
    .setValue(title)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  var headerRange = sh.getRange(startRow + 1, col, 1, width);
  headerRange
    .setValues([headers])
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  var dataRowCount = rows && rows.length ? rows.length : 0;
  if (dataRowCount > 0) {
    var dataRange = sh.getRange(startRow + 2, col, dataRowCount, width);
    dataRange.setValues(rows);

    for (var i = 0; i < formats.length; i++) {
      if (formats[i]) {
        sh.getRange(startRow + 2, col + i, dataRowCount, 1).setNumberFormat(formats[i]);
      }
    }
  }

  return {
    titleRow: startRow,
    headerRow: startRow + 1,
    dataStartRow: startRow + 2,
    dataEndRow: startRow + 1 + dataRowCount,
    dataRowCount: dataRowCount
  };
}

function _buildFundsChartFromRange_(sh, region, cfg) {
  if (!cfg || !cfg.dataRange) return null;

  var chart = sh.newChart()
    .setChartType(cfg.chartType || Charts.ChartType.PIE)
    .addRange(cfg.dataRange)
    .setOption('title', cfg.title || '')
    .setOption('width', 420)
    .setOption('height', 220)
    .setOption('legend', { position: 'right' })
    .setPosition(cfg.anchorRow || 48, cfg.anchorCol || region.startCol, 0, 0)
    .build();

  sh.insertChart(chart);
  return chart;
}
