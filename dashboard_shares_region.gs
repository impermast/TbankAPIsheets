/**
 * dashboard_shares_region.gs
 * Регион Dashboard: Акции
 *
 * Контракт региона:
 * - key = shares
 * - startCol = 12
 * - width = 10
 * - startRow = 1
 * - maxRows = 105
 */

function renderSharesDashboardRegion_(sh, region) {
  var ss = SpreadsheetApp.getActive();
  var sharesSh = ss.getSheetByName('Shares');
  var regionRange = sh.getRange(region.startRow, region.startCol, region.maxRows, region.width);

  try { regionRange.breakApart(); } catch (e) {}
  regionRange.clearContent();
  _removeChartsInRegion_(sh, region);

  if (!sharesSh) {
    dashboardWritePlaceholder_(sh, region, 'Нет данных по акциям');
    return;
  }

  if (
    typeof loadSharesDashboardData_ !== 'function' ||
    typeof mapSharesDashboardColumns_ !== 'function' ||
    typeof validateSharesDashboardColumns_ !== 'function' ||
    typeof computeSharesDashboardData_ !== 'function'
  ) {
    dashboardWritePlaceholder_(sh, region, 'Нет данных по акциям');
    return;
  }

  var raw = loadSharesDashboardData_();
  var header = (raw && raw.header) || [];
  var rows = (raw && raw.rows) || [];

  if (!header.length || !rows.length) {
    dashboardWritePlaceholder_(sh, region, 'Нет данных по акциям');
    return;
  }

  var c = mapSharesDashboardColumns_(header);
  if (!validateSharesDashboardColumns_(c)) {
    dashboardWritePlaceholder_(sh, region, 'Нет данных по акциям');
    return;
  }

  var data = computeSharesDashboardData_(rows, c);
  if (!data || !data.count) {
    dashboardWritePlaceholder_(sh, region, 'Нет данных по акциям');
    return;
  }

  dashboardWriteRegionTitle_(sh, region);

  var startCol = region.startCol;

  // B) row 3: Акции — итого
  _writeTwoColBlock_(
    sh,
    3,
    startCol,
    'Акции — итого',
    [
      ['Инвестировано', data.invested],
      ['Рыночная стоимость', data.market],
      ['P/L (руб)', data.plRub],
      ['P/L (%)', data.plPct]
    ],
    ['money', 'money', 'money', 'percent']
  );

  // C) row 10: Акции — топ-5 по позиции
  var topPositionsRows = (data.topPositions || []).slice(0, 5).map(function(item) {
    return [
      item.name || '—',
      item.ticker || '—',
      item.market,
      item.plRub,
      item.plPct
    ];
  });

  _writeTableBlock_(
    sh,
    10,
    startCol,
    'Акции — топ-5 по позиции',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    topPositionsRows,
    {
      moneyCols: [3, 4],
      percentCols: [5]
    }
  );

  // D) row 19: Акции — топ-5 по P/L
  var bestPLRows = (data.bestPL || []).slice(0, 5).map(function(item) {
    return [
      item.name || '—',
      item.ticker || '—',
      item.market,
      item.plRub,
      item.plPct
    ];
  });

  _writeTableBlock_(
    sh,
    19,
    startCol,
    'Акции — топ-5 по P/L',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    bestPLRows,
    {
      moneyCols: [3, 4],
      percentCols: [5]
    }
  );

  // E) row 28: Акции — топ-5 убыток
  var worstPLRows = (data.worstPL || []).slice(0, 5).map(function(item) {
    return [
      item.name || '—',
      item.ticker || '—',
      item.market,
      item.plRub,
      item.plPct
    ];
  });

  _writeTableBlock_(
    sh,
    28,
    startCol,
    'Акции — топ-5 убыток',
    ['Бумага', 'Тикер', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'],
    worstPLRows,
    {
      moneyCols: [3, 4],
      percentCols: [5]
    }
  );

  // F) row 37: Акции — по секторам
  var sectorsRows = (data.sectors || []).slice(0, 5).map(function(item) {
    return [
      item.name || '—',
      item.market,
      item.sharePct
    ];
  });

  _writeTableBlock_(
    sh,
    37,
    startCol,
    'Акции — по секторам',
    ['Сектор', 'Рыночная стоимость', 'Доля акций'],
    sectorsRows,
    {
      moneyCols: [2],
      percentCols: [3]
    }
  );

  // G) row 47: Акции — по странам
  var countriesRows = (data.countries || []).slice(0, 5).map(function(item) {
    return [
      item.name || '—',
      item.market,
      item.sharePct
    ];
  });

  _writeTableBlock_(
    sh,
    47,
    startCol,
    'Акции — по странам',
    ['Страна', 'Рыночная стоимость', 'Доля акций'],
    countriesRows,
    {
      moneyCols: [2],
      percentCols: [3]
    }
  );

  // H) row 57: Акции — valuation snapshot
  var valuationRows = [
    ['Median P/E', data.valuation && data.valuation.peMedian != null ? data.valuation.peMedian : ''],
    ['Median P/B', data.valuation && data.valuation.pbMedian != null ? data.valuation.pbMedian : ''],
    ['Median EV/EBITDA', data.valuation && data.valuation.evEbitdaMedian != null ? data.valuation.evEbitdaMedian : ''],
    ['Median ROE', data.valuation && data.valuation.roeMedian != null ? data.valuation.roeMedian : ''],
    ['Median Beta', data.valuation && data.valuation.betaMedian != null ? data.valuation.betaMedian : '']
  ];

  _writeTableBlock_(
    sh,
    57,
    startCol,
    'Акции — valuation snapshot',
    ['Метрика', 'Значение'],
    valuationRows,
    {}
  );

  // I) row 66: pie chart по секторам
  var sectorHeaderRow = 38;
  var sectorDataStartRow = 39;
  var sectorDataRowsCount = Math.max(1, sectorsRows.length);

  _buildChartFromRange_(
    sh,
    region,
    {
      title: 'Акции — структура по секторам',
      chartType: Charts.ChartType.PIE,
      row: 66,
      ranges: [
        sh.getRange(sectorHeaderRow, startCol, sectorDataRowsCount + 1, 2)
      ],
      width: 400,
      height: 220
    }
  );

  // J) row 86: bar chart топ-5 по позиции
  var topHeaderRow = 11;
  var topDataStartRow = 12;
  var topDataRowsCount = Math.max(1, topPositionsRows.length);

  _buildChartFromRange_(
    sh,
    region,
    {
      title: 'Акции — топ-5 по позиции',
      chartType: Charts.ChartType.BAR,
      row: 86,
      ranges: [
        sh.getRange(topHeaderRow, startCol, topDataRowsCount + 1, 1),
        sh.getRange(topHeaderRow, startCol + 2, topDataRowsCount + 1, 1)
      ],
      width: 400,
      height: 220
    }
  );
}

function _removeChartsInRegion_(sh, region) {
  var charts = sh.getCharts() || [];
  var startCol = region.startCol;
  var endCol = region.startCol + region.width - 1;
  var startRow = region.startRow;
  var endRow = region.startRow + region.maxRows - 1;

  charts.forEach(function(chart) {
    try {
      var info = chart.getContainerInfo();
      var col = info.getAnchorColumn();
      var row = info.getAnchorRow();

      if (col >= startCol && col <= endCol && row >= startRow && row <= endRow) {
        sh.removeChart(chart);
      }
    } catch (e) {}
  });
}

function _writeTwoColBlock_(sh, startRow, startCol, title, rows, valueFormats) {
  var width = 2;
  var titleRange = sh.getRange(startRow, startCol, 1, width);
  titleRange.setValues([[title, '']]);
  titleRange.setFontWeight('bold');

  var dataRows = rows || [];
  if (!dataRows.length) return;

  sh.getRange(startRow + 1, startCol, dataRows.length, width).setValues(dataRows);

  sh.getRange(startRow + 1, startCol, dataRows.length, 1).setFontWeight('bold');
  sh.getRange(startRow + 1, startCol + 1, dataRows.length, 1).setHorizontalAlignment('right');

  var formats = valueFormats || [];
  for (var i = 0; i < formats.length; i++) {
    var fmt = formats[i];
    var cell = sh.getRange(startRow + 1 + i, startCol + 1, 1, 1);

    if (fmt === 'money') cell.setNumberFormat('#,##0.00');
    if (fmt === 'percent') cell.setNumberFormat('0.00%');
  }
}

function _writeTableBlock_(sh, startRow, startCol, title, headers, rows, options) {
  var width = headers.length;
  var titleRange = sh.getRange(startRow, startCol, 1, width);
  var titleValues = [headers.map(function(_, idx) { return idx === 0 ? title : ''; })];
  titleRange.setValues(titleValues);
  titleRange.setFontWeight('bold');

  var headerRange = sh.getRange(startRow + 1, startCol, 1, width);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');

  var dataRows = rows || [];
  if (dataRows.length) {
    sh.getRange(startRow + 2, startCol, dataRows.length, width).setValues(dataRows);
  }

  var moneyCols = (options && options.moneyCols) || [];
  var percentCols = (options && options.percentCols) || [];

  moneyCols.forEach(function(colOffset) {
    if (dataRows.length) {
      sh.getRange(startRow + 2, startCol + colOffset - 1, dataRows.length, 1).setNumberFormat('#,##0.00');
    }
  });

  percentCols.forEach(function(colOffset) {
    if (dataRows.length) {
      sh.getRange(startRow + 2, startCol + colOffset - 1, dataRows.length, 1).setNumberFormat('0.00%');
    }
  });
}

function _buildChartFromRange_(sh, region, cfg) {
  var row = cfg.row;
  var chartType = cfg.chartType;
  var ranges = cfg.ranges || [];
  var width = Math.min(Number(cfg.width) || 400, 420);
  var height = Math.min(Number(cfg.height) || 220, 220);

  var builder = sh.newChart()
    .setChartType(chartType)
    .setPosition(row, region.startCol, 0, 0)
    .setOption('title', cfg.title || '')
    .setOption('width', width)
    .setOption('height', height);

  ranges.forEach(function(r) {
    builder.addRange(r);
  });

  if (chartType === Charts.ChartType.PIE) {
    builder
      .setOption('legend', { position: 'right' })
      .setOption('pieSliceText', 'value');
  }

  if (chartType === Charts.ChartType.BAR) {
    builder
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { format: '#,##0.00' })
      .setOption('bars', 'horizontal');
  }

  sh.insertChart(builder.build());
}
