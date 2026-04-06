/**
 * dashboard_bonds_region.gs
 * Регион Dashboard: Облигации
 * Отдельная отрисовка только внутри выделенного региона Dashboard.
 *
 * Контракт региона по умолчанию:
 *  - key = bonds
 *  - startCol = 23
 *  - width = 10
 *  - startRow = 1
 *  - maxRows = 115
 */

function renderBondsDashboardRegion_(sh, region) {
  region = _normalizeBondsRegion_(region);

  var baseRow = region.startRow - 1;
  var rowTitle = baseRow + 1;
  var rowTotals = baseRow + 3;
  var rowTopYtm = baseRow + 10;
  var rowTopPl = baseRow + 19;
  var rowWorstPl = baseRow + 28;
  var rowSectors = baseRow + 37;
  var rowMaturity = baseRow + 47;
  var rowCoupons = baseRow + 56;
  var rowChart1 = baseRow + 66;
  var rowChart2 = baseRow + 86;

  _removeChartsInBondsRegion_(sh, region);
  _clearBondsRegion_(sh, region);
  _writeBondsRegionTitle_(sh, region, 'Облигации');

  var data = _loadBondsRegionData_();
  if (!data.rows.length) {
    sh.getRange(rowTotals, region.startCol, 1, 2)
      .setValues([['Статус', 'Нет данных в Bonds']])
      .setFontWeight('bold');
    return;
  }

  var c = _mapBondsRegionColumns_(data.header);

  var totals = _computeBondsTotals_(data.rows, c);
  var topYtm = _computeBondsTopYtm_(data.rows, c, new Date());
  var topPl = _computeBondsTopPl_(data.rows, c, false);
  var worstPl = _computeBondsTopPl_(data.rows, c, true);
  var sectors = _computeBondsSectorSummary_(data.rows, c);
  var maturity = _computeBondsMaturitySummary_(data.rows, c, new Date());
  var coupon6m = _computeBondsCoupon6mSummary_(data.rows, c, new Date());

  _writeBondsTwoColBlock_(
    sh,
    rowTotals,
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

  _writeBondsTableBlock_(
    sh,
    rowTopYtm,
    region.startCol,
    'Облигации — top YTM',
    ['Бумага', 'YTM (%)', 'До погашения (лет)', 'FIGI'],
    topYtm.map(function (x) {
      return [x.name, x.ytm, x.years, x.figi];
    }),
    {
      dataFormats: {
        2: '0.00%',
        3: '0.00'
      },
      maxDataRows: 5
    }
  );

  _writeBondsTableBlock_(
    sh,
    rowTopPl,
    region.startCol,
    'Облигации — top P/L',
    ['Бумага', 'P/L (руб)', 'P/L (%)'],
    topPl.map(function (x) {
      return [x.name, x.plRub, x.plPct];
    }),
    {
      dataFormats: {
        2: '#,##0.00',
        3: '0.00%'
      },
      maxDataRows: 5
    }
  );

  _writeBondsTableBlock_(
    sh,
    rowWorstPl,
    region.startCol,
    'Облигации — worst P/L',
    ['Бумага', 'P/L (руб)', 'P/L (%)'],
    worstPl.map(function (x) {
      return [x.name, x.plRub, x.plPct];
    }),
    {
      dataFormats: {
        2: '#,##0.00',
        3: '0.00%'
      },
      maxDataRows: 5
    }
  );

  _writeBondsTableBlock_(
    sh,
    rowSectors,
    region.startCol,
    'Облигации — по секторам',
    ['Сектор', 'Рыночная стоимость', 'P/L (руб)'],
    sectors.map(function (x) {
      return [x.sector, x.market, x.plRub];
    }),
    {
      dataFormats: {
        2: '#,##0.00',
        3: '#,##0.00'
      },
      maxDataRows: 8
    }
  );

  var maturityBlock = _writeBondsTableBlock_(
    sh,
    rowMaturity,
    region.startCol,
    'Облигации — по сроку',
    ['Год', 'Рыночная стоимость'],
    maturity.map(function (x) {
      return [x.year, x.market];
    }),
    {
      dataFormats: {
        2: '#,##0.00'
      },
      maxDataRows: 7
    }
  );

  var couponBlock = _writeBondsTableBlock_(
    sh,
    rowCoupons,
    region.startCol,
    'Облигации — купоны 6м',
    ['Месяц', 'Выплаты'],
    coupon6m.map(function (x) {
      return [x.month, x.amount];
    }),
    {
      dataFormats: {
        2: '#,##0.00'
      },
      maxDataRows: 8
    }
  );

  if (maturityBlock && maturityBlock.dataRowsCount > 0) {
    _buildBondsChartFromRange_(
      sh,
      maturityBlock.chartRange,
      Charts.ChartType.COLUMN,
      rowChart1,
      region.startCol,
      'Облигации — сроки погашения'
    );
  }

  if (couponBlock && couponBlock.dataRowsCount > 0) {
    _buildBondsChartFromRange_(
      sh,
      couponBlock.chartRange,
      Charts.ChartType.LINE,
      rowChart2,
      region.startCol,
      'Облигации — купоны 6м'
    );
  }
}

/* =========================
 * Data read / map
 * ========================= */

function _loadBondsRegionData_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Bonds');
  if (!sh) return { header: [], rows: [] };

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { header: [], rows: [] };

  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return { header: header, rows: rows };
}

function _mapBondsRegionColumns_(hdr) {
  function idx() {
    for (var i = 0; i < arguments.length; i++) {
      var name = arguments[i];
      var k = hdr.indexOf(name);
      if (k >= 0) return k + 1;
    }
    return 0;
  }

  return {
    name: idx('Название'),
    figi: idx('FIGI'),
    sector: idx('Сектор'),
    qty: idx('Кол-во'),
    price: idx('Текущая цена'),
    nominal: idx('Номинал'),
    couponPerYear: idx('купон/год'),
    couponValue: idx('Размер купона'),
    maturity: idx('Дата погашения'),
    nextCoupon: idx('Следующий купон'),
    market: idx('Рыночная стоимость'),
    invested: idx('Инвестировано'),
    plRub: idx('P/L (руб)'),
    plPct: idx('P/L (%)')
  };
}

/* =========================
 * Calculations
 * ========================= */

function _computeBondsTotals_(rows, c) {
  var invested = 0;
  var market = 0;
  var plRub = 0;

  rows.forEach(function (r) {
    var inv = _bondNumber_(r, c.invested);
    var mkt = _bondNumber_(r, c.market);
    var pl = _bondNumber_(r, c.plRub);

    if (!isFinite(inv) || !isFinite(mkt) || !isFinite(pl)) return;

    invested += inv;
    market += mkt;
    plRub += pl;
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

    // Приближённая формула YTM в долях.
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

    if (!isFinite(market) || !isFinite(plRub)) return;

    if (!map[sector]) {
      map[sector] = { sector: sector, market: 0, plRub: 0 };
    }

    map[sector].market += market;
    map[sector].plRub += plRub;
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
  for (var i = 0; i < 6; i++) {
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

/* =========================
 * Region render helpers
 * ========================= */

function _removeChartsInBondsRegion_(sh, region) {
  var startRow = region.startRow;
  var endRow = region.startRow + region.maxRows - 1;
  var startCol = region.startCol;
  var endCol = region.startCol + region.width - 1;

  sh.getCharts().forEach(function (chart) {
    try {
      var info = chart.getContainerInfo();
      var row = info.getAnchorRow();
      var col = info.getAnchorColumn();

      if (row >= startRow && row <= endRow && col >= startCol && col <= endCol) {
        sh.removeChart(chart);
      }
    } catch (e) {}
  });
}

function _clearBondsRegion_(sh, region) {
  var rg = sh.getRange(region.startRow, region.startCol, region.maxRows, region.width);
  try { rg.breakApart(); } catch (e) {}
  rg.clearContent();
  rg.clearFormat();
  rg.clearDataValidations();
  try { rg.setWrap(true); } catch (e2) {}
}

function _writeBondsRegionTitle_(sh, region, title) {
  if (typeof dashboardWriteRegionTitle_ === 'function') {
    try {
      var titleRegion = {
        key: region.key,
        startCol: region.startCol,
        width: region.width,
        startRow: region.startRow,
        maxRows: region.maxRows,
        title: title,
        label: title
      };
      dashboardWriteRegionTitle_(sh, titleRegion);
      return;
    } catch (e) {}
  }

  var rg = sh.getRange(region.startRow, region.startCol, 1, region.width);
  try { rg.breakApart(); } catch (e2) {}
  rg.merge();
  rg
    .setValue(title)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#E5E7EB');
}

function _writeBondsTwoColBlock_(sh, startRow, startCol, title, rows, valueFormats) {
  sh.getRange(startRow, startCol, 1, 2)
    .setValues([[title, '']])
    .setFontWeight('bold');

  if (!rows || !rows.length) return null;

  var dataRange = sh.getRange(startRow + 1, startCol, rows.length, 2);
  dataRange.setValues(rows);
  sh.getRange(startRow + 1, startCol, rows.length, 1).setFontWeight('bold');

  if (valueFormats && valueFormats.length) {
    for (var i = 0; i < valueFormats.length; i++) {
      if (!valueFormats[i]) continue;
      sh.getRange(startRow + 1 + i, startCol + 1).setNumberFormat(valueFormats[i]);
    }
  }

  return {
    titleRow: startRow,
    dataStartRow: startRow + 1,
    dataRowsCount: rows.length
  };
}

function _writeBondsTableBlock_(sh, startRow, startCol, title, headers, rows, opts) {
  opts = opts || {};
  var safeRows = (rows || []).slice(0, opts.maxDataRows || rows.length);
  var dataFormats = opts.dataFormats || {};

  sh.getRange(startRow, startCol, 1, headers.length)
    .setValues([headers.map(function (_, i) { return i === 0 ? title : ''; })])
    .setFontWeight('bold');

  sh.getRange(startRow + 1, startCol, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold');

  if (safeRows.length) {
    sh.getRange(startRow + 2, startCol, safeRows.length, headers.length).setValues(safeRows);

    Object.keys(dataFormats).forEach(function (k) {
      var colIndex = Number(k);
      if (!isFinite(colIndex) || colIndex < 1 || colIndex > headers.length) return;
      sh.getRange(startRow + 2, startCol + colIndex - 1, safeRows.length, 1).setNumberFormat(dataFormats[k]);
    });
  }

  return {
    titleRow: startRow,
    headerRow: startRow + 1,
    dataStartRow: startRow + 2,
    dataRowsCount: safeRows.length,
    chartRange: sh.getRange(startRow + 1, startCol, 1 + safeRows.length, headers.length)
  };
}

function _buildBondsChartFromRange_(sh, range, chartType, anchorRow, anchorCol, title) {
  if (!range || range.getNumRows() < 2) return null;

  var chart = sh.newChart()
    .setChartType(chartType)
    .addRange(range)
    .setPosition(anchorRow, anchorCol, 0, 0)
    .setOption('title', title)
    .setOption('legend', { position: 'none' })
    .setOption('width', 420)
    .setOption('height', 220)
    .build();

  sh.insertChart(chart);
  return chart;
}

/* =========================
 * Primitive helpers
 * ========================= */

function _normalizeBondsRegion_(region) {
  region = region || {};
  return {
    key: region.key || 'bonds',
    startCol: Number(region.startCol) || 23,
    width: Number(region.width) || 10,
    startRow: Number(region.startRow) || 1,
    maxRows: Number(region.maxRows) || 115
  };
}

function _bondValue_(r, ci) {
  return ci ? r[ci - 1] : '';
}

function _bondText_(r, ci) {
  var v = _bondValue_(r, ci);
  var s = String(v == null ? '' : v).trim();
  return s;
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
