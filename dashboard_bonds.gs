/**
 * bonds_dashboard.gs
 * Перестраивает лист Dashboard на основе листа Bonds.
 *
 * Блоки:
 *  - BondsDashboard: оркестратор
 *  - DashboardCtx: контекст/конфиг
 *  - BondsData: чтение/маппинг/валидация
 *  - BondsCalc: расчёты (base/derived)
 *  - DashRender: вывод таблиц/диаграмм/секций
 *  - HistoryStore: буфер A12:C28 (Дата/Инвестировано/Стоимость)
 */

function buildBondsDashboard() {
  BondsDashboard.build();
}

var BondsDashboard = (function () {

  function build() {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      showSnack_('Другая операция обновляет Dashboard. Повторите позже.', 'Dashboard', 3000);
      return;
    }

    try {
      var ctx = DashboardCtx.create_();
      if (!ctx.src) { showSnack_('Лист Bonds не найден', 'Dashboard', 3000); return; }

      // snapshot истории до очистки
      var histSnapshot = HistoryStore.snapshot_(ctx);

      ctx.dst.clear();

      // загрузка данных + схема колонок
      var data = BondsData.load_(ctx);
      if (!data.rows.length) { showSnack_('Нет данных для сводки', 'Dashboard', 2500); return; }

      var cols = BondsData.mapColumns_(data.header);
      if (!BondsData.validate_(cols)) {
        showSnack_('Не хватает обязательных колонок (Риск/Сектор/Рыночная/…)', 'Dashboard', 3500);
        return;
      }

      // расчёты
      var base = BondsCalc.computeBase_(data.rows, cols, ctx);
      var derived = BondsCalc.computeDerived_(base, data.rows, cols, ctx);

      // таблицы
      DashRender.renderTables_(ctx, base, derived);

      // история A12:C28 (до диаграмм)
      var histBlock = HistoryStore.update_(histSnapshot, base.invested, base.market, ctx);
      HistoryStore.write_(ctx, histBlock);

      // диаграммы
      DashRender.removeCharts_(ctx);
      DashRender.renderCharts_(ctx, base, derived);

      // секции/экстра
      DashRender.renderExtras_(ctx, base, derived);

      ctx.dst.autoResizeColumns(1, 26);

      showSnack_('Dashboard обновлён', 'Dashboard', 2000);
    } finally {
      lock.releaseLock();
    }
  }

  return { build: build };
})();

var DashboardCtx = (function () {

  function create_() {
    var ss = SpreadsheetApp.getActive();
    var src = ss.getSheetByName('Bonds');
    var dst = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');

    return {
      ss: ss,
      src: src,
      dst: dst,
      now: new Date(),
      cfg: getConfig_()
    };
  }

  function getConfig_() {
    return {
      history: { r1: 12, r2: 28, c1: 1, w: 3 }, // A12:C28
      blocks: {
        kpi:   { r: 1, c: 1 },
        risk:  { r: 1, c: 4 },
        sec:   { r: 1, c: 7 },
        year:  { r: 1, c: 10 },
        coup:  { r: 1, c: 13 },
        ext:   { c: 20 } // T
      },
      layout: {
        chartsTopRow: 20,
        safeStartRowAfterHistory: 30,
        historyChart: { rowOffset: 18, col: 1, heightRows: 18 },
        cmpBlockHeightRows: 18
      },
      paletteMain: ['#4F46E5','#22C55E','#EAB308','#EF4444','#06B6D4','#A855F7','#F59E0B','#94A3B8','#10B981','#3B82F6'],
      paletteRisk: ['#10B981','#EAB308','#EF4444']
    };
  }

  return { create_: create_ };
})();

var BondsData = (function () {

  function load_(ctx) {
    var src = ctx.src;
    var lastRow = src.getLastRow();
    var lastCol = src.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { header: [], rows: [] };

    var header = src.getRange(1, 1, 1, lastCol).getValues()[0];
    var rows = src.getRange(2, 1, lastRow - 1, lastCol).getValues();
    return { header: header, rows: rows };
  }

  function mapColumns_(hdr) {
    function idx(name) { var i = hdr.indexOf(name); return (i >= 0) ? (i + 1) : 0; }
    return {
      name: idx('Название'),
      figi: idx('FIGI'),
      riskNum: idx('Риск (ручн.)'),
      sector: idx('Сектор'),
      qty: idx('Кол-во'),
      price: idx('Текущая цена'),
      nominal: idx('Номинал'),
      cupPY: idx('купон/год'),
      cupVal: idx('Размер купона'),
      maturity: idx('Дата погашения'),
      nextCp: idx('Следующий купон'),

      mkt: idx('Рыночная стоимость'),
      inv: idx('Инвестировано'),
      pl: idx('P/L (руб)'),
      plPct: idx('P/L (%)'),

      couponPct: idx('Купон, %') // optional
    };
  }

  function validate_(c) {
    return !!(c.riskNum && c.sector && c.mkt && c.inv && c.pl && c.plPct && c.price);
  }

  return { load_: load_, mapColumns_: mapColumns_, validate_: validate_ };
})();

var BondsCalc = (function () {

  function computeBase_(rows, c, ctx) {
    var today = ctx.now;
    var horizon = addMonths_(today, 6);

    function strip(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
    function r2(x) { return Math.round(Number(x || 0) * 100) / 100; }

    var invested = 0, market = 0, plRub = 0;
    var sectors = {};
    var sectorPl = {};
    var riskCnt = { 'Низкий': 0, 'Средний': 0, 'Высокий': 0 };

    var wCouponNum = 0, wCouponDen = 0;
    var wYtmNum = 0, wYtmDen = 0;

    var scatterRiskYield = [['Риск','YTM (%)','Тултип']];
    var topYtmCandidates = [];

    var monthAgg = {};
    var byYear = {};

    rows.forEach(function (r) {
      function val(ci) { return (ci ? r[ci - 1] : ''); }

      var riskNum = Number(val(c.riskNum));
      if (!isNaN(riskNum)) {
        if (riskNum <= 1) riskCnt['Низкий']++;
        else if (riskNum <= 4) riskCnt['Средний']++;
        else riskCnt['Высокий']++;
      }

      var sec = String(val(c.sector) || 'other');
      var mkt = Number(val(c.mkt)) || 0;
      var inv = Number(val(c.inv)) || 0;
      var pl  = Number(val(c.pl))  || 0;

      invested += inv; market += mkt; plRub += pl;
      sectors[sec]  = (sectors[sec]  || 0) + mkt;
      sectorPl[sec] = (sectorPl[sec] || 0) + pl;

      if (c.maturity) {
        var mv = val(c.maturity);
        if (mv) {
          var md = (mv instanceof Date) ? mv : (isNaN(Date.parse(mv)) ? null : new Date(Date.parse(mv)));
          if (md) {
            var y = md.getFullYear();
            byYear[y] = (byYear[y] || 0) + mkt;
          }
        }
      }

      var coupPctAlt = NaN;
      if (c.couponPct) coupPctAlt = Number(val(c.couponPct));
      var cupPY  = Number(val(c.cupPY));
      var cupVal = Number(val(c.cupVal));
      var price  = Number(val(c.price));

      var coupPctCalc = NaN;
      if (price > 0 && cupPY > 0 && cupVal >= 0) {
        coupPctCalc = (cupVal * cupPY) / price * 100;
      }

      var coupUse = !isNaN(coupPctAlt) ? coupPctAlt : coupPctCalc;
      if (!isNaN(coupUse) && mkt > 0) {
        wCouponNum += (coupUse / 100) * mkt;
        wCouponDen += mkt;
      }

      var nominal = Number(val(c.nominal));
      if (!(nominal > 0)) nominal = 1000;

      var matStr = val(c.maturity);
      var yearsToMat = null;
      if (matStr) {
        var d = (matStr instanceof Date) ? matStr : (isNaN(Date.parse(matStr)) ? null : new Date(Date.parse(matStr)));
        if (d) yearsToMat = Math.max(0.25, (strip(d) - strip(today)) / (365 * 24 * 3600 * 1000));
      }

      var name = String(val(c.name) || '');
      var figi = String(val(c.figi) || '');

      var ytmPct = null;
      if (price > 0 && nominal > 0 && cupPY > 0 && cupVal >= 0 && yearsToMat) {
        var C = cupVal * cupPY;
        ytmPct = ((C + (nominal - price) / yearsToMat) / ((nominal + price) / 2)) * 100;

        if (isFinite(ytmPct)) {
          if (mkt > 0) { wYtmNum += (ytmPct / 100) * mkt; wYtmDen += mkt; }
          if (!isNaN(riskNum)) {
            var tip = name + '\n' + figi + '\nYTM: ' + r2(ytmPct) + '%';
            scatterRiskYield.push([riskNum, r2(ytmPct), tip]);
          }
          topYtmCandidates.push({ name: name, figi: figi, ytm: ytmPct, years: yearsToMat });
        }
      }

      var qty = Number(val(c.qty)) || 0;
      var nextStr = val(c.nextCp);
      if (qty > 0 && cupPY > 0 && cupVal > 0 && nextStr) {
        var first = (nextStr instanceof Date) ? nextStr : (isNaN(Date.parse(nextStr)) ? null : new Date(Date.parse(nextStr)));
        if (first) {
          var periodDays = Math.max(15, Math.round(365 / cupPY));
          var d2 = strip(first);
          while (d2 <= horizon) {
            if (d2 >= strip(today)) {
              var key = d2.getFullYear() + '-' + ('0' + (d2.getMonth() + 1)).slice(-2);
              monthAgg[key] = (monthAgg[key] || 0) + (cupVal * qty);
            }
            d2 = addDays_(d2, periodDays);
          }
        }
      }
    });

    return {
      invested: invested,
      market: market,
      plRub: plRub,
      sectors: sectors,
      sectorPl: sectorPl,
      riskCnt: riskCnt,
      wCouponNum: wCouponNum,
      wCouponDen: wCouponDen,
      wYtmNum: wYtmNum,
      wYtmDen: wYtmDen,
      scatterRiskYield: scatterRiskYield,
      topYtmCandidates: topYtmCandidates,
      monthAgg: monthAgg,
      byYear: byYear
    };
  }

  function computeDerived_(base, rows, c, ctx) {
    function r2(x) { return Math.round(Number(x || 0) * 100) / 100; }

    var plPctTotal = base.invested > 0 ? (base.plRub / base.invested * 100) : 0;
    var wCouponPct = base.wCouponDen > 0 ? (base.wCouponNum / base.wCouponDen * 100) : 0;
    var wYtmPct    = base.wYtmDen > 0 ? (base.wYtmNum / base.wYtmDen * 100) : 0;

    var secArr = Object.keys(base.sectors).sort().map(function (s) { return [s, r2(base.sectors[s])]; });

    var years = Object.keys(base.byYear).sort();
    var yrArr = years.map(function (y) { return [Number(y), r2(base.byYear[y])]; });

    var monthsSorted = Object.keys(base.monthAgg).sort();
    var monArr = monthsSorted.map(function (k) { return [k, r2(base.monthAgg[k])]; });

    base.topYtmCandidates.sort(function (a, b) { return (b.ytm || -1e9) - (a.ytm || -1e9); });
    var top5 = base.topYtmCandidates.slice(0, 5);

    var secPlArr = Object.keys(base.sectorPl).map(function (s) { return { sec: s, pl: base.sectorPl[s] }; });
    secPlArr.sort(function (a, b) { return a.pl - b.pl; });
    var worstSectors = secPlArr.slice(0, Math.min(3, secPlArr.length));

    var afterHistory = ctx.cfg.layout.chartsTopRow
      + ctx.cfg.layout.historyChart.rowOffset
      + ctx.cfg.layout.historyChart.heightRows
      + 2;

    var cmpStartRow = Math.max(
      ctx.cfg.layout.safeStartRowAfterHistory,
      ctx.cfg.layout.chartsTopRow + Math.max(years.length, monthsSorted.length) + 2,
      afterHistory
    );

    var top3StartRow = Math.max(
      cmpStartRow + ctx.cfg.layout.cmpBlockHeightRows,
      ctx.cfg.layout.chartsTopRow + Math.max(years.length, monthsSorted.length) + 12,
      afterHistory + 12
    );

    var items = rows.map(function (r) {
      function val(ci) { return (ci ? r[ci - 1] : ''); }
      var name = String(val(c.name) || '');
      var plPct = Number(val(c.plPct));

      var cupPct = null;
      if (c.couponPct) {
        var tmp = Number(val(c.couponPct));
        if (!isNaN(tmp)) cupPct = tmp;
      }
      if (cupPct == null) {
        var price = Number(val(c.price));
        var cupPY = Number(val(c.cupPY));
        var cupVal = Number(val(c.cupVal));
        if (price > 0 && cupPY > 0 && cupVal >= 0) cupPct = (cupVal * cupPY) / price * 100;
      }

      return {
        name: name,
        cupPct: isFinite(cupPct) ? r2(cupPct) : '',
        plPct: isFinite(plPct) ? r2(plPct) : null
      };
    }).filter(function (x) { return x.name && x.plPct != null; });

    var best3 = items.slice().sort(function (a, b) { return b.plPct - a.plPct; }).slice(0, 3);
    var worst3 = items.slice().sort(function (a, b) { return a.plPct - b.plPct; }).slice(0, 3);

    return {
      plPctTotal: plPctTotal,
      wCouponPct: wCouponPct,
      wYtmPct: wYtmPct,
      secArr: secArr,
      years: years,
      yrArr: yrArr,
      monthsSorted: monthsSorted,
      monArr: monArr,
      top5: top5,
      worstSectors: worstSectors,
      cmpStartRow: cmpStartRow,
      top3StartRow: top3StartRow,
      best3: best3,
      worst3: worst3
    };
  }

  return { computeBase_: computeBase_, computeDerived_: computeDerived_ };
})();

var DashRender = (function () {

  function renderTables_(ctx, base, d) {
    function r2(x) { return Math.round(Number(x || 0) * 100) / 100; }
    var dst = ctx.dst;

    var kpi = [
      ['Показатель', 'Значение'],
      ['Инвестировано', r2(base.invested)],
      ['Рыночная стоимость', r2(base.market)],
      ['P/L (руб)', r2(base.plRub)],
      ['P/L (%)', r2(d.plPctTotal)],
      ['Средневзв. купонная доходность (%)', r2(d.wCouponPct)],
      ['Средневзв. YTM (%)', r2(d.wYtmPct)]
    ];
    dst.getRange(ctx.cfg.blocks.kpi.r, ctx.cfg.blocks.kpi.c, kpi.length, 2).setValues(kpi);
    dst.getRange(ctx.cfg.blocks.kpi.r, ctx.cfg.blocks.kpi.c, 1, 2).setFontWeight('bold');

    var riskTable = [
      ['Категория риска', 'Кол-во'],
      ['Низкий', base.riskCnt['Низкий']],
      ['Средний', base.riskCnt['Средний']],
      ['Высокий', base.riskCnt['Высокий']]
    ];
    dst.getRange(ctx.cfg.blocks.risk.r, ctx.cfg.blocks.risk.c, riskTable.length, 2).setValues(riskTable).setFontWeight('bold');

    dst.getRange(ctx.cfg.blocks.sec.r, ctx.cfg.blocks.sec.c, 1, 2).setValues([['Сектор', 'Рыночная стоимость']]).setFontWeight('bold');
    if (d.secArr.length) dst.getRange(ctx.cfg.blocks.sec.r + 1, ctx.cfg.blocks.sec.c, d.secArr.length, 2).setValues(d.secArr);

    dst.getRange(ctx.cfg.blocks.year.r, ctx.cfg.blocks.year.c, 1, 2).setValues([['Год погашения', 'Рыночная стоимость']]).setFontWeight('bold');
    if (d.years.length) dst.getRange(ctx.cfg.blocks.year.r + 1, ctx.cfg.blocks.year.c, d.yrArr.length, 2).setValues(d.yrArr);

    dst.getRange(ctx.cfg.blocks.coup.r, ctx.cfg.blocks.coup.c, 1, 2).setValues([['Месяц', 'Купоны (₽)']]).setFontWeight('bold');
    if (d.monthsSorted.length) dst.getRange(ctx.cfg.blocks.coup.r + 1, ctx.cfg.blocks.coup.c, d.monArr.length, 2).setValues(d.monArr);
  }

  function removeCharts_(ctx) {
    ctx.dst.getCharts().forEach(function (ch) { ctx.dst.removeChart(ch); });
  }

  function renderCharts_(ctx, base, d) {
    var dst = ctx.dst;

    var riskRange = dst.getRange(ctx.cfg.blocks.risk.r, ctx.cfg.blocks.risk.c, 4, 2);
    var riskChart = dst.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(riskRange)
      .setPosition(1, 16, 0, 0)
      .setOption('title', 'Распределение по рискам (шт.)')
      .setOption('legend', { position: 'none' })
      .setOption('colors', [ctx.cfg.paletteRisk[0]])
      .build();
    dst.insertChart(riskChart);

    if (d.secArr.length) {
      var secRange = dst.getRange(ctx.cfg.blocks.sec.r, ctx.cfg.blocks.sec.c, Math.max(2, d.secArr.length + 1), 2);
      var pie = dst.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(secRange)
        .setPosition(ctx.cfg.layout.chartsTopRow, 1, 0, 0)
        .setOption('title', 'Структура по секторам (рыночная стоимость)')
        .setOption('legend', { position: 'right' })
        .setOption('pieSliceText', 'percentage')
        .setOption('colors', ctx.cfg.paletteMain.slice(0, Math.max(3, d.secArr.length)))
        .build();
      dst.insertChart(pie);
    }

    if (d.years.length) {
      var yrRange = dst.getRange(ctx.cfg.blocks.year.r, ctx.cfg.blocks.year.c, Math.max(2, d.years.length + 1), 2);
      var matChart = dst.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(yrRange)
        .setPosition(ctx.cfg.layout.chartsTopRow, 7, 0, 0)
        .setOption('title', 'Сроки погашения (рыночная стоимость)')
        .setOption('legend', { position: 'none' })
        .setOption('colors', ['#3B82F6'])
        .build();
      dst.insertChart(matChart);
    }

    if (d.monthsSorted.length) {
      var monRange = dst.getRange(ctx.cfg.blocks.coup.r, ctx.cfg.blocks.coup.c, Math.max(2, d.monthsSorted.length + 1), 2);
      var monChart = dst.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(monRange)
        .setPosition(ctx.cfg.layout.chartsTopRow, 13, 0, 0)
        .setOption('title', 'График купонных выплат (6 месяцев)')
        .setOption('legend', { position: 'none' })
        .setOption('colors', ['#22C55E'])
        .build();
      dst.insertChart(monChart);
    }

    // --- История портфеля (A12:C28): Инвестировано vs Стоимость ---
    var h = ctx.cfg.history;
    var histRange = dst.getRange(h.r1, h.c1, h.r2 - h.r1 + 1, h.w);

    // min/max по B,C (без шапки и пустых)
    var histVals = histRange.getValues();
    var vmin = null, vmax = null;

    for (var i = 1; i < histVals.length; i++) {
      for (var j = 1; j <= 2; j++) { // B,C
        var v = Number(histVals[i][j]);
        if (!isFinite(v)) continue;
        if (vmin === null || v < vmin) vmin = v;
        if (vmax === null || v > vmax) vmax = v;
      }
    }

    // если данных нет — не строим
    if (vmin !== null && vmax !== null) {
      var span = vmax - vmin;
      var pad = span > 0 ? span * 0.05 : Math.max(1, Math.abs(vmax) * 0.02); // 5% или 2% если всё ровно
      var vwMin = vmin - pad;
      var vwMax = vmax + pad;

      var histChart = dst.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(histRange)
        .setPosition(
          ctx.cfg.layout.chartsTopRow + ctx.cfg.layout.historyChart.rowOffset,
          ctx.cfg.layout.historyChart.col,
          0, 0
        )
        .setOption('title', 'История портфеля: Инвестировано vs Стоимость')
        .setOption('legend', { position: 'bottom', alignment: 'center' })
        .setOption('hAxis', { title: 'Дата', format: 'dd.MM' })
        .setOption('vAxis', {
          title: '₽',
          viewWindowMode: 'explicit',
          viewWindow: { min: vwMin, max: vwMax }
        })
        .setOption('pointSize', 5)
        .setOption('series', {
          0: { labelInLegend: 'Инвестировано', pointSize: 5 },
          1: { labelInLegend: 'Стоимость',    pointSize: 5 }
        })
        .build();

      dst.insertChart(histChart);
    }


    var cmpData = [
      ['Метрика', 'Значение'],
      ['Средневзв. YTM (%)', Math.round(Number(d.wYtmPct || 0) * 100) / 100],
      ['Купонная доходность (%)', Math.round(Number(d.wCouponPct || 0) * 100) / 100]
    ];
    dst.getRange(d.cmpStartRow, 1, cmpData.length, 2).setValues(cmpData).setFontWeight('bold');

    var cmpRange = dst.getRange(d.cmpStartRow, 1, cmpData.length, 2);
    var cmpChart = dst.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(cmpRange)
      .setPosition(d.cmpStartRow, 4, 0, 0)
      .setOption('title', 'YTM vs Купонная доходность (средневзв., %)')
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#4F46E5'])
      .build();
    dst.insertChart(cmpChart);

    if (base.scatterRiskYield.length > 1) {
      dst.getRange(d.cmpStartRow, 7, base.scatterRiskYield.length, 3).setValues(base.scatterRiskYield).setFontWeight('bold');
      var scRange = dst.getRange(d.cmpStartRow, 7, base.scatterRiskYield.length, 3);
      var scChart = dst.newChart()
        .setChartType(Charts.ChartType.SCATTER)
        .addRange(scRange)
        .setPosition(d.cmpStartRow, 11, 0, 0)
        .setOption('title', 'Риск vs Доходность к погашению (YTM)')
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', { title: 'Риск (баллы)' })
        .setOption('vAxis', { title: 'YTM (%)' })
        .setOption('series', { 0: { pointSize: 5 } })
        .build();
      dst.insertChart(scChart);
    }
  }

  function renderExtras_(ctx, base, d) {
    function r2(x) { return Math.round(Number(x || 0) * 100) / 100; }

    var dst = ctx.dst;
    var EXT_COL = ctx.cfg.blocks.ext.c;
    var row = 1;

    var ytmTab = [['Top YTM (5)', 'YTM (%)', 'До погашения (лет)', 'FIGI']];
    d.top5.forEach(function (x) { ytmTab.push([x.name, r2(x.ytm), r2(x.years), x.figi]); });

    dst.getRange(row, EXT_COL, ytmTab.length, ytmTab[0].length).setValues(ytmTab);
    dst.getRange(row, EXT_COL, 1, ytmTab[0].length).setFontWeight('bold');
    row += ytmTab.length + 2;

    if (d.best3.length) {
      dst.getRange(row, EXT_COL, 1, 3).setValues([['Top P/L (3) — best', '', '']]).setFontWeight('bold');
      dst.getRange(row + 1, EXT_COL, 1, 3).setValues([['Название', 'Купон (%)', 'P/L (%)']]).setFontWeight('bold');
      var bestRows = d.best3.map(function (x) { return [x.name, x.cupPct, x.plPct]; });
      dst.getRange(row + 2, EXT_COL, bestRows.length, 3).setValues(bestRows);
      row += 2 + bestRows.length + 2;
    }

    if (d.worst3.length) {
      dst.getRange(row, EXT_COL, 1, 3).setValues([['Top P/L (3) — worst', '', '']]).setFontWeight('bold');
      dst.getRange(row + 1, EXT_COL, 1, 3).setValues([['Название', 'Купон (%)', 'P/L (%)']]).setFontWeight('bold');
      var worstRows = d.worst3.map(function (x) { return [x.name, x.cupPct, x.plPct]; });
      dst.getRange(row + 2, EXT_COL, worstRows.length, 3).setValues(worstRows);
      row += 2 + worstRows.length + 2;
    }

    var drawTab = [['P/L по секторам (TOP-худшие)', 'P/L (руб)']];
    d.worstSectors.forEach(function (x) { drawTab.push([x.sec, r2(x.pl)]); });

    dst.getRange(row, EXT_COL, drawTab.length, drawTab[0].length).setValues(drawTab);
    dst.getRange(row, EXT_COL, 1, drawTab[0].length).setFontWeight('bold');

    writeTop3Visual_(dst, d.top3StartRow, d.best3, d.worst3);

    dst.autoResizeColumns(1, 20);
  }

  function writeTop3Visual_(dst, startRow, best3, worst3) {
    var colA = 1, colE = 5;

    dst.getRange(startRow, colA, 1, 3).setValues([['Топ-3 прибыльные облигации (P/L, %)', '', '']]).setFontWeight('bold');
    dst.getRange(startRow, colE, 1, 3).setValues([['Топ-3 убыточные облигации (P/L, %)', '', '']]).setFontWeight('bold');

    dst.getRange(startRow + 1, colA, 1, 3).setValues([['Название', 'Купон (%)', 'P/L (%)']]).setFontWeight('bold');
    dst.getRange(startRow + 1, colE, 1, 3).setValues([['Название', 'Купон (%)', 'P/L (%)']]).setFontWeight('bold');

    if (best3.length) {
      var bestRows = best3.map(function (x) { return [x.name, x.cupPct, x.plPct]; });
      dst.getRange(startRow + 2, colA, bestRows.length, 3).setValues(bestRows);
    }
    if (worst3.length) {
      var worstRows = worst3.map(function (x) { return [x.name, x.cupPct, x.plPct]; });
      dst.getRange(startRow + 2, colE, worstRows.length, 3).setValues(worstRows);
    }
  }

  return {
    renderTables_: renderTables_,
    removeCharts_: removeCharts_,
    renderCharts_: renderCharts_,
    renderExtras_: renderExtras_
  };
})();

var HistoryStore = (function () {

  function snapshot_(ctx) {
    var h = ctx.cfg.history;
    try {
      if (ctx.dst.getLastRow() >= h.r1) {
        return ctx.dst.getRange(h.r1, h.c1, h.r2 - h.r1 + 1, h.w).getValues();
      }
    } catch (e) {}
    return null;
  }

  function update_(snapshot, investedRub, marketRub, ctx) {
    var h = ctx.cfg.history;
    var N = (h.r2 - h.r1);
    var now = ctx.now;

    var entries = [];

    if (snapshot && snapshot.length) {
      for (var i = 1; i < snapshot.length; i++) {
        var d = snapshot[i][0];
        var inv = snapshot[i][1];
        var mkt = snapshot[i][2];

        if (d === '' || d == null) continue;

        var dd = (d instanceof Date) ? d : new Date(d);
        if (!isFinite(dd.getTime())) continue;

        inv = Number(inv); mkt = Number(mkt);
        if (!isFinite(inv)) inv = null;
        if (!isFinite(mkt)) mkt = null;

        entries.push({ dt: dd, invested: inv, market: mkt });
      }
    }

    entries.push({
      dt: now,
      invested: Number(investedRub) || 0,
      market: Number(marketRub) || 0
    });

    if (entries.length > N) entries = entries.slice(entries.length - N);

    var out = [];
    out.push(['Дата', 'Инвестировано', 'Стоимость']);

    while (entries.length < N) entries.unshift(null);

    for (var k = 0; k < N; k++) {
      var e = entries[k];
      if (!e) out.push(['', '', '']);
      else out.push([e.dt, e.invested == null ? '' : e.invested, e.market == null ? '' : e.market]);
    }

    return out;
  }

  function write_(ctx, block) {
    var dst = ctx.dst;
    var h = ctx.cfg.history;

    dst.getRange(h.r1, h.c1, h.r2 - h.r1 + 1, h.w).setValues(block);
    dst.getRange(h.r1, h.c1, 1, h.w).setFontWeight('bold');
    dst.getRange(h.r1 + 1, h.c1, h.r2 - h.r1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    dst.getRange(h.r1 + 1, h.c1 + 1, h.r2 - h.r1, 2).setNumberFormat('#,##0.00');
  }

  return { snapshot_: snapshot_, update_: update_, write_: write_ };
})();
