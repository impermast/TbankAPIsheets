function buildBondsDashboard(){
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { showSnack_('Другая операция обновляет Dashboard. Повторите позже.','Dashboard',3000); return; }

  try {
    var ss = SpreadsheetApp.getActive();
    var src = ss.getSheetByName('Bonds');
    if(!src){ showSnack_('Лист Bonds не найден','Dashboard',3000); return; }

    var dst = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');
    dst.clear();

    // --- Индексы колонок ---
    var hdr = src.getRange(1,1,1,src.getLastColumn()).getValues()[0];
    function idx(name){ var i = hdr.indexOf(name); return (i>=0)? (i+1) : 0; }

    var cName    = idx('Название');
    var cFIGI    = idx('FIGI');
    var cRiskNum = idx('Риск (ручн.)');
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
      showSnack_('Не хватает обязательных колонок (Риск/Сектор/Рыночная/…)','Dashboard',3500);
      return;
    }

    var lastRow = src.getLastRow();
    if(lastRow < 2){ showSnack_('Нет данных для сводки','Dashboard',2500); return; }

    var rows = src.getRange(2,1,lastRow-1,src.getLastColumn()).getValues();

    // --- Агрегации ---
    var invested = 0, market = 0, plRub = 0;
    var sectors = {};
    var sectorPl = {};                       // ← NEW: P/L по секторам (для просадки)
    var riskCnt = { 'Низкий':0, 'Средний':0, 'Высокий':0 };
    var wCouponNum = 0, wCouponDen = 0;

    var wYtmNum = 0, wYtmDen = 0;
    var scatterRiskYield = [['Риск','YTM (%)','Тултип']];

    var topYtmCandidates = [];               // ← NEW: сбор кандидатов для Top-5 YTM

    var monthAgg = {};
    var today = new Date();
    var horizon = addMonths_(today, 6);
    var strip   = function(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); };
    var r2      = function(x){ return Math.round(Number(x||0)*100)/100; };

    rows.forEach(function(r){
      function val(ci){ return (ci? r[ci-1] : ''); }

      // Риск: бинирование
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
      sectors[sec]  = (sectors[sec] || 0) + mkt;
      sectorPl[sec] = (sectorPl[sec]|| 0) + pl;             // ← NEW: аккумулируем P/L по сектору

      var coupPctAltIdx = idx('Купон, %');
      var coupPctAlt = Number(coupPctAltIdx ? val(coupPctAltIdx) : NaN);
      var coupPctCalc = Number(val(cCupPY) ? (((Number(val(cCupVal))||0) * Number(val(cCupPY))) / (Number(val(cPrice))||0) * 100) : NaN);
      var coupUse = !isNaN(coupPctAlt) ? coupPctAlt : coupPctCalc;
      if(!isNaN(coupUse) && mkt>0){ wCouponNum += (coupUse/100)*mkt; wCouponDen += mkt; }

      // YTM (приближённо)
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
          // ← NEW: собираем кандидатов (можно фильтровать по liqudity/mkt>0)
          topYtmCandidates.push({ name:name, figi:figi, ytm:ytmPct, years:yearsToMat });
        }
      }

      // купоны по месяцам (6 мес)
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

    // KPI
    var kpi = [
      ['Показатель','Значение'],
      ['Инвестировано', r2(invested)],
      ['Рыночная стоимость', r2(market)],
      ['P/L (руб)', r2(plRub)],
      ['P/L (%)', r2(plPctTotal)],
      ['Средневзв. купонная доходность (%)', r2(wCouponPct)],
      ['Средневзв. YTM (%)', r2(wYtmPct)]
    ];
    dst.getRange(1,1,kpi.length,2).setValues(kpi);
    dst.getRange(1,1,1,2).setFontWeight('bold');

    // Риски
    var riskTable = [
      ['Категория риска','Кол-во'],
      ['Низкий',  riskCnt['Низкий']],
      ['Средний', riskCnt['Средний']],
      ['Высокий', riskCnt['Высокий']]
    ];
    dst.getRange(1,4,riskTable.length,riskTable[0].length).setValues(riskTable).setFontWeight('bold');

    // Сектора (рыночная стоимость)
    var secArr = Object.keys(sectors).sort().map(function(s){ return [s, r2(sectors[s])]; });
    dst.getRange(1,7,1,2).setValues([['Сектор','Рыночная стоимость']]).setFontWeight('bold');
    if(secArr.length) dst.getRange(2,7,secArr.length,2).setValues(secArr);

    // По годам погашения
    var byYear = {};
    if(cMaturity){
      rows.forEach(function(r){
        var mkt = Number(r[cMkt-1])||0;
        var v = r[cMaturity-1];
        if(v){
          var d = (v instanceof Date)? v : (isNaN(Date.parse(v))? null : new Date(Date.parse(v)));
          if(d){ var y = d.getFullYear(); byYear[y] = (byYear[y]||0) + mkt; }
        }
      });
    }
    var years = Object.keys(byYear).sort();
    dst.getRange(1,10,1,2).setValues([['Год погашения','Рыночная стоимость']]).setFontWeight('bold');
    if(years.length){
      var yrArr = years.map(function(y){ return [Number(y), r2(byYear[y])]; });
      dst.getRange(2,10,yrArr.length,2).setValues(yrArr);
    }

    // Купоны по месяцам
    var monthsSorted = Object.keys(monthAgg).sort();
    dst.getRange(1,13,1,2).setValues([['Месяц','Купоны (₽)']]).setFontWeight('bold');
    if(monthsSorted.length){
      var monArr = monthsSorted.map(function(k){ return [k, r2(monthAgg[k])]; });
      dst.getRange(2,13,monArr.length,2).setValues(monArr);
    }

    // Очистить старые диаграммы
    dst.getCharts().forEach(function(ch){ dst.removeChart(ch); });

    // Палитры (без изменений)
    var paletteMain = ['#4F46E5','#22C55E','#EAB308','#EF4444','#06B6D4','#A855F7','#F59E0B','#94A3B8','#10B981','#3B82F6'];
    var paletteRisk = ['#10B981','#EAB308','#EF4444'];

    // Диаграммы (как у тебя) …
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

    var cmpStartRow = Math.max(22, 20 + Math.max(years.length, monthsSorted.length) + 2);
    var cmpData = [
      ['Метрика','Значение'],
      ['Средневзв. YTM (%)', r2(wYtmPct)],
      ['Купонная доходность (%)', r2(wCouponPct)]
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

    // ====== NEW: Таблица Top-5 по YTM и Таблица «Просадка P/L по секторам» ======
    var EXT_COL = 20; // T
    var row = 1;

    // Top-5 по YTM
    topYtmCandidates.sort(function(a,b){ return (b.ytm||-1e9) - (a.ytm||-1e9); });
    var top5 = topYtmCandidates.slice(0, 5);
    var ytmTab = [['Top YTM (5)','YTM (%)','До погашения (лет)','FIGI']];
    top5.forEach(function(x){
      ytmTab.push([x.name, r2(x.ytm), r2(x.years), x.figi]);
    });
    dst.getRange(row, EXT_COL, ytmTab.length, ytmTab[0].length).setValues(ytmTab);
    dst.getRange(row, EXT_COL, 1, ytmTab[0].length).setFontWeight('bold');

    row += ytmTab.length + 2;

        // --- Топ-3 прибыльные/убыточные облигации по P/L (%) ---
    (function writeTop3Bonds(){
      // индексы уже посчитаны выше
      var cCouponPct = idx('Купон, %'); // может не быть

      // вспомогалки
      function valOf(row, ci){ return (ci ? row[ci-1] : ''); }
      function r2(x){ return Math.round(Number(x||0)*100)/100; }

      // соберём массив с нужными полями
      var items = rows.map(function(r){
        var name = String(valOf(r, cName) || '');
        var plPct = Number(valOf(r, cPLPct));
        // купон %: берём готовый столбец, иначе приблизительно считаем
        var cupPct = null;
        var price  = Number(valOf(r, cPrice));
        var cupPY  = Number(valOf(r, cCupPY));
        var cupVal = Number(valOf(r, cCupVal));

        if (cCouponPct){ 
          var tmp = Number(valOf(r, cCouponPct));
          if (!isNaN(tmp)) cupPct = tmp;
        }
        if (cupPct == null && price>0 && cupPY>0 && cupVal>=0){
          cupPct = (cupVal * cupPY) / price * 100;
        }

        return {
          name: name,
          cupPct: isFinite(cupPct) ? r2(cupPct) : '',
          plPct: isFinite(plPct) ? r2(plPct) : null
        };
      }).filter(function(x){ return x.name && x.plPct!=null; });

      if (!items.length) return;

      // топ-3: прибыльные (по убыванию) и убыточные (по возрастанию)
      var best = items.slice().sort(function(a,b){ return b.plPct - a.plPct; }).slice(0,3);
      var worst= items.slice().sort(function(a,b){ return a.plPct - b.plPct; }).slice(0,3);

      // найдём безопасную позицию под таблицы
      var startRow = Math.max(22, 20 + Math.max(years.length, monthsSorted.length) + 12); // ниже нижних чартов
      var colA = 1, colE = 5;

      // заголовки
      dst.getRange(startRow, colA, 1, 3).setValues([['Топ-3 прибыльные облигации (P/L, %)', '', '']]).setFontWeight('bold');
      dst.getRange(startRow, colE, 1, 3).setValues([['Топ-3 убыточные облигации (P/L, %)', '', '']]).setFontWeight('bold');

      // шапки таблиц
      dst.getRange(startRow+1, colA, 1, 3).setValues([['Название','Купон (%)','P/L (%)']]).setFontWeight('bold');
      dst.getRange(startRow+1, colE, 1, 3).setValues([['Название','Купон (%)','P/L (%)']]).setFontWeight('bold');

      // данные
      var bestRows  = best.map(function(x){ return [x.name, x.cupPct, x.plPct]; });
      var worstRows = worst.map(function(x){ return [x.name, x.cupPct, x.plPct]; });

      if (bestRows.length)  dst.getRange(startRow+2, colA, bestRows.length, 3).setValues(bestRows);
      if (worstRows.length) dst.getRange(startRow+2, colE, worstRows.length, 3).setValues(worstRows);
    })();

    dst.autoResizeColumns(1, 20);



    // Просадка P/L по секторам (TOP-3 худших)
    var secPlArr = Object.keys(sectorPl).map(function(s){ return {sec:s, pl:sectorPl[s]}; });
    secPlArr.sort(function(a,b){ return a.pl - b.pl; }); // по возрастанию (хуже — раньше)
    var worst = secPlArr.slice(0, Math.min(3, secPlArr.length));

    var drawTab = [['P/L по секторам (TOP-худшие)','P/L (руб)']];
    worst.forEach(function(x){ drawTab.push([x.sec, r2(x.pl)]); });

    dst.getRange(row, EXT_COL, drawTab.length, drawTab[0].length).setValues(drawTab);
    dst.getRange(row, EXT_COL, 1, drawTab[0].length).setFontWeight('bold');

    // Автоширина
    dst.autoResizeColumns(1, 26);

    showSnack_('Dashboard обновлён','Dashboard',2000);
  } finally {
    lock.releaseLock();
  }
}
