/**
 * bonds_dashboard.gs
 * Построение листа «Dashboard» с лок-синхронизацией
 */

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
    var cRiskNum = idx('Риск (ручн.)');       // ручной риск (по Rules)
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
    var riskCnt = { 'Низкий':0, 'Средний':0, 'Высокий':0 };
    var wCouponNum = 0, wCouponDen = 0;

    var wYtmNum = 0, wYtmDen = 0;
    var scatterRiskYield = [['Риск','YTM (%)','Тултип']];

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
      sectors[sec] = (sectors[sec]||0) + mkt;

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

    // Сектора
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

    // Палитры
    var paletteMain = ['#4F46E5','#22C55E','#EAB308','#EF4444','#06B6D4','#A855F7','#F59E0B','#94A3B8','#10B981','#3B82F6'];
    var paletteRisk = ['#10B981','#EAB308','#EF4444'];

    // Диаграмма: Риски
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

    // Диаграмма: Сектора (pie)
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

    // Диаграмма: Погашения по годам
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

    // Диаграмма: Купоны по месяцам
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

    // Диаграмма: YTM vs Купонная доходность
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

    // Scatter: Риск vs YTM
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

    dst.autoResizeColumns(1, 20);
    showSnack_('Dashboard обновлён','Dashboard',2000);
  } finally {
    lock.releaseLock();
  }
}

// helpers (повторно используем)
function addDays_(d, n){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()+n); }
function addMonths_(d, n){ return new Date(d.getFullYear(), d.getMonth()+n, 1); }
