/**
 * dashboard_exports.gs
 * Экспорт диаграмм и текстовой статистики с листа Dashboard.
 *
 * Что делает:
 *  - Экспорт PNG диаграмм (по реестру + динамически найденные диапазоны для "Метрика" и "Scatter")
 *  - Читает секции из колонки T (Top YTM, Top P/L best/worst, худшие сектора)
 *  - Читает историю портфеля из A12:C28 (дата/инвестировано/стоимость) и формирует краткую сводку
 *
 * Требует:
 *  - buildBondsDashboard() уже отрисовал Dashboard
 *  - tgSendMessage_(), tgSendPhoto_() определены в tgbot_api.gs
 */

// --- Реестр диаграмм с фиксированным top-left диапазона данных ---
var DASHBOARD_CHARTS = {
  risk:     { topLeft: { row: 1, col: 4  }, caption: 'Распределение по рискам (шт.)' },
  sectors:  { topLeft: { row: 1, col: 7  }, caption: 'Структура по секторам (рыночная стоимость)' },
  maturity: { topLeft: { row: 1, col: 10 }, caption: 'Сроки погашения (рыночная стоимость)' },
  coupons:  { topLeft: { row: 1, col: 13 }, caption: 'График купонных выплат (6 месяцев)' },
  history: { topLeft: { row: 12, col: 1 }, caption: 'История портфеля: Инвестировано vs Стоимость' }
};

  // --- Секции в колонке T ---
var DASHBOARD_SECTIONS = {
  topYtm:   { header: 'Top YTM (5)',                  title: 'Топ-5 облигаций по YTM (%):' },
  bestPL:   { header: 'Top P/L (3) — best',           title: 'Топ-3 прибыльные облигации (P/L, %):' },
  worstPL:  { header: 'Top P/L (3) — worst',          title: 'Топ-3 убыточные облигации (P/L, %):' },
  worstSec: { header: 'P/L по секторам (TOP-худшие)', title: 'Просадка P/L по секторам (худшие):' }
};

function buildAndSendDashboardPackage() {
  

  buildBondsDashboard();
  sendDashboardGreets_();
  sendDashboardStatsFromDashboard_();

  // по умолчанию отправляем 4 фиксированных, + 2 динамических если найдутся
  var count = sendDashboardChartsToTelegram_(['sectors','coupons','maturity','history','risk','ytmVsCoupon','scatter']);
  showSnack_('Отправлено диаграмм: ' + count, 'Telegram', 2500);
}


// -------------------- Диаграммы --------------------

function _findChartByTopLeft_(sheet, row, col) {
  var charts = sheet.getCharts();
  for (var i = 0; i < charts.length; i++) {
    var ranges = charts[i].getRanges();
    for (var j = 0; j < ranges.length; j++) {
      var r = ranges[j];
      if (r.getSheet().getName() === sheet.getName() && r.getRow() === row && r.getColumn() === col) {
        return charts[i];
      }
    }
  }
  return null;
}

function _findCellExact_(sheet, text, startRow, startCol, numRows, numCols) {
  var vals = sheet.getRange(startRow, startCol, numRows, numCols).getValues();
  text = String(text);
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[0].length; c++) {
      if (String(vals[r][c]).trim() === text) {
        return { row: startRow + r, col: startCol + c };
      }
    }
  }
  return null;
}

// "Метрика | Значение" лежит в колонках A:B, ищем "Метрика" как маркер
function _guessCmpTopLeft_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var hit = _findCellExact_(sheet, 'Метрика', 1, 1, Math.min(lastRow, 400), 6);
  if (!hit) return null;
  return { row: hit.row, col: hit.col };
}

// "Риск | YTM (%) | Тултип" лежит в G:I рядом с cmpStartRow
function _guessScatterTopLeft_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var vals = sheet.getRange(1, 1, Math.min(lastRow, 600), 12).getValues();
  for (var r = 0; r < vals.length; r++) {
    var g = String(vals[r][6] || '').trim();  // col G = 7 -> idx 6
    var h = String(vals[r][7] || '').trim();  // col H
    var i = String(vals[r][8] || '').trim();  // col I
    if (g === 'Риск' && h === 'YTM (%)' && i === 'Тултип') {
      return { row: r + 1, col: 7 };
    }
  }
  return null;
}

function exportDashboardCharts_(aliases) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Dashboard');
  if (!sh) throw new Error('Лист Dashboard не найден. Сначала выполните buildBondsDashboard().');

  var out = {};
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');

  var registry = {};
  Object.keys(DASHBOARD_CHARTS).forEach(function(k){ registry[k] = DASHBOARD_CHARTS[k]; });

  // динамические (не в фиксированном реестре)
  registry.ytmVsCoupon = { topLeft: _guessCmpTopLeft_(sh),     caption: 'YTM vs Купонная доходность (средневзв., %)' };
  registry.scatter     = { topLeft: _guessScatterTopLeft_(sh), caption: 'Риск vs Доходность к погашению (YTM)' };

  var keys = (aliases && aliases.length) ? aliases : Object.keys(registry);

  keys.forEach(function(alias){
    var cfg = registry[alias];
    if (!cfg || !cfg.topLeft) { out[alias] = null; return; }

    var chart = _findChartByTopLeft_(sh, cfg.topLeft.row, cfg.topLeft.col);
    if (!chart) { out[alias] = null; return; }

    var blob = chart.getAs('image/png').setName('dashboard_' + alias + '_' + ts + '.png');
    out[alias] = { blob: blob, caption: cfg.caption };
  });

  return out;
}

// -------------------- Секции (колонка T) --------------------

function _readSectionTableT_(sheet, headerText, extCol, maxWidth) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  extCol = extCol || 20;     // T
  maxWidth = maxWidth || 6;  // читаем вправо до 6 колонок

  // найти строку, где в колонке T стоит headerText
  var colVals = sheet.getRange(1, extCol, lastRow, 1).getValues();
  var headerRow = -1;
  for (var r = 0; r < colVals.length; r++) {
    if (String(colVals[r][0]).trim() === String(headerText)) { headerRow = r + 1; break; }
  }
  if (headerRow < 1) return [];

  // определить, является ли строка headerRow одновременно шапкой таблицы
  var headerLine = sheet.getRange(headerRow, extCol, 1, maxWidth).getValues()[0];
  var nonEmpty = 0;
  for (var i = 0; i < maxWidth; i++) if (String(headerLine[i] || '').trim()) nonEmpty++;

  var dataStartRow = null;

  // если в строке заголовка есть >1 непустой ячейки, то это шапка таблицы (как Top YTM, Worst sectors)
  if (nonEmpty > 1) dataStartRow = headerRow + 1;
  else dataStartRow = headerRow + 2; // иначе: отдельный заголовок + отдельная строка-шапка

  var out = [];
  for (var rr = dataStartRow; rr <= lastRow; rr++) {
    var rowVals = sheet.getRange(rr, extCol, 1, maxWidth).getValues()[0];
    if (!String(rowVals[0] || '').trim()) break;
    out.push(rowVals);
  }
  return out;
}

// -------------------- История A12:C28 --------------------

function _readHistoryABC_(sheet) {
  var R1 = 12, R2 = 28, C1 = 1, W = 3;
  var block = sheet.getRange(R1, C1, R2 - R1 + 1, W).getValues();

  // block[0] — шапка; данные с 1
  var points = [];
  for (var i = 1; i < block.length; i++) {
    var d = block[i][0];
    var inv = block[i][1];
    var mkt = block[i][2];
    if (d === '' || d == null) continue;

    var dd = (d instanceof Date) ? d : new Date(d);
    if (!isFinite(dd.getTime())) continue;

    inv = Number(inv);
    mkt = Number(mkt);
    if (!isFinite(inv)) inv = null;
    if (!isFinite(mkt)) mkt = null;

    points.push({ dt: dd, invested: inv, market: mkt });
  }
  return points;
}

function _formatMoney_(x) {
  if (x == null || !isFinite(Number(x))) return '—';
  return (Math.round(Number(x) * 100) / 100).toFixed(2);
}

function _historySummaryText_(points) {
  if (!points || !points.length) return 'История портфеля: нет данных (A12:C28 пусто).';

  // берем последние непустые по market
  var p = points.filter(function(x){ return x.market != null; });
  if (!p.length) return 'История портфеля: нет значений стоимости.';

  var last = p[p.length - 1];
  var prev = (p.length >= 2) ? p[p.length - 2] : null;
  var first = p[0];

  var lines = [];
  lines.push('Стоимость портфеля: ' + _formatMoney_(last.market) + ' ₽');
  if (last.invested != null) lines.push('Инвестировано: ' + _formatMoney_(last.invested) + ' ₽');

  if (prev && prev.market != null) {
    var d1 = last.market - prev.market;
    var p1 = (prev.market !== 0) ? (d1 / prev.market * 100) : null;
    lines.push('Δ к прошлому запуску: ' + _formatMoney_(d1) + ' ₽' + (p1 != null ? (' (' + _formatMoney_(p1) + '%)') : ''));
  }

  if (first && first.market != null && first !== last) {
    var d2 = last.market - first.market;
    var p2 = (first.market !== 0) ? (d2 / first.market * 100) : null;
    lines.push('Δ за окно: ' + _formatMoney_(d2) + ' ₽' + (p2 != null ? (' (' + _formatMoney_(p2) + '%)') : ''));
  }

  return lines.join('\n');
}

// -------------------- Отправка --------------------

function sendDashboardGreets_() {
  var tz = Session.getScriptTimeZone() || 'Etc/GMT';
  var now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  var text = 'Отчёт по портфелю на ' + now + '.\n' +
             'Обновлены данные по бумагам и Dashboard.';
  tgSendMessage_(text);
}

function sendDashboardStatsFromDashboard_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Dashboard');
  if (!sh) { tgSendMessage_('Лист Dashboard не найден.'); return; }

  var extCol = 20; // T
  var topYtm   = _readSectionTableT_(sh, DASHBOARD_SECTIONS.topYtm.header, extCol, 6);
  var bestPL   = _readSectionTableT_(sh, DASHBOARD_SECTIONS.bestPL.header, extCol, 6);
  var worstPL  = _readSectionTableT_(sh, DASHBOARD_SECTIONS.worstPL.header, extCol, 6);
  var worstSec = _readSectionTableT_(sh, DASHBOARD_SECTIONS.worstSec.header, extCol, 6);

  var lines = [];

  if (topYtm.length) {
    lines.push(DASHBOARD_SECTIONS.topYtm.title);
    topYtm.forEach(function(r, i){
      var name = String(r[0] || '');
      var ytm  = (r[1] != null && r[1] !== '') ? Number(r[1]).toFixed(2) : '—';
      var yrs  = (r[2] != null && r[2] !== '') ? Number(r[2]).toFixed(2) : '—';
      lines.push((i + 1) + '. ' + name + ' — YTM ' + ytm + '%, до погашения ' + yrs + ' г.');
    });
    lines.push('');
  }

  if (bestPL.length) {
    lines.push(DASHBOARD_SECTIONS.bestPL.title);
    bestPL.forEach(function(r, i){
      var name = String(r[0] || '');
      var cup  = (r[1] != null && r[1] !== '') ? Number(r[1]).toFixed(2) : '—';
      var pl   = (r[2] != null && r[2] !== '') ? Number(r[2]).toFixed(2) : '—';
      lines.push((i + 1) + '. ' + name + ' — купон ' + cup + '%, P/L ' + pl + '%');
    });
    lines.push('');
  }

  if (worstPL.length) {
    lines.push(DASHBOARD_SECTIONS.worstPL.title);
    worstPL.forEach(function(r, i){
      var name = String(r[0] || '');
      var cup  = (r[1] != null && r[1] !== '') ? Number(r[1]).toFixed(2) : '—';
      var pl   = (r[2] != null && r[2] !== '') ? Number(r[2]).toFixed(2) : '—';
      lines.push((i + 1) + '. ' + name + ' — купон ' + cup + '%, P/L ' + pl + '%');
    });
    lines.push('');
  }

  if (worstSec.length) {
    lines.push(DASHBOARD_SECTIONS.worstSec.title);
    worstSec.forEach(function(r){
      var sec = String(r[0] || '');
      var pl  = (r[1] != null && r[1] !== '') ? Number(r[1]).toFixed(2) : '0.00';
      lines.push('- ' + sec + ': ' + pl + ' ₽');
    });
    lines.push('');
  }

  // история A12:C28
  lines.push(_historySummaryText_(_readHistoryABC_(sh)));

  if (!lines.length) lines.push('Статистика на Dashboard не найдена.');
  tgSendMessage_(lines.join('\n'));
}

function sendDashboardChartsToTelegram_(aliases) {
  var pack = exportDashboardCharts_(aliases);
  var sent = 0;
  Object.keys(pack).forEach(function(alias){
    var item = pack[alias];
    if (item && item.blob) {
      tgSendPhoto_(item.blob, item.caption);
      sent++;
    }
  });
  return sent;
}


