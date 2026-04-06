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
  history:  { topLeft: { row: 12, col: 1 }, caption: 'История портфеля: Инвестировано vs Стоимость' }
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
  var sh = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  if (!sh) { tgSendMessage_('Лист Dashboard не найден.'); return; }

  var extCol = 20; // T

  // Старые bond-секции: читаем как и раньше
  var topYtm   = _readSectionTableT_(sh, DASHBOARD_SECTIONS.topYtm.header, extCol, 6);
  var bestPL   = _readSectionTableT_(sh, DASHBOARD_SECTIONS.bestPL.header, extCol, 6);
  var worstPL  = _readSectionTableT_(sh, DASHBOARD_SECTIONS.worstPL.header, extCol, 6);
  var worstSec = _readSectionTableT_(sh, DASHBOARD_SECTIONS.worstSec.header, extCol, 6);

  function trim(v) {
    return String(v == null ? '' : v).trim();
  }

  function firstNonEmptyAfter0(row) {
    for (var i = 1; i < row.length; i++) {
      var v = trim(row[i]);
      if (v) return v;
    }
    return '—';
  }

  function isKnownHeader(text) {
    return text === 'Портфель — итого' ||
           text === 'Классы активов' ||
           text === 'Акции — топ-5 по позиции' ||
           text === 'Акции — топ-5 по P/L' ||
           text === DASHBOARD_SECTIONS.topYtm.header ||
           text === DASHBOARD_SECTIONS.bestPL.header ||
           text === DASHBOARD_SECTIONS.worstPL.header ||
           text === DASHBOARD_SECTIONS.worstSec.header;
  }

  function looksLikeHeaderRow(row) {
    var joined = row.map(function(x){ return trim(x).toLowerCase(); }).join('|');
    return joined.indexOf('класс') >= 0 ||
           joined.indexOf('бумаг') >= 0 ||
           joined.indexOf('название') >= 0 ||
           joined.indexOf('тикер') >= 0 ||
           joined.indexOf('рыночная стоимость') >= 0 ||
           joined.indexOf('инвестировано') >= 0 ||
           joined.indexOf('p/l') >= 0 ||
           joined.indexOf('доля портфеля') >= 0 ||
           joined.indexOf('показатель') >= 0 ||
           joined.indexOf('метрика') >= 0;
  }

  function findHeaderAnywhere(headerText, startCol) {
    startCol = startCol || 1;
    var lastRow = Math.max(sh.getLastRow(), 1);
    var lastCol = Math.max(sh.getLastColumn(), 1);
    if (startCol > lastCol) return null;

    var vals = sh.getRange(1, startCol, lastRow, lastCol - startCol + 1).getDisplayValues();
    var target = trim(headerText);

    for (var r = 0; r < vals.length; r++) {
      for (var c = 0; c < vals[r].length; c++) {
        if (trim(vals[r][c]) === target) {
          return { row: r + 1, col: startCol + c };
        }
      }
    }
    return null;
  }

  function readSectionTableAnywhere(headerText, startCol, maxWidth) {
    var pos = findHeaderAnywhere(headerText, startCol || extCol);
    if (!pos) return [];

    var lastRow = Math.max(sh.getLastRow(), 1);
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var width = Math.min(maxWidth || 8, lastCol - pos.col + 1);
    if (width < 1) return [];

    var out = [];
    var started = false;

    for (var rr = pos.row + 1; rr <= lastRow; rr++) {
      var rowVals = sh.getRange(rr, pos.col, 1, width).getDisplayValues()[0];
      var first = trim(rowVals[0]);
      var nonEmpty = false;

      for (var i = 0; i < rowVals.length; i++) {
        if (trim(rowVals[i])) {
          nonEmpty = true;
          break;
        }
      }

      if (!nonEmpty) {
        if (started) break;
        continue;
      }

      if (isKnownHeader(first)) break;

      started = true;
      out.push(rowVals);
    }

    return out;
  }

  function appendOldBondBlocks(lines) {
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
  }

  // Новые секции справа
  var portfolioTotal = readSectionTableAnywhere('Портфель — итого', extCol, 8);
  var assetClasses   = readSectionTableAnywhere('Классы активов', extCol, 8);
  var sharesTopPos   = readSectionTableAnywhere('Акции — топ-5 по позиции', extCol, 8);
  var sharesTopPL    = readSectionTableAnywhere('Акции — топ-5 по P/L', extCol, 8);

  var hasNewSections = !!(
    portfolioTotal.length ||
    assetClasses.length ||
    sharesTopPos.length ||
    sharesTopPL.length
  );

  var lines = [];

  // Если новые секции не найдены — старый текст
  if (!hasNewSections) {
    appendOldBondBlocks(lines);
    lines.push(_historySummaryText_(_readHistoryABC_(sh)));
    if (!lines.length) lines.push('Статистика на Dashboard не найдена.');
    tgSendMessage_(lines.join('\n').replace(/\n{3,}/g, '\n\n'));
    return;
  }

  // Портфель
  if (portfolioTotal.length) {
    lines.push('Портфель:');

    var wantedPortfolio = ['Инвестировано', 'Рыночная стоимость', 'P/L (руб)', 'P/L (%)'];
    var portfolioMap = {};

    portfolioTotal.forEach(function(r) {
      if (looksLikeHeaderRow(r)) return;
      var key = trim(r[0]);
      if (key) portfolioMap[key] = firstNonEmptyAfter0(r);
    });

    wantedPortfolio.forEach(function(key) {
      if (portfolioMap[key] != null) lines.push('- ' + key + ' — ' + portfolioMap[key]);
    });

    Object.keys(portfolioMap).forEach(function(key) {
      if (wantedPortfolio.indexOf(key) === -1) {
        lines.push('- ' + key + ' — ' + portfolioMap[key]);
      }
    });

    lines.push('');
  }

  // Классы активов
  if (assetClasses.length) {
    lines.push('Классы активов:');

    var wantedClasses = ['Облигации', 'Фонды', 'Акции', 'Опционы'];
    var classMap = {};

    assetClasses.forEach(function(r) {
      if (looksLikeHeaderRow(r)) return;

      var key = trim(r[0]);
      if (!key) return;

      var count = trim(r[1]);
      var market = trim(r[3]);
      var share = trim(r[6]);

      var parts = [];
      if (market) parts.push('стоимость ' + market);
      if (share) parts.push('доля ' + share);
      if (!parts.length && count) parts.push('бумаг ' + count);
      if (!parts.length) parts.push(firstNonEmptyAfter0(r));

      classMap[key] = parts.join(', ');
    });

    wantedClasses.forEach(function(key) {
      if (classMap[key] != null) lines.push('- ' + key + ' — ' + classMap[key]);
    });

    Object.keys(classMap).forEach(function(key) {
      if (wantedClasses.indexOf(key) === -1) {
        lines.push('- ' + key + ' — ' + classMap[key]);
      }
    });

    lines.push('');
  }

  // Акции
  if (sharesTopPos.length || sharesTopPL.length) {
    lines.push('Акции:');

    var shareMap = {};
    var shareOrder = [];

    function ensureShare(name, ticker) {
      var key = name;
      if (!shareMap[key]) {
        shareMap[key] = {
          ticker: ticker || '',
          pos: '—',
          pl: '—'
        };
        shareOrder.push(key);
      } else if (!shareMap[key].ticker && ticker) {
        shareMap[key].ticker = ticker;
      }
      return key;
    }

    sharesTopPos.forEach(function(r) {
      if (looksLikeHeaderRow(r)) return;

      var name = trim(r[0]);
      var ticker = trim(r[1]);
      if (!name) return;

      var key = ensureShare(name, ticker);
      var market = trim(r[2]);
      var plRub = trim(r[3]);
      var plPct = trim(r[4]);

      var plText = '—';
      if (plRub && plPct) plText = plRub + ' (' + plPct + ')';
      else if (plRub) plText = plRub;
      else if (plPct) plText = plPct;

      if (market) shareMap[key].pos = market;
      if (plText !== '—') shareMap[key].pl = plText;
    });

    sharesTopPL.forEach(function(r) {
      if (looksLikeHeaderRow(r)) return;

      var name = trim(r[0]);
      var ticker = trim(r[1]);
      if (!name) return;

      var key = ensureShare(name, ticker);
      var market = trim(r[2]);
      var plRub = trim(r[3]);
      var plPct = trim(r[4]);

      var plText = '—';
      if (plRub && plPct) plText = plRub + ' (' + plPct + ')';
      else if (plRub) plText = plRub;
      else if (plPct) plText = plPct;

      if (market && shareMap[key].pos === '—') shareMap[key].pos = market;
      if (plText !== '—') shareMap[key].pl = plText;
    });

    shareOrder.slice(0, 5).forEach(function(name, i) {
      var item = shareMap[name] || {};
      var label = name;
      if (item.ticker) label += ' (' + item.ticker + ')';
      lines.push((i + 1) + '. ' + label + ' — позиция ' + (item.pos || '—') + ', P/L ' + (item.pl || '—'));
    });

    lines.push('');
  }

  // Старые bond-блоки
  appendOldBondBlocks(lines);

  if (!lines.length) lines.push('Статистика на Dashboard не найдена.');
  tgSendMessage_(lines.join('\n').replace(/\n{3,}/g, '\n\n'));
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
