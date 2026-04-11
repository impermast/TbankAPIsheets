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
  if (typeof buildPortfolioDashboard === 'function') {
    buildPortfolioDashboard();
  } else {
    buildBondsDashboard();
  }

  SpreadsheetApp.flush();
  Utilities.sleep(1200);

  sendDashboardGreets_();
  sendDashboardStatsFromDashboard_();

  SpreadsheetApp.flush();
  Utilities.sleep(500);

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
    var g = String(vals[r][6] || '').trim();
    var h = String(vals[r][7] || '').trim();
    var i = String(vals[r][8] || '').trim();
    if (g === 'Риск' && h === 'YTM (%)' && i === 'Тултип') {
      return { row: r + 1, col: 7 };
    }
  }
  return null;
}

function _dashboardChartRangeLooksValid_(range, sheet) {
  try {
    if (!range || !sheet) return false;
    if (range.getSheet().getName() !== sheet.getName()) return false;
    if (range.getRow() < 1 || range.getColumn() < 1) return false;
    if (range.getNumRows() < 2 || range.getNumColumns() < 2) return false;
    if (range.getRow() > sheet.getLastRow() || range.getColumn() > sheet.getLastColumn()) return false;

    var sampleRows = Math.min(range.getNumRows(), 8);
    var sampleCols = Math.min(range.getNumColumns(), 4);
    var values = range.offset(0, 0, sampleRows, sampleCols).getDisplayValues();
    var hasHeader = false;
    var hasData = false;
    var rr;
    var cc;

    for (cc = 0; cc < values[0].length; cc++) {
      if (String(values[0][cc] || '').trim()) {
        hasHeader = true;
        break;
      }
    }

    for (rr = 1; rr < values.length; rr++) {
      for (cc = 0; cc < values[rr].length; cc++) {
        if (String(values[rr][cc] || '').trim()) {
          hasData = true;
          break;
        }
      }
      if (hasData) break;
    }

    return hasHeader && hasData;
  } catch (e) {
    Logger.log('_dashboardChartRangeLooksValid_ error: ' + (e && e.message ? e.message : e));
    return false;
  }
}

function _dashboardChartLooksValid_(chart, sheet) {
  try {
    if (!chart || !sheet) return false;
    var ranges = chart.getRanges();
    if (!ranges || !ranges.length) return false;

    for (var i = 0; i < ranges.length; i++) {
      if (_dashboardChartRangeLooksValid_(ranges[i], sheet)) return true;
    }
    return false;
  } catch (e) {
    Logger.log('_dashboardChartLooksValid_ error: ' + (e && e.message ? e.message : e));
    return false;
  }
}

function exportDashboardCharts_(aliases) {
  SpreadsheetApp.flush();

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Dashboard');
  if (!sh) return {};

  var out = {};
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');

  var registry = {};
  Object.keys(DASHBOARD_CHARTS).forEach(function(k){ registry[k] = DASHBOARD_CHARTS[k]; });

  // динамические (не в фиксированном реестре)
  registry.ytmVsCoupon = { topLeft: _guessCmpTopLeft_(sh),     caption: 'YTM vs Купонная доходность (средневзв., %)' };
  registry.scatter     = { topLeft: _guessScatterTopLeft_(sh), caption: 'Риск vs Доходность к погашению (YTM)' };

  var keys = (aliases && aliases.length) ? aliases : Object.keys(registry);
  var lastRow = Math.max(sh.getLastRow(), 1);
  var lastCol = Math.max(sh.getLastColumn(), 1);

  keys.forEach(function(alias){
    out[alias] = null;

    try {
      var cfg = registry[alias];
      if (!cfg || !cfg.topLeft) return;
      if (cfg.topLeft.row < 1 || cfg.topLeft.col < 1) return;
      if (cfg.topLeft.row > lastRow || cfg.topLeft.col > lastCol) return;

      var chart = _findChartByTopLeft_(sh, cfg.topLeft.row, cfg.topLeft.col);
      if (!chart) return;
      if (!_dashboardChartLooksValid_(chart, sh)) return;

      var blob = chart.getAs('image/png');
      if (!blob) return;

      out[alias] = {
        blob: blob.setName('dashboard_' + alias + '_' + ts + '.png'),
        caption: cfg.caption
      };
    } catch (e) {
      Logger.log('exportDashboardCharts_ [' + alias + '] error: ' + (e && e.stack ? e.stack : e));
      out[alias] = null;
    }
  });

  return out;
}

// -------------------- Секции (колонка T) --------------------

function _readSectionTableT_(sheet, headerText, extCol, maxWidth) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  extCol = extCol || 20;
  maxWidth = maxWidth || 6;

  var colVals = sheet.getRange(1, extCol, lastRow, 1).getValues();
  var headerRow = -1;
  for (var r = 0; r < colVals.length; r++) {
    if (String(colVals[r][0]).trim() === String(headerText)) { headerRow = r + 1; break; }
  }
  if (headerRow < 1) return [];

  var headerLine = sheet.getRange(headerRow, extCol, 1, maxWidth).getValues()[0];
  var nonEmpty = 0;
  for (var i = 0; i < maxWidth; i++) if (String(headerLine[i] || '').trim()) nonEmpty++;

  var dataStartRow = nonEmpty > 1 ? headerRow + 1 : headerRow + 2;
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
  tgSendMessage_('Dashboard обновлён: ' + now);
}

function sendDashboardStatsFromDashboard_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  if (!sh) { tgSendMessage_('Лист Dashboard не найден.'); return; }

  var extCol = 20;
  var messages = [];

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

  function formatPl(plRub, plPct) {
    if (trim(plRub) && trim(plPct)) return trim(plRub) + ' (' + trim(plPct) + ')';
    if (trim(plRub)) return trim(plRub);
    if (trim(plPct)) return trim(plPct);
    return '—';
  }

  function pushMessage(title, lines) {
    var body = [];
    var i;

    if (title) body.push(title);
    for (i = 0; i < (lines || []).length; i++) {
      if (trim(lines[i])) body.push(lines[i]);
    }

    var text = body.join('\n').replace(/\n{3,}/g, '\n\n').trim();
    if (text) messages.push(text);
  }

  function sendMessages() {
    if (!messages.length) {
      tgSendMessage_('Статистика на Dashboard не найдена.');
      return;
    }

    for (var i = 0; i < messages.length; i++) {
      try {
        tgSendMessage_(messages[i]);
      } catch (e) {
        Logger.log('sendDashboardStatsFromDashboard_ message[' + i + '] error: ' + (e && e.stack ? e.stack : e));
      }
    }
  }

  function isKnownHeader(text) {
    text = trim(text);
    return text === 'Портфель — итого' ||
      text === 'Классы активов' ||
      text === 'Акции — топ-5 по позиции' ||
      text === 'Акции — топ-5 по P/L' ||
      text === 'Акции — топ-5 убыток' ||
      text === 'Облигации — top YTM' ||
      text === 'Облигации — top P/L' ||
      text === 'Облигации — worst P/L' ||
      text === 'Фонды — top-5 по позиции' ||
      text === 'Фонды — top-5 по P/L' ||
      text === DASHBOARD_SECTIONS.topYtm.header ||
      text === DASHBOARD_SECTIONS.bestPL.header ||
      text === DASHBOARD_SECTIONS.worstPL.header ||
      text === DASHBOARD_SECTIONS.worstSec.header;
  }

  function looksLikeHeaderRow(row) {
    var joined = row.map(function(x) { return trim(x).toLowerCase(); }).join('|');
    return joined.indexOf('класс') >= 0 ||
      joined.indexOf('бумаг') >= 0 ||
      joined.indexOf('название') >= 0 ||
      joined.indexOf('тикер') >= 0 ||
      joined.indexOf('рыночная стоимость') >= 0 ||
      joined.indexOf('инвестировано') >= 0 ||
      joined.indexOf('p/l') >= 0 ||
      joined.indexOf('доля') >= 0 ||
      joined.indexOf('показатель') >= 0 ||
      joined.indexOf('метрика') >= 0 ||
      joined.indexOf('ytm') >= 0 ||
      joined.indexOf('до погашения') >= 0 ||
      joined.indexOf('figi') >= 0;
  }

  function findHeaderAnywhere(headerText) {
    var lastRow = Math.max(sh.getLastRow(), 1);
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var vals = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var target = trim(headerText);

    for (var r = 0; r < vals.length; r++) {
      for (var c = 0; c < vals[r].length; c++) {
        if (trim(vals[r][c]) === target) {
          return { row: r + 1, col: c + 1 };
        }
      }
    }
    return null;
  }

  function readSectionTableAnywhere(headerText, maxWidth) {
    var pos = findHeaderAnywhere(headerText);
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

  function dataRows(rows, limit) {
    var out = [];
    var i;
    if (!rows || !rows.length) return out;

    for (i = 0; i < rows.length; i++) {
      if (looksLikeHeaderRow(rows[i])) continue;
      if (!trim(rows[i][0])) continue;
      out.push(rows[i]);
      if (limit && out.length >= limit) break;
    }
    return out;
  }

  function buildHistoryLines(points, includeTotals) {
    if (!points || !points.length) return [];

    var p = points.filter(function(x) { return x.market != null; });
    if (!p.length) return [];

    var lines = [];
    var last = p[p.length - 1];
    var prev = p.length >= 2 ? p[p.length - 2] : null;
    var first = p[0];
    var lastDt = Utilities.formatDate(last.dt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

    lines.push('Последняя точка: ' + lastDt);

    if (includeTotals && last.market != null) {
      lines.push('Стоимость: ' + _formatMoney_(last.market) + ' ₽');
      if (last.invested != null) lines.push('Инвестировано: ' + _formatMoney_(last.invested) + ' ₽');
    }

    if (prev && prev.market != null) {
      var d1 = last.market - prev.market;
      var p1 = prev.market !== 0 ? (d1 / prev.market * 100) : null;
      lines.push('Δ к прошлому запуску: ' + _formatMoney_(d1) + ' ₽' + (p1 != null ? (' (' + _formatMoney_(p1) + '%)') : ''));
    }

    if (first && first.market != null && first !== last) {
      var d2 = last.market - first.market;
      var p2 = first.market !== 0 ? (d2 / first.market * 100) : null;
      lines.push('Δ за окно истории: ' + _formatMoney_(d2) + ' ₽' + (p2 != null ? (' (' + _formatMoney_(p2) + '%)') : ''));
    }

    return lines;
  }

  var headerNames = [
    'Портфель — итого',
    'Классы активов',
    'Акции — топ-5 по позиции',
    'Акции — топ-5 по P/L',
    'Акции — топ-5 убыток',
    'Облигации — top YTM',
    'Облигации — top P/L',
    'Облигации — worst P/L',
    'Фонды — top-5 по позиции',
    'Фонды — top-5 по P/L'
  ];

  var newLayoutFound = false;
  for (var h = 0; h < headerNames.length; h++) {
    if (findHeaderAnywhere(headerNames[h])) {
      newLayoutFound = true;
      break;
    }
  }

  if (!newLayoutFound) {
    var topYtmOld = _readSectionTableT_(sh, DASHBOARD_SECTIONS.topYtm.header, extCol, 6);
    var bestPLOld = _readSectionTableT_(sh, DASHBOARD_SECTIONS.bestPL.header, extCol, 6);
    var worstPLOld = _readSectionTableT_(sh, DASHBOARD_SECTIONS.worstPL.header, extCol, 6);
    var worstSecOld = _readSectionTableT_(sh, DASHBOARD_SECTIONS.worstSec.header, extCol, 6);

    var oldBondLines = [];
    var oldRiskLines = [];

    if (topYtmOld.length) {
      oldBondLines.push('Top YTM:');
      topYtmOld.slice(0, 3).forEach(function(r, i) {
        var name = String(r[0] || '');
        var ytm = (r[1] != null && r[1] !== '') ? Number(r[1]).toFixed(2) : '—';
        var yrs = (r[2] != null && r[2] !== '') ? Number(r[2]).toFixed(2) : '—';
        oldBondLines.push((i + 1) + '. ' + name + ' — YTM ' + ytm + '%, до погашения ' + yrs + ' г.');
      });
    }

    if (bestPLOld.length) {
      if (oldBondLines.length) oldBondLines.push('');
      oldBondLines.push('Лучший P/L:');
      bestPLOld.slice(0, 3).forEach(function(r, i) {
        var name = String(r[0] || '');
        var cup = (r[1] != null && r[1] !== '') ? Number(r[1]).toFixed(2) : '—';
        var pl = (r[2] != null && r[2] !== '') ? Number(r[2]).toFixed(2) : '—';
        oldBondLines.push((i + 1) + '. ' + name + ' — купон ' + cup + '%, P/L ' + pl + '%');
      });
    }

    if (worstPLOld.length) {
      oldRiskLines.push('Худший P/L:');
      worstPLOld.slice(0, 3).forEach(function(r, i) {
        var name = String(r[0] || '');
        var cup = (r[1] != null && r[1] !== '') ? Number(r[1]).toFixed(2) : '—';
        var pl = (r[2] != null && r[2] !== '') ? Number(r[2]).toFixed(2) : '—';
        oldRiskLines.push((i + 1) + '. ' + name + ' — купон ' + cup + '%, P/L ' + pl + '%');
      });
    }

    if (worstSecOld.length) {
      if (oldRiskLines.length) oldRiskLines.push('');
      oldRiskLines.push('Слабые сектора:');
      worstSecOld.slice(0, 3).forEach(function(r) {
        var sec = String(r[0] || '');
        var pl = (r[1] != null && r[1] !== '') ? Number(r[1]).toFixed(2) : '0.00';
        oldRiskLines.push('- ' + sec + ': ' + pl + ' ₽');
      });
    }

    pushMessage('Облигации', oldBondLines);
    pushMessage('Риски и просадка', oldRiskLines);
    pushMessage('Динамика портфеля', buildHistoryLines(_readHistoryABC_(sh), true));
    sendMessages();
    return;
  }

  var portfolioTotal = readSectionTableAnywhere('Портфель — итого', 8);
  var assetClasses = readSectionTableAnywhere('Классы активов', 8);
  var sharesTopPos = readSectionTableAnywhere('Акции — топ-5 по позиции', 5);
  var sharesTopPL = readSectionTableAnywhere('Акции — топ-5 по P/L', 5);
  var sharesTopLoss = readSectionTableAnywhere('Акции — топ-5 убыток', 5);
  var bondsTopYtm = readSectionTableAnywhere('Облигации — top YTM', 4);
  var bondsTopPL = readSectionTableAnywhere('Облигации — top P/L', 3);
  var bondsWorstPL = readSectionTableAnywhere('Облигации — worst P/L', 3);
  var fundsTopPos = readSectionTableAnywhere('Фонды — top-5 по позиции', 5);
  var fundsTopPL = readSectionTableAnywhere('Фонды — top-5 по P/L', 5);

  var summaryLines = [];
  var portfolioRows = dataRows(portfolioTotal);
  var classRows = dataRows(assetClasses);

  if (portfolioRows.length) {
    var portfolioMap = {};
    portfolioRows.forEach(function(r) {
      portfolioMap[trim(r[0])] = firstNonEmptyAfter0(r);
    });

    if (portfolioMap['Инвестировано']) summaryLines.push('Инвестировано: ' + portfolioMap['Инвестировано']);
    if (portfolioMap['Рыночная стоимость']) summaryLines.push('Стоимость: ' + portfolioMap['Рыночная стоимость']);

    var plParts = [];
    if (portfolioMap['P/L (руб)']) plParts.push(portfolioMap['P/L (руб)']);
    if (portfolioMap['P/L (%)']) plParts.push(portfolioMap['P/L (%)']);
    if (plParts.length) summaryLines.push('P/L: ' + plParts.join(' | '));
  }

  if (classRows.length) {
    if (summaryLines.length) summaryLines.push('');
    summaryLines.push('Классы активов:');
    classRows.slice(0, 4).forEach(function(r) {
      var label = trim(r[0]);
      var market = trim(r[3]);
      var share = trim(r[6]);
      var count = trim(r[1]);
      var parts = [];

      if (market) parts.push('стоимость ' + market);
      if (share) parts.push('доля ' + share);
      if (!parts.length && count) parts.push('бумаг ' + count);
      if (!parts.length) parts.push(firstNonEmptyAfter0(r));

      summaryLines.push('- ' + label + ' — ' + parts.join(', '));
    });
  }

  pushMessage('Сводка по портфелю', summaryLines);

  var keyLines = [];
  var sharesPosRows = dataRows(sharesTopPos, 3);
  var sharesPlRows = dataRows(sharesTopPL, 3);
  var sharesLossRows = dataRows(sharesTopLoss, 3);
  var fundsPosRows = dataRows(fundsTopPos, 3);
  var fundsPlRows = dataRows(fundsTopPL, 3);

  if (sharesPosRows.length) {
    keyLines.push('Акции — крупнейшие позиции:');
    sharesPosRows.forEach(function(r, i) {
      var label = trim(r[0]);
      if (trim(r[1])) label += ' (' + trim(r[1]) + ')';
      keyLines.push((i + 1) + '. ' + label + ' — позиция ' + (trim(r[2]) || '—') + ', P/L ' + formatPl(r[3], r[4]));
    });
  } else if (sharesPlRows.length) {
    keyLines.push('Акции — лучший P/L:');
    sharesPlRows.forEach(function(r, i) {
      var label = trim(r[0]);
      if (trim(r[1])) label += ' (' + trim(r[1]) + ')';
      keyLines.push((i + 1) + '. ' + label + ' — P/L ' + formatPl(r[3], r[4]));
    });
  }

  if (sharesLossRows.length) {
    if (keyLines.length) keyLines.push('');
    keyLines.push('Акции — убыток:');
    sharesLossRows.forEach(function(r, i) {
      var label = trim(r[0]);
      if (trim(r[1])) label += ' (' + trim(r[1]) + ')';
      keyLines.push((i + 1) + '. ' + label + ' — P/L ' + formatPl(r[3], r[4]));
    });
  }

  if (fundsPosRows.length) {
    if (keyLines.length) keyLines.push('');
    keyLines.push('Фонды — крупнейшие позиции:');
    fundsPosRows.forEach(function(r, i) {
      var label = trim(r[0]);
      if (trim(r[1])) label += ' (' + trim(r[1]) + ')';
      keyLines.push((i + 1) + '. ' + label + ' — позиция ' + (trim(r[2]) || '—') + ', P/L ' + formatPl(r[3], r[4]));
    });
  } else if (fundsPlRows.length) {
    if (keyLines.length) keyLines.push('');
    keyLines.push('Фонды — лучший P/L:');
    fundsPlRows.forEach(function(r, i) {
      var label = trim(r[0]);
      if (trim(r[1])) label += ' (' + trim(r[1]) + ')';
      keyLines.push((i + 1) + '. ' + label + ' — P/L ' + formatPl(r[3], r[4]));
    });
  }

  pushMessage('Ключевые позиции', keyLines);

  var bondLines = [];
  var bondsYtmRows = dataRows(bondsTopYtm, 3);
  var bondsTopPlRows = dataRows(bondsTopPL, 3);
  var bondsWorstPlRows = dataRows(bondsWorstPL, 3);

  if (bondsYtmRows.length) {
    bondLines.push('Top YTM:');
    bondsYtmRows.forEach(function(r, i) {
      var line = (i + 1) + '. ' + trim(r[0]) + ' — YTM ' + (trim(r[1]) || '—');
      if (trim(r[2])) line += ', до погашения ' + trim(r[2]);
      bondLines.push(line);
    });
  }

  if (bondsTopPlRows.length) {
    if (bondLines.length) bondLines.push('');
    bondLines.push('Лучший P/L:');
    bondsTopPlRows.forEach(function(r, i) {
      bondLines.push((i + 1) + '. ' + trim(r[0]) + ' — P/L ' + formatPl(r[1], r[2]));
    });
  }

  if (bondsWorstPlRows.length) {
    if (bondLines.length) bondLines.push('');
    bondLines.push('Худший P/L:');
    bondsWorstPlRows.forEach(function(r, i) {
      bondLines.push((i + 1) + '. ' + trim(r[0]) + ' — P/L ' + formatPl(r[1], r[2]));
    });
  }

  pushMessage('Облигации', bondLines);

  var historyLines = buildHistoryLines(_readHistoryABC_(sh), !summaryLines.length);
  pushMessage('Динамика портфеля', historyLines);

  if (!messages.length) {
    pushMessage('', [_historySummaryText_(_readHistoryABC_(sh))]);
  }

  sendMessages();
}

function sendDashboardChartsToTelegram_(aliases) {
  SpreadsheetApp.flush();

  var pack = exportDashboardCharts_(aliases);
  var sent = 0;
  var order = (aliases && aliases.length) ? aliases.slice() : Object.keys(pack);

  order.forEach(function(alias) {
    var item = pack[alias];
    if (!item || !item.blob) return;

    try {
      tgSendPhoto_(item.blob, item.caption || '');
      sent++;
    } catch (e) {
      Logger.log('sendDashboardChartsToTelegram_ [' + alias + '] error: ' + (e && e.stack ? e.stack : e));
    }
  });

  return sent;
}
