/**
 * dashboard_exports.gs
 * Унифицированный экспорт диаграмм с «Dashboard» и отправка в Telegram.
 *
 * Идентификация диаграмм идёт по верхнему левому углу диапазона данных,
 * который используется при построении в buildBondsDashboard():
 *   - Сектора:  row=1, col=7  (Dashboard!G1)
 *   - Купоны:   row=1, col=13 (Dashboard!M1)
 * Если поменяешь раскладку на листе — обнови координаты ниже.
 */

// Реестр экспортируемых диаграмм
var DASHBOARD_CHARTS = {
  sectors: { topLeft: { row: 1, col: 7 },  caption: 'Структура по секторам (рыночная стоимость)' },
  coupons: { topLeft: { row: 1, col: 13 }, caption: 'График купонных выплат (6 месяцев)' }
};

/** Найти диаграмму по верхнему левому углу её исходного диапазона данных. */
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

/**
 * Экспорт указанных диаграмм с Dashboard в PNG.
 * @param {string[]=} aliases — например ['sectors','coupons'] (по умолчанию — все из реестра)
 * @return {Object} map alias -> { blob: Blob, caption: string } | null
 */
function exportDashboardCharts_(aliases) {
  aliases = (aliases && aliases.length) ? aliases : Object.keys(DASHBOARD_CHARTS);
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Dashboard');
  if (!sh) throw new Error('Лист Dashboard не найден. Сначала выполните buildBondsDashboard().');

  var out = {};
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');

  aliases.forEach(function (alias) {
    var cfg = DASHBOARD_CHARTS[alias];
    if (!cfg) { out[alias] = null; return; }

    var chart = _findChartByTopLeft_(sh, cfg.topLeft.row, cfg.topLeft.col);
    if (!chart) { out[alias] = null; return; }

    var blob = chart.getAs('image/png').setName('dashboard_' + alias + '_' + ts + '.png');
    out[alias] = { blob: blob, caption: cfg.caption };
  });

  return out;
}

/**
 * Прочитать «экспорт-секцию» в колонке T (EXT_COL=20): возвращает массив строк (без заголовка).
 * Ожидается раскладка:
 *   T: Заголовок секции (строка)
 *   T+1..: шапка таблицы, далее строки таблицы.
 * Ищем секцию по T: "headerText".
 */
function readDashboardSection_(headerText){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Dashboard');
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  if (lastRow < 1) return [];

  var EXT_COL = 20; // T
  var MAX_W   = 6;  // читаем до 6 колонок вправо
  var vals = sh.getRange(1, EXT_COL, lastRow, MAX_W).getValues();

  // найти строку с headerText в первом столбце (T)
  var startRow = -1;
  for (var r=1; r<=lastRow; r++){
    var cell = sh.getRange(r, EXT_COL).getValue();
    if (String(cell).trim() === headerText){ startRow = r; break; }
  }
  if (startRow < 1) return [];

  // шапка на следующей строке; данные ещё ниже — до первой пустой строки в T
  var dataStart = startRow + 2;
  var out = [];
  for (var rr=dataStart; rr<=lastRow; rr++){
    var rowVals = sh.getRange(rr, EXT_COL, 1, MAX_W).getValues()[0];
    if (!String(rowVals[0]).trim()) break; // конец секции
    out.push(rowVals);
  }
  return out;
}

/** Приветственный текст — plain text, без markdown/html. */
function sendDashboardGreets_(){
  var tz = Session.getScriptTimeZone() || 'Etc/GMT';
  var now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var greetText = 'Добрый день! Короткий отчёт о портфеле на ' + now + '.\n' +
                  'Обновлены данные по бумагам и дашборду. Ниже — ключевые графики и показатели.';
  tgSendMessage_(greetText);
}

/**
 * Сформировать и отправить короткую статистику (plain text) из экспорт-секций:
 *  - Top YTM (5)
 *  - Top P/L (3) — best
 *  - Top P/L (3) — worst
 *  - P/L по секторам (TOP-худшие)
 */
function sendDashboardStatsFromDashboard_(){
  var topYtm   = readDashboardSection_('Top YTM (5)');                 // [Название, YTM, Years, FIGI]
  var bestPL   = readDashboardSection_('Top P/L (3) — best');          // [Название, Купон, P/L%]
  var worstPL  = readDashboardSection_('Top P/L (3) — worst');         // [Название, Купон, P/L%]
  var worstSec = readDashboardSection_('P/L по секторам (TOP-худшие)'); // [Сектор, PL]

  var lines = [];

  if (topYtm.length){
    lines.push('Топ-5 облигаций по YTM (%):');
    topYtm.forEach(function(r, i){
      var name = String(r[0]||'');
      var ytm  = r[1]!=null ? Number(r[1]).toFixed(2) : '—';
      var yrs  = r[2]!=null ? Number(r[2]).toFixed(2) : '—';
      lines.push((i+1)+'. '+name+' — YTM '+ytm+'%, до погашения '+yrs+' г.');
    });
    lines.push('');
  }

  if (bestPL.length){
    lines.push('Топ-3 прибыльные облигации (P/L, %):');
    bestPL.forEach(function(r,i){
      var name = String(r[0]||'');
      var cup  = r[1]!=null ? Number(r[1]).toFixed(2) : '—';
      var pl   = r[2]!=null ? Number(r[2]).toFixed(2) : '—';
      lines.push((i+1)+'. '+name+' — купон '+cup+'%, P/L '+pl+'%');
    });
    lines.push('');
  }

  if (worstPL.length){
    lines.push('Топ-3 убыточные облигации (P/L, %):');
    worstPL.forEach(function(r,i){
      var name = String(r[0]||'');
      var cup  = r[1]!=null ? Number(r[1]).toFixed(2) : '—';
      var pl   = r[2]!=null ? Number(r[2]).toFixed(2) : '—';
      lines.push((i+1)+'. '+name+' — купон '+cup+'%, P/L '+pl+'%');
    });
    lines.push('');
  }

  if (worstSec.length){
    lines.push('Просадка P/L по секторам (худшие):');
    worstSec.forEach(function(r){
      var sec = String(r[0]||'');
      var pl  = r[1]!=null ? Number(r[1]).toFixed(2) : '0';
      lines.push('- '+sec+': '+pl+' ₽');
    });
  }

  if (!lines.length) lines.push('Статистика на Dashboard не найдена.');
  tgSendMessage_(lines.join('\n'));
}

/**
 * Отправить выбранные диаграммы в Telegram (по умолчанию — все).
 * Требует tgbot_api.gs (tgSendPhoto_).
 */
function sendDashboardChartsToTelegram_(aliases) {
  var pack = exportDashboardCharts_(aliases);
  var sent = 0;
  Object.keys(pack).forEach(function (alias) {
    var item = pack[alias];
    if (item && item.blob) {
      tgSendPhoto_(item.blob, item.caption);
      sent++;
    }
  });
  return sent;
}

/** Комбо: перестроить дашборд и отправить ключевые графики. Удобно дергать из шедулера. */
function buildAndSendDashboardCharts_() {
  setStatus_('Обновляю Dashboard и отправляю графики…');
  buildBondsDashboard();
  var count = sendDashboardChartsToTelegram_(['sectors', 'coupons']);
  showSnack_('Отправлено диаграмм: ' + count, 'Telegram', 2500);
}
