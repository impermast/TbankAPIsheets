/**
 * funds_update.gs
 * Лист «Funds»: новый порядок колонок под требования и минимально нужные поля из API.
 * - Убрали «Риск (ручн.)»
 * - Оставили «ИИ», «Заметка»
 * - FIGI → Тикер → Название (asset.brand.name)
 * - Кол-во → Средняя цена → Текущая цена
 * - Комиссии фонда (TER, %) → Заблокирован (TCS)
 * - Расчётные: Инвестировано, Рыночная стоимость, P/L (руб), P/L (%)
 * - Фокус (asset.security.etf.focusType)
 * - Сектор (asset.brand.sector)
 * - Валюта (asset.security.etf.nominalCurrency)
 * - Информация бренда (asset.brand.info) — перенос строк и ограниченная ширина
 */

const FUNDS_SHEET = 'Funds';

const FUNDS_HEADERS = [
  // Блок 1 — приоритетные
  'ИИ',
  'Заметка',

  // Блок 2 — идентификация
  'FIGI',
  'Тикер',
  'Название',

  // Блок 3 — базовые числа
  'Кол-во',
  'Средняя цена',
  'Текущая цена',

  // Блок 4 — атрибуты и флаги
  'Комиссии фонда (TER, %)',
  'Заблокирован (TCS)',

  // Блок 5 — расчётные
  'Инвестировано',
  'Рыночная стоимость',
  'P/L (руб)',
  'P/L (%)',

  // Блок 6 — ETF/бренд
  'Фокус',
  'Сектор',
  'Валюта',
  'Информация бренда'
];

// ===== Публичные действия =====
function updateFundsFull(){
  setStatus_('Funds • полное обновление…');
  var figis = readInputFigisByType_('etf');
  if (!figis.length) { showSnack_('Нет FIGI в Input!B (Фонды)','Funds',2500); return; }

  // 1) Создать/очистить лист и заголовки
  var sh = SpreadsheetApp.getActive().getSheetByName(FUNDS_SHEET) || SpreadsheetApp.getActive().insertSheet(FUNDS_SHEET);
  sh.clear();
  sh.getRange(1,1,1,FUNDS_HEADERS.length).setValues([FUNDS_HEADERS]);
  sh.setFrozenRows(1);

  // 2) Портфель: qty/avg
  var pfMap = safeFetchAllPortfolios_(); // {figi:[{accountId,accountName,qty,avg,avg_fifo}]}

  // 3) Инфо ETF + Asset/Brand и маркет-данные
  var infoBy = fetchFundsInfo_(figis);       // instrument + asset.brand + etf
  var mdBy   = fetchFundsMarketData_(figis); // last price

  // 4) Строки
  var rows = [];
  figis.forEach(function(figi){
    var pf = (pfMap[figi]||[]);
    var totalQty = pf.reduce(function(s,x){ return s + (Number(x.qty)||0); }, 0);
    var avg = (function(){
      for (var i=0;i<pf.length;i++){
        var v = (pf[i].avg != null) ? Number(pf[i].avg) : (pf[i].avg_fifo != null ? Number(pf[i].avg_fifo) : null);
        if (v != null) return v;
      }
      return null;
    })();
    var bi = infoBy[figi] || {};
    var md = mdBy[figi]   || {};

    rows.push([
      '',                               // ИИ
      '',                               // Заметка

      figi,                             // FIGI
      (bi.ticker || ''),                // Тикер
      (bi.brandName || bi.name || ''),  // Название (brand.name приоритет)

      (totalQty != null ? totalQty : ''),                 // Кол-во
      (avg != null ? avg : ''),                           // Средняя
      (md.lastPrice != null ? md.lastPrice : ''),        // Текущая

      (bi.terPct != null ? bi.terPct : ''),              // Комиссии фонда (TER, %)
      (bi.blockedTcs ? 'Да' : 'Нет'),                    // Заблокирован (TCS)

      '', '', '', '',                                    // расчётные — формулы ниже

      (bi.focus || ''),                                  // Фокус (etf.focusType)
      (bi.brandSector || ''),                            // Сектор (brand.sector)
      (bi.nominalCurrency || ''),                        // Валюта (etf.nominalCurrency)
      (bi.brandInfo || '')                               // Информация бренда
    ]);
  });

  if (!rows.length) { showSnack_('Пусто — нечего обновлять','Funds',2000); return; }

  // sanity-check
  var expectedCols = FUNDS_HEADERS.length;
  var bad = rows.findIndex(function(r){ return r.length !== expectedCols; });
  if (bad !== -1) throw new Error('Ширина строки №'+(bad+2)+' = '+rows[bad].length+' != '+expectedCols+' (FUNDS_HEADERS)');

  sh.getRange(2,1,rows.length, expectedCols).setValues(rows);

  // Формулы + автоформатирование
  applyFundsFormulas_(sh, 2, rows.length);
  formatFundsSheet_(sh);

  showSnack_('Готово: Funds обновлён ('+rows.length+' строк)','Funds',2500);
}

function updateFundPricesOnly(){
  setStatus_('Funds • только цены…');
  var sh = SpreadsheetApp.getActive().getSheetByName(FUNDS_SHEET);
  if (!sh) { showSnack_('Лист Funds не найден','Funds',2000); return; }
  var last = sh.getLastRow();
  if (last < 2) { showSnack_('Нет данных для обновления цен','Funds',2000); return; }

  var figis = sh.getRange(2, FUNDS_HEADERS.indexOf('FIGI')+1, last-1, 1).getValues().flat().filter(String);
  if (!figis.length) { showSnack_('FIGI не найдены','Funds',2000); return; }

  var mdBy = fetchFundsMarketData_(figis);
  var colPrice = FUNDS_HEADERS.indexOf('Текущая цена')+1;
  var priceArr = figis.map(function(f){
    var md = mdBy[f] || {};
    return [md.lastPrice != null ? md.lastPrice : ''];
  });
  sh.getRange(2, colPrice, priceArr.length, 1).setValues(priceArr);

  // Пересчитать формулы и обновить форматирование
  applyFundsFormulas_(sh, 2, last - 1);

  showSnack_('Цены обновлены и формулы пересчитаны','Funds • Prices',2000);
}

// ===== Helpers: API → данные =====
function fetchFundsInfo_(figis){
  var out = {};
  figis.forEach(function(figi){
    var inst = callInstrumentsEtfByFigi_(figi) || {};
    var assetUid = inst.assetUid || inst.asset_uid || null;

    var asset = null, etf = null, brand = null;
    if (assetUid && typeof callInstrumentsGetAssetByUid_ === 'function'){
      try {
        asset = callInstrumentsGetAssetByUid_(assetUid) || null;
        brand = asset && asset.brand ? asset.brand : null;
        etf   = asset && asset.security && asset.security.etf ? asset.security.etf : null;
      } catch(e){}
    }

    // TER — основное: AssetEtf.total_expense (в %), fallback: instrument.expenseRatio
    var terPct = null;
    if (etf && etf.totalExpense) {
      terPct = (typeof moneyToNumber === 'function') ? moneyToNumber(etf.totalExpense) : qToNumber(etf.totalExpense);
    } else if (inst.expenseRatio != null) {
      terPct = Number(inst.expenseRatio);
    }

    out[figi] = {
      name: inst.name || inst.ticker || '',
      ticker: inst.ticker || '',
      blockedTcs: (inst.blockedTcaFlag === true),

      // ETF/Asset
      focus: etf ? (etf.focusType || '') : '',
      nominalCurrency: etf ? (etf.nominalCurrency || '') : '',
      brandName: brand ? (brand.name || '') : '',
      brandInfo: brand ? (brand.info || '') : '',
      brandSector: brand ? (brand.sector || '') : '',

      terPct: (terPct != null ? terPct : null)
    };
  });
  return out;
}

function fetchFundsMarketData_(figis){
  var out = {};
  var last = callMarketLastPrices_(figis);
  last.forEach(function(x){
    out[x.figi] = out[x.figi] || {};
    out[x.figi].lastPrice = (x.lastPrice != null) ? Number(x.lastPrice) : null;
    out[x.figi].lastTime  = x.time || '';
  });
  return out;
}

/** Формулы для расчётных столбцов */
function applyFundsFormulas_(sh, startRow, numRows){
  var idx = function(name){ return FUNDS_HEADERS.indexOf(name)+1; };

  // Определяем разделитель аргументов формул по локали
  var loc = (SpreadsheetApp.getActive().getSpreadsheetLocale() || '').toLowerCase();
  var SEP = /^(bg|cs|da|de|el|es|et|fi|fr|hr|hu|it|lt|lv|nl|pl|pt|ro|ru|sk|sl|sr|sv|tr)/.test(loc) ? ';' : ',';

  var cQty   = idx('Кол-во');
  var cAvg   = idx('Средняя цена');
  var cPrice = idx('Текущая цена');

  var cInv   = idx('Инвестировано');
  var cMkt   = idx('Рыночная стоимость');
  var cPL    = idx('P/L (руб)');
  var cPLPct = idx('P/L (%)');

  function d(from,to){ return from - to; }
  var R2 = function(expr){ return 'ROUND(' + expr + SEP + '2)'; };

  // Инвестировано = Qty * Avg
  sh.getRange(startRow, cInv, numRows, 1)
    .setFormulaR1C1(
      '=IF(OR(LEN(RC['+d(cQty,cInv)+'])=0' + SEP + 'LEN(RC['+d(cAvg,cInv)+'])=0)' + SEP +
      '""' + SEP +
      R2('RC['+d(cQty,cInv)+']*RC['+d(cAvg,cInv)+']') + ')'
    );

  // Рыночная = Qty * Price
  sh.getRange(startRow, cMkt, numRows, 1)
    .setFormulaR1C1(
      '=IF(OR(LEN(RC['+d(cQty,cMkt)+'])=0' + SEP + 'LEN(RC['+d(cPrice,cMkt)+'])=0)' + SEP +
      '""' + SEP +
      R2('RC['+d(cQty,cMkt)+']*RC['+d(cPrice,cMkt)+']') + ')'
    );

  // P/L (руб) = Рыночная - Инвестировано
  sh.getRange(startRow, cPL, numRows, 1)
    .setFormulaR1C1(
      '=IF(OR(LEN(RC['+d(cMkt,cPL)+'])=0' + SEP + 'LEN(RC['+d(cInv,cPL)+'])=0)' + SEP +
      '""' + SEP +
      R2('RC['+d(cMkt,cPL)+']-RC['+d(cInv,cPL)+']') + ')'
    );

  // P/L (%) = P/L / Инвестировано * 100
  sh.getRange(startRow, cPLPct, numRows, 1)
    .setFormulaR1C1(
      '=IF(OR(LEN(RC['+d(cPL,cPLPct)+'])=0' + SEP + 'LEN(RC['+d(cInv,cPLPct)+'])=0)' + SEP +
      '""' + SEP +
      R2('(RC['+d(cPL,cPLPct)+']/RC['+d(cInv,cPLPct)+'])*100') + ')'
    );
}

/** Форматирование таблицы: автоширина (кроме исключений), фикс. высота строк, ширина «Название» и «Информация бренда» */
function formatFundsSheet_(sh){
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var totalCols = FUNDS_HEADERS.length;
  var COL_NAME = FUNDS_HEADERS.indexOf('Название') + 1;
  var COL_INFO = FUNDS_HEADERS.indexOf('Информация бренда') + 1;
  if (COL_NAME <= 0 || COL_INFO <= 0) return;

  // 0) На всякий случай — применяем все предыдущие изменения
  SpreadsheetApp.flush();

  // 1) Авто-подбор ширины КРОМЕ «Название» и «Информация бренда»
  autoResizeExcept_(sh, totalCols, [COL_NAME, COL_INFO]);

  // 2) Фиксируем высоту строк: шапка ~32px, данные — 30px
  try { sh.setRowHeight(1, 32); } catch(e){}
  try { sh.setRowHeights(2, lastRow - 1, 30); } catch(e){}

  // 3) «Информация бренда»: перенос строк, фикс. ширина, выравнивание по верхнему краю
  var rngInfo = sh.getRange(2, COL_INFO, Math.max(0, lastRow - 1), 1);
  try { rngInfo.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); } catch(e) { rngInfo.setWrap(true); }
  sh.setColumnWidth(COL_INFO, 320);        // можно менять число
  rngInfo.setVerticalAlignment('top');

  // 4) «Название»: без переноса (CLIP), фикс. ширина, по центру по вертикали
  var rngName = sh.getRange(2, COL_NAME, Math.max(0, lastRow - 1), 1);
  try { rngName.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); } catch(e) { rngName.setWrap(false); }
  sh.setColumnWidth(COL_NAME, 220);        // можно менять число
  rngName.setVerticalAlignment('middle');

  // 5) Применяем все размеры немедленно
  SpreadsheetApp.flush();
}

/** Авто-подбор ширины для всех колонок, кроме указанных в exceptCols (массив индексов 1-based) */
function autoResizeExcept_(sh, totalCols, exceptCols){
  exceptCols = (exceptCols || []).slice().sort(function(a,b){ return a-b; });
  var from = 1;

  function resizeRange(start, end){
    var width = end - start + 1;
    if (width > 0) {
      try { sh.autoResizeColumns(start, width); } catch(e){}
    }
  }

  for (var i = 0; i < exceptCols.length; i++){
    var ex = exceptCols[i];
    if (ex > from) resizeRange(from, ex - 1);
    from = ex + 1;
  }
  if (from <= totalCols) resizeRange(from, totalCols);
}

