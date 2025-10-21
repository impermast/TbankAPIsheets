/**
 * options_update.gs
 * Лист «Options»: создание/обновление, цены, формулы и форматирование.
 * Работает как с FIGI, так и с UID (для опционов).
 */

const OPTIONS_SHEET = 'Options';

const OPTIONS_HEADERS = [
  // Блок 1 — заметки
  'ИИ',
  'Заметка',

  // Блок 2 — идентификация
  'ID (FIGI/UID)',   // здесь храним ровно то, что пришло из Input (FIGI или uid:XXXX)
  'FIGI',
  'UID',
  'Тикер',
  'Название',

  // Блок 3 — спецификация опциона
  'Базовый актив',   // basicAsset / basicAssetUid / basicAssetPositionUid (что есть)
  'Направление',     // CALL / PUT
  'Стиль',           // AMERICAN / EUROPEAN (если дан)
  'Экспирация',      // ISO
  'Страйк',
  'Валюта',
  'Лот',

  // Блок 4 — позиции/цены
  'Кол-во',
  'Средняя цена',
  'Текущая цена',

  // Блок 5 — расчётные
  'Инвестировано',
  'Рыночная стоимость',
  'P/L (руб)',
  'P/L (%)'
];

// ========================= ПУБЛИЧНЫЕ ДЕЙСТВИЯ =========================
function calcAvgFromOpsForUid_(uid, daysBack){
  var fromIso = new Date(Date.now() - (daysBack||180)*24*3600*1000).toISOString();
  var toIso   = new Date().toISOString();
  var accs = callUsersGetAccounts_() || [];
  var bestAvg = null, bestQty = 0;

  accs.forEach(function(a){
    var cursor = null, qty = 0, avg = null;
    for(;;){
      var res = tinkoffFetch(
        'tinkoff.public.invest.api.contract.v1.OperationsService/GetOperationsByCursor',
        { accountId: a.accountId, instrumentId: uid, from: fromIso, to: toIso, limit: 1000, cursor: cursor },
        { allow404: true }
      ) || {};
      var items = res.items || res.operations || [];
      if (!items.length) break;

      for (var i=0;i<items.length;i++){
        var it = items[i];
        var typ = String(it.operationType || it.type || '').toUpperCase();
        var q   = qToNumber(it.quantityExecuted || it.quantity || it.lots);
        var px  = moneyToNumber(it.price);
        if (!px && q) {
          var pay = moneyToNumber(it.payment); // обычно со знаком
          if (pay) px = Math.abs(pay) / q;
        }
        if (!q) continue;

        if (typ.indexOf('BUY') >= 0){
          var newQty = qty + q;
          var newAvg = (qty>0 && avg!=null) ? ((avg*qty + px*q) / newQty) : px;
          qty = newQty; avg = newAvg;
        } else if (typ.indexOf('SELL') >= 0){
          qty = Math.max(0, qty - q);
          if (qty === 0) avg = null;
        }
      }
      if (!res.hasNext || !res.nextCursor) break;
      cursor = res.nextCursor;
    }
    if (qty>0 && avg!=null && qty>=bestQty){ bestQty = qty; bestAvg = avg; }
  });

  return (bestAvg!=null ? Number(bestAvg) : null);
}


function updateOptionsFull(){
  setStatus_('Options • полное обновление…');

  // Читаем колонку D «Опционы»: FIGI ИЛИ uid:XXXXXXXX… (как договорились)
  var ids = readInputFigisByType_('option');
  if (!ids || !ids.length) { showSnack_('Нет идентификаторов в Input!D (Опционы)','Options',2500); return; }

  // Собираем позиции по всем счётам: карта id -> [{qty, avg, avg_fifo}]
  var pfMapById = fetchAllPositionsByAnyId_();

  // Готовим лист
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(OPTIONS_SHEET) || ss.insertSheet(OPTIONS_SHEET);
  sh.clear();
  sh.getRange(1,1,1,OPTIONS_HEADERS.length).setValues([OPTIONS_HEADERS]);
  sh.setFrozenRows(1);

  var rows = [];

  ids.forEach(function(rawId){
    if (!rawId) return;

    var parsed = parseId_(rawId); // {id, kind:'FIGI'|'UID'}
    if (!parsed) return;

    // --- Инструмент: OptionBy по FIGI/UID
    var inst = null;
    try{
      if (parsed.kind === 'FIGI') inst = callInstrumentsOptionByFigi_(parsed.id);
      else                        inst = callInstrumentsOptionByUid_(parsed.id);
    }catch(_){}

    // Поля инструмента (без «угадываний»: только то, что пришло)
    var figi  = inst && (inst.figi || inst.instrumentFigi) || (parsed.kind==='FIGI' ? parsed.id : '');
    var uid   = inst && (inst.uid || inst.instrumentUid) || (parsed.kind==='UID'  ? parsed.id : '');
    var ticker= inst && inst.ticker || '';
    var name  = (inst && (inst.name || inst.ticker)) || '';

    var direction = mapOptionDirection_(inst && (inst.direction || inst.optionDirection || inst.kind));
    var style     = mapOptionStyle_    (inst && (inst.style     || inst.optionStyle));
    var expiry    = normalizeDateLike_(inst && (inst.expirationDate || inst.maturityDate || inst.expireDate || inst.lastTradeDate || inst.expiration));
    var strike    = numberFromMoneyOrQuotation_(inst && (inst.strikePrice || inst.strike));
    var lot       = inst ? (inst.lot != null ? inst.lot : inst.lotSize) : '';
    var currency  = (inst && (inst.currency || inst.buyCurrency || inst.sellCurrency)) ||
                    (inst && inst.strikePrice && (inst.strikePrice.currencyCode || inst.strikePrice.currency)) || '';

    var underlying = (inst && (inst.basicAsset || inst.underlying || inst.basic_asset)) ||
                     (inst && (inst.basicAssetUid || inst.basicAssetPositionUid)) || '';

    // Позиции по этому ID (FIGI или UID)
    var pfArr = pfMapById[parsed.id] || (figi ? pfMapById[figi] : []) || [];
    var totalQty = pfArr.reduce(function(s,x){ return s + (Number(x.qty)||0); }, 0);
    var avg = (function(){
      for (var i=0;i<pfArr.length;i++){
        var v = (pfArr[i].avg != null) ? Number(pfArr[i].avg) : (pfArr[i].avg_fifo != null ? Number(pfArr[i].avg_fifo) : null);
        if (v != null) return v;
      }
      return null;
    })();
    if ((avg==null || isNaN(avg))) avg = calcAvgFromOpsForUid_(uid, 365);


    rows.push([
      '',                 // ИИ
      '',                 // Заметка
      rawId,              // ID (FIGI/UID как в Input)
      figi || '',         // FIGI
      uid  || '',         // UID
      ticker,             // Тикер
      name,               // Название
      underlying,         // Базовый актив / UID базового актива
      direction,          // CALL/PUT
      style,              // Стиль
      expiry,             // Экспирация
      (strike != null ? strike : ''), // Страйк
      currency,           // Валюта
      (lot != null ? lot : ''),       // Лот
      (totalQty != null ? totalQty : ''), // Кол-во
      (avg != null ? avg : ''),           // Средняя цена
      '',                                  // Текущая цена (ниже проставим)
      '', '', '', ''                       // расчётные — формулы
    ]);
  });

  if (!rows.length){ showSnack_('Пусто — нечего обновлять','Options',2000); return; }

  var expected = OPTIONS_HEADERS.length;
  var bad = rows.findIndex(function(r){ return r.length !== expected; });
  if (bad !== -1) throw new Error('Ширина строки №'+(bad+2)+' = '+rows[bad].length+' != '+expected+' (OPTIONS_HEADERS)');

  sh.getRange(2,1,rows.length,expected).setValues(rows);

  // Цены + формулы
  updateOptionPricesOnly();
}

function updateOptionPricesOnly(){
  setStatus_('Options • только цены…');

  var sh = SpreadsheetApp.getActive().getSheetByName(OPTIONS_SHEET);
  if (!sh) { showSnack_('Лист Options не найден','Options',2000); return; }

  var last = sh.getLastRow();
  if (last < 2) { showSnack_('Нет данных для обновления','Options',2000); return; }

  var colId    = OPTIONS_HEADERS.indexOf('ID (FIGI/UID)') + 1;
  var colPrice = OPTIONS_HEADERS.indexOf('Текущая цена') + 1;

  var ids = sh.getRange(2, colId, last-1, 1).getValues().flat().map(String).filter(Boolean);
  if (!ids.length) { showSnack_('Идентификаторы не найдены','Options',2000); return; }

  // Берём цены сразу по FIGI и UID — что найдётся
  var idList = ids.map(function(x){ var p = parseId_(x); return p ? p.id : null; }).filter(Boolean);
  var priceMap = fetchLastPricesByIds_(idList); // map by figi И by uid

  var prices = ids.map(function(x){
    var p = parseId_(x);
    if (!p) return [''];
    var rec = priceMap[p.id];
    return [ rec != null ? Number(rec) : '' ];
  });
  sh.getRange(2, colPrice, prices.length, 1).setValues(prices);

  applyOptionsFormulas_(sh, 2, last-1);
  formatOptionsSheet_(sh);
  showSnack_('Цены обновлены и формулы пересчитаны','Options • Prices',2000);
}

// ========================= ФОРМУЛЫ / ФОРМАТ =========================

function applyOptionsFormulas_(sh, startRow, numRows){
  var idx = function(name){ return OPTIONS_HEADERS.indexOf(name)+1; };

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

  sh.getRange(startRow, cInv, numRows, 1).setFormulaR1C1(
    '=IF(OR(LEN(RC['+d(cQty,cInv)+'])=0' + SEP + 'LEN(RC['+d(cAvg,cInv)+'])=0)' + SEP +
    '""' + SEP + R2('RC['+d(cQty,cInv)+']*RC['+d(cAvg,cInv)+']') + ')'
  );
  sh.getRange(startRow, cMkt, numRows, 1).setFormulaR1C1(
    '=IF(OR(LEN(RC['+d(cQty,cMkt)+'])=0' + SEP + 'LEN(RC['+d(cPrice,cMkt)+'])=0)' + SEP +
    '""' + SEP + R2('RC['+d(cQty,cMkt)+']*RC['+d(cPrice,cMkt)+']') + ')'
  );
  sh.getRange(startRow, cPL, numRows, 1).setFormulaR1C1(
    '=IF(OR(LEN(RC['+d(cMkt,cPL)+'])=0' + SEP + 'LEN(RC['+d(cInv,cPL)+'])=0)' + SEP +
    '""' + SEP + R2('RC['+d(cMkt,cPL)+']-RC['+d(cInv,cPL)+']') + ')'
  );
  sh.getRange(startRow, cPLPct, numRows, 1).setFormulaR1C1(
    '=IF(OR(LEN(RC['+d(cPL,cPLPct)+'])=0' + SEP + 'LEN(RC['+d(cInv,cPLPct)+'])=0)' + SEP +
    '""' + SEP + R2('(RC['+d(cPL,cPLPct)+']/RC['+d(cInv,cPLPct)+'])*100') + ')'
  );
}

function formatOptionsSheet_(sh){
  var last = sh.getLastRow();
  if (last < 2) return;

  SpreadsheetApp.flush();
  sh.autoResizeColumns(1, OPTIONS_HEADERS.length);

  var COL_NAME  = OPTIONS_HEADERS.indexOf('Название') + 1;
  var COL_UL    = OPTIONS_HEADERS.indexOf('Базовый актив') + 1;
  if (COL_NAME > 0) sh.setColumnWidth(COL_NAME, 220);
  if (COL_UL   > 0) sh.setColumnWidth(COL_UL,   140);

  try { sh.setRowHeight(1, 32); } catch(e){}
  try { sh.setRowHeights(2, last-1, 30); } catch(e){}
}

// ========================= ВСПОМОГАТЕЛЬНЫЕ =========================

/** Разбор ID из Input: "FIGI" или "uid:xxxxxxxx-..." */
function parseId_(v){
  if (!v) return null;
  var s = String(v).trim();
  if (!s) return null;
  if (/^uid:/i.test(s)) return { id: s.slice(4), kind:'UID' };
  return { id: s, kind:'FIGI' };
}

/** Собрать карту позиций по ЛЮБОМУ ID (FIGI/UID). */
function fetchAllPositionsByAnyId_(){
  var out = {}; // id -> [{qty, avg, avg_fifo}]
  var accs = callUsersGetAccounts_() || [];
  accs.forEach(function(a){
    try{
      var raw = tinkoffFetch('tinkoff.public.invest.api.contract.v1.OperationsService/GetPositions',
                             { accountId: a.accountId }, {allow404:true}) || {};

      ['options','positions','securities','futures','derivatives'].forEach(function(key){
        var arr = raw[key];
        if (!Array.isArray(arr)) return;
        arr.forEach(function(p){
          var id = (p.figi || p.instrumentFigi || p.instrumentUid || p.uid || '').trim();
          if (!id) return;
          var qty = qToNumber(p.quantity) ?? (p.balance != null ? Number(p.balance) : null);
          var avg = moneyToNumber(p.averagePositionPrice || p.averagePositionPriceFifo || p.averagePositionPriceNoNkd);
          var avg_fifo = moneyToNumber(p.averagePositionPriceFifo || p.averagePositionPrice || p.averagePositionPriceNoNkd);
          if (!out[id]) out[id] = [];
          out[id].push({ qty: qty, avg: avg, avg_fifo: avg_fifo });
        });
      });
    }catch(_){}
  });
  return out;
}

/** Вытянуть lastPrice по списку FIGI/UID. Возвращает map по ИД. */
function fetchLastPricesByIds_(ids){
  var map = {}; if (!ids || !ids.length) return map;

  for (var i=0;i<ids.length;i+=300){
    var chunk = ids.slice(i, i+300);
    try{
      var d = tinkoffFetch('tinkoff.public.invest.api.contract.v1.MarketDataService/GetLastPrices',
                           { instrumentId: chunk }, {allow404:true}) || {};
      var arr = d.lastPrices || [];
      arr.forEach(function(x){
        var price = qToNumber(x.price || x.lastPrice);
        var f = x.figi || x.instrumentFigi || '';
        var u = x.instrumentUid || x.uid || '';
        if (f) map[f] = price;
        if (u) map[u] = price;
      });
    }catch(_){}
  }
  return map;
}

function mapOptionDirection_(x){
  if (!x) return '';
  var s = String(x).toLowerCase();
  if (s.indexOf('call') >= 0) return 'CALL';
  if (s.indexOf('put')  >= 0) return 'PUT';
  return String(x).toUpperCase();
}
function mapOptionStyle_(x){
  if (!x) return '';
  var s = String(x).toLowerCase();
  if (s.indexOf('american') >= 0) return 'AMERICAN';
  if (s.indexOf('european') >= 0) return 'EUROPEAN';
  return String(x).toUpperCase();
}
function normalizeDateLike_(v){
  try{
    if (!v) return '';
    if (v instanceof Date) return v.toISOString();
    if (typeof v === 'string') return v;
    if (typeof v === 'object'){
      if (v.seconds != null){
        var ms = Number(v.seconds) * 1000 + Math.floor(Number(v.nanos||0)/1e6);
        return new Date(ms).toISOString();
      }
      if (v.year && v.month && v.day){
        var dt = new Date(v.year, (v.month-1)||0, v.day||1);
        return dt.toISOString();
      }
    }
  }catch(e){}
  return String(v);
}
function numberFromMoneyOrQuotation_(x){
  if (!x) return null;
  try{
    if (typeof moneyToNumber === 'function') return moneyToNumber(x);
    if (typeof qToNumber === 'function')     return qToNumber(x);
  }catch(e){}
  var units = Number(x.units || 0), nano = Number(x.nano || 0)/1e9;
  var v = units + nano;
  return isNaN(v) ? null : v;
}
