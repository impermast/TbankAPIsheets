/**
 * utils.gs
 * Вспомогательные функции: кэш, конвертеры, рендеринг HTML-таблиц,
 * Input-утилиты и общие helpers.
 */

// ============================== CACHE ===============================
function cacheGet_(key){ return CacheService.getScriptCache().get(key); }
function cachePut_(key, val, ttlSec){ CacheService.getScriptCache().put(key, val, ttlSec || 1800); }

// ======================= CONVERTERS / HELPERS ======================
function qToNumber(q) {
  if (!q) return null;
  var units = Number(q.units || 0);
  var nano  = Number(q.nano  || 0);
  return units + nano / 1e9;
}
function moneyToNumber(m) {
  if (!m) return null;
  if (m.units != null || m.nano != null) return qToNumber(m);
  if (m.value != null) return Number(m.value);
  return null;
}
function tsToIso(ts) {
  if (!ts) return '';
  if (typeof ts === 'string') return ts;
  if (typeof ts === 'object' && ts.seconds != null) {
    var ms = Number(ts.seconds) * 1000 + Math.round(Number(ts.nanos || 0) / 1e6);
    return new Date(ms).toISOString();
  }
  return '';
}
function dateOrEmpty_(iso) {
  if (!iso) return '';
  try { return Utilities.formatDate(new Date(iso), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
  catch (e) { return ''; }
}
function bondPricePctToCurrency_(price, nominal) {
  if (price == null || isNaN(price)) return null;
  var p = Number(price);
  var n = Number(nominal);
  if (!isNaN(n) && n > 0) {
    var threshold = Math.max(200, n * 0.25);
    if (p <= threshold) return Math.round((p * n / 100) * 100) / 100;
  }
  return Math.round(p * 100) / 100;
}
function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function getOrCreateInputSheet_() { return getOrCreateSheet_('Input'); }
function readInputFigis_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Input');
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 1).getValues().flat();
  var seen = {};
  return vals.map(function (v) { return String(v || '').trim(); })
             .filter(function (v) { if (!v || seen[v]) return false; seen[v] = 1; return true; });
}
function htmlEscape_(s){ s=(s==null)?'':String(s); return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function quotationToNumber(q){
  if (!q) return null;
  var units = Number(q.units || 0);
  var nano  = Number(q.nano || 0) / 1e9;
  var val = units + nano;
  return isNaN(val) ? null : val;
}
function round2_(x) { if (x == null || isNaN(x)) return null; return Math.round(Number(x) * 100) / 100; }
function mapCouponType_(v) {
  var key = (v == null) ? '' : String(v);
  var map = {
    '0':'Неопределенное','1':'Постоянный','2':'Плавающий','3':'Дисконт','4':'Ипотечный','5':'Фиксированный','6':'Переменный','7':'Прочее',
    'COUPON_TYPE_UNSPECIFIED':'Неопределенное','COUPON_TYPE_CONSTANT':'Постоянный','COUPON_TYPE_FLOATING':'Плавающий','COUPON_TYPE_DISCOUNT':'Дисконт',
    'COUPON_TYPE_MORTGAGE':'Ипотечный','COUPON_TYPE_FIX':'Фиксированный','COUPON_TYPE_VARIABLE':'Переменный','COUPON_TYPE_OTHER':'Прочее'
  };
  return map[key] || '';
}
function mapRiskLevel_(v) {
  var key = (v == null) ? '' : String(v);
  var map = {
    '0':'Высокий','1':'Средний','2':'Низкий',
    'RISK_LEVEL_HIGH':'Высокий','RISK_LEVEL_MODERATE':'Средний','RISK_LEVEL_LOW':'Низкий'
  };
  return map[key] || '';
}

// ========================= Input загрузчики ========================
/** Прочитать FIGI из листа Input по типу: A=bond, B=etf, C=share, D=option */
function readInputFigisByType_(type) {
  var sh = SpreadsheetApp.getActive().getSheetByName('Input');
  if (!sh) return [];
  var colIdx = { bond: 1, etf: 2, share: 3, option: 4 }[type];
  if (!colIdx) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, colIdx, last - 1, 1).getValues().flat();
  var seen = {};
  return vals.map(function(v){ return String(v||'').trim(); })
             .filter(function(v){ if(!v || seen[v]) return false; seen[v]=1; return true; });
}
/** Универсальная загрузка FIGI всех типов в Input (A–D) */
function loadInputFigisAllTypes_(){
  setStatus_('Сканирую портфели и собираю FIGI…');

  var accs = callUsersGetAccounts_();
  if(!accs.length){ showSnack_('Нет доступа к аккаунтам','Загрузка FIGI',3000); return; }

  var uniqFigis = {};
  var optionUids = {};

  accs.forEach(function(a){
    var p1 = callPortfolioGetPortfolio_(a.accountId) || [];
    var p2 = callPortfolioGetPositions_(a.accountId) || [];
    p1.concat(p2).forEach(function(p){ if(p && p.figi) uniqFigis[p.figi]=1; });

    try{
      var raw = tinkoffFetch('tinkoff.public.invest.api.contract.v1.OperationsService/GetPositions',
                             { accountId: a.accountId }, {allow404:true}) || {};
      ['options','positions','securities','futures','derivatives'].forEach(function(key){
        var arr = raw[key];
        if (!Array.isArray(arr)) return;
        arr.forEach(function(p){
          if (!p) return;
          var hasFigi = !!(p.figi || p.instrumentFigi);
          var uid = p.instrumentUid || p.uid;
          if (!hasFigi && uid) optionUids[uid] = 1;
        });
      });
    }catch(e){}
  });

  var allFigis = Object.keys(uniqFigis);

  setStatus_('Распознаю опционы по UID…');
  var extraOptionFigis = [];
  Object.keys(optionUids).forEach(function(uid){
    try{
      var opt = callInstrumentsOptionByUid_(uid);
      if (opt && opt.figi) extraOptionFigis.push(opt.figi);
    }catch(e){}
  });

  setStatus_('Классифицирую FIGI по типам…');
  var byType = { bond:[], etf:[], share:[], option:[] };

  allFigis.forEach(function(figi){
    try{ if (callInstrumentsBondByFigi_(figi))   { byType.bond.push(figi);  return; } }catch(e){}
    try{ if (callInstrumentsEtfByFigi_(figi))    { byType.etf.push(figi);   return; } }catch(e){}
    try{ if (callInstrumentsShareByFigi_(figi))  { byType.share.push(figi); return; } }catch(e){}
    try{ if (callInstrumentsOptionByFigi_(figi)) { byType.option.push(figi);return; } }catch(e){}
  });

  if (extraOptionFigis.length){
    var seen = {};
    byType.option.concat(extraOptionFigis).forEach(function(f){ if(f) seen[f]=1; });
    byType.option = Object.keys(seen);
  }

  var sh = getOrCreateSheet_('Input');
  sh.clear();
  sh.getRange(1,1,1,4).setValues([['Облигации','Фонды','Акции','Опционы']]);

  var maxLen = Math.max(byType.bond.length, byType.etf.length, byType.share.length, byType.option.length, 0);
  if (maxLen > 0) {
    var data = [];
    for (var i=0;i<maxLen;i++){
      data.push([
        byType.bond[i]   || '',
        byType.etf[i]    || '',
        byType.share[i]  || '',
        byType.option[i] || ''
      ]);
    }
    sh.getRange(2,1,maxLen,4).setValues(data);
  }
  sh.autoResizeColumns(1,4);
  setStatus_('Готово: FIGI загружены в Input (A–D)', 2);
}
/** Выгрузить FIGI заданного типа в отдельный лист (однотипный) */
function loadInputFigisFromPortfolio(type){
  var typeMap = { bond: 'Облигации', share: 'Акции', etf: 'Фонды', option: 'Опционы' };
  if(!typeMap[type]){ showSnack_('Неизвестный тип: '+type, 'Загрузка FIGИ', 2500); return; }

  var accs = callUsersGetAccounts_();
  if(!accs.length){ showSnack_('Нет доступа к аккаунтам', 'Загрузка FIGИ', 3000); return; }

  var uniq = {};
  accs.forEach(function(a){
    var p1 = callPortfolioGetPortfolio_(a.accountId);
    var p2 = callPortfolioGetPositions_(a.accountId);
    p1.concat(p2).forEach(function(p){ if(p.figi) uniq[p.figi]=1; });
  });

  var allFigis = Object.keys(uniq);
  if(!allFigis.length){ showSnack_('В портфеле нет позиций', 'Загрузка FIGИ', 2500); return; }

  var matched = [];
  for(var i=0;i<allFigis.length;i++){
    var f = allFigis[i], ok=false;
    try{
      if (type==='bond')   { ok = !!callInstrumentsBondByFigi_(f); }
      else if (type==='share')  { ok = !!callInstrumentsShareByFigi_(f); }
      else if (type==='etf')    { ok = !!callInstrumentsEtfByFigi_(f); }
      else if (type==='option') { ok = !!callInstrumentsOptionByFigi_(f); }
    }catch(e){ ok=false; }
    if(ok) matched.push(f);
  }

  var sh = getOrCreateInputSheet_();
  sh.clear();
  sh.getRange(1,1).setValue('FIGI');
  if(matched.length){
    var data = matched.map(function(f){ return [f]; });
    sh.getRange(2,1,data.length,1).setValues(data);
  }
  sh.autoResizeColumns(1,1);
  showSnack_('Загружено FIGI: '+matched.length+' ('+typeMap[type]+')', 'Загрузка FIGИ', 3000);
}

// ============================ Render helpers =======================
function _renderKVTable(rows){
  var tr = rows.map(function(pair){
    var k = _html(pair[0] != null ? String(pair[0]) : '');
    var v = _html(pair[1] != null ? String(pair[1]) : '');
    return '<tr><td style="padding:6px 8px;border:1px solid #ddd;background:#fafafa;width:40%">'+k+
           '</td><td style="padding:6px 8px;border:1px solid #ddd">'+v+'</td></tr>';
  }).join('');
  return '<table style="border-collapse:collapse;width:100%;font-size:13px">'+tr+'</table>';
}
function _renderFlatTable(kv){
  var head = '<thead><tr><th style="text-align:left;padding:6px 8px;border:1px solid #ddd;background:#f1f5f9;width:45%">Ключ</th><th style="text-align:left;padding:6px 8px;border:1px solid #ddd;background:#f1f5f9">Значение</th></tr></thead>';
  var body = kv.map(function(x){
    return '<tr><td style="padding:6px 8px;border:1px solid #eee;background:#fafafa;vertical-align:top">'+_html(x.k)+'</td>'+
           '<td style="padding:6px 8px;border:1px solid #eee;white-space:pre-wrap;font-family:monospace">'+_html(x.v)+'</td></tr>';
  }).join('');
  return '<div style="max-height:480px;overflow:auto;border:1px solid #ddd">'+
         '<table style="border-collapse:collapse;width:100%;font-size:12px">'+head+'<tbody>'+body+'</tbody></table></div>';
}
function _flatten(obj, root){
  var out = [];
  var seen = new WeakSet();
  function walk(node, path){
    if (node === null || node === undefined) { out.push({k:path, v:String(node)}); return; }
    var t = typeof node;
    if (t === 'string' || t === 'number' || t === 'boolean'){ out.push({k:path, v:String(node)}); return; }
    if (t !== 'object'){ out.push({k:path, v:String(node)}); return; }
    if (seen.has(node)){ out.push({k:path, v:'[Circular]'}); return; }
    seen.add(node);

    if (Array.isArray(node)){
      if (node.length === 0){ out.push({k:path, v:'[]'}); return; }
      for (var i=0;i<node.length;i++) walk(node[i], path + '['+i+']');
      return;
    }
    var isQuotation = (node.units != null && node.nano != null && Object.keys(node).length <= 3);
    var isMoneyVal  = (node.currencyCode || node.currency || node.units != null && node.nano != null);
    if (isQuotation){
      var val = _qToNumber(node);
      out.push({k:path, v: JSON.stringify(node) + (val != null ? '  // = ' + val : '')});
      return;
    }
    if (isMoneyVal && node.units != null && node.nano != null){
      var mv = _qToNumber(node);
      var cur = node.currencyCode || node.currency || '';
      out.push({k:path, v: JSON.stringify(node) + (mv != null ? '  // = ' + mv + (cur ? (' ' + cur) : '') : '')});
      return;
    }
    var keys = Object.keys(node);
    if (keys.length === 0){ out.push({k:path, v:'{}'}); return; }
    keys.sort().forEach(function(k){ walk(node[k], path ? (path + '.' + k) : k); });
  }
  walk(obj, root || '');
  return out;
}
function _qToNumber(q){
  try{
    if (typeof qToNumber === 'function') return qToNumber(q);
    var u = Number(q && q.units || 0);
    var n = Number(q && q.nano || 0)/1e9;
    var v = u + n;
    return isNaN(v) ? null : v;
  }catch(e){ return null; }
}
function _html(s){
  try{ if (typeof htmlEscape_ === 'function') return htmlEscape_(s); }catch(e){}
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

/** ===== Date helpers (глобальные) ===== */
function addDays_(d, n){
  // Возвращаем дату без времени (локальная TZ), сдвинутую на n дней
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}
function addMonths_(d, n){
  // Для горизонта купонов в дашборде нам нужен первый день через n месяцев
  return new Date(d.getFullYear(), d.getMonth() + n, 1);
}


/** instrumentMetaByFigi_: берём только имя/тикер/класс. Быстро и с кэшем. */
function instrumentMetaByFigi_(figi){
  var ck = 'meta:'+figi;
  var c  = cacheGet_(ck); if (c) return JSON.parse(c);
  var out = { class:'?', name:'', ticker:'' };
  try{ var o=callInstrumentsOptionByFigi_(figi); if(o){ out.class='option'; out.name=o.name||o.ticker||''; out.ticker=o.ticker||''; cachePut_(ck, JSON.stringify(out), 12*3600); return out; } }catch(_){}
  try{ var b=callInstrumentsBondByFigi_(figi);   if(b){ out.class='bond';   out.name=b.name||b.ticker||''; out.ticker=b.ticker||''; cachePut_(ck, JSON.stringify(out), 12*3600); return out; } }catch(_){}
  try{ var e=callInstrumentsEtfByFigi_(figi);    if(e){ out.class='etf';    out.name=e.name||e.ticker||''; out.ticker=e.ticker||''; cachePut_(ck, JSON.stringify(out), 12*3600); return out; } }catch(_){}
  try{ var s=callInstrumentsShareByFigi_(figi);  if(s){ out.class='share';  out.name=s.name||s.ticker||''; out.ticker=s.ticker||''; cachePut_(ck, JSON.stringify(out), 12*3600); return out; } }catch(_){}
  cachePut_(ck, JSON.stringify(out), 2*3600);
  return out;
}

