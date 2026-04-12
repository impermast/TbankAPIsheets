/**
 * ui.gs
 * UI-вспомогательные окна и панели.
 * Зависимости: utils (htmlEscape_, _renderKVTable, _renderFlatTable, _flatten, _html),
 *              tinkoff_api (все call*), прочие твои *update* функции.
 */

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Тинькофф • Портфель')
    .addItem('Информация по FIGI…', 'menuShowInstrumentInfoByFigi')
    .addSeparator()
    .addItem('Полный цикл обновления', 'menuRunFullRefreshCycle')
    .addItem('Обновить данные портфеля', 'menuRefreshPortfolioData')
    .addItem('Обновить форматирование', 'menuApplyPortfolioFormatting')
    .addItem('Обновить Dashboard', 'menuBuildDashboard')
    .addItem('Отправить в Telegram', 'menuSendTelegramReport')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Настройки')
        .addItem('Задать/сменить токен', 'uiSetToken')
        .addItem('Загрузить FIGI в Input (все типы)', 'menuLoadFigisAllTypes')
        .addItem('Проверка доступа', 'menuSettingsDebugAccess')
        .addItem('Список аккаунтов', 'menuSettingsShowAccounts')
        .addItem('Сводка активов', 'menuSettingsShowAllAssets')
    )
    .addToUi();
}

function _uiResolveFunctionName_(candidates) {
  var arr = candidates || [];
  for (var i = 0; i < arr.length; i++) {
    var name = arr[i];
    if (!name) continue;
    if (typeof globalThis[name] === 'function') return name;
  }
  return '';
}

function _uiInvokeByName_(fnName) {
  var fn = fnName ? globalThis[fnName] : null;
  if (typeof fn !== 'function') {
    throw new Error('Функция не найдена: ' + fnName);
  }
  return fn();
}

function _uiSyncInputIfAvailable_() {
  var fnName = _uiResolveFunctionName_(['loadInputFigisAllTypes_']);
  if (!fnName) return false;
  _uiInvokeByName_(fnName);
  return true;
}

function _uiRunFormatting_() {
  var fnName = _uiResolveFunctionName_(['runPortfolioFormating_', 'runPortfolioFormating']);
  if (!fnName) {
    throw new Error('Функция форматирования не найдена: runPortfolioFormating_ / runPortfolioFormating');
  }
  return _uiInvokeByName_(fnName);
}

function _uiBuildDashboard_() {
  var fnName = _uiResolveFunctionName_(['buildPortfolioDashboard', 'buildBondsDashboard']);
  if (!fnName) {
    throw new Error('Не найдена функция сборки Dashboard');
  }
  return _uiInvokeByName_(fnName);
}

function _uiSendCurrentDashboardToTelegram_() {
  if (typeof buildAndSendDashboardPackage === 'function') {
    return buildAndSendDashboardPackage();
  }

  SpreadsheetApp.flush();
  Utilities.sleep(800);

  var sent = 0;

  if (typeof sendDashboardStatsFromDashboard_ === 'function') {
    sendDashboardStatsFromDashboard_();
  }

  if (typeof sendDashboardChartsToTelegram_ === 'function') {
    sent = Number(sendDashboardChartsToTelegram_(['sectors','coupons','maturity','history','risk','ytmVsCoupon','scatter'])) || 0;
  }

  return sent;
}

function _uiShowAccountsFallback_() {
  try {
    if (typeof callUsersGetAccounts_ !== 'function') {
      showPanel_('Список аккаунтов', '<div>Функция callUsersGetAccounts_ недоступна.</div>');
      return;
    }

    var accs = callUsersGetAccounts_() || [];
    if (!accs.length) {
      showPanel_('Список аккаунтов', '<div>Аккаунты не найдены.</div>');
      return;
    }

    var rows = accs.map(function(a) {
      return [
        String(a.accountId || ''),
        String(a.name || ''),
        String(a.type || a.accountType || ''),
        String(a.status || '')
      ];
    });

    var html =
      '<div style="margin-bottom:8px">Аккаунтов: <b>' + rows.length + '</b></div>' +
      '<table style="border-collapse:collapse;width:100%">' +
      '<thead><tr>' +
      '<th style="text-align:left;padding:6px 8px;border:1px solid #ddd;background:#f1f5f9">AccountId</th>' +
      '<th style="text-align:left;padding:6px 8px;border:1px solid #ddd;background:#f1f5f9">Имя</th>' +
      '<th style="text-align:left;padding:6px 8px;border:1px solid #ddd;background:#f1f5f9">Тип</th>' +
      '<th style="text-align:left;padding:6px 8px;border:1px solid #ddd;background:#f1f5f9">Статус</th>' +
      '</tr></thead><tbody>' +
      rows.map(function(r) {
        return '<tr>' +
          '<td style="padding:6px 8px;border:1px solid #eee">' + htmlEscape_(r[0]) + '</td>' +
          '<td style="padding:6px 8px;border:1px solid #eee">' + htmlEscape_(r[1]) + '</td>' +
          '<td style="padding:6px 8px;border:1px solid #eee">' + htmlEscape_(r[2]) + '</td>' +
          '<td style="padding:6px 8px;border:1px solid #eee">' + htmlEscape_(r[3]) + '</td>' +
          '</tr>';
      }).join('') +
      '</tbody></table>';

    showPanel_('Список аккаунтов', html);
  } catch (e) {
    showPanel_('Список аккаунтов — ошибка', '<div>' + htmlEscape_(e && e.message) + '</div>');
  }
}

/** Задание/смена токена (хранится в пользовательских свойствах) */
function uiSetToken(){
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Tinkoff Invest API токен', 'Вставьте токен (начинается с t.)', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var t = (resp.getResponseText() || '').trim();
  if (!t) { showSnack_('Пустой ввод — отмена','Токен',2500); return; }
  setTinkoffToken(t);
  showSnack_('Сохранено. Ключ TINKOFF_TOKEN (пользовательские свойства)','Токен',3000);
}

/** Новые menu entrypoints */

function menuRunFullRefreshCycle() {
  try {
    if (typeof weeklyRefreshAndNotify === 'function') {
      weeklyRefreshAndNotify();
      return;
    }

    menuRefreshPortfolioData();
    menuApplyPortfolioFormatting();
    menuBuildDashboard();
    menuSendTelegramReport();
  } catch (e) {
    showPanel_('Полный цикл обновления — ошибка', '<div>' + htmlEscape_(e && e.message) + '</div>');
  }
}

function menuRefreshPortfolioData() {
  try {
    setStatus_('Портфель • обновление данных…', 5);

    var synced = false;
    try {
      synced = _uiSyncInputIfAvailable_();
    } catch (e1) {
      Logger.log('menuRefreshPortfolioData sync warning: ' + (e1 && e1.message || e1));
    }

    if (typeof updateBondsFull !== 'function') throw new Error('updateBondsFull не найдена');
    if (typeof updateFundsFull !== 'function') throw new Error('updateFundsFull не найдена');
    if (typeof updateSharesFull !== 'function') throw new Error('updateSharesFull не найдена');
    if (typeof updateOptionsFull !== 'function') throw new Error('updateOptionsFull не найдена');

    updateBondsFull();
    updateFundsFull();
    updateSharesFull();
    updateOptionsFull();

    SpreadsheetApp.flush();

    showSnack_(
      'Данные портфеля обновлены' + (synced ? ' + Input синхронизирован' : ''),
      'Портфель',
      2500
    );
  } catch (e) {
    showPanel_('Обновить данные портфеля — ошибка', '<div>' + htmlEscape_(e && e.message) + '</div>');
  }
}

function menuApplyPortfolioFormatting() {
  try {
    setStatus_('Портфель • форматирование…', 4);
    var result = _uiRunFormatting_();

    if (result && result.skipped) {
      showSnack_('Форматирование пропущено' + (result.reason ? ': ' + result.reason : ''), 'Форматирование', 2500);
    } else {
      showSnack_('Форматирование обновлено', 'Форматирование', 2000);
    }
  } catch (e) {
    showPanel_('Обновить форматирование — ошибка', '<div>' + htmlEscape_(e && e.message) + '</div>');
  }
}

function menuSendTelegramReport() {
  try {
    setStatus_('Telegram • отправка…', 4);
    var sent = _uiSendCurrentDashboardToTelegram_();
    showSnack_('Отправка в Telegram выполнена' + (sent ? ' (графиков: ' + sent + ')' : ''), 'Telegram', 2500);
  } catch (e) {
    showPanel_('Отправить в Telegram — ошибка', '<div>' + htmlEscape_(e && e.message) + '</div>');
  }
}

function menuSettingsDebugAccess() {
  if (typeof debugPortfolioAccess === 'function') {
    debugPortfolioAccess();
    return;
  }
  showSnack_('Функция проверки доступа не найдена', 'Настройки', 2500);
}

function menuSettingsShowAccounts() {
  if (typeof portfolioShowAccounts === 'function') {
    portfolioShowAccounts();
    return;
  }
  _uiShowAccountsFallback_();
}

function menuSettingsShowAllAssets() {
  if (typeof portfolioShowAllAssets === 'function') {
    portfolioShowAllAssets();
    return;
  }
  showSnack_('Функция сводки активов не найдена', 'Настройки', 2500);
}

/** Старые menu helper functions оставлены для совместимости */
function menuUpdateFundPrices(){ updateFundPricesOnly(); }
function menuFullUpdateFunds(){  updateFundsFull(); }
function menuUpdateBondPrices(){ updateBondPricesOnly(); }
function menuFullUpdateBonds(){  updateBondsFull(); }
function menuUpdateSharePrices(){ updateSharePricesOnly(); }
function menuFullUpdateShares(){ updateSharesFull(); }
function menuBuildDashboard(){ _uiBuildDashboard_(); }
function menuFullUpdateOptions(){ updateOptionsFull(); }
function menuUpdateOptionPrices(){ updateOptionPricesOnly(); }

/** Единая загрузка FIGИ всех типов */
function menuLoadFigisAllTypes(){ loadInputFigisAllTypes_(); }

/** Новое: Информация по FIGI (универсально для bond/etf/share/option) */
function menuShowInstrumentInfoByFigi(){
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Информация по FIGI', 'Введите FIGI (например, TCS00A123456):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var figi = (resp.getResponseText() || '').trim();
  if (!figi) { showSnack_('Пусто — отмена', 'Информация по FIGI', 2500); return; }
  showInstrumentInfoByFigi(figi);
}

/** Показ плашки статуса (встроенное модальное окно Apps Script) */
function setStatus_(title, autoCloseSec){
  try {
    var safe = (title && String(title).trim()) ? String(title).trim() : 'Выполняется скрипт';
    var html = HtmlService.createHtmlOutput(
      '<div style="width:1px;height:1px;"></div>' +
      (autoCloseSec ? '<script>setTimeout(function(){google.script.host.close()},'+(autoCloseSec*1000)+');</script>' : '')
    ).setWidth(1).setHeight(1);
    SpreadsheetApp.getUi().showModalDialog(html, safe);
  } catch (e){
    // В триггере UI нет — логируем и идём дальше
    Logger.log('setStatus_ skipped (no UI): ' + (e && e.message));
  }
}

/** Короткое уведомление через ту же плашку */
function showSnack_(message, title, ms){
  var txt = title ? (String(title).trim() + ' • ' + String(message||'')) : String(message||'');
  setStatus_(txt, ms ? Math.ceil(ms/1000) : null);
}

/** Широкая модельная панель (не блокирует работу листа) */
function showPanel_(title, html){
  var out = HtmlService.createHtmlOutput(
    '<div style="font:13px/1.4 -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif;padding:12px;">' +
      '<div style="font-weight:600;margin-bottom:8px;">'+ htmlEscape_(title) +'</div>' +
      html +
      '<div style="margin-top:10px;color:#6b7280;">Если панель мешает — закройте её крестиком.</div>' +
    '</div>'
  )
  .setWidth(1280)
  .setHeight(820);

  SpreadsheetApp.getUi().showModelessDialog(out, title);
}

// ================= Debug & портфельные утилиты =================

function debugPortfolioAccess(){
  try{
    var accs=callUsersGetAccounts_();
    if(!accs.length){
      showPanel_('Проверка доступа', '<div>Нет доступа к аккаунтам. Проверьте права токена.</div>');
      return;
    }
    var rows=[], totalPos=0, totalQty=0;
    for(var i=0;i<accs.length;i++){
      var a=accs[i];
      var r1=tinkoffFetchRaw_('tinkoff.public.invest.api.contract.v1.OperationsService/GetPortfolio',{accountId:a.accountId});
      var r2=tinkoffFetchRaw_('tinkoff.public.invest.api.contract.v1.OperationsService/GetPositions',{accountId:a.accountId});
      var cnt=0, qty=0;
      if(r1.code===200){
        try{
          var j1=JSON.parse(r1.text)||{};
          var p1=(j1.positions||[]);
          cnt+=p1.length;
          for(var k=0;k<p1.length;k++){
            var q=qToNumber(p1[k].quantity||p1[k].balance);
            if(q) qty+=Number(q);
          }
        }catch(e){}
      }
      if(r2.code===200){
        try{
          var j2=JSON.parse(r2.text)||{};
          var p2=(j2.positions||j2.securities||[]);
          cnt+=p2.length;
          for(var m=0;m<p2.length;m++){
            var q2=qToNumber(p2[m].quantity||p2[m].balance);
            if(q2) qty+=Number(q2);
          }
        }catch(e){}
      }
      totalPos+=cnt; totalQty+=qty;
      rows.push('<tr><td>'+htmlEscape_(a.accountId)+'</td><td>'+htmlEscape_(a.name||'')+
                '</td><td>'+r1.code+'/'+r2.code+'</td><td style="text-align:right">'+cnt+
                '</td><td style="text-align:right">'+qty+'</td></tr>');
    }
    var html = '<div style="margin-bottom:8px">Аккаунтов: <b>'+accs.length+
               '</b>, позиций суммарно: <b>'+totalPos+
               '</b>, qty: <b>'+totalQty+'</b></div>'+
               '<table style="border-collapse:collapse;width:100%">'+
               '<thead><tr><th>AccountId</th><th>Имя</th><th>HTTP (Portfolio/Positions)</th>'+
               '<th style="text-align:right">Позиций</th><th style="text-align:right">Qty</th></tr></thead>'+
               '<tbody>'+rows.join('')+'</tbody></table>';
    showPanel_('Проверка доступа к портфелю', html);
  }catch(e){
    showPanel_('Проверка доступа — ошибка', '<div>'+htmlEscape_(e && e.message)+'</div>');
  }
}

/** Кэшированный маппинг UID -> FIGI (для опционов) */
function figiFromUid_(uid){
  if(!uid) return null;
  var ck = 'figiByUid:'+uid;
  var c = cacheGet_(ck); if(c) return c === 'null' ? null : c;
  var inst = null;
  try {
    inst = callInstrumentsOptionByUid_(uid);
  } catch(_) {}
  var figi = inst && inst.figi ? String(inst.figi) : null;
  cachePut_(ck, figi==null ? 'null' : figi, 12*3600);
  return figi;
}

/** Универсальные метаданные по id: сначала FIGI, потом UID (опционы) */
function instrumentMetaByAnyId_(id){
  if(!id) return { class:'?', name:'', ticker:'' };
  var isUid = String(id).startsWith('uid:');
  var figi = isUid ? figiFromUid_(String(id).slice(4)) : String(id);

  var ck = 'metaAny:'+ (figi||id);
  var c  = cacheGet_(ck); if (c) return JSON.parse(c);

  var out = { class:'?', name:'', ticker:'' };
  try{ var o= figi && callInstrumentsOptionByFigi_(figi); if(o){ out.class='option'; out.name=o.name||o.ticker||''; out.ticker=o.ticker||''; cachePut_(ck, JSON.stringify(out), 12*3600); return out; } }catch(_){}
  try{ var b= figi && callInstrumentsBondByFigi_(figi);   if(b){ out.class='bond';   out.name=b.name||b.ticker||''; out.ticker=b.ticker||''; cachePut_(ck, JSON.stringify(out), 12*3600); return out; } }catch(_){}
  try{ var e= figi && callInstrumentsEtfByFigi_(figi);    if(e){ out.class='etf';    out.name=e.name||e.ticker||''; out.ticker=e.ticker||''; cachePut_(ck, JSON.stringify(out), 12*3600); return out; } }catch(_){}
  try{ var s= figi && callInstrumentsShareByFigi_(figi);  if(s){ out.class='share';  out.name=s.name||s.ticker||''; out.ticker=s.ticker||''; cachePut_(ck, JSON.stringify(out), 12*3600); return out; } }catch(_){}
  cachePut_(ck, JSON.stringify(out), 2*3600);
  return out;
}

function portfolioShowAllAssets(){
  try{
    var accs = callUsersGetAccounts_();
    if(!accs.length){ showPanel_('Все активы', '<div>Аккаунтов не найдено.</div>'); return; }

    var raw = [];
    accs.forEach(function(a){
      [callPortfolioGetPortfolio_(a.accountId)||[], callPortfolioGetPositions_(a.accountId)||[]].forEach(function(lst){
        lst.forEach(function(p){
          if(!p) return;
          var figi = p.figi || p.instrumentFigi || '';
          var uid  = p.instrumentUid || p.uid || '';
          var qty  = qToNumber(p.quantity) ?? (p.balance!=null ? Number(p.balance) : null);
          if (!figi && !uid) return;
          raw.push({ accountId:a.accountId, accountName:a.name||'', figi:figi||null, uid:uid||null, qty:qty });
        });
      });
    });
    if(!raw.length){ showPanel_('Все активы', '<div>Позиции не найдены.</div>'); return; }

    var uidToFigi = {};
    var uidList = [];
    raw.forEach(function(r){
      if(!r.figi && r.uid && !uidToFigi[r.uid]) uidList.push(r.uid);
    });
    uidList.forEach(function(uid){
      var ck = 'uid2figi:'+uid;
      var c  = cacheGet_(ck);
      if (c){ uidToFigi[uid] = c; return; }
      try{
        var opt = callInstrumentsOptionByUid_(uid);
        if (opt && opt.figi){ uidToFigi[uid] = opt.figi; cachePut_(ck, opt.figi, 3600); }
      }catch(_){}
    });
    raw.forEach(function(r){
      if(!r.figi && r.uid && uidToFigi[r.uid]) r.figi = uidToFigi[r.uid];
    });

    var byKey = {};
    raw.forEach(function(r){
      var id = r.figi ? r.figi : ('uid:'+r.uid);
      var k  = r.accountId+'|'+id;
      if(!byKey[k]) byKey[k] = { acc:(r.accountName||r.accountId), id:id, figi:r.figi||'', uid:r.uid||'', qty:0 };
      if(r.qty!=null) byKey[k].qty += Number(r.qty);
    });
    var rowsRaw = Object.keys(byKey).map(function(k){ return byKey[k]; });

    var figis = rowsRaw.map(function(r){ return r.figi; }).filter(function(x){ return !!x; });
    var seenF = {}; figis = figis.filter(function(f){ if(seenF[f]) return false; seenF[f]=1; return true; });
    var priceMap = {};
    for (var i=0;i<figis.length;i+=300){
      var chunk = figis.slice(i,i+300);
      try{
        (callMarketLastPrices_(chunk)||[]).forEach(function(x){
          priceMap[x.figi] = (x.lastPrice!=null ? Number(x.lastPrice) : null);
        });
      }catch(_){}
    }

    var metaCache = {};
    function getMeta(id){
      if (metaCache[id]) return metaCache[id];
      var out = { class:'?', name:'', ticker:'', currency:'' };
      var ck  = 'meta:'+id;
      var c   = cacheGet_(ck); if(c){ out = JSON.parse(c); metaCache[id]=out; return out; }

      var figi = id.indexOf('uid:')===0 ? '' : id;
      var uid  = id.indexOf('uid:')===0 ? id.slice(4) : '';

      if (uid){
        try{
          var o = callInstrumentsOptionByUid_(uid);
          if (o){ out.class='option'; out.name=o.name||o.ticker||''; out.ticker=o.ticker||''; out.currency=o.currency||o.buyCurrency||o.sellCurrency||''; cachePut_(ck, JSON.stringify(out), 12*3600); metaCache[id]=out; return out; }
        }catch(_){}
      }

      if (figi){
        try{ var o=callInstrumentsOptionByFigi_(figi); if(o){ out.class='option'; out.name=o.name||o.ticker||''; out.ticker=o.ticker||''; out.currency=o.currency||o.buyCurrency||o.sellCurrency||''; cachePut_(ck, JSON.stringify(out), 12*3600); metaCache[id]=out; return out; } }catch(_){}
        try{ var b=callInstrumentsBondByFigi_(figi);   if(b){ out.class='bond';   out.name=b.name||b.ticker||''; out.ticker=b.ticker||''; out.currency=b.currency||b.buyCurrency||b.sellCurrency||''; cachePut_(ck, JSON.stringify(out), 12*3600); metaCache[id]=out; return out; } }catch(_){}
        try{ var e=callInstrumentsEtfByFigi_(figi);    if(e){ out.class='etf';    out.name=e.name||e.ticker||''; out.ticker=e.ticker||''; out.currency=e.currency||e.buyCurrency||e.sellCurrency||''; cachePut_(ck, JSON.stringify(out), 12*3600); metaCache[id]=out; return out; } }catch(_){}
        try{ var s=callInstrumentsShareByFigi_(figi);  if(s){ out.class='share';  out.name=s.name||s.ticker||''; out.ticker=s.ticker||''; out.currency=s.currency||s.buyCurrency||s.sellCurrency||''; cachePut_(ck, JSON.stringify(out), 12*3600); metaCache[id]=out; return out; } }catch(_){}
      }

      cachePut_(ck, JSON.stringify(out), 2*3600);
      metaCache[id]=out;
      return out;
    }

    var rows = rowsRaw.map(function(r){
      var meta  = getMeta(r.figi ? r.figi : ('uid:'+r.uid));
      var price = r.figi ? priceMap[r.figi] : null;
      var qty   = (r.qty!=null ? Number(r.qty) : null);
      var mv    = (price!=null && qty!=null) ? price*qty : null;

      function f2(x){ return (x==null||isNaN(x)) ? '' : Number(x).toFixed(2); }
      var cur = meta.currency ? (' '+meta.currency) : '';

      return {
        acc:   r.acc,
        class: meta.class || '?',
        name:  meta.name || meta.ticker || (r.figi || ('uid:'+r.uid)),
        id:    r.figi || ('uid:'+r.uid),
        qty:   (qty==null? '' : f2(qty)),
        price: (price==null? '' : f2(price)+cur),
        value: (mv==null? '' : f2(mv)+cur)
      };
    });

    rows.sort(function(a,b){
      return (a.acc||'').localeCompare(b.acc||'') ||
             (a.class||'').localeCompare(b.class||'') ||
             (a.name||'').localeCompare(b.name||'');
    });

    var head = ['Счёт','Класс','Название','FIGI/UID','Кол-во','Цена','Стоимость'];
    var th = head.map(function(h){ return '<th style="text-align:left;padding:6px 8px;border:1px solid #ddd;background:#f1f5f9">'+htmlEscape_(h)+'</th>'; }).join('');
    var tr = rows.map(function(r){
      function td(v, right){ return '<td style="padding:6px 8px;border:1px solid #eee'+(right?';text-align:right':'')+'">'+htmlEscape_(v)+'</td>'; }
      return '<tr>'+td(r.acc)+td(r.class)+td(r.name)+td(r.id)+td(r.qty,true)+td(r.price,true)+td(r.value,true)+'</tr>';
    }).join('');
    var html = '<div style="margin-bottom:8px">Всего строк: <b>'+rows.length+'</b></div>'+
               '<div style="max-height:640px;overflow:auto;border:1px solid #ddd">'+
               '<table style="border-collapse:collapse;width:100%;font-size:12px"><thead><tr>'+th+'</tr></thead><tbody>'+tr+'</tbody></table></div>';

    showPanel_('Все активы (включая опционы)', html);

  }catch(e){
    showPanel_('Все активы — ошибка', '<div>'+htmlEscape_(e && e.message)+'</div>');
  }
}

// =============== Окно «Информация по FIGI» ===============

function showInstrumentInfoByFigi(figi){
  try {
    setStatus_('Чтение инструмента по FIGI…', 5);

    var inst = null, kind = '';
    var probes = [
      {kind:'bond',   fn: callInstrumentsBondByFigi_},
      {kind:'etf',    fn: callInstrumentsEtfByFigi_},
      {kind:'share',  fn: callInstrumentsShareByFigi_},
      {kind:'?',      fn: callInstrumentsShareByFigi_},
      {kind:'option', fn: callInstrumentsOptionByFigi_}
    ];
    for (var i=0;i<probes.length;i++){
      try {
        inst = probes[i].fn(figi);
        if (inst){ kind = probes[i].kind; break; }
      } catch(e){}
    }
    if (!inst){
      showPanel_('Информация по FIGI', '<div>Инструмент не найден по FIGI: <b>'+_html(figi)+'</b></div>');
      return;
    }

    var lastArr = [];
    try { lastArr = callMarketLastPrices_([figi]) || []; } catch(e){}
    var last = lastArr[0] || null;
    var marketObj = {
      lastPrice: (last && last.lastPrice != null) ? Number(last.lastPrice) : null,
      lastTime:  (last && last.time) ? String(last.time) : null
    };
    try {
      var ts = callMarketGetTradingStatus_(figi);
      if (ts) marketObj.tradingStatus = ts.tradingStatus || ts.status || null;
    } catch(e){}

    var asset = null;
    var assetUid = inst.assetUid || inst.asset_uid || null;
    if (assetUid){
      try { asset = callInstrumentsGetAssetByUid_(assetUid) || null; } catch(e){}
    }

    var bondNextCoupon = null;
    if (kind === 'bond' && typeof fetchBondsNextCoupons_ === 'function'){
      try {
        var map = fetchBondsNextCoupons_([figi]) || {};
        bondNextCoupon = map[figi] || null;
      } catch(e){}
    }

    var sections = [];

    var summaryRows = [
      ['Тип инструмента', kind],
      ['FIGI', figi],
      ['Название', inst.name || inst.ticker || ''],
      ['Тикер', inst.ticker || ''],
      ['Валюта', inst.currency || inst.buyCurrency || inst.sellCurrency || ''],
      ['Биржа', inst.exchange || inst.realExchange || ''],
      ['Заблокирован (TCS)', (inst.blockedTcaFlag === true ? 'Да' : (inst.blockedTcaFlag === false ? 'Нет' : ''))],
      ['Текущая цена', marketObj.lastPrice != null ? String(marketObj.lastPrice) : ''],
      ['Время цены', marketObj.lastTime || ''],
      ['Trading status', marketObj.tradingStatus || '']
    ];
    if (kind === 'bond'){
      summaryRows.push(['След. купон — дата', (bondNextCoupon && bondNextCoupon.date ? String(bondNextCoupon.date) : '')]);
      summaryRows.push(['След. купон — сумма', (bondNextCoupon && bondNextCoupon.value != null ? String(bondNextCoupon.value) : '')]);
    }
    sections.push('<h3 style="margin:16px 0 8px">Сводка</h3>' + _renderKVTable(summaryRows));

    sections.push('<h3 style="margin:16px 0 8px">Instrument ('+kind+')</h3>' + _renderFlatTable(_flatten(inst, 'instrument')));
    if (asset) sections.push('<h3 style="margin:16px 0 8px">Asset</h3>' + _renderFlatTable(_flatten(asset, 'asset')));
    sections.push('<h3 style="margin:16px 0 8px">Market</h3>' + _renderFlatTable(_flatten(marketObj, 'market')));

    var html = sections.join('<div style="height:12px"></div>');
    showPanel_('Информация по FIGI', html);

  } catch (e){
    showPanel_('Информация по FIGI — ошибка', '<div>'+_html(e && e.message)+'</div>');
  }
}
