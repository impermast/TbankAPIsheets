/**
 * portfolio_ui.gs
 * Вспомогательные UI: нижние «снэки», списки аккаунтов, загрузка FIGI, купон по FIGI и debug.
 */



/** Показывает текст в штатной плашке “Выполняется скрипт …”.
 *  title — что именно показать в плашке.
 *  autoCloseSec — через сколько секунд закрыть (необязательно; если не задано — висит, пока скрипт не завершится/не перезапишется новым вызовом).
 */
function setStatus_(title, autoCloseSec){
  var safe = (title && String(title).trim()) ? String(title).trim() : 'Выполняется скрипт';
  var html = HtmlService.createHtmlOutput(
    // мини-контейнер + опциональное автозакрытие
    '<div style="width:1px;height:1px;"></div>' +
    (autoCloseSec ? '<script>setTimeout(function(){google.script.host.close()},'+(autoCloseSec*1000)+');</script>' : '')
  ).setWidth(1).setHeight(1);
  // важное: НЕ пустой title, иначе ошибка
  SpreadsheetApp.getUi().showModalDialog(html, safe);
}

/** Алиас на setStatus_: вместо своих “снеков” просто пишем статус в плашку. */
function showSnack_(message, title, ms){
  // Склеиваем: "Заголовок • Сообщение" или просто Сообщение
  var txt = title ? (String(title).trim() + ' • ' + String(message||'')) : String(message||'');
  setStatus_(txt, ms ? Math.ceil(ms/1000) : null);
}


function showPanel_(title, html){
  var out = HtmlService.createHtmlOutput(
    '<div style="font:13px/1.4 -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif;padding:12px;">' +
      '<div style="font-weight:600;margin-bottom:8px;">'+ htmlEscape_(title) +'</div>' +
      html +
      '<div style="margin-top:10px;color:#6b7280;">Если панель мешает — закройте её крестиком.</div>' +
    '</div>'
  ).setTitle(title).setWidth(360);
  SpreadsheetApp.getUi().showSidebar(out);
}

// ===== Debug & портфельные утилиты =====
function debugPortfolioAccess(){
  try{
    var accs=callUsersGetAccounts_();
    if(!accs.length){ showPanel_('Проверка доступа', '<div>Нет доступа к аккаунтам. Проверьте права токена.</div>'); return; }
    var rows=[], totalPos=0, totalQty=0;
    for(var i=0;i<accs.length;i++){
      var a=accs[i];
      var r1=tinkoffFetchRaw_('tinkoff.public.invest.api.contract.v1.OperationsService/GetPortfolio',{accountId:a.accountId});
      var r2=tinkoffFetchRaw_('tinkoff.public.invest.api.contract.v1.OperationsService/GetPositions',{accountId:a.accountId});
      var cnt=0, qty=0;
      if(r1.code===200){ try{ var j1=JSON.parse(r1.text)||{}; var p1=(j1.positions||[]); cnt+=p1.length; for(var k=0;k<p1.length;k++){ var q=qToNumber(p1[k].quantity||p1[k].balance); if(q) qty+=Number(q); } }catch(e){} }
      if(r2.code===200){ try{ var j2=JSON.parse(r2.text)||{}; var p2=(j2.positions||j2.securities||[]); cnt+=p2.length; for(var m=0;m<p2.length;m++){ var q2=qToNumber(p2[m].quantity||p2[m].balance); if(q2) qty+=Number(q2); } }catch(e){} }
      totalPos+=cnt; totalQty+=qty;
      rows.push('<tr><td>'+htmlEscape_(a.accountId)+'</td><td>'+htmlEscape_(a.name||'')+'</td><td>'+r1.code+'/'+r2.code+'</td><td style="text-align:right">'+cnt+'</td><td style="text-align:right">'+qty+'</td></tr>');
    }
    var html = '<div style="margin-bottom:8px">Аккаунтов: <b>'+accs.length+
               '</b>, позиций суммарно: <b>'+totalPos+
               '</b>, qty: <b>'+totalQty+'</b></div>'+
               '<table style="border-collapse:collapse;width:100%">'+
               '<thead><tr><th>AccountId</th><th>Имя</th><th>HTTP (Portfolio/Positions)</th><th style="text-align:right">Позиций</th><th style="text-align:right">Qty</th></tr></thead>'+
               '<tbody>'+rows.join('')+'</tbody></table>';
    showPanel_('Проверка доступа к портфелю', html);
  }catch(e){
    showPanel_('Проверка доступа — ошибка', '<div>'+htmlEscape_(e && e.message)+'</div>');
  }
}
function portfolioShowAccounts(){
  try{
    var accs=callUsersGetAccounts_();
    if(!accs.length){ showPanel_('Аккаунты', '<div>Аккаунтов не найдено.</div>'); return; }
    var rows = accs.map(function(a,i){
      return '<tr><td>'+(i+1)+'</td><td>'+htmlEscape_(a.accountId)+'</td><td>'+htmlEscape_(a.name||'')+'</td></tr>';
    }).join('');
    var html = '<table style="border-collapse:collapse;width:100%"><thead><tr><th>#</th><th>AccountId</th><th>Имя</th></tr></thead><tbody>'+rows+'</tbody></table>';
    showPanel_('Список аккаунтов ('+accs.length+')', html);
  }catch(e){
    showPanel_('Аккаунты — ошибка', '<div>'+htmlEscape_(e && e.message)+'</div>');
  }
}
function portfolioShowBalances(){
  try{
    var accs=callUsersGetAccounts_();
    if(!accs.length){ showPanel_('Баланс', '<div>Аккаунтов не найдено.</div>'); return; }
    var total=0, rows=[];
    for(var i=0;i<accs.length;i++){
      var a=accs[i];
      var d=tinkoffFetch('tinkoff.public.invest.api.contract.v1.OperationsService/GetPortfolio',{accountId:a.accountId},{allow404:true});
      var val = d ? moneyToNumber(d.totalAmountPortfolio) : null;
      var v = (val!=null) ? Number(val) : null;
      if(v!=null) total+=v;
      rows.push('<tr><td>'+htmlEscape_(a.accountId)+'</td><td>'+htmlEscape_(a.name||'')+'</td><td style="text-align:right">'+(v!=null? v.toFixed(2) : 'n/a')+'</td></tr>');
    }
    var html = '<div style="margin-bottom:8px">Суммарная оценка портфеля: <b>'+ (isNaN(total)? 'n/a' : total.toFixed(2)) +'</b></div>'+
               '<table style="border-collapse:collapse;width:100%"><thead><tr><th>AccountId</th><th>Имя</th><th style="text-align:right">Оценка</th></tr></thead><tbody>'+rows+'</tbody></table>';
    showPanel_('Баланс по аккаунтам', html);
  }catch(e){
    showPanel_('Баланс — ошибка', '<div>'+htmlEscape_(e && e.message)+'</div>');
  }
}
function portfolioShowPositionsCount(){
  try{
    var accs=callUsersGetAccounts_();
    if(!accs.length){ showPanel_('Позиции', '<div>Аккаунтов не найдено.</div>'); return; }
    var total=0, rows=[];
    for(var i=0;i<accs.length;i++){
      var a=accs[i];
      var p1=callPortfolioGetPortfolio_(a.accountId);
      var p2=callPortfolioGetPositions_(a.accountId);
      var s={}, all=p1.concat(p2);
      for(var k=0;k<all.length;k++){ if(all[k].figi) s[all[k].figi]=1; }
      var cnt=Object.keys(s).length; total+=cnt;
      rows.push('<tr><td>'+htmlEscape_(a.accountId)+'</td><td>'+htmlEscape_(a.name||'')+'</td><td style="text-align:right">'+cnt+'</td></tr>');
    }
    var html = '<div style="margin-bottom:8px">Всего позиций по всем аккаунтам: <b>'+total+'</b></div>'+
               '<table style="border-collapse:collapse;width:100%"><thead><tr><th>AccountId</th><th>Имя</th><th style="text-align:right">Позиций</th></tr></thead><tbody>'+rows+'</tbody></table>';
    showPanel_('Позиции по аккаунтам', html);
  }catch(e){
    showPanel_('Позиции — ошибка', '<div>'+htmlEscape_(e && e.message)+'</div>');
  }
}

// Загрузка FIGI из портфеля в Input
function loadInputFigisFromPortfolio(type){
  var typeMap = { bond: 'Облигации', share: 'Акции', etf: 'Фонды' };
  if(!typeMap[type]){ showSnack_('Неизвестный тип: '+type, 'Загрузка FIGI', 2500); return; }

  var accs = callUsersGetAccounts_();
  if(!accs.length){ showSnack_('Нет доступа к аккаунтам', 'Загрузка FIGI', 3000); return; }

  var uniq = {};
  accs.forEach(function(a){
    var p1 = callPortfolioGetPortfolio_(a.accountId);
    var p2 = callPortfolioGetPositions_(a.accountId);
    p1.concat(p2).forEach(function(p){ if(p.figi) uniq[p.figi]=1; });
  });
  var allFigis = Object.keys(uniq);
  if(!allFigis.length){ showSnack_('В портфеле нет позиций', 'Загрузка FIGI', 2500); return; }

  var matched = [];
  for(var i=0;i<allFigis.length;i++){
    var f = allFigis[i], ok=false;
    try{
      if(type==='bond'){ ok = !!callInstrumentsBondByFigi_(f); }
      else if(type==='share'){ ok = !!callInstrumentsShareByFigi_(f); }
      else if(type==='etf'){ ok = !!callInstrumentsEtfByFigi_(f); }
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
  showSnack_('Загружено FIGI: '+matched.length+' ('+typeMap[type]+')', 'Загрузка FIGI', 3000);
}

// Купонное модальное окно
function showBondCouponInfoByFigi(figi){
  try{
    var s = getBondCouponSnapshot_(figi);
    function fmt(v, suffix){ if(v==null || v==='') return '—'; return (typeof v==='number'? String(round2_(v)) : String(v)) + (suffix||''); }
    var lines = [
      'Название: ' + (s.name || '—') + (s.ticker ? ' ('+s.ticker+')' : ''),
      'FIGI: ' + s.figi,
      'Номинал: ' + fmt(s.nominal),
      'Валюта: ' + (s.currency || '—'),
      '',
      'Тип купона: ' + (s.couponTypeDesc || '—'),
      'Частота купонов/год: ' + (s.couponsPerYear || '—'),
      'Размер купона: ' + fmt(s.couponValue),
      'Купон, %: ' + fmt(s.couponRate, '%'),
      'Купонный период (дней): ' + (s.couponPeriodDays || '—'),
      'Следующий купон (дата): ' + (s.nextCouponDate || '—') + (s.couponNumber ? (' (№'+s.couponNumber+')') : ''),
      'Дата фиксации реестра: ' + (s.fixDate || '—'),
      '',
      'Текущая цена: ' + fmt(s.lastPrice),
      'НКД (сегодня): ' + fmt(s.accrued),
      'Время цены: ' + (s.priceTime || '—'),
      'Прибл. купонная доходность: ' + fmt(s.approxYieldPct, '%'),
      '',
      'Уровень риска: ' + (s.riskLevelDesc || '—')
    ];
    showPanel_('Купон по FIGI', '<pre style="white-space:pre-wrap;margin:0">'+htmlEscape_(lines.join('\n'))+'</pre>');
  }catch(e){
    showSnack_((e && e.message) || String(e), 'Ошибка', 4000);
  }
}
