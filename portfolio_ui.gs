/** portfolio_ui.gs — отладочные панели и загрузка FIGI, купонная карточка */

function htmlEscape_(s){ s=(s==null)?'':String(s); return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
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

function debugPortfolioAccess(){
  try{
    var accs=callUsersGetAccounts_();
    if(!accs.length){ showPanel_('Проверка доступа', '<div>Нет доступа к аккаунтам. Проверьте права токена.</div>'); return; }
    var rows=[], totalPos=0, totalQty=0;
    for(var i=0;i<accs.length;i++){
      var a=accs[i];
      var r1=tinkoffFetch('tinkoff.public.invest.api.contract.v1.OperationsService/GetPortfolio',{accountId:a.accountId},{allow404:true});
      var r2=tinkoffFetch('tinkoff.public.invest.api.contract.v1.OperationsService/GetPositions',{accountId:a.accountId},{allow404:true});
      var cnt=0, qty=0;
      var p1 = (r1 && r1.positions) || [];
      var p2 = (r2 && (r2.positions||r2.securities)) || [];
      cnt += p1.length + p2.length;
      p1.concat(p2).forEach(function(p){ var q=qToNumber(p.quantity||p.balance); if(q) qty+=Number(q); });
      totalPos+=cnt; totalQty+=qty;
      rows.push('<tr><td>'+htmlEscape_(a.accountId)+'</td><td>'+htmlEscape_(a.name||'')+'</td><td style="text-align:right">'+cnt+'</td><td style="text-align:right">'+qty+'</td></tr>');
    }
    var html = '<div style="margin-bottom:8px">Аккаунтов: <b>'+accs.length+
               '</b>, позиций суммарно: <b>'+totalPos+
               '</b>, qty: <b>'+totalQty+'</b></div>'+
               '<table style="border-collapse:collapse;width:100%">'+
               '<thead><tr><th>AccountId</th><th>Имя</th><th style="text-align:right">Позиций</th><th style="text-align:right">Qty</th></tr></thead>'+
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
    var html = '<div style="margin-bottom:8px">Всего позиций: <b>'+total+'</b></div>'+
               '<table style="border-collapse:collapse;width:100%"><thead><tr><th>AccountId</th><th>Имя</th><th style="text-align:right">Позиций</th></tr></thead><tbody>'+rows+'</tbody></table>';
    showPanel_('Позиции по аккаунтам', html);
  }catch(e){
    showPanel_('Позиции — ошибка', '<div>'+htmlEscape_(e && e.message)+'</div>');
  }
}

/** Загрузка FIGI из портфеля в лист Input */
function loadInputFigisFromPortfolio(type){
  var typeMap = { bond: 'Облигации', share: 'Акции', etf: 'Фонды' };
  if(!typeMap[type]){ notify_('Загрузка FIGI', 'Неизвестный тип: '+type, 5); return; }

  var accs = callUsersGetAccounts_();
  if(!accs.length){ notify_('Загрузка FIGI', 'Нет доступа к аккаунтам', 6); return; }

  var uniq = {};
  accs.forEach(function(a){
    var p1 = callPortfolioGetPortfolio_(a.accountId);
    var p2 = callPortfolioGetPositions_(a.accountId);
    p1.concat(p2).forEach(function(p){ if(p.figi) uniq[p.figi]=1; });
  });
  var allFigis = Object.keys(uniq);
  if(!allFigis.length){ notify_('Загрузка FIGI', 'В портфеле нет позиций', 5); return; }

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
  notify_('Загрузка FIGI', 'Загружено FIGI: '+matched.length+' ('+typeMap[type]+')', 6);
}

/** Купонная карточка */
function getBondCouponSnapshot_(figi){
  var card = callInstrumentsBondByFigi_(figi);
  if(!card) throw new Error('FIGI не найден как облигация или нет доступа: '+figi);

  var nominal        = (card.nominal!=null) ? (qToNumber(card.nominal) ?? moneyToNumber(card.nominal)) : null;
  var couponRate     = (card.couponRate!=null) ? qToNumber(card.couponRate) : null;
  var couponValue    = (card.couponValue!=null) ? moneyToNumber(card.couponValue) :
                       (card.couponNominal!=null ? moneyToNumber(card.couponNominal) : null);
  var couponsPerYear = (card.couponQuantityPerYear!=null) ? Number(card.couponQuantityPerYear) : null;
  var couponTypeCard = card.couponType;
  var riskDesc       = mapRiskLevel_(card.riskLevel);

  var now=new Date();
  var fromIso=new Date(now.getTime()-30*24*3600*1000).toISOString();
  var toIso  =new Date(now.getTime()+3*365*24*3600*1000).toISOString();
  var events = callInstrumentsGetBondCouponsCached_(figi, fromIso, toIso) || [];
  var future=null,lastPast=null;
  events.forEach(function(c){
    var dtIso = tsToIso(c.couponDate || c.coupon_date || c.couponDateLt || c.date);
    if(!dtIso) return;
    var dt=new Date(dtIso);
    var val = moneyToNumber(c.payOneBond || c.pay_one_bond || c.couponValue || c.value);
    var typ = c.couponType || c.coupon_type;
    var per = (c.couponPeriod!=null?Number(c.couponPeriod):(c.coupon_period!=null?Number(c.coupon_period):null));
    var num = c.couponNumber || c.coupon_number || null;
    var fix = tsToIso(c.fixDate || c.fix_date);
    var rec = {dtIso:dtIso,value:(val!=null?Number(val):null),type:typ,period:per,number:num,fixIso:fix};
    if(dt.getTime()>=now.getTime()){
      if(!future || new Date(future.dtIso)>dt) future=rec;
    } else {
      if(!lastPast || new Date(lastPast.dtIso)<dt) lastPast=rec;
    }
  });
  var chosen = future || lastPast || null;

  var perYearFromEvent = (chosen && chosen.period && chosen.period>0) ? Math.max(1, Math.round(365/ chosen.period)) : null;
  if(couponsPerYear==null && perYearFromEvent!=null) couponsPerYear = perYearFromEvent;
  if(couponValue==null && chosen && chosen.value!=null) couponValue = chosen.value;
  if(couponRate==null && couponValue!=null && nominal && couponsPerYear){
    couponRate = (couponValue*couponsPerYear/nominal)*100;
  }
  var couponTypeDesc = mapCouponType_((chosen && chosen.type!=null)? chosen.type : couponTypeCard);

  var lastArr = callMarketLastPrices_([figi]);
  var lp = (lastArr && lastArr.length) ? lastArr[0] : null;
  var lastPrice = lp ? bondPricePctToCurrency_(lp.lastPrice, nominal) : null;
  var priceTime = lp ? lp.time : '';
  var accrued   = callMarketAccruedInterestsToday_(figi);
  var approxYieldPct = (couponValue!=null && couponsPerYear && lastPrice) ? (couponValue*couponsPerYear/lastPrice)*100 : null;

  return {
    figi: figi,
    name: card.name || card.placementName || '',
    ticker: card.ticker || '',
    currency: card.currency || card.currencyCode || '',
    nominal: nominal,
    couponRate: couponRate,
    couponValue: couponValue,
    couponsPerYear: couponsPerYear,
    couponTypeDesc: couponTypeDesc,
    riskLevelDesc: riskDesc,
    nextCouponDate: chosen ? dateOrEmpty_(chosen.dtIso) : '',
    couponNumber: chosen ? (chosen.number||'') : '',
    couponPeriodDays: chosen ? (chosen.period||'') : '',
    fixDate: chosen ? dateOrEmpty_(chosen.fixIso) : '',
    lastPrice: (lastPrice!=null? round2_(lastPrice): null),
    priceTime: priceTime || '',
    accrued: (accrued!=null? round2_(accrued): null),
    approxYieldPct: (approxYieldPct!=null? round2_(approxYieldPct): null)
  };
}

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
    notify_('Купон по FIGI', lines.join(' • '), 6);
  }catch(e){
    notify_('Ошибка', (e && e.message) || String(e), 6);
  }
}
