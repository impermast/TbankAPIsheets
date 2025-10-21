function loadInputFigisAllTypes_(){
  setStatus_('Сканирую портфели и собираю FIGI/UID…');

  var accs = callUsersGetAccounts_();
  if (!accs.length){ showSnack_('Нет доступа к аккаунтам','Загрузка FIGI',3000); return; }

  // FIGI всех классов + UID опционов (если FIGI нет)
  var uniqFigis = {};
  var optionUids = {};

  accs.forEach(function(a){
    // как и раньше берём обе выборки
    var p1 = callPortfolioGetPortfolio_(a.accountId) || [];
    var p2 = callPortfolioGetPositions_(a.accountId) || [];

    p1.concat(p2).forEach(function(p){
      if (!p) return;
      var figi = p.figi || p.instrumentFigi;
      var uid  = p.instrumentUid || p.uid;

      if (figi) {
        uniqFigis[figi] = 1;
      } else if (uid) {
        // у опционов часто нет FIGI — запомним UID
        optionUids[uid] = 1;
      }
    });

    // подстраховка: «сыро» читаем GetPositions и выдёргиваем UID опционов
    try{
      var raw = tinkoffFetch('tinkoff.public.invest.api.contract.v1.OperationsService/GetPositions',
                             { accountId: a.accountId }, {allow404:true}) || {};
      (raw.options || []).forEach(function(p){
        if (!p) return;
        var hasFigi = !!(p.figi || p.instrumentFigi);
        var uid = p.instrumentUid || p.uid;
        if (!hasFigi && uid) optionUids[uid] = 1;
      });
    }catch(e){}
  });

  var allFigis = Object.keys(uniqFigis);

  setStatus_('Классифицирую FIGI по типам…');
  var byType = { bond:[], etf:[], share:[], option:[] };

  // FIGI классифицируем через *By(FIGI)
  allFigis.forEach(function(figi){
    try{ if (callInstrumentsBondByFigi_(figi))   { byType.bond.push(figi);   return; } }catch(_){}
    try{ if (callInstrumentsEtfByFigi_(figi))    { byType.etf.push(figi);    return; } }catch(_){}
    try{ if (callInstrumentsShareByFigi_(figi))  { byType.share.push(figi);  return; } }catch(_){}
    try{ if (callInstrumentsOptionByFigi_(figi)) { byType.option.push(figi); return; } }catch(_){}
  });

  // UID опционов добавляем как uid:<UID>
  Object.keys(optionUids).forEach(function(uid){
    byType.option.push('uid:' + uid);
  });

  // дедуп
  Object.keys(byType).forEach(function(k){
    var seen = {}; byType[k] = byType[k].filter(function(x){ if(seen[x]) return false; seen[x]=1; return true; });
  });

  // Выгрузка в Input: A=Облигации, B=Фонды, C=Акции, D=Опционы (FIGI или uid:UID)
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
  setStatus_('Готово: FIGI/UID загружены в Input (A–D)', 2);
}
