/**
 * input_loader.gs
 * Единая загрузка FIGI из портфелей по типам в лист Input:
 * A: Облигации, B: Фонды, C: Акции, D: Опционы
 */

// type: 'bond' | 'etf' | 'share' | 'option'
function readInputFigisByType_(type) {
  var sh = SpreadsheetApp.getActive().getSheetByName('Input');
  if (!sh) return [];
  // Заголовки по колонкам: A=Облигации, B=Фонды, C=Акции, D=Опционы
  var colIdx = { bond: 1, etf: 2, share: 3, option: 4 }[type];
  if (!colIdx) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, colIdx, last - 1, 1).getValues().flat();
  var seen = {};
  return vals.map(function(v){ return String(v||'').trim(); })
             .filter(function(v){ if(!v || seen[v]) return false; seen[v]=1; return true; });
}

function loadInputFigisAllTypes_(){
  setStatus_('Сканирую портфели и собираю FIGI…');

  var accs = callUsersGetAccounts_();
  if(!accs.length){ showSnack_('Нет доступа к аккаунтам','Загрузка FIGI',3000); return; }

  // Собираем уникальные FIGI по всем аккаунтам
  var uniq = {};
  accs.forEach(function(a){
    var p1 = callPortfolioGetPortfolio_(a.accountId);
    var p2 = callPortfolioGetPositions_(a.accountId);
    p1.concat(p2).forEach(function(p){ if(p && p.figi) uniq[p.figi]=1; });
  });
  var allFigis = Object.keys(uniq);
  if(!allFigis.length){ showSnack_('В портфеле нет позиций','Загрузка FIGI',3000); return; }

  setStatus_('Классифицирую FIGI по типам…');

  // Классификация по типам: bond / etf / share / option
  var byType = { bond:[], etf:[], share:[], option:[] };
  allFigis.forEach(function(figi){
    try{
      if (callInstrumentsBondByFigi_(figi))      { byType.bond.push(figi); return; }
    }catch(e){}
    try{
      if (callInstrumentsEtfByFigi_(figi))       { byType.etf.push(figi); return; }
    }catch(e){}
    try{
      if (callInstrumentsShareByFigi_(figi))     { byType.share.push(figi); return; }
    }catch(e){}
    try{
      if (callInstrumentsOptionByFigi_(figi))    { byType.option.push(figi); return; }
    }catch(e){}
    // если ничего не нашли — игнорим
  });

  var sh = getOrCreateSheet_('Input');
  sh.clear();

  // Заголовки по колонкам
  var headers = [['Облигации','Фонды','Акции','Опционы']];
  sh.getRange(1,1,1,4).setValues(headers);

  // максимальная длина
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
