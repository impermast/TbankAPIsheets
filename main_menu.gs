/** main_menu.gs — меню, уведомления, диалог FIGI */

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Тинькофф • Портфель')
    .addSubMenu(
      ui.createMenu('Облигации')
        .addItem('Обновить только цены', 'menuUpdateBondPrices')
        .addItem('Обновить все данные', 'menuFullUpdateBonds')
        .addSeparator()
        .addItem('Показать информацию об облигации по FIGI…', 'menuShowBondCouponByFigi')
    )
    .addSubMenu(
      ui.createMenu('Портфель / Аккаунт')
        .addItem('Проверка доступа (коды/счётчики)', 'debugPortfolioAccess')
        .addItem('Список аккаунтов', 'portfolioShowAccounts')
        .addItem('Сводка балансов по аккаунтам', 'portfolioShowBalances')
        .addItem('Счётчик позиций по аккаунтам', 'portfolioShowPositionsCount')
    )
    .addSubMenu(
      ui.createMenu('Загрузка FIGI из API')
        .addItem('Облигации', 'menuLoadFigisBonds')
        .addItem('Акции',     'menuLoadFigisShares')
        .addItem('Фонды',     'menuLoadFigisEtfs')
    )
    .addSubMenu(
      ui.createMenu('Сводка (Dashboard)')
        .addItem('Создать/обновить сводный лист', 'menuBuildDashboard')
    )
    .addSeparator()
    .addItem('Задать/сменить токен', 'uiSetToken')
    .addToUi();
}

/** Уведомления */
function notify_(title, msg, sec){
  try{
    SpreadsheetApp.getActive().toast(String(msg||''), String(title||''), Math.max(1, Number(sec||3)));
  }catch(e){}
}

/** «Снекбар» снизу, автозакрытие */
function showSnackbar_(text, millis){
  var html = HtmlService.createHtmlOutput(
    '<div id="snackbar" style="position:fixed;left:50%;bottom:24px;transform:translateX(-50%);' +
    'background:#111827;color:#fff;padding:10px 16px;border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.2);' +
    'font:13px -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif;z-index:99999;">' +
    (text ? String(text).replace(/[<>&]/g, s=>({"<":"&lt;",">":"&gt;","&":"&amp;"}[s])) : '') +
    '</div>' +
    '<script>setTimeout(function(){google.script.host.close();},'+ (millis||1800) +');</script>'
  ).setWidth(10).setHeight(10); // компактно
  SpreadsheetApp.getUi().showModelessDialog(html, '');
}

/** Диалог ввода FIGI в HTML */
function menuShowBondCouponByFigi(){
  var html = HtmlService.createHtmlOutput(
    '<div style="font:13px -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif;padding:14px 16px;min-width:320px">' +
    '<div style="font-weight:600;margin-bottom:6px">Купон по FIGI</div>' +
    '<input id="figi" placeholder="например, TCS00A123456" style="width:100%;padding:8px;border:1px solid #d1d5db;border-radius:8px"/>' +
    '<div style="margin-top:10px;display:flex;gap:8px;justify-content:flex-end">' +
      '<button onclick="google.script.host.close()" style="padding:6px 10px;border-radius:8px;border:1px solid #e5e7eb;background:#fff">Отмена</button>' +
      '<button onclick="runFigi()" style="padding:6px 10px;border-radius:8px;border:0;background:#4f46e5;color:#fff">Показать</button>' +
    '</div>' +
    '<script>function runFigi(){var f=document.getElementById("figi").value.trim(); if(!f){alert("Введите FIGI");return;} google.script.run.withSuccessHandler(()=>google.script.host.close()).showBondCouponInfoByFigi(f);}</script>' +
    '</div>'
  ).setTitle('Купон по FIGI');
  SpreadsheetApp.getUi().showModalDialog(html, 'Купон по FIGI');
}

/** Токен теперь в UserProperties */
function uiSetToken(){
  var ui=SpreadsheetApp.getUi();
  var resp=ui.prompt('Tinkoff Invest API токен','Вставьте токен (начинается с t.)',ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()===ui.Button.OK){
    var t=(resp.getResponseText()||'').trim();
    if(!t){ notify_('Токен','Пусто — отмена',3); return; }
    setTinkoffToken(t);
    showSnackbar_('Токен сохранён (UserProperties)', 1800);
  }
}

/** Проксируем пункты меню */
function menuUpdateBondPrices(){ updateBondPricesOnly(); }
function menuFullUpdateBonds(){  updateBondsFull(); }
function menuLoadFigisBonds(){  loadInputFigisFromPortfolio('bond'); }
function menuLoadFigisShares(){ loadInputFigisFromPortfolio('share'); }
function menuLoadFigisEtfs(){   loadInputFigisFromPortfolio('etf'); }
function menuBuildDashboard(){   buildBondsDashboard(); }
