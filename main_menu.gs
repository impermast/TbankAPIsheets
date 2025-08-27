/**
 * main_menu.gs
 * Простое и понятное меню
 */

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Тинькофф • Портфель')

    // Блок: Облигации
    .addSubMenu(
      ui.createMenu('Облигации')
        .addItem('Обновить только цены', 'menuUpdateBondPrices')   // читает FIGI из Input!A (Облигации)
        .addItem('Полное обновление', 'menuFullUpdateBonds')        // читает FIGI из Input!A (Облигации)
        .addSeparator()
        .addItem('Показать купон по FIGI…', 'menuShowBondCouponByFigi')
    )

    // Блок: Портфель / Аккаунт
    .addSubMenu(
      ui.createMenu('Портфель')
        .addItem('Проверка доступа', 'debugPortfolioAccess')
        .addItem('Список аккаунтов', 'portfolioShowAccounts')
        .addItem('Сводка балансов', 'portfolioShowBalances')
        .addItem('Счётчик позиций', 'portfolioShowPositionsCount')
    )

    // Блок: Данные
    .addSubMenu(
      ui.createMenu('Данные')
        .addItem('Загрузить FIGI в Input (все типы)', 'menuLoadFigisAllTypes') // A: Облигации, B: Фонды, C: Акции, D: Опционы
    )

    // Блок: Сводка
    .addSubMenu(
      ui.createMenu('Сводка')
        .addItem('Обновить Dashboard', 'menuBuildDashboard')
    )

    // Блок: Настройки
    .addSubMenu(
      ui.createMenu('Настройки')
        .addItem('Задать/сменить токен', 'uiSetToken')
    )

    .addToUi();
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

/** Меню → действия */
function menuUpdateBondPrices(){ updateBondPricesOnly(); }
function menuFullUpdateBonds(){  updateBondsFull(); }
function menuBuildDashboard(){   buildBondsDashboard(); }

/** Диалог «Купон по FIGI» */
function menuShowBondCouponByFigi(){
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Купон по FIGI', 'Введите FIGI облигации (например, TCS00A123456):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var figi = (resp.getResponseText() || '').trim();
  if (!figi) { showSnack_('Пусто — отмена', 'Купон по FIGI', 2500); return; }
  showBondCouponInfoByFigi(figi);
}

/** Единая загрузка FIGI всех типов в лист Input (A: облигации, B: фонды, C: акции, D: опционы) */
function menuLoadFigisAllTypes(){ loadInputFigisAllTypes_(); }
