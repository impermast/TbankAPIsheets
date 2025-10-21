/**
 * main_menu.gs
 * Простое и понятное меню
 */

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Тинькофф • Портфель')

    // ГЛАВНОЕ: Информация по FIGI
    .addItem('Информация по FIGI…', 'menuShowInstrumentInfoByFigi')
    .addSeparator() 
    // Блок: Облигации
    .addSubMenu(
      ui.createMenu('Облигации')
        .addItem('Обновить только цены', 'menuUpdateBondPrices')   // читает FIGI из Input!A (Облигации)
        .addItem('Полное обновление', 'menuFullUpdateBonds')        // читает FIGI из Input!A (Облигации)
    )

    // Блок: Фонды
    .addSubMenu(
      ui.createMenu('Фонды')
        .addItem('Обновить только цены', 'menuUpdateFundPrices')   // читает FIGI из Input!B (Фонды)
        .addItem('Полное обновление', 'menuFullUpdateFunds')
    )
    // В onOpen():
    .addSubMenu(
      ui.createMenu('Опционы')
        .addItem('Обновить только цены', 'menuUpdateOptionPrices')
        .addItem('Полное обновление', 'menuFullUpdateOptions')
    )

    // Блок: Портфель / Аккаунт
    .addSeparator() 
    .addSubMenu(
      ui.createMenu('Портфель')
        .addItem('Проверка доступа', 'debugPortfolioAccess')
        .addItem('Список аккаунтов', 'portfolioShowAccounts')
        .addItem('Сводка активов', 'portfolioShowAllAssets')
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
function menuUpdateFundPrices(){ updateFundPricesOnly(); }
function menuFullUpdateFunds(){  updateFundsFull(); }
function menuUpdateBondPrices(){ updateBondPricesOnly(); }
function menuFullUpdateBonds(){  updateBondsFull(); }
function menuBuildDashboard(){   buildBondsDashboard(); }
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
