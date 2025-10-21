/**
 * scheduler.gs
 * Еженедельные/ежемесячные задачи и установка/снятие триггеров.
 * Зависимости:
 *  - updateBondsFull(), updateFundsFull(), updateOptionsFull(), buildBondsDashboard()
 *  - sendDashboardChartsToTelegram_(aliases) из dashboard_exports.gs
 *  - tgSendPhoto_(blob, caption), tgSendMessage_(text)
 */

/** ===== Еженедельное автообновление (понедельник) ===== */
function weeklyRefreshAndNotify(){
  var t0 = Date.now();
  try {

    // 1) приветствие
    sendDashboardGreets_();

    if (typeof updateBondsFull     === 'function') updateBondsFull();
    Utilities.sleep(800);
    if (typeof updateFundsFull     === 'function') updateFundsFull();
    Utilities.sleep(800);
    if (typeof updateOptionsFull   === 'function') updateOptionsFull();
    Utilities.sleep(800);
    if (typeof buildBondsDashboard === 'function') buildBondsDashboard();

    SpreadsheetApp.flush();
    Utilities.sleep(500);

    

    // 2) диаграммы (сектора + купоны)
    sendDashboardChartsToTelegram_(['sectors','coupons']);

    // 3) короткая статистика из экспорт-секций
    sendDashboardStatsFromDashboard_();

    Logger.log('weeklyRefreshAndNotify OK, elapsed='+(Date.now()-t0)+'ms');
  } catch (e){
    Logger.log('weeklyRefreshAndNotify ERROR: ' + (e && e.stack || e));
    try { tgSendMessage_('Автообновление: ошибка: ' + (e && e.message)); } catch(_) {}
  }
}


/** Поставить еженедельный триггер на понедельник в заданный час (по таймзоне проекта). */
function installWeeklyAutoRefresh(hour){
  hour = Number(hour); if (isNaN(hour)) hour = 13;
  removeWeeklyAutoRefreshTriggers();
  ScriptApp.newTrigger('weeklyRefreshAndNotify')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(hour)
    .create();
  Logger.log('Weekly trigger set: SATURDAY at ' + hour + ':00');
}

/** Удалить все триггеры этой задачи. */
function removeWeeklyAutoRefreshTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(tr){
    if (tr.getHandlerFunction && tr.getHandlerFunction() === 'weeklyRefreshAndNotify'){
      ScriptApp.deleteTrigger(tr);
    }
  });
  Logger.log('Removed weeklyRefreshAndNotify triggers');
}

/** ===== (опционально) Ежемесячное автообновление ===== */
function monthlyAutoRefreshJob(){
  var t0 = Date.now();
  try {
    if (typeof updateBondsFull      === 'function') updateBondsFull();
    if (typeof updateFundsFull      === 'function') updateFundsFull();
    if (typeof updateOptionsFull    === 'function') updateOptionsFull();
    if (typeof buildBondsDashboard  === 'function') buildBondsDashboard();
    // При желании можно и сюда добавить отправку:
    // sendDashboardChartsToTelegram_(['sectors','coupons']);
    Logger.log('monthlyAutoRefreshJob: OK in ' + (Date.now()-t0) + ' ms');
  } catch (e){
    Logger.log('monthlyAutoRefreshJob: ERROR ' + (e && e.message));
  }
}

function installMonthlyAutoRefresh(day, hour){
  day  = Number(day)  || 1;
  hour = Number(hour) || 8;
  removeMonthlyAutoRefreshTriggers();
  ScriptApp.newTrigger('monthlyAutoRefreshJob')
    .timeBased()
    .onMonthDay(day)
    .atHour(hour)
    .create();
  Logger.log('Monthly trigger set: day=' + day + ', hour=' + hour);
}

function removeMonthlyAutoRefreshTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(tr){
    if (tr.getHandlerFunction && tr.getHandlerFunction() === 'monthlyAutoRefreshJob'){
      ScriptApp.deleteTrigger(tr);
    }
  });
  Logger.log('Removed monthlyAutoRefreshJob triggers');
}

function listProjectTriggers(){
  var trs = ScriptApp.getProjectTriggers();
  if (!trs.length){ Logger.log('Нет триггеров'); return; }
  trs.forEach(function(tr){
    Logger.log('Trigger: ' + (tr.getUniqueId ? tr.getUniqueId() : '?')
      + ', handler=' + tr.getHandlerFunction()
      + ', type=' + tr.getEventType());
  });
}
