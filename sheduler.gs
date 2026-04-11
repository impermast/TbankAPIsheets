/**
 * scheduler.gs
 * Еженедельные/ежемесячные задачи и установка/снятие триггеров.
 * Зависимости:
 * updateBondsFull(), updateFundsFull(), updateSharesFull(), updateOptionsFull(), buildBondsDashboard()/buildPortfolioDashboard()
 * sendDashboardChartsToTelegram_(aliases), sendDashboardStatsFromDashboard_(), tgSendMessage_(text)
 */

function _schedulerCreateSummary_(jobName) {
  return {
    jobName: jobName || 'job',
    startedAtMs: Date.now(),
    startedAtIso: new Date().toISOString(),
    currentStep: 'init',
    success: false,
    error: '',
    accountsCount: 0,
    inputSynced: false,
    inputSyncSkipped: false,
    dashboardBuilt: false,
    chartsSent: 0,
    statsSent: false,
    rowCounts: {
      Bonds: 0,
      Funds: 0,
      Shares: 0,
      Options: 0
    },
    warnings: [],
    steps: [],
    elapsedMs: 0
  };
}

function _schedulerWarn_(summary, message) {
  var msg = String(message || '').trim();
  if (!msg) return;
  summary.warnings.push(msg);
  Logger.log('[Scheduler][WARN] ' + msg);
}

function _schedulerSafeTelegram_(text) {
  if (typeof tgSendMessage_ !== 'function') return false;
  try {
    tgSendMessage_(String(text || ''));
    return true;
  } catch (e) {
    Logger.log('[Scheduler][TG ERROR] ' + (e && e.stack || e));
    return false;
  }
}

function _schedulerFormatDuration_(ms) {
  ms = Math.max(0, Number(ms) || 0);
  var totalSec = Math.round(ms / 1000);
  var min = Math.floor(totalSec / 60);
  var sec = totalSec % 60;
  return min + ' мин ' + sec + ' сек';
}

function _schedulerBoolText_(value, skipped) {
  if (skipped) return 'пропущено';
  return value ? 'да' : 'нет';
}

function _schedulerCountSheetRows_(sheetName) {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sh) return 0;
    return Math.max(0, sh.getLastRow() - 1);
  } catch (e) {
    Logger.log('[Scheduler][COUNT ERROR] ' + sheetName + ': ' + (e && e.message || e));
    return 0;
  }
}

function _schedulerAuditSheets_(summary) {
  summary.rowCounts.Bonds = _schedulerCountSheetRows_('Bonds');
  summary.rowCounts.Funds = _schedulerCountSheetRows_('Funds');
  summary.rowCounts.Shares = _schedulerCountSheetRows_('Shares');
  summary.rowCounts.Options = _schedulerCountSheetRows_('Options');

  var dash = null;
  try {
    dash = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  } catch (e) {}

  summary.dashboardBuilt = !!(dash && dash.getLastRow() > 0 && dash.getLastColumn() > 0);

  if (!summary.dashboardBuilt) {
    _schedulerWarn_(summary, 'Dashboard отсутствует или пуст после сборки');
  }

  return summary.rowCounts;
}

function _schedulerBuildSummaryText_(summary) {
  var lines = [];
  lines.push('Service summary: ' + summary.jobName);
  lines.push('Статус: ' + (summary.success ? 'OK' : 'ERROR'));
  if (summary.error) lines.push('Ошибка: ' + summary.error);
  lines.push('Аккаунтов: ' + (summary.accountsCount || 0));
  lines.push('Input sync: ' + _schedulerBoolText_(summary.inputSynced, summary.inputSyncSkipped));
  lines.push('Dashboard: ' + _schedulerBoolText_(summary.dashboardBuilt, false));
  lines.push(
    'Строки: Bonds ' + summary.rowCounts.Bonds +
    ' | Funds ' + summary.rowCounts.Funds +
    ' | Shares ' + summary.rowCounts.Shares +
    ' | Options ' + summary.rowCounts.Options
  );
  lines.push('Графиков отправлено: ' + (summary.chartsSent || 0));
  lines.push('Warnings: ' + summary.warnings.length);
  if (summary.warnings.length) {
    summary.warnings.slice(0, 5).forEach(function(w) {
      lines.push('- ' + w);
    });
  }
  lines.push('Длительность: ' + _schedulerFormatDuration_(summary.elapsedMs));
  return lines.join('\n');
}

function _schedulerRunStep_(summary, stepName, fn, isCritical) {
  var t0 = Date.now();
  summary.currentStep = stepName;
  Logger.log('[Scheduler] STEP START: ' + stepName);

  try {
    var result = fn();
    summary.steps.push({
      name: stepName,
      status: 'ok',
      elapsedMs: Date.now() - t0
    });
    Logger.log('[Scheduler] STEP OK: ' + stepName + ' in ' + (Date.now() - t0) + ' ms');
    return result;
  } catch (e) {
    var msg = stepName + ': ' + (e && e.message ? e.message : e);
    summary.steps.push({
      name: stepName,
      status: isCritical ? 'error' : 'warning',
      elapsedMs: Date.now() - t0,
      message: msg
    });

    Logger.log('[Scheduler] STEP ' + (isCritical ? 'ERROR' : 'WARN') + ': ' + msg);
    if (e && e.stack) Logger.log(e.stack);

    if (isCritical) throw new Error(msg);

    _schedulerWarn_(summary, msg);
    return null;
  }
}

function _schedulerPreflight_(summary) {
  var missing = [];

  if (typeof callUsersGetAccounts_ !== 'function') missing.push('callUsersGetAccounts_');
  if (typeof updateBondsFull !== 'function') missing.push('updateBondsFull');
  if (typeof updateFundsFull !== 'function') missing.push('updateFundsFull');
  if (typeof updateSharesFull !== 'function') missing.push('updateSharesFull');
  if (typeof updateOptionsFull !== 'function') missing.push('updateOptionsFull');
  if (!(typeof buildPortfolioDashboard === 'function' || typeof buildBondsDashboard === 'function')) {
    missing.push('buildPortfolioDashboard/buildBondsDashboard');
  }

  if (missing.length) {
    throw new Error('Отсутствуют критические функции: ' + missing.join(', '));
  }

  var accounts;
  try {
    accounts = callUsersGetAccounts_() || [];
  } catch (e) {
    throw new Error('Не удалось получить список аккаунтов: ' + (e && e.message ? e.message : e));
  }

  if (!accounts.length) {
    throw new Error('Нет доступных аккаунтов для обновления');
  }

  summary.accountsCount = accounts.length;

  if (typeof loadInputFigisAllTypes_ !== 'function') {
    _schedulerWarn_(summary, 'loadInputFigisAllTypes_ не найдена — sync Input будет пропущен');
  }
  if (typeof tgSendMessage_ !== 'function') {
    _schedulerWarn_(summary, 'tgSendMessage_ не найдена — сервисные Telegram-сообщения будут пропущены');
  }
  if (typeof sendDashboardChartsToTelegram_ !== 'function') {
    _schedulerWarn_(summary, 'sendDashboardChartsToTelegram_ не найдена — отправка графиков будет пропущена');
  }
  if (typeof sendDashboardStatsFromDashboard_ !== 'function') {
    _schedulerWarn_(summary, 'sendDashboardStatsFromDashboard_ не найдена — отправка статистики будет пропущена');
  }
}

function _schedulerRunPipeline_(opts) {
  opts = opts || {};

  var summary = _schedulerCreateSummary_(opts.jobName || 'pipeline');
  var chartAliases = opts.chartAliases || ['sectors','coupons','maturity','history','risk','ytmVsCoupon','scatter'];

  try {
    _schedulerRunStep_(summary, 'preflight', function() {
      _schedulerPreflight_(summary);
    }, true);

    if (opts.sendGreeting !== false) {
      _schedulerRunStep_(summary, 'greeting', function() {
        _schedulerSafeTelegram_(
          'Запускаю ' + summary.jobName +
          ': sync Input → update sheets → rebuild Dashboard → audit → Telegram'
        );
      }, false);
    }

    _schedulerRunStep_(summary, 'sync_input', function() {
      if (typeof loadInputFigisAllTypes_ === 'function') {
        loadInputFigisAllTypes_();
        summary.inputSynced = true;
        Utilities.sleep(800);
      } else {
        summary.inputSyncSkipped = true;
      }
    }, false);

    _schedulerRunStep_(summary, 'update_bonds', function() {
      updateBondsFull();
      Utilities.sleep(800);
    }, true);

    _schedulerRunStep_(summary, 'update_funds', function() {
      updateFundsFull();
      Utilities.sleep(800);
    }, true);

    _schedulerRunStep_(summary, 'update_shares', function() {
      updateSharesFull();
      Utilities.sleep(800);
    }, true);

    _schedulerRunStep_(summary, 'update_options', function() {
      updateOptionsFull();
      Utilities.sleep(800);
    }, true);

    _schedulerRunStep_(summary, 'build_dashboard', function() {
      if (typeof buildPortfolioDashboard === 'function') {
        buildPortfolioDashboard();
      } else {
        buildBondsDashboard();
      }
      SpreadsheetApp.flush();
      Utilities.sleep(500);
    }, true);

    _schedulerRunStep_(summary, 'post_run_audit', function() {
      _schedulerAuditSheets_(summary);
      if (!summary.dashboardBuilt) {
        throw new Error('Dashboard не построен или пуст после обновления');
      }
    }, true);

    if (opts.sendCharts) {
      _schedulerRunStep_(summary, 'send_charts', function() {
        if (typeof sendDashboardChartsToTelegram_ === 'function') {
          summary.chartsSent = Number(sendDashboardChartsToTelegram_(chartAliases)) || 0;
        }
      }, false);
    }

    if (opts.sendStats) {
      _schedulerRunStep_(summary, 'send_stats', function() {
        if (typeof sendDashboardStatsFromDashboard_ === 'function') {
          sendDashboardStatsFromDashboard_();
          summary.statsSent = true;
        }
      }, false);
    }

    summary.success = true;
    summary.elapsedMs = Date.now() - summary.startedAtMs;

    if (opts.sendServiceSummary) {
      _schedulerRunStep_(summary, 'send_service_summary', function() {
        _schedulerSafeTelegram_(_schedulerBuildSummaryText_(summary));
      }, false);
    }

    Logger.log('[Scheduler] ' + summary.jobName + ' OK, elapsed=' + summary.elapsedMs + 'ms');
    return summary;

  } catch (e) {
    summary.success = false;
    summary.error = e && e.message ? e.message : String(e || 'unknown error');
    summary.elapsedMs = Date.now() - summary.startedAtMs;

    try {
      _schedulerAuditSheets_(summary);
    } catch (_) {}

    Logger.log('[Scheduler] ' + summary.jobName + ' ERROR at step ' + summary.currentStep + ': ' + summary.error);
    if (e && e.stack) Logger.log(e.stack);

    if (opts.notifyErrors !== false) {
      _schedulerSafeTelegram_(
        summary.jobName + ': ошибка на шаге "' + summary.currentStep + '"\n' + summary.error
      );
      if (opts.sendServiceSummary !== false) {
        _schedulerSafeTelegram_(_schedulerBuildSummaryText_(summary));
      }
    }

    return summary;
  }
}

/** ===== Еженедельное автообновление (понедельник) ===== */
function weeklyRefreshAndNotify() {
  _schedulerRunPipeline_({
    jobName: 'weeklyRefreshAndNotify',
    sendGreeting: true,
    sendCharts: true,
    sendStats: true,
    sendServiceSummary: true,
    notifyErrors: true,
    chartAliases: ['sectors','coupons','maturity','history','risk','ytmVsCoupon','scatter']
  });
}

/** Поставить еженедельный триггер на понедельник в заданный час (по таймзоне проекта). */
function installWeeklyAutoRefresh(hour) {
  hour = Number(hour);
  if (isNaN(hour)) hour = 13;

  removeWeeklyAutoRefreshTriggers();

  ScriptApp.newTrigger('weeklyRefreshAndNotify')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(hour)
    .create();

  Logger.log('Weekly trigger set: MONDAY at ' + hour + ':00');
}

/** Удалить все триггеры этой задачи. */
function removeWeeklyAutoRefreshTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(tr) {
    if (tr.getHandlerFunction && tr.getHandlerFunction() === 'weeklyRefreshAndNotify') {
      ScriptApp.deleteTrigger(tr);
    }
  });
  Logger.log('Removed weeklyRefreshAndNotify triggers');
}

/** ===== (опционально) Ежемесячное автообновление ===== */
function monthlyAutoRefreshJob() {
  _schedulerRunPipeline_({
    jobName: 'monthlyAutoRefreshJob',
    sendGreeting: false,
    sendCharts: false,
    sendStats: false,
    sendServiceSummary: false,
    notifyErrors: true
  });
}

function installMonthlyAutoRefresh(day, hour) {
  day = Number(day) || 1;
  hour = Number(hour) || 8;

  removeMonthlyAutoRefreshTriggers();

  ScriptApp.newTrigger('monthlyAutoRefreshJob')
    .timeBased()
    .onMonthDay(day)
    .atHour(hour)
    .create();

  Logger.log('Monthly trigger set: day=' + day + ', hour=' + hour);
}

function removeMonthlyAutoRefreshTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(tr) {
    if (tr.getHandlerFunction && tr.getHandlerFunction() === 'monthlyAutoRefreshJob') {
      ScriptApp.deleteTrigger(tr);
    }
  });
  Logger.log('Removed monthlyAutoRefreshJob triggers');
}

function listProjectTriggers() {
  var trs = ScriptApp.getProjectTriggers();
  if (!trs.length) {
    Logger.log('Нет триггеров');
    return;
  }

  trs.forEach(function(tr) {
    Logger.log(
      'Trigger: ' + (tr.getUniqueId ? tr.getUniqueId() : '?') +
      ', handler=' + tr.getHandlerFunction() +
      ', type=' + tr.getEventType()
    );
  });
}
