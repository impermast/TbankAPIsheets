/**
 * scheduler.gs
 * Еженедельные/ежемесячные задачи и установка/снятие триггеров.
 *
 * End-to-end pipeline:
 * 1) preflight-check
 * 2) greeting
 * 3) sync Input
 * 4) update Bonds / Funds / Shares / Options
 * 5) post-processing formatting
 * 6) build Dashboard
 * 7) flush / stabilization
 * 8) Telegram charts
 * 9) Telegram stats
 * 10) short service summary
 *
 * Публичный контракт сохраняется:
 * - weeklyRefreshAndNotify()
 * - installWeeklyAutoRefresh(hour)
 * - removeWeeklyAutoRefreshTriggers()
 * - monthlyAutoRefreshJob()
 * - installMonthlyAutoRefresh(day, hour)
 * - removeMonthlyAutoRefreshTriggers()
 * - listProjectTriggers()
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

    dependencies: {
      inputSyncFn: '',
      formattingFn: '',
      dashboardFn: '',
      tgMessage: false,
      tgCharts: false,
      tgStats: false
    },

    inputSync: {
      attempted: false,
      synced: false,
      skipped: false,
      reason: ''
    },

    updates: {
      Bonds: false,
      Funds: false,
      Shares: false,
      Options: false
    },

    formatting: {
      attempted: false,
      applied: false,
      skipped: false,
      reason: '',
      result: null,
      functionName: ''
    },

    dashboard: {
      attempted: false,
      built: false,
      functionName: ''
    },

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

  summary.dashboard.built = !!(dash && dash.getLastRow() > 0 && dash.getLastColumn() > 0);

  if (!summary.dashboard.built) {
    _schedulerWarn_(summary, 'Dashboard отсутствует или пуст после сборки');
  }

  return summary.rowCounts;
}

function _schedulerResolveFunctionName_(candidates) {
  var arr = candidates || [];
  for (var i = 0; i < arr.length; i++) {
    var name = arr[i];
    if (!name) continue;
    if (typeof globalThis[name] === 'function') return name;
  }
  return '';
}

function _schedulerInvokeByName_(fnName) {
  var fn = fnName ? globalThis[fnName] : null;
  if (typeof fn !== 'function') {
    throw new Error('Функция не найдена: ' + fnName);
  }
  return fn();
}

function _schedulerBuildSummaryText_(summary) {
  var lines = [];

  lines.push('Service summary: ' + summary.jobName);
  lines.push('Статус: ' + (summary.success ? 'OK' : 'ERROR'));

  if (summary.error) {
    lines.push('Ошибка: ' + summary.error);
  }

  lines.push('Аккаунтов: ' + (summary.accountsCount || 0));

  if (summary.inputSync.synced) {
    lines.push('Input sync: synced');
  } else if (summary.inputSync.skipped) {
    lines.push('Input sync: skipped' + (summary.inputSync.reason ? ' (' + summary.inputSync.reason + ')' : ''));
  } else {
    lines.push('Input sync: no');
  }

  lines.push(
    'Updates: ' +
    'Bonds ' + (summary.updates.Bonds ? 'OK' : '—') + ' | ' +
    'Funds ' + (summary.updates.Funds ? 'OK' : '—') + ' | ' +
    'Shares ' + (summary.updates.Shares ? 'OK' : '—') + ' | ' +
    'Options ' + (summary.updates.Options ? 'OK' : '—')
  );

  if (summary.formatting.applied) {
    lines.push('Formatting: applied');
  } else if (summary.formatting.skipped) {
    lines.push('Formatting: skipped' + (summary.formatting.reason ? ' (' + summary.formatting.reason + ')' : ''));
  } else {
    lines.push('Formatting: no');
  }

  lines.push(
    'Dashboard: ' +
    (summary.dashboard.built ? 'built' : 'not built') +
    (summary.dashboard.functionName ? ' [' + summary.dashboard.functionName + ']' : '')
  );

  lines.push(
    'Строки: Bonds ' + summary.rowCounts.Bonds +
    ' | Funds ' + summary.rowCounts.Funds +
    ' | Shares ' + summary.rowCounts.Shares +
    ' | Options ' + summary.rowCounts.Options
  );

  lines.push('Telegram stats: ' + (summary.statsSent ? 'sent' : 'skipped'));
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

    if (isCritical) {
      throw new Error(msg);
    }

    _schedulerWarn_(summary, msg);
    return null;
  }
}

function _schedulerPreflight_(summary) {
  var missingCritical = [];

  if (typeof callUsersGetAccounts_ !== 'function') missingCritical.push('callUsersGetAccounts_');
  if (typeof updateBondsFull !== 'function') missingCritical.push('updateBondsFull');
  if (typeof updateFundsFull !== 'function') missingCritical.push('updateFundsFull');
  if (typeof updateSharesFull !== 'function') missingCritical.push('updateSharesFull');
  if (typeof updateOptionsFull !== 'function') missingCritical.push('updateOptionsFull');

  summary.dependencies.inputSyncFn = _schedulerResolveFunctionName_(['loadInputFigisAllTypes_']);
  summary.dependencies.formattingFn = _schedulerResolveFunctionName_(['runPortfolioFormating_', 'runPortfolioFormating']);
  summary.dependencies.dashboardFn = _schedulerResolveFunctionName_(['buildPortfolioDashboard', 'buildBondsDashboard']);
  summary.dependencies.tgMessage = (typeof tgSendMessage_ === 'function');
  summary.dependencies.tgCharts = (typeof sendDashboardChartsToTelegram_ === 'function');
  summary.dependencies.tgStats = (typeof sendDashboardStatsFromDashboard_ === 'function');

  if (!summary.dependencies.dashboardFn) {
    missingCritical.push('buildPortfolioDashboard/buildBondsDashboard');
  }

  if (missingCritical.length) {
    throw new Error('Отсутствуют критические функции: ' + missingCritical.join(', '));
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

  if (!summary.dependencies.inputSyncFn) {
    _schedulerWarn_(summary, 'loadInputFigisAllTypes_ не найдена — sync Input будет пропущен');
  }

  if (!summary.dependencies.formattingFn) {
    _schedulerWarn_(summary, 'runPortfolioFormating_/runPortfolioFormating не найдена — post-processing будет пропущен');
  }

  if (!summary.dependencies.tgMessage) {
    _schedulerWarn_(summary, 'tgSendMessage_ не найдена — сервисные Telegram-сообщения будут пропущены');
  }

  if (!summary.dependencies.tgCharts) {
    _schedulerWarn_(summary, 'sendDashboardChartsToTelegram_ не найдена — отправка графиков будет пропущена');
  }

  if (!summary.dependencies.tgStats) {
    _schedulerWarn_(summary, 'sendDashboardStatsFromDashboard_ не найдена — отправка статистики будет пропущена');
  }
}

function _schedulerRunPipeline_(opts) {
  opts = opts || {};

  var summary = _schedulerCreateSummary_(opts.jobName || 'pipeline');
  var chartAliases = opts.chartAliases || ['sectors', 'coupons', 'maturity', 'history', 'risk', 'ytmVsCoupon', 'scatter'];
  var stabilizationMs = Number(opts.stabilizationMs);
  if (!isFinite(stabilizationMs) || stabilizationMs < 0) stabilizationMs = 800;

  try {
    _schedulerRunStep_(summary, 'preflight-check', function() {
      _schedulerPreflight_(summary);
    }, true);

    if (opts.sendGreeting !== false) {
      _schedulerRunStep_(summary, 'greeting', function() {
        _schedulerSafeTelegram_(
          'Запускаю ' + summary.jobName +
          ': preflight → sync Input → update sheets → formatting → Dashboard → Telegram'
        );
      }, false);
    }

    _schedulerRunStep_(summary, 'sync-input', function() {
      summary.inputSync.attempted = true;

      if (summary.dependencies.inputSyncFn) {
        _schedulerInvokeByName_(summary.dependencies.inputSyncFn);
        summary.inputSync.synced = true;
      } else {
        summary.inputSync.skipped = true;
        summary.inputSync.reason = 'function missing';
      }
    }, false);

    _schedulerRunStep_(summary, 'update-bonds', function() {
      updateBondsFull();
      summary.updates.Bonds = true;
    }, true);

    _schedulerRunStep_(summary, 'update-funds', function() {
      updateFundsFull();
      summary.updates.Funds = true;
    }, true);

    _schedulerRunStep_(summary, 'update-shares', function() {
      updateSharesFull();
      summary.updates.Shares = true;
    }, true);

    _schedulerRunStep_(summary, 'update-options', function() {
      updateOptionsFull();
      summary.updates.Options = true;
    }, true);

    _schedulerRunStep_(summary, 'post-processing-formatting', function() {
      summary.formatting.attempted = true;
      summary.formatting.functionName = summary.dependencies.formattingFn || '';

      if (!summary.dependencies.formattingFn) {
        summary.formatting.skipped = true;
        summary.formatting.reason = 'function missing';
        return;
      }

      var formattingResult = _schedulerInvokeByName_(summary.dependencies.formattingFn);
      summary.formatting.result = formattingResult || null;

      if (formattingResult && formattingResult.skipped) {
        summary.formatting.skipped = true;
        summary.formatting.reason = formattingResult.reason || 'formatting skipped by module';
      } else {
        summary.formatting.applied = true;
      }
    }, false);

    _schedulerRunStep_(summary, 'build-dashboard', function() {
      summary.dashboard.attempted = true;
      summary.dashboard.functionName = summary.dependencies.dashboardFn || '';

      _schedulerInvokeByName_(summary.dependencies.dashboardFn);
    }, true);

    _schedulerRunStep_(summary, 'flush-stabilization', function() {
      SpreadsheetApp.flush();
      Utilities.sleep(stabilizationMs);
    }, true);

    _schedulerRunStep_(summary, 'post-run-audit', function() {
      _schedulerAuditSheets_(summary);
      if (!summary.dashboard.built) {
        throw new Error('Dashboard не построен или пуст после обновления');
      }
    }, true);

    if (opts.sendCharts) {
      _schedulerRunStep_(summary, 'telegram-charts', function() {
        if (summary.dependencies.tgCharts) {
          summary.chartsSent = Number(sendDashboardChartsToTelegram_(chartAliases)) || 0;
        }
      }, false);
    }

    if (opts.sendStats) {
      _schedulerRunStep_(summary, 'telegram-stats', function() {
        if (summary.dependencies.tgStats) {
          sendDashboardStatsFromDashboard_();
          summary.statsSent = true;
        }
      }, false);
    }

    summary.success = true;
    summary.elapsedMs = Date.now() - summary.startedAtMs;

    if (opts.sendServiceSummary) {
      _schedulerRunStep_(summary, 'service-summary', function() {
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
    chartAliases: ['sectors', 'coupons', 'maturity', 'history', 'risk', 'ytmVsCoupon', 'scatter'],
    stabilizationMs: 1200
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
    notifyErrors: true,
    stabilizationMs: 1200
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
