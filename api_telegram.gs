/**
 * api_telegram.gs
 * Telegram helpers + webhook commands.
 *
 * Script Properties:
 *   TELEGRAM_BOT_TOKEN = bot1234567890:XXXXX
 *   TELEGRAM_CHAT_ID   = 123456789
 *
 * Fallback keys:
 *   TG_BOT_TOKEN
 *   TG_CHAT_ID
 */

// ===================== Config =====================

function tgToken_() {
  var sp = PropertiesService.getScriptProperties();
  var up = PropertiesService.getUserProperties();

  var token =
    sp.getProperty('TELEGRAM_BOT_TOKEN') ||
    up.getProperty('TELEGRAM_BOT_TOKEN') ||
    sp.getProperty('TG_BOT_TOKEN') ||
    up.getProperty('TG_BOT_TOKEN');

  if (!token) {
    throw new Error('Telegram: не найден TELEGRAM_BOT_TOKEN или TG_BOT_TOKEN');
  }

  return String(token).trim();
}

function tgDefaultChatId_() {
  var sp = PropertiesService.getScriptProperties();
  var up = PropertiesService.getUserProperties();

  var chatId =
    sp.getProperty('TELEGRAM_CHAT_ID') ||
    up.getProperty('TELEGRAM_CHAT_ID') ||
    sp.getProperty('TG_CHAT_ID') ||
    up.getProperty('TG_CHAT_ID');

  if (!chatId) {
    throw new Error('Telegram: не найден TELEGRAM_CHAT_ID или TG_CHAT_ID');
  }

  return String(chatId).trim();
}

function tgConfig_() {
  return {
    token: tgToken_(),
    chatId: tgDefaultChatId_()
  };
}

function tgRememberChatIdIfMissing_(chatId) {
  chatId = String(chatId || '').trim();
  if (!chatId) return;

  var sp = PropertiesService.getScriptProperties();

  var existing =
    sp.getProperty('TELEGRAM_CHAT_ID') ||
    sp.getProperty('TG_CHAT_ID');

  if (String(existing || '').trim()) return;

  sp.setProperty('TELEGRAM_CHAT_ID', chatId);
  sp.setProperty('TG_CHAT_ID', chatId);
}

function tgIsAllowedChat_(chatId) {
  chatId = String(chatId || '').trim();
  if (!chatId) return false;

  var sp = PropertiesService.getScriptProperties();

  var allowed =
    sp.getProperty('TELEGRAM_CHAT_ID') ||
    sp.getProperty('TG_CHAT_ID') ||
    '';

  allowed = String(allowed || '').trim();

  if (!allowed) return true;

  return chatId === allowed;
}

// ===================== Telegram API =====================

function tgPost_(methodName, payload) {
  payload = payload || {};

  if (payload.chat_id == null) {
    payload.chat_id = tgDefaultChatId_();
  }

  var url = 'https://api.telegram.org/bot' + tgToken_() + '/' + methodName;

  var options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();

  if (code !== 200) {
    Logger.log(
      'Telegram API ' +
      methodName +
      ' → ' +
      code +
      ' ' +
      resp.getContentText()
    );
  }

  return resp;
}

function tgSendMessage_(text, opts) {
  opts = opts || {};

  var payload = {
    text: String(text || ''),
    parse_mode: opts.parse_mode || 'HTML',
    disable_web_page_preview: opts.disable_preview == null ? true : !!opts.disable_preview
  };

  if (opts.reply_to_message_id != null) {
    payload.reply_to_message_id = opts.reply_to_message_id;
  }

  return tgPost_('sendMessage', payload);
}

function tgSendMessageTo_(chatId, text, opts) {
  opts = opts || {};

  var payload = {
    chat_id: String(chatId),
    text: String(text || ''),
    parse_mode: opts.parse_mode || 'HTML',
    disable_web_page_preview: opts.disable_preview == null ? true : !!opts.disable_preview
  };

  if (opts.reply_to_message_id != null) {
    payload.reply_to_message_id = opts.reply_to_message_id;
  }

  return tgPost_('sendMessage', payload);
}

function tgSendPhoto_(photo, caption, opts) {
  opts = opts || {};

  var payload = {
    photo: photo,
    caption: String(caption || ''),
    parse_mode: opts.parse_mode || 'HTML'
  };

  if (opts.has_spoiler) {
    payload.has_spoiler = true;
  }

  return tgPost_('sendPhoto', payload);
}

function tgSendDocument_(documentBlobOrUrl, caption, opts) {
  opts = opts || {};

  var payload = {
    document: documentBlobOrUrl,
    caption: String(caption || ''),
    parse_mode: opts.parse_mode || 'HTML'
  };

  return tgPost_('sendDocument', payload);
}

function logTelegramUpdates() {
  var resp = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + tgToken_() + '/getUpdates',
    { muteHttpExceptions: true }
  );

  Logger.log(resp.getContentText());
}

// ===================== Webhook =====================

function doPost(e) {
  try {
    var upd = JSON.parse(e.postData && e.postData.contents || '{}');
    tgHandleUpdate_(upd);
  } catch (err) {
    Logger.log('doPost error: ' + (err && err.stack || err));
  }

  return ContentService
    .createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

function tgHandleUpdate_(upd) {
  if (!upd) return;
  if (_tgSeenUpdate_(upd.update_id)) return;
  if (!upd.message || typeof upd.message.text !== 'string') return;

  var text = String(upd.message.text || '').trim();
  var chatId = String(upd.message.chat && upd.message.chat.id || '').trim();

  if (!tgIsAllowedChat_(chatId)) return;

  tgRememberChatIdIfMissing_(chatId);

  var cmd = text.split(/\s+/)[0].split('@')[0].toLowerCase();

  if (cmd === '/start' || cmd === '/help' || cmd === '/commands') {
    tgSendMessageTo_(
      chatId,
      'Команды:\n' +
      '/refresh — обновить таблицу, построить Dashboard и отправить отчёт'
    );
    return;
  }

  if (cmd === '/refresh') {
    tgSendMessageTo_(chatId, 'Запускаю обновление. Отчёт придёт сюда.');
    _enqueueRefreshJobOnce_();
    return;
  }

  if (cmd.charAt(0) === '/') {
    tgSendMessageTo_(chatId, 'Неизвестная команда. Напишите /help');
  }
}

function _tgSeenUpdate_(updateId) {
  if (typeof updateId !== 'number') return false;

  var props = PropertiesService.getScriptProperties();
  var last = Number(props.getProperty('TG_LAST_UPDATE_ID') || 0);

  if (updateId <= last) return true;

  props.setProperty('TG_LAST_UPDATE_ID', String(updateId));
  return false;
}

// ===================== Refresh trigger =====================

function _enqueueRefreshJobOnce_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) return;

  try {
    var props = PropertiesService.getScriptProperties();
    var flag = 'TG_REFRESH_JOB_SCHEDULED';
    var oldTs = Number(props.getProperty(flag) || 0);
    var now = Date.now();

    if (oldTs && now - oldTs < 30 * 60 * 1000) {
      return;
    }

    props.deleteProperty(flag);
    _deleteTriggersByHandler_('tgRefreshJob_');

    ScriptApp.newTrigger('tgRefreshJob_')
      .timeBased()
      .after(1000)
      .create();

    props.setProperty(flag, String(now));
  } finally {
    lock.releaseLock();
  }
}

function tgRefreshJob_() {
  try {
    if (typeof weeklyRefreshAndNotify !== 'function') {
      throw new Error('weeklyRefreshAndNotify() не найдена');
    }

    weeklyRefreshAndNotify();

  } catch (e) {
    Logger.log('tgRefreshJob_ error: ' + (e && e.stack || e));

    try {
      tgSendMessage_(
        'Refresh: ошибка\n' +
        String(e && e.message ? e.message : e)
      );
    } catch (_) {}

  } finally {
    PropertiesService.getScriptProperties().deleteProperty('TG_REFRESH_JOB_SCHEDULED');
    _deleteTriggersByHandler_('tgRefreshJob_');
  }
}

function _deleteTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function(tr) {
    if (tr.getHandlerFunction && tr.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(tr);
    }
  });
}

// ===================== Webhook setup =====================

function tgSetWebhookUrl_(execUrl, dropPending) {
  if (!execUrl || !/\/exec$/.test(execUrl)) {
    throw new Error('Передай URL веб-приложения, который оканчивается на /exec');
  }

  var resp = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + tgToken_() + '/setWebhook',
    {
      method: 'post',
      payload: {
        url: execUrl,
        drop_pending_updates: dropPending ? 'true' : 'false'
      },
      muteHttpExceptions: true
    }
  );

  Logger.log(
    'setWebhook: ' +
    resp.getResponseCode() +
    ' ' +
    resp.getContentText()
  );
}

function tgDeleteWebhook() {
  var resp = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + tgToken_() + '/deleteWebhook',
    {
      method: 'post',
      muteHttpExceptions: true
    }
  );

  Logger.log(
    'deleteWebhook: ' +
    resp.getResponseCode() +
    ' ' +
    resp.getContentText()
  );
}

function tgGetWebhookInfo() {
  var resp = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + tgToken_() + '/getWebhookInfo',
    {
      muteHttpExceptions: true
    }
  );

  Logger.log(
    'getWebhookInfo: ' +
    resp.getResponseCode() +
    ' ' +
    resp.getContentText()
  );
}

function sos() {
  tgDeleteWebhook();
  tgGetWebhookInfo();
}

// ===================== Dashboard charts =====================

function sendDashboardChartsToTelegram_(aliases) {
  SpreadsheetApp.flush();

  var pack = exportDashboardCharts_(aliases);
  var sent = 0;
  var order = aliases && aliases.length ? aliases.slice() : Object.keys(pack);

  for (var i = 0; i < order.length; i++) {
    var alias = order[i];
    var item = pack[alias];

    if (!item || !item.blob) continue;

    try {
      tgSendPhoto_(item.blob, item.caption || '');
      sent++;
    } catch (e) {
      Logger.log(
        'sendDashboardChartsToTelegram_ [' +
        alias +
        '] error: ' +
        (e && e.stack ? e.stack : e)
      );
    }
  }

  return sent;
}
