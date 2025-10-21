/**
 * tgbot_api.gs
 * Простые помощники для Telegram: читают токен/чат из Script Properties.
 *
 * Ожидаемые ключи в Project → Script properties:
 *   TELEGRAM_BOT_TOKEN = bot1234567890:XXXXX
 *   TELEGRAM_CHAT_ID   = 123456789  (или @channel_username)
 *
 * Поддерживаются альтернативные имена: TG_BOT_TOKEN, TG_CHAT_ID (fallback).
 * Если переменные лежат в User Properties — тоже подхватим как запасной вариант.
 */

/** Прочитать конфиг из Script/User Properties. Бросает ошибку, если чего-то нет. */
function tgConfig_() {
  var sp = PropertiesService.getScriptProperties();
  var up = PropertiesService.getUserProperties();

  var token =
    (sp.getProperty('TELEGRAM_BOT_TOKEN') || up.getProperty('TELEGRAM_BOT_TOKEN')) ||
    (sp.getProperty('TG_BOT_TOKEN')       || up.getProperty('TG_BOT_TOKEN'));

  var chatId =
    (sp.getProperty('TELEGRAM_CHAT_ID') || up.getProperty('TELEGRAM_CHAT_ID')) ||
    (sp.getProperty('TG_CHAT_ID')       || up.getProperty('TG_CHAT_ID'));

  if (!token) throw new Error('Telegram: не найден TELEGRAM_BOT_TOKEN (или TG_BOT_TOKEN) в Script/User Properties');
  if (!chatId) throw new Error('Telegram: не найден TELEGRAM_CHAT_ID (или TG_CHAT_ID) в Script/User Properties');

  return { token: String(token).trim(), chatId: String(chatId).trim() };
}

/** Универсальный POST: если в payload есть Blob — отправляем multipart/form-data. */
function tgPost_(methodName, payload) {
  var cfg = tgConfig_();
  var url = 'https://api.telegram.org/bot' + cfg.token + '/' + methodName;

  // Определяем, есть ли Blob в значениях payload
  var hasBlob = false;
  if (payload && typeof payload === 'object') {
    for (var k in payload) {
      if (payload[k] && typeof payload[k].getAs === 'function') { hasBlob = true; break; }
    }
  }

  var options = {
    method: 'post',
    muteHttpExceptions: true
  };

  // Всегда передаём chat_id, если его нет в payload
  if (payload && payload.chat_id == null) payload.chat_id = cfg.chatId;

  if (hasBlob) {
    // multipart/form-data — достаточно передать объект с Blob
    options.payload = payload;
  } else {
    // application/x-www-form-urlencoded — удобен для простых запросов
    options.payload = payload;
  }

  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  if (code !== 200) {
    Logger.log('Telegram API ' + methodName + ' → ' + code + ' ' + resp.getContentText());
  }
  return resp;
}

/**
 * Отправить текстовое сообщение.
 * @param {string} text
 * @param {object=} opts  { parse_mode?: 'HTML'|'MarkdownV2'|..., disable_preview?: boolean, reply_to_message_id?: number }
 */
function tgSendMessage_(text, opts) {
  opts = opts || {};
  var payload = {
    text: String(text || ''),
    parse_mode: (opts.parse_mode || 'HTML'),
    disable_web_page_preview: (opts.disable_preview == null ? true : !!opts.disable_preview)
  };
  if (opts.reply_to_message_id != null) payload.reply_to_message_id = opts.reply_to_message_id;
  return tgPost_('sendMessage', payload);
}

/**
 * Отправить фото.
 * @param {Blob|string} photo - Blob (картинка) ИЛИ строка (URL/файл_id).
 * @param {string=} caption   - подпись (можно с HTML/MarkdownV2 — см. opts.parse_mode).
 * @param {object=} opts      - { parse_mode?: 'HTML', has_spoiler?: boolean }
 */
function tgSendPhoto_(photo, caption, opts) {
  opts = opts || {};
  var payload = {
    photo: photo, // Apps Script сам сделает multipart, если это Blob
    caption: caption || '',
    parse_mode: (opts.parse_mode || 'HTML')
  };
  if (opts.has_spoiler) payload.has_spoiler = true;
  return tgPost_('sendPhoto', payload);
}

/** (Опционально) Отправить документ (xlsx/pdf/zip и т.д.). */
function tgSendDocument_(documentBlobOrUrl, caption, opts) {
  opts = opts || {};
  var payload = {
    document: documentBlobOrUrl,
    caption: caption || '',
    parse_mode: (opts.parse_mode || 'HTML')
  };
  return tgPost_('sendDocument', payload);
}

/** Быстрый просмотр апдейтов — помогает найти chat_id при отладке. */
function logTelegramUpdates() {
  var cfg = tgConfig_();
  var url = 'https://api.telegram.org/bot' + cfg.token + '/getUpdates';
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log(resp.getContentText());
}


/**
 * tgbot_webhook_min.gs
 * Простой вебхук: /start → ответ в чат + сохранение chat_id + одноразовый запуск weeklyRefreshAndNotify().
 * Требует TELEGRAM_BOT_TOKEN в Script Properties. Для рассылок берётся TELEGRAM_CHAT_ID (мы его сохраняем на /start).
 */

/** ==== Мини-утилиты ==== */
function _tgToken_() {
  var t = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!t) throw new Error('TELEGRAM_BOT_TOKEN не задан в свойствах скрипта');
  return t;
}
function tgSendMessageTo_(chatId, text) {
  var url = 'https://api.telegram.org/bot' + _tgToken_() + '/sendMessage';
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: { chat_id: String(chatId), text: String(text || '') },
    muteHttpExceptions: true
  });
  Logger.log('sendMessageTo [' + chatId + ']: ' + resp.getResponseCode() + ' ' + resp.getContentText());
}
function tgSendMessage_(text) { // оставляем совместимость с существующим кодом
  var chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
  if (!chatId) throw new Error('TELEGRAM_CHAT_ID не задан — сначала отправьте /start боту');
  tgSendMessageTo_(chatId, text);
}

/** ==== Вебхук ==== */
/** ================== WEBHOOK HANDLER (минимальный и стабильный) ================== */
function doPost(e) {
  try {
    var upd = JSON.parse(e.postData && e.postData.contents || '{}');
    tgHandleUpdate_(upd);
  } catch (err) {
    Logger.log('doPost error: ' + (err && err.stack || err));
  }
  // ВАЖНО: всегда быстро отдаём 200 OK — иначе Телеграм ретраит и вы получите дубли.
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

/** Обработка одного апдейта */
function tgHandleUpdate_(upd) {
  // 0) Дедуп по update_id — защитит от ретраев Телеги
  if (_tgSeenUpdate_(upd.update_id)) return;

  // 1) Принимаем только текстовые "message" (игнорируем edited_message, my_chat_member и т.п.)
  if (!upd || !upd.message || typeof upd.message.text !== 'string') return;

  var text = (upd.message.text || '').trim();
  var chatId = String(upd.message.chat && upd.message.chat.id || '');
  var ONLY_CHAT = (PropertiesService.getScriptProperties().getProperty('TG_CHAT_ID') || '').trim();
  if (ONLY_CHAT && String(chatId) !== String(ONLY_CHAT)) {
    // чужие чаты игнорим
    return;
  }

  // 2) Простая маршрутизация по командам. Любая неизвестная команда — короткий help, но
  //    только если это САМОМУ текст начинается с '/'. Обычный текст игнорим.
  if (text === '/start') {
    tgSendMessage_('Принял команду. Запускаю разовое обновление и сбор дашборда — пришлю отчёт сюда.');
    _enqueueWeeklyJobOnce_();  // ставим одноразовый триггер, задача выполнится вне вебхука
    return;
  }

  if (text === '/help' || text === '/commands') {
    tgSendMessage_(
      'Команды:\n' +
      '/start — запустить разовое обновление и прислать отчёт\n' +
      '/help — показать команды'
    );
    return;
  }

  if (text.charAt(0) === '/') {
    tgSendMessage_('Не знаю эту команду. Напишите /help');
    return;
  }

  // Обычный текст — молча игнорим (чтобы бот не «болтал» без причины)
}

/** ===== ДЕДУПЛИКАЦИЯ АПДЕЙТОВ ПО update_id ===== */
function _tgSeenUpdate_(updateId) {
  if (typeof updateId !== 'number') return false;
  var props = PropertiesService.getScriptProperties();
  var last = Number(props.getProperty('TG_LAST_UPDATE_ID') || 0);
  if (updateId <= last) return true; // уже видели/ответили
  props.setProperty('TG_LAST_UPDATE_ID', String(updateId));
  return false;
}

/** ====== ОДНОРАЗОВЫЙ ЗАПУСК ТЯЖЁЛОЙ ЗАДАЧИ ВНЕ ВЕБХУКА ====== */
function _enqueueWeeklyJobOnce_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) return; // если одновременно жмут /start — выполнит один
  try {
    var props = PropertiesService.getScriptProperties();
    var FLAG = 'WEEKLY_JOB_SCHEDULED';
    if (props.getProperty(FLAG)) return; // уже поставлено — не дублируем

    ScriptApp.newTrigger('weeklyRefreshAndNotify')  // ваша существующая тяжёлая задача
      .timeBased()
      .after(1000) // через секунду
      .create();

    props.setProperty(FLAG, String(Date.now()));
  } finally {
    lock.releaseLock();
  }
}

// (необязательно) сброс флага — вызовите в конце weeklyRefreshAndNotify
function _weeklyJobClearFlag_() {
  PropertiesService.getScriptProperties().deleteProperty('WEEKLY_JOB_SCHEDULED');
}


/** ==== Утилиты вебхука ==== */
function tgSetWebhookUrl_(execUrl, dropPending) {
  if (!execUrl || !/\/exec$/.test(execUrl)) {
    throw new Error('Передай сюда URL веб-приложения, который оканчивается на /exec');
  }
  var resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + _tgToken_() + '/setWebhook', {
    method: 'post',
    payload: {
      url: execUrl,
      drop_pending_updates: dropPending ? 'true' : 'false'
    },
    muteHttpExceptions: true
  });
  Logger.log('setWebhook: ' + resp.getResponseCode() + ' ' + resp.getContentText());
}

function tgDeleteWebhook() {
  var resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + _tgToken_() + '/deleteWebhook', {
    method: 'post',
    muteHttpExceptions: true
  });
  Logger.log('deleteWebhook: ' + resp.getResponseCode() + ' ' + resp.getContentText());
}
function tgGetWebhookInfo() {
  var resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + _tgToken_() + '/getWebhookInfo', { muteHttpExceptions: true });
  Logger.log('getWebhookInfo: ' + resp.getResponseCode() + ' ' + resp.getContentText());
}

function sos(){
  tgDeleteWebhook();         // снять старый
  tgGetWebhookInfo();        // проверить: url должен быть .../exec, pending_update_count: 0

}