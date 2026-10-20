/**
 * Скрипт для Google Таблицы — принимает ответы с сайта-приглашения.
 * Создаёт два листа: «Анкета» (кто придёт) и «Тілектер» (пожелания).
 * Инструкция по установке — в файле README.md.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var p = (e && e.parameter) || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();

    if (p.type === 'wish') {
      var w = getSheet_(ss, 'Тілектер', ['Уақыты', 'Аты', 'Тілегі']);
      w.appendRow([now, p.name || '', p.message || '']);
    } else {
      var r = getSheet_(ss, 'Анкета', ['Уақыты', 'Аты-жөні', 'Келесіз бе?']);
      r.appendRow([now, p.name || '', p.attendance || '']);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('OK — скрипт работает');
}

function getSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
