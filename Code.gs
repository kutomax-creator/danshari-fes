/**
 * 断捨離フェス - バックエンド (Google Apps Script)
 *
 * セットアップ:
 * 1. 新しい Google スプレッドシートを作成
 * 2. 拡張機能 > Apps Script を開く
 * 3. このファイルの内容を全部貼り付けて保存
 * 4. 右上「デプロイ」> 新しいデプロイ > 種類「ウェブアプリ」
 *      - 実行するユーザー: 自分
 *      - アクセスできるユーザー: 全員
 * 5. デプロイ後に表示される URL (.../exec) を index.html の APPS_SCRIPT_URL にコピペ
 * 6. 初回アクセス時に「承認が必要です」と出るので承認する
 *
 * シートスキーマ: id | text | reactions(JSON) | ts | color
 * 旧スキーマ(votes数値)との後方互換あり
 */

var SHEET_NAME = 'Comments';
var VALID_EMOJIS = ['👍','👏','🔥','❓'];

function doGet(e) {
  var action   = (e.parameter.action || 'list');
  var callback = e.parameter.callback;
  var result;

  try {
    var sheet = getSheet_();
    var cache = CacheService.getScriptCache();

    if (action === 'list') {
      var cached = cache.get('list');
      if (cached) {
        result = JSON.parse(cached);
      } else {
        result = { ok: true, comments: getComments_(sheet), title: getTitle_() };
        cache.put('list', JSON.stringify(result), 2);
      }

    } else {
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(5000);

        if (action === 'submit') {
          var text  = (e.parameter.text  || '').toString().trim().slice(0, 280);
          var color = (e.parameter.color || '').toString().replace(/[^#0-9a-fA-F]/g, '').slice(0, 7);
          if (text) {
            var initR = JSON.stringify({'👍':0,'👏':0,'🔥':0,'❓':0});
            sheet.appendRow([Utilities.getUuid(), text, initR, new Date().getTime(), color]);
          }
          result = { ok: true, comments: getComments_(sheet), title: getTitle_() };

        } else if (action === 'vote') {
          voteComment_(sheet, e.parameter.id, e.parameter.type);
          result = { ok: true, comments: getComments_(sheet), title: getTitle_() };

        } else if (action === 'setTitle') {
          setTitle_((e.parameter.title || '').toString().slice(0, 60));
          result = { ok: true, comments: getComments_(sheet), title: getTitle_() };

        } else if (action === 'reset') {
          resetComments_(sheet);
          result = { ok: true, comments: [], title: getTitle_() };

        } else {
          result = { ok: false, error: 'unknown action' };
        }
        cache.remove('list');
      } finally {
        try { lock.releaseLock(); } catch(e2) {}
      }
    }
  } catch(err) {
    result = { ok: false, error: String(err) };
  }

  var json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function parseReactions_(val) {
  if (!val && val !== 0) return {'👍':0,'👏':0,'🔥':0,'❓':0};
  if (typeof val === 'number') return {'👍':val,'👏':0,'🔥':0,'❓':0};
  var s = String(val);
  if (!s || s === '0') return {'👍':0,'👏':0,'🔥':0,'❓':0};
  try {
    var n = Number(s);
    if (!isNaN(n)) return {'👍':n,'👏':0,'🔥':0,'❓':0};
    return JSON.parse(s);
  } catch(e) {
    return {'👍':0,'👏':0,'🔥':0,'❓':0};
  }
}

function getSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id','text','reactions','ts','color']);
  }
  return sheet;
}

function getComments_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = Math.max(sheet.getLastColumn(), 4);
  var values  = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    out.push({
      id:        String(row[0]),
      text:      String(row[1]),
      reactions: parseReactions_(row[2]),
      ts:        Number(row[3]) || 0,
      color:     lastCol >= 5 ? String(row[4] || '') : ''
    });
  }
  return out;
}

function voteComment_(sheet, id, type) {
  if (!id) return;
  var emoji = (VALID_EMOJIS.indexOf(type) !== -1) ? type : '👍';
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      var cell      = sheet.getRange(i + 2, 3);
      var reactions = parseReactions_(cell.getValue());
      reactions[emoji] = (reactions[emoji] || 0) + 1;
      cell.setValue(JSON.stringify(reactions));
      break;
    }
  }
}

function resetComments_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
}

function getTitle_() {
  return PropertiesService.getScriptProperties().getProperty('title') || '断捨離フェス';
}

function setTitle_(t) {
  if (!t) return;
  PropertiesService.getScriptProperties().setProperty('title', t);
}
