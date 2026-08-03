/**
 * まるつけ — Web API
 *
 * すべて POST。本文は JSON だが Content-Type は text/plain で送ってもらう。
 * application/json にすると CORS の事前確認（OPTIONS）が飛び、GAS はそれに応答できないため。
 *
 * 形式:  { action: "...", token: "...", ... }
 * 返り値: { ok: true, data: ... } または { ok: false, error: "..." }
 */

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'リクエストの形式が不正です。' });
  }

  try {
    return json_({ ok: true, data: route_(body) });
  } catch (err) {
    return json_({ ok: false, error: (err && err.message) ? err.message : String(err) });
  }
}

/** ブラウザで直接開いたときの案内。データは一切返さない */
function doGet() {
  return ContentService
    .createTextOutput('まるつけ API です。この URL はアプリから使います。')
    .setMimeType(ContentService.MimeType.TEXT);
}

function selfCheck_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = { sheets: {} };

  if (!ss) return { error: 'スプレッドシートに紐付いていません。' };

  for (var key in SHEETS) {
    var name = SHEETS[key];
    var sh = ss.getSheetByName(name);
    if (!sh) { out.sheets[name] = 'ありません'; continue; }

    var head = [];
    if (sh.getLastColumn() > 0) {
      head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
               .map(function (v) { return String(v).trim(); });
    }
    var missing = (HEADERS[name] || []).filter(function (h) { return head.indexOf(h) < 0; });
    out.sheets[name] = {
      行数: Math.max(0, sh.getLastRow() - 1),
      足りない列: missing
    };
  }

  out.アクセスキー設定済み = !!getAccessKey_();
  return out;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function route_(b) {
  var action = String(b.action || '');

  // 鍵が要らないのはこの2つだけ。どちらもデータを返さない
  if (action === 'ping')  return { ok: true, time: Date.now() };

  /**
   * 不具合を切り分けるための自己点検。認証は要らない。
   * 名前・ハッシュ・IDなど中身は一切返さない。構造と件数だけ。
   * スキーマは公開リポジトリにあるので、これで増える情報は無い。
   */
  if (action === 'selfCheck') return selfCheck_();
  // ここから先は必ずアクセスキーを確認する
  requireKey_(b.key);
  var teacher = teacherOf_(b.teacher);

  switch (action) {
    /** 起動直後の一括取得。マスタと記録をまとめて返す */
    case 'bootstrap':
      return {
        masters: getMasters_(),
        records: getRecords_({ days: b.days })
      };

    case 'masters':
      return getMasters_();

    case 'records':
      return getRecords_({ studentId: b.studentId, days: b.days });

    case 'addRecord':
      return addRecord_(teacher, b);

    case 'deleteRecord':
      return deleteRecord_(teacher, b.id);

    case 'setCause':
      return setCause_(teacher, b.id, b.cause);

    /** 応答が落ちたときの後始末。受付キーで記録が入ったか確かめる */
    case 'findRecord':
      return findRecordByKey_(b.recKey);

    case 'addSection':
      return addSection_(teacher, b.bookId, b.label);

    case 'addStudent':
      return addStudent_(teacher, b.name, b.grade);

    case 'addBook':
      return addBook_(teacher, b);

    default:
      throw new Error('不明な操作です: ' + action);
  }
}
