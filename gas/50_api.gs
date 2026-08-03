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

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function route_(b) {
  var action = String(b.action || '');

  // 認証不要なのはこの2つだけ
  if (action === 'ping')  return { ok: true, time: Date.now() };
  if (action === 'login') return login_(b.loginId, b.password, b.key);

  // ここから先は必ずトークンを確認する
  var teacher = requireTeacher_(b.token);

  switch (action) {
    case 'logout':
      logout_(b.token);
      return true;

    case 'me':
      return { teacher: { id: teacher.id, name: teacher.name } };

    /** ログイン直後の一括取得。マスタと直近の記録をまとめて返す */
    case 'bootstrap':
      return {
        teacher: { id: teacher.id, name: teacher.name },
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
      return findRecordByKey_(b.key);

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
