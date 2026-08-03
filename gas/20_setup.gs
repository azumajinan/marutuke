/**
 * まるつけ — 初期化と講師アカウントの管理
 *
 * ここの関数はスクリプトエディタから手で実行する。Web API からは呼べない。
 */

/**
 * シートを一式作る。既にあるシートには触らない。
 * スクリプトエディタで一度だけ実行する。
 */
function initialize() {
  var ss = ss_();
  var created = [];

  for (var key in SHEETS) {
    var name = SHEETS[key];
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      created.push(name);
    }
    if (sh.getLastRow() === 0) {
      var head = HEADERS[name];
      sh.getRange(1, 1, 1, head.length).setValues([head]);
      sh.getRange(1, 1, 1, head.length)
        .setFontWeight('bold')
        .setBackground('#EEF0F5');
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, head.length);
    }
  }

  // 既定でできる「シート1」が空なら消す
  var first = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1 && first.getLastRow() === 0) {
    ss.deleteSheet(first);
  }

  var msg = created.length
    ? '作成したシート: ' + created.join(', ')
    : 'すべてのシートは既にあります。';
  Logger.log(msg);
  return msg;
}

/**
 * 講師アカウントを作る。
 *
 * 使い方: 下の値を書き換えて、この関数をスクリプトエディタから実行する。
 * 実行したらパスワードの部分は消しておくこと（スクリプトは共同編集者に見える）。
 */
function createTeacherManually() {
  var 名前     = '高橋';
  var ログインID = 'takahashi';
  var パスワード  = 'ここに決めたパスワード';

  Logger.log(createTeacher(名前, ログインID, パスワード));
}

function createTeacher(name, loginId, password) {
  if (!name || !loginId || !password) throw new Error('名前・ログインID・パスワードは必須です。');
  if (String(password).length < 8) throw new Error('パスワードは8文字以上にしてください。');

  var teachers = readAll_(SHEETS.TEACHER);
  for (var i = 0; i < teachers.length; i++) {
    if (String(teachers[i]['ログインID']).trim() === String(loginId).trim()) {
      throw new Error('そのログインIDは既に使われています: ' + loginId);
    }
  }

  var salt = newSalt_();
  appendRow_(SHEETS.TEACHER, {
    'id':            newId_('t'),
    '名前':           name,
    'ログインID':      loginId,
    'パスワードハッシュ': hashPassword_(password, salt),
    'ソルト':          salt,
    '有効':           true,
    '作成日時':        new Date()
  });
  return '作成しました: ' + name + '（ログインID: ' + loginId + '）';
}

/** パスワードを変える。上と同じくエディタから実行する */
function changePasswordManually() {
  var ログインID    = 'takahashi';
  var 新しいパスワード = 'ここに新しいパスワード';

  Logger.log(changePassword(ログインID, 新しいパスワード));
}

function changePassword(loginId, newPassword) {
  if (String(newPassword).length < 8) throw new Error('パスワードは8文字以上にしてください。');

  var teachers = readAll_(SHEETS.TEACHER);
  for (var i = 0; i < teachers.length; i++) {
    if (String(teachers[i]['ログインID']).trim() === String(loginId).trim()) {
      var salt = newSalt_();
      updateRow_(SHEETS.TEACHER, teachers[i]._row, {
        'ソルト':          salt,
        'パスワードハッシュ': hashPassword_(newPassword, salt)
      });
      // その講師のセッションを全部切る
      revokeSessionsOf_(teachers[i]['id']);
      return 'パスワードを変更しました: ' + loginId + '（ログイン中の端末は切断されます）';
    }
  }
  throw new Error('その講師が見つかりません: ' + loginId);
}

/**
 * 動作確認用のサンプルデータを入れる。
 * 実運用のスプレッドシートでは実行しないこと。
 */
function seedSampleData() {
  var b1 = newId_('b'), b2 = newId_('b');

  appendRow_(SHEETS.STUDENT, { 'id': newId_('s'), '名前': '田中 涼',   '学年': '中3', '有効': true, '作成日時': new Date() });
  appendRow_(SHEETS.STUDENT, { 'id': newId_('s'), '名前': '鈴木 陽向', '学年': '中2', '有効': true, '作成日時': new Date() });

  appendRow_(SHEETS.BOOK, { 'id': b1, '書名': '新中学問題集 数学3年',     '略称': '新中問 数3',  '科目': '数学', '総ページ': 168, '有効': true, '作成日時': new Date() });
  appendRow_(SHEETS.BOOK, { 'id': b2, '書名': '中学必修テキスト 数学2年', '略称': '必修テキ 数2', '科目': '数学', '総ページ': 152, '有効': true, '作成日時': new Date() });

  var u1 = [['式の計算',8,27],['平方根',28,45],['二次方程式',46,69],['二次関数',70,95],['相似',96,121]];
  var u2 = [['式の計算',6,25],['連立方程式',26,49],['一次関数',50,77],['図形の性質',78,105],['確率',106,127]];

  u1.forEach(function (u) {
    appendRow_(SHEETS.UNIT, { 'id': newId_('u'), '教材id': b1, '教材略称': '新中問 数3', '単元名': u[0], '開始ページ': u[1], '終了ページ': u[2] });
  });
  u2.forEach(function (u) {
    appendRow_(SHEETS.UNIT, { 'id': newId_('u'), '教材id': b2, '教材略称': '必修テキ 数2', '単元名': u[0], '開始ページ': u[1], '終了ページ': u[2] });
  });

  ['基本問題','練習問題','章末問題'].forEach(function (s) {
    appendRow_(SHEETS.SECTION, { 'id': newId_('h'), '教材id': b1, '教材略称': '新中問 数3', '見出し': s });
  });
  ['確認問題','練習問題','章末問題'].forEach(function (s) {
    appendRow_(SHEETS.SECTION, { 'id': newId_('h'), '教材id': b2, '教材略称': '必修テキ 数2', '見出し': s });
  });

  return 'サンプルデータを入れました。';
}

/** 期限切れのセッション行を掃除する。日次トリガーに設定してもよい */
function cleanupSessions() {
  var rows = readAll_(SHEETS.SESSION);
  var limit = Date.now() - AUTH.SESSION_HOURS * 3600 * 1000;
  var removed = 0;
  // 下から消す。上から消すと行番号がずれる
  for (var i = rows.length - 1; i >= 0; i--) {
    var last = new Date(rows[i]['最終利用日時'] || rows[i]['発行日時']).getTime();
    if (!last || last < limit) {
      deleteRow_(SHEETS.SESSION, rows[i]._row);
      removed++;
    }
  }
  return removed + ' 件のセッションを削除しました。';
}
