/**
 * まるつけ — 初期化と講師アカウントの管理
 *
 * ここの関数はスクリプトエディタから手で実行する。Web API からは呼べない。
 */

/**
 * シートを一式作る。何度実行してもよい。
 *
 * 既にあるシートのデータには触らないが、あとから増えた列だけは右端に足す。
 * 列の順番は見出し名で引いているので、右端に足しても既存の並びは壊れない。
 */
function initialize() {
  var ss = ss_();
  var created = [];
  var addedCols = [];

  for (var key in SHEETS) {
    var name = SHEETS[key];
    var head = HEADERS[name];
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      created.push(name);
    }
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, head.length).setValues([head]);
      sh.getRange(1, 1, 1, head.length)
        .setFontWeight('bold')
        .setBackground('#EEF0F5');
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, head.length);
    } else {
      // 足りない見出しだけ右端に付け足す（列を増やしたときの移行）
      var map = headerMap_(sh);
      var missing = head.filter(function (h) { return map[h] === undefined; });
      if (missing.length) {
        var at = sh.getLastColumn() + 1;
        sh.getRange(1, at, 1, missing.length).setValues([missing])
          .setFontWeight('bold')
          .setBackground('#EEF0F5');
        addedCols.push(name + '（' + missing.join('、') + '）');
      }
    }
  }

  // 既定でできる「シート1」が空なら消す
  var first = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1 && first.getLastRow() === 0) {
    ss.deleteSheet(first);
  }

  var parts = [];
  if (created.length)   parts.push('作成したシート: ' + created.join('、'));
  if (addedCols.length) parts.push('列を足しました: ' + addedCols.join(' / '));
  var msg = parts.length ? parts.join('\n') : 'シートは最新の状態です。';
  Logger.log(msg);
  return msg;
}

/**
 * 講師を足す。ログインは無いので名前だけ。
 * 記録に「誰が付けたか」を残すためのもの。
 */
function createTeacher(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('名前が空です。');

  var id = newId_('t');
  appendRow_(SHEETS.TEACHER, {
    'id':    id,
    '名前':   name,
    '有効':   true,
    '作成日時': new Date()
  });
  return '追加しました: ' + name;
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

