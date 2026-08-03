/**
 * まるつけ — スプレッドシートのメニュー
 *
 * 講師・生徒・教材の管理をスクリプトエディタでやらずに済ませる。
 * スプレッドシートを開くと上部に「まるつけ」メニューが出る。
 *
 * メニューから呼ぶ関数は末尾に _ を付けないこと。
 * _ 付きは GAS が「外から呼べない関数」として扱うため、メニューからも呼べない。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('まるつけ')
    .addItem('シートを作る（最初に1回）', 'menuInitialize')
    .addSeparator()
    .addItem('講師を追加…', 'menuAddTeacher')
    .addItem('パスワードを変更…', 'menuChangePassword')
    .addSeparator()
    .addItem('手で足した行を整える', 'menuNormalize')
    .addSeparator()
    .addItem('古いセッションを掃除', 'menuCleanupSessions')
    .addItem('全員をログアウトさせる', 'menuLogoutAll')
    .addToUi();
}

/* ---------- 画面まわりの小道具 ---------- */

function ui_() { return SpreadsheetApp.getUi(); }

/** メニューから呼ぶ処理を包む。例外はダイアログで見せる */
function run_(title, fn) {
  try {
    var msg = fn();
    if (msg) ui_().alert(title, msg, ui_().ButtonSet.OK);
  } catch (err) {
    ui_().alert(title, 'できませんでした。\n\n' + ((err && err.message) || err), ui_().ButtonSet.OK);
  }
}

/** 入力を求める。キャンセルなら null */
function ask_(title, prompt) {
  var res = ui_().prompt(title, prompt, ui_().ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui_().Button.OK) return null;
  return String(res.getResponseText() || '').trim();
}

/* ---------- メニューの中身 ---------- */

function menuInitialize() {
  run_('シートを作る', function () { return initialize(); });
}

function menuAddTeacher() {
  run_('講師を追加', function () {
    var name = ask_('講師を追加', '名前を入れてください（例: 高橋）');
    if (name === null) return '';
    if (!name) throw new Error('名前が空です。');

    var loginId = ask_('講師を追加', 'ログインIDを入れてください。\n半角英数字がおすすめです（例: takahashi）');
    if (loginId === null) return '';
    if (!loginId) throw new Error('ログインIDが空です。');

    var pw = ask_('講師を追加',
      name + ' さんのパスワードを入れてください（8文字以上）。\n\n' +
      'この文字はシートには残りません。ハッシュだけが保存されます。');
    if (pw === null) return '';

    return createTeacher(name, loginId, pw) +
      '\n\nこのログインIDとパスワードで画面にログインできます。';
  });
}

function menuChangePassword() {
  run_('パスワードを変更', function () {
    var loginId = ask_('パスワードを変更', 'どの講師ですか。ログインIDを入れてください。');
    if (loginId === null) return '';
    if (!loginId) throw new Error('ログインIDが空です。');

    var pw = ask_('パスワードを変更', '新しいパスワードを入れてください（8文字以上）。');
    if (pw === null) return '';

    return changePassword(loginId, pw);
  });
}

function menuCleanupSessions() {
  run_('古いセッションを掃除', function () { return cleanupSessions(); });
}

function menuLogoutAll() {
  run_('全員をログアウトさせる', function () {
    var res = ui_().alert('全員をログアウトさせる',
      'ログイン中の端末をすべて切ります。\n次に使うときは、もう一度ログインが要ります。\n\n実行しますか。',
      ui_().ButtonSet.YES_NO);
    if (res !== ui_().Button.YES) return '';

    var rows = readAll_(SHEETS.SESSION);
    for (var i = rows.length - 1; i >= 0; i--) deleteRow_(SHEETS.SESSION, rows[i]._row);
    return rows.length + ' 件のセッションを切りました。';
  });
}

function menuNormalize() {
  run_('手で足した行を整える', function () { return normalizeSheets(); });
}

/* ---------- 手で足した行の面倒を見る ---------- */

/**
 * 生徒や教材は、シートに名前を打つだけで足せるようにしたい。
 * ただし id が無いと記録と結びつかないので、空欄を埋めて回る。
 *
 * - id が空 → 振る
 * - 有効 が空 → true にする（打ったばかりの行を「無効」扱いにしない）
 * - 作成日時 が空 → 今
 * - 単元・見出しの 教材id が空 → 教材略称から引く（略称だけ打てば済むように）
 * - 逆に 教材略称 が空 → 教材id から埋める
 */
function normalizeSheets() {
  return withLock_(function () {
    var lines = [];

    /* 講師も 有効 を立てる。パスワードが無ければどのみちログインできないので、
       ここで止める意味はない。止めるなら 有効 を手で FALSE にする */
    lines.push(fillBasic_(SHEETS.TEACHER, 't', true));
    lines.push(fillBasic_(SHEETS.STUDENT, 's', true));
    lines.push(fillBasic_(SHEETS.BOOK,    'b', true));
    lines.push(fillChild_(SHEETS.UNIT,    'u'));
    lines.push(fillChild_(SHEETS.SECTION, 'h'));

    var warn = warnings_();
    var body = lines.filter(function (s) { return s; }).join('\n');
    if (!body) body = '直すところはありませんでした。';
    return body + (warn ? '\n\n― 気になる点 ―\n' + warn : '');
  });
}

/** id・有効・作成日時 を埋める */
function fillBasic_(name, prefix, defaultActive) {
  var rows = readAll_(name);
  var n = 0;

  rows.forEach(function (r) {
    var patch = {};
    if (!String(r['id'] || '').trim()) patch['id'] = newId_(prefix);
    if (r['有効'] === '' || r['有効'] === null || r['有効'] === undefined) {
      if (defaultActive) patch['有効'] = true;
    }
    if (r['作成日時'] === '' || r['作成日時'] === null || r['作成日時'] === undefined) {
      patch['作成日時'] = new Date();
    }
    if (Object.keys(patch).length) { updateRow_(name, r._row, patch); n++; }
  });

  return n ? name + 'シート: ' + n + ' 行を整えました。' : '';
}

/** 教材にぶら下がる行（単元・見出し）。教材略称 ⇄ 教材id を補い合う */
function fillChild_(name, prefix) {
  var books = readAll_(SHEETS.BOOK);
  var byShort = {}, byId = {};
  books.forEach(function (b) {
    var short = String(b['略称'] || b['書名'] || '').trim();
    if (short) byShort[short] = b;
    if (b['id']) byId[String(b['id']).trim()] = b;
  });

  var rows = readAll_(name);
  var n = 0;

  rows.forEach(function (r) {
    var patch = {};
    if (!String(r['id'] || '').trim()) patch['id'] = newId_(prefix);

    var bookId = String(r['教材id'] || '').trim();
    var short  = String(r['教材略称'] || '').trim();

    if (!bookId && short && byShort[short]) {
      patch['教材id'] = byShort[short]['id'];
    } else if (bookId && !short && byId[bookId]) {
      patch['教材略称'] = byId[bookId]['略称'] || byId[bookId]['書名'];
    }

    if (Object.keys(patch).length) { updateRow_(name, r._row, patch); n++; }
  });

  return n ? name + 'シート: ' + n + ' 行を整えました。' : '';
}

/** 直せないけれど知らせておきたいこと */
function warnings_() {
  var out = [];

  var noPw = readAll_(SHEETS.TEACHER).filter(function (t) {
    return !String(t['パスワードハッシュ'] || '').trim();
  });
  if (noPw.length) {
    out.push('・パスワードが設定されていない講師が ' + noPw.length + ' 人います（' +
      noPw.map(function (t) { return t['名前'] || t['ログインID'] || '無名'; }).join('、') +
      '）。\n  「パスワードを変更…」で設定するまでログインできません。');
  }

  var bookIds = {};
  readAll_(SHEETS.BOOK).forEach(function (b) { if (b['id']) bookIds[String(b['id']).trim()] = true; });

  [SHEETS.UNIT, SHEETS.SECTION].forEach(function (name) {
    var orphan = readAll_(name).filter(function (r) {
      var id = String(r['教材id'] || '').trim();
      return !id || !bookIds[id];
    });
    if (orphan.length) {
      out.push('・' + name + 'シートに、どの教材か分からない行が ' + orphan.length +
        ' 行あります。\n  教材略称の綴りが教材シートと一致しているか確かめてください。');
    }
  });

  var dup = {}, dupList = [];
  readAll_(SHEETS.TEACHER).forEach(function (t) {
    var lid = String(t['ログインID'] || '').trim();
    if (!lid) return;
    if (dup[lid]) { if (dupList.indexOf(lid) < 0) dupList.push(lid); }
    dup[lid] = true;
  });
  if (dupList.length) {
    out.push('・ログインIDが重複しています: ' + dupList.join('、') +
      '\n  先に見つかった方だけがログインできます。どちらかを直してください。');
  }

  return out.join('\n');
}
