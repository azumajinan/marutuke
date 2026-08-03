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
    .addItem('アクセスURLを表示', 'menuShowUrl')
    .addItem('アクセスURLを作り直す', 'menuResetUrl')
    .addSeparator()
    .addItem('講師を追加…', 'menuAddTeacher')
    .addItem('手で足した行を整える', 'menuNormalize')
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
    return createTeacher(name) + '\n\n画面を開いたときに、この名前を選べるようになります。';
  });
}

/** 画面のある場所。アクセスURLを組み立てるのに使う */
var APP_BASE_URL = 'https://azumajinan.github.io/marutuke/';

function accessUrl_() {
  return APP_BASE_URL + '?k=' + ensureAccessKey_();
}

/**
 * URL を見せるダイアログ。
 *
 * alert() は本文を手で選んでコピーすることになる。URL が折り返すので
 * 鍵の途中までしか取れず、「鍵の無いURL」を開いて入れない事故が起きた。
 * ボタン1つでコピーできる形にして、選択操作そのものを無くす。
 */
function showUrlDialog_(title, lead) {
  var url = accessUrl_();
  var html =
    '<style>' +
    'body{font:13px/1.7 -apple-system,"Helvetica Neue",sans-serif;margin:0;padding:16px;color:#1b2330}' +
    'p{margin:0 0 12px}.lead{font-weight:600}' +
    'input{width:100%;box-sizing:border-box;font:12px/1.5 monospace;padding:8px;' +
    'border:1px solid #c6ccd8;border-radius:6px;background:#f7f8fb}' +
    '.row{display:flex;gap:8px;align-items:center;margin:10px 0 14px}' +
    'button{font:600 13px/1 inherit;padding:9px 16px;border:0;border-radius:6px;' +
    'background:#2f3e7e;color:#fff;cursor:pointer}' +
    '.ok{color:#1a7f4b;font-weight:600}' +
    '.warn{color:#8a3324;background:#fdf2ee;padding:10px 12px;border-radius:6px}' +
    'a{color:#2f3e7e}' +
    '</style>' +
    '<p class="lead">' + esc_(lead) + '</p>' +
    '<input id="u" type="text" readonly value="' + esc_(url) + '">' +
    '<div class="row">' +
    '<button onclick="cp()">コピー</button>' +
    '<a href="' + esc_(url) + '" target="_blank" rel="noopener">この端末で開く</a>' +
    '<span id="done" class="ok"></span>' +
    '</div>' +
    '<p class="warn">このURLを知っている人は誰でも記録を読み書きできます。' +
    'SNSや公開の場に貼らないでください。渡すのは講師だけに。</p>' +
    '<p>スマホで使うときは、コピーしたURLを自分宛のメッセージなどで送って開いてください。<br>' +
    '一度開けば端末が覚えるので、次からは鍵なしのURLでも入れます。</p>' +
    '<script>' +
    'function cp(){var e=document.getElementById("u");e.focus();e.setSelectionRange(0,e.value.length);' +
    'try{document.execCommand("copy");document.getElementById("done").textContent="コピーしました";}' +
    'catch(err){document.getElementById("done").textContent="Ctrl+C を押してください";}}' +
    '<\/script>';

  ui_().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(560).setHeight(360),
    title
  );
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function menuShowUrl() {
  run_('アクセスURL', function () {
    showUrlDialog_('アクセスURL', 'このURLを開けば、そのまま使えます。ログインはありません。');
    return '';
  });
}

function menuResetUrl() {
  run_('アクセスURLを作り直す', function () {
    var res = ui_().alert('アクセスURLを作り直す',
      '今のURLは使えなくなります。\n各講師の端末で、新しいURLを開き直す必要があります。\n\n' +
      '端末を紛失したときなどに使ってください。\n\n作り直しますか。',
      ui_().ButtonSet.YES_NO);
    if (res !== ui_().Button.YES) return '';

    resetAccessKey_();
    showUrlDialog_('新しいアクセスURL', '古いURLはもう使えません。各講師にこのURLを配り直してください。');
    return '';
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

function isBlank_(v) { return v === '' || v === null || v === undefined; }

/** id・有効・作成日時 を埋める */
function fillBasic_(name, prefix, defaultActive) {
  ensureHeaders_(name);
  var now = new Date();

  var n = bulkFix_(name, function (r) {
    var patch = {};
    if (!String(r['id'] || '').trim()) patch['id'] = newId_(prefix);
    if (isBlank_(r['有効']) && defaultActive) patch['有効'] = true;
    if (isBlank_(r['作成日時'])) patch['作成日時'] = now;
    return Object.keys(patch).length ? patch : null;
  });

  return n ? name + 'シート: ' + n + ' 行を整えました。' : '';
}

/** 教材にぶら下がる行（単元・見出し）。教材略称 ⇄ 教材id を補い合う */
function fillChild_(name, prefix) {
  ensureHeaders_(name);

  var byShort = {}, byId = {};
  readAll_(SHEETS.BOOK).forEach(function (b) {
    var short = String(b['略称'] || b['書名'] || '').trim();
    if (short) byShort[short] = b;
    if (b['id']) byId[String(b['id']).trim()] = b;
  });

  var n = bulkFix_(name, function (r) {
    var patch = {};
    if (!String(r['id'] || '').trim()) patch['id'] = newId_(prefix);

    var bookId = String(r['教材id'] || '').trim();
    var short  = String(r['教材略称'] || '').trim();

    if (!bookId && short && byShort[short]) {
      patch['教材id'] = byShort[short]['id'];
    } else if (bookId && !short && byId[bookId]) {
      patch['教材略称'] = byId[bookId]['略称'] || byId[bookId]['書名'];
    }
    return Object.keys(patch).length ? patch : null;
  });

  return n ? name + 'シート: ' + n + ' 行を整えました。' : '';
}

/** 直せないけれど知らせておきたいこと */
function warnings_() {
  var out = [];

  var noName = readAll_(SHEETS.TEACHER).filter(function (t) {
    return !String(t['名前'] || '').trim();
  });
  if (noName.length) {
    out.push('・名前の無い講師の行が ' + noName.length + ' 行あります。\n' +
      '  画面の講師選びに出てこないので、名前を入れるか行ごと消してください。');
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

  var seen = {}, dupList = [];
  readAll_(SHEETS.STUDENT).forEach(function (s) {
    var id = String(s['id'] || '').trim();
    if (!id) return;
    if (seen[id] && dupList.indexOf(id) < 0) dupList.push(id);
    seen[id] = true;
  });
  if (dupList.length) {
    out.push('・生徒のidが重複しています: ' + dupList.join('、') +
      '\n  記録がどちらの生徒のものか分からなくなります。片方を直してください。');
  }

  return out.join('\n');
}
