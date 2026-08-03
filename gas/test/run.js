/* GAS の外側だけを偽物に差し替えて、ロジックを実際に動かす */
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

/* ---- 偽スプレッドシート ---- */
const store = {};                       // name -> 2次元配列（1行目が見出し）
function mkSheet(name) {
  const rows = store[name] || (store[name] = []);
  return {
    getName: () => name,
    getLastRow: () => rows.length,
    getLastColumn: () => (rows.length ? Math.max(...rows.map(r => r.length)) : 0),
    getRange(r, c, nr = 1, nc = 1) {
      return {
        getValues() {
          const out = [];
          for (let i = 0; i < nr; i++) {
            const row = rows[r - 1 + i] || [];
            const line = [];
            for (let j = 0; j < nc; j++) line.push(row[c - 1 + j] === undefined ? '' : row[c - 1 + j]);
            out.push(line);
          }
          return out;
        },
        setValues(v) {
          for (let i = 0; i < v.length; i++) {
            const target = rows[r - 1 + i] || (rows[r - 1 + i] = []);
            for (let j = 0; j < v[i].length; j++) target[c - 1 + j] = v[i][j];
          }
          return this;
        },
        setValue(v) { return this.setValues([[v]]); },
        setFontWeight() { return this; },
        setBackground() { return this; }
      };
    },
    appendRow(row) { rows.push(row.slice()); },
    deleteRow(n) { rows.splice(n - 1, 1); },
    setFrozenRows() {}, autoResizeColumns() {}
  };
}
/* メニューが押した順に記録される。中身の確認用 */
const uiLog = [];
global.SpreadsheetApp = {
  getUi: () => ({
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' },
    Button: { OK: 'OK', YES: 'YES', NO: 'NO' },
    createMenu(name) {
      const m = { _items: [] };
      m.addItem = (label, fn) => { m._items.push([label, fn]); return m; };
      m.addSeparator = () => m;
      m.addToUi = () => { uiLog.push({ menu: name, items: m._items }); };
      return m;
    },
    alert: (...a) => { uiLog.push({ alert: a }); return 'OK'; },
    prompt: (...a) => { uiLog.push({ prompt: a }); return { getSelectedButton: () => 'OK', getResponseText: () => '' }; }
  }),
  getActiveSpreadsheet: () => ({
    getSheetByName: n => (store[n] ? mkSheet(n) : null),
    insertSheet: n => { store[n] = []; return mkSheet(n); },
    getSheets: () => Object.keys(store).map(mkSheet),
    deleteSheet: sh => { delete store[sh.getName()]; }
  })
};

global.Utilities = {
  DigestAlgorithm: { SHA_256: 'SHA_256' },
  Charset: { UTF_8: 'UTF_8' },
  computeDigest(_alg, str) {
    const buf = crypto.createHash('sha256').update(str, 'utf8').digest();
    return Array.from(buf).map(b => (b > 127 ? b - 256 : b));
  }
};

const cache = {};
global.CacheService = {
  getScriptCache: () => ({
    get: k => (cache[k] === undefined ? null : cache[k]),
    put: (k, v) => { cache[k] = v; },
    remove: k => { delete cache[k]; }
  })
};
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
global.Logger = { log: () => {} };
global.ContentService = {
  MimeType: { JSON: 'JSON', TEXT: 'TEXT' },
  createTextOutput(t) { return { _t: t, setMimeType() { return this; }, getContent() { return this._t; } }; }
};

/* ---- 本体を読み込む ---- */
// 間接 eval はグローバルスコープで走るので、var と function 宣言が globalThis に載る
for (const f of ['00_config', '10_sheet', '20_setup', '30_auth', '40_data', '50_api', '60_menu']) {
  (0, eval)(fs.readFileSync(path.join(__dirname, '..', f + '.gs'), 'utf8'));
}

/* ---- テスト ---- */
let pass = 0, fail = 0;
function t(name, fn) {
  try { fn(); console.log('  OK   ' + name); pass++; }
  catch (e) { console.log('  FAIL ' + name + '\n       ' + e.message); fail++; }
}
function eq(a, b, m) {
  const A = JSON.stringify(a), B = JSON.stringify(b);
  if (A !== B) throw new Error((m || '') + ' 期待 ' + B + ' / 実際 ' + A);
}
function throws(fn, re, m) {
  try { fn(); } catch (e) { if (!re || re.test(e.message)) return; throw new Error((m||'')+' 文言が違う: '+e.message); }
  throw new Error((m || '') + ' 例外が出なかった');
}

console.log('\n■ 初期化');
t('シートが7枚できる', () => { initialize(); eq(Object.keys(store).length, 7); });
t('二度実行しても壊れない', () => { initialize(); eq(Object.keys(store).length, 7); });

console.log('\n■ 認証');
t('講師を作れる', () => { createTeacher('高橋', 'takahashi', 'sup3r-secret-pw'); eq(readAll_(SHEETS.TEACHER).length, 1); });
t('パスワードは平文で保存されない', () => {
  const row = readAll_(SHEETS.TEACHER)[0];
  if (JSON.stringify(row).includes('sup3r-secret-pw')) throw new Error('平文が残っている');
});
t('同じログインIDは拒否', () => throws(() => createTeacher('別人', 'takahashi', 'another-pass1')));
t('8文字未満は拒否',   () => throws(() => createTeacher('短い', 'mijikai', 'abc'), /8文字/));
t('正しいパスワードでログインできる', () => {
  const r = login_('takahashi', 'sup3r-secret-pw');
  if (!r.token || r.token.length !== 64) throw new Error('トークンが不正: ' + r.token);
  eq(r.teacher.name, '高橋');
  globalThis.TOKEN = r.token;
});
t('誤ったパスワードは弾く', () => throws(() => login_('takahashi', 'wrong-password'), /違います/));
t('存在しないIDも同じ文言', () => throws(() => login_('nobody', 'whatever12'), /違います/));
t('5回失敗でロックされる', () => {
  for (let i = 0; i < 6; i++) { try { login_('lockme', 'bad'); } catch (e) {} }
  throws(() => login_('lockme', 'bad'), /ロック/);
});
t('トークン無しは弾く',   () => throws(() => requireTeacher_(''), /ログイン/));
t('偽トークンは弾く',     () => throws(() => requireTeacher_('0'.repeat(64)), /有効期限|ログイン/));
t('正しいトークンは通る', () => eq(requireTeacher_(TOKEN).name, '高橋'));

console.log('\n■ API の入口が認証を強制しているか');
const OPEN = ['ping', 'login'];
t('login/ping 以外は全部トークンを要求する', () => {
  const actions = ['logout','me','bootstrap','masters','records','addRecord',
                   'deleteRecord','setCause','addSection','addStudent','addBook'];
  actions.forEach(a => {
    const res = JSON.parse(doPost({ postData: { contents: JSON.stringify({ action: a }) } }).getContent());
    if (res.ok !== false) throw new Error(a + ' がトークン無しで通った');
  });
});
t('ping は認証不要', () => {
  const res = JSON.parse(doPost({ postData: { contents: JSON.stringify({ action: 'ping' }) } }).getContent());
  eq(res.ok, true);
});
t('壊れた本文でも落ちない', () => {
  const res = JSON.parse(doPost({ postData: { contents: '{{{' } }).getContent());
  eq(res.ok, false);
});
t('未知の操作は拒否', () => {
  const res = JSON.parse(doPost({ postData: { contents: JSON.stringify({ action: 'dropTable', token: TOKEN }) } }).getContent());
  eq(res.ok, false);
});

console.log('\n■ マスタと記録');
t('サンプル投入', () => { seedSampleData(); const m = getMasters_(); eq(m.students.length, 2); eq(m.books.length, 2); });
t('単元がページ順', () => {
  const b = getMasters_().books[0];
  eq(b.units.map(u => u.from), [8,28,46,70,96]);
});
t('見出しは教材ごと', () => {
  const bs = getMasters_().books;
  eq(bs[0].sections, ['基本問題','練習問題','章末問題']);
  eq(bs[1].sections, ['確認問題','練習問題','章末問題']);
});

const T = requireTeacher_(TOKEN);
const M = getMasters_();
const S = M.students[0].id, B = M.books[0].id;

t('記録を追加できる', () => {
  const r = addRecord_(T, { student: S, book: B, page: 78, section: '', major: 1, q: 3, result: 2, cause: 3 });
  if (!r.id) throw new Error('idが無い');
  eq(r.result, 2); eq(r.cause, 3);
});
t('シートには日本語で入る', () => {
  const row = readAll_(SHEETS.RECORD)[0];
  eq(row['結果'], '×'); eq(row['つまずき'], '知識なし');
  eq(row['生徒名'], '田中 涼');
});
t('○ につまずきは付かない', () => {
  const r = addRecord_(T, { student: S, book: B, page: 78, section: '', major: 1, q: 4, result: 1, cause: 3 });
  eq(r.cause, 0);
});
t('不正な結果は拒否', () => throws(() => addRecord_(T, { student: S, book: B, result: 9 }), /不正/));
t('存在しない生徒は拒否', () => throws(() => addRecord_(T, { student: 'nope', book: B, result: 1 }), /見つかりません/));
t('読み戻せる', () => {
  const rs = getRecords_({ days: 0 });
  eq(rs.length, 2);
  eq(rs[0].result, 2); eq(rs[0].cause, 3); eq(rs[0].page, 78);
});
t('生徒で絞れる', () => {
  eq(getRecords_({ studentId: S, days: 0 }).length, 2);
  eq(getRecords_({ studentId: M.students[1].id, days: 0 }).length, 0);
});
t('つまずきを変えられる', () => {
  const id = getRecords_({ days: 0 })[0].id;
  setCause_(T, id, 1);
  eq(getRecords_({ days: 0 })[0].cause, 1);
});
t('正解の記録には分類を付けられない', () => {
  const ok = getRecords_({ days: 0 }).filter(r => r.result === 1)[0];
  throws(() => setCause_(T, ok.id, 2), /不正解/);
});
t('記録を消せる', () => {
  const id = getRecords_({ days: 0 })[0].id;
  deleteRecord_(T, id);
  eq(getRecords_({ days: 0 }).length, 1);
});
t('無い記録の削除はエラーにしない', () => eq(deleteRecord_(T, 'r_nope'), true));

console.log('\n■ 応答が落ちたときの再送（受付キー）');
t('同じ受付キーは二重に入らない', () => {
  const n0 = getRecords_({ days: 0 }).length;
  const p = { key: 'k-abc', student: S, book: B, page: 90, section: '', major: 3, q: 1, result: 2, cause: 2 };
  const a = addRecord_(T, p);
  const b = addRecord_(T, p);                    // 画面が再送した想定
  eq(getRecords_({ days: 0 }).length, n0 + 1);
  eq(b.id, a.id);
  eq(b.duplicate, true);
  eq(b.result, 2); eq(b.cause, 2); eq(b.page, 90);
});
t('受付キーが無ければ従来どおり毎回入る', () => {
  const n0 = getRecords_({ days: 0 }).length;
  const p = { student: S, book: B, page: 91, section: '', major: 1, q: 1, result: 1 };
  addRecord_(T, p); addRecord_(T, p);
  eq(getRecords_({ days: 0 }).length, n0 + 2);
});
t('受付キーから記録を引ける', () => {
  const f = findRecordByKey_('k-abc');
  if (!f || !f.id) throw new Error('引けない');
  eq(f.id, getRecords_({ days: 0 }).filter(r => r.page === 90)[0].id);
});
t('無いキーは null', () => eq(findRecordByKey_('k-nope'), null));
t('findRecord も認証が要る', () => {
  const res = JSON.parse(doPost({ postData: { contents: JSON.stringify({ action: 'findRecord', key: 'k-abc' }) } }).getContent());
  eq(res.ok, false);
});

console.log('\n■ 列を増やしたときの移行');
t('既にあるシートに足りない列を付け足す', () => {
  const sh = store[SHEETS.RECORD];
  const col = sh[0].indexOf('受付キー');
  if (col < 0) throw new Error('前提が崩れている');
  sh.forEach(r => r.splice(col, 1));            // 古いシートを再現（受付キー列が無い）
  eq(store[SHEETS.RECORD][0].indexOf('受付キー'), -1);
  initialize();
  const after = store[SHEETS.RECORD][0].indexOf('受付キー');
  if (after < 0) throw new Error('列が足されていない');
  eq(after, store[SHEETS.RECORD][0].length - 1);   // 右端に足す
});
t('列を足しても既存の行は読める', () => {
  const rs = getRecords_({ days: 0 });
  if (!rs.length) throw new Error('読めなくなった');
  eq(rs.filter(r => r.page === 90).length, 1);
});

console.log('\n■ 見出しの追加');
t('追加できる', () => { addSection_(T, B, '入試対策'); eq(getMasters_().books[0].sections.includes('入試対策'), true); });
t('重複は増えない', () => {
  addSection_(T, B, '入試対策');
  eq(getMasters_().books[0].sections.filter(s => s === '入試対策').length, 1);
});
t('空文字は拒否', () => throws(() => addSection_(T, B, '   '), /空/));

console.log('\n■ 手でシートを触ったときの耐性');
t('列を入れ替えても読める', () => {
  const rows = store[SHEETS.STUDENT];
  const order = [2, 0, 1, 3, 4];                       // 名前, id, 学年, 有効, 作成日時
  store[SHEETS.STUDENT] = rows.map(r => order.map(i => r[i]));
  // 見出しも一緒に動くので readAll_ は列名で追随できるはず
  eq(getMasters_().students.length, 2);
  eq(getMasters_().students[0].name, '田中 涼');
});
t('空行があっても読み飛ばす', () => {
  store[SHEETS.STUDENT].push(['', '', '', '', '']);
  eq(getMasters_().students.length, 2);
});
t('結果が壊れた行は捨てる', () => {
  const before = getRecords_({ days: 0 }).length;
  const sh = store[SHEETS.RECORD];
  const head = sh[0];
  const bad = head.map(() => '');
  bad[head.indexOf('id')] = 'r_broken';
  bad[head.indexOf('日時')] = new Date();
  bad[head.indexOf('結果')] = 'あ';
  sh.push(bad);
  eq(getRecords_({ days: 0 }).length, before);        // 壊れた行は数に入らない
});
t('無効な生徒は一覧から消える', () => {
  const sh = store[SHEETS.STUDENT];
  const col = sh[0].indexOf('有効');
  sh[1][col] = false;
  eq(getMasters_().students.length, 1);
});

console.log('\n■ シートから手で足した行の面倒');
t('メニューが組み立てられる', () => {
  uiLog.length = 0;
  onOpen();
  eq(uiLog[0].menu, 'まるつけ');
  /* メニューから呼ぶ関数は _ 付きだと GAS が呼べない */
  uiLog[0].items.forEach(it => {
    if (/_$/.test(it[1])) throw new Error(it[0] + ' の呼び先が _ 付き: ' + it[1]);
    if (typeof globalThis[it[1]] !== 'function') throw new Error(it[1] + ' が無い');
  });
});
t('名前だけ打った生徒に id が振られ、有効になる', () => {
  const sh = store[SHEETS.STUDENT], head = sh[0];
  const row = head.map(() => '');
  row[head.indexOf('名前')] = '手打ち 太郎';
  sh.push(row);
  normalizeSheets();
  const added = readAll_(SHEETS.STUDENT).filter(r => r['名前'] === '手打ち 太郎')[0];
  if (!/^s_/.test(added['id'])) throw new Error('id が振られていない: ' + added['id']);
  eq(toBool_(added['有効']), true);
  if (!added['作成日時']) throw new Error('作成日時が空');
});
t('生徒一覧にすぐ出る', () => {
  eq(getMasters_().students.filter(s => s.name === '手打ち 太郎').length, 1);
});
t('単元は教材略称だけ打てば教材idが埋まる', () => {
  const short = readAll_(SHEETS.BOOK)[0]['略称'];
  const sh = store[SHEETS.UNIT], head = sh[0];
  const row = head.map(() => '');
  row[head.indexOf('教材略称')] = short;
  row[head.indexOf('単元名')] = '手打ち単元';
  row[head.indexOf('開始ページ')] = 122;
  row[head.indexOf('終了ページ')] = 140;
  sh.push(row);
  normalizeSheets();
  const u = readAll_(SHEETS.UNIT).filter(r => r['単元名'] === '手打ち単元')[0];
  eq(u['教材id'], readAll_(SHEETS.BOOK)[0]['id']);
  if (!/^u_/.test(u['id'])) throw new Error('id が振られていない');
});
t('二度実行しても id は振り直されない', () => {
  const before = readAll_(SHEETS.STUDENT).map(r => r['id']).join(',');
  normalizeSheets();
  eq(readAll_(SHEETS.STUDENT).map(r => r['id']).join(','), before);
});
t('パスワード未設定の講師を知らせる', () => {
  const sh = store[SHEETS.TEACHER], head = sh[0];
  const row = head.map(() => '');
  row[head.indexOf('名前')] = '新人 花子';
  row[head.indexOf('ログインID')] = 'hanako';
  sh.push(row);
  const msg = normalizeSheets();
  if (!/新人 花子/.test(msg)) throw new Error('警告に出ない: ' + msg);
  if (!/ログインできません/.test(msg)) throw new Error('文言が違う: ' + msg);
});
t('パスワード未設定ではログインできない', () => throws(() => login_('hanako', 'whatever12'), /違います/));
t('メニューから設定すればログインできる', () => {
  changePassword('hanako', 'hanako-no-pass');
  eq(requireTeacher_(login_('hanako', 'hanako-no-pass').token).name, '新人 花子');
});
t('教材に結びつかない単元を知らせる', () => {
  const sh = store[SHEETS.UNIT], head = sh[0];
  const row = head.map(() => '');
  row[head.indexOf('教材略称')] = '存在しない教材';
  row[head.indexOf('単元名')] = '迷子';
  sh.push(row);
  const msg = normalizeSheets();
  if (!/どの教材か分からない/.test(msg)) throw new Error('警告に出ない: ' + msg);
});
t('ログインIDの重複を知らせる', () => {
  const sh = store[SHEETS.TEACHER], head = sh[0];
  const row = head.map(() => '');
  row[head.indexOf('id')] = 't_dup';
  row[head.indexOf('名前')] = '別人';
  row[head.indexOf('ログインID')] = 'hanako';
  sh.push(row);
  const msg = normalizeSheets();
  if (!/重複/.test(msg)) throw new Error('警告に出ない: ' + msg);
});

console.log('\n■ ログアウト');
t('ログアウトでトークンが無効になる', () => {
  logout_(TOKEN);
  throws(() => requireTeacher_(TOKEN), /有効期限|ログイン/);
});
t('パスワード変更で全セッションが切れる', () => {
  const a = login_('takahashi', 'sup3r-secret-pw');
  changePassword('takahashi', 'brand-new-password');
  throws(() => requireTeacher_(a.token), /有効期限|ログイン/);
  const b = login_('takahashi', 'brand-new-password');
  eq(requireTeacher_(b.token).name, '高橋');
});

console.log('\n' + (fail ? '✗ ' + fail + ' 件失敗 / ' : '') + pass + ' 件成功');
process.exit(fail ? 1 : 0);
