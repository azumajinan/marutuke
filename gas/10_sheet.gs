/**
 * まるつけ — シート入出力の共通処理
 *
 * 列は見出し名で引く。手で列を入れ替えても動く。
 */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('シートがありません: ' + name + '（初期化 initialize() を実行してください）');
  return sh;
}

/** 見出し行を { 列名: 0始まりの位置 } にする */
function headerMap_(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol === 0) throw new Error('見出し行がありません: ' + sh.getName());
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < head.length; i++) {
    var key = String(head[i]).trim();
    if (key) map[key] = i;
  }
  return map;
}

/**
 * シート全体を { 列名: 値 } の配列で返す。
 * 行番号が要る操作のために _row（1始まりの実際の行番号）を添える。
 */
function readAll_(name) {
  var sh = sheet_(name);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var map = headerMap_(sh);
  var values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var out = [];

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    // 全列が空の行は読み飛ばす（手で行を消したあとの残骸対策）
    var empty = true;
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== '' && row[c] !== null) { empty = false; break; }
    }
    if (empty) continue;

    var obj = { _row: r + 2 };
    for (var key in map) obj[key] = row[map[key]];
    out.push(obj);
  }
  return out;
}

/** { 列名: 値 } を見出しの順に並べて末尾に追加する */
function appendRow_(name, obj) {
  var sh = sheet_(name);
  var map = headerMap_(sh);
  var row = [];
  for (var key in map) row[map[key]] = '';
  for (var k in obj) {
    if (map[k] === undefined) continue;
    row[map[k]] = obj[k];
  }
  for (var i = 0; i < row.length; i++) if (row[i] === undefined) row[i] = '';
  sh.appendRow(row);
  return sh.getLastRow();
}

/** 指定行の一部の列だけ書き換える */
function updateRow_(name, rowNumber, obj) {
  var sh = sheet_(name);
  var map = headerMap_(sh);
  for (var k in obj) {
    if (map[k] === undefined) continue;
    sh.getRange(rowNumber, map[k] + 1).setValue(obj[k]);
  }
}

function deleteRow_(name, rowNumber) {
  sheet_(name).deleteRow(rowNumber);
}

/**
 * 衝突しないID。時刻を36進で並べ、末尾に乱数を足す。
 * 講師2人が同時に採点しても衝突しない程度で、シート上でも読める長さ。
 *
 * ここは Math.random でよい。IDは秘密ではなく、当てられても
 * トークンを持たない相手には何もできないため。
 * 秘密の値（トークン・ソルト）には使わないこと。randomHex_ を使う。
 */
function newId_(prefix) {
  var t = Date.now().toString(36);
  var r = Math.floor(Math.random() * 1679616).toString(36); // 36^4
  while (r.length < 4) r = '0' + r;
  return prefix + '_' + t + r;
}

/** 書き込みは必ずロックを取る。講師2人が同時に採点するため */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('混み合っています。少し待ってからもう一度お試しください。');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function toBool_(v) {
  if (v === true) return true;
  if (v === false || v === '' || v === null || v === undefined) return false;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'はい' || s === '1' || s === '有効' || s === 'yes';
}

function toInt_(v, fallback) {
  var n = parseInt(v, 10);
  return isNaN(n) ? (fallback === undefined ? 0 : fallback) : n;
}
