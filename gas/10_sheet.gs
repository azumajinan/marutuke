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

/**
 * 書こうとした列がまだ無ければ、右端に足してから書く。
 *
 * 列を増やしたのに initialize を流し忘れる、という事故を防ぐためのもの。
 * 黙って書き捨てると、あとから「保存したはずの値が無い」という形で表に出る。
 * 足すのは HEADERS に定義がある列だけ。知らない名前は無視する。
 */
function ensureColumns_(sh, obj) {
  var map = headerMap_(sh);
  var known = HEADERS[sh.getName()] || [];
  var missing = [];
  for (var k in obj) {
    if (map[k] === undefined && known.indexOf(k) >= 0 && missing.indexOf(k) < 0) missing.push(k);
  }
  if (!missing.length) return map;

  var at = sh.getLastColumn() + 1;
  sh.getRange(1, at, 1, missing.length).setValues([missing])
    .setFontWeight('bold')
    .setBackground('#EEF0F5');
  return headerMap_(sh);
}

/** { 列名: 値 } を見出しの順に並べて末尾に追加する */
function appendRow_(name, obj) {
  var sh = sheet_(name);
  var map = ensureColumns_(sh, obj);
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
  var map = ensureColumns_(sh, obj);
  for (var k in obj) {
    if (map[k] === undefined) continue;
    sh.getRange(rowNumber, map[k] + 1).setValue(obj[k]);
  }
}

function deleteRow_(name, rowNumber) {
  sheet_(name).deleteRow(rowNumber);
}

/**
 * シート全体を1回読んで、1回で書き戻す。
 *
 * updateRow_ は1セルごとに通信するので、行数が増えると実行時間の上限に当たる。
 * 生徒200行を1行ずつ直すと往復が800回になり、6分では終わらない。
 * ここはシート全体を配列で受け取り、直した列だけまとめて書き戻す。
 *
 * patcher(行オブジェクト) が { 列名: 値 } を返せばその行を直す。null なら触らない。
 * 戻り値は直した行数。
 */
function bulkFix_(name, patcher) {
  var sh = sheet_(name);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return 0;

  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var idx = {};
  for (var c = 0; c < values[0].length; c++) {
    var h = String(values[0][c]).trim();
    if (h) idx[h] = c;
  }

  var changed = 0;
  var touched = {};                       // 書き戻す列だけ覚える

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var blank = true;
    for (var i = 0; i < row.length; i++) {
      if (row[i] !== '' && row[i] !== null) { blank = false; break; }
    }
    if (blank) continue;

    var obj = { _row: r + 1 };
    for (var k in idx) obj[k] = row[idx[k]];

    var patch = patcher(obj);
    if (!patch) continue;

    var did = false;
    for (var key in patch) {
      if (idx[key] === undefined) continue;
      row[idx[key]] = patch[key];
      touched[idx[key]] = true;
      did = true;
    }
    if (did) changed++;
  }

  /* 触った列だけ書き戻す。シート全体を上書きすると、
     こちらが関知しない列の書式や式まで巻き込むため */
  for (var col in touched) {
    var c1 = Number(col);
    var colVals = [];
    for (var rr = 1; rr < values.length; rr++) colVals.push([values[rr][c1]]);
    sh.getRange(2, c1 + 1, colVals.length, 1).setValues(colVals);
  }
  return changed;
}

/** HEADERS にあって実際のシートに無い列を、右端にまとめて足す */
function ensureHeaders_(name) {
  var sh = sheet_(name);
  var head = HEADERS[name] || [];
  if (!head.length) return [];
  var map = headerMap_(sh);
  var missing = head.filter(function (h) { return map[h] === undefined; });
  if (!missing.length) return [];

  var at = sh.getLastColumn() + 1;
  sh.getRange(1, at, 1, missing.length).setValues([missing])
    .setFontWeight('bold')
    .setBackground('#EEF0F5');
  return missing;
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
