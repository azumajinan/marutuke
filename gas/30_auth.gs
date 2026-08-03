/**
 * まるつけ — 認証
 *
 * エンドポイントは匿名公開になるため、防御はここに一本化される。
 * データを読み書きする操作は例外なくトークンを要求すること。
 */

/**
 * 乱数の出どころ。
 *
 * Math.random() は暗号用途ではない。出力をいくつか見れば内部状態を復元でき、
 * 続きが予測できる。Utilities.getUuid() は Java の SecureRandom を使うので、
 * こちらを種にしてハッシュに通す。
 */
function randomHex_() {
  var seed = Utilities.getUuid() + '|' + Utilities.getUuid() + '|' + Date.now();
  return bytesToHex_(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed, Utilities.Charset.UTF_8)
  );
}

/** ソルトは秘密ではなく重複しなければよいが、乱数の出どころは揃えておく */
function newSalt_() {
  return randomHex_().slice(0, 32);
}

function bytesToHex_(bytes) {
  var s = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] + 256) % 256;
    s += (b < 16 ? '0' : '') + b.toString(16);
  }
  return s;
}

/**
 * ソルト付きで SHA-256 を繰り返す。
 * GAS には bcrypt が無いので、反復回数で総当たりを遅くする。
 *
 * rounds は必ず「そのハッシュを作ったときの回数」を渡すこと。
 * 省略すると今の既定値になる。照合で取り違えると誰もログインできなくなる。
 */
function hashPassword_(password, salt, rounds) {
  var n = toInt_(rounds, 0) || AUTH.HASH_ROUNDS;
  var cur = salt + '|' + password;
  for (var i = 0; i < n; i++) {
    cur = bytesToHex_(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cur, Utilities.Charset.UTF_8)
    );
  }
  return cur;
}

/** 比較にかかる時間を入力に依存させない */
function safeEquals_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

/* ---------- 総当たりの抑制 ---------- */

function failureKey_(loginId) {
  return 'fail_' + String(loginId).toLowerCase();
}

function isLockedOut_(loginId) {
  var n = CacheService.getScriptCache().get(failureKey_(loginId));
  return n !== null && toInt_(n, 0) >= AUTH.MAX_FAILURES;
}

/** 同じ入力の数え直しを避ける。再送は人の試行ではない */
function alreadyCounted_(key) {
  if (!key) return false;
  var k = 'try:' + String(key);
  var c = CacheService.getScriptCache();
  if (c.get(k)) return true;
  c.put(k, '1', 300);
  return false;
}

function noteFailure_(loginId) {
  var cache = CacheService.getScriptCache();
  var key = failureKey_(loginId);
  var n = toInt_(cache.get(key), 0) + 1;
  cache.put(key, String(n), AUTH.LOCKOUT_MINUTES * 60);
}

function clearFailures_(loginId) {
  CacheService.getScriptCache().remove(failureKey_(loginId));
}

/* ---------- ログイン ---------- */

/**
 * @param key 画面が1回の入力ごとに作る文字列。
 *   応答が落ちて再送されたとき、失敗を二重に数えないために使う。
 *   数えるのは「人が入力し直した回数」であって、通信のやり直しではない。
 */
function login_(loginId, password, key) {
  if (!loginId || !password) throw new Error('ログインIDとパスワードを入力してください。');

  if (isLockedOut_(loginId)) {
    throw new Error('ログインの失敗が続いたため、' + AUTH.LOCKOUT_MINUTES + '分間ロックしています。');
  }

  var teachers = readAll_(SHEETS.TEACHER);
  var found = null;
  for (var i = 0; i < teachers.length; i++) {
    if (String(teachers[i]['ログインID']).trim() === String(loginId).trim()) { found = teachers[i]; break; }
  }

  // 見つからない場合も同じ文言を返す。IDの存在を教えない
  var FAIL = 'ログインIDまたはパスワードが違います。';
  if (!found || !toBool_(found['有効'])) {
    if (!alreadyCounted_(key)) noteFailure_(loginId);
    throw new Error(FAIL);
  }

  /* 「反復回数」列が無い、または空の行は、列を足す前に作られたもの */
  var rounds = toInt_(found['反復回数'], 0) || AUTH.LEGACY_ROUNDS;

  var hash = hashPassword_(password, found['ソルト'], rounds);
  if (!safeEquals_(hash, found['パスワードハッシュ'])) {
    if (!alreadyCounted_(key)) noteFailure_(loginId);
    throw new Error(FAIL);
  }

  clearFailures_(loginId);

  /* 合っていた。古い回数なら今の回数で入れ直す。本人しか知らないこの瞬間にしかできない */
  if (rounds !== AUTH.HASH_ROUNDS) {
    var salt2 = newSalt_();
    updateRow_(SHEETS.TEACHER, found._row, {
      'ソルト':          salt2,
      'パスワードハッシュ': hashPassword_(password, salt2, AUTH.HASH_ROUNDS),
      '反復回数':        AUTH.HASH_ROUNDS
    });
  }

  var token = newToken_();
  var now = new Date();
  withLock_(function () {
    appendRow_(SHEETS.SESSION, {
      'token':      token,
      '講師id':      found['id'],
      '発行日時':     now,
      '最終利用日時':  now
    });
  });

  return {
    token: token,
    teacher: { id: found['id'], name: found['名前'] }
  };
}

/**
 * セッショントークン。64桁の16進、256ビット。
 *
 * これはパスワードと同じ力を持つ（12時間、全データに触れる）。
 * 推測できてはいけないので、必ず randomHex_ を使うこと。
 */
function newToken_() {
  return randomHex_();
}

/**
 * トークンを検証して講師を返す。無効なら例外。
 * データを触るハンドラは必ず最初にこれを通すこと。
 */
function requireTeacher_(token) {
  if (!token) throw new Error('ログインしてください。');

  var sessions = readAll_(SHEETS.SESSION);
  var hit = null;
  for (var i = 0; i < sessions.length; i++) {
    if (safeEquals_(sessions[i]['token'], token)) { hit = sessions[i]; break; }
  }
  if (!hit) throw new Error('ログインの有効期限が切れました。もう一度ログインしてください。');

  var issued = new Date(hit['発行日時']).getTime();
  if (!issued || Date.now() - issued > AUTH.SESSION_HOURS * 3600 * 1000) {
    deleteRow_(SHEETS.SESSION, hit._row);
    throw new Error('ログインの有効期限が切れました。もう一度ログインしてください。');
  }

  var teachers = readAll_(SHEETS.TEACHER);
  for (var j = 0; j < teachers.length; j++) {
    if (teachers[j]['id'] === hit['講師id'] && toBool_(teachers[j]['有効'])) {
      // 最終利用の記録は毎回書くと重いので、5分以上空いたときだけ
      var last = new Date(hit['最終利用日時'] || hit['発行日時']).getTime();
      if (Date.now() - last > 5 * 60 * 1000) {
        updateRow_(SHEETS.SESSION, hit._row, { '最終利用日時': new Date() });
      }
      return { id: teachers[j]['id'], name: teachers[j]['名前'], _sessionRow: hit._row };
    }
  }
  throw new Error('このアカウントは使えなくなっています。');
}

function logout_(token) {
  var sessions = readAll_(SHEETS.SESSION);
  for (var i = 0; i < sessions.length; i++) {
    if (safeEquals_(sessions[i]['token'], token)) {
      deleteRow_(SHEETS.SESSION, sessions[i]._row);
      return true;
    }
  }
  return true;
}

function revokeSessionsOf_(teacherId) {
  var sessions = readAll_(SHEETS.SESSION);
  for (var i = sessions.length - 1; i >= 0; i--) {
    if (sessions[i]['講師id'] === teacherId) deleteRow_(SHEETS.SESSION, sessions[i]._row);
  }
}
