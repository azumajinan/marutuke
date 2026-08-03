/**
 * まるつけ — 認証
 *
 * エンドポイントは匿名公開になるため、防御はここに一本化される。
 * データを読み書きする操作は例外なくトークンを要求すること。
 */

function newSalt_() {
  var bytes = [];
  for (var i = 0; i < 16; i++) bytes.push(Math.floor(Math.random() * 256) - 128);
  return bytesToHex_(bytes);
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
 * ハッシュはスプレッドシート内にしか無く、シートの閲覧権を持つのは塾の関係者だけ、
 * という前提に立っている。パスワードは十分な長さにすること。
 */
function hashPassword_(password, salt) {
  var cur = salt + '|' + password;
  for (var i = 0; i < AUTH.HASH_ROUNDS; i++) {
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

function login_(loginId, password) {
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
    noteFailure_(loginId);
    throw new Error(FAIL);
  }

  var hash = hashPassword_(password, found['ソルト']);
  if (!safeEquals_(hash, found['パスワードハッシュ'])) {
    noteFailure_(loginId);
    throw new Error(FAIL);
  }

  clearFailures_(loginId);

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

function newToken_() {
  var bytes = [];
  for (var i = 0; i < 32; i++) bytes.push(Math.floor(Math.random() * 256) - 128);
  return bytesToHex_(bytes);
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
