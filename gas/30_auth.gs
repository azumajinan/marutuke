/**
 * まるつけ — アクセスの許可
 *
 * ログインは無い。「アクセスキーの入った URL を知っていること」が許可の条件。
 * 鍵は URL の ?k=... に入っていて、画面は初回に端末へ保存する。
 *
 * エンドポイントは匿名公開なので、防御はこの鍵だけに依存する。
 * 鍵は公開リポジトリには置かない。スクリプトのプロパティに持つ。
 *
 * 照合は文字列比較だけ。パスワードのようなハッシュの繰り返しは無く、一瞬で終わる。
 */

var ACCESS_KEY_PROP = 'ACCESS_KEY';

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

function bytesToHex_(bytes) {
  var s = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] + 256) % 256;
    s += (b < 16 ? '0' : '') + b.toString(16);
  }
  return s;
}

/** 比較にかかる時間を入力に依存させない */
function safeEquals_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

/* ---------- アクセスキー ---------- */

function getAccessKey_() {
  return String(PropertiesService.getScriptProperties().getProperty(ACCESS_KEY_PROP) || '');
}

/** 鍵を作り直す。前の URL はその場で使えなくなる */
function resetAccessKey_() {
  var key = randomHex_().slice(0, 32);
  PropertiesService.getScriptProperties().setProperty(ACCESS_KEY_PROP, key);
  return key;
}

/** 無ければ作る。初回の取り違えを防ぐため、勝手には作り直さない */
function ensureAccessKey_() {
  return getAccessKey_() || resetAccessKey_();
}

/**
 * データを触る操作は必ずこれを通す。
 *
 * 鍵が未設定のまま素通しにはしない。設定し忘れが「誰でも書ける状態」になるより、
 * 「誰も使えない状態」で気付く方がよい。
 */
function requireKey_(key) {
  var expected = getAccessKey_();
  if (!expected) {
    throw new Error('まだ使える状態になっていません。スプレッドシートの「まるつけ」メニューからアクセスURLを作ってください。');
  }
  if (!key || !safeEquals_(key, expected)) {
    throw new Error('このURLでは使えません。正しいURLを管理者から受け取ってください。');
  }
  return true;
}
