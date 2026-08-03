/**
 * まるつけ — データ取得と書き込み
 *
 * SPEC.md の導出ルールに従う。状態を持つ列は作らず、すべて記録から導く。
 */

/* ---------- マスタ ---------- */

/**
 * 送られてきた講師idを名簿に照らす。
 * 見つからなければ空。記録は残す（誰が付けたか分からない記録の方がまし）。
 */
function teacherOf_(id) {
  id = String(id || '').trim();
  if (!id) return { id: '', name: '' };
  var rows = readAll_(SHEETS.TEACHER);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['id']).trim() === id) {
      return { id: rows[i]['id'], name: rows[i]['名前'] || '' };
    }
  }
  return { id: '', name: '' };
}

function getMasters_() {
  var teachers = readAll_(SHEETS.TEACHER)
    .filter(function (r) { return toBool_(r['有効']); })
    .map(function (r) { return { id: r['id'], name: r['名前'] }; });

  var students = readAll_(SHEETS.STUDENT)
    .filter(function (r) { return toBool_(r['有効']); })
    .map(function (r) {
      return { id: r['id'], name: r['名前'], grade: r['学年'] || '' };
    });

  var units = readAll_(SHEETS.UNIT);
  var sections = readAll_(SHEETS.SECTION);

  var books = readAll_(SHEETS.BOOK)
    .filter(function (r) { return toBool_(r['有効']); })
    .map(function (b) {
      return {
        id: b['id'],
        title: b['書名'],
        short: b['略称'] || b['書名'],
        subject: b['科目'] || '',
        pages: toInt_(b['総ページ'], 0),
        units: units
          .filter(function (u) { return u['教材id'] === b['id']; })
          .map(function (u) {
            return {
              id: u['id'],
              name: u['単元名'],
              from: toInt_(u['開始ページ'], 0),
              to: toInt_(u['終了ページ'], 0)
            };
          })
          .sort(function (x, y) { return x.from - y.from; }),
        sections: sections
          .filter(function (s) { return s['教材id'] === b['id']; })
          .map(function (s) { return s['見出し']; })
      };
    });

  return { teachers: teachers, students: students, books: books };
}

/* ---------- 記録 ---------- */

/**
 * 記録を返す。既定は直近 DEFAULT_RECORD_DAYS 日。
 * 全期間が要るとき（分析の推移など）は days に 0 を渡す。
 */
function getRecords_(opt) {
  opt = opt || {};
  var days = opt.days === undefined ? DEFAULT_RECORD_DAYS : toInt_(opt.days, DEFAULT_RECORD_DAYS);
  var since = days > 0 ? Date.now() - days * 24 * 3600 * 1000 : 0;

  var rows = readAll_(SHEETS.RECORD);
  var out = [];

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (opt.studentId && r['生徒id'] !== opt.studentId) continue;

    var at = new Date(r['日時']).getTime();
    if (!at || (since && at < since)) continue;

    var result = TEXT_TO_RESULT[String(r['結果']).trim()];
    if (!result) continue; // 結果が読めない行は捨てる（手で壊した行の保険）

    out.push({
      id:       r['id'],
      at:       at,
      teacher:  r['講師id'],
      student:  r['生徒id'],
      book:     r['教材id'],
      page:     toInt_(r['ページ'], 0),
      section:  String(r['見出し'] === null || r['見出し'] === undefined ? '' : r['見出し']).trim(),
      major:    toInt_(r['大問'], 1),
      q:        toInt_(r['小問'], 1),
      result:   result,
      cause:    TEXT_TO_CAUSE[String(r['つまずき'] || '').trim()] || 0
    });
  }

  out.sort(function (a, b) { return a.at - b.at; });
  return out;
}

function addRecord_(teacher, p) {
  if (!p.student) throw new Error('生徒が指定されていません。');
  if (!p.book)    throw new Error('教材が指定されていません。');

  var result = toInt_(p.result, 0);
  if (!RESULT_TO_TEXT[result]) throw new Error('結果の値が不正です: ' + p.result);

  var cause = toInt_(p.cause, 0);
  if (CAUSE_TO_TEXT[cause] === undefined) cause = 0;
  if (result !== 2) cause = 0; // つまずきは不正解のときだけ

  // 名前は「読むための控え」。書いた時点の名前を固定する
  var masters = getMasters_();
  var st = null, bk = null;
  masters.students.forEach(function (s) { if (s.id === p.student) st = s; });
  masters.books.forEach(function (b) { if (b.id === p.book) bk = b; });
  if (!st) throw new Error('その生徒は見つかりません。');
  if (!bk) throw new Error('その教材は見つかりません。');

  /* アクセスキー（p.key）とは別物。こちらは記録1件を指す受付キー */
  var key = String(p.recKey || '').trim();
  var id = newId_('r');
  var now = new Date();

  var existing = withLock_(function () {
    /* 応答が落ちて画面が再送してきた場合。書かずに前の結果を返す。
       ロックの中で見てから書くので、二重に入ることはない */
    if (key) {
      var rows = readAll_(SHEETS.RECORD);
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i]['受付キー'] || '').trim() !== key) continue;
        return {
          id: rows[i]['id'],
          at: new Date(rows[i]['日時']).getTime(),
          teacher: rows[i]['講師id'],
          student: rows[i]['生徒id'], book: rows[i]['教材id'],
          page: toInt_(rows[i]['ページ'], 1),
          section: String(rows[i]['見出し'] || ''),
          major: toInt_(rows[i]['大問'], 1), q: toInt_(rows[i]['小問'], 1),
          result: TEXT_TO_RESULT[String(rows[i]['結果']).trim()] || 1,
          cause: TEXT_TO_CAUSE[String(rows[i]['つまずき'] || '').trim()] || 0,
          duplicate: true
        };
      }
    }

    appendRow_(SHEETS.RECORD, {
      'id':       id,
      '日時':      now,
      '講師id':    teacher.id,
      '講師名':    teacher.name,
      '生徒id':    st.id,
      '生徒名':    st.name,
      '教材id':    bk.id,
      '教材略称':   bk.short,
      'ページ':     toInt_(p.page, 1),
      '見出し':     String(p.section || ''),
      '大問':      toInt_(p.major, 1),
      '小問':      toInt_(p.q, 1),
      '結果':      RESULT_TO_TEXT[result],
      'つまずき':   CAUSE_TO_TEXT[cause],
      '受付キー':   key
    });
    return null;
  });

  if (existing) return existing;

  return {
    id: id, at: now.getTime(), teacher: teacher.id,
    student: st.id, book: bk.id,
    page: toInt_(p.page, 1), section: String(p.section || ''),
    major: toInt_(p.major, 1), q: toInt_(p.q, 1),
    result: result, cause: cause
  };
}

/**
 * 取り消し。記録した本人の直近の1件を消す用途を想定している。
 *
 * 既に無い場合もエラーにしない。応答が落ちて画面が再送してくることがあり、
 * そのとき「消せませんでした」と出すのは嘘になるため。消えていれば目的は足りている。
 */
function deleteRecord_(teacher, id) {
  return withLock_(function () {
    var rows = readAll_(SHEETS.RECORD);
    for (var i = 0; i < rows.length; i++) {
      if (rows[i]['id'] !== id) continue;
      deleteRow_(SHEETS.RECORD, rows[i]._row);
      return true;
    }
    return true;
  });
}

/** 受付キーから記録を1件引く。入ったのか入らなかったのか分からないときに使う */
function findRecordByKey_(key) {
  key = String(key || '').trim();
  if (!key) return null;
  var rows = readAll_(SHEETS.RECORD);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['受付キー'] || '').trim() !== key) continue;
    return { id: rows[i]['id'], at: new Date(rows[i]['日時']).getTime() };
  }
  return null;
}

function setCause_(teacher, id, cause) {
  var c = toInt_(cause, 0);
  if (CAUSE_TO_TEXT[c] === undefined) throw new Error('つまずきの値が不正です: ' + cause);

  return withLock_(function () {
    var rows = readAll_(SHEETS.RECORD);
    for (var i = 0; i < rows.length; i++) {
      if (rows[i]['id'] !== id) continue;
      if (TEXT_TO_RESULT[String(rows[i]['結果']).trim()] !== 2) {
        throw new Error('不正解の記録にしか分類は付けられません。');
      }
      updateRow_(SHEETS.RECORD, rows[i]._row, { 'つまずき': CAUSE_TO_TEXT[c] });
      return true;
    }
    throw new Error('その記録は見つかりません。');
  });
}

/* ---------- マスタの追加 ---------- */

function addSection_(teacher, bookId, label) {
  label = String(label || '').trim();
  if (!label) throw new Error('見出しが空です。');

  return withLock_(function () {
    var existing = readAll_(SHEETS.SECTION);
    for (var i = 0; i < existing.length; i++) {
      if (existing[i]['教材id'] === bookId &&
          String(existing[i]['見出し']).trim() === label) {
        return label; // 既にある。エラーにはしない
      }
    }
    var book = null;
    readAll_(SHEETS.BOOK).forEach(function (b) { if (b['id'] === bookId) book = b; });
    if (!book) throw new Error('その教材は見つかりません。');

    appendRow_(SHEETS.SECTION, {
      'id':     newId_('h'),
      '教材id':  bookId,
      '教材略称': book['略称'] || book['書名'],
      '見出し':   label
    });
    return label;
  });
}

function addStudent_(teacher, name, grade) {
  name = String(name || '').trim();
  if (!name) throw new Error('生徒名が空です。');
  var id = newId_('s');
  withLock_(function () {
    appendRow_(SHEETS.STUDENT, {
      'id': id, '名前': name, '学年': String(grade || ''),
      '有効': true, '作成日時': new Date()
    });
  });
  return { id: id, name: name, grade: String(grade || '') };
}

function addBook_(teacher, p) {
  var title = String(p.title || '').trim();
  if (!title) throw new Error('書名が空です。');
  var id = newId_('b');
  var short = String(p.short || title).trim();

  withLock_(function () {
    appendRow_(SHEETS.BOOK, {
      'id': id, '書名': title, '略称': short,
      '科目': String(p.subject || ''), '総ページ': toInt_(p.pages, 0),
      '有効': true, '作成日時': new Date()
    });
    (p.units || []).forEach(function (u) {
      appendRow_(SHEETS.UNIT, {
        'id': newId_('u'), '教材id': id, '教材略称': short,
        '単元名': String(u.name || ''),
        '開始ページ': toInt_(u.from, 0), '終了ページ': toInt_(u.to, 0)
      });
    });
    (p.sections || []).forEach(function (s) {
      appendRow_(SHEETS.SECTION, {
        'id': newId_('h'), '教材id': id, '教材略称': short, '見出し': String(s)
      });
    });
  });
  return { id: id };
}
