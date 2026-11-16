/**
 * 展示アンケート集約 — 共有サーバー（Google Apps Script）
 *
 * 設定手順（集計担当が1回だけ。PC推奨）
 *  1. https://script.google.com で「新しいプロジェクト」を作る
 *  2. 最初からあるコードを消して、このコードを貼り付けて保存
 *  3. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *       実行するユーザー：自分 ／ アクセスできるユーザー：全員
 *     →「デプロイ」→ アクセスを許可
 *  4. 表示されたURL（…/exec）をアプリの「共有サーバーの設定」に貼って「接続」
 *
 * データはGoogleドライブの「展示アンケート集約データ」に、1人1行（ID単位）で保存されます。
 * 集計から外す行は「除外」列に何か入力します。
 * コードを直したら「デプロイ」→「デプロイを管理」→ 編集 →「新バージョン」（URLは変わりません）。
 */
var APP = 'seal-survey';
var SHEET_NAME = 'データ';
var BOOK_NAME = '展示アンケート集約データ';
var HEADER = ['ID', '入力者', '会社名', '日付', '①合計', '②合計', '最終更新', '保存番号', '受信日時', '除外', 'データ'];
var ID_RE = /^\d{8}-[0-9A-HJKMNP-TV-Z]{6}$/;

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'list') return out_(listAll_());
    return out_({ok: true, app: APP, version: 1, time: new Date().toISOString()});
  } catch (err) { return out_({ok: false, error: String(err && err.message || err)}); }
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e && e.postData ? e.postData.contents : ''); }
  catch (err) { return out_({ok: false, error: 'JSONを読み取れません'}); }
  try {
    if (!req || typeof req !== 'object') return out_({ok: false, error: 'データがありません'});
    if (req.action === 'put') return out_(put_(req.sessions));
    if (req.action === 'exclude') return out_(exclude_(req.ids, req.exclude !== false));
    if (req.action === 'list') return out_(listAll_());
    if (req.action === 'ping') return out_({ok: true, app: APP, version: 1, time: new Date().toISOString()});
    return out_({ok: false, error: '不明な操作です'});
  } catch (err) { return out_({ok: false, error: String(err && err.message || err)}); }
}

/** エディタで一度実行すると、保存先のスプレッドシートを作ってURLをログに出します（任意）。 */
function setup() {
  var url = withLock_(function () { return {ok: true, url: sheet_().getParent().getUrl()}; }).url;
  Logger.log('保存先: ' + url);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return {ok: false, busy: true, error: '混み合っています。少し待ってから自動で再送します'};
  try { return fn(); } finally { try { SpreadsheetApp.flush(); } catch (err) {} lock.releaseLock(); }
}

function sheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID'), ss = null, created = false;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (err) { ss = null; } }
  if (!ss) { try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (err) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create(BOOK_NAME); created = true; }
  if (ss.getId() !== id) props.setProperty('SHEET_ID', ss.getId());
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = created ? ss.getSheets()[0].setName(SHEET_NAME) : ss.insertSheet(SHEET_NAME);
    sh.getRange('A:D').setNumberFormat('@');
    sh.getRange('G:H').setNumberFormat('@');
    sh.getRange('J:K').setNumberFormat('@');
    sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADER.length).getValues().map(function (r) {
    var d = null;
    try { d = JSON.parse(r[10]); } catch (err) { d = null; }
    return {id: String(r[0]).replace(/^'/, ''), data: d, excluded: r[9]};
  });
}

// 書き込み時はID列（A）だけを読み、比べる行のデータ（K）だけを1セル読む（行が増えても速い）
function readIds_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]).replace(/^'/, ''); });
}
function cellJson_(sh, r) { try { return JSON.parse(sh.getRange(r, 11).getValue()); } catch (err) { return null; } }

function safe_(s) { s = String(s); return /^[=+\-@\t\r]/.test(s) ? "'" + s : s; }
function stamp_(s) { return (s && s.updatedAt) || ''; }
// 保存番号（revision）で新旧を決め、時刻は同じ番号のときだけ使う（端末の時計がずれていても正しく判定）
function newer_(a, b) { return a.revision > b.revision || (a.revision === b.revision && stamp_(a) > stamp_(b)); }
function same_(a, b) { return a.revision === b.revision && stamp_(a) === stamp_(b); }

function check_(s) {
  if (!s || typeof s !== 'object' || !/(^|-)seal-survey$/.test(String(s.format)) || s.version !== 2 || !ID_RE.test(String(s.id))) return null;
  var m = s.meta;
  if (!m || typeof m.name !== 'string' || typeof m.company !== 'string' || typeof m.date !== 'string' || typeof m.createdAt !== 'string') return null;
  if (!m.name || !m.company || m.name.length > 80 || m.company.length > 120 || !/^\d{4}-\d{2}-\d{2}$/.test(m.date)) return null;
  var day = new Date(m.date + 'T00:00:00Z');
  if (isNaN(day.getTime()) || day.toISOString().slice(0, 10) !== m.date) return null;
  if (typeof s.revision !== 'number' || s.revision < 0 || Math.floor(s.revision) !== s.revision) return null;
  if (JSON.stringify(s).length > 40000) return null;
  var q1 = 0, q2 = 0, ok = true;
  var nonNeg = function (v) { return typeof v === 'number' && v >= 0 && Math.floor(v) === v && v <= 1000000; };
  ['1', '2'].forEach(function (k) {
    var d = s.days && s.days[k];
    if (!d) { ok = false; return; }
    ['student', 'adult'].forEach(function (t) {
      var f = d[t];
      if (!f || !Array.isArray(f.matrix) || !Array.isArray(f.q2) || !nonNeg(f.unknown)) { ok = false; return; }
      f.matrix.forEach(function (row) { if (!Array.isArray(row)) { ok = false; return; } row.forEach(function (v) { if (!nonNeg(v)) ok = false; else q1 += v; }); });
      q1 += f.unknown;
      f.q2.forEach(function (v) { if (!nonNeg(v)) ok = false; else q2 += v; });
    });
  });
  return ok ? {q1: q1, q2: q2} : null;
}

function put_(sessions) {
  if (!Array.isArray(sessions)) return {ok: false, error: 'データがありません'};
  return withLock_(function () {
    var sh = sheet_(), ids = readIds_(sh), index = {}, results = [];
    ids.forEach(function (id, k) { if (index[id] === undefined) index[id] = k; });
    sessions.slice(0, 100).forEach(function (s) {
      var v = check_(s);
      if (!v) { results.push({id: s && s.id ? String(s.id).slice(0, 40) : '', status: 'invalid'}); return; }
      var row = [s.id, safe_(s.meta.name), safe_(s.meta.company), s.meta.date, v.q1, v.q2, stamp_(s), String(s.revision), new Date(), '', JSON.stringify(s)];
      var i = index[s.id];
      if (i === undefined) {
        sh.appendRow(row);
        index[s.id] = ids.length;
        ids.push(s.id);
        results.push({id: s.id, status: 'saved', revision: s.revision});
        return;
      }
      var cur = cellJson_(sh, i + 2);
      if (cur && cur.meta && cur.meta.createdAt !== s.meta.createdAt) { results.push({id: s.id, status: 'conflict'}); return; }
      if (cur && typeof cur.revision === 'number' && !newer_(s, cur)) { results.push({id: s.id, status: same_(s, cur) ? 'saved' : 'stale', revision: cur.revision}); return; }
      sh.getRange(i + 2, 1, 1, 9).setValues([row.slice(0, 9)]);   // 「除外」列（J）はそのまま
      sh.getRange(i + 2, 11).setValue(row[10]);
      results.push({id: s.id, status: 'saved', revision: s.revision});
    });
    return {ok: true, results: results, time: new Date().toISOString()};
  });
}

function listAll_() {
  var read = function () {
    var best = {}, order = [], sessions = [], excluded = 0;
    readRows_(sheet_()).forEach(function (r) {
      if (!r.data || typeof r.data !== 'object') return;
      var cur = best[r.id];
      if (!cur) { best[r.id] = r; order.push(r.id); }
      else if (typeof r.data.revision === 'number' && (typeof cur.data.revision !== 'number' || newer_(r.data, cur.data))) best[r.id] = r;
    });
    order.forEach(function (id) {
      var r = best[id];
      if (String(r.excluded || '').trim()) { excluded++; return; }
      sessions.push(r.data);
    });
    return {ok: true, sessions: sessions, excluded: excluded, time: new Date().toISOString()};
  };
  return PropertiesService.getScriptProperties().getProperty('SHEET_ID') ? read() : withLock_(read);
}

function exclude_(ids, flag) {
  if (!Array.isArray(ids)) return {ok: false, error: 'IDがありません'};
  return withLock_(function () {
    var sh = sheet_(), all = readIds_(sh), want = {}, changed = 0;
    ids.slice(0, 500).forEach(function (id) { if (ID_RE.test(String(id))) want[id] = true; });
    all.forEach(function (id, i) {
      if (want[id]) { sh.getRange(i + 2, 10).setValue(flag ? '除外' : ''); changed++; }
    });
    return {ok: true, changed: changed, time: new Date().toISOString()};
  });
}
