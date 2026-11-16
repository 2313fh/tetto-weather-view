/**
 * 展示アンケート集約 — 共有サーバー（Google Apps Script）
 *
 * 設定手順（PCのブラウザー推奨・約5分。代表者が1回だけ）
 *  1. https://script.google.com を開き「新しいプロジェクト」を作る
 *  2. 最初からあるコードを全部消して、このコードを貼り付けて保存（Ctrl+S / ⌘+S）
 *  3. 右上「デプロイ」→「新しいデプロイ」→ 種類の選択（歯車）で「ウェブアプリ」
 *       次のユーザーとして実行：自分
 *       アクセスできるユーザー：全員
 *     →「デプロイ」→「アクセスを承認」（「詳細」→「…（安全ではないページ）に移動」→「許可」）
 *  4. 表示された「ウェブアプリ」のURL（…/exec）をコピーし、アプリの「共有サーバー」に貼り付けて「接続」
 *
 * データは、あなたのGoogleドライブに自動で作られるスプレッドシート
 * 「展示アンケート集約データ」に、入力者ごと（ID単位）に1行で保存されます。
 * 同じIDは新しい内容で上書きされ、行が重複することはありません。
 * 集計から外したい行は「除外」列に何か入力してください（例：除外）。
 * コードを直したときは「デプロイ」→「デプロイを管理」→ 編集 →「新バージョン」で更新します（URLは変わりません）。
 */
var APP = 'kyudentd-seal-survey';
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
  try { return fn(); } finally { lock.releaseLock(); }
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

function safe_(s) { s = String(s); return /^[=+\-@\t\r]/.test(s) ? "'" + s : s; }
function stamp_(s) { return (s && s.updatedAt) || ''; }
function newer_(a, b) { return stamp_(a) > stamp_(b) || (stamp_(a) === stamp_(b) && a.revision > b.revision); }
function same_(a, b) { return stamp_(a) === stamp_(b) && a.revision === b.revision; }

function check_(s) {
  if (!s || typeof s !== 'object' || s.format !== APP || s.version !== 2 || !ID_RE.test(String(s.id))) return null;
  var m = s.meta;
  if (!m || typeof m.name !== 'string' || typeof m.company !== 'string' || typeof m.date !== 'string' || typeof m.createdAt !== 'string') return null;
  if (!m.name || !m.company || m.name.length > 80 || m.company.length > 120 || !/^\d{4}-\d{2}-\d{2}$/.test(m.date)) return null;
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
    var sh = sheet_(), rows = readRows_(sh), index = {}, results = [];
    rows.forEach(function (r, i) { index[r.id] = i; });
    sessions.slice(0, 100).forEach(function (s) {
      var v = check_(s);
      if (!v) { results.push({id: s && s.id ? String(s.id).slice(0, 40) : '', status: 'invalid'}); return; }
      var i = index[s.id];
      var row = [s.id, safe_(s.meta.name), safe_(s.meta.company), s.meta.date, v.q1, v.q2, stamp_(s), String(s.revision), new Date(), '', JSON.stringify(s)];
      if (i === undefined) {
        sh.appendRow(row);
        index[s.id] = rows.length;
        rows.push({id: s.id, data: s, excluded: ''});
        results.push({id: s.id, status: 'saved', revision: s.revision});
        return;
      }
      var cur = rows[i].data;
      if (cur && cur.meta && cur.meta.createdAt !== s.meta.createdAt) { results.push({id: s.id, status: 'conflict'}); return; }
      if (cur && !newer_(s, cur)) { results.push({id: s.id, status: same_(s, cur) ? 'saved' : 'stale', revision: cur.revision}); return; }
      row[9] = rows[i].excluded;
      sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
      rows[i].data = s;
      results.push({id: s.id, status: 'saved', revision: s.revision});
    });
    return {ok: true, results: results, time: new Date().toISOString()};
  });
}

function listAll_() {
  return withLock_(function () {
    var sessions = [], excluded = 0;
    readRows_(sheet_()).forEach(function (r) {
      if (!r.data) return;
      if (String(r.excluded || '').trim()) { excluded++; return; }
      sessions.push(r.data);
    });
    return {ok: true, sessions: sessions, excluded: excluded, time: new Date().toISOString()};
  });
}

function exclude_(ids, flag) {
  if (!Array.isArray(ids)) return {ok: false, error: 'IDがありません'};
  return withLock_(function () {
    var sh = sheet_(), rows = readRows_(sh), want = {}, changed = 0;
    ids.slice(0, 500).forEach(function (id) { if (ID_RE.test(String(id))) want[id] = true; });
    rows.forEach(function (r, i) {
      if (want[r.id]) { sh.getRange(i + 2, 10).setValue(flag ? '除外' : ''); changed++; }
    });
    return {ok: true, changed: changed, time: new Date().toISOString()};
  });
}
