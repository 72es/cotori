// cotori 回饋收集程式（第 3 版：加上跨裝置同步）
// 貼進 Google 試算表的「擴充功能 → Apps Script」，部署成網頁應用程式。
// 「回饋」工作表：大家主動送出的星評、留言、提問（記名字）。
// 「閱讀」工作表：閱讀行為（讀完、點開路過、點原文），不記名字。

const FB_SHEET = '回饋';
const FB_HEADERS = ['時間', '名字', '類型', '期別', '文章ID', '文章標題', '星等', '內容', '送出ID'];
const READ_SHEET = '閱讀';
const READ_HEADERS = ['時間', '事件', '期別', '文章ID', '分類', '停留秒數', '送出ID'];
const FB_TYPES = ['rating', 'note', 'question'];
const READ_TYPES = ['read', 'open_skip', 'open_source', 'save'];

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function clip_(v, n) {
  return String(v == null ? '' : v).slice(0, n);
}

function seen_(sh, col, sid) {
  if (!sid) return false;
  const last = sh.getLastRow();
  if (last < 2) return false;
  const n = Math.min(300, last - 1);
  const ids = sh.getRange(last - n + 1, col, n, 1).getValues().flat();
  return ids.indexOf(sid) >= 0;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.type === 'register' || d.type === 'sync') return json_(sync_(d));
    const sid = clip_(d.sid, 40);
    if (READ_TYPES.indexOf(d.type) >= 0) {
      const sh = sheet_(READ_SHEET, READ_HEADERS);
      if (seen_(sh, 7, sid)) return ContentService.createTextOutput('dup');
      sh.appendRow([new Date(), d.type, clip_(d.edition, 20), clip_(d.itemId, 60), clip_(d.sec, 20), Number(d.secs) || '', sid]);
      return ContentService.createTextOutput('ok');
    }
    const type = FB_TYPES.indexOf(d.type) >= 0 ? d.type : 'note';
    const stars = Number(d.stars);
    const sh = sheet_(FB_SHEET, FB_HEADERS);
    if (seen_(sh, 9, sid)) return ContentService.createTextOutput('dup');
    sh.appendRow([
      new Date(),
      clip_(d.name, 30),
      type,
      clip_(d.edition, 20),
      clip_(d.itemId, 60),
      clip_(d.title, 120),
      stars >= 1 && stars <= 5 ? stars : '',
      clip_(d.text, 1000),
      sid
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('cotori ok');
}


// ───────── 跨裝置同步 ─────────
// 「同步」工作表：一人一列。只存「代稱」「PIN 的雜湊」和一包閱讀紀錄（已讀、收藏、星評、留言草稿、提問、偏好）。
// 不存 PIN 原碼。每日出刊的排程不讀這張表，閱讀統計也不從這裡來。
// 忘記 PIN：把那個人的整列刪掉，他重新用「初次使用」設定即可（紀錄會從他手機上的那份重新帶上來）。
const SYNC_SHEET = '同步';
const SYNC_HEADERS = ['代稱', 'PIN雜湊', '更新時間', '資料1', '資料2', '資料3', '資料4', '資料5', '資料6', '資料7', '資料8'];
const CHUNK = 45000;          // Google 試算表一格最多 5 萬字
const DAY = 86400000;

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function salt_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SYNC_SALT');
  if (!s) { s = Utilities.getUuid(); p.setProperty('SYNC_SALT', s); }
  return s;
}

function hash_(h) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt_() + '|' + h, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}

function key_(name) {
  return String(name || '').trim().toLowerCase().slice(0, 30);
}

function findRow_(sh, k) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const names = sh.getRange(2, 1, last - 1, 1).getValues().flat().map(function (v) { return key_(v); });
  const i = names.indexOf(k);
  return i < 0 ? 0 : i + 2;
}

function readBlob_(sh, row) {
  const cells = sh.getRange(row, 4, 1, 8).getValues()[0].join('');
  try { return cells ? JSON.parse(cells) : {}; } catch (e) { return {}; }
}

function writeBlob_(sh, row, obj) {
  const s = JSON.stringify(obj);
  if (s.length > CHUNK * 8) throw new Error('too big');
  const parts = [];
  for (let i = 0; i < 8; i++) parts.push(s.slice(i * CHUNK, (i + 1) * CHUNK));
  sh.getRange(row, 4, 1, 8).setValues([parts]);
  sh.getRange(row, 3).setValue(new Date());
}

function sync_(d) {
  const k = key_(d.name);
  const h = String(d.h || '');
  if (!k || !/^[0-9a-f]{64}$/.test(h)) return { ok: false, err: 'bad' };
  const sh = sheet_(SYNC_SHEET, SYNC_HEADERS);
  sh.getRange('A:B').setNumberFormat('@');
  let row = findRow_(sh, k);
  const st = d.state || {};
  if (d.type === 'register') {
    if (row) return { ok: false, err: 'exists' };
    sh.appendRow([String(d.name).trim().slice(0, 30), hash_(h), new Date()]);
    row = sh.getLastRow();
    const clean = mergeState_({}, st, Date.now());
    writeBlob_(sh, row, clean);
    return { ok: true, state: clean, name: String(d.name).trim().slice(0, 30) };
  }
  if (!row) return { ok: false, err: 'noone' };
  const meta = sh.getRange(row, 1, 1, 3).getValues()[0];
  if (meta[1] !== hash_(h)) return { ok: false, err: 'badpin' };
  const merged = mergeState_(readBlob_(sh, row), st, Date.now());
  writeBlob_(sh, row, merged);
  return { ok: true, state: merged, name: String(meta[0]) };
}

// 兩份紀錄合併：已讀取聯集；收藏、星評、偏好看時間，新的蓋舊的（取消收藏也會同步）；提問依送出ID合併。
function mergeState_(a, b, now) {
  a = a || {}; b = b || {};
  const out = { v: 1, read: {}, saved: {}, fb: {}, asked: [], prefs: null };
  [a.read, b.read].forEach(function (m) {
    Object.keys(m || {}).forEach(function (id) {
      const t = Number(m[id]) || 1;
      if (now - t < 90 * DAY || t < 1e12) out.read[id] = Math.max(out.read[id] || 0, t);
    });
  });
  [a.saved, b.saved].forEach(function (m) {
    Object.keys(m || {}).forEach(function (id) {
      const x = m[id]; if (!x || typeof x !== 'object') return;
      const cur = out.saved[id];
      if (!cur || (Number(x.t) || 0) > (Number(cur.t) || 0)) out.saved[id] = x;
    });
  });
  Object.keys(out.saved).forEach(function (id) {
    const x = out.saved[id];
    if (x.del && now - (Number(x.t) || 0) > 90 * DAY) delete out.saved[id];
  });
  [a.fb, b.fb].forEach(function (m) {
    Object.keys(m || {}).forEach(function (id) {
      const x = m[id]; if (!x || typeof x !== 'object') return;
      const cur = out.fb[id];
      if (!cur || (Number(x.t) || 0) >= (Number(cur.t) || 0)) out.fb[id] = x;
    });
  });
  const seen = {};
  (a.asked || []).concat(b.asked || []).forEach(function (q) {
    if (!q) return;
    const k = q.sid || (q.ts + '|' + q.text);
    if (seen[k]) return; seen[k] = 1; out.asked.push(q);
  });
  out.asked.sort(function (x, y) { return (x.ts || 0) - (y.ts || 0); });
  out.asked = out.asked.slice(-200);
  const pa = a.prefs, pb = b.prefs;
  out.prefs = !pa ? (pb || null) : !pb ? pa : ((Number(pb.t) || 0) >= (Number(pa.t) || 0) ? pb : pa);
  return out;
}
