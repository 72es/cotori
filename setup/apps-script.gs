// cotori 回饋收集程式（第 2 版）
// 貼進 Google 試算表的「擴充功能 → Apps Script」，部署成網頁應用程式。
// 「回饋」工作表：大家主動送出的星評、留言、提問（記名字）。
// 「閱讀」工作表：閱讀行為（讀完、點開跳過沒差、點原文），不記名字。

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
