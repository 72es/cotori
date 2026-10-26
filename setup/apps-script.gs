// kotori 回饋收集程式
// 貼進 Google 試算表的「擴充功能 → Apps Script」，部署成網頁應用程式。
// 每一筆回饋會新增一列到名為「回饋」的工作表。

const SHEET_NAME = '回饋';
const HEADERS = ['時間', '名字', '類型', '期別', '文章ID', '文章標題', '星等', '內容', '送出ID'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function clip_(v, n) {
  return String(v == null ? '' : v).slice(0, n);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const type = ['rating', 'note', 'question'].indexOf(d.type) >= 0 ? d.type : 'note';
    const stars = Number(d.stars);
    const sh = sheet_();
    // 同一個送出ID 只記一次（手機重送時不會重複）
    const sid = clip_(d.sid, 40);
    if (sid) {
      const last = sh.getLastRow();
      if (last > 1) {
        const ids = sh.getRange(Math.max(2, last - 300), 9, Math.min(300, last - 1), 1).getValues().flat();
        if (ids.indexOf(sid) >= 0) return ContentService.createTextOutput('dup');
      }
    }
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
  return ContentService.createTextOutput('kotori ok');
}
