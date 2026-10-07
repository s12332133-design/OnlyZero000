/** Google 試算表排行榜。修改 ID 後先執行 setup，再部署為網頁應用程式。
 * GET  /exec?action=top 讀取前 10 名（callback 可選，僅供公開資料 JSONP）。
 * POST /exec  JSON 或表單：action=submit, nickname=暱稱, score=分數。
 * 同暱稱只顯示最佳成績；暱稱不是身分驗證。時間模式請使用秒。
 */
const CONFIG = Object.freeze({
  SPREADSHEET_ID: '17j_YBi5cFj6lB8gURuErbM_lPxzIHazUAS2Cn_ZsZoM',
  SHEET_NAME: 'Leaderboard',
  ORDER: 'desc', // desc：分數越高越好；asc：時間越短越好
  MAX_VALUE: 1000000
});
const HEADERS = ['nickname', 'score', 'submittedAt', 'mode'];

function setup() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange('A:A').setNumberFormat('@');
      sheet.getRange('C:C').setNumberFormat('yyyy-MM-dd HH:mm:ss');
    }
    sheet.getRange(1, 5).setValue('requestId');
    return sheet.getName();
  } finally { lock.releaseLock(); }
}

function sheet_() {
  if (!['asc', 'desc'].includes(CONFIG.ORDER)) throw new Error('ORDER 必須是 asc 或 desc');
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) throw new Error('請先在 Apps Script 執行 setup');
  const actual = sheet.getRange(1, 1, 1, 4).getValues()[0];
  if (actual.join('|') !== HEADERS.join('|')) throw new Error('排行榜欄位不符，請使用新的工作表');
  return sheet;
}

function doGet(e) {
  const p = e && e.parameter || {};
  let result;
  try {
    if (p.action === 'status') {
      if (!/^[a-f0-9-]{36}$/.test(p.requestId || '')) throw new Error('requestId 格式不正確');
      const sheet = sheet_();
      const count = sheet.getLastRow() - 1;
      const found = count > 0 && sheet.getRange(2, 5, count, 1).getValues().some(r => r[0] === p.requestId);
      result = { ok: true, saved: !!found };
    } else {
      if ((p.action || 'top') !== 'top') throw new Error('GET 僅支援 top 或 status');
      result = { ok: true, order: CONFIG.ORDER, leaderboard: top_() };
    }
  } catch (error) { result = error_(error); }
  if (p.callback) {
    if (!/^[A-Za-z_$][A-Za-z0-9_$]{0,63}$/.test(p.callback)) {
      return json_({ ok: false, error: 'callback 格式不正確' });
    }
    return ContentService.createTextOutput(p.callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(result);
}

function doPost(e) {
  try {
    if (e && e.postData && e.postData.contents.length > 4096) throw new Error('資料太大');
    let p = e && e.parameter || {};
    // 支援 application/json 與 text/plain JSON，亦支援 URLSearchParams 表單。
    if (e && e.postData && /^\s*\{/.test(e.postData.contents)) {
      p = JSON.parse(e.postData.contents);
    }
    if (p.action !== 'submit') throw new Error('POST 必須指定 action=submit');
    const entry = validate_(p);
    const requestId = p.requestId || '';
    if (requestId && !/^[a-f0-9-]{36}$/.test(requestId)) throw new Error('requestId 格式不正確');
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = sheet_();
      // 前置單引號確保暱稱以文字儲存，不被當成試算表公式。
      const count = sheet.getLastRow() - 1;
      const exists = requestId && count > 0 && sheet.getRange(2, 5, count, 1).getValues().some(r => r[0] === requestId);
      if (!exists) sheet.appendRow(["'" + entry.nickname, entry.score, new Date(), CONFIG.ORDER, requestId]);
      SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }
    return json_({ ok: true, saved: entry });
  } catch (error) { return json_(error_(error)); }
}

function validate_(p) {
  if (typeof p.nickname !== 'string') throw new Error('請輸入暱稱');
  const nickname = p.nickname.normalize('NFKC').trim();
  if (!nickname || Array.from(nickname).length > 20 || /[\u0000-\u001f\u007f]/.test(nickname)) {
    throw new Error('暱稱須為 1～20 字，不可包含換行或控制字元');
  }
  if (!['number', 'string'].includes(typeof p.score) || String(p.score).trim() === '') {
    throw new Error('請提供有效的分數或秒數');
  }
  const score = Number(p.score);
  if (!Number.isFinite(score) || score < 0 || score > CONFIG.MAX_VALUE) {
    throw new Error('分數或秒數超出範圍');
  }
  return { nickname: nickname, score: score };
}

function top_() {
  const sheet = sheet_();
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  const best = new Map();
  rows.forEach(function(row, index) {
    if (row[3] !== CONFIG.ORDER || typeof row[1] !== 'number' || !Number.isFinite(row[1])) return;
    const nickname = String(row[0]).trim();
    if (!nickname) return;
    const item = { nickname: nickname, score: row[1], row: index };
    const previous = best.get(nickname);
    if (!previous || (CONFIG.ORDER === 'desc' ? item.score > previous.score : item.score < previous.score)) {
      best.set(nickname, item);
    }
  });
  return Array.from(best.values()).sort(function(a, b) {
    return (CONFIG.ORDER === 'desc' ? b.score - a.score : a.score - b.score) || a.row - b.row;
  }).slice(0, 10).map(function(item, index) {
    return { rank: index + 1, nickname: item.nickname, score: item.score };
  });
}

function error_(error) {
  console.error(error && error.stack || error);
  return { ok: false, error: error instanceof Error ? error.message : '處理失敗' };
}
function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
