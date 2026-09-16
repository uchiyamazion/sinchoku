/**
 * 案件進捗管理アプリ（独立版）
 * このスプレッドシート専用のGASプロジェクトとして使用してください。
 * fron-kanri / ac-inspection とは無関係の独立したバックエンドです。
 */

const SS = () => SpreadsheetApp.getActiveSpreadsheet();
const sheet = (name) => SS().getSheetByName(name);

function makeRes(data, status) {
  status = status || 'success';
  const payload = JSON.stringify({ status, data });
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}
function makeErr(msg) {
  return makeRes(msg, 'error');
}

function sheetToObjects(sh) {
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0];
  return vals.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function deleteRowById(sh, id) {
  const vals = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][idCol] == id) {
      sh.deleteRow(i + 1);
      return makeRes({ id });
    }
  }
  return makeErr('対象レコードが見つかりません: ' + id);
}

// ════════════════════════════════════════════════
// doGet / doPost ルーター
// ════════════════════════════════════════════════

function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action || 'deal_list';
    switch (action) {
      case 'deal_list': return dealList();
      default:          return makeErr('不明なaction: ' + action);
    }
  } catch (err) {
    return makeErr(err.toString());
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const payload = body.payload || body;

    switch (action) {
      case 'deal_upsert': return dealUpsert(payload.data || payload);
      case 'deal_delete': return dealDelete(payload.id || payload);
      default:            return makeErr('不明なaction: ' + action);
    }
  } catch (err) {
    return makeErr(err.toString());
  }
}

// ════════════════════════════════════════════════
// 案件管理（見積〜受注〜工事〜請求のパイプライン）
// ════════════════════════════════════════════════

const DEAL_SHEET_NAME = '案件管理';
const DEAL_HEADERS = [
  'id', 'customerName', 'projectName',
  'quoteDate', 'quoteAmount',
  'status',                 // 見積中 / 受注 / 失注 / 工事中 / 完了
  'orderDate', 'lostReason', 'winFactor',
  'constructionStart', 'constructionEnd',
  'supplier', 'purchaseDate', 'purchaseAmount', 'billingMonth',
  'assignee', 'note',
  'createdAt', 'updatedAt'
];

function ensureDealSheet_() {
  let sh = sheet(DEAL_SHEET_NAME);
  if (!sh) {
    sh = SS().insertSheet(DEAL_SHEET_NAME);
    sh.appendRow(DEAL_HEADERS);
  }
  return sh;
}

// 初回セットアップ用：手動で一度実行するとシートを作成します
function setup() {
  ensureDealSheet_();
  Logger.log('セットアップ完了：「' + DEAL_SHEET_NAME + '」シートを作成しました。');
}

function nextDealId_(sh) {
  const vals = sh.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < vals.length; i++) {
    const m = String(vals[i][0]).match(/^D(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  return 'D' + String(max + 1).padStart(4, '0');
}

// 経過日数（見積中のみ計算。それ以外はnull）
function calcElapsedDays_(quoteDate, status) {
  if (status !== '見積中' || !quoteDate) return null;
  const d = new Date(quoteDate);
  if (isNaN(d.getTime())) return null;
  const now = new Date();
  const diffMs = now.setHours(0,0,0,0) - new Date(d).setHours(0,0,0,0);
  return Math.floor(diffMs / 86400000);
}

// 期間重複判定
function rangesOverlap_(aStart, aEnd, bStart, bEnd) {
  if (!aStart || !bStart) return false;
  const as = new Date(aStart).getTime();
  const ae = new Date(aEnd || aStart).getTime();
  const bs = new Date(bStart).getTime();
  const be = new Date(bEnd || bStart).getTime();
  return as <= be && bs <= ae;
}

function dealList() {
  const sh = ensureDealSheet_();
  const rows = sheetToObjects(sh);
  rows.forEach(r => {
    r.elapsedDays = calcElapsedDays_(r.quoteDate, r.status);
  });
  return makeRes(rows);
}

function dealUpsert(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ensureDealSheet_();
    const vals = sh.getDataRange().getValues();
    const headers = vals[0];
    const idCol = headers.indexOf('id');
    const now = new Date();

    // 工事日重複チェック（自分自身と失注案件は除外）
    const conflicts = [];
    if (data.constructionStart) {
      for (let i = 1; i < vals.length; i++) {
        const rowId = vals[i][idCol];
        if (rowId === data.id) continue;
        const rowStatus = vals[i][headers.indexOf('status')];
        if (rowStatus === '失注') continue;
        const rowStart = vals[i][headers.indexOf('constructionStart')];
        const rowEnd = vals[i][headers.indexOf('constructionEnd')];
        if (rangesOverlap_(data.constructionStart, data.constructionEnd, rowStart, rowEnd)) {
          conflicts.push({
            id: rowId,
            projectName: vals[i][headers.indexOf('projectName')],
            customerName: vals[i][headers.indexOf('customerName')],
            constructionStart: rowStart,
            constructionEnd: rowEnd
          });
        }
      }
    }

    if (data.id) {
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][idCol] == data.id) {
          data.updatedAt = now;
          const row = headers.map(h => (data[h] !== undefined ? data[h] : vals[i][headers.indexOf(h)]));
          sh.getRange(i + 1, 1, 1, headers.length).setValues([row]);
          return makeRes({ id: data.id, conflicts });
        }
      }
    }

    data.id = nextDealId_(sh);
    data.createdAt = now;
    data.updatedAt = now;
    const row = headers.map(h => (data[h] !== undefined ? data[h] : ''));
    sh.appendRow(row);
    return makeRes({ id: data.id, conflicts });
  } catch (err) {
    return makeErr('dealUpsert error: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

function dealDelete(id) {
  const sh = ensureDealSheet_();
  return deleteRowById(sh, id);
}
