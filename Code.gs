/**
 * 案件進捗管理アプリ（実データ連携版）
 * 「自社案件進捗管理表」の実シート（営業部・営業部（ビルメン）・１課・２課・３課・函館・旭川）を
 * 直接読み込み、既存列には一切書き込まない。
 * ステータス・失注理由・受注決め手・工事日・仕入情報などの追加項目は
 * 別シート「案件進捗_追加情報」に 案件No(+拠点) をキーとして保持し、読み込み時にJOINして返す。
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

// ════════════════════════════════════════════════
// 設定：既存シート一覧・列位置
// ════════════════════════════════════════════════

const SOURCE_SHEETS = ['営業部', '営業部（ビルメン）', '１課', '２課', '３課', '函館', '旭川'];
const HEADER_ROW = 8;
const DATA_START_ROW = 9;

// B列(2)〜S列(19)の共通レイアウト
const COL = {
  no: 2, dealNo: 3, customerName: 4, siteName: 5, projectName: 6,
  summary: 7, occurredDate: 8, source: 9, assignee: 10,
  quoteAmount: 11, plannedProfit: 12, rank: 13, expectedOrderMonth: 14,
  currentStatus: 15, pendingIssue: 16, confirmedAmount: 17, profit: 18, profitRate: 19
};

const EXTRA_SHEET_NAME = '案件進捗_追加情報';
const EXTRA_HEADERS = [
  'id', 'status', 'quoteDate', 'lostReason', 'winFactor',
  'constructionStart', 'constructionEnd',
  'supplier', 'purchaseDate', 'purchaseAmount', 'billingMonth',
  'updatedAt'
];

function ensureExtraSheet_() {
  let sh = sheet(EXTRA_SHEET_NAME);
  if (!sh) {
    sh = SS().insertSheet(EXTRA_SHEET_NAME);
    sh.appendRow(EXTRA_HEADERS);
  }
  return sh;
}

// 初回セットアップ用：一度手動実行してください
function setup() {
  ensureExtraSheet_();
  Logger.log('セットアップ完了：「' + EXTRA_SHEET_NAME + '」シートを作成しました。');
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
      case 'deal_extra_upsert': return dealExtraUpsert(payload.data || payload);
      default:                  return makeErr('不明なaction: ' + action);
    }
  } catch (err) {
    return makeErr(err.toString());
  }
}

// ════════════════════════════════════════════════
// 既存シートの読み込み（読み取り専用）
// ════════════════════════════════════════════════

function readSourceSheet_(sheetName) {
  const sh = sheet(sheetName);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const numRows = lastRow - DATA_START_ROW + 1;
  const vals = sh.getRange(DATA_START_ROW, 1, numRows, 19).getValues();

  const out = [];
  vals.forEach((row, i) => {
    const dealNo = row[COL.dealNo - 1];
    const customerName = row[COL.customerName - 1];
    // 実績行やただの空行はスキップ（案件No・顧客名がどちらも空なら非データ行）
    if ((dealNo === '' || dealNo === null) && (customerName === '' || customerName === null)) return;

    const rowNum = DATA_START_ROW + i;
    const id = sheetName + '::' + (dealNo !== '' && dealNo !== null ? dealNo : ('r' + rowNum));

    out.push({
      id: id,
      branch: sheetName,
      dealNo: dealNo || '',
      rowNum: rowNum,
      customerName: customerName || '',
      siteName: row[COL.siteName - 1] || '',
      projectName: row[COL.projectName - 1] || '',
      summary: row[COL.summary - 1] || '',
      occurredDate: row[COL.occurredDate - 1] || '',
      source: row[COL.source - 1] || '',
      assignee: row[COL.assignee - 1] || '',
      quoteAmount: row[COL.quoteAmount - 1] || '',
      plannedProfit: row[COL.plannedProfit - 1] || '',
      rank: row[COL.rank - 1] || '',
      expectedOrderMonth: row[COL.expectedOrderMonth - 1] || '',
      currentStatus: row[COL.currentStatus - 1] || '',
      pendingIssue: row[COL.pendingIssue - 1] || '',
      confirmedAmount: row[COL.confirmedAmount - 1] || '',
      profit: row[COL.profit - 1] || '',
      profitRate: row[COL.profitRate - 1] || '',
    });
  });
  return out;
}

function readAllSources_() {
  let all = [];
  SOURCE_SHEETS.forEach(name => {
    all = all.concat(readSourceSheet_(name));
  });
  return all;
}

function readExtraMap_() {
  const sh = ensureExtraSheet_();
  const vals = sh.getDataRange().getValues();
  const map = {};
  if (vals.length < 2) return map;
  const headers = vals[0];
  for (let i = 1; i < vals.length; i++) {
    const obj = {};
    headers.forEach((h, c) => { obj[h] = vals[i][c]; });
    if (obj.id) map[obj.id] = obj;
  }
  return map;
}

// 経過日数（見積提出日 優先、無ければ発生日。ステータスが受注系/失注なら計算しない）
function calcElapsedDays_(baseDate, status) {
  if (status && ['受注', '工事中', '完了', '失注'].indexOf(status) >= 0) return null;
  if (!baseDate) return null;
  const d = new Date(baseDate);
  if (isNaN(d.getTime())) return null;
  const now = new Date();
  const diffMs = now.setHours(0,0,0,0) - new Date(d).setHours(0,0,0,0);
  return Math.floor(diffMs / 86400000);
}

function dealList() {
  const sources = readAllSources_();
  const extraMap = readExtraMap_();

  const merged = sources.map(d => {
    const extra = extraMap[d.id] || {};
    const status = extra.status || '';
    const baseDate = extra.quoteDate || d.occurredDate;
    return Object.assign({}, d, {
      status: status,
      quoteDate: extra.quoteDate || '',
      lostReason: extra.lostReason || '',
      winFactor: extra.winFactor || '',
      constructionStart: extra.constructionStart || '',
      constructionEnd: extra.constructionEnd || '',
      supplier: extra.supplier || '',
      purchaseDate: extra.purchaseDate || '',
      purchaseAmount: extra.purchaseAmount || '',
      billingMonth: extra.billingMonth || '',
      elapsedDays: calcElapsedDays_(baseDate, status)
    });
  });

  return makeRes(merged);
}

// ════════════════════════════════════════════════
// 追加情報の upsert（既存シートには一切書き込まない）
// ════════════════════════════════════════════════

function rangesOverlap_(aStart, aEnd, bStart, bEnd) {
  if (!aStart || !bStart) return false;
  const as = new Date(aStart).getTime();
  const ae = new Date(aEnd || aStart).getTime();
  const bs = new Date(bStart).getTime();
  const be = new Date(bEnd || bStart).getTime();
  return as <= be && bs <= ae;
}

function dealExtraUpsert(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!data.id) return makeErr('idが指定されていません');

    const sh = ensureExtraSheet_();
    const vals = sh.getDataRange().getValues();
    const headers = vals[0];
    const idCol = headers.indexOf('id');

    // 工事日重複チェック（他案件の追加情報のみが対象。失注は除外）
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
          conflicts.push({ id: rowId, constructionStart: rowStart, constructionEnd: rowEnd });
        }
      }
    }

    data.updatedAt = new Date();

    for (let i = 1; i < vals.length; i++) {
      if (vals[i][idCol] == data.id) {
        const row = headers.map(h => (data[h] !== undefined ? data[h] : vals[i][headers.indexOf(h)]));
        sh.getRange(i + 1, 1, 1, headers.length).setValues([row]);
        return makeRes({ id: data.id, conflicts });
      }
    }

    const row = headers.map(h => (data[h] !== undefined ? data[h] : ''));
    sh.appendRow(row);
    return makeRes({ id: data.id, conflicts });
  } catch (err) {
    return makeErr('dealExtraUpsert error: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}
