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

// ════════════════════════════════════════════════
// 年度のシート
// 「旭川_2027」のように「拠点名_年度」の名前のタブを、その拠点の年度シートとして一緒に読み込む。
// 年度（4月始まり）の入ったシートは、同じ拠点の案件として扱う（Noは年度シートごとに1から振り直せる）。
// 拠点名だけのタブ（2026年度のシート）は、これまでどおり。idは「タブ名::案件No」なので、既存の追加情報は、そのまま使える。
// ════════════════════════════════════════════════
function sourceTabs_() {
  const tabs = SOURCE_SHEETS.map(b => ({ tab: b, branch: b, fy: null }));
  SS().getSheets().forEach(s => {
    const n = s.getName();
    const m = n.match(/^(.+)_(\d{4})$/);
    if (m && SOURCE_SHEETS.indexOf(m[1]) >= 0) tabs.push({ tab: n, branch: m[1], fy: Number(m[2]) });
  });
  return tabs;
}
function isSourceTab_(name) { return sourceTabs_().some(t => t.tab === name); }
// 書き込み対象のタブ名（クライアントが sheet を送ってくればそれ、なければ branch）
function tabOf_(data) { return data.sheet || data.branch; }
function currentFy_() {
  const p = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-M').split('-').map(Number);
  return p[1] >= 4 ? p[0] : p[0] - 1;
}
// 新規作成の書き込み先：今年度の年度シート（例：旭川_2027）があればそれ、なければ拠点名のシート
function resolveCreateTab_(data) {
  if (data.sheet && isSourceTab_(data.sheet)) return data.sheet;
  const cand = data.branch + '_' + currentFy_();
  return (sheet(cand) && isSourceTab_(cand)) ? cand : data.branch;
}
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
  'paymentTerms', 'receiptMonth', 'advanceAmount', 'advanceMonth',
  'nextActionDate', 'nextActionNote', 'excluded',
  'partsSupplier', 'partsDate', 'partsAmount', 'partsBillingMonth',
  'outsourceSupplier', 'outsourceDate', 'outsourceAmount', 'outsourceBillingMonth',
  'orderedAt',
  'updatedAt'
];

function ensureExtraSheet_() {
  let sh = sheet(EXTRA_SHEET_NAME);
  if (!sh) {
    sh = SS().insertSheet(EXTRA_SHEET_NAME);
    sh.appendRow(EXTRA_HEADERS);
    return sh;
  }
  // 既存シートに新しい項目の見出しが無ければ、末尾に追加する（列の順序は問わない）
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const cur = sh.getRange(1, 1, 1, lastCol).getValues()[0].filter(String);
  const missing = EXTRA_HEADERS.filter(h => cur.indexOf(h) < 0);
  if (missing.length) {
    sh.getRange(1, cur.length + 1, 1, missing.length).setValues([missing]);
  }
  return sh;
}

// 月の値（"2026-11" など）がスプレッドシート側で日付に変換されていても "yyyy-MM" に揃える
function normMonthStr_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return v || '';
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
      case 'target_list': return targetList();
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

    let result;
    switch (action) {
      case 'deal_extra_upsert': result = dealExtraUpsert(payload.data || payload); break;
      case 'deal_core_update':  result = dealCoreUpdate(payload.data || payload); break;
      case 'deal_core_create':  result = dealCoreCreate(payload.data || payload); break;
      case 'deal_delete':       result = dealDelete(payload.data || payload); break;
      case 'fy_sheets_create':  result = createFiscalYearSheets(payload.data || payload); break;
      case 'values_rename':     result = valuesRename(payload.data || payload); break;
      case 'deal_conflict_check': return dealConflictCheck(payload.data || payload);
      case 'target_save':       return targetSave(payload.data || payload);
      default:                  return makeErr('不明なaction: ' + action);
    }
    // 書き込み系の処理が終わったら、一覧キャッシュを破棄する（次の読み込みで最新を取り直す）
    invalidateDealCache_();
    return result;
  } catch (err) {
    return makeErr(err.toString());
  }
}

// ════════════════════════════════════════════════
// 既存シートの読み込み（読み取り専用）
// ════════════════════════════════════════════════

function readSourceSheet_(sheetName, branchName, fy) {
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
      branch: branchName || sheetName,
      sheet: sheetName,
      sheetFy: fy || '',
      dealNo: dealNo || '',
      rowNum: rowNum,
      no: row[COL.no - 1] || '',
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
  sourceTabs_().forEach(t => {
    all = all.concat(readSourceSheet_(t.tab, t.branch, t.fy));
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

// ════════════════════════════════════════════════
// 一覧のキャッシュ（CacheService）
// 1キー100KBの上限があるため、JSON文字列を分割して保存する
// ════════════════════════════════════════════════

const DEAL_CACHE_TTL_SEC = 90;     // キャッシュの有効秒数（元シートを直接編集した場合の反映遅れの上限）
const DEAL_CACHE_CHUNK = 30000;    // 1チャンクあたりの文字数（日本語でも100KB未満に収める）

function getDealCache_() {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get('dl_n'));
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push('dl_' + i);
    const got = c.getAll(keys);
    let out = '';
    for (let i = 0; i < n; i++) {
      if (got['dl_' + i] === undefined || got['dl_' + i] === null) return null;
      out += got['dl_' + i];
    }
    return out;
  } catch (err) { return null; }
}

function putDealCache_(str) {
  try {
    const c = CacheService.getScriptCache();
    const obj = {};
    let n = 0;
    for (let i = 0; i < str.length; i += DEAL_CACHE_CHUNK) {
      obj['dl_' + n] = str.substring(i, i + DEAL_CACHE_CHUNK);
      n++;
    }
    obj['dl_n'] = String(n);
    c.putAll(obj, DEAL_CACHE_TTL_SEC);
  } catch (err) { /* キャッシュに失敗しても通常動作に影響させない */ }
}

function invalidateDealCache_() {
  try { CacheService.getScriptCache().remove('dl_n'); } catch (err) {}
}

function dealList() {
  const cachedStr = getDealCache_();
  if (cachedStr) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }
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
      orderedAt: extra.orderedAt || '',
      constructionStart: extra.constructionStart || '',
      constructionEnd: extra.constructionEnd || '',
      supplier: extra.supplier || '',
      purchaseDate: extra.purchaseDate || '',
      purchaseAmount: extra.purchaseAmount || '',
      billingMonth: normMonthStr_(extra.billingMonth),
      paymentTerms: extra.paymentTerms || '',
      receiptMonth: normMonthStr_(extra.receiptMonth),
      advanceAmount: extra.advanceAmount || '',
      advanceMonth: normMonthStr_(extra.advanceMonth),
      nextActionDate: extra.nextActionDate || '',
      nextActionNote: extra.nextActionNote || '',
      excluded: extra.excluded ? '1' : '',
      partsSupplier: extra.partsSupplier || '',
      partsDate: extra.partsDate || '',
      partsAmount: extra.partsAmount || '',
      partsBillingMonth: normMonthStr_(extra.partsBillingMonth),
      outsourceSupplier: extra.outsourceSupplier || '',
      outsourceDate: extra.outsourceDate || '',
      outsourceAmount: extra.outsourceAmount || '',
      outsourceBillingMonth: normMonthStr_(extra.outsourceBillingMonth),
      updatedAt: extra.updatedAt || '',
      elapsedDays: calcElapsedDays_(baseDate, status)
    });
  });

  const payloadStr = JSON.stringify({ status: 'success', data: merged });
  putDealCache_(payloadStr);
  return ContentService.createTextOutput(payloadStr).setMimeType(ContentService.MimeType.JSON);
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

// ════════════════════════════════════════════════
// 元シートの基本情報を直接更新（行ズレ確認つき）
// ════════════════════════════════════════════════

const EDITABLE_CORE_FIELDS = [
  'customerName', 'siteName', 'projectName', 'summary', 'occurredDate', 'source',
  'assignee', 'quoteAmount', 'plannedProfit', 'rank', 'expectedOrderMonth',
  'currentStatus', 'pendingIssue', 'confirmedAmount', 'profit', 'profitRate'
];


// ════════════════════════════════════════════════
// 競合検知（同時編集による上書き防止）
// ════════════════════════════════════════════════

function normVal_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString();
  return (v === null || v === undefined) ? '' : String(v);
}

// 案件Noで行ズレを補正して行番号を返す（見つからなければ -1）
function resolveRowNum_(sh, rowNum, origDealNo) {
  rowNum = Number(rowNum);
  if (!origDealNo) return rowNum;
  if (String(sh.getRange(rowNum, COL.dealNo).getValue()) === String(origDealNo)) return rowNum;
  const lastRow = sh.getLastRow();
  for (let r = DATA_START_ROW; r <= lastRow; r++) {
    if (String(sh.getRange(r, COL.dealNo).getValue()) === String(origDealNo)) return r;
  }
  return -1;
}

// 自分が変更した項目(changed)について、読み込み時の値(base)と現在のシートの値が違えば競合
function coreConflicts_(sh, rowNum, changed, base) {
  const out = [];
  if (!changed || !changed.length || !base) return out;
  changed.forEach(f => {
    if (!COL[f] || base[f] === undefined) return;
    const cur = normVal_(sh.getRange(rowNum, COL[f]).getValue());
    if (cur !== normVal_(base[f])) out.push({ field: f, current: cur });
  });
  return out;
}

// 追加情報の行が、読み込み後に他の人に更新されていれば競合（1秒未満の差は誤差として無視）
function extraConflict_(vals, id, baseUpdatedAt) {
  const headers = vals[0];
  const idCol = headers.indexOf('id');
  const uCol = headers.indexOf('updatedAt');
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][idCol] == id) {
      const cur = vals[i][uCol];
      if (!cur) return null;
      const curT = new Date(cur).getTime();
      const baseT = baseUpdatedAt ? new Date(baseUpdatedAt).getTime() : 0;
      if (!baseT || Math.abs(curT - baseT) >= 1000) return { updatedAt: new Date(cur) };
      return null;
    }
  }
  return null;
}

// 保存前の競合チェック（読み取りのみ・書き込みなし）
function dealConflictCheck(data) {
  try {
    const result = { core: [], extra: null };
    if (data.branch && data.rowNum && data.changedCore && data.changedCore.length) {
      const sh = sheet(tabOf_(data));
      if (sh) {
        const r = resolveRowNum_(sh, data.rowNum, data.origDealNo);
        if (r > 0) result.core = coreConflicts_(sh, r, data.changedCore, data.baseCore);
      }
    }
    if (data.extraId && data.baseUpdatedAt !== undefined) {
      const vals = ensureExtraSheet_().getDataRange().getValues();
      result.extra = extraConflict_(vals, data.extraId, data.baseUpdatedAt);
    }
    return makeRes(result);
  } catch (err) {
    return makeErr('dealConflictCheck error: ' + err.toString());
  }
}

// ════════════════════════════════════════════════
// 目標金額（部署 × 年度 × 上期/下期）
// 年度は4月始まり。上期＝4〜9月、下期＝10〜3月。金額は円で保存。
// ════════════════════════════════════════════════

const TARGET_SHEET_NAME = '目標設定';
const TARGET_HEADERS = ['fy', 'half', 'branch', 'amount', 'updatedAt'];

function ensureTargetSheet_() {
  let sh = sheet(TARGET_SHEET_NAME);
  if (!sh) {
    sh = SS().insertSheet(TARGET_SHEET_NAME);
    sh.appendRow(TARGET_HEADERS);
  }
  return sh;
}

function targetList() {
  try {
    const vals = ensureTargetSheet_().getDataRange().getValues();
    const out = [];
    for (let i = 1; i < vals.length; i++) {
      const fy = Number(vals[i][0]);
      const half = String(vals[i][1] || '');
      const branch = String(vals[i][2] || '');
      const amount = Number(vals[i][3]);
      if (!fy || !half || !branch || isNaN(amount)) continue;
      out.push({ fy: fy, half: half, branch: branch, amount: amount });
    }
    return makeRes(out);
  } catch (err) {
    return makeErr('targetList error: ' + err.toString());
  }
}

// data: { fy: 2026, items: [{ half: '上期', branch: '１課', amount: 50000000 | '' }, ...] }
// amount が空（''）の項目は削除扱い
function targetSave(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const fy = Number(data.fy);
    if (!fy || !data.items || !data.items.length) return makeErr('年度または目標が指定されていません');
    const sh = ensureTargetSheet_();
    const vals = sh.getDataRange().getValues();
    const rowOf = {};
    for (let i = 1; i < vals.length; i++) {
      rowOf[Number(vals[i][0]) + '|' + vals[i][1] + '|' + vals[i][2]] = i + 1;
    }
    const delRows = [];
    const now = new Date();
    data.items.forEach(it => {
      if (['上期', '下期'].indexOf(it.half) < 0 || !it.branch) return;
      const key = fy + '|' + it.half + '|' + it.branch;
      const empty = (it.amount === '' || it.amount === null || it.amount === undefined);
      const r = rowOf[key];
      if (empty) {
        if (r) delRows.push(r);
        return;
      }
      const amt = Number(it.amount);
      if (isNaN(amt)) return;
      if (r) {
        sh.getRange(r, 1, 1, 5).setValues([[fy, it.half, it.branch, amt, now]]);
      } else {
        sh.appendRow([fy, it.half, it.branch, amt, now]);
      }
    });
    delRows.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
    return makeRes({ fy: fy });
  } catch (err) {
    return makeErr('targetSave error: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════
// 表記ゆれの一括統一（顧客名・現場名・源泉・担当者・仕入先）
// from に挙げた値と完全一致するセルだけを to に書き換える（値が違うセルには触れない）
// ════════════════════════════════════════════════

const RENAMEABLE_FIELDS = ['customerName', 'siteName', 'source', 'assignee'];

function valuesRename(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const field = data.field;
    const to = String(data.to === undefined || data.to === null ? '' : data.to).trim();
    const from = (data.from || []).map(String);
    if (!to || !from.length) return makeErr('統一先または統一元が指定されていません');
    const fromSet = {};
    from.forEach(v => { fromSet[v] = true; });

    let changed = 0;
    if (field === 'supplier') {
      const sh = ensureExtraSheet_();
      const vals = sh.getDataRange().getValues();
      // 機器の仕入先・部品・部材の仕入先・外注先をまとめて統一する
      const cols = ['supplier', 'partsSupplier', 'outsourceSupplier'].map(h => vals[0].indexOf(h)).filter(c => c >= 0);
      if (!cols.length) return makeErr('仕入先の列が見つかりません');
      cols.forEach(col => {
        for (let i = 1; i < vals.length; i++) {
          const v = String(vals[i][col]);
          if (fromSet[v] && v !== to) { sh.getRange(i + 1, col + 1).setValue(to); changed++; }
        }
      });
    } else if (RENAMEABLE_FIELDS.indexOf(field) >= 0) {
      sourceTabs_().forEach(t => {
        const sh = sheet(t.tab);
        if (!sh) return;
        const lastRow = sh.getLastRow();
        if (lastRow < DATA_START_ROW) return;
        const col = COL[field];
        const vals = sh.getRange(DATA_START_ROW, col, lastRow - DATA_START_ROW + 1, 1).getValues();
        vals.forEach((r, i) => {
          const v = String(r[0]);
          if (fromSet[v] && v !== to) { sh.getRange(DATA_START_ROW + i, col).setValue(to); changed++; }
        });
      });
    } else {
      return makeErr('対象外の項目です: ' + field);
    }
    return makeRes({ changed: changed });
  } catch (err) {
    return makeErr('valuesRename error: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════
// 年度シートの作成（例：2027年度）
// 各拠点の現在のシート（拠点名のタブ）を複製して「拠点名_2027」を作り、案件の入力欄だけを空にする。
// B列のNo・数式のセル・書式・列幅はそのまま。すでにあるタブは、作り直さずに飛ばす（何度押しても安全）。
// ════════════════════════════════════════════════
function createFiscalYearSheets(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const fy = Number(data.fy);
    if (!fy || fy < 2000 || fy > 2100) return makeErr('年度が正しくありません: ' + data.fy);
    const inputCols = ['dealNo'].concat(EDITABLE_CORE_FIELDS).map(f => COL[f]);
    const created = [], skipped = [], missing = [];
    SOURCE_SHEETS.forEach(base => {
      const name = base + '_' + fy;
      if (sheet(name)) { skipped.push(name); return; }
      const src = sheet(base);
      if (!src) { missing.push(base); return; }
      const copy = src.copyTo(SS());
      copy.setName(name);
      const lastRow = copy.getLastRow();
      if (lastRow >= DATA_START_ROW) {
        const n = lastRow - DATA_START_ROW + 1;
        inputCols.forEach(c => {
          const rng = copy.getRange(DATA_START_ROW, c, n, 1);
          const fs = rng.getFormulas();
          // 数式のセルは、数式のまま残す。それ以外（入力された値）は空にする
          rng.setValues(fs.map(r => [r[0] ? r[0] : '']));
        });
      }
      created.push(name);
    });
    invalidateDealCache_();
    return makeRes({ fy: fy, created: created, skipped: skipped, missing: missing });
  } catch (err) {
    return makeErr('createFiscalYearSheets error: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

function dealCoreCreate(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!data.branch || SOURCE_SHEETS.indexOf(data.branch) < 0) {
      return makeErr('拠点(branch)が不正です: ' + data.branch);
    }
    const tab = resolveCreateTab_(data);
    const sh = sheet(tab);
    if (!sh) return makeErr('シートが見つかりません: ' + tab);

    const lastRow = Math.max(sh.getLastRow(), HEADER_ROW);

    // 新規行の位置：No列にだけ番号が入っている「空きの定型行」があれば、その最初の行を使う。
    // （番号が下まで先に振ってある運用のため、getLastRow()+1 だと遠く離れた下の行に入ってしまう）
    // 空きの定型行が無ければ、従来どおり末尾に追加し、Noは既存の最大値+1を採番する。
    let newRow = -1;
    let maxNo = 0;
    if (lastRow >= DATA_START_ROW) {
      const blk = sh.getRange(DATA_START_ROW, COL.no, lastRow - DATA_START_ROW + 1, COL.projectName - COL.no + 1).getValues(); // B〜F列
      blk.forEach((r, i) => {
        const isNum = (r[0] !== '' && r[0] !== null && isFinite(Number(r[0])));
        if (isNum) maxNo = Math.max(maxNo, Number(r[0]));
        const blank = [r[1], r[2], r[3], r[4]].every(v => v === '' || v === null);
        if (newRow < 0 && isNum && blank) newRow = DATA_START_ROW + i;
      });
    }
    const reuseRow = newRow > 0;
    if (!reuseRow) newRow = lastRow + 1;

    // 案件Noは指定があればそのまま使用。未入力なら空欄のまま作成する（自動採番はしない）
    const dealNo = (data.dealNo && String(data.dealNo).trim()) ? String(data.dealNo).trim() : '';

    if (!reuseRow) sh.getRange(newRow, COL.no).setValue(maxNo + 1);
    if (dealNo) sh.getRange(newRow, COL.dealNo).setValue(dealNo);

    EDITABLE_CORE_FIELDS.forEach(f => {
      if (data[f] === undefined || data[f] === '') return;
      sh.getRange(newRow, COL[f]).setValue(data[f]);
    });

    // 案件Noが空の場合は、他の未入力行と同じ「拠点::r行番号」形式のidにする（一覧取得時と一致させる）
    const id = dealNo ? (tab + '::' + dealNo) : (tab + '::r' + newRow);

    const newNo = sh.getRange(newRow, COL.no).getValue();
    return makeRes({ id: id, branch: data.branch, sheet: tab, rowNum: newRow, dealNo: dealNo, no: newNo });
  } catch (err) {
    return makeErr('dealCoreCreate error: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

function dealCoreUpdate(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!data.branch || !data.rowNum) return makeErr('branch / rowNum が指定されていません');
    const tab = tabOf_(data);
    const sh = sheet(tab);
    if (!sh) return makeErr('シートが見つかりません: ' + tab);

    let rowNum = Number(data.rowNum);
    const origDealNo = data.origDealNo !== undefined ? data.origDealNo : data.dealNo;

    // 他の人が行を追加・削除して行がズレていないか、元の案件Noで突合確認
    if (origDealNo) {
      const curDealNo = sh.getRange(rowNum, COL.dealNo).getValue();
      if (String(curDealNo) !== String(origDealNo)) {
        const lastRow = sh.getLastRow();
        let found = -1;
        for (let r = DATA_START_ROW; r <= lastRow; r++) {
          if (String(sh.getRange(r, COL.dealNo).getValue()) === String(origDealNo)) { found = r; break; }
        }
        if (found < 0) {
          return makeErr('対象の行が見つかりませんでした（案件No: ' + origDealNo + '）。ページを再読み込みしてから、もう一度お試しください。');
        }
        rowNum = found;
      }
    }

    // 競合検知：自分が変更した項目が、読み込み後に他の人に変更されていないか確認
    if (!data.force && data.changedCore) {
      const cc = coreConflicts_(sh, rowNum, data.changedCore, data.baseCore);
      if (cc.length) return makeRes({ core: cc }, 'conflict');
    }

    EDITABLE_CORE_FIELDS.forEach(f => {
      if (data[f] === undefined) return;
      // changedCore が指定されている場合は、変更した項目だけ書き込む（他の人の編集を上書きしない）
      if (data.changedCore && data.changedCore.indexOf(f) < 0) return;
      const colIdx = COL[f];
      if (!colIdx) return;
      sh.getRange(rowNum, colIdx).setValue(data[f]);
    });

    // 案件No自体が変更された場合は、その列も更新し、追加情報シート側の紐付けIDも付け替える
    let newDealNo = origDealNo;
    if (data.dealNo !== undefined && String(data.dealNo) !== String(origDealNo)) {
      newDealNo = data.dealNo;
      sh.getRange(rowNum, COL.dealNo).setValue(newDealNo);
      if (origDealNo) {
        renameExtraId_(tab + '::' + origDealNo, tab + '::' + newDealNo);
      }
    }

    return makeRes({ id: tab + '::' + newDealNo, rowNum: rowNum, dealNo: newDealNo });
  } catch (err) {
    return makeErr('dealCoreUpdate error: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════
// 案件の削除
// 元シートの行は削除せず（行がズレるため）、案件No・基本情報の入力欄だけを空にする。
// B列のNoと、数式が入っているセルはそのまま残す＝新規作成時に「空きの定型行」として再利用される。
// 追加情報シートの行は削除する。消した内容は「削除履歴」シートに控える。
// ════════════════════════════════════════════════

const DELETE_LOG_SHEET_NAME = '削除履歴';
const DELETE_LOG_HEADERS = ['deletedAt', 'branch', 'rowNum', 'no', 'dealNo', 'customerName', 'siteName', 'projectName', 'coreJson', 'extraJson'];

function dealDelete(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tab = tabOf_(data);
    if (!tab || !isSourceTab_(tab) || !data.rowNum) {
      return makeErr('branch / rowNum が正しく指定されていません');
    }
    const sh = sheet(tab);
    if (!sh) return makeErr('シートが見つかりません: ' + tab);

    // 行ズレの補正（案件Noで突合）。案件Noが無い案件は、行番号をそのまま使う
    const rowNum = resolveRowNum_(sh, data.rowNum, data.origDealNo);
    if (rowNum < DATA_START_ROW) {
      return makeErr('対象の行が見つかりませんでした（案件No: ' + data.origDealNo + '）。ページを再読み込みしてから、もう一度お試しください。');
    }

    // 競合検知：読み込み後に、他の人がこの行を書き換えていないか
    if (!data.force && data.baseCore) {
      const keys = Object.keys(data.baseCore).filter(f => COL[f]);
      const cc = coreConflicts_(sh, rowNum, keys, data.baseCore);
      if (cc.length) return makeRes({ core: cc }, 'conflict');
    }

    // 控えを残す（行の内容と、追加情報シートの行）
    const rowVals = sh.getRange(rowNum, 1, 1, 19).getValues()[0];
    const core = {};
    Object.keys(COL).forEach(f => { core[f] = normVal_(rowVals[COL[f] - 1]); });
    if (!core.dealNo && !core.customerName && !core.siteName && !core.projectName) {
      return makeErr('この行はすでに空です');
    }

    const exSh = ensureExtraSheet_();
    const exVals = exSh.getDataRange().getValues();
    const exHeaders = exVals[0];
    const idCol = exHeaders.indexOf('id');
    const extraId = data.id || (tab + '::' + (core.dealNo ? core.dealNo : 'r' + rowNum));
    let exRowIdx = -1, extra = null;
    for (let i = 1; i < exVals.length; i++) {
      if (exVals[i][idCol] == extraId) {
        exRowIdx = i + 1;
        extra = {};
        exHeaders.forEach((h, c) => { extra[h] = normVal_(exVals[i][c]); });
        break;
      }
    }

    let log = sheet(DELETE_LOG_SHEET_NAME);
    if (!log) { log = SS().insertSheet(DELETE_LOG_SHEET_NAME); log.appendRow(DELETE_LOG_HEADERS); }
    log.appendRow([new Date(), tab, rowNum, core.no, core.dealNo, core.customerName, core.siteName, core.projectName,
      JSON.stringify(core), extra ? JSON.stringify(extra) : '']);

    // 元シート：案件No＋基本情報の入力欄を空にする（数式のセルは残す）
    ['dealNo'].concat(EDITABLE_CORE_FIELDS).forEach(f => {
      const cell = sh.getRange(rowNum, COL[f]);
      if (!cell.getFormula()) cell.clearContent();
    });

    // 追加情報シート：該当の行を削除
    if (exRowIdx > 0) exSh.deleteRow(exRowIdx);

    return makeRes({ id: extraId, rowNum: rowNum });
  } catch (err) {
    return makeErr('dealDelete error: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

// 追加情報シートの行のidを付け替える（案件No変更時に使用）
function renameExtraId_(oldId, newId) {
  const sh = ensureExtraSheet_();
  const vals = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][idCol] == oldId) {
      sh.getRange(i + 1, idCol + 1).setValue(newId);
      return true;
    }
  }
  return false;
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

    // 競合検知：読み込み後に他の人が更新していないか確認
    if (!data.force && data.baseUpdatedAt !== undefined) {
      const ec = extraConflict_(vals, data.id, data.baseUpdatedAt);
      if (ec) return makeRes({ extra: ec }, 'conflict');
    }

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

    // 受注日時の自動記録：受注系ステータスへ「切り替わった」ときだけ記録する
    // （既に受注系だった案件には後から日時を付けない＝過去案件の日数が狂わないように）
    const oCol = headers.indexOf('orderedAt');
    if (oCol >= 0) {
      const WON_STATUSES = ['受注', '工事中', '完了'];
      const sCol = headers.indexOf('status');
      let exOrdered = '', exStatus = '';
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][idCol] == data.id) { exOrdered = vals[i][oCol]; exStatus = vals[i][sCol]; break; }
      }
      const effSt = data.status !== undefined ? data.status : exStatus;
      if (WON_STATUSES.indexOf(effSt) >= 0) {
        data.orderedAt = exOrdered || (WON_STATUSES.indexOf(exStatus) >= 0 ? '' : new Date());
      } else {
        data.orderedAt = '';
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
