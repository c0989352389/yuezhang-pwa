/**
 * 月帳 PWA — Google Apps Script 後端 v1.4.0
 * 部署：「以網頁應用程式部署」→ 執行身分=自己 / 存取權=任何人
 *
 * 設定：
 *   1. 將 SHEET_ID 替換為你的 Google 試算表 ID
 *   2. 專案設定 → 指令碼屬性 → 新增 OCR_SPACE_API_KEY (註冊 https://ocr.space/ocrapi 拿 free key)
 *   3. (Drive API 已不需要,可以移除)
 *
 * 試算表需有以下 sheet（不存在會自動建）：
 *   記帳明細：日期、金額、類別、說明、來源、備註、ID、建立時間、用戶
 */
const SHEET_ID = '1LVxI70AqgM6wsZptlfI0H6WKPMiuh-tnnkLUZBSjSvA';
const SHEET_NAME = '記帳明細';
const TZ = 'Asia/Taipei';
const CACHE_TTL = 300; // 5 分鐘

const ENTRY_HEADER = ['日期', '金額', '類別', '說明', '來源', '備註', 'ID', '建立時間', '用戶', '帳戶'];

const ACCOUNT_SHEET = '帳戶設定';
const ACCOUNT_HEADER = ['名稱', '圖示', '排序', '狀態'];
const DEFAULT_ACCOUNTS = [
  ['現金', '💵', 1, 'active'],
  ['信用卡', '💳', 2, 'active'],
  ['街口', '🟢', 3, 'active'],
  ['悠遊卡', '🟦', 4, 'active'],
  ['匯款', '🏦', 5, 'active'],
];

const CATEGORY_SHEET = '類別設定';
const CATEGORY_HEADER = ['名稱', '圖示', '排序', '狀態', '類型'];
// 類型: expense / income / both — 影響在記帳頁顯示位置
const DEFAULT_CATEGORIES = [
  ['🍜 餐飲', '🍜', 1, 'active', 'expense'],
  ['🚌 交通', '🚌', 2, 'active', 'expense'],
  ['🛍️ 購物', '🛍️', 3, 'active', 'expense'],
  ['🎬 娛樂', '🎬', 4, 'active', 'expense'],
  ['💊 醫療', '💊', 5, 'active', 'expense'],
  ['🏠 住家', '🏠', 6, 'active', 'expense'],
  ['📱 通訊', '📱', 7, 'active', 'expense'],
  ['💰 收入', '💰', 8, 'active', 'income'],
  ['📊 投資', '📊', 9, 'active', 'both'],
  ['📦 其他', '📦', 10, 'active', 'both'],
];

const CATEGORY_MAP = {
  '🍜 餐飲': ['早餐','午餐','晚餐','飲料','咖啡','奶茶','珍奶','吃','食','餐','麵','飯','便當','外送','火鍋','燒烤','壽司','拉麵','麥當勞','肯德基','星巴克','starbucks','foodpanda','ubereats'],
  '🚌 交通': ['捷運','mrt','公車','計程車','taxi','uber','停車','加油','油費','youbike','ubike','火車','高鐵','台鐵','機票'],
  '🛍️ 購物': ['買','超商','7-11','全家','全聯','超市','好市多','costco','大潤發','衣服','鞋子','包包','網購','momo','pchome','蝦皮'],
  '🎬 娛樂': ['電影','遊戲','ktv','旅遊','旅行','門票','展覽','課程','訂閱','netflix','spotify','健身'],
  '💊 醫療': ['藥','診所','醫院','掛號','健保','醫療','牙醫','眼科','健檢'],
  '🏠 住家': ['水費','電費','瓦斯','房租','管理費','修繕','家具','家電','清潔用品'],
  '📱 通訊': ['電話費','網路費','手機費','wifi','電信','中華電信','遠傳','台哥大'],
  '💰 收入': ['薪資','薪水','工資','獎金','收入','轉帳收','匯款','退款','退稅'],
  '📊 投資': ['股票','基金','投資','理財','保險費','etf'],
};

// ============ 快取 ============

function _cacheKey_(suffix) {
  return 'yuezhang_' + suffix + '_' + _cacheVer_();
}

function _cacheVer_() {
  const cache = CacheService.getScriptCache();
  let v = cache.get('yuezhang_ver');
  if (!v) {
    v = String(Date.now());
    cache.put('yuezhang_ver', v, CACHE_TTL);
  }
  return v;
}

function _bumpCache_() {
  CacheService.getScriptCache().put('yuezhang_ver', String(Date.now()), CACHE_TTL);
}

function _cacheGet_(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (_) { return null; }
}

function _cachePut_(key, val) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(val), CACHE_TTL);
  } catch (_) {}
}

// 一次讀整張表並快取（多個 handler 共用,大幅減少 sheet 讀取）
function _getAllRows_() {
  const cacheKey = _cacheKey_('rows');
  const cached = _cacheGet_(cacheKey);
  if (cached) return cached;

  const sheet = _getSheet_();
  const rows = sheet.getDataRange().getValues();
  const result = { header: rows[0] || [], data: rows.slice(1) };
  _cachePut_(cacheKey, result);
  return result;
}

// ============ 工具函式 ============

function _toDateStr_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{1,2}-\d{1,2}/.test(s)) {
    const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  }
  return '';
}

function _toNumber_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function _headerIdx_(header) {
  const m = {};
  header.forEach(function (name, i) { m[String(name).trim()] = i; });
  return m;
}

function _rowToEntry_(row, idx, fallbackId) {
  return {
    id:          String(row[idx['ID']] || fallbackId),
    date:        _toDateStr_(row[idx['日期']]),
    amount:      _toNumber_(row[idx['金額']]),
    category:    String(row[idx['類別']] || '📦 其他'),
    description: String(row[idx['說明']] || ''),
    source:      String(row[idx['來源']] || 'manual'),
    note:        String(row[idx['備註']] || ''),
    user:        String(row[idx['用戶']] != null ? row[idx['用戶']] : ''),
    account:     String(row[idx['帳戶']] != null ? row[idx['帳戶']] : ''),
  };
}

function _formatDate_(date) {
  return Utilities.formatDate(date instanceof Date ? date : new Date(date), TZ, 'yyyy-MM-dd');
}

function _today_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

// ============ HTTP 入口 ============

function doGet(e) {
  return route_(e, (e.parameter || {}));
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return jsonErr('JSON 格式錯誤');
  }
  return route_(e, body);
}

function route_(e, payload) {
  try {
    const action = payload.action;
    if (!action) return jsonErr('缺少 action');
    const data = _dispatch_(action, payload);
    _maybeInvalidate_(action);
    return jsonOk(data);
  } catch (err) {
    return jsonErr(err && err.message ? err.message : String(err), 500);
  }
}

const _WRITE_ACTIONS_ = { addEntry: 1, addEntries: 1, updateEntry: 1, deleteEntry: 1, processOCR: 0, saveAccount: 1, deleteAccount: 1, saveCategory: 1, deleteCategory: 1 };

function _maybeInvalidate_(action) {
  if (_WRITE_ACTIONS_[action]) _bumpCache_();
}

function _dispatch_(action, payload) {
  switch (action) {
    case 'ping':          return { ok: true, ts: Date.now(), version: '1.2.0', msg: '月帳 API 正常運作' };
    case 'getHomeData':   return handleGetHomeData_(payload);
    case 'getEntries':    return handleGetEntries_(payload);
    case 'getMonthStats': return handleGetMonthStats_(payload);
    case 'addEntry':      return handleAddEntry_(payload);
    case 'addEntries':    return handleAddEntries_(payload);
    case 'updateEntry':   return handleUpdateEntry_(payload);
    case 'deleteEntry':   return handleDeleteEntry_(payload);
    case 'processOCR':    return handleProcessOCR_(payload);
    case 'getAccounts':   return handleGetAccounts_(payload);
    case 'saveAccount':   return handleSaveAccount_(payload);
    case 'deleteAccount': return handleDeleteAccount_(payload);
    case 'getCategories': return handleGetCategories_(payload);
    case 'saveCategory':  return handleSaveCategory_(payload);
    case 'deleteCategory':return handleDeleteCategory_(payload);
    default: throw new Error('未知 action: ' + action);
  }
}

// ============ 回應 ============

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(message, code) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: message, code: code || 400 }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============ Sheet 存取 ============

function _getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(ENTRY_HEADER);
    sheet.getRange(1, 1, 1, ENTRY_HEADER.length)
      .setFontWeight('bold')
      .setBackground('#0d1525')
      .setFontColor('#4fd9b3');
    sheet.setFrozenRows(1);
    return sheet;
  }
  const lastCol = sheet.getLastColumn();
  const header = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const have = {};
  header.forEach(function (h) { have[String(h).trim()] = true; });
  ENTRY_HEADER.forEach(function (col) {
    if (!have[col]) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col)
        .setFontWeight('bold').setBackground('#0d1525').setFontColor('#4fd9b3');
    }
  });
  return sheet;
}

// ============ Handler ============

// 一次回首頁需要的全部資料(今日明細+本月統計),省一次冷啟動
function handleGetHomeData_(p) {
  const userFilter = p.user ? String(p.user).trim() : '';
  const all = _getAllRows_();
  const idx = _headerIdx_(all.header);
  const userColIdx = idx['用戶'];

  const now = new Date();
  const year  = parseInt(p.year, 10)  || now.getFullYear();
  const month = parseInt(p.month, 10) || (now.getMonth() + 1);
  const today = _today_();
  const monthPrefix = year + '-' + ('0' + month).slice(-2);

  const todayEntries = [];
  const stats = { income: 0, expense: 0, categories: {}, daily: {} };

  for (let i = 0; i < all.data.length; i++) {
    const row = all.data[i];
    const dateStr = _toDateStr_(row[idx['日期']]);
    if (!dateStr) continue;

    if (userFilter && userColIdx != null) {
      const rowUser = String(row[userColIdx] || '').trim();
      if (rowUser !== userFilter) continue;
    }

    // 本月統計
    if (dateStr.indexOf(monthPrefix) === 0) {
      const amount = _toNumber_(row[idx['金額']]);
      const category = String(row[idx['類別']] || '📦 其他');
      const day = parseInt(dateStr.substring(8, 10), 10);

      if (amount >= 0) stats.income += amount;
      else             stats.expense += Math.abs(amount);

      stats.categories[category] = (stats.categories[category] || 0) + Math.abs(amount);

      if (!stats.daily[day]) stats.daily[day] = { income: 0, expense: 0 };
      if (amount >= 0) stats.daily[day].income  += amount;
      else             stats.daily[day].expense += Math.abs(amount);
    }

    // 今日明細
    if (dateStr === today) {
      todayEntries.push(_rowToEntry_(row, idx, i + 1));
    }
  }

  todayEntries.sort(function (a, b) { return b.date.localeCompare(a.date); });
  return {
    today: today,
    entries: todayEntries,
    year: year,
    month: month,
    stats: stats,
    user: userFilter,
  };
}

function handleAddEntry_(p) {
  const entry = p.entry;
  if (!entry) throw new Error('entry 不可為空');

  const sheet = _getSheet_();
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = _headerIdx_(header);

  const now = new Date();
  const id = Utilities.getUuid().substring(0, 8);
  const date = entry.date || _formatDate_(now);
  const amount = _toNumber_(entry.amount);
  const category = entry.category || classifyText_(entry.description || '');
  const user = String(p.user || entry.user || '').trim();

  const row = new Array(header.length).fill('');
  row[idx['日期']]     = date;
  row[idx['金額']]     = amount;
  row[idx['類別']]     = category;
  row[idx['說明']]     = entry.description || '';
  row[idx['來源']]     = entry.source || 'manual';
  row[idx['備註']]     = entry.note || '';
  row[idx['ID']]       = id;
  row[idx['建立時間']] = now;
  if (idx['用戶'] != null) row[idx['用戶']] = user;
  const account = String(p.account || entry.account || '').trim();
  if (idx['帳戶'] != null) row[idx['帳戶']] = account;
  sheet.appendRow(row);

  return { id: id, category: category, amount: amount, user: user, account: account };
}

function handleUpdateEntry_(p) {
  const id = String(p.id || (p.entry && p.entry.id) || '').trim();
  if (!id) throw new Error('缺少 id');
  const patch = p.entry || p.patch || {};

  const sheet = _getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('沒有資料');

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = _headerIdx_(header);
  if (idx['ID'] == null) throw new Error('Sheet 缺少 ID 欄');

  const idCol = idx['ID'] + 1;
  const idValues = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  let rowNum = -1;
  for (let i = 0; i < idValues.length; i++) {
    if (String(idValues[i][0]).trim() === id) { rowNum = i + 2; break; }
  }
  if (rowNum < 0) throw new Error('找不到 ID: ' + id);

  // 依 patch 寫入指定欄位 (依 header idx, 不 hard-code 順序)
  const fieldMap = {
    date: '日期', amount: '金額', category: '類別', description: '說明',
    source: '來源', note: '備註', user: '用戶', account: '帳戶',
  };
  Object.keys(fieldMap).forEach(function (k) {
    if (patch[k] !== undefined && idx[fieldMap[k]] != null) {
      let v = patch[k];
      if (k === 'amount') v = _toNumber_(v);
      sheet.getRange(rowNum, idx[fieldMap[k]] + 1).setValue(v);
    }
  });

  return { id: id, updated: true };
}

function handleDeleteEntry_(p) {
  const id = String(p.id || '').trim();
  if (!id) throw new Error('缺少 id');

  const sheet = _getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('沒有資料');

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = _headerIdx_(header);
  if (idx['ID'] == null) throw new Error('Sheet 缺少 ID 欄');

  const idCol = idx['ID'] + 1;
  const idValues = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < idValues.length; i++) {
    if (String(idValues[i][0]).trim() === id) {
      sheet.deleteRow(i + 2);
      return { id: id, deleted: true };
    }
  }
  throw new Error('找不到 ID: ' + id);
}

function handleAddEntries_(p) {
  const entries = p.entries;
  if (!Array.isArray(entries) || !entries.length) return { added: 0 };
  const user = p.user || '';
  entries.forEach(function (e) { handleAddEntry_({ entry: e, user: user }); });
  return { added: entries.length };
}

function handleGetEntries_(p) {
  const all = _getAllRows_();
  if (!all.data.length) return { entries: [] };

  const idx = _headerIdx_(all.header);
  const start = p.startDate ? p.startDate : null;
  const end   = p.endDate   ? p.endDate   : null;
  const userFilter = p.user ? String(p.user).trim() : '';

  const entries = [];
  for (let i = 0; i < all.data.length; i++) {
    const dateStr = _toDateStr_(all.data[i][idx['日期']]);
    if (!dateStr) continue;
    if (start && dateStr < start) continue;
    if (end   && dateStr > end)   continue;
    const entry = _rowToEntry_(all.data[i], idx, i + 1);
    if (userFilter && entry.user !== userFilter) continue;
    entries.push(entry);
  }
  entries.sort(function (a, b) { return b.date.localeCompare(a.date); });
  return { entries: entries };
}

function handleGetMonthStats_(p) {
  const all = _getAllRows_();
  const idx = all.data.length ? _headerIdx_(all.header) : null;

  const now = new Date();
  const year  = parseInt(p.year, 10)  || now.getFullYear();
  const month = parseInt(p.month, 10) || (now.getMonth() + 1);
  const userFilter = p.user ? String(p.user).trim() : '';
  const stats = { income: 0, expense: 0, categories: {}, daily: {} };

  if (!idx) return { year: year, month: month, stats: stats, user: userFilter };

  const prefix = year + '-' + ('0' + month).slice(-2);
  const userColIdx = idx['用戶'];

  for (let i = 0; i < all.data.length; i++) {
    const row = all.data[i];
    const dateStr = _toDateStr_(row[idx['日期']]);
    if (!dateStr || dateStr.indexOf(prefix) !== 0) continue;

    if (userFilter) {
      const rowUser = userColIdx != null ? String(row[userColIdx] || '').trim() : '';
      if (rowUser !== userFilter) continue;
    }

    const amount = _toNumber_(row[idx['金額']]);
    const category = String(row[idx['類別']] || '📦 其他');
    const day = parseInt(dateStr.substring(8, 10), 10);

    if (amount >= 0) stats.income += amount;
    else             stats.expense += Math.abs(amount);

    stats.categories[category] = (stats.categories[category] || 0) + Math.abs(amount);

    if (!stats.daily[day]) stats.daily[day] = { income: 0, expense: 0 };
    if (amount >= 0) stats.daily[day].income  += amount;
    else             stats.daily[day].expense += Math.abs(amount);
  }
  return { year: year, month: month, stats: stats, user: userFilter };
}

// OCR 改用 OCR.space API (Google Drive OCR 上傳已停用)
const OCR_API_URL = 'https://api.ocr.space/parse/image';

function handleProcessOCR_(p) {
  const base64 = p.imageBase64;
  if (!base64) throw new Error('缺少 imageBase64');

  const apiKey = PropertiesService.getScriptProperties().getProperty('OCR_SPACE_API_KEY');
  if (!apiKey) throw new Error('未設定 OCR_SPACE_API_KEY,請至 GAS 專案設定 → 指令碼屬性新增');

  const payload = {
    apikey: apiKey,
    base64Image: 'data:image/jpeg;base64,' + base64,
    language: 'cht',     // 繁體中文
    OCREngine: '1',      // engine 1 支援 cht
    scale: 'true',
    isTable: 'false',
    detectOrientation: 'true',
  };

  const res = UrlFetchApp.fetch(OCR_API_URL, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) {
    throw new Error('OCR API HTTP ' + code + ': ' + body.substring(0, 200));
  }

  let json;
  try { json = JSON.parse(body); } catch (e) { throw new Error('OCR 回應非 JSON: ' + body.substring(0, 200)); }

  if (json.IsErroredOnProcessing) {
    const msg = json.ErrorMessage;
    const errStr = (msg && msg.join) ? msg.join('; ') : (msg || '未知錯誤');
    throw new Error('OCR 解析失敗: ' + errStr);
  }

  const text = (json.ParsedResults && json.ParsedResults[0] && json.ParsedResults[0].ParsedText) || '';
  return { rawText: text, transactions: parseBankText_(text) };
}

// 偵錯用:確認 OCR_SPACE_API_KEY 是否設好
function testOcrKey() {
  const k = PropertiesService.getScriptProperties().getProperty('OCR_SPACE_API_KEY');
  if (!k) { Logger.log('❌ OCR_SPACE_API_KEY 未設定'); return; }
  Logger.log('✅ API key 已設定,前 4 碼: ' + k.substring(0, 4) + '...');
}

// ============ 解析 / 分類 ============

function parseBankText_(text) {
  const lines = String(text || '').split('\n').map(function (l) { return l.trim(); }).filter(function (l) { return l.length > 1; });
  const bank = _parseBankFormat_(lines);
  if (bank.length) return bank;
  // 找不到銀行對帳單格式 → 試試收據/發票
  const r = _parseReceipt_(lines);
  return r ? [r] : [];
}

function _parseBankFormat_(lines) {
  const results = [];
  const P1 = /^(\d{1,2})[\/\-](\d{1,2})\s+(.+?)\s+([\+\-]?\d[\d,]*(?:\.\d{1,2})?)$/;
  const P2 = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s+(.+?)\s+([\+\-]?\d[\d,]*(?:\.\d{1,2})?)$/;
  const thisYear = new Date().getFullYear();

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    let m = line.match(P2);
    if (m) {
      const amt = parseFloat(m[5].replace(/,/g, ''));
      if (!isNaN(amt) && Math.abs(amt) >= 1) {
        results.push({
          date: m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2),
          amount: amt,
          description: m[4].trim(),
          category: classifyText_(m[4]),
          source: 'ocr',
        });
      }
      continue;
    }
    m = line.match(P1);
    if (m) {
      const amt = parseFloat(m[4].replace(/,/g, ''));
      if (!isNaN(amt) && Math.abs(amt) >= 1) {
        results.push({
          date: thisYear + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2),
          amount: amt,
          description: m[3].trim(),
          category: classifyText_(m[3]),
          source: 'ocr',
        });
      }
    }
  }
  return results;
}

// 收據/發票模式:整張單抓 1 筆支出
function _parseReceipt_(lines) {
  // 跳過這些「不是金額」的行(發票號/隨機碼/賣方/買方/統編/條碼)
  const SKIP_LINE = /(隨機碼|賣方|買方|統一編號|統編|發票號|BR-|單號|門店|分機|電話|地址)/;
  // 排除年份(1900-2099) — 4 碼數字無逗號小數點且在年份範圍
  const isYearLike = function (raw) { return /^(19|20)\d{2}$/.test(raw); };
  // 從一行抓所有合理金額候選
  const extractAmounts = function (rawLine) {
    // 清理:中文逗號→英文逗號,且消除「數字+空白+數字」之間的空白(OCR 常把 1,280 讀成 1, 280)
    const line = String(rawLine || '')
      .replace(/，/g, ',')
      .replace(/(\d)[ \t]+(?=[,\d])/g, '$1')
      .replace(/(,)[ \t]+(?=\d)/g, '$1');
    // 優先抓帶千分位的數字(1,280 / 12,345),再退而求其次抓純數字
    const m = line.match(/(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)/g);
    if (!m) return [];
    return m.map(function (s) {
      const raw = s.replace(/,/g, '');
      const hasComma = s.indexOf(',') !== -1;
      const hasDecimal = s.indexOf('.') !== -1;
      const n = parseFloat(raw);
      return { raw: raw, n: n, hasComma: hasComma, hasDecimal: hasDecimal, str: s };
    }).filter(function (o) {
      if (isNaN(o.n)) return false;
      if (o.raw.length > 7) return false; // 太長(發票號/隨機碼)
      if (o.n < 1 || o.n >= 1000000) return false;
      if (isYearLike(o.raw) && !o.hasComma && !o.hasDecimal) return false; // 跳過年份
      return true;
    });
  };

  // 1. 找金額 — 優先「總計/合計/小計/應收/金額」附近的數字
  const TOTAL_KW = /(總\s*計|合\s*計|小\s*計|應\s*收|應\s*付|金\s*額|TOTAL|TOTAL\s*AMOUNT|消費金額)/i;
  let amount = 0;
  for (let i = 0; i < lines.length; i++) {
    if (SKIP_LINE.test(lines[i])) continue;
    if (TOTAL_KW.test(lines[i])) {
      // 同行優先 — 取最大候選
      let cands = extractAmounts(lines[i]);
      // 若同行沒有,看下一行
      if (!cands.length && i + 1 < lines.length) cands = extractAmounts(lines[i + 1]);
      if (cands.length) {
        amount = cands.reduce(function (max, c) { return c.n > max ? c.n : max; }, 0);
        break;
      }
    }
  }

  // 2. 兜底 — 找最大金額(優先有逗號的,因為金額常寫 1,280)
  if (!amount) {
    let bestComma = 0, bestPlain = 0;
    for (let i = 0; i < lines.length; i++) {
      if (SKIP_LINE.test(lines[i])) continue;
      const cands = extractAmounts(lines[i]);
      cands.forEach(function (c) {
        if (c.hasComma && c.n > bestComma) bestComma = c.n;
        if (!c.hasComma && c.n > bestPlain) bestPlain = c.n;
      });
    }
    amount = bestComma || bestPlain;
  }
  if (!amount) return null;

  // 3. 找日期 — 西元 yyyy/mm/dd 或 yyyy-mm-dd 或 民國 yyy年mm月dd日
  let dateStr = _today_();
  for (let i = 0; i < lines.length; i++) {
    let m = lines[i].match(/(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
    if (m) {
      dateStr = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
      break;
    }
    m = lines[i].match(/(\d{2,3})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
    if (m) {
      const y = parseInt(m[1], 10) + (parseInt(m[1], 10) > 200 ? 0 : 1911);
      dateStr = y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
      break;
    }
  }

  // 4. 找商家名 — 第一行有中文且沒大量數字的行
  let merchant = '';
  for (let i = 0; i < Math.min(lines.length, 5); i++) {
    const l = lines[i];
    if (/[一-龥]{2,}/.test(l) && (l.match(/\d/g) || []).length < 4) {
      merchant = l.replace(/[\s\-_]+$/, '').substring(0, 30);
      break;
    }
  }
  if (!merchant) merchant = '收據';

  return {
    date: dateStr,
    amount: -Math.abs(amount), // 收據預設為支出
    description: merchant,
    category: classifyText_(merchant),
    source: 'ocr',
  };
}

function classifyText_(text) {
  if (!text) return '📦 其他';
  const lo = String(text).toLowerCase();
  const cats = Object.keys(CATEGORY_MAP);
  for (let i = 0; i < cats.length; i++) {
    const kws = CATEGORY_MAP[cats[i]];
    for (let j = 0; j < kws.length; j++) {
      if (lo.indexOf(kws[j]) !== -1) return cats[i];
    }
  }
  return '📦 其他';
}

// ============ 帳戶/類別設定 sheet ============

function _ensureSettingSheet_(name, header, defaultRows) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(header);
    sheet.getRange(1, 1, 1, header.length)
      .setFontWeight('bold').setBackground('#0d1525').setFontColor('#4fd9b3');
    sheet.setFrozenRows(1);
    if (defaultRows && defaultRows.length) {
      sheet.getRange(2, 1, defaultRows.length, header.length).setValues(defaultRows);
    }
  }
  return sheet;
}

function _readSettingRows_(sheetName, header, defaults) {
  const sheet = _ensureSettingSheet_(sheetName, header, defaults);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const idx = _headerIdx_(data[0]);
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[idx['名稱']]) continue;
    const o = {};
    Object.keys(idx).forEach(function (k) { o[k] = r[idx[k]]; });
    rows.push(o);
  }
  rows.sort(function (a, b) {
    const sa = Number(a['排序']) || 0, sb = Number(b['排序']) || 0;
    return sa - sb;
  });
  return rows;
}

function handleGetAccounts_(p) {
  const rows = _readSettingRows_(ACCOUNT_SHEET, ACCOUNT_HEADER, DEFAULT_ACCOUNTS);
  const includeInactive = !!p.includeInactive;
  const list = rows.filter(function (r) { return includeInactive || r['狀態'] !== 'inactive'; })
    .map(function (r) {
      return { name: String(r['名稱']), icon: String(r['圖示'] || ''), sort: Number(r['排序']) || 0, status: String(r['狀態'] || 'active') };
    });
  return { accounts: list };
}

function handleSaveAccount_(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('缺少 name');
  const sheet = _ensureSettingSheet_(ACCOUNT_SHEET, ACCOUNT_HEADER, DEFAULT_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  const idx = _headerIdx_(data[0]);
  const oldName = String(p.oldName || '').trim();
  // 編輯現有
  if (oldName) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx['名稱']]).trim() === oldName) {
        sheet.getRange(i + 1, idx['名稱']  + 1).setValue(name);
        if (p.icon !== undefined) sheet.getRange(i + 1, idx['圖示'] + 1).setValue(p.icon);
        if (p.sort !== undefined) sheet.getRange(i + 1, idx['排序'] + 1).setValue(Number(p.sort) || 0);
        if (p.status !== undefined) sheet.getRange(i + 1, idx['狀態'] + 1).setValue(p.status);
        return { name: name, updated: true };
      }
    }
    throw new Error('找不到帳戶: ' + oldName);
  }
  // 新增
  // 檢查重名
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx['名稱']]).trim() === name) throw new Error('帳戶名稱已存在: ' + name);
  }
  const row = new Array(data[0].length).fill('');
  row[idx['名稱']] = name;
  row[idx['圖示']] = p.icon || '';
  row[idx['排序']] = Number(p.sort) || (data.length);
  row[idx['狀態']] = p.status || 'active';
  sheet.appendRow(row);
  return { name: name, created: true };
}

function handleDeleteAccount_(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('缺少 name');
  const sheet = _ensureSettingSheet_(ACCOUNT_SHEET, ACCOUNT_HEADER, DEFAULT_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  const idx = _headerIdx_(data[0]);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx['名稱']]).trim() === name) {
      sheet.deleteRow(i + 1);
      return { name: name, deleted: true };
    }
  }
  throw new Error('找不到帳戶: ' + name);
}

function handleGetCategories_(p) {
  const rows = _readSettingRows_(CATEGORY_SHEET, CATEGORY_HEADER, DEFAULT_CATEGORIES);
  const includeInactive = !!p.includeInactive;
  const list = rows.filter(function (r) { return includeInactive || r['狀態'] !== 'inactive'; })
    .map(function (r) {
      return { name: String(r['名稱']), icon: String(r['圖示'] || ''), sort: Number(r['排序']) || 0, status: String(r['狀態'] || 'active'), type: String(r['類型'] || 'expense') };
    });
  return { categories: list };
}

function handleSaveCategory_(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('缺少 name');
  const sheet = _ensureSettingSheet_(CATEGORY_SHEET, CATEGORY_HEADER, DEFAULT_CATEGORIES);
  const data = sheet.getDataRange().getValues();
  const idx = _headerIdx_(data[0]);
  const oldName = String(p.oldName || '').trim();
  if (oldName) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx['名稱']]).trim() === oldName) {
        sheet.getRange(i + 1, idx['名稱']  + 1).setValue(name);
        if (p.icon !== undefined) sheet.getRange(i + 1, idx['圖示'] + 1).setValue(p.icon);
        if (p.sort !== undefined) sheet.getRange(i + 1, idx['排序'] + 1).setValue(Number(p.sort) || 0);
        if (p.status !== undefined) sheet.getRange(i + 1, idx['狀態'] + 1).setValue(p.status);
        if (p.type !== undefined && idx['類型'] != null) sheet.getRange(i + 1, idx['類型'] + 1).setValue(p.type);
        return { name: name, updated: true };
      }
    }
    throw new Error('找不到類別: ' + oldName);
  }
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx['名稱']]).trim() === name) throw new Error('類別名稱已存在: ' + name);
  }
  const row = new Array(data[0].length).fill('');
  row[idx['名稱']] = name;
  row[idx['圖示']] = p.icon || '';
  row[idx['排序']] = Number(p.sort) || (data.length);
  row[idx['狀態']] = p.status || 'active';
  if (idx['類型'] != null) row[idx['類型']] = p.type || 'expense';
  sheet.appendRow(row);
  return { name: name, created: true };
}

function handleDeleteCategory_(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('缺少 name');
  const sheet = _ensureSettingSheet_(CATEGORY_SHEET, CATEGORY_HEADER, DEFAULT_CATEGORIES);
  const data = sheet.getDataRange().getValues();
  const idx = _headerIdx_(data[0]);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx['名稱']]).trim() === name) {
      sheet.deleteRow(i + 1);
      return { name: name, deleted: true };
    }
  }
  throw new Error('找不到類別: ' + name);
}

// ============ 初始化 ============

function setup() {
  _getSheet_();
  _ensureSettingSheet_(ACCOUNT_SHEET, ACCOUNT_HEADER, DEFAULT_ACCOUNTS);
  _ensureSettingSheet_(CATEGORY_SHEET, CATEGORY_HEADER, DEFAULT_CATEGORIES);
  SpreadsheetApp.flush();
  Logger.log('✅ 初始化完成');
}
