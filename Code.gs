/**
 * 月帳 PWA — Google Apps Script 後端
 * 部署：「以網頁應用程式部署」→ 執行身分=自己 / 存取權=任何人
 *
 * 設定：
 *   1. 將 SHEET_ID 替換為你的 Google 試算表 ID
 *   2. 服務 → 加入 Drive API v2（OCR 用）
 *
 * 試算表需有以下 sheet（不存在會自動建）：
 *   記帳明細：日期、金額、類別、說明、來源、備註、ID、建立時間
 */
const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_NAME = '記帳明細';
const TZ = 'Asia/Taipei';

const ENTRY_HEADER = ['日期', '金額', '類別', '說明', '來源', '備註', 'ID', '建立時間'];

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

// ============ 工具函式 ============

// Sheet 把日期自動轉成 Date 時用此防呆,把任何形式轉回 'yyyy-MM-dd'
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

// 依 header 名稱建立 idx map (避免 hard-code 陣列順序)
function _headerIdx_(header) {
  const m = {};
  header.forEach(function (name, i) { m[String(name).trim()] = i; });
  return m;
}

// 依 header idx 把 row array 轉成 entry 物件
function _rowToEntry_(row, idx, fallbackId) {
  return {
    id:          String(row[idx['ID']] || fallbackId),
    date:        _toDateStr_(row[idx['日期']]),
    amount:      _toNumber_(row[idx['金額']]),
    category:    String(row[idx['類別']] || '📦 其他'),
    description: String(row[idx['說明']] || ''),
    source:      String(row[idx['來源']] || 'manual'),
    note:        String(row[idx['備註']] || ''),
  };
}

function _formatDate_(date) {
  return Utilities.formatDate(date instanceof Date ? date : new Date(date), TZ, 'yyyy-MM-dd');
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
    return jsonOk(data);
  } catch (err) {
    return jsonErr(err && err.message ? err.message : String(err), 500);
  }
}

function _dispatch_(action, payload) {
  switch (action) {
    case 'ping':          return { ok: true, ts: Date.now(), version: '1.0.0', msg: '月帳 API 正常運作' };
    case 'getEntries':    return handleGetEntries_(payload);
    case 'getMonthStats': return handleGetMonthStats_(payload);
    case 'addEntry':      return handleAddEntry_(payload);
    case 'addEntries':    return handleAddEntries_(payload);
    case 'processOCR':    return handleProcessOCR_(payload);
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
  }
  return sheet;
}

// ============ Handler ============

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

  // 依 header idx 組 row,而不是 hard-code 順序
  const row = new Array(header.length).fill('');
  row[idx['日期']]     = date;
  row[idx['金額']]     = amount;
  row[idx['類別']]     = category;
  row[idx['說明']]     = entry.description || '';
  row[idx['來源']]     = entry.source || 'manual';
  row[idx['備註']]     = entry.note || '';
  row[idx['ID']]       = id;
  row[idx['建立時間']] = now;
  sheet.appendRow(row);

  return { id: id, category: category, amount: amount };
}

function handleAddEntries_(p) {
  const entries = p.entries;
  if (!Array.isArray(entries) || !entries.length) return { added: 0 };
  entries.forEach(function (e) { handleAddEntry_({ entry: e }); });
  return { added: entries.length };
}

function handleGetEntries_(p) {
  const sheet = _getSheet_();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { entries: [] };

  const idx = _headerIdx_(rows[0]);
  const start = p.startDate ? p.startDate : null;
  const end   = p.endDate   ? p.endDate   : null;

  const entries = [];
  for (let i = 1; i < rows.length; i++) {
    const dateStr = _toDateStr_(rows[i][idx['日期']]);
    if (!dateStr) continue;
    if (start && dateStr < start) continue;
    if (end   && dateStr > end)   continue;
    entries.push(_rowToEntry_(rows[i], idx, i));
  }
  entries.sort(function (a, b) { return b.date.localeCompare(a.date); });
  return { entries: entries };
}

function handleGetMonthStats_(p) {
  const sheet = _getSheet_();
  const rows = sheet.getDataRange().getValues();
  const idx = rows.length ? _headerIdx_(rows[0]) : null;

  const now = new Date();
  const year  = parseInt(p.year, 10)  || now.getFullYear();
  const month = parseInt(p.month, 10) || (now.getMonth() + 1);
  const stats = { income: 0, expense: 0, categories: {}, daily: {} };

  if (!idx || rows.length <= 1) {
    return { year: year, month: month, stats: stats };
  }

  const prefix = year + '-' + ('0' + month).slice(-2);

  for (let i = 1; i < rows.length; i++) {
    const dateStr = _toDateStr_(rows[i][idx['日期']]);
    if (!dateStr || dateStr.indexOf(prefix) !== 0) continue;

    const amount = _toNumber_(rows[i][idx['金額']]);
    const category = String(rows[i][idx['類別']] || '📦 其他');
    const day = parseInt(dateStr.substring(8, 10), 10);

    if (amount >= 0) stats.income += amount;
    else             stats.expense += Math.abs(amount);

    stats.categories[category] = (stats.categories[category] || 0) + Math.abs(amount);

    if (!stats.daily[day]) stats.daily[day] = { income: 0, expense: 0 };
    if (amount >= 0) stats.daily[day].income  += amount;
    else             stats.daily[day].expense += Math.abs(amount);
  }
  return { year: year, month: month, stats: stats };
}

function handleProcessOCR_(p) {
  const base64 = p.imageBase64;
  const mimeType = p.mimeType || 'image/jpeg';
  if (!base64) throw new Error('缺少 imageBase64');

  const decoded = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(decoded, mimeType, 'ocr_' + Date.now() + '.jpg');
  const file = Drive.Files.insert(
    { title: 'yuezhang_ocr_tmp', mimeType: 'application/vnd.google-apps.document' },
    blob,
    { ocr: true, ocrLanguage: 'zh-TW' }
  );
  const text = DocumentApp.openById(file.id).getBody().getText();
  DriveApp.getFileById(file.id).setTrashed(true);
  return { rawText: text, transactions: parseBankText_(text) };
}

// ============ 解析 / 分類 ============

function parseBankText_(text) {
  const lines = String(text || '').split('\n').map(function (l) { return l.trim(); }).filter(function (l) { return l.length > 3; });
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

// ============ 初始化 ============

function setup() {
  _getSheet_();
  SpreadsheetApp.flush();
  Logger.log('✅ 初始化完成');
}
