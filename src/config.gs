// ====== CONFIG ======
const TZ = 'Asia/Tokyo';

const SHEET = {
  SETTINGS: 'Settings',
  FORM_TEMPLATES: 'FormTemplates',
  FORM_MAP: 'FormMap',
  DEPT_APPROVERS: 'DeptApprovers',  // 部署×承認者メール
  WORKERS: '作業員マスタ',
  JOBS: '業務NOマスタ',
  ORDERS: '工番マスタ',
  BATCH_LOGS: 'BatchLogs',
};

// フォーム内の質問タイトル（両テンプレフォームで統一すること）
const Q = {
  TYPE: '申請種別',
  DEPT: '部署',
  WORKER: '作業員',
  DATE: '作業実施日',
  ORDER: '工番',  // 旧テンプレ互換（削除対象の検出用に残す）
  // 工番：モーダルで選択した工番コードをそのままプリフィル（最大3件）
  ORDER_1: '工番1',
  ORDER_2: '工番2',
  ORDER_3: '工番3',
  // 旧形式の検出・削除用（プレフィックス＋番号分離方式）
  _OLD_ORDER_PREFIX_1: '工番プレフィックス_1',
  _OLD_ORDER_NUMBER_1: '工番番号_1',
  _OLD_ORDER_PREFIX_2: '工番プレフィックス_2',
  _OLD_ORDER_NUMBER_2: '工番番号_2',
  _OLD_ORDER_PREFIX_3: '工番プレフィックス_3',
  _OLD_ORDER_NUMBER_3: '工番番号_3',
  WORK_CONTENT: '業務内容',
  REASON: '理由',
  REASON_DETAIL: '補足理由',
  OT_HOURS: '予定時間',
  HD_HOURS: '予定時間',
};

const REASONS = [
  '急ぎ作業を要するため',
  '納期遅延解消のため',
  'マシントラブル対応のため',
  '業者対応のため',
  '顧客対応のため',
  'その他（→補足理由必須）',
];

const OT_HOURS = ['0.5','1.0','1.5','2.0','2.5','3.0','3.5','4.0'];
const HD_HOURS = ['半日','1日'];

// ====== UTIL ======
function fmtDate_(d, pattern='yyyy-MM-dd') {
  return Utilities.formatDate(d, TZ, pattern);
}

function getDb_() {
  // このスクリプトがDBスプレッドシートに紐づく前提なら SpreadsheetApp.getActive()
  return SpreadsheetApp.getActive();
}

var _settingsCache = null;

function getSettings_() {
  if (_settingsCache) return _settingsCache;
  const ss = getDb_();
  const sh = ss.getSheetByName(SHEET.SETTINGS);
  const values = sh.getDataRange().getValues();
  // 想定：1行目説明、2行目ヘッダ、3行目以降 data (key,value,memo)
  const map = {};
  for (let r = 2; r < values.length; r++) {
    const key = String(values[r][0] || '').trim();
    const val = values[r][1];
    if (key) map[key] = val;
  }
  _settingsCache = map;
  return map;
}

function requireSheet_(name) {
  const ss = getDb_();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function normalize_(s) {
  return String(s ?? '').trim()
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, c =>
      String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
}

var _headerCache = {};

function getSheetHeaderIndex_(sheetName, headerRowNo=1) {
  const sh = requireSheet_(sheetName);
  const cacheKey = sheetName + '_' + headerRowNo;
  if (_headerCache[cacheKey]) {
    return { sh, header: _headerCache[cacheKey].header, idx: _headerCache[cacheKey].idx };
  }
  const header = sh.getRange(headerRowNo, 1, 1, sh.getLastColumn()).getValues()[0].map(h => normalize_(h));
  const idx = buildHeaderIndex_(header);
  _headerCache[cacheKey] = { header, idx };
  return { sh, header, idx };
}
