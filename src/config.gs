// ====== CONFIG ======
const TZ = 'Asia/Tokyo';

const SHEET = {
  SETTINGS: 'Settings',
  FORM_TEMPLATES: 'FormTemplates',
  FORM_MAP: 'FormMap',
  DEPT_APPROVERS: 'DeptApprovers',  // 部署×承認者メール
  DEPTS: '部署マスタ',
  WORKERS: '作業員マスタ',
  JOBS: '業務NOマスタ',
  ORDERS: '工番マスタ',
  BATCH_LOGS: 'BatchLogs',
};

// フォーム内の質問タイトル（テンプレフォームにこの通り配置してください）
const Q = {
  TYPE: '申請種別',
  DEPT: '部署',
  WORKER: '作業員', // 例：A001 今泉雄二
  DATE: '作業実施日',
  ORDER: '工番',
  JOB: '業務ID（業務NO）',
  REASON: '理由',
  REASON_DETAIL: '補足理由',
  OT_HOURS: '予定時間（残業）',
  HD_HOURS: '予定時間（休日）',
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

function getSettings_() {
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
  return map;
}

function requireSheet_(name) {
  const ss = getDb_();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function normalize_(s) {
  return String(s ?? '').trim();
}

function getSheetHeaderIndex_(sheetName, headerRowNo=1) {
  const sh = requireSheet_(sheetName);
  const header = sh.getRange(headerRowNo, 1, 1, sh.getLastColumn()).getValues()[0].map(h => normalize_(h));
  return { sh, header, idx: buildHeaderIndex_(header) };
}
