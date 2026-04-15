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
  STAMP_MAP: 'StampMap',        // メール→印鑑画像FileIDマッピング
  COMPANY_CALENDAR: '社内カレンダーマスタ', // 年度カレンダー（休日/祝日/出勤日）
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

// 休憩スケジュール（時刻ベース）
const BREAK_WINDOWS = [
  { startH: 10, startM:  0, endH: 10, endM: 10 },  // 10時休憩
  { startH: 12, startM: 15, endH: 13, endM:  0 },  // 昼休憩
  { startH: 15, startM:  0, endH: 15, endM: 10 },  // 15時休憩
  { startH: 17, startM: 10, endH: 17, endM: 20 },  // 残業1回目休憩
  { startH: 19, startM: 20, endH: 19, endM: 30 },  // 残業2回目休憩
  { startH: 21, startM: 30, endH: 21, endM: 40 },  // 残業3回目休憩
];

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

// ====== 過去日遡及申請ヘルパー ======
// 年度開始は 3/16。当該日が3/16以降ならその年の3/16、以前なら前年の3/16を返す。

function getFiscalYearStartDate_(dateObj) {
  // JSTで年月日を取り出す
  var y = Number(fmtDate_(dateObj, 'yyyy'));
  var m = Number(fmtDate_(dateObj, 'MM'));
  var d = Number(fmtDate_(dateObj, 'dd'));
  // 3/16以降 → その年の3/16、それ以前 → 前年の3/16
  var startYear = (m > 3 || (m === 3 && d >= 16)) ? y : (y - 1);
  // 年度開始日（JST 00:00）
  return new Date(startYear + '-03-16T00:00:00+09:00');
}

/**
 * targetDateStr('yyyy-MM-dd') が現在の年度内かつ今日以前か判定。
 * @return {{ok: boolean, error: string}}
 */
function isWithinCurrentFiscalYear_(targetDateStr) {
  if (!targetDateStr || !/^\d{4}-\d{2}-\d{2}$/.test(String(targetDateStr))) {
    return { ok: false, error: '作業実施日の形式が不正です（yyyy-MM-dd）: ' + targetDateStr };
  }
  var target = new Date(targetDateStr + 'T00:00:00+09:00');
  if (isNaN(target.getTime())) {
    return { ok: false, error: '作業実施日が無効です: ' + targetDateStr };
  }
  var now = new Date();
  var todayStr = fmtDate_(now, 'yyyy-MM-dd');
  if (targetDateStr > todayStr) {
    return { ok: false, error: '未来日は申請できません: ' + targetDateStr };
  }
  var fyStart = getFiscalYearStartDate_(now);
  var fyStartStr = fmtDate_(fyStart, 'yyyy-MM-dd');
  if (targetDateStr < fyStartStr) {
    return { ok: false, error: '年度開始日（' + fyStartStr + '）より前の日付は申請できません。' };
  }
  return { ok: true, error: '' };
}

/**
 * 休日出勤の作業実施日バリデーション。
 * - 通常申請（isRetroactive=false）: 年度開始日（3/16）以降であれば、今日・未来日ともに許可。
 *   休日出勤は今週末・来週末への先出し申請が通常運用のため。
 * - 遡及申請（isRetroactive=true）: 年度開始日（3/16）以降で、かつ今日以前の日付のみ許可。
 * どちらの場合も年度開始日より前は拒否。
 * @param {string} targetDateStr 'yyyy-MM-dd'
 * @param {boolean} isRetroactive
 * @return {{ok: boolean, error: string}}
 */
function isHolidayDateValid_(targetDateStr, isRetroactive) {
  if (!targetDateStr || !/^\d{4}-\d{2}-\d{2}$/.test(String(targetDateStr))) {
    return { ok: false, error: '作業実施日の形式が不正です（yyyy-MM-dd）: ' + targetDateStr };
  }
  var target = new Date(targetDateStr + 'T00:00:00+09:00');
  if (isNaN(target.getTime())) {
    return { ok: false, error: '作業実施日が無効です: ' + targetDateStr };
  }
  var now = new Date();
  var todayStr = fmtDate_(now, 'yyyy-MM-dd');
  var fyStart = getFiscalYearStartDate_(now);
  var fyStartStr = fmtDate_(fyStart, 'yyyy-MM-dd');

  // 年度開始より前は常に拒否
  if (targetDateStr < fyStartStr) {
    return { ok: false, error: '年度開始日（' + fyStartStr + '）より前の日付は申請できません。' };
  }
  if (isRetroactive === true) {
    // 遡及: 今日以前のみ（未来日は遡及ではない）
    if (targetDateStr > todayStr) {
      return { ok: false, error: '遡及申請では未来日を選択できません: ' + targetDateStr };
    }
  }
  // 通常の休日出勤: 未来日も許可（年度内なら任意）
  return { ok: true, error: '' };
}

/**
 * 申請が過去日遡及かどうか判定。
 * submittedAt の日付（JST）が targetDate より後ならば true。
 *
 * @param {Date|string} submittedAt  Date オブジェクト推奨。
 *   文字列を渡す場合は ISO 8601 形式かつタイムゾーン指定付き
 *   （例: '2026-04-07T10:00:00+09:00'）でなければ JST 解釈がズレる可能性がある。
 * @param {string} targetDateStr 'yyyy-MM-dd'
 * @return {boolean}
 */
function computeIsRetroactive_(submittedAt, targetDateStr) {
  if (!submittedAt || !targetDateStr) return false;
  try {
    // Date オブジェクトが渡されればそのまま使用。文字列の場合は new Date() で解釈するが
    // タイムゾーン指定がないとローカル解釈になるため、呼び出し側で Date を渡すことを推奨。
    var subDate;
    if (submittedAt instanceof Date) {
      subDate = submittedAt;
    } else {
      subDate = new Date(String(submittedAt));
      if (isNaN(subDate.getTime())) {
        console.warn('computeIsRetroactive_: submittedAt の文字列パースに失敗: ' + submittedAt);
        return false;
      }
    }
    if (isNaN(subDate.getTime())) return false;
    var subYmd = fmtDate_(subDate, 'yyyy-MM-dd');
    return subYmd > String(targetDateStr);
  } catch (e) {
    return false;
  }
}

/**
 * targetDate と 今日 (todayYmd 'yyyy-MM-dd') の差分日数。
 * 非過去（targetDate >= today）なら 0 を返す。
 */
function computeDaysAgo_(targetDateStr, todayYmd) {
  if (!targetDateStr || !todayYmd) return 0;
  try {
    var t = new Date(String(targetDateStr) + 'T00:00:00+09:00');
    var today = new Date(String(todayYmd) + 'T00:00:00+09:00');
    if (isNaN(t.getTime()) || isNaN(today.getTime())) return 0;
    var diffMs = today.getTime() - t.getTime();
    if (diffMs <= 0) return 0;
    return Math.floor(diffMs / (24 * 60 * 60 * 1000));
  } catch (e) {
    return 0;
  }
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
