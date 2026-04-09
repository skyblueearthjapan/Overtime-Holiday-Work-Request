// ====== 共通：対象日の全申請を取得（canceled以外） ======

function listAllRequestsByDate_(dateObj) {
  var { sh, idx } = getSheetHeaderIndex_('Requests', 1);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var ymd = fmtDate_(dateObj, 'yyyy-MM-dd');
  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (!status || status === 'canceled') continue;

    var targetDateVal = row[idx['targetDate']];
    var targetYmd;
    try {
      targetYmd = targetDateVal instanceof Date
        ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
        : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
    } catch (e) { continue; }
    if (targetYmd !== ymd) continue;

    out.push({
      requestId: normalize_(row[idx['requestId']]),
      requestType: normalize_(row[idx['requestType(overtime/holiday)']]),
      status: status,
      dept: normalize_(row[idx['dept']]),
      workerName: normalize_(row[idx['workerName']]),
      workerEmail: normalize_(row[idx['workerEmail']]),
      approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
      submittedAt: row[idx['submittedAt']],
      approvedAt: row[idx['approvedAt']],
      approvedBy2: idx['approvedBy2'] !== undefined ? normalize_(row[idx['approvedBy2']]) : '',
      pdfFileId: normalize_(row[idx['pdfFileId']]),
      pdfGeneratedAt: row[idx['pdfGeneratedAt']],
    });
  }

  out.sort(function(a, b) { return (a.dept + a.workerName).localeCompare(b.dept + b.workerName, 'ja'); });
  return out;
}

// ====== 共通：対象日の承認済み申請を取得 ======

function listApprovedRequestsByDate_(dateObj) {
  const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const ymd = fmtDate_(dateObj, 'yyyy-MM-dd');
  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (status !== 'approved') continue;

    const targetDateVal = row[idx['targetDate']];
    const targetYmd = targetDateVal instanceof Date
      ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
      : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
    if (targetYmd !== ymd) continue;

    out.push({
      requestId: normalize_(row[idx['requestId']]),
      requestType: normalize_(row[idx['requestType(overtime/holiday)']]),
      dept: normalize_(row[idx['dept']]),
      workerName: normalize_(row[idx['workerName']]),
      workerEmail: normalize_(row[idx['workerEmail']]),
      approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
      submittedAt: row[idx['submittedAt']],
      approvedAt: row[idx['approvedAt']],
      pdfFileId: normalize_(row[idx['pdfFileId']]),
      pdfGeneratedAt: row[idx['pdfGeneratedAt']],
    });
  }

  // 部署→氏名で軽く整列
  out.sort((a,b)=> (a.dept+a.workerName).localeCompare(b.dept+b.workerName,'ja'));
  return out;
}

// ====== 共通：WorkLogs を requestId で引く Map ======

function buildWorkLogsMapByRequestId_() {
  const sh = requireSheet_('WorkLogs');
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return new Map();

  const header = sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0].map(h=>normalize_(h));
  const idx = buildHeaderIndex_(header);

  const ridCol = idx['requestId'];
  if (ridCol === undefined) throw new Error('WorkLogsに requestId 列がありません。');

  // TZバグ回避: DateオブジェクトはスプレッドシートTZでフォーマットして元のJST文字列を復元
  var ssTz = getSafeTimeZone_(sh);

  const values = sh.getRange(3,1,lastRow-2,sh.getLastColumn()).getValues();
  const map = new Map();

  for (const row of values) {
    const rid = normalize_(row[ridCol]);
    if (!rid) continue;
    map.set(rid, {
      actualStartAt: toJstString_(row[idx['actualStartAt']], ssTz),
      actualEndAt: toJstString_(row[idx['actualEndAt']], ssTz),
      actualMinutes: Number(row[idx['actualMinutes']] || 0),
      breakMinutes: Number(row[idx['breakMinutes']] || 0),
      netMinutes: Number(row[idx['netMinutes']] || 0),
    });
  }
  return map;
}

/**
 * WorkLogsのDate/文字列値をJST文字列に安全に変換するヘルパー。
 * GASがJST文字列をスプレッドシートTZで解釈してDate化するバグを回避。
 */
function toJstString_(val, ssTz) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, ssTz, 'yyyy-MM-dd HH:mm:ss');
  }
  return String(val);
}

/**
 * スプレッドシートのタイムゾーンを安全に取得するヘルパー。
 * getSpreadsheetTimeZone() が null/エラーの場合は
 * Session.getScriptTimeZone() → TZ定数 の順にフォールバック。
 */
function getSafeTimeZone_(sheetOrNull) {
  // 1) スプレッドシートTZ
  try {
    var ss = sheetOrNull
      ? (sheetOrNull.getParent ? sheetOrNull.getParent() : sheetOrNull)
      : getDb_();
    var tz = ss.getSpreadsheetTimeZone();
    if (tz && typeof tz === 'string') return tz;
  } catch (_) {}
  // 2) スクリプトTZ（appsscript.json の timeZone）
  try {
    var sTz = Session.getScriptTimeZone();
    if (sTz && typeof sTz === 'string') return sTz;
  } catch (_) {}
  // 3) フォールバック
  return TZ;
}

// ====== 共通：分→「X時間Y分」表記 ======

function fmtMinutesJa_(mins) {
  const m = Math.max(0, Number(mins || 0));
  const h = Math.floor(m / 60);
  const r = m % 60;
  if (h === 0) return `${r}分`;
  if (r === 0) return `${h}時間`;
  return `${h}時間${r}分`;
}

// ====================================================================
// 夕方メール（17–18 / 18–19）— 承認時間（予定）を報告
// ====================================================================

// ====== 今週末〜の休日出勤申請を取得 ======

function listHolidayRequestsForWeekend_(dateObj) {
  var { sh, idx } = getSheetHeaderIndex_('Requests', 1);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // 今週末の日付範囲を算出（api_getDashboardと同じロジック）
  var now = dateObj;
  var dow = now.getDay();
  var monday = new Date(now);
  monday.setDate(now.getDate() - (dow === 0 ? 6 : dow - 1));
  monday.setHours(0, 0, 0, 0);
  var saturday = new Date(monday);
  saturday.setDate(monday.getDate() + 5);
  var nextMonday = new Date(monday);
  nextMonday.setDate(monday.getDate() + 7);

  var weekendStart = fmtDate_(saturday, 'yyyy-MM-dd');
  var weekendEnd = fmtDate_(nextMonday, 'yyyy-MM-dd');

  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (!status || status === 'canceled') continue;

    var requestType = normalize_(row[idx['requestType(overtime/holiday)']]);
    if (requestType !== 'holiday') continue;

    var targetDateVal = row[idx['targetDate']];
    var targetYmd;
    try {
      targetYmd = targetDateVal instanceof Date
        ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
        : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
    } catch (e) { continue; }

    if (targetYmd < weekendStart || targetYmd > weekendEnd) continue;

    var dayNames = ['日','月','火','水','木','金','土'];
    var d = new Date(targetYmd + 'T00:00:00');
    var dateLabel = (d.getMonth()+1) + '/' + d.getDate() + '(' + dayNames[d.getDay()] + ')';

    out.push({
      requestId: normalize_(row[idx['requestId']]),
      requestType: requestType,
      status: status,
      dept: normalize_(row[idx['dept']]),
      workerName: normalize_(row[idx['workerName']]),
      workerEmail: normalize_(row[idx['workerEmail']]),
      approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
      submittedAt: row[idx['submittedAt']],
      approvedAt: row[idx['approvedAt']],
      approvedBy2: idx['approvedBy2'] !== undefined ? normalize_(row[idx['approvedBy2']]) : '',
      targetDate: targetYmd,
      targetDateLabel: dateLabel,
    });
  }

  out.sort(function(a, b) {
    if (a.targetDate !== b.targetDate) return a.targetDate < b.targetDate ? -1 : 1;
    return (a.dept + a.workerName).localeCompare(b.dept + b.workerName, 'ja');
  });
  return out;
}

// ====== 本文生成（部署別に見やすく） ======

function buildEveningMailBody_(dateObj) {
  const settings = getSettings_();
  const appUrl = normalize_(settings['APP_URL']) || '';

  // 残業: 当日のみ
  const allItems = listAllRequestsByDate_(dateObj);
  const overtimeItems = allItems.filter(function(it) { return it.requestType === 'overtime'; });

  // 休日: 今週末〜の範囲
  const holidayItems = listHolidayRequestsForWeekend_(dateObj);

  const dateLabel = fmtDate_(dateObj, 'yyyy/MM/dd');
  const dateParam = fmtDate_(dateObj, 'yyyy-MM-dd');

  if (overtimeItems.length === 0 && holidayItems.length === 0) {
    return [
      `【残業・休日出勤 申請状況報告】${dateLabel}`,
      '',
      '本日分の残業申請および今週末の休日出勤申請はありません。',
      '',
      appUrl ? `詳細（アプリ）：${appUrl}` : '',
    ].filter(Boolean).join('\n');
  }

  const lines = [];
  lines.push(`【残業・休日出勤 申請状況報告】${dateLabel}`);
  lines.push('');

  // 残業セクション
  if (overtimeItems.length > 0) {
    lines.push('【本日の残業申請】');
    lines.push('');
    const otGroups = new Map();
    for (const it of overtimeItems) {
      if (!otGroups.has(it.dept)) otGroups.set(it.dept, []);
      otGroups.get(it.dept).push(it);
    }
    for (const [dept, arr] of otGroups.entries()) {
      lines.push(`■ ${dept}`);
      for (const it of arr) {
        let statusLabel = '未承認';
        if (it.approvedBy2) {
          statusLabel = '二次承認済';
        } else if (it.status === 'approved') {
          statusLabel = '承認済';
        }
        lines.push(`- ${it.workerName}：残業 ${fmtMinutesJa_(it.approvedMinutes)}（${statusLabel}）`);
      }
      lines.push('');
    }
  }

  // 休日セクション
  if (holidayItems.length > 0) {
    lines.push('【今週末の休日出勤申請】');
    lines.push('');
    const hdGroups = new Map();
    for (const it of holidayItems) {
      if (!hdGroups.has(it.dept)) hdGroups.set(it.dept, []);
      hdGroups.get(it.dept).push(it);
    }
    for (const [dept, arr] of hdGroups.entries()) {
      lines.push(`■ ${dept}`);
      for (const it of arr) {
        let statusLabel = '未承認';
        if (it.approvedBy2) {
          statusLabel = '二次承認済';
        } else if (it.status === 'approved') {
          statusLabel = '承認済';
        }
        lines.push(`- ${it.workerName}：休日出勤 ${it.targetDateLabel} ${fmtMinutesJa_(it.approvedMinutes)}（${statusLabel}）`);
      }
      lines.push('');
    }
  }

  if (appUrl) {
    lines.push(`詳細（アプリ）：${appUrl}`);
    lines.push(`二次承認ページ：${appUrl}?page=approve2&date=${dateParam}`);
    lines.push('');
  }

  lines.push('※本メールは自動送信です。');
  return lines.join('\n');
}

// ====== 夕方メール送信（手動実行可） ======
// 夕方2回（17–18 / 18–19）は同じ関数を2回トリガーでOK。
// その時間点の承認状況が反映される。

function sendEveningMail_() {
  const settings = getSettings_();
  const to = normalize_(settings['HR_MAIL_TO']);
  if (!to) throw new Error('Settingsに HR_MAIL_TO が未設定です。');

  const now = new Date();
  const subject = `【残業・休日出勤】申請状況報告 ${fmtDate_(now,'yyyy/MM/dd')}`;
  const body = buildEveningMailBody_(now);

  GmailApp.sendEmail(to, subject, body);
}

// ====================================================================
// データ蓄積スプレッドシート（総務確認用台帳）
// ====================================================================

/** 蓄積SSのヘッダ定義 */
var ACC_HEADERS_ = [
  '管理番号', '申請種別', '作業実施日', '曜日',
  '部署', '作業員ID', '氏名',
  '理由', '補足理由', '業務内容',
  '承認時間(分)', '承認時間', '開始時刻', '終了時刻',
  '実働(分)', '実働', '休憩(分)', '休憩', '実残業/実働(分)', '実残業/実働',
  '申請日', '1次承認日', '1次承認者', '2次承認日', '2次承認者',
  '承認状態', 'PDF'
];

/** 蓄積SSの列幅 */
var ACC_COL_WIDTHS_ = [
  140, // 管理番号
  80,  // 申請種別
  100, // 作業実施日
  50,  // 曜日
  100, // 部署
  90,  // 作業員ID
  100, // 氏名
  160, // 理由
  160, // 補足理由
  160, // 業務内容
  80,  // 承認時間(分)
  100, // 承認時間
  80,  // 開始時刻
  80,  // 終了時刻
  80,  // 実働(分)
  100, // 実働
  80,  // 休憩(分)
  100, // 休憩
  80,  // 実残業/実働(分)
  100, // 実残業/実働
  100, // 申請日
  100, // 1次承認日
  120, // 1次承認者
  100, // 2次承認日
  120, // 2次承認者
  80,  // 承認状態
  60   // PDF
];

var DOWS_ACC_ = ['日','月','火','水','木','金','土'];

/**
 * 蓄積スプレッドシートを取得（なければ新規作成してSettingsに保存）。
 * @return {Spreadsheet}
 */
function getAccumulationSS_() {
  var settings = getSettings_();
  var ssId = normalize_(settings['ACCUMULATION_SS_ID']);
  var folderId = '1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B';
  var ss;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
      return ss;
    } catch (e) {
      Logger.log('蓄積SS取得失敗（ID: ' + ssId + '）、新規作成します: ' + e.message);
    }
  }

  ss = SpreadsheetApp.create('残業・休日出勤 データ蓄積（総務確認用）');
  ssId = ss.getId();
  try {
    var file = DriveApp.getFileById(ssId);
    var folder = DriveApp.getFolderById(folderId);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {
    Logger.log('フォルダ移動スキップ: ' + e.message);
  }
  var settingsSh = requireSheet_('Settings');
  settingsSh.appendRow(['ACCUMULATION_SS_ID', ssId]);
  Logger.log('蓄積SS新規作成: ' + ssId);
  return ss;
}

/**
 * 蓄積SSの年度シートを取得（なければ作成＋書式設定）
 * @param {Spreadsheet} ss
 * @param {string} sheetType - '残業' or '休日出勤'
 * @param {number} fy - 年度
 * @return {Sheet}
 */
function getAccFySheet_(ss, sheetType, fy) {
  var sheetName = sheetType + '_' + fy + '年度';
  var sh = ss.getSheetByName(sheetName);
  if (sh) return sh;

  sh = ss.insertSheet(sheetName, 0);
  sh.getRange(1, 1, 1, ACC_HEADERS_.length).setValues([ACC_HEADERS_]);

  // ヘッダ書式
  var hr = sh.getRange(1, 1, 1, ACC_HEADERS_.length);
  hr.setFontWeight('bold');
  hr.setBackground('#1a365d');
  hr.setFontColor('#ffffff');
  hr.setFontSize(10);
  hr.setHorizontalAlignment('center');
  hr.setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 32);

  // 列幅
  for (var c = 0; c < ACC_COL_WIDTHS_.length; c++) {
    sh.setColumnWidth(c + 1, ACC_COL_WIDTHS_[c]);
  }

  // _tmpシート（初回作成時のデフォルトシート）を削除
  var tmp = ss.getSheetByName('Sheet1');
  if (tmp) { try { ss.deleteSheet(tmp); } catch (e) {} }
  tmp = ss.getSheetByName('シート1');
  if (tmp) { try { ss.deleteSheet(tmp); } catch (e) {} }

  Logger.log('蓄積年度シート作成: ' + sheetName);
  return sh;
}

/**
 * 年度算出（3/16始まり: 休暇届appと統一）
 */
function computeAccFiscalYear_(dateObj) {
  var d = dateObj || new Date();
  var year = d.getFullYear();
  var month = d.getMonth() + 1;
  var day = d.getDate();
  if (month < 3 || (month === 3 && day < 16)) return year - 1;
  return year;
}

/**
 * 時刻文字列を抽出するヘルパー
 */
function extractTimeStr_(val) {
  if (!val) return '';
  if (val instanceof Date) return fmtDate_(val, 'HH:mm');
  var s = String(val);
  var m = s.match(/(\d{1,2}:\d{2})/);
  return m ? m[1] : s.replace(/^\d{4}-\d{2}-\d{2}\s?/, '').substring(0, 5);
}

/**
 * 当日分のデータを蓄積スプレッドシートに追記する。
 * 既に同日データがあれば削除→再追記。
 */
function appendToAccumulationSS_(dateObj) {
  var ss = getAccumulationSS_();
  var ymd = fmtDate_(dateObj, 'yyyy/MM/dd');
  var items = listApprovedRequestsByDate_(dateObj);
  var workMap = buildWorkLogsMapByRequestId_();
  var fy = computeAccFiscalYear_(dateObj);

  var overtimeRows = [];
  var holidayRows = [];

  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var wl = workMap.get(it.requestId) || {};
    var d = dateObj instanceof Date ? dateObj : new Date(dateObj);
    var row = buildAccRow_(it, wl, d);

    if (it.requestType === 'overtime') {
      overtimeRows.push(row);
    } else {
      holidayRows.push(row);
    }
  }

  writeAccFySheet_(ss, '残業', fy, ymd, overtimeRows);
  writeAccFySheet_(ss, '休日出勤', fy, ymd, holidayRows);
}

/**
 * 蓄積行データを構築
 */
function buildAccRow_(req, wl, targetDate) {
  var d = targetDate instanceof Date ? targetDate : new Date(targetDate);
  var ymd = fmtDate_(d, 'yyyy/MM/dd');
  var dow = DOWS_ACC_[d.getDay()];

  // 作業員ID取得
  var workerId = '';
  if (req.workerCode) {
    workerId = req.workerCode;
  } else {
    // Requestsシートから取得
    try {
      var reqInfo = getSheetHeaderIndex_('Requests', 1);
      var reqSh = reqInfo.sh;
      var reqIdx = reqInfo.idx;
      var lastRow = reqSh.getLastRow();
      if (lastRow >= 2) {
        var vals = reqSh.getRange(2, 1, lastRow - 1, reqSh.getLastColumn()).getValues();
        for (var r = 0; r < vals.length; r++) {
          if (normalize_(vals[r][reqIdx['requestId']]) === req.requestId) {
            workerId = normalize_(vals[r][reqIdx['workerCode']] || '');
            break;
          }
        }
      }
    } catch (e) { /* ignore */ }
  }

  var startStr = extractTimeStr_(wl.actualStartAt);
  var endStr = extractTimeStr_(wl.actualEndAt);

  // 申請日
  var submittedStr = '';
  if (req.submittedAt) {
    var sd = req.submittedAt instanceof Date ? req.submittedAt : new Date(req.submittedAt);
    if (!isNaN(sd.getTime())) submittedStr = fmtDate_(sd, 'yyyy/MM/dd');
  }

  // 1次承認
  var approved1Str = '';
  if (req.approvedAt) {
    var a1 = req.approvedAt instanceof Date ? req.approvedAt : new Date(req.approvedAt);
    if (!isNaN(a1.getTime())) approved1Str = fmtDate_(a1, 'yyyy/MM/dd');
  }
  var approver1 = req.approvedBy || '';

  // 2次承認
  var approved2Str = '';
  if (req.approvedAt2) {
    var a2 = req.approvedAt2 instanceof Date ? req.approvedAt2 : new Date(req.approvedAt2);
    if (!isNaN(a2.getTime())) approved2Str = fmtDate_(a2, 'yyyy/MM/dd');
  }
  var approver2 = req.approvedBy2 || '';

  // 承認状態
  var statusLabel = req.status || '';
  if (statusLabel === 'approved') statusLabel = '承認済';
  else if (statusLabel === 'submitted') statusLabel = '申請中';
  else if (statusLabel === 'canceled') statusLabel = '取消';

  // 申請種別ラベル
  var typeLabel = req.requestType === 'overtime' ? '残業' : '休日出勤';

  var pdfFileId = req.pdfFileId || '';

  var approvedMin = Number(req.approvedMinutes || 0);
  var actualMin = wl.actualMinutes || 0;
  var breakMin = wl.breakMinutes || 0;
  var netMin = wl.netMinutes || 0;

  return [
    req.requestId,
    typeLabel,
    ymd,
    dow,
    req.dept,
    workerId,
    req.workerName,
    req.reason || '',
    req.reasonDetail || '',
    req.workContent || '',
    approvedMin,
    fmtMinutesJa_(approvedMin),
    startStr,
    endStr,
    actualMin,
    fmtMinutesJa_(actualMin),
    breakMin,
    fmtMinutesJa_(breakMin),
    netMin,
    fmtMinutesJa_(netMin),
    submittedStr,
    approved1Str,
    approver1,
    approved2Str,
    approver2,
    statusLabel,
    pdfFileId  // 後でハイパーリンクに変換
  ];
}

/**
 * 年度シートに書き込み（同日データ削除→再追記）
 */
function writeAccFySheet_(ss, sheetType, fy, ymd, rows) {
  var sh = getAccFySheet_(ss, sheetType, fy);

  // 既存の同日データを削除（作業実施日=3列目で判定、下から削除）
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var existingValues = sh.getRange(2, 3, lastRow - 1, 1).getValues();
    for (var i = existingValues.length - 1; i >= 0; i--) {
      var cellVal = existingValues[i][0];
      var cellYmd = '';
      if (cellVal instanceof Date) cellYmd = fmtDate_(cellVal, 'yyyy/MM/dd');
      else cellYmd = String(cellVal);
      if (cellYmd === ymd) sh.deleteRow(i + 2);
    }
  }

  // 追記＋書式設定
  for (var j = 0; j < rows.length; j++) {
    var writeRow = sh.getLastRow() + 1;
    sh.getRange(writeRow, 1, 1, rows[j].length).setValues([rows[j]]);
    formatAccRow_(sh, writeRow, rows[j]);
  }
}

/**
 * 指定requestIdの申請データを蓄積SSにupsert。
 * 作業完了時・二次承認時にリアルタイムで呼び出す。
 */
function upsertAccumulationRow_(requestId) {
  try {
    var req = getRequestById_(requestId);
    if (!req) { Logger.log('upsert蓄積: 申請が見つかりません: ' + requestId); return; }

    // Requestsシートから追加情報を取得
    var reqFull = getRequestFullData_(requestId);
    if (reqFull) {
      req.submittedAt = reqFull.submittedAt || req.submittedAt;
      req.approvedAt2 = reqFull.approvedAt2 || '';
      req.approvedBy2 = reqFull.approvedBy2 || '';
      req.workerCode = reqFull.workerCode || '';
    }

    var workMap = buildWorkLogsMapByRequestId_();
    var wl = workMap.get(requestId) || {};

    var ss = getAccumulationSS_();

    var d = req.targetDate instanceof Date ? req.targetDate : new Date(req.targetDate);
    var fy = computeAccFiscalYear_(d);

    // PDF状態を最新で再取得
    var reqLatest = getRequestById_(requestId);
    if (reqLatest) req.pdfFileId = reqLatest.pdfFileId || '';

    var rowData = buildAccRow_(req, wl, d);

    var sheetType = req.requestType === 'overtime' ? '残業' : '休日出勤';
    var sh = getAccFySheet_(ss, sheetType, fy);

    // requestId（管理番号=1列目）で既存行を検索
    var lastRow = sh.getLastRow();
    var existingRowNo = -1;
    if (lastRow >= 2) {
      var ridColValues = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ridColValues.length; i++) {
        if (normalize_(ridColValues[i][0]) === requestId) {
          existingRowNo = i + 2;
          break;
        }
      }
    }

    var writeRow;
    if (existingRowNo > 0) {
      sh.getRange(existingRowNo, 1, 1, rowData.length).setValues([rowData]);
      writeRow = existingRowNo;
    } else {
      writeRow = lastRow + 1;
      sh.getRange(writeRow, 1, 1, rowData.length).setValues([rowData]);
    }

    formatAccRow_(sh, writeRow, rowData);

    Logger.log('upsert蓄積完了: ' + requestId + ' → ' + sheetType + '_' + fy + '年度' + (existingRowNo > 0 ? '(上書き)' : '(追記)'));
  } catch (e) {
    Logger.log('upsert蓄積エラー（処理続行）: ' + e.message);
  }
}

/**
 * 蓄積SSから指定requestIdの行を削除する。
 * 申請キャンセル時に呼ばれる。エラー時はログ出力のみ（処理続行）。
 */
function removeAccumulationRow_(requestId, requestType, targetDate) {
  try {
    var ss = getAccumulationSS_();
    var d = targetDate instanceof Date ? targetDate : new Date(targetDate);
    var fy = computeAccFiscalYear_(d);
    var sheetType = requestType === 'overtime' ? '残業' : '休日出勤';
    var sh = ss.getSheetByName(sheetType + '_' + fy + '年度');
    if (!sh) return; // シートが無ければ何もしない

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    // requestId（1列目）で行を検索（下から削除で行番号ずれ防止）
    var ridColValues = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = ridColValues.length - 1; i >= 0; i--) {
      if (normalize_(ridColValues[i][0]) === requestId) {
        sh.deleteRow(i + 2);
        Logger.log('蓄積SS行削除完了: ' + requestId + ' → ' + sheetType + '_' + fy + '年度 行' + (i + 2));
        return;
      }
    }
    Logger.log('蓄積SS行削除: 該当行なし（' + requestId + '）');
  } catch (e) {
    Logger.log('removeAccumulationRow_ エラー（処理続行）: ' + e.message);
  }
}

/**
 * Requestsシートからフル情報を取得（getRequestById_では取れない列を補完）
 */
function getRequestFullData_(requestId) {
  try {
    var info = getSheetHeaderIndex_('Requests', 1);
    var sh = info.sh;
    var idx = info.idx;
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return null;

    var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < values.length; i++) {
      if (normalize_(values[i][idx['requestId']]) === requestId) {
        var row = values[i];
        return {
          submittedAt: row[idx['submittedAt']] || '',
          approvedAt2: idx['approvedAt2'] !== undefined ? row[idx['approvedAt2']] : '',
          approvedBy2: idx['approvedBy2'] !== undefined ? normalize_(row[idx['approvedBy2']]) : '',
          workerCode: idx['workerCode'] !== undefined ? normalize_(row[idx['workerCode']]) : '',
        };
      }
    }
  } catch (e) {
    Logger.log('getRequestFullData_ エラー: ' + e.message);
  }
  return null;
}

/**
 * 蓄積行の書式設定（偶数行背景色＋PDFハイパーリンク）
 */
function formatAccRow_(sh, rowNo, rowData) {
  var colCount = ACC_HEADERS_.length;
  var range = sh.getRange(rowNo, 1, 1, colCount);
  range.setFontSize(10);
  range.setVerticalAlignment('middle');

  // 偶数行に薄い背景色
  if ((rowNo % 2) === 0) {
    range.setBackground('#f8fafc');
  } else {
    range.setBackground('#ffffff');
  }

  // PDF列（27列目＝ACC_HEADERS_の最終列）にハイパーリンク
  var pdfColIdx = ACC_HEADERS_.indexOf('PDF');
  var pdfFileId = rowData[pdfColIdx];
  if (pdfFileId) {
    var pdfUrl = 'https://drive.google.com/file/d/' + pdfFileId + '/view';
    sh.getRange(rowNo, pdfColIdx + 1).setFormula('=HYPERLINK("' + pdfUrl + '","PDF")');
  }
}

// ====================================================================
// 蓄積SSマイグレーション：既存シートに時間表記列を追加
// ====================================================================

/**
 * 既存の蓄積スプレッドシート全シートに時間表記列（承認時間/実働/休憩/実残業）を追加し、
 * 既存行の分データから「X時間Y分」形式を一括生成する。
 * メニューから手動実行する想定。
 */
function migrateAccumulationTimeColumns() {
  var ss;
  try {
    ss = getAccumulationSS_();
  } catch (e) {
    Logger.log('蓄積SSが見つかりません: ' + e.message);
    SpreadsheetApp.getUi().alert('蓄積スプレッドシートが見つかりません。');
    return;
  }

  var sheets = ss.getSheets();
  var migratedCount = 0;

  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var sheetName = sh.getName();
    // 残業_XXXX年度 or 休日出勤_XXXX年度 のみ対象
    if (!/^(残業|休日出勤)_\d{4}年度$/.test(sheetName)) continue;

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) continue;

    var header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });

    // 既にマイグレーション済みかチェック（「承認時間」列が存在するか）
    if (header.indexOf('承認時間') >= 0) {
      Logger.log(sheetName + ': マイグレーション済みのためスキップ');
      continue;
    }

    // 旧ヘッダの列位置を特定
    var oldApprovedMinCol = header.indexOf('承認時間(分)');   // 旧10列目(0-indexed)
    var oldActualMinCol   = header.indexOf('実働(分)');       // 旧13列目
    var oldBreakMinCol    = header.indexOf('休憩(分)');       // 旧14列目
    var oldNetMinCol      = header.indexOf('実残業/実働(分)'); // 旧15列目

    if (oldApprovedMinCol < 0) {
      Logger.log(sheetName + ': 承認時間(分)列が見つかりません、スキップ');
      continue;
    }

    // 新列を挿入（後ろから挿入してインデックスずれを防ぐ）
    // 挿入位置: 各(分)列の直後に1列ずつ挿入
    // 逆順に挿入: netMin → breakMin → actualMin → approvedMin
    var insertions = [
      { afterCol: oldNetMinCol,      newHeader: '実残業/実働' },
      { afterCol: oldBreakMinCol,    newHeader: '休憩' },
      { afterCol: oldActualMinCol,   newHeader: '実働' },
      { afterCol: oldApprovedMinCol, newHeader: '承認時間' },
    ];

    for (var ins = 0; ins < insertions.length; ins++) {
      var colPos = insertions[ins].afterCol + 1; // 1-indexed, 挿入先は直後
      // 前の挿入で列がずれた分を補正
      for (var prev = ins + 1; prev < insertions.length; prev++) {
        // 自分より前（左）に挿入された列数分だけ右にずらす
      }
      // 逆順挿入なので、既に右側で挿入した列数を加算
      var offset = 0;
      for (var k = 0; k < ins; k++) {
        if (insertions[k].afterCol > insertions[ins].afterCol) offset++;
      }
      sh.insertColumnAfter(colPos + 1 + offset); // afterCol は 0-indexed → +1 で 1-indexed
      sh.getRange(1, colPos + 2 + offset).setValue(insertions[ins].newHeader);
    }

    // 再読み込み: 挿入後の最新ヘッダを取得
    lastCol = sh.getLastColumn();
    header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });

    // 新しい列位置を取得
    var newApprovedHmCol = header.indexOf('承認時間');
    var newActualHmCol   = header.indexOf('実働');
    var newBreakHmCol    = header.indexOf('休憩');
    var newNetHmCol      = header.indexOf('実残業/実働');

    // 分列の新位置（挿入後にずれているため再取得）
    var minApprovedCol = header.indexOf('承認時間(分)');
    var minActualCol   = header.indexOf('実働(分)');
    var minBreakCol    = header.indexOf('休憩(分)');
    var minNetCol      = header.indexOf('実残業/実働(分)');

    // 既存データ行を一括変換
    if (lastRow >= 2) {
      var data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      var updates = []; // [[row, col, value], ...]

      for (var r = 0; r < data.length; r++) {
        var rowNum = r + 2;
        updates.push([rowNum, newApprovedHmCol + 1, fmtMinutesJa_(data[r][minApprovedCol])]);
        updates.push([rowNum, newActualHmCol + 1,   fmtMinutesJa_(data[r][minActualCol])]);
        updates.push([rowNum, newBreakHmCol + 1,    fmtMinutesJa_(data[r][minBreakCol])]);
        updates.push([rowNum, newNetHmCol + 1,      fmtMinutesJa_(data[r][minNetCol])]);
      }

      // 一括書き込み
      for (var u = 0; u < updates.length; u++) {
        sh.getRange(updates[u][0], updates[u][1]).setValue(updates[u][2]);
      }
    }

    // 列幅設定
    if (newApprovedHmCol >= 0) sh.setColumnWidth(newApprovedHmCol + 1, 100);
    if (newActualHmCol >= 0)   sh.setColumnWidth(newActualHmCol + 1, 100);
    if (newBreakHmCol >= 0)    sh.setColumnWidth(newBreakHmCol + 1, 100);
    if (newNetHmCol >= 0)      sh.setColumnWidth(newNetHmCol + 1, 100);

    // ヘッダ書式を再適用
    var hr = sh.getRange(1, 1, 1, lastCol);
    hr.setFontWeight('bold');
    hr.setBackground('#1a365d');
    hr.setFontColor('#ffffff');
    hr.setFontSize(10);
    hr.setHorizontalAlignment('center');

    SpreadsheetApp.flush();
    migratedCount++;
    Logger.log(sheetName + ': マイグレーション完了（' + (lastRow - 1) + '行変換）');
  }

  var msg = 'マイグレーション完了: ' + migratedCount + 'シートを更新しました。';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* trigger実行時はUI不可 */ }
}

// ====================================================================
// 作業完了未記録の申請を抽出（過去7日以内の approved + actualEndAt 空）
// ====================================================================

function listIncompleteRequests_() {
  const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const now = new Date();
  const todayYmd = fmtDate_(now, 'yyyy-MM-dd');

  // 7日前
  const sevenDaysAgo = new Date(now);
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  const limitYmd = fmtDate_(sevenDaysAgo, 'yyyy-MM-dd');

  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const wlMap = buildWorkLogsMapByRequestId_();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (status !== 'approved') continue;

    const targetDateVal = row[idx['targetDate']];
    let targetYmd;
    try {
      targetYmd = targetDateVal instanceof Date
        ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
        : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
    } catch (e) { continue; }

    // 今日の申請は除外（まだ勤務前の可能性）
    if (targetYmd >= todayYmd) continue;
    // 7日より前は除外
    if (targetYmd < limitYmd) continue;

    const requestId = normalize_(row[idx['requestId']]);
    const wl = wlMap.get(requestId) || {};

    // actualEndAt が空 → 作業完了未記録
    if (wl.actualEndAt) continue;

    out.push({
      requestId,
      requestType: normalize_(row[idx['requestType(overtime/holiday)']]),
      dept: normalize_(row[idx['dept']]),
      workerName: normalize_(row[idx['workerName']]),
      targetDate: targetYmd,
      approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
    });
  }

  // 部署→氏名→日付でソート
  out.sort((a, b) => {
    if (a.dept !== b.dept) return a.dept.localeCompare(b.dept, 'ja');
    if (a.workerName !== b.workerName) return a.workerName.localeCompare(b.workerName, 'ja');
    return a.targetDate.localeCompare(b.targetDate);
  });
  return out;
}

// ====================================================================
// 翌朝メール（7–8）— 実績一覧（CSV/Excel添付）＋ PDF作成件数
// ====================================================================

// ====== 朝レポ用データ生成（Requests × WorkLogs） ======

function buildMorningReportRows_(dateObj) {
  const items = listApprovedRequestsByDate_(dateObj);
  const workMap = buildWorkLogsMapByRequestId_();

  // 出力列（CSV/Excelの列）
  const header = [
    '日付',
    '部署',
    '氏名',
    '種別',
    '承認時間(分)',
    '承認時間',
    '開始',
    '終了',
    '実働(分)',
    '実働',
    '休憩(分)',
    '休憩',
    '実残業/実働(分)',
    '実残業/実働',
    'PDF作成',
    'requestId',
  ];

  const rows = [header];

  for (const it of items) {
    const wl = workMap.get(it.requestId) || {};
    const typeJa = it.requestType === 'overtime' ? '残業' : '休日出勤';

    const start = wl.actualStartAt instanceof Date ? fmtDate_(wl.actualStartAt,'HH:mm') : '';
    const end = wl.actualEndAt instanceof Date ? fmtDate_(wl.actualEndAt,'HH:mm') : '';

    const actualMin = wl.actualMinutes || 0;
    const breakMin = wl.breakMinutes || 0;
    const netMin = wl.netMinutes || 0;

    const pdfDone = it.pdfFileId ? '作成済' : '未作成';

    rows.push([
      fmtDate_(dateObj, 'yyyy/MM/dd'),
      it.dept,
      it.workerName,
      typeJa,
      it.approvedMinutes,
      fmtMinutesJa_(it.approvedMinutes),
      start,
      end,
      actualMin,
      fmtMinutesJa_(actualMin),
      breakMin,
      fmtMinutesJa_(breakMin),
      netMin,
      fmtMinutesJa_(netMin),
      pdfDone,
      it.requestId,
    ]);
  }
  return rows;
}

// ====== CSV 作成 ======

function makeCsvBlob_(rows, filename) {
  const esc = (v) => {
    const s = String(v ?? '');
    // CSV安全化（カンマ・改行・ダブルクォート）
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g,'""')}"`;
    return s;
  };

  const csv = rows.map(r => r.map(esc).join(',')).join('\n');
  return Utilities.newBlob(csv, 'text/csv', filename);
}

// ====== Excel（xlsx）作成（テンポラリSS → xlsxエクスポート） ======

function exportRowsToXlsxBlob_(rows, filename) {
  // 一時スプレッドシート作成
  const tmp = SpreadsheetApp.create(`TMP_EXPORT_${filename}_${Date.now()}`);
  const ssId = tmp.getId();
  const sh = tmp.getSheets()[0];
  sh.setName('Report');

  // 書き込み
  sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  SpreadsheetApp.flush();

  // xlsx エクスポート
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx`;
  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` }});
  const blob = res.getBlob().setName(filename);

  // 一時ファイル削除（ゴミ箱へ）
  DriveApp.getFileById(ssId).setTrashed(true);

  return blob;
}

// ====== PDF作成件数カウント ======

function countGeneratedPdfsForDate_(dateObj) {
  const items = listApprovedRequestsByDate_(dateObj);
  let overtime = 0;
  let holiday = 0;

  for (const it of items) {
    if (!it.pdfFileId) continue;
    if (it.requestType === 'overtime') overtime++;
    if (it.requestType === 'holiday') holiday++;
  }
  return { overtime, holiday, total: overtime + holiday };
}

// ====== 朝メール送信（本文にテキストテーブル＋リンク） ======

function sendMorningMail_(targetDateObj) {
  const settings = getSettings_();
  const to = normalize_(settings['HR_MAIL_TO']);
  if (!to) throw new Error('Settingsに HR_MAIL_TO が未設定です。');

  // 引数がなければ前日をデフォルトとする（手動実行時も前日分）
  const targetDate = targetDateObj || (function() {
    var d = new Date(); d.setDate(d.getDate() - 1); return d;
  })();
  const dateLabel = fmtDate_(targetDate, 'yyyy/MM/dd');
  const appUrl = normalize_(settings['APP_URL']) || '';
  const pdfFolderUrl = normalize_(settings['PDF_DRIVE_FOLDER_URL']) || 'https://drive.google.com/drive/folders/1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B?usp=sharing';
  const accSsId = normalize_(settings['ACCUMULATION_SS_ID']);
  const accSsUrl = accSsId ? 'https://docs.google.com/spreadsheets/d/' + accSsId + '/edit' : '';

  const items = listApprovedRequestsByDate_(targetDate);
  const workMap = buildWorkLogsMapByRequestId_();
  const counts = countGeneratedPdfsForDate_(targetDate);

  const subject = `【残業・休日出勤】実績報告 ${dateLabel}`;

  const bodyLines = [];

  // 作業完了未記録アラートを先頭に挿入
  const incompleteList = listIncompleteRequests_();
  if (incompleteList.length > 0) {
    bodyLines.push('【⚠ 作業完了未記録のお知らせ】');
    bodyLines.push('以下の作業員は作業完了ボタンが押されておらず、実働時間が記録されていません。');
    bodyLines.push('該当の作業員に連絡し、TOP画面から手入力で実績を記録するよう依頼してください。');
    bodyLines.push('※ 総務部画面（日次詳細）からも手動入力が可能です。');
    bodyLines.push('');

    // 日付別にグループ化（同じ日付のリンクをまとめる）
    const incByDate = new Map();
    for (const it of incompleteList) {
      if (!incByDate.has(it.targetDate)) incByDate.set(it.targetDate, []);
      incByDate.get(it.targetDate).push(it);
    }

    for (const [targetDate, dateItems] of incByDate.entries()) {
      // 部署別にサブグループ
      const deptMap = new Map();
      for (const it of dateItems) {
        if (!deptMap.has(it.dept)) deptMap.set(it.dept, []);
        deptMap.get(it.dept).push(it);
      }
      const dateStr = targetDate.replace(/-/g, '/');
      bodyLines.push(`▼ ${dateStr} の未記録`);
      for (const [dept, arr] of deptMap.entries()) {
        bodyLines.push(`  ■ ${dept}`);
        for (const it of arr) {
          const typeJa = it.requestType === 'overtime' ? '残業' : '休日出勤';
          bodyLines.push(`    - ${it.workerName}：${typeJa} 承認${fmtMinutesJa_(it.approvedMinutes)}`);
        }
      }
      if (appUrl) {
        bodyLines.push(`  → 作業員TOP画面：${appUrl}?page=top`);
        bodyLines.push(`  → 総務部手動入力：${appUrl}?page=admin&date=${targetDate}`);
      }
      bodyLines.push('');
    }
    bodyLines.push('---');
    bodyLines.push('');
  }

  bodyLines.push(`【残業・休日出勤 実績報告】${dateLabel}`);
  bodyLines.push('');

  if (items.length === 0) {
    bodyLines.push('本日の承認済み申請はありません。');
    bodyLines.push('');
  } else {
    // 部署別にグループ化
    const groups = new Map();
    for (const it of items) {
      if (!groups.has(it.dept)) groups.set(it.dept, []);
      groups.get(it.dept).push(it);
    }

    for (const [dept, arr] of groups.entries()) {
      bodyLines.push(`■ ${dept}`);
      for (const it of arr) {
        const wl = workMap.get(it.requestId) || {};
        const typeJa = it.requestType === 'overtime' ? '残業' : '休日出勤';
        const pdfStatus = it.pdfFileId ? 'PDF作成済' : 'PDF未作成';
        const actualMin = wl.actualMinutes || 0;
        const breakMin = wl.breakMinutes || 0;
        const netMin = wl.netMinutes || 0;

        bodyLines.push(`- ${it.workerName}：${typeJa} 承認${fmtMinutesJa_(it.approvedMinutes)} / 実働${fmtMinutesJa_(actualMin)} / 休憩${fmtMinutesJa_(breakMin)} / 実残業${fmtMinutesJa_(netMin)} [${pdfStatus}]`);
      }
      bodyLines.push('');
    }
  }

  bodyLines.push(`PDF作成状況：残業 ${counts.overtime}件 / 休日 ${counts.holiday}件 / 合計 ${counts.total}件`);
  bodyLines.push('');
  bodyLines.push(`PDFフォルダ：${pdfFolderUrl}`);
  if (accSsUrl) bodyLines.push(`データ蓄積：${accSsUrl}`);
  if (appUrl) bodyLines.push(`詳細（アプリ）：${appUrl}`);
  bodyLines.push('');
  bodyLines.push('※本メールは自動送信です。');

  GmailApp.sendEmail(to, subject, bodyLines.join('\n'));
}

// ====================================================================
// 朝バッチ（PDF一括生成 → 朝メール送信）
// トリガーからはこの関数を呼ぶ（sendMorningMail_の代わり）
// ====================================================================

function morningBatch_() {
  // 朝バッチは「前日分」の実績を対象とする
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  // 1) 承認済み＆PDF未生成の申請をまとめてPDF化
  const pdfResult = batchGeneratePdfs_(yesterday);
  Logger.log('morningBatch_ PDF: ok=' + pdfResult.ok + ' skip=' + pdfResult.skip + ' fail=' + pdfResult.fail);

  // 2) BatchLogsに記録
  logBatchResult_('morningBatch', yesterday, pdfResult);

  // 3) データ蓄積スプレッドシートに追記
  try {
    appendToAccumulationSS_(yesterday);
  } catch (e) {
    Logger.log('蓄積SSエラー: ' + e.message);
  }

  // 4) 朝メール送信（前日分）→ 統合メールに移行済み
  // sendMorningMail_(yesterday);
}

/**
 * メール送信なしの朝バッチ（PDF生成+データ蓄積のみ）
 * 統合メール通知システム移行後に使用
 */
function morningBatchNoMail_() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  const pdfResult = batchGeneratePdfs_(yesterday);
  Logger.log('morningBatchNoMail_ PDF: ok=' + pdfResult.ok + ' skip=' + pdfResult.skip + ' fail=' + pdfResult.fail);

  logBatchResult_('morningBatchNoMail', yesterday, pdfResult);

  try {
    appendToAccumulationSS_(yesterday);
  } catch (e) {
    Logger.log('蓄積SSエラー: ' + e.message);
  }
}
