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
// データ蓄積スプレッドシート
// ====================================================================

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

  ss = SpreadsheetApp.create('残業・休日出勤 データ蓄積');
  ssId = ss.getId();
  // 指定フォルダに移動
  try {
    var file = DriveApp.getFileById(ssId);
    var folder = DriveApp.getFolderById(folderId);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {
    Logger.log('フォルダ移動スキップ: ' + e.message);
  }
  // Settings に保存
  var settingsSh = requireSheet_('Settings');
  settingsSh.appendRow(['ACCUMULATION_SS_ID', ssId]);
  Logger.log('蓄積SS新規作成: ' + ssId);
  return ss;
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

  var accHeader = ['日付', '部署', '氏名', '承認時間(分)', '開始', '終了', '実働(分)', '休憩(分)', '実残業/実働(分)', 'PDF', 'requestId'];

  // 残業・休日出勤でそれぞれ分割
  var overtimeRows = [];
  var holidayRows = [];

  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var wl = workMap.get(it.requestId) || {};
    var startStr = wl.actualStartAt || '';
    var endStr = wl.actualEndAt || '';
    if (startStr instanceof Date) startStr = fmtDate_(startStr, 'HH:mm');
    else if (startStr) startStr = String(startStr).replace(/^\d{4}-\d{2}-\d{2}\s?/, '').substring(0, 5);
    if (endStr instanceof Date) endStr = fmtDate_(endStr, 'HH:mm');
    else if (endStr) endStr = String(endStr).replace(/^\d{4}-\d{2}-\d{2}\s?/, '').substring(0, 5);

    var row = [
      ymd,
      it.dept,
      it.workerName,
      it.approvedMinutes,
      startStr,
      endStr,
      wl.actualMinutes || 0,
      wl.breakMinutes || 0,
      wl.netMinutes || 0,
      it.pdfFileId ? '作成済' : '未作成',
      it.requestId,
    ];

    if (it.requestType === 'overtime') {
      overtimeRows.push(row);
    } else {
      holidayRows.push(row);
    }
  }

  // シートに書き込み
  writeAccumulationSheet_(ss, '残業', accHeader, ymd, overtimeRows);
  writeAccumulationSheet_(ss, '休日出勤', accHeader, ymd, holidayRows);
}

/**
 * 蓄積SSの指定シートに書き込む。同日データがあれば削除→再追記。
 */
function writeAccumulationSheet_(ss, sheetName, header, ymd, rows) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(header);
  }

  // 既存の同日データを削除（日付列=1列目で判定、下から削除）
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var existingValues = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = existingValues.length - 1; i >= 0; i--) {
      var cellVal = existingValues[i][0];
      var cellYmd = '';
      if (cellVal instanceof Date) {
        cellYmd = fmtDate_(cellVal, 'yyyy/MM/dd');
      } else {
        cellYmd = String(cellVal);
      }
      if (cellYmd === ymd) {
        sh.deleteRow(i + 2);
      }
    }
  }

  // 追記
  for (var j = 0; j < rows.length; j++) {
    sh.appendRow(rows[j]);
  }
}

/**
 * 指定requestIdの申請データを蓄積SSにupsert（既存行があれば上書き、なければ追記）。
 * 作業完了時・二次承認時にリアルタイムで呼び出す。
 * 失敗しても呼び出し元の処理には影響しない。
 */
function upsertAccumulationRow_(requestId) {
  try {
    var req = getRequestById_(requestId);
    if (!req) { Logger.log('upsert蓄積: 申請が見つかりません: ' + requestId); return; }

    var workMap = buildWorkLogsMapByRequestId_();
    var wl = workMap.get(requestId) || {};

    var ss = getAccumulationSS_();

    // 日付フォーマット
    var d = req.targetDate instanceof Date ? req.targetDate : new Date(req.targetDate);
    var ymd = fmtDate_(d, 'yyyy/MM/dd');

    // 開始・終了の時刻部分を抽出
    var startStr = wl.actualStartAt || '';
    var endStr = wl.actualEndAt || '';
    if (startStr instanceof Date) startStr = fmtDate_(startStr, 'HH:mm');
    else if (startStr) startStr = String(startStr).replace(/^\d{4}-\d{2}-\d{2}\s?/, '').substring(0, 5);
    if (endStr instanceof Date) endStr = fmtDate_(endStr, 'HH:mm');
    else if (endStr) endStr = String(endStr).replace(/^\d{4}-\d{2}-\d{2}\s?/, '').substring(0, 5);

    // PDF状態を最新で再取得（generatePdf後にRequestsが更新されている可能性）
    var reqLatest = getRequestById_(requestId);
    var pdfStatus = (reqLatest && reqLatest.pdfFileId) ? '作成済' : '未作成';

    var rowData = [
      ymd,
      req.dept,
      req.workerName,
      Number(req.approvedMinutes || 0),
      startStr,
      endStr,
      wl.actualMinutes || 0,
      wl.breakMinutes || 0,
      wl.netMinutes || 0,
      pdfStatus,
      requestId,
    ];

    var sheetName = req.requestType === 'overtime' ? '残業' : '休日出勤';
    var accHeader = ['日付', '部署', '氏名', '承認時間(分)', '開始', '終了', '実働(分)', '休憩(分)', '実残業/実働(分)', 'PDF', 'requestId'];

    var sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
      sh.appendRow(accHeader);
    }

    // requestId列（11列目=インデックス10）で既存行を検索
    var lastRow = sh.getLastRow();
    var existingRowNo = -1;
    if (lastRow >= 2) {
      var ridColValues = sh.getRange(2, accHeader.length, lastRow - 1, 1).getValues(); // requestId列
      for (var i = 0; i < ridColValues.length; i++) {
        if (normalize_(ridColValues[i][0]) === requestId) {
          existingRowNo = i + 2;
          break;
        }
      }
    }

    if (existingRowNo > 0) {
      // 上書き
      sh.getRange(existingRowNo, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // 追記
      sh.appendRow(rowData);
    }

    Logger.log('upsert蓄積完了: ' + requestId + ' → ' + sheetName + (existingRowNo > 0 ? '(上書き)' : '(追記)'));
  } catch (e) {
    Logger.log('upsert蓄積エラー（処理続行）: ' + e.message);
  }
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
    '休憩(分)',
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
      breakMin,
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

function sendMorningMail_() {
  const settings = getSettings_();
  const to = normalize_(settings['HR_MAIL_TO']);
  if (!to) throw new Error('Settingsに HR_MAIL_TO が未設定です。');

  const now = new Date();
  const dateLabel = fmtDate_(now, 'yyyy/MM/dd');
  const appUrl = normalize_(settings['APP_URL']) || '';
  const pdfFolderUrl = normalize_(settings['PDF_DRIVE_FOLDER_URL']) || 'https://drive.google.com/drive/folders/1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B?usp=sharing';
  const accSsId = normalize_(settings['ACCUMULATION_SS_ID']);
  const accSsUrl = accSsId ? 'https://docs.google.com/spreadsheets/d/' + accSsId + '/edit' : '';

  const items = listApprovedRequestsByDate_(now);
  const workMap = buildWorkLogsMapByRequestId_();
  const counts = countGeneratedPdfsForDate_(now);

  const subject = `【残業・休日出勤】実績報告 ${dateLabel}`;

  const bodyLines = [];

  // 作業完了未記録アラートを先頭に挿入
  const incompleteList = listIncompleteRequests_();
  if (incompleteList.length > 0) {
    bodyLines.push('【⚠ 作業完了未記録のお知らせ】');
    bodyLines.push('以下の作業員は作業完了ボタンが押されておらず、実働時間が記録されていません。');
    bodyLines.push('確認の上、総務部画面（日次詳細）より手動で実績を入力してください。');
    bodyLines.push('');

    const incGroups = new Map();
    for (const it of incompleteList) {
      if (!incGroups.has(it.dept)) incGroups.set(it.dept, []);
      incGroups.get(it.dept).push(it);
    }
    for (const [dept, arr] of incGroups.entries()) {
      bodyLines.push(`■ ${dept}`);
      for (const it of arr) {
        const typeJa = it.requestType === 'overtime' ? '残業' : '休日出勤';
        const dateStr = it.targetDate.replace(/-/g, '/');
        bodyLines.push(`  - ${it.workerName}：${typeJa} 承認${fmtMinutesJa_(it.approvedMinutes)}（${dateStr}）`);
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
  const today = new Date();

  // 1) 承認済み＆PDF未生成の申請をまとめてPDF化
  const pdfResult = batchGeneratePdfs_(today);
  Logger.log('morningBatch_ PDF: ok=' + pdfResult.ok + ' skip=' + pdfResult.skip + ' fail=' + pdfResult.fail);

  // 2) BatchLogsに記録
  logBatchResult_('morningBatch', today, pdfResult);

  // 3) データ蓄積スプレッドシートに追記
  try {
    appendToAccumulationSS_(today);
  } catch (e) {
    Logger.log('蓄積SSエラー: ' + e.message);
  }

  // 4) 朝メール送信（本文にテキストテーブル＋リンク）
  sendMorningMail_();
}
