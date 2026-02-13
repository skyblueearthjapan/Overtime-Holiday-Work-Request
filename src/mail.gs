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

  const values = sh.getRange(3,1,lastRow-2,sh.getLastColumn()).getValues();
  const map = new Map();

  for (const row of values) {
    const rid = normalize_(row[ridCol]);
    if (!rid) continue;
    map.set(rid, {
      actualStartAt: row[idx['actualStartAt']],
      actualEndAt: row[idx['actualEndAt']],
      actualMinutes: Number(row[idx['actualMinutes']] || 0),
      breakMinutes: Number(row[idx['breakMinutes']] || 0),
      netMinutes: Number(row[idx['netMinutes']] || 0),
    });
  }
  return map;
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

// ====== 本文生成（部署別に見やすく） ======

function buildEveningMailBody_(dateObj) {
  const settings = getSettings_();
  const appUrl = normalize_(settings['APP_URL']) || '';

  const items = listApprovedRequestsByDate_(dateObj);
  const dateLabel = fmtDate_(dateObj, 'yyyy/MM/dd');

  if (items.length === 0) {
    return [
      `【残業・休日出勤 承認時間（予定）報告】${dateLabel}`,
      '',
      '本日分の「承認済み」申請はありません。',
      '',
      appUrl ? `詳細（アプリ）：${appUrl}` : '',
    ].filter(Boolean).join('\n');
  }

  // dept -> rows
  const groups = new Map();
  for (const it of items) {
    if (!groups.has(it.dept)) groups.set(it.dept, []);
    groups.get(it.dept).push(it);
  }

  const lines = [];
  lines.push(`【残業・休日出勤 承認時間（予定）報告】${dateLabel}`);
  lines.push('');
  lines.push('承認済みの申請について、予定（承認）時間を報告します。');
  lines.push('');

  for (const [dept, arr] of groups.entries()) {
    lines.push(`■ ${dept}`);
    for (const it of arr) {
      const typeJa = it.requestType === 'overtime' ? '残業' : '休日出勤';
      lines.push(`- ${it.workerName}：${typeJa} ${fmtMinutesJa_(it.approvedMinutes)}（承認済）`);
    }
    lines.push('');
  }

  if (appUrl) {
    lines.push(`詳細（アプリ）：${appUrl}`);
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
  const subject = `【残業・休日出勤】承認時間（予定）報告 ${fmtDate_(now,'yyyy/MM/dd')}`;
  const body = buildEveningMailBody_(now);

  GmailApp.sendEmail(to, subject, body);
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

// ====== 朝メール送信（CSV + Excel 添付） ======

function sendMorningMail_() {
  const settings = getSettings_();
  const to = normalize_(settings['HR_MAIL_TO']);
  if (!to) throw new Error('Settingsに HR_MAIL_TO が未設定です。');

  const now = new Date();
  const dateLabel = fmtDate_(now, 'yyyy/MM/dd');
  const appUrl = normalize_(settings['APP_URL']) || '';

  const rows = buildMorningReportRows_(now);
  const counts = countGeneratedPdfsForDate_(now);

  const subject = `【残業・休日出勤】実績一覧（CSV/Excel添付） ${dateLabel}`;

  const bodyLines = [];
  bodyLines.push(`【残業・休日出勤 実績一覧】${dateLabel}`);
  bodyLines.push('');
  bodyLines.push('本メールには、承認済み申請の「実績（開始/終了/実働/休憩/実残業）」一覧を添付しています。');
  bodyLines.push('');
  bodyLines.push(`PDF作成状況：`);
  bodyLines.push(`- 残業申請書PDF：${counts.overtime} 件`);
  bodyLines.push(`- 休日申請書PDF：${counts.holiday} 件`);
  bodyLines.push(`- 合計：${counts.total} 件`);
  bodyLines.push('');
  if (appUrl) bodyLines.push(`詳細（アプリ）：${appUrl}`);
  bodyLines.push('');
  bodyLines.push('※本メールは自動送信です。');

  const csvName = `実績一覧_${fmtDate_(now,'yyyyMMdd')}.csv`;
  const xlsxName = `実績一覧_${fmtDate_(now,'yyyyMMdd')}.xlsx`;

  const csvBlob = makeCsvBlob_(rows, csvName);
  const xlsxBlob = exportRowsToXlsxBlob_(rows, xlsxName);

  GmailApp.sendEmail(to, subject, bodyLines.join('\n'), {
    attachments: [csvBlob, xlsxBlob],
  });
}
