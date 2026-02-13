// ====== 総務部画面：権限チェック共通 ======

function assertAdmin_() {
  const email = Session.getActiveUser().getEmail();
  if (!isAdmin_(email)) throw new Error('総務部（管理者）権限がありません。');
  return email;
}

// ====== 年度（4/1開始）ユーティリティ ======

function getFiscalYearStart_(d) {
  const dt = new Date(d);
  const y = dt.getFullYear();
  const m = dt.getMonth() + 1; // 1-12
  // 4月以上ならその年の4/1、1-3月なら前年の4/1
  const fy = (m >= 4) ? y : y - 1;
  return new Date(fy, 3, 1, 0, 0, 0); // Apr=3
}

function getFiscalYearEnd_(d) {
  const start = getFiscalYearStart_(d);
  return new Date(start.getFullYear() + 1, 3, 1, 0, 0, 0); // 次年度4/1（exclusive扱い推奨）
}

function monthKey_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM');
}

function toDate_(v) {
  if (v instanceof Date) return v;
  const d = new Date(v);
  if (isNaN(d.getTime())) return null;
  return d;
}

// ====== Requests×WorkLogs を join して "行データ" を作る（総務の基礎） ======

function buildJoinedRecords_(dateFrom, dateTo) {
  // dateFrom/dateTo: Date、dateToはexclusive推奨
  const reqShInfo = getSheetHeaderIndex_('Requests', 1);
  const reqSh = reqShInfo.sh;
  const reqIdx = reqShInfo.idx;

  // WorkLogs map（既に作成済み）
  const wlMap = buildWorkLogsMapByRequestId_();

  // Requests読み込み
  const lastRow = reqSh.getLastRow();
  if (lastRow < 2) return [];
  const reqValues = reqSh.getRange(2,1,lastRow-1,reqSh.getLastColumn()).getValues();

  const out = [];
  for (const row of reqValues) {
    const status = normalize_(row[reqIdx['status(submitted/approved/canceled)']]);
    if (!status || status === 'canceled') continue;

    const targetDate = toDate_(row[reqIdx['targetDate']]);
    if (!targetDate) continue;

    if (targetDate < dateFrom || targetDate >= dateTo) continue;

    const requestId = normalize_(row[reqIdx['requestId']]);
    const requestType = normalize_(row[reqIdx['requestType(overtime/holiday)']]);
    const dept = normalize_(row[reqIdx['dept']]);
    const workerName = normalize_(row[reqIdx['workerName']]);
    const workerCode = reqIdx['workerCode'] !== undefined ? normalize_(row[reqIdx['workerCode']]) : '';
    const approvedMinutes = Number(row[reqIdx['approvedMinutes']] || 0);

    const pdfFileId = reqIdx['pdfFileId'] !== undefined ? normalize_(row[reqIdx['pdfFileId']]) : '';
    const pdfGeneratedAt = reqIdx['pdfGeneratedAt'] !== undefined ? row[reqIdx['pdfGeneratedAt']] : '';

    const wl = wlMap.get(requestId) || {};
    const netMinutes = Number(wl.netMinutes || 0);

    out.push({
      requestId,
      status,
      requestType, // overtime/holiday
      dept,
      workerCode,
      workerName,
      targetDate,
      approvedMinutes,
      actualStartAt: wl.actualStartAt || '',
      actualEndAt: wl.actualEndAt || '',
      actualMinutes: Number(wl.actualMinutes || 0),
      breakMinutes: Number(wl.breakMinutes || 0),
      netMinutes,
      pdfFileId,
      pdfGeneratedAt,
    });
  }
  return out;
}

// ====================================================================
// 総務部API：日次一覧（任意日）
// ====================================================================

function api_adminDailyDetail(dateYmd, deptFilter) {
  assertAdmin_();
  const d = new Date(dateYmd);
  if (isNaN(d.getTime())) throw new Error('日付が不正です（yyyy-mm-dd）');

  const from = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0,0,0);
  const to = new Date(d.getFullYear(), d.getMonth(), d.getDate()+1, 0,0,0);

  let recs = buildJoinedRecords_(from, to);

  if (deptFilter && deptFilter !== 'ALL') {
    recs = recs.filter(r => r.dept === deptFilter);
  }

  // PDF未作成（承認済み＆実績完了済み＆pdf空）
  const pdfMissing = recs.filter(r => r.status === 'approved' && r.netMinutes > 0 && !r.pdfFileId).length;

  return {
    date: fmtDate_(from,'yyyy-MM-dd'),
    deptFilter: deptFilter || 'ALL',
    total: recs.length,
    pdfMissing,
    items: recs.sort((a,b)=>(a.dept+a.workerName).localeCompare(b.dept+b.workerName,'ja')),
  };
}

// ====================================================================
// 総務部API：月次40h監視（最重要）
// ====================================================================

function api_adminMonthlySummary(yearMonth, deptFilter) {
  assertAdmin_();

  // yearMonth: "yyyy-mm"
  const [yStr, mStr] = String(yearMonth).split('-');
  const y = Number(yStr), m = Number(mStr);
  if (!y || !m || m<1 || m>12) throw new Error('yearMonthは yyyy-mm で指定してください');

  const from = new Date(y, m-1, 1, 0,0,0);
  const to = new Date(y, m, 1, 0,0,0);

  let recs = buildJoinedRecords_(from, to);
  if (deptFilter && deptFilter !== 'ALL') recs = recs.filter(r => r.dept === deptFilter);

  // 個人集計
  const map = new Map(); // key=dept|name
  for (const r of recs) {
    const key = `${r.dept}__${r.workerName}`;
    if (!map.has(key)) {
      map.set(key, {
        dept: r.dept,
        workerName: r.workerName,
        overtimeNet: 0,
        holidayNet: 0,
        totalNet: 0,
        approvedTotal: 0,
        pdfMissing: 0,
      });
    }
    const agg = map.get(key);
    agg.approvedTotal += Number(r.approvedMinutes||0);
    if (r.requestType === 'overtime') agg.overtimeNet += r.netMinutes;
    if (r.requestType === 'holiday') agg.holidayNet += r.netMinutes;
    agg.totalNet += r.netMinutes;

    // PDF未作成（承認済み＆実績あり＆pdf空）
    if (r.status === 'approved' && r.netMinutes > 0 && !r.pdfFileId) agg.pdfMissing++;
  }

  const people = Array.from(map.values());
  people.sort((a,b)=> b.totalNet - a.totalNet);

  // KPI
  const LIMIT40 = 2400, LIMIT60 = 3600;
  const over40 = people.filter(p => p.totalNet >= LIMIT40).length;
  const over60 = people.filter(p => p.totalNet >= LIMIT60).length;
  const pdfMissingTotal = people.reduce((s,p)=>s+p.pdfMissing,0);

  // グラフ用（棒グラフ）
  const chart = {
    labels: people.map(p => `${p.workerName}`),
    values: people.map(p => p.totalNet), // 分
    limit40: LIMIT40,
    limit60: LIMIT60,
  };

  return {
    yearMonth,
    deptFilter: deptFilter || 'ALL',
    kpi: {
      totalPeople: people.length,
      over40,
      over60,
      pdfMissingTotal,
      totalNetMinutes: people.reduce((s,p)=>s+p.totalNet,0),
    },
    people,
    chart,
  };
}

// ====================================================================
// 総務部API：年度（4/1〜3/31）月別推移（部署別/全体）
// ====================================================================

function api_adminFiscalTrend(baseDateYmd, deptFilter) {
  assertAdmin_();

  const base = new Date(baseDateYmd);
  if (isNaN(base.getTime())) throw new Error('baseDateYmdが不正です（yyyy-mm-dd）');

  const fyStart = getFiscalYearStart_(base);
  const fyEnd = getFiscalYearEnd_(base);

  let recs = buildJoinedRecords_(fyStart, fyEnd);
  if (deptFilter && deptFilter !== 'ALL') recs = recs.filter(r => r.dept === deptFilter);

  // yyyy-mm -> sum
  const sumByMonth = new Map();
  for (const r of recs) {
    const mk = monthKey_(r.targetDate);
    sumByMonth.set(mk, (sumByMonth.get(mk) || 0) + r.netMinutes);
  }

  // 12ヶ月分並べる（4月開始）
  const labels = [];
  const values = [];
  const cursor = new Date(fyStart);
  for (let i=0; i<12; i++) {
    const mk = monthKey_(cursor);
    labels.push(mk);
    values.push(sumByMonth.get(mk) || 0);
    cursor.setMonth(cursor.getMonth()+1);
  }

  return {
    fiscalYearStart: fmtDate_(fyStart,'yyyy-MM-dd'),
    fiscalYearEnd: fmtDate_(fyEnd,'yyyy-MM-dd'),
    deptFilter: deptFilter || 'ALL',
    labels,
    values, // 分
  };
}

// ====================================================================
// 総務部API：部署プルダウン用
// ====================================================================

function api_adminDeptOptions() {
  assertAdmin_();
  const depts = loadDeptList_(); // 既に作成済み
  return ['ALL', ...depts];
}

// ====================================================================
// Web アプリ：doGet ルーティング
// ====================================================================

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'top';

  if (page === 'admin') {
    assertAdmin_();
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('総務部管理')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 既存トップ（部署選択等）
  return HtmlService.createHtmlOutput('TODO: existing routing');
}
