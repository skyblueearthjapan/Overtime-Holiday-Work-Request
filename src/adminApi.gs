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

  const [yStr, mStr] = String(yearMonth).split('-');
  const y = Number(yStr), m = Number(mStr);
  if (!y || !m || m < 1 || m > 12) throw new Error('yearMonthは yyyy-mm で指定してください');

  const from = new Date(y, m - 1, 1, 0, 0, 0);
  const to = new Date(y, m, 1, 0, 0, 0);

  let recs = buildJoinedRecords_(from, to);
  if (deptFilter && deptFilter !== 'ALL') recs = recs.filter(r => r.dept === deptFilter);

  // 月の進捗（日数）
  const today = new Date();
  const daysInMonth = new Date(y, m, 0).getDate();
  const todayInSameMonth = (today >= from && today < to);
  const dayOfMonth = todayInSameMonth ? today.getDate() : daysInMonth; // 過去月は月末扱い

  const LIMIT40 = 2400;
  const LIMIT60 = 3600;

  // 個人集計
  const map = new Map(); // key=dept|name
  const deptSum = new Map(); // dept -> totalNet

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
        // 予測
        pacePerDay: 0,
        projectedTotal: 0,
        projectedOver40: false,
        projectedOver60: false,
        remain40: LIMIT40,
        remain60: LIMIT60,
      });
    }
    const agg = map.get(key);

    agg.approvedTotal += Number(r.approvedMinutes || 0);
    if (r.requestType === 'overtime') agg.overtimeNet += r.netMinutes;
    if (r.requestType === 'holiday') agg.holidayNet += r.netMinutes;
    agg.totalNet += r.netMinutes;

    if (r.status === 'approved' && r.netMinutes > 0 && !r.pdfFileId) agg.pdfMissing++;

    deptSum.set(r.dept, (deptSum.get(r.dept) || 0) + r.netMinutes);
  }

  const people = Array.from(map.values());

  // 予測計算（今月累計 / 経過日 × 月日数）
  for (const p of people) {
    const pace = p.totalNet / Math.max(1, dayOfMonth);
    p.pacePerDay = pace;
    p.projectedTotal = Math.round(pace * daysInMonth);
    p.projectedOver40 = p.projectedTotal >= LIMIT40;
    p.projectedOver60 = p.projectedTotal >= LIMIT60;
    p.remain40 = Math.max(0, LIMIT40 - p.totalNet);
    p.remain60 = Math.max(0, LIMIT60 - p.totalNet);
  }

  // 並び：実績降順
  people.sort((a,b)=> b.totalNet - a.totalNet);

  const over40 = people.filter(p => p.totalNet >= LIMIT40).length;
  const over60 = people.filter(p => p.totalNet >= LIMIT60).length;
  const projOver40 = people.filter(p => p.projectedOver40).length;
  const projOver60 = people.filter(p => p.projectedOver60).length;
  const pdfMissingTotal = people.reduce((s,p)=>s+p.pdfMissing,0);

  // 注意対象：実績30h超 or 予測40h超
  const watch = people
    .filter(p => p.totalNet >= 1800 || p.projectedOver40)
    .map(p => ({
      dept: p.dept,
      workerName: p.workerName,
      totalNet: p.totalNet,
      projectedTotal: p.projectedTotal,
      remain40: p.remain40,
      remain60: p.remain60,
      projectedOver40: p.projectedOver40,
      projectedOver60: p.projectedOver60,
      pdfMissing: p.pdfMissing,
    }));

  // グラフ用（棒：個人）
  const chartPeople = {
    labels: people.map(p => p.workerName),
    values: people.map(p => p.totalNet), // 分
    projected: people.map(p => p.projectedTotal), // 分
    limit40: LIMIT40,
    limit60: LIMIT60,
  };

  // グラフ用（円：部署）
  const deptLabels = Array.from(deptSum.keys()).sort((a,b)=>a.localeCompare(b,'ja'));
  const chartDept = {
    labels: deptLabels,
    values: deptLabels.map(d => deptSum.get(d) || 0),
  };

  return {
    yearMonth,
    deptFilter: deptFilter || 'ALL',
    monthInfo: {
      daysInMonth,
      dayOfMonth,
      isCurrentMonth: todayInSameMonth,
    },
    kpi: {
      totalPeople: people.length,
      over40,
      over60,
      projectedOver40: projOver40,
      projectedOver60: projOver60,
      pdfMissingTotal,
      totalNetMinutes: people.reduce((s,p)=>s+p.totalNet,0),
    },
    people,
    watch,
    charts: {
      people: chartPeople,
      dept: chartDept,
    }
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
// 総務部API：特別条項（60h超）年度内回数カウント
// ====================================================================

function api_adminSpecialClauseCount(baseDateStr, deptFilter) {
  assertAdmin_();

  const base = new Date(baseDateStr);
  if (isNaN(base.getTime())) throw new Error('基準日が不正です');

  const year = base.getMonth() >= 3 ? base.getFullYear() : base.getFullYear() - 1;
  const fiscalStart = new Date(year, 3, 1);
  const fiscalEnd = new Date(year + 1, 3, 1);

  const recs = buildJoinedRecords_(fiscalStart, fiscalEnd)
    .filter(r => r.netMinutes > 0);

  const LIMIT60 = 3600;

  // 個人×月 の netMinutes 合算
  const monthlyMap = new Map(); // key: dept|name|yyyy-mm
  for (const r of recs) {
    const ym = Utilities.formatDate(r.targetDate, Session.getScriptTimeZone(), 'yyyy-MM');
    const key = `${r.dept}|${r.workerName}|${ym}`;
    monthlyMap.set(key, (monthlyMap.get(key) || 0) + r.netMinutes);
  }

  // 60h超の月をカウント
  const countMap = new Map(); // dept|name -> count60
  for (const [key, total] of monthlyMap.entries()) {
    if (total >= LIMIT60) {
      const [dept, name] = key.split('|');
      const k = `${dept}|${name}`;
      countMap.set(k, (countMap.get(k) || 0) + 1);
    }
  }

  const result = [];
  for (const [key, cnt] of countMap.entries()) {
    const [dept, name] = key.split('|');
    if (deptFilter && deptFilter !== 'ALL' && dept !== deptFilter) continue;

    result.push({
      dept,
      workerName: name,
      count60: cnt,
      isDanger: cnt >= 6,
      isWarn: cnt >= 5
    });
  }

  result.sort((a,b)=> b.count60 - a.count60);

  return {
    fiscalYear: `${year}-04-01 ~ ${year+1}-03-31`,
    list: result
  };
}

// ====================================================================
// 総務部API：未承認滞留監視
// ====================================================================

function api_adminPendingWatch() {
  assertAdmin_();

  const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  const now = new Date();
  const result = [];

  for (const row of values) {
    const status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (status !== 'submitted') continue;

    const submittedAt = new Date(row[idx['submittedAt']]);
    const hours = (now - submittedAt) / (1000*60*60);

    result.push({
      dept: normalize_(row[idx['dept']]),
      workerName: normalize_(row[idx['workerName']]),
      submittedAt: submittedAt,
      hoursPending: Math.floor(hours),
      isOver48h: hours >= 48
    });
  }

  result.sort((a,b)=> b.hoursPending - a.hoursPending);

  return result;
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
  const appUrl = ScriptApp.getService().getUrl();

  // ページ名のホワイトリスト（ここ以外は top に落とす）
  const allowed = new Set(['top', 'approver', 'admin']);
  const safePage = allowed.has(page) ? page : 'top';

  // admin のみ権限チェック（エラー時はフレンドリーな画面を返す）
  if (safePage === 'admin') {
    const email = Session.getActiveUser().getEmail();
    const adminResult = isAdminWithDebug_(email);
    if (!adminResult.ok) {
      const t = HtmlService.createTemplateFromFile('no_auth');
      t.APP_URL = appUrl;
      t.message = '総務部（管理者）権限がありません。\n'
        + 'ApproverMap シートに role=admin で登録されているか、\n'
        + 'Settings シートの ADMIN_EMAILS にメールアドレスが含まれているか確認してください。\n'
        + '（現在のアカウント: ' + (email || '取得不可') + '）\n\n'
        + '--- デバッグ情報 ---\n' + adminResult.debug;
      return t.evaluate()
        .setTitle('権限エラー')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  const t = HtmlService.createTemplateFromFile(safePage);
  t.APP_URL = appUrl;

  return t.evaluate()
    .setTitle('残業・休日出勤申請')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** HTML include 用（共通パーツ読込用） */
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
