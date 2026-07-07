// ====================================================================
// スタッフビュー：月次 残業・休日出勤 集計（締め期間 16-15 / 誰でも検索）
//   - 判定基準：実績net（承認済）で確定。承認予定/申請中は「見込み」として併記。
//   - 40h=2400分 / 60h=3600分。40時間超は年6回まで（上限60時間）。
//   - 年度：3/16〜翌3/15（既存 getFiscalYearStart_ の4/1とは別物）。
//   本ページ専用。権限ガードなし（全社員が検索可）。
// ====================================================================

const OT_LIMIT40_ = 2400; // 40時間（分）
const OT_LIMIT60_ = 3600; // 60時間（分）

/** 締め月ラベル 'yyyy-MM' を返す（日<=15 は前月分、日>=16 は当月分）。 */
function closingYmOfDate_(d) {
  const dt = (d instanceof Date) ? d : new Date(d);
  let y = dt.getFullYear();
  let m = dt.getMonth() + 1; // 1-12
  if (dt.getDate() <= 15) {
    m -= 1;
    if (m === 0) { m = 12; y -= 1; }
  }
  return y + '-' + ('0' + m).slice(-2);
}

/** 締め月 'yyyy-MM' の期間 [from, to)（to は exclusive = (M+1)/16）。 */
function closingPeriodOf_(yearMonth) {
  const parts = String(yearMonth).split('-');
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  if (!y || !m || m < 1 || m > 12) throw new Error('yearMonth は yyyy-MM で指定してください: ' + yearMonth);
  const from = new Date(y, m - 1, 16, 0, 0, 0); // M/16
  const to = new Date(y, m, 16, 0, 0, 0);       // (M+1)/16（exclusive）
  return { from: from, to: to, yearMonth: y + '-' + ('0' + m).slice(-2) };
}

/** 当該日が属する年度（3/16始まり）の開始日 3/16 を返す。 */
function otFiscalYearStart_(d) {
  const dt = (d instanceof Date) ? d : new Date(d);
  const y = dt.getFullYear();
  const m = dt.getMonth() + 1;
  const day = dt.getDate();
  const fy = (m > 3 || (m === 3 && day >= 16)) ? y : y - 1;
  return new Date(fy, 2, 16, 0, 0, 0); // 3/16（3月= index 2）
}

/** レコードのワーカーキー（workerCode 優先、無ければ dept|氏名）。 */
function otWorkerKey_(r) {
  return r.workerCode ? ('C:' + r.workerCode) : ('N:' + r.dept + '|' + r.workerName);
}

/**
 * スタッフビュー用：締め月の集計を返す（個人 or 一覧）。
 * @param {string} yearMonth 締め月 'yyyy-MM'（省略時は現在の締め月）
 * @param {string} dept 部署名（''/'ALL' で全部署）
 * @param {string} workerToken スタッフ選択値 "A001 氏名"（'' で一覧）
 * @return {Object} personal or list
 */
function api_getClosingSummary(yearMonth, dept, workerToken) {
  if (!yearMonth) yearMonth = closingYmOfDate_(new Date());
  const period = closingPeriodOf_(yearMonth);

  // 年度（3/16始まり）: この締め月が属する年度を1回スキャン
  const fyStart = otFiscalYearStart_(period.from);
  const fyEnd = new Date(fyStart.getFullYear() + 1, fyStart.getMonth(), fyStart.getDate(), 0, 0, 0);

  const deptSel = (dept && dept !== 'ALL') ? dept : '';

  let fyRecs = buildJoinedRecords_(fyStart, fyEnd); // targetDate ∈ [fyStart, fyEnd)、canceled除外
  if (deptSel) fyRecs = fyRecs.filter(function(r) { return r.dept === deptSel; });

  const monthRecs = fyRecs.filter(function(r) {
    return r.targetDate >= period.from && r.targetDate < period.to;
  });

  // 期間表示用（末日 = to - 1日 = (M+1)/15）
  const lastDay = new Date(period.to.getFullYear(), period.to.getMonth(), period.to.getDate() - 1);
  const meta = {
    yearMonth: period.yearMonth,
    monthLabel: period.from.getFullYear() + '年 ' + (period.from.getMonth() + 1) + '月分',
    periodFrom: fmtDate_(period.from, 'M/d'),
    periodTo: fmtDate_(lastDay, 'M/d'),
    fyLabel: fyStart.getFullYear() + '年度',
    limit40: OT_LIMIT40_,
    limit60: OT_LIMIT60_
  };

  // ---- 個人ビュー ----
  if (workerToken) {
    const token = String(workerToken).trim();
    const sp = token.indexOf(' ');
    const code = sp >= 0 ? token.slice(0, sp) : token;
    const name = sp >= 0 ? token.slice(sp + 1).trim() : '';
    const match = function(r) {
      if (code && r.workerCode === code) return true;
      if (name && r.workerName === name && (!deptSel || r.dept === deptSel)) return true;
      return false;
    };

    const mine = monthRecs.filter(match);

    let otNet = 0, holNet = 0, pendMin = 0, pendCount = 0;
    for (let i = 0; i < mine.length; i++) {
      const r = mine[i];
      const actualized = (r.status === 'approved' && r.netMinutes > 0);
      if (r.status === 'approved') {
        if (r.requestType === 'overtime') otNet += r.netMinutes;
        else if (r.requestType === 'holiday') holNet += r.netMinutes;
      }
      if (!actualized) { pendMin += Number(r.approvedMinutes || 0); pendCount++; }
    }
    const totalNet = otNet + holNet;

    const detail = mine.slice().sort(function(a, b) { return a.targetDate - b.targetDate; })
      .map(function(r) {
        return {
          date: fmtDate_(r.targetDate, 'M/d'),
          type: r.requestType, // overtime / holiday
          approvedMinutes: Number(r.approvedMinutes || 0),
          netMinutes: Number(r.netMinutes || 0),
          breakMinutes: Number(r.breakMinutes || 0),
          status: r.status,    // approved / submitted
          pdfFileId: r.pdfFileId || ''
        };
      });

    // 年度内の 40時間超 締め月数（承認済 net で判定）
    const byYm = {};
    for (let j = 0; j < fyRecs.length; j++) {
      const r = fyRecs[j];
      if (!match(r) || r.status !== 'approved') continue;
      const ym = closingYmOfDate_(r.targetDate);
      byYm[ym] = (byYm[ym] || 0) + r.netMinutes;
    }
    let annualOver = 0;
    for (const k in byYm) { if (byYm[k] >= OT_LIMIT40_) annualOver++; }

    let wName = name, wDept = deptSel;
    if (mine.length) { wName = mine[0].workerName || name; wDept = mine[0].dept || deptSel; }

    return {
      mode: 'personal',
      meta: meta,
      worker: { workerCode: code, workerName: wName, dept: wDept },
      kpi: { otNet: otNet, holNet: holNet, totalNet: totalNet, pendMin: pendMin, pendCount: pendCount },
      detail: detail,
      annualOver: annualOver
    };
  }

  // ---- 一覧ビュー ----
  const groups = new Map();
  for (let i = 0; i < monthRecs.length; i++) {
    const r = monthRecs[i];
    const key = otWorkerKey_(r);
    if (!groups.has(key)) {
      groups.set(key, {
        workerCode: r.workerCode || '', workerName: r.workerName, dept: r.dept,
        otNet: 0, holNet: 0, totalNet: 0, pendMin: 0, annualOver: 0
      });
    }
    const g = groups.get(key);
    if (r.status === 'approved') {
      if (r.requestType === 'overtime') g.otNet += r.netMinutes;
      else if (r.requestType === 'holiday') g.holNet += r.netMinutes;
      g.totalNet = g.otNet + g.holNet;
    }
    const actualized = (r.status === 'approved' && r.netMinutes > 0);
    if (!actualized) g.pendMin += Number(r.approvedMinutes || 0);
  }

  // 年度内 40h超 締め月数（キー別）
  const annualByKey = {};
  for (let j = 0; j < fyRecs.length; j++) {
    const r = fyRecs[j];
    if (r.status !== 'approved') continue;
    const key = otWorkerKey_(r);
    const ym = closingYmOfDate_(r.targetDate);
    if (!annualByKey[key]) annualByKey[key] = {};
    annualByKey[key][ym] = (annualByKey[key][ym] || 0) + r.netMinutes;
  }

  const people = Array.from(groups.values());
  for (let p = 0; p < people.length; p++) {
    const person = people[p];
    const key = person.workerCode ? ('C:' + person.workerCode) : ('N:' + person.dept + '|' + person.workerName);
    const mm = annualByKey[key] || {};
    let c = 0;
    for (const ym in mm) { if (mm[ym] >= OT_LIMIT40_) c++; }
    person.annualOver = c;
  }
  people.sort(function(a, b) { return b.totalNet - a.totalNet; });

  return { mode: 'list', meta: meta, dept: deptSel, people: people };
}
