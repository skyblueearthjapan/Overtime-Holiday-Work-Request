#!/usr/bin/env node
// 画面（5 ページ＋権限エラー）を api 無しで確かめるための見本サーバー（仕様 12-1）。
// GAS 版と同じ形の固定データを返す。業務の決まり（権限・休憩の計算など）は簡単にしか再現しない。
//
//   node vps-app/web/dev/mock-server.js [port]        既定 8766 → http://localhost:8766/?page=top
//   MOCK_ADMIN=0   管理者でない人として返す（admin・approve2 は権限エラー、approver は承認者として開く）
//
// 呼び出しの記録: GET /__calls（関数名・引数・X-Idempotency-Key）。GET /__reset で記録と失敗の予約を消す。
// 失敗の見本: GET /__fail?status=500&count=1        次の書き込みを status で失敗させる（処理はしない）
//            GET /__fail?status=500&count=1&after=1 処理した後に status を返す（届いたのに応答が失敗）
//            GET /__fail?status=401&count=1        次の呼び出し（読む関数も）と /auth/me をログイン切れにする
//            GET /__fail?status=409&count=3        次の書き込みを 3 回「処理中」で返す
// 画面は nginx と同じ CSP（script-src 'self'）で返す（インラインの script や on 属性を使うと動かない）。
'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = Number(process.argv[2]) || 8766;
const VPS = path.resolve(__dirname, '..', '..');
const STATIC_DIR = path.join(VPS, 'web', 'static');
const TPL_DIR = path.join(VPS, 'api', 'app', 'templates', 'pages');
const ADMIN = process.env.MOCK_ADMIN !== '0';
const USER = { email: ADMIN ? 'imaizumi@lineworks-local.info' : 'shonin@lineworks-local.info' };
const PORTAL_URL = 'https://script.google.com/a/macros/lineworks-local.info/s/PORTAL_MOCK/exec';
const CSP = "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; " +
  "connect-src 'self'; frame-ancestors 'self'; base-uri 'self'; form-action 'self'";

// src/config.gs と同じ値
const REASONS = ['急ぎ作業を要するため', '納期遅延解消のため', 'マシントラブル対応のため', '業者対応のため', '顧客対応のため', 'その他（→補足理由必須）'];
const OT_HOURS = ['0.5', '1.0', '1.5', '2.0', '2.5', '3.0', '3.5', '4.0'];
const HD_HOURS = ['半日', '1日'];

// ---- 日付（JST） ----
function pad(n) { return String(n).padStart(2, '0'); }
function jst(d) { return new Date((d ? d.getTime() : Date.now()) + 9 * 3600 * 1000); }
function ymd(d) { const j = jst(d); return j.getUTCFullYear() + '-' + pad(j.getUTCMonth() + 1) + '-' + pad(j.getUTCDate()); }
function ymdhms(d) { const j = jst(d); return ymd(d) + ' ' + pad(j.getUTCHours()) + ':' + pad(j.getUTCMinutes()) + ':' + pad(j.getUTCSeconds()); }
function addDays(s, n) { const d = new Date(s + 'T00:00:00+09:00'); return ymd(new Date(d.getTime() + n * 86400000)); }
function label(s) { const j = jst(new Date(s + 'T00:00:00+09:00')); return (j.getUTCMonth() + 1) + '/' + j.getUTCDate() + '(' + '日月火水木金土'[j.getUTCDay()] + ')'; }
const TODAY = ymd();
const YESTERDAY = addDays(TODAY, -1);

// ---- 固定データ ----
const DEPTS = ['品質管理課', '製造1課', '製造2課'];
const WORKERS = {
  '製造1課': ['1001 山田 太郎', '1002 佐藤 花子', "1003 x');alert(1)//"],
  '製造2課': ['2001 鈴木 一郎', '2002 "><img src=x onerror=alert(1)>'],
  '品質管理課': ['3001 高橋 次郎']
};
const ORDERS = [
  { code: 'K-1001', product: '制御盤', dest: 'A工業' },
  { code: 'K-1002', product: '分電盤', dest: 'B製作所' },
  { code: 'K-2001', product: '操作盤', dest: 'C電機' }
];
let seq = 0;
const reqs = [];
function mk(o) {
  const r = Object.assign({
    requestId: 'mock-' + String(++seq).padStart(4, '0') + '-0000-0000-000000000000',
    requestType: 'overtime', status: 'submitted', dept: '製造1課', workerCode: '1001', workerName: '山田 太郎',
    targetDate: TODAY, approvedMinutes: 120, actualStartAt: '', actualEndAt: '', actualMinutes: 0, breakMinutes: 0,
    netMinutes: 0, pdfFileId: '', approvedBy: '', approvedAt: '', approvedBy2: '', approvedAt2: '',
    submittedAt: TODAY + ' 09:00:00', reason: REASONS[0]
  }, o);
  reqs.push(r);
  return r;
}
mk({});
mk({ workerCode: '1002', workerName: '佐藤 花子', status: 'approved', approvedBy: 'shonin@lineworks-local.info', approvedAt: TODAY + ' 10:00:00' });
mk({ workerCode: '1003', workerName: "x');alert(1)//", status: 'approved', approvedAt: TODAY + ' 10:00:00' });
mk({ dept: '製造2課', workerCode: '2002', workerName: '"><img src=x onerror=alert(1)>', status: 'approved', approvedAt: TODAY + ' 10:00:00',
  actualStartAt: TODAY + ' 17:20:00', actualEndAt: TODAY + ' 19:30:00', actualMinutes: 130, breakMinutes: 0, netMinutes: 130, pdfFileId: 'mockPdf1' });
mk({ targetDate: YESTERDAY, submittedAt: YESTERDAY + ' 09:00:00', workerName: '期限 超過', workerCode: '1009' });
mk({ requestType: 'holiday', targetDate: addDays(TODAY, 2), approvedMinutes: 480, status: 'approved', approvedAt: TODAY + ' 08:00:00' });
mk({ requestType: 'holiday', dept: '品質管理課', workerCode: '3001', workerName: '高橋 次郎', targetDate: TODAY, approvedMinutes: 240,
  status: 'approved', actualStartAt: TODAY + ' 08:00:00', submittedAt: addDays(TODAY, -3) + ' 12:00:00' });

function find(id) {
  const r = reqs.find((x) => x.requestId === id);
  if (!r) { throw new Error('申請が見つかりません: ' + id); }
  return r;
}
function isRetro(r) { return r.submittedAt.slice(0, 10) > r.targetDate; }
function daysAgo(r) {
  const d = (new Date(TODAY + 'T00:00:00Z') - new Date(r.targetDate + 'T00:00:00Z')) / 86400000;
  return d > 0 ? d : 0;
}
function dashItem(r) {
  return {
    requestId: r.requestId, requestType: r.requestType, status: r.status, dept: r.dept, workerCode: r.workerCode,
    workerName: r.workerName, targetDate: r.targetDate, targetDateLabel: r.targetDate === TODAY ? '' : label(r.targetDate),
    approvedMinutes: r.approvedMinutes, actualStartAt: r.actualStartAt, actualEndAt: r.actualEndAt,
    actualMinutes: r.actualMinutes, breakMinutes: r.breakMinutes, netMinutes: r.netMinutes, pdfFileId: r.pdfFileId,
    approvedBy2: r.approvedBy2, isOverdue: r.targetDate < TODAY && !r.actualEndAt, reason: r.reason
  };
}
function minutesBetween(a, b) { return Math.round((new Date(b.replace(' ', 'T') + '+09:00') - new Date(a.replace(' ', 'T') + '+09:00')) / 60000); }
function complete(r, start, end) {
  if (end <= start) { throw new Error('終了時刻は開始時刻より後にしてください。'); }
  r.actualStartAt = start; r.actualEndAt = end;
  r.actualMinutes = minutesBetween(start, end);
  r.breakMinutes = r.actualMinutes >= 130 ? 10 : 0;  // 見本（本物は時刻窓方式）
  r.netMinutes = r.actualMinutes - r.breakMinutes;
  return { ok: true, requestId: r.requestId, actualMinutes: r.actualMinutes, breakMinutes: r.breakMinutes, netMinutes: r.netMinutes, pdf: null };
}
function hm(s, def) { return s ? TODAY + ' ' + s + ':00' : def; }
function approverItem(r) {
  return {
    requestId: r.requestId, requestType: r.requestType, status: r.status, dept: r.dept, workerName: r.workerName,
    targetDate: r.targetDate, targetDateLabel: label(r.targetDate), approvedMinutes: r.approvedMinutes,
    submittedAt: r.submittedAt, isOverdue: r.status === 'submitted' && r.targetDate < TODAY,
    isRetroactive: isRetro(r), daysAgo: daysAgo(r)
  };
}
function monthly(deptFilter) {
  const people = {};
  reqs.filter((r) => r.status !== 'canceled' && (deptFilter === 'ALL' || !deptFilter || r.dept === deptFilter)).forEach((r) => {
    const k = r.dept + '__' + r.workerName;
    const p = people[k] = people[k] || { dept: r.dept, workerName: r.workerName, totalNet: 0, overtimeNet: 0, holidayNet: 0, approvedTotal: 0, pdfMissing: 0 };
    p.totalNet += r.netMinutes * 10;
    if (r.requestType === 'overtime') { p.overtimeNet += r.netMinutes * 10; } else { p.holidayNet += r.netMinutes * 10; }
    p.approvedTotal += r.approvedMinutes;
    if (r.status === 'approved' && r.netMinutes > 0 && !r.pdfFileId) { p.pdfMissing++; }
  });
  const list = Object.values(people).map((p) => Object.assign(p, {
    projectedTotal: Math.round(p.totalNet * 1.6), remain40: Math.max(0, 2400 - p.totalNet), remain60: Math.max(0, 3600 - p.totalNet),
    over40: p.totalNet >= 2400, over60: p.totalNet >= 3600
  })).map((p) => Object.assign(p, { projectedOver40: p.projectedTotal >= 2400, projectedOver60: p.projectedTotal >= 3600 }))
    .sort((a, b) => b.totalNet - a.totalNet);
  return {
    yearMonth: TODAY.slice(0, 7), deptFilter: deptFilter,
    monthInfo: { daysInMonth: 30, dayOfMonth: Number(TODAY.slice(8)), isCurrentMonth: true },
    kpi: { totalPeople: list.length, over40: list.filter((p) => p.over40).length, over60: list.filter((p) => p.over60).length,
      projectedOver40: list.filter((p) => p.projectedOver40).length, projectedOver60: list.filter((p) => p.projectedOver60).length,
      pdfMissingTotal: list.reduce((s, p) => s + p.pdfMissing, 0), totalNetMinutes: list.reduce((s, p) => s + p.totalNet, 0) },
    people: list, watch: list.filter((p) => p.totalNet >= 1800 || p.projectedOver40),
    charts: { people: { labels: list.map((p) => p.workerName), values: list.map((p) => p.totalNet), projected: list.map((p) => p.projectedTotal), limit40: 2400, limit60: 3600 },
      dept: { labels: DEPTS, values: DEPTS.map((d) => list.filter((p) => p.dept === d).reduce((s, p) => s + p.totalNet, 0)) } }
  };
}
function closingMeta() {
  const ym = TODAY.slice(0, 7);
  return { yearMonth: ym, monthLabel: ym.slice(0, 4) + '年 ' + Number(ym.slice(5)) + '月分', periodFrom: Number(ym.slice(5)) + '/16', periodTo: '', fyLabel: ym.slice(0, 4) + '年度', limit40: 2400, limit60: 3600 };
}

const RPC = {
  api_getDeptList: () => DEPTS.slice(),
  api_getWorkersByDept: (dept) => (dept ? (WORKERS[dept] || []).slice() : []),
  api_getWorkerInfo: () => null,  // 今のマスタに「Googleアカウント」列が無いので常に null（gap-4 F-12）
  api_getOrderList: () => ORDERS.slice(),
  api_getHolidayCandidateDates: () => [1, 2, 3, 4].map((n) => {
    const d = addDays(TODAY, n);
    return { date: d, label: d.slice(5).replace('-', '/') + label(d).replace(/^[^(]*/, ''), type: n === 3 ? 'holiday' : 'weekend', holidayName: n === 3 ? '見本の祝日' : undefined };
  }),
  api_getDashboard: () => ({
    today: TODAY,
    overtime: reqs.filter((r) => r.requestType === 'overtime' && r.status !== 'canceled' && (r.targetDate === TODAY || !r.actualEndAt)).map(dashItem),
    holiday: reqs.filter((r) => r.requestType === 'holiday' && r.status !== 'canceled').map(dashItem)
  }),
  api_submitRequest: (data) => {
    if (!data || !data.workerLabel) { return { ok: false, error: '作業員を選択してください。' }; }
    const sp = data.workerLabel.indexOf(' ');
    const r = mk({ requestType: data.requestType, dept: data.dept, workerCode: data.workerLabel.slice(0, sp), workerName: data.workerLabel.slice(sp + 1),
      targetDate: data.targetDate, approvedMinutes: data.requestType === 'overtime' ? Math.round(Number(data.hours) * 60) : (data.hours === '1日' ? 480 : 240),
      submittedAt: ymdhms(), reason: data.reason });
    return { ok: true, success: true, requestId: r.requestId };
  },
  api_markOvertimeDone: (id, end) => { const r = find(id); return complete(r, r.targetDate + ' 17:20:00', hm(end, ymdhms())); },
  api_markHolidayStart: (id, start) => { const r = find(id); r.actualStartAt = hm(start, ymdhms()); return { ok: true, requestId: id, actualStartAt: r.actualStartAt }; },
  api_markHolidayDone: (id, end) => {
    const r = find(id);
    if (!r.actualStartAt) { throw new Error('休日開始が記録されていません（開始ボタンを押してください）。'); }
    return complete(r, r.actualStartAt, hm(end, ymdhms()));
  },
  api_editRequest: (id, patch) => {
    const r = find(id);
    if (r.status === 'canceled') { throw new Error('取消済みの申請は編集できません。'); }
    r.approvedMinutes = r.requestType === 'overtime' ? Math.round(Number(patch.hours) * 60) : (patch.hours === '1日' ? 480 : 240);
    const out = { ok: true, newStatus: r.status, newMinutes: r.approvedMinutes };
    if (r.requestType === 'holiday' && patch.targetDate && patch.targetDate !== r.targetDate) { r.targetDate = patch.targetDate; out.newTargetDate = r.targetDate; }
    return out;
  },
  api_deleteRequest: (id) => {
    const r = find(id);
    if (r.status === 'canceled') { throw new Error('既にキャンセル済みです。'); }
    r.status = 'canceled';
    return { ok: true };
  },
  api_getApproverDepts: () => (ADMIN
    ? { isAdmin: true, approvers: [{ email: 'shonin@lineworks-local.info', name: '承認 一郎', depts: ['製造1課', '製造2課'] }, { email: 'qc@lineworks-local.info', name: '品管 二郎', depts: ['品質管理課'] }], currentEmail: USER.email }
    : { isAdmin: false, approvers: [{ email: USER.email, name: '承認 一郎', depts: ['製造1課', '製造2課'] }], currentEmail: USER.email }),
  api_getApproverDashboard: (depts) => {
    const ds = Array.isArray(depts) ? depts : [depts];
    if (ds.length === 0 || !ds[0]) { return { today: '', overtime: [], holiday: [], isAdmin: ADMIN, error: '部署が指定されていません。' }; }
    const mine = reqs.filter((r) => ds.includes(r.dept) && (r.status === 'submitted' || r.targetDate === TODAY));
    return { today: TODAY, overtime: mine.filter((r) => r.requestType === 'overtime').map(approverItem), holiday: mine.filter((r) => r.requestType === 'holiday').map(approverItem), isAdmin: ADMIN, selectedDepts: ds };
  },
  api_approveRequest: (id) => {
    const r = find(id);
    if (r.status === 'approved') { return { ok: true, message: '既に承認済みです。' }; }
    r.status = 'approved'; r.approvedBy = USER.email; r.approvedAt = ymdhms();
    return { ok: true, requestId: id, approvedBy: r.approvedBy, approvedAt: r.approvedAt };
  },
  api_rejectRequest: (id) => {
    const r = find(id);
    if (r.status !== 'submitted') { throw new Error('未承認状態でのみ却下可能です（現在: ' + r.status + '）。'); }
    r.status = 'canceled'; r.approvedBy = USER.email + '(却下)';
    return { ok: true, requestId: id };
  },
  api_approveRequestsBatch: (ids) => {
    if (ids.length > 50) { throw new Error('一括承認は50件までです。'); }
    return { ok: true, results: ids.map((id) => Object.assign({ requestId: id, ok: true }, RPC.api_approveRequest(id))) };
  },
  api_approverMonthlySummary: (ym, depts) => {
    const all = monthly('ALL');
    const ds = Array.isArray(depts) ? depts : [depts];
    all.people = all.people.filter((p) => ds.includes(p.dept));
    all.watch = all.watch.filter((p) => ds.includes(p.dept));
    return all;
  },
  api_getSecondaryApprovalList: (day) => ({
    date: day,
    items: reqs.filter((r) => r.targetDate === day && r.status !== 'canceled').map((r) => ({
      requestId: r.requestId, requestType: r.requestType, status: r.status, dept: r.dept, workerName: r.workerName,
      targetDate: r.targetDate, isOverdue: r.targetDate < day, approvedMinutes: r.approvedMinutes,
      submittedAt: r.submittedAt.slice(0, 16), approvedBy: r.approvedBy, approvedAt: (r.approvedAt || '').slice(0, 16),
      approvedBy2: r.approvedBy2, approvedAt2: (r.approvedAt2 || '').slice(0, 16), pdfFileId: r.pdfFileId,
      isRetroactive: isRetro(r), daysAgo: daysAgo(r)
    }))
  }),
  api_secondaryApprove: (id) => {
    const r = find(id);
    if (r.approvedBy2) { return { ok: true, message: '既に二次承認済みです。' }; }
    if (r.status !== 'approved') { throw new Error('一次承認が完了していません（現在: ' + r.status + '）。'); }
    r.approvedBy2 = USER.email; r.approvedAt2 = ymdhms();
    return { ok: true, requestId: id, approvedBy2: r.approvedBy2, approvedAt2: r.approvedAt2, pdfRegenerated: false };
  },
  api_secondaryApproveBatch: (ids) => {
    if (ids.length > 50) { throw new Error('一括承認は50件までです。'); }
    return { ok: true, results: ids.map((id) => { try { return Object.assign({ requestId: id }, RPC.api_secondaryApprove(id)); } catch (e) { return { requestId: id, ok: false, error: e.message }; } }) };
  },
  api_adminDeptOptions: () => ['ALL'].concat(DEPTS),
  api_adminMonthlySummary: (ym, dept) => monthly(dept),
  api_adminFiscalTrend: (day, dept) => {
    const y = Number(day.slice(0, 4)) - (Number(day.slice(5, 7)) < 4 ? 1 : 0);
    const labels = Array.from({ length: 12 }, (_, i) => { const m = 4 + i; return (m > 12 ? y + 1 : y) + '-' + pad(m > 12 ? m - 12 : m); });
    return { fiscalYearStart: y + '-04-01', fiscalYearEnd: (y + 1) + '-04-01', deptFilter: dept, labels: labels, values: labels.map((_, i) => (i * 397) % 3000) };
  },
  // 9-1 #22: 日付は文字で返す（GAS では Date が入って null になっていた）。過去申請のバッジのために isRetroactive・daysAgo も
  api_adminDailyDetail: (day, dept) => {
    const items = reqs.filter((r) => r.targetDate === day && r.status !== 'canceled' && (dept === 'ALL' || r.dept === dept)).map((r) => ({
      requestId: r.requestId, status: r.status, requestType: r.requestType, dept: r.dept, workerCode: r.workerCode, workerName: r.workerName,
      targetDate: r.targetDate, approvedMinutes: r.approvedMinutes, actualStartAt: r.actualStartAt, actualEndAt: r.actualEndAt,
      actualMinutes: r.actualMinutes, breakMinutes: r.breakMinutes, netMinutes: r.netMinutes, pdfFileId: r.pdfFileId,
      pdfGeneratedAt: r.pdfFileId ? r.targetDate + ' 20:00:00' : '', isRetroactive: isRetro(r), daysAgo: daysAgo(r)
    })).sort((a, b) => (a.dept + a.workerName).localeCompare(b.dept + b.workerName, 'ja'));
    return { date: day, deptFilter: dept, total: items.length, pdfMissing: items.filter((i) => i.status === 'approved' && i.netMinutes > 0 && !i.pdfFileId).length, items: items };
  },
  api_adminManualComplete: (id, end, start) => complete(find(id), find(id).targetDate + ' ' + start + ':00', find(id).targetDate + ' ' + end + ':00'),
  api_adminSpecialClauseCount: () => ({ fiscalYear: '2026-04-01 ~ 2027-03-31', list: [{ dept: '製造1課', workerName: '山田 太郎', count60: 5, isDanger: false, isWarn: true }] }),
  api_adminPendingWatch: () => reqs.filter((r) => r.status === 'submitted').map((r) => ({
    requestType: r.requestType, dept: r.dept, workerName: r.workerName, targetDateLabel: label(r.targetDate), isPastDate: r.targetDate < TODAY,
    submittedAt: r.submittedAt, hoursPending: Math.max(0, Math.floor((Date.now() - new Date(r.submittedAt.replace(' ', 'T') + '+09:00')) / 3600000)),
    isOver48h: false, targetDate: r.targetDate, isRetroactive: isRetro(r), daysAgo: daysAgo(r)
  })),
  api_getClosingSummary: (ym, dept, token) => {
    if (token) {
      const sp = token.indexOf(' ');
      return { mode: 'personal', meta: closingMeta(), worker: { workerCode: token.slice(0, sp), workerName: token.slice(sp + 1), dept: dept },
        kpi: { otNet: 600, holNet: 480, totalNet: 1080, pendMin: 120, pendCount: 1 },
        detail: [{ date: '11/20', type: 'overtime', approvedMinutes: 120, netMinutes: 110, breakMinutes: 10, status: 'approved', pdfFileId: 'mockPdf1' }],
        annualOver: 1 };
    }
    return { mode: 'list', meta: closingMeta(), dept: dept, people: [
      { workerCode: '1001', workerName: '山田 太郎', dept: '製造1課', otNet: 1500, holNet: 480, totalNet: 1980, pendMin: 0, annualOver: 0 },
      { workerCode: '1003', workerName: "x');alert(1)//", dept: '製造1課', otNet: 300, holNet: 0, totalNet: 300, pendMin: 60, annualOver: 0 },
      { workerCode: '2002', workerName: '"><img src=x onerror=alert(1)>', dept: '製造2課', otNet: 2500, holNet: 0, totalNet: 2500, pendMin: 0, annualOver: 2 }
    ] };
  },
  // 4-2 管理者だけ
  ops_getStatus: () => ({ rowSync: { pending: 2, oldestSeconds: 14, blocked: [] }, gasJobs: { pending: 1, running: 0, failed: 1,
    failedJobs: [{ id: 41, request_id: reqs[3].requestId, action: 'afterComplete', last_error: 'GAS が応答しませんでした', attempts: 8 }] },
    lastGasPing: { at: ymdhms(), ok: true }, morningBatch: { at: TODAY + ' 06:50:12', ok: true }, reconcileFixed: 0 }),
  ops_retryJob: () => ({ ok: true }),
  ops_syncSheetConfigNow: () => ({ ok: true, counts: { settings: 12, dept_approvers: 13 }, overlaps: [], mismatches: [] }),
  ops_runMasterSync: () => ({ ok: true, tabs: { '作業員マスタ': { ok: true, rows: 80 } } }),
  ops_regeneratePdf: () => ({ ok: true, jobId: 101 }),
  ops_batchPdfForDate: () => ({ ok: true, jobId: 102 }),
  ops_rebuildLedger: () => ({ ok: true, jobId: 103 }),
  ops_recalcBreaks: () => ({ ok: true, changed: 0 }),
  ops_revokeSessions: () => ({ ok: true, revoked: 1 })
};
const READ_ONLY = ['api_getDeptList', 'api_getWorkersByDept', 'api_getWorkerInfo', 'api_getOrderList', 'api_getHolidayCandidateDates',
  'api_getDashboard', 'api_getApproverDepts', 'api_getApproverDashboard', 'api_approverMonthlySummary', 'api_getSecondaryApprovalList',
  'api_adminDeptOptions', 'api_adminMonthlySummary', 'api_adminFiscalTrend', 'api_adminDailyDetail', 'api_adminSpecialClauseCount',
  'api_adminPendingWatch', 'api_getClosingSummary', 'ops_getStatus'];
const ADMIN_ONLY = /^(ops_|api_admin|api_getSecondaryApprovalList|api_secondaryApprove)/;

// ---- テンプレート（pages.py と同じ値を差し込む。Jinja2 の代わりに {{ X }} を置き換えるだけ） ----
function escHtml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&#34;').replace(/'/g, '&#39;'); }
function render(file, vars) {
  return fs.readFileSync(path.join(TPL_DIR, file), 'utf8')
    .replace(/\{\{ data_json\|safe \}\}/g, () => vars.data_json || '')
    .replace(/\{\{ (\w+) \}\}/g, (all, k) => escHtml(vars[k] === undefined ? '' : vars[k]));
}
function jsonForScript(o) { return JSON.stringify(o).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026'); }
function validDate(s) { return /^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(s || '') && ymd(new Date(s + 'T00:00:00+09:00')) === s ? s : ''; }

const calls = [];
const done = {};
let failNext = null;
function takeFail(isWrite) {
  if (!failNext || (!isWrite && failNext.status !== 401)) { return null; }
  const f = failNext;
  f.count--;
  if (f.count <= 0) { failNext = null; }
  return f;
}

function send(res, status, body, type, extra) {
  const text = (typeof body === 'string' || Buffer.isBuffer(body)) ? body : JSON.stringify(body);
  res.writeHead(status, Object.assign({ 'Content-Type': type || 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }, extra || {}));
  res.end(text);
}

const TYPES = { '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml' };

http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  const p = url.pathname;
  if (p === '/__calls') { return send(res, 200, calls); }
  if (p === '/__reset') { calls.length = 0; failNext = null; return send(res, 200, { ok: true }); }
  if (p === '/__fail') {
    failNext = { status: Number(url.searchParams.get('status')) || 500, count: Number(url.searchParams.get('count')) || 1, after: url.searchParams.get('after') === '1' };
    return send(res, 200, { ok: true, failNext: failNext });
  }
  if (p === '/auth/me') {
    if (failNext && failNext.status === 401) { return send(res, 401, { ok: false, error: 'ログインが必要です。' }); }
    return send(res, 200, { email: USER.email, name: USER.email.split('@')[0] });
  }
  if (p.startsWith('/auth/')) {
    const next = url.searchParams.get('next') || '/';
    return send(res, 200, '<!doctype html><meta charset="utf-8"><title>mock login</title><p>見本サーバーのログイン画面（' + escHtml(p) + '）</p>' +
      '<p><a href="' + escHtml(next.startsWith('/') && !next.startsWith('//') ? next : '/') + '">ログインしたことにして戻る</a></p>', 'text/html; charset=utf-8', { 'Content-Security-Policy': CSP });
  }
  if (p.startsWith('/static/')) {
    const file = path.normalize(path.join(STATIC_DIR, p.slice('/static/'.length)));
    if (!file.startsWith(STATIC_DIR) || !fs.existsSync(file) || !fs.statSync(file).isFile()) { return send(res, 404, 'not found', 'text/plain'); }
    return send(res, 200, fs.readFileSync(file), TYPES[path.extname(file)] || 'application/octet-stream');
  }
  if (p.startsWith('/api/rpc/')) {
    const name = decodeURIComponent(p.slice('/api/rpc/'.length));
    let raw = '';
    req.on('data', (c) => { raw += c; });
    req.on('end', () => {
      let args;
      try { args = raw ? JSON.parse(raw) : []; } catch (e) { return send(res, 400, { ok: false, error: '入力が不正です' }); }
      const key = req.headers['x-idempotency-key'] || '';
      calls.push({ name: name, args: args, idempotencyKey: key, at: ymdhms() });
      if (req.method !== 'POST') { return send(res, 405, { ok: false, error: 'POST だけです' }); }
      if (!Object.prototype.hasOwnProperty.call(RPC, name)) { return send(res, 501, { ok: false, error: '未登録の関数です: ' + name }); }
      if (!Array.isArray(args)) { return send(res, 400, { ok: false, error: '入力が不正です' }); }
      const write = READ_ONLY.indexOf(name) < 0;
      const fail = takeFail(write);
      if (fail && fail.status === 401) { return send(res, 401, { ok: false, error: 'ログインが必要です。', login: '/auth/login' }); }
      if (!ADMIN && ADMIN_ONLY.test(name)) { return send(res, 403, { ok: false, error: '総務部（管理者）権限がありません。' }); }
      if (write && key && done[key]) { return send(res, 200, done[key]); }
      if (fail && fail.status === 409) { return send(res, 409, { ok: false, error: '同じ操作を処理中です。' }); }
      if (fail && !fail.after) { return send(res, fail.status, { ok: false, error: 'サーバーでエラーが起きました。時間をおいてもう一度お試しください。' }); }
      let body;
      try {
        body = { ok: true, result: RPC[name].apply(null, args) };
      } catch (e) {
        return send(res, 400, { ok: false, error: e.message });
      }
      if (write && key) { done[key] = body; }
      if (fail) { return send(res, fail.status, { ok: false, error: 'サーバーでエラーが起きました。時間をおいてもう一度お試しください。' }); }
      // 実際のサーバーの遅さに近づける
      setTimeout(() => send(res, 200, JSON.parse(JSON.stringify(body))), write ? 300 : 120);
    });
    return undefined;
  }
  if (p === '/') {
    const pageParam = url.searchParams.get('page') || 'top';
    if (pageParam === 'redirect') { res.writeHead(302, { Location: '/?page=top' }); return res.end(); }
    const page = ['top', 'approver', 'admin', 'approve2', 'staffSummary'].includes(pageParam) ? pageParam : 'top';
    const common = { APP_URL: 'http://localhost:' + PORT + '/', PORTAL_URL: PORTAL_URL };
    let html;
    if (!ADMIN && (page === 'admin' || page === 'approve2')) {
      html = render('no_auth.html', Object.assign({ message: '総務部（管理者）権限がありません。\nApproverMap シートに role=admin で登録されているか、\nSettings シートの ADMIN_EMAILS にメールアドレスが含まれているか確認してください。\n（現在のアカウント: ' + USER.email + '）' }, common));
    } else {
      const data = { PAGE: page, DATE_PARAM: validDate(url.searchParams.get('date')), REASONS: REASONS, OT_HOURS: OT_HOURS, HD_HOURS: HD_HOURS };
      html = render(page + '.html', Object.assign({ data_json: jsonForScript(data) }, common));
    }
    return send(res, 200, html, 'text/html; charset=utf-8', { 'Content-Security-Policy': CSP });
  }
  return send(res, 404, 'not found', 'text/plain');
}).listen(PORT, () => {
  console.log('見本サーバー: http://localhost:' + PORT + '/?page=top  （' + (ADMIN ? '管理者' : '承認者') + 'として）');
});
