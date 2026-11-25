// 仕様（docs/vps-implementation-spec.md）4-3 の画面の 29 関数と、4-2 の管理者だけの 9 関数。試験が照らし合わせる元。
'use strict';

const SCREEN = [
  'api_getDeptList', 'api_getWorkersByDept', 'api_getWorkerInfo', 'api_getOrderList',
  'api_getHolidayCandidateDates', 'api_getDashboard', 'api_submitRequest', 'api_markOvertimeDone',
  'api_markHolidayStart', 'api_markHolidayDone', 'api_editRequest', 'api_deleteRequest',
  'api_getApproverDepts', 'api_getApproverDashboard', 'api_approveRequest', 'api_rejectRequest',
  'api_approveRequestsBatch', 'api_approverMonthlySummary', 'api_getSecondaryApprovalList', 'api_secondaryApprove',
  'api_secondaryApproveBatch', 'api_adminDeptOptions', 'api_adminMonthlySummary', 'api_adminFiscalTrend',
  'api_adminDailyDetail', 'api_adminManualComplete', 'api_adminSpecialClauseCount', 'api_adminPendingWatch',
  'api_getClosingSummary'
];
const OPS = [
  'ops_getStatus', 'ops_retryJob', 'ops_syncSheetConfigNow', 'ops_runMasterSync', 'ops_regeneratePdf',
  'ops_batchPdfForDate', 'ops_rebuildLedger', 'ops_recalcBreaks', 'ops_revokeSessions'
];
// 画面の「使われていないコード」だけが呼ぶ関数（仕様 4-1 で出さない。呼ばれることは無い）
const DEAD_CODE_ONLY = ['api_getFormUrl'];

module.exports = { SCREEN, OPS, ALL: SCREEN.concat(OPS), DEAD_CODE_ONLY };
