// ====== 承認権限チェック（ApproverMap） ======

function isAdmin_(email) {
  return isAdminWithDebug_(email).ok;
}

/**
 * デバッグ情報付き管理者チェック
 * @return {{ ok: boolean, debug: string }}
 */
function isAdminWithDebug_(email) {
  const log = [];
  if (!email) return { ok: false, debug: 'email が空です' };
  log.push('チェック対象: ' + email);

  // 1. Settings の ADMIN_EMAILS / GENERAL_AFFAIRS_EMAILS をチェック（カンマ区切り複数可）
  try {
    const settings = getSettings_();
    const raw = settings['ADMIN_EMAILS'];
    const gaRaw = settings['GENERAL_AFFAIRS_EMAILS'];
    log.push('Settings ADMIN_EMAILS 生値: [' + String(raw) + ']');
    log.push('Settings GENERAL_AFFAIRS_EMAILS 生値: [' + String(gaRaw) + ']');
    const combined = String(raw || '') + ',' + String(gaRaw || '');
    const adminEmails = combined.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    log.push('パース後: ' + JSON.stringify(adminEmails));
    if (adminEmails.indexOf(email.toLowerCase()) >= 0) {
      return { ok: true, debug: log.join('\n') };
    }
    log.push('Settings チェック: 不一致');
  } catch (e) {
    log.push('Settings エラー: ' + e.message);
  }

  // 2. ApproverMap で role=admin チェック
  try {
    const sh = requireSheet_('ApproverMap');
    const values = sh.getDataRange().getValues();
    log.push('ApproverMap 行数: ' + values.length);
    // ヘッダ想定：部署, 承認者メール, role(approver/admin), 有効フラグ
    const H = values[0].map(h => normalize_(h));
    log.push('ApproverMap ヘッダ(row1): ' + JSON.stringify(H));
    const idx = {
      dept: H.indexOf('部署'),
      mail: H.indexOf('承認者メール'),
      role: H.indexOf('role(approver/admin)'),
      enabled: H.indexOf('有効フラグ'),
    };
    log.push('ApproverMap idx: ' + JSON.stringify(idx));
    for (let r=1; r<values.length; r++) {
      const row = values[r];
      const mail = normalize_(row[idx.mail]);
      const role = normalize_(row[idx.role]);
      const enabled = idx.enabled >= 0 ? row[idx.enabled] : true;
      if (enabled === false || String(enabled).toLowerCase()==='false') continue;
      if (mail === email && role === 'admin') {
        return { ok: true, debug: log.join('\n') };
      }
    }
    log.push('ApproverMap チェック: 該当なし');
  } catch (e) {
    log.push('ApproverMap エラー: ' + e.message);
  }

  return { ok: false, debug: log.join('\n') };
}

function canApproveDept_(email, dept) {
  if (isAdmin_(email)) return true;

  // DeptApprovers / ApproverMap 両対応（getDeptApproverMap_ は deptAuth.gs）
  const map = getDeptApproverMap_();
  const set = map[dept];
  if (set && set.has(email.toLowerCase())) return true;

  return false;
}

// ====== 承認者画面用：部署別「本日申請」取得 ======

function api_getTodayRequestsForDept(dept) {
  const normDept = normalize_(dept);
  const email = Session.getActiveUser().getEmail();
  if (!canApproveDept_(email, normDept)) throw new Error('この部署の承認権限がありません。');

  const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
  const values = sh.getDataRange().getValues();

  const today = fmtDate_(new Date(), 'yyyy-MM-dd');
  const out = [];

  for (let r=1; r<values.length; r++) {
    const row = values[r];
    const status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (!status || status === 'canceled') continue;

    const rowDept = normalize_(row[idx['dept']]);
    if (rowDept !== normDept) continue;

    const targetDateVal = row[idx['targetDate']];
    const targetDate = targetDateVal instanceof Date ? fmtDate_(targetDateVal,'yyyy-MM-dd') : fmtDate_(new Date(targetDateVal),'yyyy-MM-dd');
    if (targetDate !== today) continue;

    out.push({
      requestId: row[idx['requestId']],
      requestType: row[idx['requestType(overtime/holiday)']],
      status,
      dept: rowDept,
      workerName: row[idx['workerName']],
      workerEmail: row[idx['workerEmail']],
      targetDate,
      approvedMinutes: row[idx['approvedMinutes']],
      submittedAt: row[idx['submittedAt']],
      approvedAt: row[idx['approvedAt']],
    });
  }
  return { today, dept: normDept, items: out };
}

// ====== 承認者ダッシュボード（全承認対象部署の残業+休日） ======

function api_getApproverDashboard(deptsOrDept) {
  try {
    const email = Session.getActiveUser().getEmail() || '';
    const admin = !email || isAdmin_(email);

    if (!deptsOrDept || (Array.isArray(deptsOrDept) && deptsOrDept.length === 0)) {
      return { today: '', weekendStart: '', weekendEnd: '',
               overtime: [], holiday: [], isAdmin: admin,
               error: '部署が指定されていません。' };
    }

    // 配列 or 文字列どちらも対応
    var deptList = Array.isArray(deptsOrDept)
      ? deptsOrDept.map(function(d){ return normalize_(d); })
      : [normalize_(deptsOrDept)];

    // 権限チェック（管理者でなければ各部署をチェック）
    if (!admin) {
      for (var i = 0; i < deptList.length; i++) {
        if (!canApproveDept_(email, deptList[i])) {
          return { today: '', weekendStart: '', weekendEnd: '',
                   overtime: [], holiday: [], isAdmin: false,
                   error: '部署「' + deptList[i] + '」の承認権限がありません。' };
        }
      }
    }

    var allowedDepts = new Set(deptList);

    const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
    const values = sh.getDataRange().getValues();

    const now = new Date();
    const today = fmtDate_(now, 'yyyy-MM-dd');

    // 今週末の日付範囲
    const dow = now.getDay();
    const monday = new Date(now);
    monday.setDate(now.getDate() - (dow === 0 ? 6 : dow - 1));
    monday.setHours(0, 0, 0, 0);
    const saturday = new Date(monday);
    saturday.setDate(monday.getDate() + 5);
    const nextMonday = new Date(monday);
    nextMonday.setDate(monday.getDate() + 7);
    const weekendStart = fmtDate_(saturday, 'yyyy-MM-dd');
    const weekendEnd = fmtDate_(nextMonday, 'yyyy-MM-dd');

    const dayNames = ['日','月','火','水','木','金','土'];
    const overtime = [];
    const holiday = [];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const status = normalize_(row[idx['status(submitted/approved/canceled)']]);
      if (!status) continue;

      const rowDept = normalize_(row[idx['dept']]);
      if (allowedDepts && !allowedDepts.has(rowDept)) continue;

      // targetDate のパース（不正な日付はスキップ）
      const targetDateVal = row[idx['targetDate']];
      let targetDate;
      try {
        targetDate = targetDateVal instanceof Date
          ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
          : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
      } catch (e) {
        continue; // 不正な日付はスキップ
      }

      const requestType = normalize_(row[idx['requestType(overtime/holiday)']]);
      const requestId = normalize_(row[idx['requestId']]);

      // submittedAt を文字列に変換（シリアライズ問題回避）
      const submittedAtRaw = row[idx['submittedAt']];
      let submittedAt = '';
      if (submittedAtRaw instanceof Date) {
        submittedAt = fmtDate_(submittedAtRaw, 'yyyy-MM-dd HH:mm:ss');
      } else if (submittedAtRaw) {
        submittedAt = String(submittedAtRaw);
      }

      // 未処理の過去申請にはアラームフラグ
      var isOverdue = false;

      const item = {
        requestId: requestId || '',
        requestType: requestType || '',
        status: status || '',
        dept: rowDept || '',
        workerName: normalize_(row[idx['workerName']]) || '',
        targetDate: targetDate || '',
        targetDateLabel: '',
        approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
        submittedAt: submittedAt,
        isOverdue: false,
      };

      if (requestType === 'overtime') {
        if (status === 'submitted') {
          // 未処理: targetDateが今日以前なら全て表示（過去分はアラーム）
          if (targetDate <= today) {
            if (targetDate < today) {
              item.isOverdue = true;
            }
            overtime.push(item);
          }
        } else {
          // approved/canceled: 当日のみ表示
          if (targetDate === today) {
            overtime.push(item);
          }
        }
      } else if (requestType === 'holiday') {
        if (status === 'submitted') {
          // 未処理: targetDateが今日以前 or 今週末範囲なら表示
          if (targetDate <= today) {
            if (targetDate < today) {
              item.isOverdue = true;
            }
            const d = new Date(targetDate + 'T00:00:00');
            item.targetDateLabel = (d.getMonth()+1) + '/' + d.getDate() + '(' + dayNames[d.getDay()] + ')';
            holiday.push(item);
          } else if (targetDate >= weekendStart && targetDate <= weekendEnd) {
            const d = new Date(targetDate + 'T00:00:00');
            item.targetDateLabel = (d.getMonth()+1) + '/' + d.getDate() + '(' + dayNames[d.getDay()] + ')';
            holiday.push(item);
          }
        } else {
          // approved/canceled: 当日 or 今週末範囲で当日のみ
          if (targetDate === today) {
            const d = new Date(targetDate + 'T00:00:00');
            item.targetDateLabel = (d.getMonth()+1) + '/' + d.getDate() + '(' + dayNames[d.getDay()] + ')';
            holiday.push(item);
          } else if (targetDate >= weekendStart && targetDate <= weekendEnd) {
            const d = new Date(targetDate + 'T00:00:00');
            item.targetDateLabel = (d.getMonth()+1) + '/' + d.getDate() + '(' + dayNames[d.getDay()] + ')';
            holiday.push(item);
          }
        }
      }
    }

    // ソート：未承認を先頭、次に部署→名前
    const sortFn = function(a, b) {
      if (a.status !== b.status) {
        if (a.status === 'submitted') return -1;
        if (b.status === 'submitted') return 1;
      }
      if (a.dept !== b.dept) return a.dept < b.dept ? -1 : 1;
      return a.workerName < b.workerName ? -1 : a.workerName > b.workerName ? 1 : 0;
    };
    overtime.sort(sortFn);
    holiday.sort(function(a, b) {
      if (a.status !== b.status) {
        if (a.status === 'submitted') return -1;
        if (b.status === 'submitted') return 1;
      }
      if (a.targetDate !== b.targetDate) return a.targetDate < b.targetDate ? -1 : 1;
      if (a.dept !== b.dept) return a.dept < b.dept ? -1 : 1;
      return a.workerName < b.workerName ? -1 : a.workerName > b.workerName ? 1 : 0;
    });

    return { today: today, weekendStart: weekendStart, weekendEnd: weekendEnd,
             overtime: overtime, holiday: holiday, isAdmin: admin,
             selectedDepts: deptList };
  } catch (err) {
    // エラーをクライアントに伝えるため、エラー情報を含むオブジェクトを返す
    console.error('api_getApproverDashboard エラー: ' + err.message + '\n' + err.stack);
    return { today: '', weekendStart: '', weekendEnd: '',
             overtime: [], holiday: [], isAdmin: true,
             error: err.message };
  }
}

// ====== 承認実行（承認ボタン） ======

function api_approveRequest(requestId) {
  const email = Session.getActiveUser().getEmail() || 'unknown';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('Requestsが空です。');

    const values = sh.getRange(2, 1, lastRow-1, sh.getLastColumn()).getValues();
    let rowNo = -1;
    let dept = '';
    let status = '';

    for (let i=0; i<values.length; i++) {
      const row = values[i];
      if (normalize_(row[idx['requestId']]) === requestId) {
        rowNo = i + 2;
        dept = normalize_(row[idx['dept']]);
        status = normalize_(row[idx['status(submitted/approved/canceled)']]);
        break;
      }
    }
    if (rowNo === -1) throw new Error('requestIdが見つかりません。');
    if (!canApproveDept_(email, dept)) throw new Error('この申請の承認権限がありません。');
    if (status === 'approved') return { ok: true, message: '既に承認済みです。' };
    if (status === 'canceled') throw new Error('キャンセル済みです。');

    const now = new Date();
    sh.getRange(rowNo, idx['status(submitted/approved/canceled)']+1).setValue('approved');
    if (idx['approvedAt'] !== undefined) sh.getRange(rowNo, idx['approvedAt']+1).setValue(now);
    if (idx['approvedBy'] !== undefined) sh.getRange(rowNo, idx['approvedBy']+1).setValue(email);

    return { ok: true, requestId, approvedBy: email, approvedAt: fmtDate_(now, 'yyyy-MM-dd HH:mm:ss') };
  } finally {
    lock.releaseLock();
  }
}

// ====== 却下API ======

function api_rejectRequest(requestId) {
  var email = Session.getActiveUser().getEmail() || 'unknown';
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var { sh, idx } = getSheetHeaderIndex_('Requests', 1);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('Requestsが空です。');

    var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    var rowNo = -1;
    var dept = '';
    var status = '';

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (normalize_(row[idx['requestId']]) === requestId) {
        rowNo = i + 2;
        dept = normalize_(row[idx['dept']]);
        status = normalize_(row[idx['status(submitted/approved/canceled)']]);
        break;
      }
    }
    if (rowNo === -1) throw new Error('requestIdが見つかりません。');
    if (!canApproveDept_(email, dept)) throw new Error('この申請の承認権限がありません。');
    if (status !== 'submitted') throw new Error('未承認状態でのみ却下可能です（現在: ' + status + '）。');

    var now = new Date();
    sh.getRange(rowNo, idx['status(submitted/approved/canceled)'] + 1).setValue('canceled');
    if (idx['approvedBy'] !== undefined) sh.getRange(rowNo, idx['approvedBy'] + 1).setValue(email + '(却下)');
    if (idx['approvedAt'] !== undefined) sh.getRange(rowNo, idx['approvedAt'] + 1).setValue(now);

    return { ok: true, requestId: requestId };
  } finally {
    lock.releaseLock();
  }
}

// ====== 一括承認API ======

function api_approveRequestsBatch(requestIds) {
  if (!requestIds || requestIds.length === 0) return { ok: true, results: [] };
  if (requestIds.length > 50) throw new Error('一括承認は50件までです。');

  const email = Session.getActiveUser().getEmail() || 'unknown';
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, results: [] };

    const numCols = sh.getLastColumn();
    const values = sh.getRange(2, 1, lastRow-1, numCols).getValues();
    const now = new Date();
    const results = [];
    const targetSet = {};
    for (var k = 0; k < requestIds.length; k++) targetSet[requestIds[k]] = true;

    // 一括書き込み用：{ rowNo → { colIdx → value } }
    const pendingWrites = [];

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var rid = normalize_(row[idx['requestId']]);
      if (!rid || !targetSet[rid]) continue;

      var rowNo = i + 2;
      var dept = normalize_(row[idx['dept']]);
      var status = normalize_(row[idx['status(submitted/approved/canceled)']]);

      if (status === 'approved') { results.push({requestId:rid,ok:true,message:'既に承認済み'}); delete targetSet[rid]; continue; }
      if (status === 'canceled') { results.push({requestId:rid,ok:false,error:'キャンセル済み'}); delete targetSet[rid]; continue; }
      if (!canApproveDept_(email, dept)) { results.push({requestId:rid,ok:false,error:'承認権限なし'}); delete targetSet[rid]; continue; }

      pendingWrites.push({ rowNo: rowNo });
      results.push({requestId:rid,ok:true,approvedBy:email,approvedAt:fmtDate_(now,'yyyy-MM-dd HH:mm:ss')});
      delete targetSet[rid];
    }

    // setValues一括書き込み（行ごとにまとめて1回のsetValueで処理）
    var statusCol = idx['status(submitted/approved/canceled)'] + 1;
    var atCol = idx['approvedAt'] !== undefined ? idx['approvedAt'] + 1 : -1;
    var byCol = idx['approvedBy'] !== undefined ? idx['approvedBy'] + 1 : -1;
    for (var w = 0; w < pendingWrites.length; w++) {
      var rn = pendingWrites[w].rowNo;
      sh.getRange(rn, statusCol).setValue('approved');
      if (atCol > 0) sh.getRange(rn, atCol).setValue(now);
      if (byCol > 0) sh.getRange(rn, byCol).setValue(email);
    }
    if (pendingWrites.length > 0) SpreadsheetApp.flush();

    for (var missing in targetSet) results.push({requestId:missing,ok:false,error:'見つかりません'});
    return { ok: true, results: results };
  } finally { lock.releaseLock(); }
}

// ====== 承認者向け月次サマリーAPI ======

function api_approverMonthlySummary(yearMonth, deptsOrDept) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('メールアドレスが取得できません。');

  // 配列 or 文字列どちらも対応
  var deptList = Array.isArray(deptsOrDept)
    ? deptsOrDept.map(function(d){ return normalize_(d); }).filter(Boolean)
    : [normalize_(deptsOrDept)];

  if (deptList.length === 0 || !deptList[0]) throw new Error('部署が指定されていません。');

  // 権限チェック
  for (var i = 0; i < deptList.length; i++) {
    if (!canApproveDept_(email, deptList[i]))
      throw new Error('部署「' + deptList[i] + '」の月次サマリーを閲覧する権限がありません。');
  }

  // 複数部署の場合、buildMonthlySummary_を'ALL'で呼んでから部署フィルタする
  if (deptList.length === 1) {
    return buildMonthlySummary_(yearMonth, deptList[0]);
  }

  // 複数部署: ALLで取得してから対象部署のみにフィルタ
  var allData = buildMonthlySummary_(yearMonth, 'ALL');
  var deptSet = {};
  for (var j = 0; j < deptList.length; j++) deptSet[deptList[j]] = true;

  allData.people = allData.people.filter(function(p){ return deptSet[p.dept]; });
  allData.watch = allData.watch.filter(function(w){ return deptSet[w.dept]; });
  allData.deptFilter = deptList.join(', ');

  // KPI再計算
  var LIMIT40 = 2400, LIMIT60 = 3600;
  allData.kpi.totalPeople = allData.people.length;
  allData.kpi.over40 = allData.people.filter(function(p){ return p.totalNet >= LIMIT40; }).length;
  allData.kpi.over60 = allData.people.filter(function(p){ return p.totalNet >= LIMIT60; }).length;
  allData.kpi.projectedOver40 = allData.people.filter(function(p){ return p.projectedOver40; }).length;
  allData.kpi.projectedOver60 = allData.people.filter(function(p){ return p.projectedOver60; }).length;
  allData.kpi.pdfMissingTotal = allData.people.reduce(function(s,p){ return s + p.pdfMissing; }, 0);
  allData.kpi.totalNetMinutes = allData.people.reduce(function(s,p){ return s + p.totalNet; }, 0);

  // チャート再計算
  allData.charts.people = {
    labels: allData.people.map(function(p){ return p.workerName; }),
    values: allData.people.map(function(p){ return p.totalNet; }),
    projected: allData.people.map(function(p){ return p.projectedTotal; }),
    limit40: LIMIT40,
    limit60: LIMIT60,
  };

  return allData;
}

// ====== 承認者の担当部署一覧取得（承認者プロファイルベース） ======

function api_getApproverDepts() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('メールアドレスが取得できません。');
  const emailLc = email.toLowerCase();
  const profiles = getApproverProfiles_();

  if (isAdmin_(email)) {
    return { isAdmin: true, approvers: profiles, currentEmail: emailLc };
  }

  // 自分のプロファイルのみ返す
  const mine = profiles.filter(function(p){ return p.email === emailLc; });
  if (mine.length === 0) {
    throw new Error('承認者権限がありません。DeptApprovers または ApproverMap に登録してください。');
  }
  return { isAdmin: false, approvers: mine, currentEmail: emailLc };
}

// ====== 二次承認API ======

/**
 * Requestsシートに approvedBy2 / approvedAt2 列がなければ自動追加し、
 * 最新のヘッダーインデックスを返す。
 * ※ getSheetHeaderIndex_ はキャッシュを使うため、列追加後はキャッシュを破棄する。
 */
function ensureSecondaryApprovalColumns_() {
  var sh = requireSheet_('Requests');
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h){ return normalize_(h); });
  var idx = buildHeaderIndex_(header);
  var changed = false;

  if (idx['approvedBy2'] === undefined) {
    var newCol = header.length;
    sh.getRange(1, newCol + 1).setValue('approvedBy2');
    header.push('approvedBy2');
    idx['approvedBy2'] = newCol;
    changed = true;
  }
  if (idx['approvedAt2'] === undefined) {
    var newCol2 = header.length;
    sh.getRange(1, newCol2 + 1).setValue('approvedAt2');
    header.push('approvedAt2');
    idx['approvedAt2'] = newCol2;
    changed = true;
  }
  if (changed) {
    SpreadsheetApp.flush();
    // getSheetHeaderIndex_ のキャッシュを破棄
    _headerCache = {};
  }

  return { sh: sh, idx: idx };
}

/**
 * 指定日の全申請一覧を取得（二次承認画面用）
 * @param {string} dateYmd - 'yyyy-MM-dd'
 */
function api_getSecondaryApprovalList(dateYmd) {
  assertAdmin_();

  var d = new Date(dateYmd);
  if (isNaN(d.getTime())) throw new Error('日付が不正です（yyyy-MM-dd）');

  var ymd = fmtDate_(d, 'yyyy-MM-dd');
  var { sh, idx } = ensureSecondaryApprovalColumns_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { date: ymd, items: [] };

  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (!status || status === 'canceled') continue;

    // 二次承認済みはスキップ（非表示）
    var approvedBy2Val = idx['approvedBy2'] !== undefined ? normalize_(row[idx['approvedBy2']]) : '';
    if (approvedBy2Val) continue;

    var targetDateVal = row[idx['targetDate']];
    var targetYmd;
    try {
      targetYmd = targetDateVal instanceof Date
        ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
        : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
    } catch (e) { continue; }
    // 未2次承認: 対象日以前の申請も表示（期限超過として）
    // 2次承認済みは既にスキップ済み（L572）なので、ここでは未承認分のみ
    if (targetYmd > ymd) continue;  // 未来の申請は除外

    var submittedAtRaw = row[idx['submittedAt']];
    var submittedAt = '';
    if (submittedAtRaw instanceof Date) {
      submittedAt = fmtDate_(submittedAtRaw, 'yyyy-MM-dd HH:mm');
    } else if (submittedAtRaw) {
      submittedAt = String(submittedAtRaw);
    }

    var approvedAtRaw = row[idx['approvedAt']];
    var approvedAt = '';
    if (approvedAtRaw instanceof Date) {
      approvedAt = fmtDate_(approvedAtRaw, 'yyyy-MM-dd HH:mm');
    } else if (approvedAtRaw) {
      approvedAt = String(approvedAtRaw);
    }

    var approvedAt2Raw = idx['approvedAt2'] !== undefined ? row[idx['approvedAt2']] : '';
    var approvedAt2 = '';
    if (approvedAt2Raw instanceof Date) {
      approvedAt2 = fmtDate_(approvedAt2Raw, 'yyyy-MM-dd HH:mm');
    } else if (approvedAt2Raw) {
      approvedAt2 = String(approvedAt2Raw);
    }

    out.push({
      requestId: normalize_(row[idx['requestId']]),
      requestType: normalize_(row[idx['requestType(overtime/holiday)']]),
      status: status,
      dept: normalize_(row[idx['dept']]),
      workerName: normalize_(row[idx['workerName']]),
      targetDate: targetYmd,
      isOverdue: targetYmd < ymd,  // 対象日より前の申請は期限超過
      approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
      submittedAt: submittedAt,
      approvedBy: idx['approvedBy'] !== undefined ? normalize_(row[idx['approvedBy']]) : '',
      approvedAt: approvedAt,
      approvedBy2: approvedBy2Val,
      approvedAt2: approvedAt2,
      pdfFileId: idx['pdfFileId'] !== undefined ? normalize_(row[idx['pdfFileId']]) : '',
    });
  }

  // ソート: 期限超過 → 二次未承認＆一次承認済み → 日付（古い順） → 部署 → 氏名
  out.sort(function(a, b) {
    // 1. 期限超過を先頭
    var aOver = a.isOverdue ? 0 : 1;
    var bOver = b.isOverdue ? 0 : 1;
    if (aOver !== bOver) return aOver - bOver;
    // 2. 二次未承認 & 一次承認済みを優先
    var aReady = (a.status === 'approved' && !a.approvedBy2) ? 0 : 1;
    var bReady = (b.status === 'approved' && !b.approvedBy2) ? 0 : 1;
    if (aReady !== bReady) return aReady - bReady;
    // 3. 日付（古い順）
    if (a.targetDate !== b.targetDate) return a.targetDate < b.targetDate ? -1 : 1;
    // 4. 部署 → 氏名
    if (a.dept !== b.dept) return a.dept < b.dept ? -1 : 1;
    return a.workerName < b.workerName ? -1 : a.workerName > b.workerName ? 1 : 0;
  });

  return { date: ymd, items: out };
}

/**
 * 二次承認（個別）
 * @param {string} requestId
 */
function api_secondaryApprove(requestId) {
  var email = assertAdmin_();
  var approvedAt2Str = '';
  var hadExistingPdf = false; // 作業完了済み（PDF生成済み）かどうか

  // ---- Phase 1: 二次承認書き込み（ロック内） ----
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var { sh, idx } = ensureSecondaryApprovalColumns_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('Requestsが空です。');

    var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    var rowNo = -1;
    var status = '';
    var existingApprovedBy2 = '';

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (normalize_(row[idx['requestId']]) === requestId) {
        rowNo = i + 2;
        status = normalize_(row[idx['status(submitted/approved/canceled)']]);
        existingApprovedBy2 = idx['approvedBy2'] !== undefined ? normalize_(row[idx['approvedBy2']]) : '';
        // 既存PDFがあるか確認（作業完了済み＝PDF生成済み）
        hadExistingPdf = !!(idx['pdfFileId'] !== undefined && normalize_(row[idx['pdfFileId']]));
        break;
      }
    }

    if (rowNo === -1) throw new Error('requestIdが見つかりません。');
    if (status !== 'approved') throw new Error('一次承認が完了していません（現在: ' + status + '）。');
    if (existingApprovedBy2) return { ok: true, message: '既に二次承認済みです。' };

    var now = new Date();
    approvedAt2Str = fmtDate_(now, 'yyyy-MM-dd HH:mm:ss');
    if (idx['approvedBy2'] !== undefined) sh.getRange(rowNo, idx['approvedBy2'] + 1).setValue(email);
    if (idx['approvedAt2'] !== undefined) sh.getRange(rowNo, idx['approvedAt2'] + 1).setValue(now);

    // 既存PDFがある場合のみ削除＆クリア（作業完了前はPDF無し→何もしない）
    if (hadExistingPdf) {
      clearExistingPdfForRegeneration_(sh, idx, values[rowNo - 2], rowNo);
    }

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // ---- Phase 2: 既存PDFがあった場合のみ即時再生成（二次承認印付きで上書き） ----
  var pdfResult = null;
  if (hadExistingPdf) {
    try {
      _headerCache = {};
      pdfResult = generatePdfDirect_(requestId);
      Logger.log('二次承認後PDF再生成成功: ' + JSON.stringify(pdfResult));
    } catch (e) {
      Logger.log('二次承認後PDF再生成エラー: ' + e.message);
    }
  } else {
    Logger.log('二次承認: PDF未生成（作業完了前）→ 作業完了時に二次承認印付きPDFが生成されます');
  }

  // 蓄積SSにリアルタイム書き込み（失敗しても処理続行）
  upsertAccumulationRow_(requestId);

  return {
    ok: true,
    requestId: requestId,
    approvedBy2: email,
    approvedAt2: approvedAt2Str,
    pdfRegenerated: pdfResult ? (pdfResult.ok || false) : false,
  };
}

/**
 * 二次承認（一括）
 * @param {string[]} requestIds
 */
function api_secondaryApproveBatch(requestIds) {
  if (!requestIds || requestIds.length === 0) return { ok: true, results: [] };
  if (requestIds.length > 50) throw new Error('一括承認は50件までです。');

  var email = assertAdmin_();
  var pdfTargetIds = []; // PDF再生成が必要なrequestId（既存PDFありの分だけ）

  // ---- Phase 1: 二次承認書き込み（ロック内） ----
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var results = [];
  try {
    var { sh, idx } = ensureSecondaryApprovalColumns_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, results: [] };

    var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    var now = new Date();
    var targetSet = {};
    for (var k = 0; k < requestIds.length; k++) targetSet[requestIds[k]] = true;

    var pendingWrites = [];

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var rid = normalize_(row[idx['requestId']]);
      if (!rid || !targetSet[rid]) continue;

      var rowNo = i + 2;
      var status = normalize_(row[idx['status(submitted/approved/canceled)']]);
      var existingApprovedBy2 = idx['approvedBy2'] !== undefined ? normalize_(row[idx['approvedBy2']]) : '';

      if (existingApprovedBy2) {
        results.push({ requestId: rid, ok: true, message: '既に二次承認済み' });
        delete targetSet[rid];
        continue;
      }
      if (status !== 'approved') {
        results.push({ requestId: rid, ok: false, error: '一次承認未完了' });
        delete targetSet[rid];
        continue;
      }

      // 既存PDFがあるか（作業完了済みか）
      var hasPdf = !!(idx['pdfFileId'] !== undefined && normalize_(row[idx['pdfFileId']]));
      pendingWrites.push({ rowNo: rowNo, rowIdx: i, requestId: rid, hasPdf: hasPdf });
      results.push({ requestId: rid, ok: true, approvedBy2: email, approvedAt2: fmtDate_(now, 'yyyy-MM-dd HH:mm:ss') });
      delete targetSet[rid];
    }

    // 一括書き込み
    var by2Col = idx['approvedBy2'] !== undefined ? idx['approvedBy2'] + 1 : -1;
    var at2Col = idx['approvedAt2'] !== undefined ? idx['approvedAt2'] + 1 : -1;
    for (var w = 0; w < pendingWrites.length; w++) {
      var rn = pendingWrites[w].rowNo;
      if (by2Col > 0) sh.getRange(rn, by2Col).setValue(email);
      if (at2Col > 0) sh.getRange(rn, at2Col).setValue(now);

      // 既存PDFがある場合のみ削除＆クリア
      if (pendingWrites[w].hasPdf) {
        clearExistingPdfForRegeneration_(sh, idx, values[pendingWrites[w].rowIdx], rn);
        pdfTargetIds.push(pendingWrites[w].requestId);
      }
    }
    if (pendingWrites.length > 0) SpreadsheetApp.flush();

    for (var missing in targetSet) results.push({ requestId: missing, ok: false, error: '見つかりません' });
  } finally {
    lock.releaseLock();
  }

  // ---- Phase 2: 既存PDFがあった分のみ即時再生成（ロック外） ----
  if (pdfTargetIds.length > 0) {
    _headerCache = {};
    for (var p = 0; p < pdfTargetIds.length; p++) {
      try {
        generatePdfDirect_(pdfTargetIds[p]);
        Logger.log('一括二次承認PDF再生成成功: ' + pdfTargetIds[p]);
      } catch (e) {
        Logger.log('一括二次承認PDF再生成エラー(' + pdfTargetIds[p] + '): ' + e.message);
      }
      if (pdfTargetIds.length > 5) Utilities.sleep(500);
    }
  }

  // ---- Phase 3: 蓄積SSにリアルタイム書き込み（失敗しても処理続行） ----
  for (var q = 0; q < results.length; q++) {
    if (results[q].ok && results[q].approvedBy2) {
      upsertAccumulationRow_(results[q].requestId);
    }
  }

  return { ok: true, results: results };
}

/**
 * 二次承認時に既存PDFをゴミ箱に移し、pdfフィールドをクリアして翌朝バッチで再生成させる
 */
function clearExistingPdfForRegeneration_(sh, idx, row, rowNo) {
  var existingPdfId = idx['pdfFileId'] !== undefined ? normalize_(row[idx['pdfFileId']]) : '';
  if (existingPdfId) {
    try { DriveApp.getFileById(existingPdfId).setTrashed(true); } catch (e) {
      Logger.log('旧PDF削除スキップ: ' + e.message);
    }
  }
  if (idx['pdfGeneratedAt'] !== undefined) sh.getRange(rowNo, idx['pdfGeneratedAt'] + 1).setValue('');
  if (idx['pdfFileId'] !== undefined) sh.getRange(rowNo, idx['pdfFileId'] + 1).setValue('');
  if (idx['pdfFolderId'] !== undefined) sh.getRange(rowNo, idx['pdfFolderId'] + 1).setValue('');
}
