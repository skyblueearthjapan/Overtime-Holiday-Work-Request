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

function api_getApproverDashboard(dept) {
  try {
    const email = Session.getActiveUser().getEmail() || '';
    const admin = !email || isAdmin_(email);

    if (!dept) {
      return { today: '', weekendStart: '', weekendEnd: '',
               overtime: [], holiday: [], isAdmin: admin,
               error: '部署が指定されていません。' };
    }
    const normDept = normalize_(dept);

    if (!admin && !canApproveDept_(email, normDept)) {
      return { today: '', weekendStart: '', weekendEnd: '',
               overtime: [], holiday: [], isAdmin: false,
               error: 'この部署の承認権限がありません。' };
    }

    let allowedDepts = new Set([normDept]);

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
      if (!status || status === 'canceled') continue;

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
      };

      if (requestType === 'overtime' && targetDate === today) {
        overtime.push(item);
      } else if (requestType === 'holiday' && targetDate >= weekendStart && targetDate <= weekendEnd) {
        const d = new Date(targetDate + 'T00:00:00');
        item.targetDateLabel = (d.getMonth()+1) + '/' + d.getDate() + '(' + dayNames[d.getDay()] + ')';
        holiday.push(item);
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
             selectedDept: normDept };
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

function api_approverMonthlySummary(yearMonth, dept) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('メールアドレスが取得できません。');
  const normDept = normalize_(dept);
  if (!normDept) throw new Error('部署が指定されていません。');
  if (!canApproveDept_(email, normDept))
    throw new Error('この部署の月次サマリーを閲覧する権限がありません。');
  return buildMonthlySummary_(yearMonth, normDept);
}

// ====== 承認者の担当部署一覧取得 ======

function api_getApproverDepts() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('メールアドレスが取得できません。');

  // admin なら全部署
  if (isAdmin_(email)) {
    return { isAdmin: true, depts: loadDeptList_() };
  }

  // DeptApprovers / ApproverMap 両対応
  const map = getDeptApproverMap_();
  const emailLc = email.toLowerCase();
  const depts = [];
  for (const dept of Object.keys(map)) {
    if (map[dept].has(emailLc)) depts.push(dept);
  }

  if (depts.length === 0) {
    throw new Error('承認者権限がありません。DeptApprovers または ApproverMap に登録してください。');
  }

  return { isAdmin: false, depts };
}
