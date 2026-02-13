// ====== 承認権限チェック（ApproverMap） ======

function isAdmin_(email) {
  if (!email) return false;

  // 1. Settings の ADMIN_EMAILS もチェック（カンマ区切り複数可）
  try {
    const settings = getSettings_();
    const adminEmails = String(settings['ADMIN_EMAILS'] || '').split(',').map(s => s.trim().toLowerCase());
    if (adminEmails.indexOf(email.toLowerCase()) >= 0) return true;
  } catch (e) { /* Settings未設定でもOK */ }

  // 2. ApproverMap で role=admin チェック
  const sh = requireSheet_('ApproverMap');
  const values = sh.getDataRange().getValues();
  // ヘッダ想定：部署, 承認者メール, role(approver/admin), 有効フラグ
  const H = values[0].map(h => normalize_(h));
  const idx = {
    dept: H.indexOf('部署'),
    mail: H.indexOf('承認者メール'),
    role: H.indexOf('role(approver/admin)'),
    enabled: H.indexOf('有効フラグ'),
  };
  for (let r=1; r<values.length; r++) {
    const row = values[r];
    const mail = normalize_(row[idx.mail]);
    const role = normalize_(row[idx.role]);
    const enabled = idx.enabled >= 0 ? row[idx.enabled] : true;
    if (enabled === false || String(enabled).toLowerCase()==='false') continue;
    if (mail === email && role === 'admin') return true;
  }
  return false;
}

function canApproveDept_(email, dept) {
  if (isAdmin_(email)) return true;
  const sh = requireSheet_('ApproverMap');
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idx = {
    dept: H.indexOf('部署'),
    mail: H.indexOf('承認者メール'),
    role: H.indexOf('role(approver/admin)'),
    enabled: H.indexOf('有効フラグ'),
  };
  for (let r=1; r<values.length; r++) {
    const row = values[r];
    const d = normalize_(row[idx.dept]);
    const mail = normalize_(row[idx.mail]);
    const enabled = idx.enabled >= 0 ? row[idx.enabled] : true;
    if (enabled === false || String(enabled).toLowerCase()==='false') continue;
    if (mail === email && d === dept) return true;
  }
  return false;
}

// ====== 承認者画面用：部署別「本日申請」取得 ======

function api_getTodayRequestsForDept(dept) {
  const email = Session.getActiveUser().getEmail();
  if (!canApproveDept_(email, dept)) throw new Error('この部署の承認権限がありません。');

  const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
  const values = sh.getDataRange().getValues();

  const today = fmtDate_(new Date(), 'yyyy-MM-dd');
  const out = [];

  for (let r=1; r<values.length; r++) {
    const row = values[r];
    const status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (!status || status === 'canceled') continue;

    const rowDept = normalize_(row[idx['dept']]);
    if (rowDept !== dept) continue;

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
  return { today, dept, items: out };
}

// ====== 承認実行（承認ボタン） ======

function api_approveRequest(requestId) {
  const email = Session.getActiveUser().getEmail();
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

    return { ok: true, requestId, approvedBy: email, approvedAt: now };
  } finally {
    lock.releaseLock();
  }
}

// ====== 承認者の担当部署一覧取得 ======

function api_getApproverDepts() {
  const email = Session.getActiveUser().getEmail();
  const sh = requireSheet_('ApproverMap');
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idx = {
    dept: H.indexOf('部署'),
    mail: H.indexOf('承認者メール'),
    role: H.indexOf('role(approver/admin)'),
    enabled: H.indexOf('有効フラグ'),
  };

  const depts = [];
  let isAdminRole = false;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const mail = normalize_(row[idx.mail]);
    const enabled = idx.enabled >= 0 ? row[idx.enabled] : true;
    if (enabled === false || String(enabled).toLowerCase() === 'false') continue;
    if (mail !== email) continue;

    const role = normalize_(row[idx.role]);
    if (role === 'admin') isAdminRole = true;

    const dept = normalize_(row[idx.dept]);
    if (dept && !depts.includes(dept)) depts.push(dept);
  }

  if (isAdminRole) {
    return { isAdmin: true, depts: loadDeptList_() };
  }

  return { isAdmin: isAdminRole, depts };
}
