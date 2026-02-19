// ====== WorkLogs の行取得・更新 DAO ======
// WorkLogs のヘッダ行は「2行目」想定（1行目が注釈）

function findWorkLogRowNo_(requestId) {
  const sh = requireSheet_('WorkLogs');
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return -1;

  const header = sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0].map(h=>normalize_(h));
  const idx = buildHeaderIndex_(header);
  const ridCol = idx['requestId'];
  if (ridCol === undefined) throw new Error('WorkLogsに requestId 列がありません。');

  const values = sh.getRange(3,1,lastRow-2,sh.getLastColumn()).getValues();
  for (let i=0; i<values.length; i++) {
    if (normalize_(values[i][ridCol]) === requestId) return i + 3;
  }
  return -1;
}

function updateWorkLog_(requestId, patch) {
  const sh = requireSheet_('WorkLogs');
  const header = sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0].map(h=>normalize_(h));
  const idx = buildHeaderIndex_(header);

  let rowNo = findWorkLogRowNo_(requestId);
  if (rowNo === -1) {
    // 無ければ作る
    const row = new Array(header.length).fill('');
    row[idx['requestId']] = requestId;
    sh.appendRow(row);
    rowNo = sh.getLastRow();
  }

  for (const [key,val] of Object.entries(patch)) {
    if (idx[key] === undefined) continue;
    sh.getRange(rowNo, idx[key]+1).setValue(val);
  }
  return rowNo;
}

// ====== Requests から必要情報を取る ======

function getRequestById_(requestId) {
  const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const values = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  for (let i=0; i<values.length; i++) {
    const row = values[i];
    if (normalize_(row[idx['requestId']]) === requestId) {
      return {
        rowNo: i+2,
        requestId,
        requestType: normalize_(row[idx['requestType(overtime/holiday)']]),
        status: normalize_(row[idx['status(submitted/approved/canceled)']]),
        dept: normalize_(row[idx['dept']]),
        workerName: row[idx['workerName']],
        workerEmail: normalize_(row[idx['workerEmail']]),
        targetDate: row[idx['targetDate']],
        approvedMinutes: row[idx['approvedMinutes']],
        pdfGeneratedAt: row[idx['pdfGeneratedAt']],
        pdfFileId: row[idx['pdfFileId']],
        pdfFolderId: row[idx['pdfFolderId']],
      };
    }
  }
  return null;
}

// ====== 休憩控除（休憩マスタ方式） ======
// 休憩マスタ：適用区分(overtime/holiday) / 実働分_下限 / 実働分_上限 / 休憩分

function calcBreakMinutesByMaster_(type, actualMinutes) {
  const sh = requireSheet_('休憩マスタ');
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h=>normalize_(h));
  const idx = {
    type: H.indexOf('適用区分(overtime/holiday)'),
    min: H.indexOf('実働分_下限'),
    max: H.indexOf('実働分_上限'),
    brk: H.indexOf('休憩分'),
  };
  if (idx.type<0 || idx.min<0 || idx.brk<0) return 0;

  for (let r=1; r<values.length; r++) {
    const row = values[r];
    if (normalize_(row[idx.type]) !== type) continue;
    const min = Number(row[idx.min] ?? 0);
    const max = row[idx.max] === '' || row[idx.max] == null ? Infinity : Number(row[idx.max]);
    const brk = Number(row[idx.brk] ?? 0);
    if (actualMinutes >= min && actualMinutes <= max) return brk;
  }
  return 0;
}

// ====== 作業者本人チェック ======

function assertSelf_(request) {
  const email = Session.getActiveUser().getEmail();
  // workerEmailが入っている前提なら本人確認できる
  if (request.workerEmail && request.workerEmail !== email) {
    throw new Error('本人のみ操作可能です。');
  }
}

// ====== 残業：完了ボタン（17:20固定起点） ======

function api_markOvertimeDone(requestId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = getRequestById_(requestId);
    if (!req) throw new Error('申請が見つかりません。');
    if (req.requestType !== 'overtime') throw new Error('残業申請ではありません。');
    if (req.status === 'canceled') throw new Error('キャンセル済みです。');

    // 本人チェック（workerEmailが未設定ならスキップ）
    assertSelf_(req);

    const now = new Date();

    // start = targetDate 17:20（JST）
    const d = req.targetDate instanceof Date ? req.targetDate : new Date(req.targetDate);
    const start = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 17, 20, 0);

    const actualMinutes = Math.max(0, Math.round((now.getTime() - start.getTime()) / 60000));
    const breakMinutes = calcBreakMinutesByMaster_('overtime', actualMinutes);
    const netMinutes = Math.max(0, actualMinutes - breakMinutes);

    updateWorkLog_(requestId, {
      actualStartAt: start,
      actualEndAt: now,
      actualMinutes: actualMinutes,
      breakMinutes: breakMinutes,
      netMinutes: netMinutes,
      updatedAt: new Date(),
      updatedBy: Session.getActiveUser().getEmail() || 'unknown',
    });

    // 承認済みならPDF生成
    let pdf = null;
    if (req.status === 'approved') {
      pdf = generatePdfForRequest_(requestId);
    }

    return { ok:true, requestId, actualMinutes, breakMinutes, netMinutes, pdf };
  } finally {
    lock.releaseLock();
  }
}

// ====== 休日：開始ボタン ======

function api_markHolidayStart(requestId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = getRequestById_(requestId);
    if (!req) throw new Error('申請が見つかりません。');
    if (req.requestType !== 'holiday') throw new Error('休日申請ではありません。');
    if (req.status === 'canceled') throw new Error('キャンセル済みです。');

    assertSelf_(req);

    const now = new Date();
    updateWorkLog_(requestId, {
      actualStartAt: now,
      updatedAt: now,
      updatedBy: Session.getActiveUser().getEmail() || 'unknown',
    });

    return { ok:true, requestId, actualStartAt: fmtDate_(now, 'yyyy-MM-dd HH:mm:ss') };
  } finally {
    lock.releaseLock();
  }
}

// ====== 休日：完了ボタン（start/end で実績 → 休憩 → net） ======

function api_markHolidayDone(requestId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = getRequestById_(requestId);
    if (!req) throw new Error('申請が見つかりません。');
    if (req.requestType !== 'holiday') throw new Error('休日申請ではありません。');
    if (req.status === 'canceled') throw new Error('キャンセル済みです。');

    assertSelf_(req);

    // WorkLogsのstart取得
    const sh = requireSheet_('WorkLogs');
    const header = sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0].map(h=>normalize_(h));
    const idx = buildHeaderIndex_(header);
    const rowNo = findWorkLogRowNo_(requestId);
    if (rowNo === -1) throw new Error('休日開始が記録されていません（開始ボタンを押してください）。');

    const row = sh.getRange(rowNo,1,1,sh.getLastColumn()).getValues()[0];
    const start = row[idx['actualStartAt']];
    if (!(start instanceof Date)) throw new Error('休日開始が記録されていません（開始ボタンを押してください）。');

    const now = new Date();
    const actualMinutes = Math.max(0, Math.round((now.getTime() - start.getTime()) / 60000));
    const breakMinutes = calcBreakMinutesByMaster_('holiday', actualMinutes);
    const netMinutes = Math.max(0, actualMinutes - breakMinutes);

    updateWorkLog_(requestId, {
      actualEndAt: now,
      actualMinutes: actualMinutes,
      breakMinutes: breakMinutes,
      netMinutes: netMinutes,
      updatedAt: new Date(),
      updatedBy: Session.getActiveUser().getEmail() || 'unknown',
    });

    // 承認済みならPDF生成
    let pdf = null;
    if (req.status === 'approved') {
      pdf = generatePdfForRequest_(requestId);
    }

    return { ok:true, requestId, actualMinutes, breakMinutes, netMinutes, pdf };
  } finally {
    lock.releaseLock();
  }
}

// ====== 作業者の「本日の申請一覧」取得（後方互換用に残す） ======
// workerCode を指定 → その人の申請だけ返す
// 未指定（空文字） → 本日の全申請を返す

function api_getTodayRequestsForWorker(workerCode) {
  const filterCode = normalize_(workerCode || '');

  const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
  const values = sh.getDataRange().getValues();
  const today = fmtDate_(new Date(), 'yyyy-MM-dd');

  const wlMap = buildWorkLogsMapByRequestId_();
  const out = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (!status || status === 'canceled') continue;

    // 作業員コードでフィルタ（指定時のみ）
    if (filterCode) {
      const rowCode = normalize_(row[idx['workerCode']]);
      if (rowCode !== filterCode) continue;
    }

    const targetDateVal = row[idx['targetDate']];
    const targetDate = targetDateVal instanceof Date
      ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
      : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
    if (targetDate !== today) continue;

    const requestId = normalize_(row[idx['requestId']]);
    const wl = wlMap.get(requestId) || {};

    out.push({
      requestId,
      requestType: normalize_(row[idx['requestType(overtime/holiday)']]),
      status,
      dept: normalize_(row[idx['dept']]),
      workerCode: normalize_(row[idx['workerCode']]),
      workerName: normalize_(row[idx['workerName']]),
      targetDate,
      approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
      submittedAt: row[idx['submittedAt']],
      approvedAt: row[idx['approvedAt']],
      actualStartAt: wl.actualStartAt || '',
      actualEndAt: wl.actualEndAt || '',
      actualMinutes: Number(wl.actualMinutes || 0),
      breakMinutes: Number(wl.breakMinutes || 0),
      netMinutes: Number(wl.netMinutes || 0),
      pdfFileId: normalize_(row[idx['pdfFileId']]),
    });
  }

  return { today, items: out };
}

// ====== 全社ダッシュボード（残業=本日 / 休日=今週末） ======
// 全社員の申請を取得し、残業と休日に分けて返す。

function api_getDashboard() {
  try {
    const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
    const values = sh.getDataRange().getValues();
    const wlMap = buildWorkLogsMapByRequestId_();

    const now = new Date();
    const today = fmtDate_(now, 'yyyy-MM-dd');

    // 今週末の日付範囲を算出（土〜日＋振替月曜まで）
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

      const targetDateVal = row[idx['targetDate']];
      let targetDate;
      try {
        targetDate = targetDateVal instanceof Date
          ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
          : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
      } catch (e) { continue; }

      const requestType = normalize_(row[idx['requestType(overtime/holiday)']]);
      const requestId = normalize_(row[idx['requestId']]);
      const wl = wlMap.get(requestId) || {};

      // Date→文字列変換（GASシリアライズ問題回避）
      const actualStartAt = wl.actualStartAt instanceof Date
        ? fmtDate_(wl.actualStartAt, 'yyyy-MM-dd HH:mm:ss') : String(wl.actualStartAt || '');
      const actualEndAt = wl.actualEndAt instanceof Date
        ? fmtDate_(wl.actualEndAt, 'yyyy-MM-dd HH:mm:ss') : String(wl.actualEndAt || '');

      const item = {
        requestId: requestId || '',
        requestType: requestType || '',
        status: status || '',
        dept: normalize_(row[idx['dept']]) || '',
        workerCode: normalize_(row[idx['workerCode']]) || '',
        workerName: normalize_(row[idx['workerName']]) || '',
        targetDate: targetDate || '',
        targetDateLabel: '',
        approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
        actualStartAt: actualStartAt,
        actualEndAt: actualEndAt,
        actualMinutes: Number(wl.actualMinutes || 0),
        breakMinutes: Number(wl.breakMinutes || 0),
        netMinutes: Number(wl.netMinutes || 0),
        pdfFileId: normalize_(row[idx['pdfFileId']]) || '',
      };

      if (requestType === 'overtime' && targetDate === today) {
        overtime.push(item);
      } else if (requestType === 'holiday' && targetDate >= weekendStart && targetDate <= weekendEnd) {
        const d = new Date(targetDate + 'T00:00:00');
        item.targetDateLabel = (d.getMonth()+1) + '/' + d.getDate() + '(' + dayNames[d.getDay()] + ')';
        holiday.push(item);
      }
    }

    overtime.sort(function(a, b) {
      if (a.dept !== b.dept) return a.dept < b.dept ? -1 : 1;
      return a.workerName < b.workerName ? -1 : a.workerName > b.workerName ? 1 : 0;
    });

    holiday.sort(function(a, b) {
      if (a.targetDate !== b.targetDate) return a.targetDate < b.targetDate ? -1 : 1;
      if (a.dept !== b.dept) return a.dept < b.dept ? -1 : 1;
      return a.workerName < b.workerName ? -1 : a.workerName > b.workerName ? 1 : 0;
    });

    return {
      today: today, weekendStart: weekendStart, weekendEnd: weekendEnd,
      overtime: overtime, holiday: holiday,
    };
  } catch (err) {
    console.error('api_getDashboard エラー: ' + err.message + '\n' + err.stack);
    return { today: '', weekendStart: '', weekendEnd: '',
             overtime: [], holiday: [], error: err.message };
  }
}
