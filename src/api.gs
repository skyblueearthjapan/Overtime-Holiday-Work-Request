// ====== Sebastian (Hermès Agent) 連携用 外部API ======
// doPost で受けた JSON リクエストを処理し、結果を JSON で返す。
// 認証はスクリプトプロパティ SEBASTIAN_API_TOKEN との一致で行う。
//
// セットアップ:
//   1. PropertiesService.getScriptProperties().setProperty('SEBASTIAN_API_TOKEN', '長いランダム文字列')
//   2. このスクリプトを Web アプリとしてデプロイ（実行: 自分、アクセス: 全員）
//   3. 取得したURLとトークンを Hermès 側 .env に登録

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return _apiResponse_({ error: 'invalid JSON: ' + err.message });
  }

  const expected = PropertiesService.getScriptProperties().getProperty('SEBASTIAN_API_TOKEN');
  if (!expected) {
    return _apiResponse_({ error: 'server misconfigured: SEBASTIAN_API_TOKEN not set' });
  }
  if (body.token !== expected) {
    return _apiResponse_({ error: 'unauthorized' });
  }

  const action = String(body.action || '');
  try {
    let result;
    switch (action) {
      case 'ping':
        result = { ok: true, pong: new Date().toISOString() };
        break;
      case 'completeOvertime':
        if (!body.requestId) throw new Error('requestId is required');
        result = sebastian_markOvertimeDone_(String(body.requestId));
        break;
      case 'markHolidayStart':
        if (!body.requestId) throw new Error('requestId is required');
        result = sebastian_markHolidayStart_(
          String(body.requestId),
          String(body.manualStartTime || '')
        );
        break;
      case 'markHolidayDone':
        if (!body.requestId) throw new Error('requestId is required');
        result = sebastian_markHolidayDone_(
          String(body.requestId),
          String(body.manualEndTime || '')
        );
        break;
      default:
        throw new Error('unknown action: ' + action);
    }
    return _apiResponse_(result);
  } catch (err) {
    return _apiResponse_({
      error: String((err && err.message) || err),
      action: action,
    });
  }
}

function _apiResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====== 残業：完了処理（Sebastian経由） ======
// オリジナル api_markOvertimeDone と同等。assertSelf_ を省略（Sebastianが本人代理で実行するため）。
// PDF生成は status === 'approved' の場合のみ実行（既存仕様と同じ）。

function sebastian_markOvertimeDone_(requestId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = getRequestById_(requestId);
    if (!req) throw new Error('申請が見つかりません: ' + requestId);
    if (req.requestType !== 'overtime') throw new Error('残業申請ではありません: ' + requestId);
    if (req.status === 'canceled') throw new Error('キャンセル済みの申請です: ' + requestId);

    const now = new Date();

    // start = targetDate 17:20（JST）
    const d = req.targetDate instanceof Date ? req.targetDate : new Date(req.targetDate);
    const start = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 17, 20, 0);

    const actualMinutes = Math.max(0, Math.round((now.getTime() - start.getTime()) / 60000));
    const breakMinutes = calcBreakMinutesByMaster_('overtime', actualMinutes);
    const netMinutes = Math.max(0, actualMinutes - breakMinutes);

    updateWorkLog_(requestId, {
      actualStartAt: fmtDate_(start, 'yyyy-MM-dd HH:mm:ss'),
      actualEndAt:   fmtDate_(now,   'yyyy-MM-dd HH:mm:ss'),
      actualMinutes: actualMinutes,
      breakMinutes:  breakMinutes,
      netMinutes:    netMinutes,
      updatedAt:     fmtDate_(new Date(), 'yyyy-MM-dd HH:mm:ss'),
      updatedBy:     'sebastian-agent',
    });

    let pdf = null;
    if (req.status === 'approved') {
      try {
        pdf = generatePdfForRequest_(requestId);
      } catch (e) {
        // PDF失敗は致命的ではない。エラー情報だけ返して続行。
        pdf = { error: String(e.message || e) };
      }
    }

    // 蓄積SSにリアルタイム書き込み（失敗しても処理続行）
    upsertAccumulationRow_(requestId);

    return {
      ok: true,
      requestId: requestId,
      workerName: req.workerName,
      targetDate: req.targetDate,
      actualMinutes: actualMinutes,
      breakMinutes: breakMinutes,
      netMinutes: netMinutes,
      status: req.status,
      pdf: pdf,
    };
  } finally {
    lock.releaseLock();
  }
}

// ====== 休日出勤：作業開始（Sebastian経由） ======
// 既存 api_markHolidayStart と同等。assertSelf_ を省略（Sebastianが本人代理で実行するため）。
// manualStartTime: HH:mm形式（省略時は現在時刻）

function sebastian_markHolidayStart_(requestId, manualStartTime) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = getRequestById_(requestId);
    if (!req) throw new Error('申請が見つかりません: ' + requestId);
    if (req.requestType !== 'holiday') throw new Error('休日出勤申請ではありません: ' + requestId);
    if (req.status === 'canceled') throw new Error('キャンセル済みの申請です: ' + requestId);

    var startTime;
    if (manualStartTime && /^\d{2}:\d{2}$/.test(manualStartTime)) {
      const d = req.targetDate instanceof Date ? req.targetDate : new Date(req.targetDate);
      const targetYmd = fmtDate_(d, 'yyyy-MM-dd');
      startTime = new Date(targetYmd + 'T' + manualStartTime + ':00+09:00');
    } else {
      startTime = new Date();
    }

    updateWorkLog_(requestId, {
      actualStartAt: fmtDate_(startTime, 'yyyy-MM-dd HH:mm:ss'),
      updatedAt:     fmtDate_(new Date(), 'yyyy-MM-dd HH:mm:ss'),
      updatedBy:     'sebastian-agent',
    });

    return {
      ok: true,
      requestId: requestId,
      workerName: req.workerName,
      targetDate: req.targetDate,
      actualStartAt: fmtDate_(startTime, 'yyyy-MM-dd HH:mm:ss'),
    };
  } finally {
    lock.releaseLock();
  }
}

// ====== 休日出勤：作業完了（Sebastian経由） ======
// 既存 api_markHolidayDone と同等。assertSelf_ を省略。
// WorkLogsのactualStartAt（事前にstartで記録済）と end から実働分を計算。
// 休憩は calcBreakByClockTime_ で時刻ベース控除。
// PDF生成は status === 'approved' の場合のみ実行。

function sebastian_markHolidayDone_(requestId, manualEndTime) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = getRequestById_(requestId);
    if (!req) throw new Error('申請が見つかりません: ' + requestId);
    if (req.requestType !== 'holiday') throw new Error('休日出勤申請ではありません: ' + requestId);
    if (req.status === 'canceled') throw new Error('キャンセル済みの申請です: ' + requestId);

    const sh = requireSheet_('WorkLogs');
    const header = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0].map(h => normalize_(h));
    const idx = buildHeaderIndex_(header);
    const rowNo = findWorkLogRowNo_(requestId);
    if (rowNo === -1) throw new Error('休日開始が記録されていません（先にstartを実行してください）: ' + requestId);

    const row = sh.getRange(rowNo, 1, 1, sh.getLastColumn()).getValues()[0];
    const startRaw = row[idx['actualStartAt']];
    let start;
    if (startRaw instanceof Date) {
      var ssTz = getSafeTimeZone_(sh);
      var startStr = Utilities.formatDate(startRaw, ssTz, "yyyy-MM-dd'T'HH:mm:ss");
      start = new Date(startStr + '+09:00');
    } else if (startRaw) {
      start = new Date(String(startRaw).replace(' ', 'T') + '+09:00');
    }
    if (!start || isNaN(start.getTime())) {
      throw new Error('休日開始が記録されていません（先にstartを実行してください）: ' + requestId);
    }

    var end;
    if (manualEndTime && /^\d{2}:\d{2}$/.test(manualEndTime)) {
      const dEnd = req.targetDate instanceof Date ? req.targetDate : new Date(req.targetDate);
      const targetYmd = fmtDate_(dEnd, 'yyyy-MM-dd');
      end = new Date(targetYmd + 'T' + manualEndTime + ':00+09:00');
    } else {
      end = new Date();
    }

    const actualMinutes = Math.max(0, Math.round((end.getTime() - start.getTime()) / 60000));
    const breakMinutes = calcBreakByClockTime_(start, end);
    const netMinutes = Math.max(0, actualMinutes - breakMinutes);

    updateWorkLog_(requestId, {
      actualStartAt: fmtDate_(start, 'yyyy-MM-dd HH:mm:ss'),
      actualEndAt:   fmtDate_(end,   'yyyy-MM-dd HH:mm:ss'),
      actualMinutes: actualMinutes,
      breakMinutes:  breakMinutes,
      netMinutes:    netMinutes,
      updatedAt:     fmtDate_(new Date(), 'yyyy-MM-dd HH:mm:ss'),
      updatedBy:     'sebastian-agent',
    });

    let pdf = null;
    if (req.status === 'approved') {
      try {
        pdf = generatePdfForRequest_(requestId);
      } catch (e) {
        pdf = { error: String(e.message || e) };
      }
    }

    // 蓄積SSにリアルタイム書き込み（失敗しても処理続行）
    upsertAccumulationRow_(requestId);

    return {
      ok: true,
      requestId: requestId,
      workerName: req.workerName,
      targetDate: req.targetDate,
      actualStartAt: fmtDate_(start, 'yyyy-MM-dd HH:mm:ss'),
      actualEndAt:   fmtDate_(end,   'yyyy-MM-dd HH:mm:ss'),
      actualMinutes: actualMinutes,
      breakMinutes: breakMinutes,
      netMinutes: netMinutes,
      status: req.status,
      pdf: pdf,
    };
  } finally {
    lock.releaseLock();
  }
}
