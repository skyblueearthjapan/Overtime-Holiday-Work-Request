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
