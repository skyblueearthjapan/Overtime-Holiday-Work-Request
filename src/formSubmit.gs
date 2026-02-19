// ====== フォーム送信トリガー付与 ======
// ※ ポーリング方式（pollNewResponses_）に移行済みのため、
//    個別フォームに onFormSubmit トリガーは作成しない。
//    GAS のトリガー上限（20個/プロジェクト）を回避するための設計変更。

function addFormSubmitTrigger_(formId) {
  // No-op: pollNewResponses_（1分間隔の時間トリガー）で全フォームを一括処理
}

// ====== フォーム回答ポーリング（全フォーム一括チェック） ======
// 個別 onFormSubmit トリガーの代わりに、時間駆動で全フォームの新規回答を処理する。
// setupAllTriggers_ で1分間隔の時間トリガーとして登録される。

function pollNewResponses_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('pollNewResponses_: lock取得失敗（別プロセスが実行中）');
    return;
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const lastTs = props.getProperty('POLL_LAST_TS');
    let since = lastTs ? new Date(lastTs) : null;
    const now = new Date();

    // 初回実行時：24時間前を基準にして既存回答も拾う
    if (!since) {
      since = new Date(now.getTime() - 24 * 60 * 60 * 1000);
      props.setProperty('POLL_LAST_TS', since.toISOString());
      Logger.log('pollNewResponses_: 初回実行 — 24時間前を基準タイムスタンプにセット');
      // ↓ そのまま処理を続行（return しない）
    }

    // FormMap から全アクティブフォームを取得
    const sh = ensureFormMapSheet_();
    const values = sh.getDataRange().getValues();
    const H = values[0].map(h => normalize_(h));
    const formIdCol = H.indexOf('formId');
    const activeCol = H.indexOf('isActive');
    if (formIdCol < 0) return;

    let processed = 0;
    let errors = 0;
    let latestResponseTs = since; // 処理成功した最新の回答タイムスタンプ

    // リトライ時の重複処理防止用：キャッシュから処理済み回答IDを読み込む
    const cache = CacheService.getScriptCache();
    const cachedIds = cache.get('POLL_PROCESSED_IDS');
    const processedIds = new Set(cachedIds ? JSON.parse(cachedIds) : []);

    for (let r = 1; r < values.length; r++) {
      const formId = normalize_(values[r][formIdCol]);
      if (!formId) continue;
      if (activeCol >= 0) {
        const a = values[r][activeCol];
        if (a === false || String(a).toLowerCase() === 'false') continue;
      }

      try {
        const form = FormApp.openById(formId);
        // since 以降の新規回答だけ取得（GAS組込のフィルタ）
        const responses = form.getResponses(since);

        for (const resp of responses) {
          // 重複処理防止：同じ回答を2回処理しない
          const respId = resp.getId();
          if (processedIds.has(respId)) continue;

          try {
            handleFormSubmit_({ response: resp });
            processedIds.add(respId);
            processed++;
            // 成功した回答のタイムスタンプを追跡
            const respTs = resp.getTimestamp();
            if (respTs > latestResponseTs) latestResponseTs = respTs;
          } catch (e) {
            errors++;
            console.error('pollNewResponses_ handler error (form=' + formId + '): ' + e.message);
            // エラー時はタイムスタンプを進めない → 次回リトライ
          }
        }
      } catch (e) {
        console.warn('pollNewResponses_ form open error (' + formId + '): ' + e.message);
      }
    }

    // タイムスタンプ更新：エラーがなければ now まで進める
    // エラーがあれば、成功した最新の回答まで（失敗分は次回リトライ）
    if (errors === 0) {
      props.setProperty('POLL_LAST_TS', now.toISOString());
    } else if (latestResponseTs > since) {
      // 部分的に成功 → 最新成功分まで進める（ただし失敗分の前まで）
      // 安全のため since のまま据え置き（全件リトライ）
      Logger.log('pollNewResponses_: errors発生のためタイムスタンプ据え置き（次回リトライ）');
    } else {
      // 全件失敗 → タイムスタンプ据え置き
      Logger.log('pollNewResponses_: 全件エラーのためタイムスタンプ据え置き');
    }

    // 処理済みIDキャッシュを更新（10分間保持、古いIDは自然に失効）
    if (processedIds.size > 0) {
      // 最新100件だけ保持（メモリ節約）
      const idArr = Array.from(processedIds).slice(-100);
      cache.put('POLL_PROCESSED_IDS', JSON.stringify(idArr), 600);
    }

    if (processed > 0 || errors > 0) {
      Logger.log('pollNewResponses_: processed=' + processed + ' errors=' + errors);
    }
  } finally {
    lock.releaseLock();
  }
}

// ====== デバッグ：ポーリング状態確認（手動実行用） ======
// GASスクリプトエディタから「debugPollStatus_」を実行すると、
// ポーリング状態と各フォームの未処理回答数がログに出力されます。

function debugPollStatus_() {
  const props = PropertiesService.getScriptProperties();
  const lastTs = props.getProperty('POLL_LAST_TS');
  Logger.log('=== ポーリング診断 ===');
  Logger.log('POLL_LAST_TS: ' + (lastTs || '(未設定 — ポーリング未実行)'));

  if (lastTs) {
    const since = new Date(lastTs);
    const ageMin = Math.round((Date.now() - since.getTime()) / 60000);
    Logger.log('最終チェック: ' + ageMin + '分前');
  }

  // トリガー確認
  const triggers = ScriptApp.getProjectTriggers();
  const pollTrigger = triggers.find(t => t.getHandlerFunction() === 'pollNewResponses_');
  Logger.log('pollNewResponses_ トリガー: ' + (pollTrigger ? '存在する (' + pollTrigger.getEventType() + ')' : '存在しない！ setupAllTriggers_ を実行してください'));

  // FormMap の全フォームを確認
  const sh = ensureFormMapSheet_();
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const formIdCol = H.indexOf('formId');
  const typeCol = H.indexOf('type');
  const deptCol = H.indexOf('dept');
  const activeCol = H.indexOf('isActive');
  if (formIdCol < 0) {
    Logger.log('FormMapに formId 列がありません');
    return;
  }

  const since = lastTs ? new Date(lastTs) : new Date(Date.now() - 24 * 60 * 60 * 1000);

  for (let r = 1; r < values.length; r++) {
    const formId = normalize_(values[r][formIdCol]);
    if (!formId) continue;
    if (activeCol >= 0) {
      const a = values[r][activeCol];
      if (a === false || String(a).toLowerCase() === 'false') {
        Logger.log('  [SKIP] ' + values[r][typeCol] + '/' + values[r][deptCol] + ' (inactive)');
        continue;
      }
    }

    try {
      const form = FormApp.openById(formId);
      const allResp = form.getResponses();
      const newResp = form.getResponses(since);
      Logger.log('  ' + values[r][typeCol] + '/' + values[r][deptCol] +
        ' — 全回答: ' + allResp.length + ', since以降: ' + newResp.length);

      // 最新回答の詳細
      if (newResp.length > 0) {
        const latest = newResp[newResp.length - 1];
        Logger.log('    最新回答: ' + latest.getTimestamp() + ' id=' + latest.getId());

        // 各アイテムの回答を試行
        try {
          const irs = latest.getItemResponses();
          const titles = irs.map(ir => {
            try { return ir.getItem().getTitle(); }
            catch (_) { return '(削除済みアイテム)'; }
          });
          Logger.log('    回答フィールド: ' + titles.join(', '));
        } catch (e) {
          Logger.log('    getItemResponses() エラー: ' + e.message);
        }
      }
    } catch (e) {
      Logger.log('  [ERROR] ' + values[r][typeCol] + '/' + values[r][deptCol] + ': ' + e.message);
    }
  }

  Logger.log('=== 診断完了 ===');
}

// ====== デバッグ：手動で1件処理テスト ======
// 最新の未処理回答を1件だけ処理してエラーを確認する。

function debugProcessLatest_() {
  const sh = ensureFormMapSheet_();
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const formIdCol = H.indexOf('formId');
  const activeCol = H.indexOf('isActive');
  if (formIdCol < 0) return;

  const props = PropertiesService.getScriptProperties();
  const lastTs = props.getProperty('POLL_LAST_TS');
  const since = lastTs ? new Date(lastTs) : new Date(Date.now() - 24 * 60 * 60 * 1000);

  for (let r = 1; r < values.length; r++) {
    const formId = normalize_(values[r][formIdCol]);
    if (!formId) continue;
    if (activeCol >= 0) {
      const a = values[r][activeCol];
      if (a === false || String(a).toLowerCase() === 'false') continue;
    }

    try {
      const form = FormApp.openById(formId);
      const newResp = form.getResponses(since);
      if (newResp.length === 0) continue;

      const resp = newResp[newResp.length - 1];
      Logger.log('処理テスト: formId=' + formId + ' respId=' + resp.getId());

      handleFormSubmit_({ response: resp });
      Logger.log('処理成功！Requestsシートを確認してください。');

      // 成功したらタイムスタンプを更新
      props.setProperty('POLL_LAST_TS', new Date().toISOString());
      return;
    } catch (e) {
      Logger.log('処理エラー: ' + e.message + '\n' + e.stack);
    }
  }

  Logger.log('未処理の回答が見つかりませんでした。');
}

// ====== Requests/WorkLogs 共通ヘルパー ======

function buildHeaderIndex_(headerRow) {
  const idx = {};
  headerRow.forEach((h, i) => {
    const key = normalize_(h);
    if (key) idx[key] = i;
  });
  return idx;
}

function getRequestsSheet_() {
  return requireSheet_('Requests');
}

function getWorkLogsSheet_() {
  return requireSheet_('WorkLogs');
}

// ====== 工番マスタ補完 ======

function lookupOrderInfo_(orderNo) {
  const sh = requireSheet_(SHEET.ORDERS);
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idx = {
    order: H.indexOf('工番'),
    customer: H.indexOf('受注先'),
    dest: H.indexOf('納入先'),
    product: H.indexOf('品名'),
  };
  if (idx.order < 0) return null;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (normalize_(row[idx.order]) === orderNo) {
      return {
        orderNo,
        customer: idx.customer >= 0 ? normalize_(row[idx.customer]) : '',
        dest: idx.dest >= 0 ? normalize_(row[idx.dest]) : '',
        product: idx.product >= 0 ? normalize_(row[idx.product]) : '',
      };
    }
  }
  return null;
}

// ====== フォーム回答から工番1〜3をパース ======
// モーダルで選択された工番コードがそのままテキスト欄に入る。
// 全て任意（間接業務の方は工番なしで申請可能）。

function parseWorkNosFromForm_(ans) {
  const fields = [Q.ORDER_1, Q.ORDER_2, Q.ORDER_3];
  const workNos = fields.map(q => normalize_(ans.get(q) || ''));

  return {
    workNo1: workNos[0],
    workNo2: workNos[1],
    workNo3: workNos[2],
    errors: [],
  };
}

// ====== 工番1〜3をマスタ突合して補完結果を返す ======

function enrichWorkNos_(workNo1, workNo2, workNo3) {
  const result = {
    orderNo1: workNo1, customer1: '', product1: '',
    orderNo2: workNo2, customer2: '', product2: '',
    orderNo3: workNo3, customer3: '', product3: '',
    errors: [],
  };

  const lookup = (workNo, suffix) => {
    if (!workNo) return;
    const info = lookupOrderInfo_(workNo);
    if (info) {
      result['customer' + suffix] = info.dest || info.customer || '';
      result['product' + suffix] = info.product || '';
    } else {
      result.errors.push('工番マスタ未登録: ' + workNo);
    }
  };

  lookup(workNo1, '1');
  lookup(workNo2, '2');
  lookup(workNo3, '3');

  return result;
}

// ====== 作業員ラベル解析 ======

function parseWorkerChoice_(label) {
  const s = normalize_(label);
  if (!s) return { workerCode: '', workerName: '' };
  const parts = s.split(/\s+/);
  const code = parts[0] || '';
  const name = parts.slice(1).join(' ') || '';
  return { workerCode: code, workerName: name };
}

// ====== 予定時間（分）への正規化 ======

function plannedMinutesFromOvertime_(hStr) {
  // "0.5" → 30
  const n = Number(normalize_(hStr));
  if (!isFinite(n) || n <= 0) throw new Error(`残業の予定時間が不正です: ${hStr}`);
  return Math.round(n * 60);
}

function plannedMinutesFromHoliday_(label) {
  const s = normalize_(label);
  if (s === '半日') return 240;
  if (s === '1日') return 480;
  throw new Error(`休日の予定時間が不正です: ${label}`);
}

// ====== フォーム送信ハンドラ本体（installable onFormSubmit） ======
// 送信した瞬間に Requests に 1 行追加 → トップで即表示

function handleFormSubmit_(e) {
  // ※ ロックは呼び出し元（pollNewResponses_）で取得済みのため、ここでは取得しない。
  //    再取得するとデッドロックになる。
  try {
    // e.response.getItemResponses() を使う（タイトルで拾うのでフォームが増えても堅い）
    const response = e.response;
    let itemResponses;
    try {
      itemResponses = response.getItemResponses();
    } catch (irErr) {
      console.error('getItemResponses() 失敗（フォーム構造変更の影響の可能性）: ' + irErr.message);
      throw irErr;
    }

    // タイトル→回答 のMap
    // 注: ensureOrderItems_ でフォームアイテムを削除した場合、
    //     旧回答の ir.getItem() がエラーになることがあるため try/catch で保護
    const ans = new Map();
    for (const ir of itemResponses) {
      try {
        const title = normalize_(ir.getItem().getTitle());
        const value = ir.getResponse();
        ans.set(title, value);
      } catch (itemErr) {
        // 削除済みアイテムの回答 → スキップ（工番旧形式の残骸など）
        console.warn('回答アイテム読取スキップ（削除済みの可能性）: ' + itemErr.message);
      }
    }

    // 種別・部署
    const typeJa = normalize_(ans.get(Q.TYPE));
    // "残業" or "休日"
    const dept = normalize_(ans.get(Q.DEPT));
    if (!typeJa || !dept) throw new Error(`申請種別/部署が取得できません。テンプレ質問タイトルを確認: ${Q.TYPE}/${Q.DEPT}`);

    const requestType = (typeJa === '残業') ? 'overtime' : 'holiday';

    // 作業員
    const workerLabel = normalize_(ans.get(Q.WORKER));
    const { workerCode, workerName } = parseWorkerChoice_(workerLabel);
    if (!workerCode) throw new Error(`作業員が取得できません: ${Q.WORKER}`);

    // 作業員マスタからメールアドレスを補完
    const workerEmail = lookupWorkerEmail_(workerCode);

    // 実施日
    const targetDateRaw = ans.get(Q.DATE);
    const targetDate = targetDateRaw instanceof Date ? targetDateRaw : new Date(targetDateRaw);
    if (!targetDate || isNaN(targetDate.getTime())) throw new Error(`作業実施日が不正です: ${targetDateRaw}`);

    // 工番（プレフィックス選択＋5桁番号入力 ×最大3件、間接業務は全て空でもOK）
    const { workNo1, workNo2, workNo3, errors: workNoErrors } = parseWorkNosFromForm_(ans);
    // 片方だけ入力されている等の整合性エラーのみチェック（全空はOK）
    if (workNoErrors.length > 0) {
      console.warn('工番入力の注意: ' + workNoErrors.join('; '));
    }

    // 理由＋補足
    const reason = normalize_(ans.get(Q.REASON));
    const reasonDetail = normalize_(ans.get(Q.REASON_DETAIL));
    if (!reason) throw new Error(`理由が取得できません: ${Q.REASON}`);
    if (reason.startsWith('その他') && !reasonDetail) {
      // フォーム側でも必須制御するが、保険でサーバ側でもチェック
      throw new Error('理由が「その他」の場合、補足理由が必須です。');
    }

    // 予定時間（分）
    let approvedMinutes = 0;
    if (requestType === 'overtime') {
      approvedMinutes = plannedMinutesFromOvertime_(ans.get(Q.OT_HOURS));
    } else {
      approvedMinutes = plannedMinutesFromHoliday_(ans.get(Q.HD_HOURS));
    }

    // 工番マスタ突合（3件分を一括補完）
    const enriched = enrichWorkNos_(workNo1, workNo2, workNo3);

    // エラー集約（工番パース＋マスタ未登録）
    const allErrors = [...workNoErrors, ...enriched.errors];

    // Requests に追加
    const requestId = Utilities.getUuid();
    const now = new Date();

    appendRequestRow_({
      requestId,
      requestType,
      status: 'submitted',
      dept,
      workerCode,
      workerName,
      workerEmail,
      targetDate,
      submittedAt: now,
      approvedAt: '',
      approvedBy: '',
      approvedMinutes,
      reason,
      reasonDetail,
      workContent: '', // フォームに業務内容があるならここへ
      // 明細（最大3件）
      jobId1: '',
      workNo1: workNo1,
      orderNo1: enriched.orderNo1,
      customer1: enriched.customer1,
      product1: enriched.product1,
      workNo2: workNo2,
      orderNo2: enriched.orderNo2,
      customer2: enriched.customer2,
      product2: enriched.product2,
      workNo3: workNo3,
      orderNo3: enriched.orderNo3,
      customer3: enriched.customer3,
      product3: enriched.product3,
      // その他
      hrMailSentAt: '',
      pdfGeneratedAt: '',
      pdfFileId: '',
      pdfFolderId: '',
      exportError: allErrors.length > 0 ? allErrors.join(' / ') : '',
    });

    // WorkLogs プレースホルダ（requestIdの行を作っておく）
    // ※ WorkLogs書き込みは補助的なので、失敗しても Requests 書き込み済みなら
    //    処理成功とみなす（次回ポーリングで無限ループしないため）
    try {
      ensureWorkLogRow_(requestId);
    } catch (wlErr) {
      console.warn('ensureWorkLogRow_ 警告（Requests書込みは成功済み）: ' + wlErr.message);
    }

  } catch (err) {
    // フォーム送信を止めることはできないので、ログに残す
    console.error('handleFormSubmit_ error: ' + err.message + '\n' + err.stack);
    throw err; // pollNewResponses_ の errors カウンタに反映させるため再throw
  }
}

// ====== Requests への追記（ヘッダ名ベースで安全に書く） ======
// Requests の列順が変わっても壊れないよう「ヘッダ名」で位置を探す

function appendRequestRow_(obj) {
  const sh = getRequestsSheet_();
  const values = sh.getDataRange().getValues();
  const header = values[0].map(h => normalize_(h));
  const idx = buildHeaderIndex_(header);

  // ヘッダ名（この文字列はRequestsの1行目に存在する必要があります）
  const keyMap = {
    requestId: 'requestId',
    requestType: 'requestType(overtime/holiday)',
    status: 'status(submitted/approved/canceled)',
    dept: 'dept',
    workerCode: 'workerCode',
    workerName: 'workerName',
    workerEmail: 'workerEmail',
    targetDate: 'targetDate',
    submittedAt: 'submittedAt',
    approvedAt: 'approvedAt',
    approvedBy: 'approvedBy',
    approvedMinutes: 'approvedMinutes',
    reason: 'reason',
    reasonDetail: 'reasonDetail',
    workContent: 'workContent',

    jobId1: 'jobId1',
    workNo1: 'workNo1',
    orderNo1: 'orderNo1',
    customer1: 'customer1',
    product1: 'product1',

    workNo2: 'workNo2',
    orderNo2: 'orderNo2',
    customer2: 'customer2',
    product2: 'product2',

    workNo3: 'workNo3',
    orderNo3: 'orderNo3',
    customer3: 'customer3',
    product3: 'product3',

    hrMailSentAt: 'hrMailSentAt',
    pdfGeneratedAt: 'pdfGeneratedAt',
    pdfFileId: 'pdfFileId',
    pdfFolderId: 'pdfFolderId',
    exportError: 'exportError',
  };

  // 1行分の配列をヘッダ長で作り、該当キーだけ埋める
  const row = new Array(header.length).fill('');

  for (const [prop, headerName] of Object.entries(keyMap)) {
    const col = idx[headerName];
    if (col === undefined) continue; // 存在しない列はスキップ（V1の柔軟性）
    row[col] = obj[prop] ?? '';
  }

  sh.appendRow(row);
}

// ====== WorkLogs のプレースホルダ行を作る（無ければ追加） ======

function ensureWorkLogRow_(requestId) {
  const sh = getWorkLogsSheet_();
  const values = sh.getDataRange().getValues();

  // 2行目ヘッダ想定（1行目が注釈）
  const header = values[1].map(h => normalize_(h));
  const idx = buildHeaderIndex_(header);

  const ridCol = idx['requestId'];
  if (ridCol === undefined) throw new Error('WorkLogsに requestId 列がありません。');

  for (let r = 2; r < values.length; r++) {
    if (normalize_(values[r][ridCol]) === requestId) return; // 既にある
  }

  const row = new Array(header.length).fill('');
  row[ridCol] = requestId;
  if (idx['updatedAt'] !== undefined) row[idx['updatedAt']] = new Date();
  if (idx['updatedBy'] !== undefined) row[idx['updatedBy']] = 'formSubmit';

  sh.appendRow(row);
}
