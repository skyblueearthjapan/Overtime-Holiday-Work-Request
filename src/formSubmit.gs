// ====== フォーム送信トリガー付与 ======

function addFormSubmitTrigger_(formId) {
  const form = FormApp.openById(formId);

  // 重複作成防止（同じフォームに同じハンドラが既にあれば作らない）
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'handleFormSubmit_' && t.getTriggerSourceId && t.getTriggerSourceId() === formId) {
      return;
    }
  }

  ScriptApp.newTrigger('handleFormSubmit_')
    .forForm(form)
    .onFormSubmit()
    .create();
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

// ====== 業務ID ラベル解析 ======

function parseJobChoice_(label) {
  // 例： "J01 001:設計"
  const s = normalize_(label);
  if (!s) return { jobId: '', jobLabel: '' };
  const parts = s.split(/\s+/);
  return { jobId: parts[0], jobLabel: s };
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
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // e.response.getItemResponses() を使う（タイトルで拾うのでフォームが増えても堅い）
    const response = e.response;
    const itemResponses = response.getItemResponses();

    // タイトル→回答 のMap
    const ans = new Map();
    for (const ir of itemResponses) {
      const title = normalize_(ir.getItem().getTitle());
      const value = ir.getResponse();
      ans.set(title, value);
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

    // 実施日
    const targetDateRaw = ans.get(Q.DATE);
    const targetDate = targetDateRaw instanceof Date ? targetDateRaw : new Date(targetDateRaw);
    if (!targetDate || isNaN(targetDate.getTime())) throw new Error(`作業実施日が不正です: ${targetDateRaw}`);

    // 工番（プルダウンのラベルから工番コードだけ抽出）
    const orderLabel = normalize_(ans.get(Q.ORDER));
    const orderNo = orderLabel.split('｜')[0].trim(); // "123｜..." → "123"
    if (!orderNo) throw new Error(`工番が取得できません: ${Q.ORDER}`);

    // 業務ID（部署で絞っている想定）
    const jobLabel = normalize_(ans.get(Q.JOB));
    const { jobId, jobLabel: jobLabelFull } = parseJobChoice_(jobLabel);
    if (!jobId) throw new Error(`業務IDが取得できません: ${Q.JOB}`);

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

    // 工番マスタ補完
    const orderInfo = lookupOrderInfo_(orderNo) || { customer: '', dest: '', product: '' };

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
      workerEmail: '', // 作業員マスタにメール列がある場合、ここで補完する（後述の補助関数で対応可）
      targetDate,
      submittedAt: now,
      approvedAt: '',
      approvedBy: '',
      approvedMinutes,
      reason,
      workContent: '', // フォームに業務内容があるならここへ
      // 明細（今回は1件運用をV1とする）
      jobId1: jobId,
      workNo1: jobLabelFull,
      orderNo1: orderNo,
      customer1: orderInfo.dest || orderInfo.customer || '',
      product1: orderInfo.product || '',
      // その他
      hrMailSentAt: '',
      pdfGeneratedAt: '',
      pdfFileId: '',
      pdfFolderId: '',
      exportError: '',
    });

    // WorkLogs プレースホルダ（requestIdの行を作っておく）
    ensureWorkLogRow_(requestId);

  } catch (err) {
    // フォーム送信を止めることはできないので、ログに残す
    console.error(err);
    // 必要なら管理者にメール通知も可能（V2）
  } finally {
    lock.releaseLock();
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
    workContent: 'workContent',

    jobId1: 'jobId1',
    workNo1: 'workNo1',
    orderNo1: 'orderNo1',
    customer1: 'customer1',
    product1: 'product1',

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
