// ====== FormMap の自動作成/取得 ======

function ensureFormMapSheet_() {
  const ss = getDb_();
  let sh = ss.getSheetByName(SHEET.FORM_MAP);
  if (sh) return sh;

  // FormMap が無ければ自動作成（ヘッダ付き）
  sh = ss.insertSheet(SHEET.FORM_MAP);
  sh.appendRow(['type', 'dept', 'formId', 'formUrl', 'updatedAt', 'isActive']);
  return sh;
}

// ====== FormMap の検索 ======

function findFormMapRow_(type, dept) {
  const sh = ensureFormMapSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null; // ヘッダのみ → 該当なし

  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idx = {
    type: H.indexOf('type'),
    dept: H.indexOf('dept'),
    formId: H.indexOf('formId'),
    formUrl: H.indexOf('formUrl'),
    updatedAt: H.indexOf('updatedAt'),
    isActive: H.indexOf('isActive'),
  };
  if (idx.type < 0 || idx.dept < 0 || idx.formId < 0) {
    throw new Error('FormMapのヘッダが想定と違います（type/dept/formId/formUrl...）');
  }

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (normalize_(row[idx.type]) === type && normalize_(row[idx.dept]) === dept) {
      const isActive = idx.isActive >= 0 ? row[idx.isActive] : true;
      return { rowIndex: r + 1, idx, row, isActive };
    }
  }
  return null;
}

// ====== FormMap の書き込み ======

function upsertFormMap_(type, dept, formId, formUrl) {
  const sh = ensureFormMapSheet_();
  const found = findFormMapRow_(type, dept);
  const now = new Date();
  if (found) {
    const r = found.rowIndex;
    sh.getRange(r, found.idx.formId + 1).setValue(formId);
    if (found.idx.formUrl >= 0) sh.getRange(r, found.idx.formUrl + 1).setValue(formUrl);
    if (found.idx.updatedAt >= 0) sh.getRange(r, found.idx.updatedAt + 1).setValue(now);
    if (found.idx.isActive >= 0) sh.getRange(r, found.idx.isActive + 1).setValue(true);
  } else {
    sh.appendRow([type, dept, formId, formUrl, now, true]);
  }
}

// ====== テンプレフォームID取得 ======
// Settings → FormTemplates シートの順で探す

function getTemplateFormId_(type) {
  // 1) Settings から取得（OVERTIME_TEMPLATE_FORM_ID / HOLIDAY_TEMPLATE_FORM_ID）
  const settingsKey = (type === 'overtime')
    ? 'OVERTIME_TEMPLATE_FORM_ID'
    : 'HOLIDAY_TEMPLATE_FORM_ID';
  try {
    const settings = getSettings_();
    const id = normalize_(settings[settingsKey]);
    if (id) return id;
  } catch (_) { /* Settings取得失敗は無視して次へ */ }

  // 2) FormTemplates シートから取得（従来方式）
  const ss = getDb_();
  const sh = ss.getSheetByName(SHEET.FORM_TEMPLATES);
  if (!sh) {
    throw new Error(
      'テンプレフォームIDが見つかりません。\n' +
      `Settings に ${settingsKey} を設定するか、FormTemplates シートを用意してください。`
    );
  }
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idxType = H.indexOf('type');
  const idxId = H.indexOf('templateFormId');
  if (idxType < 0 || idxId < 0) throw new Error('FormTemplatesのヘッダが想定と違います（type/templateFormId）。');

  for (let r = 1; r < values.length; r++) {
    if (normalize_(values[r][idxType]) === type) {
      const id = normalize_(values[r][idxId]);
      if (!id) throw new Error(`FormTemplatesにtemplateFormIdがありません: type=${type}`);
      return id;
    }
  }
  throw new Error(`FormTemplatesにtypeがありません: ${type}`);
}

// ====== フォームURL取得（TOP画面用：未作成なら自動生成） ======

function api_getFormUrl(type, dept, workerLabel) {
  let formId = '';

  // 既存を探す
  const found = findFormMapRow_(type, dept);
  if (found) {
    const isActive = found.isActive;
    if (isActive !== false && String(isActive).toLowerCase() !== 'false') {
      formId = normalize_(found.row[found.idx.formId]);
    }
  }

  // 未作成 → 自動生成を試みる
  if (!formId) {
    const result = getOrCreateDeptForm_(type, dept);
    if (!result) return null;
    formId = result.formId;
  }

  if (!formId) return null;

  // プリフィルURLを生成（申請種別・部署・作業員を事前入力）
  return buildPrefillUrl_(formId, type, dept, workerLabel);
}

/**
 * プリフィル（事前入力）付きURLを生成する。
 * 申請種別・部署・作業員を自動セットした状態でフォームを開く。
 */
function buildPrefillUrl_(formId, requestType, dept, workerLabel) {
  const form = FormApp.openById(formId);
  let formResponse = form.createResponse();

  // 申請種別をプリフィル（"残業" or "休日"）
  const typeLabel = (requestType === 'overtime') ? '残業' : '休日';
  formResponse = addPrefill_(formResponse, form, Q.TYPE, typeLabel);

  // 部署をプリフィル
  formResponse = addPrefill_(formResponse, form, Q.DEPT, dept);

  // 作業員をプリフィル
  if (workerLabel) {
    formResponse = addPrefill_(formResponse, form, Q.WORKER, workerLabel);
  }

  // 作業実施日を今日でプリフィル
  formResponse = addDatePrefill_(formResponse, form, Q.DATE);

  return formResponse.toPrefilledUrl();
}

/**
 * フォーム回答オブジェクトにプリフィルを追加するヘルパー。
 * 質問が見つからない場合やタイプ不一致の場合はスキップ。
 */
function addPrefill_(formResponse, form, questionTitle, value) {
  const item = findItemByTitleOrNull_(form, questionTitle);
  if (!item || !value) return formResponse;

  const type = item.getType();
  let itemResponse;
  if (type === FormApp.ItemType.LIST) {
    itemResponse = item.asListItem().createResponse(value);
  } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    itemResponse = item.asMultipleChoiceItem().createResponse(value);
  } else if (type === FormApp.ItemType.TEXT) {
    itemResponse = item.asTextItem().createResponse(value);
  } else {
    return formResponse; // 未対応タイプはスキップ
  }

  return formResponse.withItemResponse(itemResponse);
}

/**
 * 日付フィールドに今日の日付をプリフィルするヘルパー。
 */
function addDatePrefill_(formResponse, form, questionTitle) {
  const item = findItemByTitleOrNull_(form, questionTitle);
  if (!item) return formResponse;

  if (item.getType() !== FormApp.ItemType.DATE) return formResponse;

  const today = new Date();
  const year = Number(Utilities.formatDate(today, TZ, 'yyyy'));
  const month = Number(Utilities.formatDate(today, TZ, 'M'));
  const day = Number(Utilities.formatDate(today, TZ, 'd'));

  const itemResponse = item.asDateItem().createResponse(year, month, day);
  return formResponse.withItemResponse(itemResponse);
}
