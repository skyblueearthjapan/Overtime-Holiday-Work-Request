// ====== FormMap の検索/書き込み ======

function findFormMapRow_(type, dept) {
  const sh = requireSheet_(SHEET.FORM_MAP);
  const values = sh.getDataRange().getValues();
  // ヘッダ想定：type, dept, formId, formUrl, updatedAt, isActive
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

function upsertFormMap_(type, dept, formId, formUrl) {
  const sh = requireSheet_(SHEET.FORM_MAP);
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

function getTemplateFormId_(type) {
  const sh = requireSheet_(SHEET.FORM_TEMPLATES);
  const values = sh.getDataRange().getValues();
  // ヘッダ想定：type, templateFormId, note
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
