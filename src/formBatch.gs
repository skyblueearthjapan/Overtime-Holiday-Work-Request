// ====== 全フォーム更新（毎朝バッチ） ======

function nightlyUpdateAllForms_() {
  const sh = ensureFormMapSheet_();
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idx = {
    type: H.indexOf('type'),
    dept: H.indexOf('dept'),
    formId: H.indexOf('formId'),
    isActive: H.indexOf('isActive'),
  };
  if (idx.type < 0 || idx.dept < 0 || idx.formId < 0) {
    throw new Error('FormMapヘッダが想定と違います。');
  }

  let ok = 0, ng = 0;
  const errors = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const isActive = idx.isActive >= 0 ? row[idx.isActive] : true;
    if (isActive === false || String(isActive).toLowerCase() === 'false') continue;

    const type = normalize_(row[idx.type]);
    const dept = normalize_(row[idx.dept]);
    const formId = normalize_(row[idx.formId]);
    if (!type || !dept || !formId) continue;

    try {
      updateDeptFormChoices_(formId, type, dept);
      ok++;
    } catch (e) {
      ng++;
      errors.push(`type=${type} dept=${dept} : ${e.message}`);
    }
  }

  Logger.log(`nightlyUpdateAllForms: ok=${ok} ng=${ng}`);
  if (errors.length) Logger.log(errors.join('\n'));
}
