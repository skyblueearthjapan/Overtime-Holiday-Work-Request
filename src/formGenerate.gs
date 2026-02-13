// ====== 部署×種別フォームを「作る or 返す」核心関数 ======

function getOrCreateDeptForm_(type, dept) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const found = findFormMapRow_(type, dept);
    if (found && found.isActive !== false && normalize_(found.row[found.idx.formId])) {
      const formId = normalize_(found.row[found.idx.formId]);
      const form = FormApp.openById(formId);
      return { formId, formUrl: form.getPublishedUrl(), created: false };
    }

    // 生成
    const templateId = getTemplateFormId_(type);
    const templateFile = DriveApp.getFileById(templateId);

    const formTitle = (type === 'overtime')
      ? `【残業申請】${dept}`
      : `【休日申請】${dept}`;

    const copyFile = templateFile.makeCopy(formTitle);
    const newFormId = copyFile.getId();
    const form = FormApp.openById(newFormId);

    // タイトル・説明
    form.setTitle(formTitle);
    form.setDescription([
      '【運用案内】',
      '・送信後の修正は原則できません。',
      '・トップ画面に申請が表示され、承認後に「承認済み」ラベルに変わります。',
      '・実績はトップ画面のボタン（残業=完了のみ／休日=開始・完了）で記録します。',
    ].join('\n'));

    // 質問の固定（部署/種別は1択にして実質編集不可）
    setDropdownChoices_(findItemByTitle_(form, Q.TYPE), [type === 'overtime' ? '残業' : '休日']);
    setDropdownChoices_(findItemByTitle_(form, Q.DEPT), [dept]);

    // 選択肢を最新マスタでセット
    updateDeptFormChoices_(form, type, dept);

    // onFormSubmit トリガーを付与
    addFormSubmitTrigger_(newFormId);

    // 保存
    upsertFormMap_(type, dept, newFormId, form.getPublishedUrl());
    return { formId: newFormId, formUrl: form.getPublishedUrl(), created: true };
  } finally {
    lock.releaseLock();
  }
}

// ====== フォームの選択肢更新（作業員・業務ID・工番・理由・予定時間） ======

function updateDeptFormChoices_(formOrFormId, type, dept) {
  const form = (typeof formOrFormId === 'string') ? FormApp.openById(formOrFormId) : formOrFormId;

  // マスタ読み込み
  const workersByDept = loadWorkersByDept_();
  const jobsByDept = loadJobsByDept_();
  const orders = loadOrderChoices_();

  const workerChoices = workersByDept.get(dept) || [];
  const jobChoices = jobsByDept.get(dept) || [];

  if (workerChoices.length === 0) {
    // 空だとフォーム送信不能になるので、最低限のダミーを入れる（またはエラーにする）
    throw new Error(`作業員候補が空です。作業員マスタを確認してください: dept=${dept}`);
  }
  if (jobChoices.length === 0) {
    throw new Error(`業務ID候補が空です。業務NOマスタを確認してください: dept=${dept}`);
  }
  if (orders.length === 0) {
    throw new Error('工番候補が空です。工番マスタを確認してください。');
  }

  // セット
  setDropdownChoices_(findItemByTitle_(form, Q.WORKER), workerChoices);
  setDropdownChoices_(findItemByTitle_(form, Q.JOB), jobChoices);
  setDropdownChoices_(findItemByTitle_(form, Q.ORDER), orders);
  setDropdownChoices_(findItemByTitle_(form, Q.REASON), REASONS);

  if (type === 'overtime') {
    setDropdownChoices_(findItemByTitle_(form, Q.OT_HOURS), OT_HOURS);
  } else {
    setDropdownChoices_(findItemByTitle_(form, Q.HD_HOURS), HD_HOURS);
  }

  // 種別/部署は1択固定（再確認）
  setDropdownChoices_(findItemByTitle_(form, Q.TYPE), [type === 'overtime' ? '残業' : '休日']);
  setDropdownChoices_(findItemByTitle_(form, Q.DEPT), [dept]);

  return true;
}
