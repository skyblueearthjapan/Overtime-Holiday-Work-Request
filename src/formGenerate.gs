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

    // 生成前にマスタを確認（作業員0人の部署はスキップ）
    const workersByDept = loadWorkersByDept_();
    if (!workersByDept.has(dept) || workersByDept.get(dept).length === 0) {
      Logger.log(`SKIP: 作業員が0人のためフォーム未生成: type=${type} dept=${dept}`);
      return null;
    }

    const templateId = getTemplateFormId_(type);
    const templateFile = DriveApp.getFileById(templateId);

    const formTitle = (type === 'overtime')
      ? `【残業申請】${dept}`
      : `【休日申請】${dept}`;

    // フォルダ指定があればそこへ複製
    let copyFile;
    try {
      const settings = getSettings_();
      const folderId = normalize_(settings['FORMS_PARENT_FOLDER_ID']);
      if (folderId) {
        copyFile = templateFile.makeCopy(formTitle, DriveApp.getFolderById(folderId));
      } else {
        copyFile = templateFile.makeCopy(formTitle);
      }
    } catch (_) {
      copyFile = templateFile.makeCopy(formTitle);
    }
    const newFormId = copyFile.getId();

    // 組織内のリンク共有を有効化（Workspace環境向け）
    // ANYONE_WITH_LINK が管理ポリシーで制限されている場合に備え、
    // 全体公開 → ドメイン内共有の順でフォールバック
    try {
      copyFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (_) {
      copyFile.setSharing(DriveApp.Access.ANYONE_WITHIN_DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    }

    const form = FormApp.openById(newFormId);

    // 新しいGoogle Formsでは複製時に「未公開(draft)」状態になるため、
    // 先に公開してからでないと各種操作がエラーになる
    try { form.setPublished(true); } catch (_) { /* 旧環境では不要 */ }

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

    // 全設定完了後に回答受付を開始（途中で呼ぶとリセットされる場合がある）
    form.setAcceptingResponses(true);

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
    Logger.log(`WARN: 作業員候補が空のためスキップ: dept=${dept}`);
    return false;
  }
  if (jobChoices.length === 0) {
    Logger.log(`WARN: 業務ID候補が空のためスキップ: dept=${dept}`);
    return false;
  }
  if (orders.length === 0) {
    Logger.log('WARN: 工番候補が空のためスキップ');
    return false;
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
