// ====== 手動メニュー（管理者が押せる） ======

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('残業休日アプリ')
    .addItem('フォームを全更新（マスタ反映）', 'nightlyUpdateAllForms_')
    .addItem('フォーム生成テスト（残業×先頭部署）', 'debugCreateFirstDeptOvertime_')
    .addToUi();
}

function debugCreateFirstDeptOvertime_() {
  const depts = loadDeptList_();
  if (!depts.length) throw new Error('部署マスタが空です。');
  const dept = depts[0];
  const res = getOrCreateDeptForm_('overtime', dept);
  Logger.log(JSON.stringify(res, null, 2));
}

// ====== トリガー作成（毎朝更新） ======

function setupTriggers_() {
  // 既存同名トリガーを重複作成しない簡易対策
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'nightlyUpdateAllForms_') return;
  }
  ScriptApp.newTrigger('nightlyUpdateAllForms_')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    // 6時台
    .nearMinute(30)
    .create();
}
