// ====== 手動メニュー（管理者が押せる） ======

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('残業休日アプリ')
    .addItem('フォームを全更新（マスタ反映）', 'nightlyUpdateAllForms_')
    .addItem('フォーム生成テスト（残業×先頭部署）', 'debugCreateFirstDeptOvertime_')
    .addSeparator()
    .addItem('夕方メール送信（手動）', 'sendEveningMail_')
    .addItem('朝メール送信（手動）', 'sendMorningMail_')
    .addSeparator()
    .addItem('マスタ転記（手動実行）', 'syncAllMasters')
    .addSeparator()
    .addItem('全トリガー初期セットアップ', 'setupAllTriggers_')
    .addToUi();
}

function debugCreateFirstDeptOvertime_() {
  const depts = loadDeptList_();
  if (!depts.length) throw new Error('部署マスタが空です。');
  const dept = depts[0];
  const res = getOrCreateDeptForm_('overtime', dept);
  Logger.log(JSON.stringify(res, null, 2));
}

// ====== トリガー作成（フォーム毎朝更新） ======

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

// ====== トリガー作成（メール：夕方2回＋朝1回） ======

function setupMailTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  const has = (fn) => triggers.some(t => t.getHandlerFunction() === fn);

  // 夕方 17:10（17時台）
  if (!triggers.some(t => t.getHandlerFunction()==='sendEveningMail_' && t.getEventType()===ScriptApp.EventType.CLOCK)) {
    ScriptApp.newTrigger('sendEveningMail_')
      .timeBased()
      .everyDays(1)
      .atHour(17)
      .nearMinute(10)
      .create();
  }

  // 夕方 18:10（18時台）…同じ関数をもう1本
  ScriptApp.newTrigger('sendEveningMail_')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .nearMinute(10)
    .create();

  // 朝 07:10（7時台）
  if (!has('sendMorningMail_')) {
    ScriptApp.newTrigger('sendMorningMail_')
      .timeBased()
      .everyDays(1)
      .atHour(7)
      .nearMinute(10)
      .create();
  }
}

// ====== トリガー作成（マスタ自動転記） ======

function setupSyncTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.some(t => t.getHandlerFunction() === 'syncAllMasters')) return;

  // 毎日 06:00 にマスタ同期（フォーム更新 06:30 より前に実行）
  ScriptApp.newTrigger('syncAllMasters')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(0)
    .create();
}

// ====== 全トリガー一括セットアップ ======

function setupAllTriggers_() {
  setupSyncTrigger_();    // マスタ転記（6:00）
  setupTriggers_();       // フォーム毎朝更新（6:30）
  setupMailTriggers_();   // 夕方2回（17:10, 18:10）＋朝（7:10）
  Logger.log('全トリガーをセットアップしました。');
}
