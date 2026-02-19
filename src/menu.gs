// ====== 手動メニュー（管理者が押せる） ======

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('残業休日アプリ')
    .addItem('フォームを全更新（マスタ反映）', 'nightlyUpdateAllForms_')
    .addItem('フォーム生成テスト（残業×先頭部署）', 'debugCreateFirstDeptOvertime_')
    .addItem('全部署フォーム一括生成', 'buildAllDeptForms')
    .addSeparator()
    .addItem('夕方メール送信（手動）', 'sendEveningMail_')
    .addItem('朝バッチ（PDF生成+朝メール）', 'morningBatch_')
    .addItem('朝メールのみ（手動）', 'sendMorningMail_')
    .addItem('PDF一括生成（本日分）', 'manualBatchPdfs_')
    .addSeparator()
    .addItem('マスタ転記（手動実行）', 'syncAllMasters')
    .addSeparator()
    .addItem('全トリガー初期セットアップ', 'setupAllTriggers_')
    .addItem('旧フォームトリガー削除', 'cleanupFormSubmitTriggers_')
    .addSeparator()
    .addItem('診断：ポーリング状態確認', 'debugPollStatus_')
    .addItem('診断：最新回答を手動処理', 'debugProcessLatest_')
    .addToUi();
}

/** PDF一括生成（手動）— 本日分の承認済み＆PDF未生成を処理 */
function manualBatchPdfs_() {
  const result = batchGeneratePdfs_(new Date());
  logBatchResult_('manualBatchPdfs', new Date(), result);
  const msg = 'PDF一括生成 完了\n'
    + '成功: ' + result.ok + ' 件\n'
    + 'スキップ（作成済）: ' + result.skip + ' 件\n'
    + '失敗: ' + result.fail + ' 件'
    + (result.errors.length ? '\n\nエラー:\n' + result.errors.join('\n') : '');
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

function debugCreateFirstDeptOvertime_() {
  const depts = loadDeptList_();
  if (!depts.length) throw new Error('作業員マスタに部署データがありません。');
  const dept = depts[0];
  const res = getOrCreateDeptForm_('overtime', dept);
  Logger.log(JSON.stringify(res, null, 2));
}

/**
 * 全部署×残業/休日 フォームを一括生成（初期展開用）。
 * FormMap に既にある部署はスキップされる。
 */
function buildAllDeptForms() {
  const depts = loadDeptList_();
  if (!depts.length) throw new Error('作業員マスタに部署データがありません。');

  const log = [];
  for (const dept of depts) {
    for (const type of ['overtime', 'holiday']) {
      try {
        const res = getOrCreateDeptForm_(type, dept);
        if (!res) {
          log.push('SKIP ' + type + ' / ' + dept + ' (作業員0人)');
          continue;
        }
        const label = res.created ? 'CREATED' : 'EXISTS';
        log.push(label + ' ' + type + ' / ' + dept + ' -> ' + res.formUrl);
      } catch (e) {
        log.push('ERROR ' + type + ' / ' + dept + ': ' + e.message);
      }
      Utilities.sleep(500); // API負荷軽減
    }
  }

  const summary = 'buildAllDeptForms:\n' + log.join('\n');
  Logger.log(summary);
  console.log(summary);
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

  // sendEveningMail_ の既存トリガー数をカウント
  const eveningCount = triggers.filter(
    t => t.getHandlerFunction() === 'sendEveningMail_'
      && t.getEventType() === ScriptApp.EventType.CLOCK
  ).length;

  // 夕方 17:10 + 18:10 の2本が必要（既に2本以上あればスキップ）
  if (eveningCount < 2) {
    // 既存をすべて削除してから再作成（半端な状態を解消）
    triggers
      .filter(t => t.getHandlerFunction() === 'sendEveningMail_')
      .forEach(t => ScriptApp.deleteTrigger(t));

    // 夕方 17:10（17時台）
    ScriptApp.newTrigger('sendEveningMail_')
      .timeBased()
      .everyDays(1)
      .atHour(17)
      .nearMinute(10)
      .create();

    // 夕方 18:10（18時台）
    ScriptApp.newTrigger('sendEveningMail_')
      .timeBased()
      .everyDays(1)
      .atHour(18)
      .nearMinute(10)
      .create();
  }

  // 朝 07:10（7時台）— PDF一括生成 + 朝メール
  if (!has('morningBatch_')) {
    ScriptApp.newTrigger('morningBatch_')
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

// ====== 旧フォームトリガー一括削除 ======
// 個別 onFormSubmit トリガー（handleFormSubmit_）をすべて削除する。
// ポーリング方式への移行時に実行が必要。

function cleanupFormSubmitTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'handleFormSubmit_') {
      ScriptApp.deleteTrigger(t);
      deleted++;
    }
  }
  Logger.log('cleanupFormSubmitTriggers_: ' + deleted + ' 個の旧トリガーを削除');
  return deleted;
}

// ====== ポーリングトリガー設定（1分間隔） ======

function setupFormPollTrigger_() {
  // 旧フォーム個別トリガーを先に削除
  cleanupFormSubmitTriggers_();

  // 既存のポーリングトリガーがあればスキップ
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.some(t => t.getHandlerFunction() === 'pollNewResponses_')) return;

  ScriptApp.newTrigger('pollNewResponses_')
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('setupFormPollTrigger_: ポーリングトリガーを作成（1分間隔）');
}

// ====== 全トリガー一括セットアップ ======

function setupAllTriggers_() {
  cleanupFormSubmitTriggers_(); // 旧フォーム個別トリガーを削除
  setupSyncTrigger_();          // マスタ転記（6:00）
  setupTriggers_();              // フォーム毎朝更新（6:30）
  setupMailTriggers_();          // 夕方2回（17:10, 18:10）＋朝バッチ（7:10）
  setupFormPollTrigger_();       // フォーム回答ポーリング（1分間隔）
  Logger.log('全トリガーをセットアップしました。');
}
