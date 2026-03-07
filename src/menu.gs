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
    .addItem('全回答を一括処理（初回/リカバリ用）', 'reprocessAllResponses_')
    .addSeparator()
    .addItem('社内カレンダーマスタ初期投入（2026年度）', 'setupCompanyCalendar2026')
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

  // 夕方2回：16:45頃（1回目）＋ 18:10（2回目）
  if (eveningCount !== 2) {
    // 既存をすべて削除してから再作成（半端な状態を解消）
    triggers
      .filter(t => t.getHandlerFunction() === 'sendEveningMail_')
      .forEach(t => ScriptApp.deleteTrigger(t));

    // 1回目 16:45頃（16時台） ※nearMinuteは0,15,30,45のみ有効
    ScriptApp.newTrigger('sendEveningMail_')
      .timeBased()
      .everyDays(1)
      .atHour(16)
      .nearMinute(45)
      .create();

    // 2回目 18:10（18時台）
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
  setupMailTriggers_();          // 夕方2回（16:45頃＋18:10）＋朝バッチ（7:10）
  setupFormPollTrigger_();       // フォーム回答ポーリング（1分間隔）
  Logger.log('全トリガーをセットアップしました。');
}

// ====== 社内カレンダーマスタ 初期投入（2026年度） ======
// PDFカレンダー「2026年度 株式会社ラインワークス カレンダー」を元に作成。
// 稼働日:257 休日:108（会社休日:89 祝日:19）
// ヘッダ: 日付, 区分(休日/祝日/出勤), 曜日, 備考
// ※ 休日=会社休日（土日含む）, 祝日=国民の祝日, 出勤=稼働日

function setupCompanyCalendar2026() {
  const ss = getDb_();
  const sheetName = SHEET.COMPANY_CALENDAR;
  let sh = ss.getSheetByName(sheetName);
  if (sh) {
    // 既存データがあれば確認
    if (sh.getLastRow() > 1) {
      const ui = SpreadsheetApp.getUi();
      const res = ui.alert('確認', '既存の社内カレンダーマスタを上書きしますか？', ui.ButtonSet.YES_NO);
      if (res !== ui.Button.YES) return;
    }
    sh.clear();
  } else {
    sh = ss.insertSheet(sheetName);
  }

  // ヘッダ
  sh.getRange(1, 1, 1, 4).setValues([['日付', '区分', '曜日', '備考']]);

  // 2026年度（2026/4/1 ～ 2027/3/31）の全休日・祝日データ
  // 区分: 休日=会社休日（土日）, 祝日=国民の祝日, 出勤土曜=土曜出勤日
  const holidays = [
    // ===== 2026年4月 (23稼働, 6会社休日, 1祝日) =====
    ['2026-04-04', '休日', '土', ''],
    ['2026-04-05', '休日', '日', ''],
    ['2026-04-11', '出勤土曜', '土', ''],
    ['2026-04-12', '休日', '日', ''],
    ['2026-04-18', '休日', '土', ''],
    ['2026-04-19', '休日', '日', ''],
    ['2026-04-25', '出勤土曜', '土', ''],
    ['2026-04-26', '休日', '日', ''],
    ['2026-04-29', '祝日', '水', '昭和の日'],

    // ===== 2026年5月 (20稼働, 7会社休日, 4祝日) =====
    ['2026-05-02', '出勤土曜', '土', ''],
    ['2026-05-03', '祝日', '日', '憲法記念日'],
    ['2026-05-04', '祝日', '月', 'みどりの日'],
    ['2026-05-05', '祝日', '火', 'こどもの日'],
    ['2026-05-06', '祝日', '水', '振替休日'],
    ['2026-05-09', '出勤土曜', '土', ''],
    ['2026-05-10', '休日', '日', ''],
    ['2026-05-16', '休日', '土', ''],
    ['2026-05-17', '休日', '日', ''],
    ['2026-05-23', '休日', '土', ''],
    ['2026-05-24', '休日', '日', ''],
    ['2026-05-30', '休日', '土', ''],
    ['2026-05-31', '休日', '日', ''],

    // ===== 2026年6月 (23稼働, 7会社休日, 0祝日) =====
    ['2026-06-06', '休日', '土', ''],
    ['2026-06-07', '休日', '日', ''],
    ['2026-06-13', '休日', '土', ''],
    ['2026-06-14', '休日', '日', ''],
    ['2026-06-20', '出勤土曜', '土', ''],
    ['2026-06-21', '休日', '日', ''],
    ['2026-06-27', '休日', '土', ''],
    ['2026-06-28', '休日', '日', ''],

    // ===== 2026年7月 (23稼働, 7会社休日, 1祝日) =====
    ['2026-07-04', '休日', '土', ''],
    ['2026-07-05', '休日', '日', ''],
    ['2026-07-11', '出勤土曜', '土', ''],
    ['2026-07-12', '休日', '日', ''],
    ['2026-07-18', '休日', '土', ''],
    ['2026-07-19', '休日', '日', ''],
    ['2026-07-20', '祝日', '月', '海の日'],
    ['2026-07-25', '休日', '土', ''],
    ['2026-07-26', '休日', '日', ''],

    // ===== 2026年8月 (19稼働, 11会社休日, 1祝日) =====
    ['2026-08-01', '休日', '土', ''],
    ['2026-08-02', '休日', '日', ''],
    ['2026-08-08', '休日', '土', ''],
    ['2026-08-09', '休日', '日', ''],
    ['2026-08-11', '祝日', '火', '山の日'],
    ['2026-08-13', '休日', '木', 'お盆休み'],
    ['2026-08-14', '休日', '金', 'お盆休み'],
    ['2026-08-15', '休日', '土', 'お盆休み'],
    ['2026-08-16', '休日', '日', ''],
    ['2026-08-22', '休日', '土', ''],
    ['2026-08-23', '休日', '日', ''],
    ['2026-08-29', '休日', '土', ''],
    ['2026-08-30', '休日', '日', ''],

    // ===== 2026年9月 (20稼働, 7会社休日, 3祝日) =====
    ['2026-09-05', '休日', '土', ''],
    ['2026-09-06', '休日', '日', ''],
    ['2026-09-12', '休日', '土', ''],
    ['2026-09-13', '休日', '日', ''],
    ['2026-09-19', '出勤土曜', '土', ''],
    ['2026-09-20', '休日', '日', ''],
    ['2026-09-21', '祝日', '月', '敬老の日'],
    ['2026-09-22', '祝日', '火', '国民の休日'],
    ['2026-09-23', '祝日', '水', '秋分の日'],
    ['2026-09-26', '休日', '土', ''],
    ['2026-09-27', '休日', '日', ''],

    // ===== 2026年10月 (22稼働, 8会社休日, 1祝日) =====
    ['2026-10-03', '休日', '土', ''],
    ['2026-10-04', '休日', '日', ''],
    ['2026-10-10', '休日', '土', ''],
    ['2026-10-11', '休日', '日', ''],
    ['2026-10-12', '祝日', '月', 'スポーツの日'],
    ['2026-10-17', '出勤土曜', '土', ''],
    ['2026-10-18', '休日', '日', ''],
    ['2026-10-24', '休日', '土', ''],
    ['2026-10-25', '休日', '日', ''],
    ['2026-10-31', '休日', '土', ''],

    // ===== 2026年11月 (20稼働, 8会社休日, 2祝日) =====
    ['2026-11-01', '休日', '日', ''],
    ['2026-11-03', '祝日', '火', '文化の日'],
    ['2026-11-07', '休日', '土', ''],
    ['2026-11-08', '休日', '日', ''],
    ['2026-11-14', '休日', '土', ''],
    ['2026-11-15', '休日', '日', ''],
    ['2026-11-21', '休日', '土', ''],
    ['2026-11-22', '休日', '日', ''],
    ['2026-11-23', '祝日', '月', '勤労感謝の日'],
    ['2026-11-28', '休日', '土', ''],
    ['2026-11-29', '休日', '日', ''],

    // ===== 2026年12月 (23稼働, 8会社休日, 0祝日) =====
    ['2026-12-05', '休日', '土', ''],
    ['2026-12-06', '休日', '日', ''],
    ['2026-12-12', '休日', '土', ''],
    ['2026-12-13', '休日', '日', ''],
    ['2026-12-19', '休日', '土', ''],
    ['2026-12-20', '休日', '日', ''],
    ['2026-12-26', '出勤土曜', '土', ''],
    ['2026-12-27', '休日', '日', ''],
    ['2026-12-29', '休日', '月', '年末休暇'],
    ['2026-12-30', '休日', '火', '年末休暇'],
    ['2026-12-31', '休日', '水', '年末休暇'],

    // ===== 2027年1月 (20稼働, 9会社休日, 2祝日) =====
    ['2027-01-01', '祝日', '木', '元旦'],
    ['2027-01-02', '休日', '金', '年始休暇'],
    ['2027-01-03', '休日', '土', '年始休暇'],
    ['2027-01-04', '休日', '日', ''],
    ['2027-01-09', '休日', '土', ''],
    ['2027-01-10', '休日', '日', ''],
    ['2027-01-11', '祝日', '月', '成人の日'],
    ['2027-01-16', '出勤土曜', '土', ''],
    ['2027-01-17', '休日', '日', ''],
    ['2027-01-23', '休日', '土', ''],
    ['2027-01-24', '休日', '日', ''],
    ['2027-01-30', '休日', '土', ''],
    ['2027-01-31', '休日', '日', ''],

    // ===== 2027年2月 (20稼働, 6会社休日, 2祝日) =====
    ['2027-02-06', '休日', '土', ''],
    ['2027-02-07', '休日', '日', ''],
    ['2027-02-11', '祝日', '木', '建国記念日'],
    ['2027-02-13', '休日', '土', ''],
    ['2027-02-14', '休日', '日', ''],
    ['2027-02-20', '休日', '土', ''],
    ['2027-02-21', '休日', '日', ''],
    ['2027-02-27', '出勤土曜', '土', ''],
    ['2027-02-28', '休日', '日', ''],

    // ===== 2027年3月 (24稼働, 5会社休日, 2祝日) =====
    ['2027-03-06', '休日', '土', ''],
    ['2027-03-07', '休日', '日', ''],
    ['2027-03-13', '出勤土曜', '土', ''],
    ['2027-03-14', '休日', '日', ''],
    ['2027-03-20', '休日', '土', ''],
    ['2027-03-21', '祝日', '日', '春分の日'],
    ['2027-03-22', '祝日', '月', '振替休日'],
    ['2027-03-27', '出勤土曜', '土', ''],
    ['2027-03-28', '休日', '日', ''],
  ];

  // データ書き込み
  if (holidays.length > 0) {
    sh.getRange(2, 1, holidays.length, 4).setValues(holidays);
  }

  // 書式設定
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#d9e2f3');
  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 80);
  sh.setColumnWidth(3, 50);
  sh.setColumnWidth(4, 150);

  // 条件付き書式：祝日行を赤背景
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('祝日')
    .setBackground('#fce4e4')
    .setRanges([sh.getRange(2, 1, holidays.length, 4)])
    .build();
  // 出勤土曜行を青背景
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('出勤土曜')
    .setBackground('#e4f0fc')
    .setRanges([sh.getRange(2, 1, holidays.length, 4)])
    .build();
  sh.setConditionalFormatRules([rule1, rule2]);

  SpreadsheetApp.flush();
  Logger.log('社内カレンダーマスタ 2026年度データを投入しました（' + holidays.length + '行）');
}
