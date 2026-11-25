// Cutover.gs - VPS 版への切り替え（移転モード）
//
// スクリプトプロパティ MIGRATED_TO_URL に新しいアプリの起点（https://ホスト名）が入っている間は「移転モード」。
// プロパティが無ければ、今までと同じ動き（この判定の読み込み以外は何も変えない）。
//
// 移転モードのとき:
// - doGet: アプリの代わりに「新しいアドレスへ移りました」のページを返す。?page= や ?date= などの
//   パラメータは新しいアドレスへそのまま引き継ぐ（統合メール通知などの本文のリンクのため）。
// - DB を開く 1 か所 getDb_()（config.gs）の先頭で止める（罠 47）。画面の関数も保守の関数もここを通る。
//   さらに、画面から呼ぶ関数（api_*）と公開の保守関数の先頭にも assertNotMigrated_() を置いた
//   （api_getHolidayCandidateDates のように DB の失敗を握りつぶして別の方式で動く関数があるため）。
// - getDb_() を通らない入口（gap-4 F-7: setup*Trigger_・cleanupFormSubmitTriggers_・
//   reprocessAllResponses_・debugProcessLatest_・debugPollStatus_ など）も先頭で止める。
// - 時刻のトリガーの関数（pollNewResponses_・nightlyUpdateAllForms_・morningBatchNoMail_・
//   morningBatch_・syncAllMasters・sendEveningMail_ など）は何もせずに戻る。
// - doPost の内部用の入口（Internal.gs の vps.* / internal.*）は止めない。token を確かめた実行の中だけ
//   _vpsInternalCall_ を true にし、getDb_() を通す。
// - doPost の Sebastian の action は、SEBASTIAN_FORWARD_URL があれば本文をそのまま VPS へ転送する。
//   無い間（切り替えの途中）は {"error":"service is moving, retry later"} を返す（計画 6-3・B-17）。
//
// 下の運用関数は、スクリプトの持ち主が GAS エディタから手で実行する（罠 48）。

var MIGRATED_TO_URL_KEY = 'MIGRATED_TO_URL';            // 移転モードの印＝新しいアプリの起点
var VPS_ORIGIN_KEY = 'VPS_ORIGIN';                      // 引数なしの包みが読む起点（前もって手で入れる）
var SEBASTIAN_FORWARD_URL_KEY = 'SEBASTIAN_FORWARD_URL'; // Sebastian の転送先（<起点>/api/sebastian）
var MIGRATION_CLOSED_FORMS_KEY = 'MIGRATION_CLOSED_FORMS'; // 閉じたフォームの記録（戻すときに使う）

// 起点の形: https://ホスト名[:ポート]（パスなし）。既定値は持たない（罠 24d）。
var VPS_ORIGIN_PATTERN_ = /^https:\/\/[A-Za-z0-9.-]+(:\d{1,5})?$/;

// このプロジェクトの時刻のトリガーの関数（GAS のトリガー画面で確かめる。gap-4 F-17）
var MIGRATION_TRIGGER_FUNCTIONS = [
  'pollNewResponses_',
  'nightlyUpdateAllForms_',
  'morningBatchNoMail_',
  'syncAllMasters',
  'morningBatch_',
  'sendEveningMail_',
  'handleFormSubmit_',
];

// 実行の中だけの印。Internal.gs の doPost の入口が token を確かめた後でだけ true にする。
var _vpsInternalCall_ = false;

// 実行ごとに 1 回だけプロパティを読む（getDb_ は 1 回の実行で何度も呼ばれるため）。
var _migratedToUrlCache_ = null;

function getMigratedToUrl_() {
  if (_migratedToUrlCache_ !== null) return _migratedToUrlCache_;
  try {
    var v = String(PropertiesService.getScriptProperties().getProperty(MIGRATED_TO_URL_KEY) || '').trim();
    _migratedToUrlCache_ = v;
    return v;
  } catch (e) {
    // プロパティが読めないときは今までと同じ動きにする（キャッシュしない）。
    Logger.log('getMigratedToUrl_: プロパティを読めませんでした: ' + e.message);
    return '';
  }
}

function invalidateMigrationCache_() {
  _migratedToUrlCache_ = null;
}

function isMigrationMode_() {
  return !!getMigratedToUrl_();
}

function normalizeVpsOrigin_(origin) {
  return String(origin || '').trim().replace(/\/+$/, '');
}

function isValidVpsOrigin_(origin) {
  return VPS_ORIGIN_PATTERN_.test(String(origin || ''));
}

function migrationMessage_() {
  var url = getMigratedToUrl_();
  return '残業・休日出勤申請は新しいアドレスへ移りました。' +
    (isValidVpsOrigin_(url) ? url + '/ を開いてください。' : '新しいアドレスは管理者にお問い合わせください。') +
    '（この画面からは操作できません）';
}

// 画面の関数・保守の関数・getDb_() の先頭で呼ぶ。移転モードなら例外にする。
function assertNotMigrated_() {
  if (isMigrationMode_()) {
    throw new Error(migrationMessage_());
  }
}

// getDb_() から呼ぶ。内部用の入口（token を確かめた実行）だけは通す。
function assertDbOpenAllowed_() {
  if (_vpsInternalCall_) return;
  assertNotMigrated_();
}

// トリガーの関数の先頭で呼ぶ。移転モードなら true（呼び出し側は何もせずに戻る）。
function skipBecauseMigrated_(fnName) {
  if (!isMigrationMode_()) return false;
  Logger.log(fnName + ': 移転モード（' + getMigratedToUrl_() + '）のため何もしません');
  return true;
}

// ===== doGet の移転ページ =====

function escapeHtmlCutover_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ?page=approve2&date=2026-11-25 → https://<起点>/?page=approve2&date=2026-11-25
function migratedTargetUrl_(origin, params) {
  var base = normalizeVpsOrigin_(origin) + '/';
  var keys = Object.keys(params || {}).sort(function(a, b) {
    if (a === 'page') return -1;
    if (b === 'page') return 1;
    return a < b ? -1 : (a > b ? 1 : 0);
  });
  var parts = [];
  keys.forEach(function(k) {
    var v = params[k];
    if (v == null) return;
    parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(String(v)));
  });
  return parts.length ? base + '?' + parts.join('&') : base;
}

// GAS の iframe の外へ出るため <base target="_top"> とボタンの target="_top" を使う（罠 29・gap-2 S4）。
// 自動の移動は「動けばよい」扱い（クリックなしの window.top の移動はブラウザが止めることがある）。
function renderMigratedPage_(e) {
  var origin = getMigratedToUrl_();
  var title = '残業・休日出勤申請（移転しました）';
  if (!isValidVpsOrigin_(origin)) {
    Logger.log('renderMigratedPage_: MIGRATED_TO_URL の形が正しくないため、固定の文言を出します');
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="utf-8"></head>' +
      '<body style="font-family:sans-serif;margin:40px 16px">' +
      '<p>残業・休日出勤申請は新しいアドレスへ移りました。新しいアドレスは管理者にお問い合わせください。</p>' +
      '</body></html>')
      .setTitle(title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var params = (e && e.parameter) ? e.parameter : {};
  var target = migratedTargetUrl_(origin, params);
  var hasParams = Object.keys(params).length > 0;
  var autoMove = hasParams
    ? '<script>setTimeout(function(){try{window.top.location.href=' +
      JSON.stringify(target).replace(/</g, '\\u003c') + ';}catch(err){}},1500);</script>'
    : '';
  var html = [
    '<!DOCTYPE html><html><head><meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<base target="_top">',
    '<style>body{font-family:sans-serif;margin:40px 16px;line-height:1.7;color:#222}',
    '.go{display:inline-block;padding:18px 36px;font-size:22px;font-weight:bold;color:#fff;background:#0a4d73;',
    'border-radius:8px;text-decoration:none}.url{font-size:14px;color:#555;word-break:break-all}</style></head><body>',
    '<h1 style="font-size:20px">残業・休日出勤申請は新しいアドレスへ移りました</h1>',
    '<p>下のボタンを押して、新しいアドレスで開いてください。</p>',
    '<p><a class="go" href="' + escapeHtmlCutover_(target) + '" target="_top">新しいアドレスで開く</a></p>',
    '<p class="url">新しいアドレス: ' + escapeHtmlCutover_(normalizeVpsOrigin_(origin) + '/') +
      '<br>ブックマークをお使いの方は、差し替えをお願いします。</p>',
    autoMove,
    '</body></html>',
  ].join('');
  return HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===== Sebastian の転送（移転モードのときだけ doPost から呼ぶ） =====

// 受けた本文（生の文字列）をそのまま VPS の /api/sebastian へ POST し、返事の本文をそのまま返す。
// 転送できないときは、今の doPost の失敗と同じ形 {error, action} を返す（HTTP 200）。
function forwardSebastianToVps_(rawBody, action) {
  var url = '';
  try {
    url = String(PropertiesService.getScriptProperties().getProperty(SEBASTIAN_FORWARD_URL_KEY) || '').trim();
  } catch (e) {
    url = '';
  }
  if (!url) {
    return _apiResponse_({ error: 'service is moving, retry later' });
  }
  var res;
  try {
    res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: String(rawBody == null ? '' : rawBody),
      muteHttpExceptions: true,
      followRedirects: false,
    });
  } catch (err) {
    Logger.log('forwardSebastianToVps_: 転送に失敗: ' + err.message);
    var fail = { error: 'forward failed: ' + String((err && err.message) || err) };
    if (action) fail.action = action;
    return _apiResponse_(fail);
  }
  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code !== 200) {
    Logger.log('forwardSebastianToVps_: VPS が HTTP ' + code + ' を返しました');
    var bad = { error: 'forward failed: HTTP ' + code };
    if (action) bad.action = action;
    return _apiResponse_(bad);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

// ===== 運用関数を動かしてよい人の確認（罠 48） =====
// Web アプリは持ち主として実行されるが、画面（google.script.run）から呼ぶと操作している人は
// 利用者のアカウントになるので、持ち主以外は通らない。エディタから実行すれば成り立つ。
function assertCutoverOperator_() {
  var active = '';
  var effective = '';
  try { active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e1) { active = ''; }
  try { effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (e2) { effective = ''; }
  if (!active || !effective || active !== effective) {
    throw new Error('この操作は、スクリプトの持ち主が GAS エディタから実行してください。');
  }
  return active;
}

// 移転モードでも読めるよう、getDb_() の判定を通さずに DB を開く（読むだけの運用関数の中でだけ使う）。
function openDbForCutoverOps_() {
  return SpreadsheetApp.getActive();
}

// FormMap の formId 一覧（読むだけ）
function listFormMapForms_() {
  var sh = openDbForCutoverOps_().getSheetByName(SHEET.FORM_MAP);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var H = values[0].map(function(h) { return normalize_(h); });
  var c = { type: H.indexOf('type'), dept: H.indexOf('dept'), formId: H.indexOf('formId'), isActive: H.indexOf('isActive') };
  if (c.formId < 0) return [];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var formId = normalize_(values[r][c.formId]);
    if (!formId) continue;
    out.push({
      formId: formId,
      type: c.type >= 0 ? normalize_(values[r][c.type]) : '',
      dept: c.dept >= 0 ? normalize_(values[r][c.dept]) : '',
      isActive: c.isActive >= 0 ? values[r][c.isActive] : '',
    });
  }
  return out;
}

// ===== GAS エディタから実行する運用関数 =====

// 例: enterMigrationMode('https://overtime.lineworks-local.net')
// 起点は必ず引数で渡す（既定値を持たない。ほかのアプリへ利用者を送らないため。罠 24d）。
function enterMigrationMode(vpsOrigin) {
  assertCutoverOperator_();
  var value = normalizeVpsOrigin_(vpsOrigin);
  if (!isValidVpsOrigin_(value)) {
    throw new Error('起点は https://ホスト名 の形（パスなし）で、引数に渡してください: ' + value);
  }
  PropertiesService.getScriptProperties().setProperty(MIGRATED_TO_URL_KEY, value);
  invalidateMigrationCache_();
  Logger.log('移転モードにしました: ' + value +
    '\nSebastian の転送はまだ止めています。VPS の確認が済んだら setSebastianForwardUrl を実行してください。');
  return value;
}

// エディタの実行ボタンは引数を渡せないので、当日はこれを実行する。
// 起点はスクリプトプロパティ VPS_ORIGIN から読む（前もって手で入れる。コードに書かない）。
function enterMigrationModeForOvertime() {
  var origin = '';
  try { origin = String(PropertiesService.getScriptProperties().getProperty(VPS_ORIGIN_KEY) || ''); } catch (e) { origin = ''; }
  if (!origin.trim()) {
    throw new Error('スクリプトプロパティ ' + VPS_ORIGIN_KEY + ' に新しいアプリの起点（https://…）を入れてから実行してください。');
  }
  return enterMigrationMode(origin);
}

// Sebastian の転送を始める（計画 B-17）。移転モードの起点 + /api/sebastian を転送先にする。
function setSebastianForwardUrl() {
  assertCutoverOperator_();
  invalidateMigrationCache_();
  var origin = getMigratedToUrl_();
  if (!isValidVpsOrigin_(origin)) {
    throw new Error('移転モードではありません。先に enterMigrationModeForOvertime を実行してください。');
  }
  var url = origin + '/api/sebastian';
  PropertiesService.getScriptProperties().setProperty(SEBASTIAN_FORWARD_URL_KEY, url);
  Logger.log('Sebastian の転送先を設定しました: ' + url);
  return url;
}

// Sebastian の転送を止める（移転モードのままなら "service is moving" を返す状態に戻る）。
function clearSebastianForwardUrl() {
  assertCutoverOperator_();
  PropertiesService.getScriptProperties().deleteProperty(SEBASTIAN_FORWARD_URL_KEY);
  Logger.log('Sebastian の転送先を消しました。');
  return true;
}

// 戻すとき。移転モードと Sebastian の転送を外し、今までと同じ動きに戻す。
// トリガーは restoreTriggersAfterRollback で、フォームは reopenFormsAfterRollback で戻す。
function exitMigrationMode() {
  assertCutoverOperator_();
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(MIGRATED_TO_URL_KEY);
  props.deleteProperty(SEBASTIAN_FORWARD_URL_KEY);
  invalidateMigrationCache_();
  Logger.log('移転モードと Sebastian の転送を外しました。');
  return true;
}

// 今の状態をログに出す。何も変えない。移転モードでも動く（getDb_ の判定を通さない。gap-4 F-8）。
function showMigrationStatus() {
  assertCutoverOperator_();
  var props = PropertiesService.getScriptProperties();
  invalidateMigrationCache_();
  var status = {
    migratedTo: getMigratedToUrl_(),
    vpsOrigin: String(props.getProperty(VPS_ORIGIN_KEY) || ''),
    sebastianForwardUrl: String(props.getProperty(SEBASTIAN_FORWARD_URL_KEY) || ''),
    internalTokenSet: !!props.getProperty('VPS_INTERNAL_TOKEN'),
    pollLastTs: String(props.getProperty('POLL_LAST_TS') || ''),
    triggers: ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); }),
    closedForms: JSON.parse(props.getProperty(MIGRATION_CLOSED_FORMS_KEY) || '[]').length,
    forms: [],
    unprocessedResponses: 0,
  };
  // 旧フォームの受付状態と、POLL_LAST_TS 以降の回答の数（取り込まれていない見込みの数）
  var since = status.pollLastTs ? new Date(status.pollLastTs) : null;
  listFormMapForms_().forEach(function(f) {
    var row = { formId: f.formId, type: f.type, dept: f.dept, accepting: null, newResponses: null };
    try {
      var form = FormApp.openById(f.formId);
      row.accepting = form.isAcceptingResponses();
      if (since && !isNaN(since.getTime())) {
        row.newResponses = form.getResponses(since).length;
        status.unprocessedResponses += row.newResponses;
      }
    } catch (e) {
      row.error = e.message;
    }
    status.forms.push(row);
  });
  Logger.log(
    '移転モード: ' + (status.migratedTo || '（なし。今までと同じ動き）') +
    '\nSebastian の転送先: ' + (status.sebastianForwardUrl || '（なし）') +
    '\n内部用の token: ' + (status.internalTokenSet ? 'あり' : 'なし') +
    '\nPOLL_LAST_TS: ' + (status.pollLastTs || '（なし）') +
    '\nトリガー: ' + (status.triggers.length ? status.triggers.join(', ') : 'なし') +
    '\n旧フォーム: ' + status.forms.length + ' 本（受付中 ' +
      status.forms.filter(function(f) { return f.accepting === true; }).length + ' 本）' +
    '\nPOLL_LAST_TS 以降の回答: ' + status.unprocessedResponses + ' 件');
  return status;
}

// 旧フォーム（FormMap の全行。本番は 21 本）の受付を閉じる。フォームと回答は消さない（gap-4 F-19）。
// 変えたものを記録し、戻すときは reopenFormsAfterRollback で記録したものだけを開け直す。
function closeAllFormsForMigration() {
  assertCutoverOperator_();
  var props = PropertiesService.getScriptProperties();
  var record = JSON.parse(props.getProperty(MIGRATION_CLOSED_FORMS_KEY) || '[]');
  var known = {};
  record.forEach(function(r) { known[r.formId] = true; });
  var closed = [];
  var already = [];
  var errors = [];
  listFormMapForms_().forEach(function(f) {
    try {
      var form = FormApp.openById(f.formId);
      if (form.isAcceptingResponses()) {
        form.setAcceptingResponses(false);
        closed.push(f);
        if (!known[f.formId]) {
          record.push({ formId: f.formId, type: f.type, dept: f.dept, closedAt: new Date().toISOString() });
          known[f.formId] = true;
        }
      } else {
        already.push(f);
      }
    } catch (e) {
      errors.push(f.formId + ' (' + f.type + '/' + f.dept + '): ' + e.message);
    }
  });
  props.setProperty(MIGRATION_CLOSED_FORMS_KEY, JSON.stringify(record));
  Logger.log('旧フォームの受付を閉じました: ' + closed.length + ' 本' +
    closed.map(function(f) { return '\n  閉じた: ' + f.type + ' / ' + f.dept + ' (' + f.formId + ')'; }).join('') +
    '\nもともと閉じていた: ' + already.length + ' 本' +
    already.map(function(f) { return '\n  そのまま: ' + f.type + ' / ' + f.dept + ' (' + f.formId + ')'; }).join('') +
    '\n失敗: ' + errors.length + ' 本' + errors.map(function(s) { return '\n  ' + s; }).join(''));
  return { closed: closed.length, alreadyClosed: already.length, errors: errors };
}

// 戻すとき。closeAllFormsForMigration で閉じたフォームだけ受付を開け直す。
function reopenFormsAfterRollback() {
  assertCutoverOperator_();
  var props = PropertiesService.getScriptProperties();
  var record = JSON.parse(props.getProperty(MIGRATION_CLOSED_FORMS_KEY) || '[]');
  var reopened = [];
  var left = [];
  var errors = [];
  record.forEach(function(r) {
    try {
      var form = FormApp.openById(r.formId);
      if (!form.isAcceptingResponses()) form.setAcceptingResponses(true);
      reopened.push(r);
    } catch (e) {
      errors.push(r.formId + ' (' + r.type + '/' + r.dept + '): ' + e.message);
      left.push(r);
    }
  });
  if (left.length) {
    props.setProperty(MIGRATION_CLOSED_FORMS_KEY, JSON.stringify(left));
  } else {
    props.deleteProperty(MIGRATION_CLOSED_FORMS_KEY);
  }
  Logger.log('旧フォームの受付を開け直しました: ' + reopened.length + ' 本' +
    reopened.map(function(r) { return '\n  開けた: ' + r.type + ' / ' + r.dept + ' (' + r.formId + ')'; }).join('') +
    '\n失敗: ' + errors.length + ' 本' + errors.map(function(s) { return '\n  ' + s; }).join(''));
  return { reopened: reopened.length, errors: errors };
}

// このプロジェクトの時刻のトリガー（と旧フォームのトリガー）を消し、消したものを記録する。
// 見えるのは実行している人のトリガーだけ（gap-4 F-17）。ほかの人の分は GAS の画面で確かめる。
function removeAllTriggersForMigration() {
  assertCutoverOperator_();
  var removed = [];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    var isClock = false;
    try { isClock = t.getEventType() === ScriptApp.EventType.CLOCK; } catch (e) { isClock = false; }
    if (isClock || MIGRATION_TRIGGER_FUNCTIONS.indexOf(fn) !== -1) {
      var id = '';
      try { id = t.getUniqueId(); } catch (e2) { id = ''; }
      ScriptApp.deleteTrigger(t);
      removed.push(fn + (id ? ' (' + id + ')' : ''));
    }
  });
  var left = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });
  Logger.log('削除したトリガー: ' + (removed.length ? removed.join(', ') : 'なし') +
    '\n残っているトリガー: ' + (left.length ? left.join(', ') : 'なし'));
  return { removed: removed, left: left };
}

// 戻すとき（exitMigrationMode の後）。計画 6-6 C-3 のとおり、朝バッチとマスタ同期の 2 本だけ作り直す
// （同じ関数のトリガーがあれば作らない）。旧フォームの 2 本は決定 #11 のとおり戻さない。
function restoreTriggersAfterRollback() {
  assertCutoverOperator_();
  invalidateMigrationCache_();
  if (isMigrationMode_()) throw new Error('移転モードのままです。先に exitMigrationMode を実行してください。');
  var before = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });
  setupSyncTrigger_();
  setupMorningBatchTrigger_();
  var after = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });
  Logger.log('トリガー（前）: ' + (before.join(', ') || 'なし') + '\nトリガー（後）: ' + (after.join(', ') || 'なし'));
  return after;
}

// 戻すとき（オーナーが旧フォームも戻すと決めたとき）。setupAllTriggers_ と同じ 4 本を作り直す。
function restoreAllTriggersAfterRollback() {
  assertCutoverOperator_();
  invalidateMigrationCache_();
  if (isMigrationMode_()) throw new Error('移転モードのままです。先に exitMigrationMode を実行してください。');
  var before = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });
  setupAllTriggers_();
  var after = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });
  Logger.log('トリガー（前）: ' + (before.join(', ') || 'なし') + '\nトリガー（後）: ' + (after.join(', ') || 'なし'));
  return after;
}

// 移転モードのメニューの 1 項目から呼ぶ。新しいアドレスを出すだけ。
function showMigratedAddress() {
  var msg = migrationMessage_();
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return msg;
}
