// Internal.gs - VPS から呼ぶ内部用の入口（計画 4-6a・決定 B）
//
// doPost（api.gs）は、本文の action が "vps." か "internal." で始まるとき（または "internal": true のとき）
// だけここへ回す。Sebastian の 4 action とは別の道で、移転モードでも止めない。
//
// 認証: 本文の token をスクリプトプロパティ VPS_INTERNAL_TOKEN（32 字以上。SEBASTIAN_API_TOKEN とは別の値）と
//       定数時間で比べる。合わなければ何もしない。
// 冪等: jobId ごとに結果をスクリプトプロパティ（VPSJOB_<jobId>）に残し、同じ jobId には処理せずに同じ結果を返す
//       （"replayed": true）。残すのは「最終の結果」（ok:true か、再試行しても変わらない失敗）だけ。
//       busy・row_not_synced・pdf_failed・internal など再試行できる失敗は残さないので、同じ jobId で送り直せる。
//       最大 VPS_JOB_MAX_ENTRIES_ 件・VPS_JOB_TTL_MS_ まで（古いものから消す）。
// 直列: 今のコードと同じスクリプトのロック（LockService.getScriptLock）を tryLock(20000) で取る。
//       取れなければ {"ok":false,"error":"busy","retryable":true}。中で呼ぶ今の関数（generatePdf* など）は
//       今と同じく自分でロックを取り直して放す（Sebastian の経路と同じ）。
// 中身: 今の関数を、今と同じ順序・条件で呼ぶだけ（PDF・台帳の作りは変えない）。
//
// action（vps.* が依頼の名前。internal.* は計画 4-6a の名前で、同じ処理への別名）:
//   vps.ping / internal.ping
//   vps.renderPdf        {requestId, jobId, mode?, regenerate?, oldPdfFileId?, ledger?, requireWorkDone?, workLog?, expect?}
//   vps.upsertLedger     {requestId, jobId}
//   vps.removeLedgerRow  {requestId, requestType, targetDate, jobId}
//   vps.morningBatch     {targetDate?, jobId}
//   vps.ledgerRewriteDay {targetDate, jobId}
//   internal.afterApprove / afterComplete / regenerate / afterDelete / batchPdf / ledgerRewriteDay（4-6 の表の経路）

var VPS_INTERNAL_API_VERSION_ = '2026-11-26.1';
var VPS_INTERNAL_TOKEN_KEY_ = 'VPS_INTERNAL_TOKEN';
var VPS_JOB_PREFIX_ = 'VPSJOB_';
var VPS_JOB_MAX_ENTRIES_ = 300;
var VPS_JOB_TTL_MS_ = 7 * 24 * 60 * 60 * 1000;
var VPS_JOB_RUNNING_STALE_MS_ = 7 * 60 * 1000; // GAS の 1 回の上限（6 分）より長い
var VPS_JOB_MAX_VALUE_LEN_ = 8500;             // スクリプトプロパティ 1 値の上限（9KB）より小さく
var VPS_LOCK_WAIT_MS_ = 20000;

// 本文がこの入口あてかどうか
function isVpsInternalRequest_(body) {
  if (!body || typeof body !== 'object') return false;
  if (body.internal === true) return true;
  var a = String(body.action || '');
  return a.indexOf('vps.') === 0 || a.indexOf('internal.') === 0;
}

// 定数時間の比較（長さの違いも同じ手間で比べる）
function vpsTimingSafeEqual_(a, b) {
  a = String(a == null ? '' : a);
  b = String(b == null ? '' : b);
  var len = Math.max(a.length, b.length);
  var diff = a.length ^ b.length;
  for (var i = 0; i < len; i++) {
    var ca = i < a.length ? a.charCodeAt(i) : 0;
    var cb = i < b.length ? b.charCodeAt(i) : 0;
    diff |= (ca ^ cb);
  }
  return diff === 0;
}

function vpsFail_(code, message, retryable, extra) {
  var out = { ok: false, error: code, message: String(message || code), retryable: !!retryable };
  if (extra) for (var k in extra) out[k] = extra[k];
  return out;
}

function vpsFmtDateTime_(v) {
  if (!v) return '';
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : fmtDate_(v, 'yyyy-MM-dd HH:mm:ss');
  return String(v);
}

// 'yyyy-MM-dd' → その日の正午（JST）の Date。今の朝バッチは「昨日の今の時刻」の Date を渡しているが、
// 使う側（fmtDate_・getFullYear など）は日付だけを見るので、正午にして時間帯のずれを避ける。
function vpsParseYmd_(ymd) {
  var s = String(ymd || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  var d = new Date(s + 'T12:00:00+09:00');
  return isNaN(d.getTime()) ? null : d;
}

// pdfGeneratedAt（Date か 'yyyy-MM-dd HH:mm:ss'＝JST の文字）→ エポックのミリ秒。無ければ null
function vpsEpochMs_(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v.getTime();
  var m = /^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}(?::\d{2})?)$/.exec(String(v).trim());
  if (!m) return null;
  var t = new Date(m[1] + 'T' + (m[2].length === 5 ? m[2] + ':00' : m[2]) + '+09:00').getTime();
  return isNaN(t) ? null : t;
}

function vpsYesterday_() {
  var d = new Date();
  d.setDate(d.getDate() - 1);
  return vpsParseYmd_(fmtDate_(d, 'yyyy-MM-dd'));
}

// ===== 冪等（jobId ごとの結果） =====

function vpsJobKey_(jobId) {
  return VPS_JOB_PREFIX_ + String(jobId);
}

function vpsJobRead_(jobId) {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(vpsJobKey_(jobId));
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    Logger.log('vpsJobRead_: ' + e.message);
    return null;
  }
}

function vpsJobWrite_(jobId, entry) {
  try {
    var json = JSON.stringify(entry);
    if (json.length > VPS_JOB_MAX_VALUE_LEN_ && entry.result) {
      // 長すぎる結果（朝バッチのエラー詳細など）は詳細を削って残す
      var slim = JSON.parse(JSON.stringify(entry));
      if (slim.result.errors) slim.result.errors = slim.result.errors.slice(0, 5).map(function(s) { return String(s).slice(0, 300); });
      slim.result.truncated = true;
      json = JSON.stringify(slim);
      if (json.length > VPS_JOB_MAX_VALUE_LEN_) {
        json = JSON.stringify({ t: entry.t, state: entry.state, action: entry.action,
          result: { ok: entry.result.ok, error: entry.result.error, truncated: true } });
      }
    }
    PropertiesService.getScriptProperties().setProperty(vpsJobKey_(jobId), json);
  } catch (e) {
    Logger.log('vpsJobWrite_: ' + e.message);
  }
}

function vpsJobDelete_(jobId) {
  try { PropertiesService.getScriptProperties().deleteProperty(vpsJobKey_(jobId)); } catch (e) { /* ignore */ }
}

// 古い記録を消す（件数と日数の上限）
function vpsJobPrune_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    var now = Date.now();
    var entries = [];
    Object.keys(all).forEach(function(k) {
      if (k.indexOf(VPS_JOB_PREFIX_) !== 0) return;
      var t = 0;
      try { t = Number(JSON.parse(all[k]).t) || 0; } catch (e) { t = 0; }
      entries.push({ k: k, t: t });
    });
    entries.sort(function(a, b) { return a.t - b.t; });
    var excess = entries.length - VPS_JOB_MAX_ENTRIES_;
    entries.forEach(function(e, i) {
      if (i < excess || now - e.t > VPS_JOB_TTL_MS_) props.deleteProperty(e.k);
    });
  } catch (e) {
    Logger.log('vpsJobPrune_: ' + e.message);
  }
}

// ===== 受け口 =====

// doPost から呼ぶ。返すのは ContentService の JSON（いつも HTTP 200）。
function handleVpsInternal_(body) {
  return _apiResponse_(dispatchVpsInternal_(body));
}

function dispatchVpsInternal_(body) {
  var started = Date.now();
  var action = String(body.action || '');
  var jobId = body.jobId != null ? String(body.jobId) : '';

  // 認証（何かを読む・書く前に）
  var props = PropertiesService.getScriptProperties();
  var expected = String(props.getProperty(VPS_INTERNAL_TOKEN_KEY_) || '');
  var sebastian = String(props.getProperty('SEBASTIAN_API_TOKEN') || '');
  if (expected.length < 32 || (sebastian && vpsTimingSafeEqual_(expected, sebastian))) {
    // 未設定・短すぎる・Sebastian と同じ値のときは誰も通さない
    return vpsFail_('unauthorized', 'unauthorized', false);
  }
  if (!vpsTimingSafeEqual_(body.token, expected)) {
    return vpsFail_('unauthorized', 'unauthorized', false);
  }

  // ここから先だけ getDb_() を通す（移転モードでも）
  _vpsInternalCall_ = true;
  try {
    var params = vpsMergeParams_(body);
    var route = VPS_INTERNAL_ROUTES_[action];
    if (!route) return vpsFail_('bad_request', 'unknown action: ' + action, false, { action: action });

    if (route.readOnly) {
      var r0 = route.fn(params);
      r0.action = action;
      r0.durationMs = Date.now() - started;
      return r0;
    }

    if (!jobId || !/^[A-Za-z0-9._:-]{1,100}$/.test(jobId)) {
      return vpsFail_('bad_request', 'jobId is required (1-100 chars of [A-Za-z0-9._:-])', false, { action: action });
    }

    // 同じ jobId の結果があれば、処理せずに返す
    var prev = vpsJobRead_(jobId);
    if (prev && prev.state === 'done' && prev.result) {
      if (prev.action && prev.action !== action) {
        return vpsFail_('bad_request', 'jobId was already used for ' + prev.action, false, { action: action, jobId: jobId });
      }
      var replay = JSON.parse(JSON.stringify(prev.result));
      replay.replayed = true;
      return replay;
    }
    if (prev && prev.state === 'running' && Date.now() - Number(prev.t || 0) < VPS_JOB_RUNNING_STALE_MS_) {
      return vpsFail_('busy', 'same jobId is still running', true, { action: action, jobId: jobId });
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(VPS_LOCK_WAIT_MS_)) {
      return vpsFail_('busy', 'script lock not available', true, { action: action, jobId: jobId });
    }
    var result;
    try {
      vpsJobWrite_(jobId, { t: Date.now(), state: 'running', action: action });
      try {
        result = route.fn(params);
      } catch (err) {
        var msg = String((err && err.message) || err);
        // 今の関数の中のロック待ち切れ（waitLock の例外）は busy として返す
        var isLock = /lock|ロック/i.test(msg) && /time|timed|タイムアウト|取得/i.test(msg);
        result = vpsFail_(isLock ? 'busy' : 'internal', msg, true);
      }
      result.action = action;
      result.jobId = jobId;
      if (params.requestId && result.requestId === undefined) result.requestId = String(params.requestId);
      result.durationMs = Date.now() - started;

      var final = result.ok === true || result.retryable === false;
      if (final) {
        vpsJobWrite_(jobId, { t: Date.now(), state: 'done', action: action, result: result });
        vpsJobPrune_();
      } else {
        vpsJobDelete_(jobId);
      }
      return result;
    } finally {
      try { lock.releaseLock(); } catch (e) { /* 中の関数が既に放している */ }
    }
  } finally {
    _vpsInternalCall_ = false;
  }
}

// 計画 4-6a の本文（args の中）と、依頼の本文（上の段）のどちらでも受ける。上の段が勝つ。
function vpsMergeParams_(body) {
  var p = {};
  var args = (body && body.args && typeof body.args === 'object') ? body.args : {};
  Object.keys(args).forEach(function(k) { p[k] = args[k]; });
  Object.keys(body || {}).forEach(function(k) {
    if (k === 'token' || k === 'args' || k === 'internal') return;
    p[k] = body[k];
  });
  return p;
}

// ===== action の表 =====

var VPS_INTERNAL_ROUTES_ = {
  'vps.ping': { readOnly: true, fn: function(p) { return vpsPing_(); } },
  'internal.ping': { readOnly: true, fn: function(p) { return vpsPing_(); } },

  'vps.renderPdf': { fn: function(p) { return vpsRenderPdf_(p); } },
  'vps.upsertLedger': { fn: function(p) { return vpsUpsertLedger_(p); } },
  'vps.removeLedgerRow': { fn: function(p) { return vpsRemoveLedgerRow_(p); } },
  'vps.morningBatch': { fn: function(p) { return vpsMorningBatch_(p); } },
  'vps.ledgerRewriteDay': { fn: function(p) { return vpsLedgerRewriteDay_(p); } },
  'vps.ledgerRebuild': { fn: function(p) { return vpsLedgerRebuild_(p); } },

  // 計画 4-6 の表の経路（今の画面の関数と同じ順序・条件）
  // 一次承認: actualEndAt があるときだけ作る。PDF が例外なら台帳は書かない（approval.gs api_approveRequest）
  'internal.afterApprove': { fn: function(p) {
    return vpsRenderPdf_(vpsWith_(p, { mode: 'auto', requireWorkDone: true, ledger: 'onSuccess', regenerate: false }));
  } },
  // 完了（画面・総務の手動入力・Sebastian）: PDF の成否に関係なく台帳を upsert。Sebastian はいつも XLOOKUP 方式
  // 仕様 7-2: statusAtEnqueue が approved のときだけ作る。wlOverride は screen・admin のときだけ（sebastian は override なし）
  'internal.afterComplete': { fn: function(p) {
    var isSeb = String(p.source || '') === 'sebastian';
    var fixed = { mode: isSeb ? 'xlookup' : 'auto', ledger: 'always', regenerate: false };
    if (isSeb) fixed.workLog = null;
    else if (p.wlOverride && typeof p.wlOverride === 'object') fixed.workLog = p.wlOverride;
    if (p.statusAtEnqueue !== undefined && p.statusAtEnqueue !== null && String(p.statusAtEnqueue) !== 'approved') {
      fixed.skipPdfReason = 'not_approved_at_enqueue';
    }
    return vpsRenderPdf_(vpsWith_(p, fixed));
  } },
  // 作り直し: secondary＝いつも直接方式・台帳はいつも / edit＝PDF_MODE・台帳あり / admin＝PDF_MODE・台帳なし
  'internal.regenerate': { fn: function(p) {
    var reason = String(p.reason || '');
    if (reason === 'secondary') return vpsRenderPdf_(vpsWith_(p, { mode: 'direct', regenerate: true, ledger: 'always' }));
    if (reason === 'edit') return vpsRenderPdf_(vpsWith_(p, { mode: 'auto', regenerate: true, ledger: 'always' }));
    if (reason === 'admin') return vpsRenderPdf_(vpsWith_(p, { mode: 'auto', regenerate: true, ledger: 'none' }));
    return vpsFail_('bad_request', 'reason must be secondary|edit|admin', false);
  } },
  'internal.afterDelete': { fn: function(p) { return vpsAfterDelete_(p); } },
  // 仕様 7-2: PDF の無い二次承認の単票など。vps.upsertLedger と同じ
  'internal.ledgerUpsert': { fn: function(p) { return vpsUpsertLedger_(p); } },
  // 朝バッチの 1 件: シートに pdfFileId があれば skipped。PDF の例外も ok:true で返す（朝バッチが数える）。台帳は書かない
  'internal.batchPdf': { fn: function(p) {
    return vpsRenderPdf_(vpsWith_(p, { mode: 'auto', ledger: 'none', regenerate: false, skipIfHasPdf: true, pdfErrorIsOk: true }));
  } },
  'internal.ledgerRewriteDay': { fn: function(p) { return vpsLedgerRewriteDay_(p); } },
  'internal.ledgerRebuild': { fn: function(p) { return vpsLedgerRebuild_(p); } },
};

function vpsWith_(p, fixed) {
  var out = {};
  Object.keys(p).forEach(function(k) { out[k] = p[k]; });
  Object.keys(fixed).forEach(function(k) { out[k] = fixed[k]; });
  return out;
}

// ===== ping =====

// 仕様 7-2 の名前（scriptId・spreadsheetId・spreadsheetTimeZone・migrated・forwardUrlSet・accumulationKeyRows・
// apiVersion）と、最初の版の名前（dbId・dbTimeZone・migrationMode・version）の両方を返す。
var INTERNAL_API_VERSION = 1;
function vpsPing_() {
  var out = {
    ok: true,
    pong: new Date().toISOString(),
    apiVersion: INTERNAL_API_VERSION,
    version: VPS_INTERNAL_API_VERSION_,
    scriptId: '',
    migrated: isMigrationMode_(),
    migrationMode: isMigrationMode_(),
    migratedTo: getMigratedToUrl_(),
    forwardUrlSet: false,
    scriptTimeZone: '',
  };
  try { out.scriptId = ScriptApp.getScriptId(); } catch (e0) { out.scriptIdError = String(e0.message || e0); }
  try { out.forwardUrlSet = !!PropertiesService.getScriptProperties().getProperty(SEBASTIAN_FORWARD_URL_KEY); } catch (e1) { /* ignore */ }
  try { out.scriptTimeZone = Session.getScriptTimeZone(); } catch (e) { /* ignore */ }
  try {
    var ss = getDb_();
    out.spreadsheetId = out.dbId = ss.getId();
    try { out.spreadsheetTimeZone = out.dbTimeZone = ss.getSpreadsheetTimeZone(); } catch (e2) { /* ignore */ }
    var s = getSettings_();
    out.settings = {
      TEMPLATE_SSID: normalize_(s['TEMPLATE_SSID']),
      PDF_ROOT_FOLDER_ID: normalize_(s['PDF_ROOT_FOLDER_ID']),
      ACCUMULATION_SS_ID: normalize_(s['ACCUMULATION_SS_ID']),
      PDF_MODE: String(s['PDF_MODE'] == null ? '' : s['PDF_MODE']),
      SOURCE_SSID: normalize_(s['SOURCE_SSID']),
      pdfModeIsDirect: normalize_(s['PDF_MODE']).indexOf('direct') >= 0,
    };
    // Settings に ACCUMULATION_SS_ID の行がいくつあるか（2 行以上なら下の行が使われる。gap-5 中-2）
    var vals = ss.getSheetByName(SHEET.SETTINGS).getDataRange().getValues();
    var n = 0;
    for (var r = 2; r < vals.length; r++) if (String(vals[r][0] || '').trim() === 'ACCUMULATION_SS_ID') n++;
    out.accumulationKeyRows = n;
  } catch (e3) {
    out.dbError = String(e3.message || e3);
  }
  return out;
}

// ===== expect（シートの行がこの値になっているはず）の確かめ =====

var VPS_EXPECT_ALIASES_ = {
  status: 'status(submitted/approved/canceled)',
  requestType: 'requestType(overtime/holiday)',
};
var VPS_WORKLOG_FIELDS_ = ['actualStartAt', 'actualEndAt', 'actualMinutes', 'breakMinutes', 'netMinutes'];

function vpsCellText_(v) {
  if (v instanceof Date) return fmtDate_(v, 'yyyy-MM-dd HH:mm:ss');
  return normalize_(v);
}

// 合わない項目の一覧を返す（空なら合っている）
function vpsCheckExpect_(requestId, expect) {
  if (!expect || typeof expect !== 'object') return [];
  var keys = Object.keys(expect);
  if (!keys.length) return [];
  var info = getSheetHeaderIndex_('Requests', 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  var row = null;
  if (lastRow >= 2) {
    var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < values.length; i++) {
      if (normalize_(values[i][idx['requestId']]) === requestId) { row = values[i]; break; }
    }
  }
  if (!row) return ['requestId'];
  var wl = null;
  var bad = [];
  keys.forEach(function(k) {
    var want = expect[k] == null ? '' : normalize_(expect[k]);
    var got;
    if (VPS_WORKLOG_FIELDS_.indexOf(k) >= 0) {
      if (!wl) wl = buildWorkLogsMapByRequestId_().get(requestId) || {};
      got = normalize_(wl[k] == null ? '' : wl[k]);
      if (k !== 'actualStartAt' && k !== 'actualEndAt' && got === '0' && want === '') got = '';
    } else {
      var col = idx[VPS_EXPECT_ALIASES_[k] || k];
      if (col === undefined) { bad.push(k); return; }
      got = vpsCellText_(row[col]);
      // 日時は秒まで、または分までの書き方を許す
      if (got !== want && /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(got) && got.slice(0, want.length) === want) got = want;
    }
    if (got !== want) bad.push(k);
  });
  return bad;
}

// ===== 台帳の守り（計画 4-7） =====

// 台帳が開けるか、書く年度タブの見出しが ACC_HEADERS_ と同じかを、今の関数を呼ぶ前に確かめる。
// 開けないときに今の getAccumulationSS_ は新しい台帳を作ってしまうので、ここで止める。
// sheetTypes: ['残業'] など。タブが無いときは今の関数が作る（今と同じ）。
function vpsLedgerPrecheck_(dateObj, sheetTypes) {
  var ssId = normalize_(getSettings_()['ACCUMULATION_SS_ID']);
  if (!ssId) return { status: 'error', error: 'ledger_unavailable', message: 'ACCUMULATION_SS_ID is not set' };
  var ss;
  try {
    ss = SpreadsheetApp.openById(ssId);
  } catch (e) {
    return { status: 'error', error: 'ledger_unavailable', message: String(e.message || e) };
  }
  var fy = computeAccFiscalYear_(dateObj);
  for (var i = 0; i < sheetTypes.length; i++) {
    var name = sheetTypes[i] + '_' + fy + '年度';
    var sh = ss.getSheetByName(name);
    if (!sh) continue;
    var width = ACC_HEADERS_.length;
    var head = sh.getRange(1, 1, 1, width).getValues()[0];
    for (var c = 0; c < width; c++) {
      var want = String(ACC_HEADERS_[c]);
      if (String(head[c] == null ? '' : head[c]) !== want) {
        return { status: 'error', error: 'ledger_header_mismatch', sheetName: name,
          message: name + ' の ' + (c + 1) + ' 列目の見出しが違います' };
      }
    }
  }
  return null;
}

function vpsUpsertLedgerCore_(requestId, req) {
  req = req || getRequestById_(requestId);
  if (!req) return { status: 'skipped', reason: 'request_not_found' };
  var d = req.targetDate instanceof Date ? req.targetDate : new Date(req.targetDate);
  var sheetType = req.requestType === 'overtime' ? '残業' : '休日出勤';
  var pre = vpsLedgerPrecheck_(d, [sheetType]);
  if (pre) return pre;
  // 今の関数は失敗をログに出して握る（今と同じ）。ここでは呼んだことだけを返す。
  upsertAccumulationRow_(requestId);
  return { status: 'upserted', sheetName: sheetType + '_' + computeAccFiscalYear_(d) + '年度' };
}

// ===== vps.renderPdf =====
// mode: auto（PDF_MODE に "direct" を含めば直接方式、ほかは XLOOKUP 方式。今の画面の経路と同じ判定）
//       direct（二次承認の作り直しと同じ。いつも generatePdfDirect_）
//       xlookup（Sebastian の完了と同じ。いつも generatePdfForRequest_）
//       secondary（direct ＋ regenerate の別名）
// regenerate / oldPdfFileId: 前の PDF をゴミ箱へ移し、Requests の PDF の 3 列を空にしてから作る
//       （clearExistingPdfForRegeneration_ と同じ。順序「捨ててから作る」も同じ）。
//       シートの pdfFileId が oldPdfFileId と違う値（＝この仕事の前の試みで作った PDF）なら、捨てずに already を返す。
// ledger: none（既定）| always（PDF の成否に関係なく upsert）| onSuccess（created か already のときだけ upsert）
// requireWorkDone: true なら WorkLogs に actualEndAt が無いとき作らない（一次承認の経路）
// workLog: 今の画面の完了・手動入力が渡す workLogOverride と同じ形（任意）
function vpsRenderPdf_(p) {
  var requestId = normalize_(p.requestId);
  if (!requestId) return vpsFail_('bad_request', 'requestId is required', false);
  var mode = String(p.mode || 'auto');
  if (['auto', 'direct', 'xlookup', 'secondary'].indexOf(mode) < 0) {
    return vpsFail_('bad_request', 'mode must be auto|direct|xlookup|secondary', false);
  }
  var ledgerMode = String(p.ledger || 'none');
  if (['none', 'always', 'onSuccess'].indexOf(ledgerMode) < 0) {
    return vpsFail_('bad_request', 'ledger must be none|always|onSuccess', false);
  }
  var oldPdfFileId = normalize_(p.oldPdfFileId);
  var regenerate = p.regenerate === true || mode === 'secondary' || !!oldPdfFileId;

  var mismatch = vpsCheckExpect_(requestId, p.expect);
  if (mismatch.length) {
    return vpsFail_('row_not_synced', 'sheet row does not match expect: ' + mismatch.join(','), true, { requestId: requestId });
  }

  var req = getRequestById_(requestId);
  if (!req) return vpsFail_('row_not_synced', 'request not found in sheet: ' + requestId, true, { requestId: requestId });

  var out = {
    ok: true, requestId: requestId, dept: req.dept, workerName: normalize_(req.workerName),
    pdfFileId: '', pdfFolderId: '', pdfGeneratedAt: '', pdfName: '', error: null,
    pdf: { status: 'skipped' }, ledger: { status: 'skipped' }, trash: { status: 'skipped' },
  };

  // 1) 前の PDF を捨てる（作り直しのとき）
  if (regenerate) {
    out.trash = vpsTrashOldPdf_(req, oldPdfFileId);
    _headerCache = {};
    req = getRequestById_(requestId);
  }

  // 2) PDF を作る（今の関数をそのまま）
  var useDirect;
  if (mode === 'direct' || mode === 'secondary') useDirect = true;
  else if (mode === 'xlookup') useDirect = false;
  else useDirect = normalize_(getSettings_()['PDF_MODE']).indexOf('direct') >= 0;
  out.pdf.method = useDirect ? 'direct' : 'xlookup';

  var canMake = true;
  if (p.skipPdfReason) {
    canMake = false;
    out.pdf = { status: 'skipped', reason: String(p.skipPdfReason), method: out.pdf.method };
  } else if (p.skipIfHasPdf === true && normalize_(req.pdfFileId)) {
    canMake = false;
    out.pdf = { status: 'skipped', reason: 'has_pdf', method: out.pdf.method };
  } else if (req.status !== 'approved') {
    canMake = false;
    out.pdf = { status: 'skipped', reason: 'not_approved', method: out.pdf.method };
  } else if (p.requireWorkDone === true) {
    var wlNow = buildWorkLogsMapByRequestId_().get(requestId) || {};
    if (!wlNow.actualEndAt) {
      canMake = false;
      out.pdf = { status: 'skipped', reason: 'work_not_done', method: out.pdf.method };
    }
  }

  if (canMake) {
    var wlOverride = (p.workLog && typeof p.workLog === 'object') ? p.workLog : undefined;
    try {
      var r = useDirect ? generatePdfDirect_(requestId, wlOverride) : generatePdfForRequest_(requestId, wlOverride);
      if (r && r.already) {
        out.pdf = { status: 'already', method: out.pdf.method, pdfFileId: normalize_(r.pdfFileId) };
      } else {
        out.pdf = { status: 'created', method: out.pdf.method, pdfFileId: r ? r.pdfFileId : '',
          pdfName: r ? r.pdfName : '', folderId: r ? r.folderId : '' };
      }
    } catch (e) {
      out.pdf = { status: 'error', method: out.pdf.method, error: String((e && e.message) || e) };
    }
  }

  // 3) 台帳（今の経路と同じ条件。PDF の後）
  var pdfOk = out.pdf.status === 'created' || out.pdf.status === 'already';
  if (ledgerMode === 'always' || (ledgerMode === 'onSuccess' && pdfOk)) {
    try {
      out.ledger = vpsUpsertLedgerCore_(requestId, null);
    } catch (e2) {
      out.ledger = { status: 'error', error: 'internal', message: String((e2 && e2.message) || e2) };
    }
  }

  // 4) シートに書かれた PDF の 3 列を返す（VPS は返事を受けたら先に DB へ記録する）
  _headerCache = {};
  var after = getRequestById_(requestId) || {};
  out.pdfFileId = normalize_(after.pdfFileId);
  out.pdfFolderId = normalize_(after.pdfFolderId);
  out.pdfGeneratedAt = vpsFmtDateTime_(after.pdfGeneratedAt);
  out.pdfGeneratedAtEpochMs = vpsEpochMs_(after.pdfGeneratedAt);
  out.pdfName = out.pdf.pdfName || '';
  out.pdf.pdfFileId = out.pdfFileId;
  out.pdf.folderId = out.pdfFolderId;
  out.pdf.pdfGeneratedAt = out.pdfGeneratedAt;
  out.pdf.pdfGeneratedAtEpochMs = out.pdfGeneratedAtEpochMs;

  if (out.pdf.status === 'error' && p.pdfErrorIsOk !== true) {
    out.ok = false;
    out.error = 'pdf_failed';
    out.message = out.pdf.error;
    out.retryable = true;
  }
  return out;
}

// 前の PDF を捨てる。今の clearExistingPdfForRegeneration_ を使う（ゴミ箱へ移せなくても記録して続ける）。
function vpsTrashOldPdf_(req, oldPdfFileId) {
  var info = getSheetHeaderIndex_('Requests', 1);
  var row = info.sh.getRange(req.rowNo, 1, 1, info.sh.getLastColumn()).getValues()[0];
  var cur = info.idx['pdfFileId'] !== undefined ? normalize_(row[info.idx['pdfFileId']]) : '';
  if (cur && oldPdfFileId && cur !== oldPdfFileId) {
    // この仕事の前の試みで作った新しい PDF。捨てない（作る段で already になる）
    return { status: 'skipped', reason: 'sheet_has_newer_pdf', fileId: oldPdfFileId, currentPdfFileId: cur };
  }
  var out = { status: 'skipped', fileId: oldPdfFileId || cur };
  if (cur) {
    // シートに前の PDF が残っている: 今と同じくゴミ箱へ移して 3 列を空にする
    var trashedOk = vpsTryTrash_(cur);
    clearExistingPdfForRegeneration_(info.sh, info.idx, vpsRowWithoutPdfId_(row, info.idx), req.rowNo);
    SpreadsheetApp.flush();
    out.status = trashedOk === true ? 'trashed' : 'error';
    if (trashedOk !== true) out.error = trashedOk;
  } else if (oldPdfFileId) {
    // VPS が先にシートの 3 列を空にした: ゴミ箱へ移すだけ
    var t = vpsTryTrash_(oldPdfFileId);
    out.status = t === true ? 'trashed' : 'error';
    if (t !== true) out.error = t;
  }
  return out;
}

// clearExistingPdfForRegeneration_ に渡す行から pdfFileId を外す（ゴミ箱へ移すのは vpsTryTrash_ で済ませ、
// 結果を返事に入れるため。3 列を空にする処理は今の関数のまま）。
function vpsRowWithoutPdfId_(row, idx) {
  var copy = row.slice();
  if (idx['pdfFileId'] !== undefined) copy[idx['pdfFileId']] = '';
  return copy;
}

function vpsTryTrash_(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return true;
  } catch (e) {
    Logger.log('旧PDF削除スキップ: ' + e.message);
    return String((e && e.message) || e);
  }
}

// ===== vps.upsertLedger =====
function vpsUpsertLedger_(p) {
  var requestId = normalize_(p.requestId);
  if (!requestId) return vpsFail_('bad_request', 'requestId is required', false);
  var mismatch = vpsCheckExpect_(requestId, p.expect);
  if (mismatch.length) return vpsFail_('row_not_synced', 'sheet row does not match expect: ' + mismatch.join(','), true);
  var req = getRequestById_(requestId);
  if (!req) return vpsFail_('row_not_synced', 'request not found in sheet: ' + requestId, true);
  var ledger = vpsUpsertLedgerCore_(requestId, req);
  if (ledger.status === 'error') return vpsFail_(ledger.error, ledger.message, ledger.error === 'ledger_unavailable', { ledger: ledger });
  return { ok: true, requestId: requestId, ledger: ledger };
}

// ===== vps.removeLedgerRow =====
// 今の api_deleteRequest と同じく removeAccumulationRow_(requestId, requestType, targetDate) を呼ぶ。
// 行はもう取消済み・シートから読めないこともあるので、種別と対象日は本文で受ける。
function vpsRemoveLedgerRow_(p) {
  var requestId = normalize_(p.requestId);
  var requestType = normalize_(p.requestType);
  var d = vpsParseYmd_(p.targetDate);
  if (!requestId) return vpsFail_('bad_request', 'requestId is required', false);
  if (requestType !== 'overtime' && requestType !== 'holiday') return vpsFail_('bad_request', 'requestType must be overtime|holiday', false);
  if (!d) return vpsFail_('bad_request', 'targetDate must be yyyy-MM-dd', false);
  var sheetType = requestType === 'overtime' ? '残業' : '休日出勤';
  var pre = vpsLedgerPrecheck_(d, [sheetType]);
  if (pre) return vpsFail_(pre.error, pre.message, pre.error === 'ledger_unavailable', { ledger: pre });
  removeAccumulationRow_(requestId, requestType, d);
  return { ok: true, requestId: requestId, ledger: { status: 'removed', sheetName: sheetType + '_' + computeAccFiscalYear_(d) + '年度' } };
}

// ===== internal.afterDelete =====
// 削除: PDF を作らない。前の PDF をゴミ箱へ（シートに残っていれば 3 列も空に）→ 台帳の行を消す。
function vpsAfterDelete_(p) {
  var requestId = normalize_(p.requestId);
  if (!requestId) return vpsFail_('bad_request', 'requestId is required', false);
  var out = { ok: true, requestId: requestId, trash: { status: 'skipped' }, ledger: { status: 'skipped' } };
  var oldPdfFileId = normalize_(p.oldPdfFileId);
  var req = getRequestById_(requestId);
  if (req && (oldPdfFileId || normalize_(req.pdfFileId))) {
    out.trash = vpsTrashOldPdf_(req, oldPdfFileId);
  } else if (oldPdfFileId) {
    var t = vpsTryTrash_(oldPdfFileId);
    out.trash = t === true ? { status: 'trashed', fileId: oldPdfFileId } : { status: 'error', fileId: oldPdfFileId, error: t };
  }
  var requestType = normalize_(p.requestType) || (req ? req.requestType : '');
  var d = vpsParseYmd_(p.targetDate) || (req ? (req.targetDate instanceof Date ? req.targetDate : new Date(req.targetDate)) : null);
  if (!d || isNaN(d.getTime()) || (requestType !== 'overtime' && requestType !== 'holiday')) {
    return vpsFail_('bad_request', 'requestType and targetDate are required when the row is not in the sheet', false);
  }
  var sheetType = requestType === 'overtime' ? '残業' : '休日出勤';
  var pre = vpsLedgerPrecheck_(d, [sheetType]);
  if (pre) {
    out.ledger = pre;
    return vpsFail_(pre.error, pre.message, pre.error === 'ledger_unavailable', { trash: out.trash, ledger: pre });
  }
  removeAccumulationRow_(requestId, requestType, d);
  out.ledger = { status: 'removed', sheetName: sheetType + '_' + computeAccFiscalYear_(d) + '年度' };
  return out;
}

// ===== vps.ledgerRewriteDay / internal.ledgerRewriteDay =====
// 朝バッチの最後と同じ appendToAccumulationSS_(対象日)（前日分の書き直し）。
function vpsLedgerRewriteDay_(p) {
  var d = vpsParseYmd_(p.targetDate);
  if (!d) return vpsFail_('bad_request', 'targetDate must be yyyy-MM-dd', false);
  var pre = vpsLedgerPrecheck_(d, ['残業', '休日出勤']);
  if (pre) return vpsFail_(pre.error, pre.message, pre.error === 'ledger_unavailable', { ledger: pre });
  appendToAccumulationSS_(d);
  return { ok: true, targetDate: fmtDate_(d, 'yyyy-MM-dd'), ledger: { status: 'rewritten' } };
}

// ===== vps.morningBatch =====
// 今の morningBatchNoMail_ と同じ 3 段を、指定の対象日で動かす:
//   batchGeneratePdfs_(対象日) → logBatchResult_('morningBatchNoMail', 対象日, 結果) → appendToAccumulationSS_(対象日)
// 対象日を省くと今と同じ「昨日」。件数が多いと 6 分の上限に近づくので、VPS は internal.batchPdf を 1 件ずつ
// 呼ぶ形（計画 4-9）も選べる。
function vpsMorningBatch_(p) {
  var d = p.targetDate ? vpsParseYmd_(p.targetDate) : vpsYesterday_();
  if (!d) return vpsFail_('bad_request', 'targetDate must be yyyy-MM-dd', false);
  var pdfResult = batchGeneratePdfs_(d);
  Logger.log('vps.morningBatch PDF: ok=' + pdfResult.ok + ' skip=' + pdfResult.skip + ' fail=' + pdfResult.fail);
  logBatchResult_('morningBatchNoMail', d, pdfResult);
  var ledger;
  var pre = vpsLedgerPrecheck_(d, ['残業', '休日出勤']);
  if (pre) {
    ledger = pre;
  } else {
    try {
      appendToAccumulationSS_(d);
      ledger = { status: 'rewritten' };
    } catch (e) {
      Logger.log('蓄積SSエラー: ' + e.message);
      ledger = { status: 'error', error: 'internal', message: String(e.message || e) };
    }
  }
  return {
    ok: true,
    targetDate: fmtDate_(d, 'yyyy-MM-dd'),
    counts: { ok: pdfResult.ok || 0, skip: pdfResult.skip || 0, fail: pdfResult.fail || 0 },
    errors: pdfResult.errors || [],
    batchLog: { status: 'appended', batchName: 'morningBatchNoMail' },
    ledger: ledger,
  };
}

// ===== vps.ledgerRebuild / internal.ledgerRebuild =====
// 今の rebuildAccumulationSS()（年度タブを _backup_<日時> へ名前を変えてから全件を書き直す）をそのまま呼ぶ。
// 台帳を開けなければ呼ばない（今の getAccumulationSS_ が新しい台帳を作るのを防ぐ）。6 分を超えれば GAS が止め、
// VPS 側は時間切れ（結果が分からない）として扱う。
function vpsLedgerRebuild_(p) {
  var pre = vpsLedgerPrecheck_(new Date(), []);
  if (pre) return vpsFail_(pre.error, pre.message, pre.error === 'ledger_unavailable', { ledger: pre });
  var r = rebuildAccumulationSS();
  if (!r || r.rebuilt !== true) {
    return vpsFail_('internal', 'rebuildAccumulationSS did not finish: ' + ((r && r.reason) || 'unknown'), true,
      { ledger: { status: 'error', error: (r && r.reason) || 'unknown' } });
  }
  return {
    ok: true,
    ledger: { status: 'rebuilt' },
    counts: { processedRows: r.processed, sheets: r.sheets, writtenRows: r.written, backups: r.backups.length },
    backups: r.backups,
    rebuildMs: r.durationMs,
  };
}
