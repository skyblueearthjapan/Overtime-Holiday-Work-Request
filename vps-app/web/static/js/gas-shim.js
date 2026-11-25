/*
 * gas-shim.js — GAS の google.script.run の代わり（VPS 版。仕様 3-4〜3-6）。
 *
 * 画面の呼び出し行は GAS 版のまま動く:
 *   google.script.run.withSuccessHandler(f).withFailureHandler(g).api_xxx(a, b)
 * → POST /api/rpc/api_xxx  本文 ["a","b"]（省略した引数は送らない）
 *
 * 返事:
 *   200 {ok:true,result}   → f(result)（result が {error}・{ok:false}・null でもそのまま成功に渡す）
 *   400 など {ok:false,error} → g(Error(error))（err.message と String(err) の両方で文が出る）
 *   409（同じ操作を処理中）→ 1 秒おきに最大 10 回、同じ冪等キーで送り直す
 *   401（ログイン切れ）    → g は呼ばない。書きかけを sessionStorage に退避し、ボタンを戻し、
 *                            beforeunload を外して、ログインへ移る（罠 20）
 *   5xx・通信エラー        → g(Error('サーバーでエラーが起きました。…'))。冪等キーは残す
 *
 * 冪等キー（罠 17）: 書き込みの関数だけ、「関数名＋引数」ごとに 1 つ。捨てるのは 200・4xx（401・409 以外）のときだけ。
 * 同じ関数・同じ引数の書き込みが返事待ちの間にもう一度来たら送らず、前の返事を両方のハンドラに渡す。
 *
 * CSP（script-src 'self'）のため、HTML の on 属性は data-on-<event>="関数名" と data-args='[JSON]' に置き換えてあり、
 * このファイルの受け口（document への委譲）が呼ぶ。引数の "$this"・"$event"・"$value" は要素・イベント・要素の値に置き換える。
 */
(function (global) {
  'use strict';

  var doc = global.document;

  // ---- 仕様 4-3 の 29 関数と 4-2 の 9 関数（閉じたリスト。ここに無い名前は google.script.run に生えない） ----
  var READ_FNS = [
    'api_getDeptList', 'api_getWorkersByDept', 'api_getWorkerInfo', 'api_getOrderList',
    'api_getHolidayCandidateDates', 'api_getDashboard', 'api_getApproverDepts', 'api_getApproverDashboard',
    'api_approverMonthlySummary', 'api_getSecondaryApprovalList', 'api_adminDeptOptions',
    'api_adminMonthlySummary', 'api_adminFiscalTrend', 'api_adminDailyDetail', 'api_adminSpecialClauseCount',
    'api_adminPendingWatch', 'api_getClosingSummary',
    'ops_getStatus'
  ];
  var WRITE_FNS = [
    'api_submitRequest', 'api_markOvertimeDone', 'api_markHolidayStart', 'api_markHolidayDone',
    'api_editRequest', 'api_deleteRequest', 'api_approveRequest', 'api_rejectRequest',
    'api_approveRequestsBatch', 'api_secondaryApprove', 'api_secondaryApproveBatch', 'api_adminManualComplete',
    'ops_retryJob', 'ops_syncSheetConfigNow', 'ops_runMasterSync', 'ops_regeneratePdf',
    'ops_batchPdfForDate', 'ops_rebuildLedger', 'ops_recalcBreaks', 'ops_revokeSessions',
    'ops_runSheetIntake'
  ];
  var ALL_FNS = READ_FNS.concat(WRITE_FNS);
  var IS_WRITE = {};
  WRITE_FNS.forEach(function (n) { IS_WRITE[n] = true; });

  var INTERNAL_MSG = 'サーバーでエラーが起きました。時間をおいてもう一度お試しください。';
  var LOGIN_MSG = 'ログインが切れました。ログイン後にもう一度押してください。';
  var RETRY_409_MAX = 10;
  var DRAFT_MAX_AGE_MS = 24 * 60 * 60 * 1000;

  function retryDelayMs() {
    return typeof global.__VPS_SHIM_RETRY_MS === 'number' ? global.__VPS_SHIM_RETRY_MS : 1000;
  }

  // ハンドラの中の例外は、GAS と同じくシムの外へ投げる（ほかのハンドラは止めない）
  function report(e) {
    global.setTimeout(function () { throw e; }, 0);
  }

  // GAS の失敗ハンドラに渡る Error と同じ使い方ができる形（err.message・String(err) の両方で文が出る）
  function gasError(msg) {
    var e = new Error(msg);
    e.toString = function () { return msg; };
    return e;
  }

  function newKey() {
    var c = global.crypto;
    if (c && typeof c.randomUUID === 'function') { return c.randomUUID(); }
    var b = new Uint8Array(16);
    if (c && typeof c.getRandomValues === 'function') { c.getRandomValues(b); }
    else { for (var i = 0; i < 16; i++) { b[i] = Math.floor(Math.random() * 256); } }
    var s = '';
    for (var j = 0; j < 16; j++) { s += (b[j] < 16 ? '0' : '') + b[j].toString(16); }
    return s;
  }

  // ---- 呼び出し ----
  var opKeys = {};     // 書き込みの「関数名＋引数」→ 冪等キー
  var inflight = {};   // 書き込みの「関数名＋引数」→ 返事待ちの呼び出し（waiter の配列）
  var currentAction = null;  // 受け口が呼んでいる最中の操作（401 でボタンを戻すため）

  function deliver(w, kind, value) {
    var h = kind === 'ok' ? w.opts.ok : w.opts.fail;
    if (typeof h !== 'function') { return; }
    try {
      if (w.opts.hasUser) { h(value, w.opts.user); } else { h(value); }
    } catch (e) { report(e); }
  }

  function takeWaiters(sig, own) {
    if (!sig) { return own; }
    var ws = inflight[sig] || [];
    delete inflight[sig];
    return ws;
  }

  function settle(sig, own, kind, value) {
    takeWaiters(sig, own).forEach(function (w) { deliver(w, kind, value); });
  }

  function errorText(data) {
    return data && typeof data.error === 'string' && data.error ? data.error : '';
  }

  function send(name, body, sig, attempt, own) {
    var headers = { 'Content-Type': 'application/json' };
    if (sig) { headers['X-Idempotency-Key'] = opKeys[sig]; }
    var p;
    try {
      p = global.fetch('/api/rpc/' + encodeURIComponent(name), {
        method: 'POST', headers: headers, body: body, credentials: 'same-origin'
      });
    } catch (e) {
      p = Promise.reject(e);
    }
    var handled = false;
    p.then(function (res) {
      return res.text().then(function (t) {
        var data = null;
        try { data = t ? JSON.parse(t) : null; } catch (e) { data = null; }
        return { status: res.status, data: data };
      }, function () { return { status: res.status, data: null }; });
    }).then(function (r) {
      handled = true;
      var st = r.status, data = r.data;
      if (st === 401) {
        // 失敗ハンドラは呼ばない（今の画面は失敗で alert と読み直しをするので二重に出る）。冪等キーは捨てない。
        var ws = takeWaiters(sig, own);
        goToLogin(ws.map(function (w) { return w.ctx; }));
        return;
      }
      if (st === 409) {
        if (attempt < RETRY_409_MAX) {
          global.setTimeout(function () { send(name, body, sig, attempt + 1, own); }, retryDelayMs());
          return;
        }
        settle(sig, own, 'fail', gasError(errorText(data) || INTERNAL_MSG));
        return;
      }
      if (st >= 500 || st < 200) {
        settle(sig, own, 'fail', gasError(INTERNAL_MSG));
        return;
      }
      if (st === 200 && data && data.ok === true) {
        if (sig) { delete opKeys[sig]; }
        settle(sig, own, 'ok', data.result === undefined ? null : data.result);
        return;
      }
      if (st >= 400) {
        // 業務エラー（400）・権限（403）など。サーバーは答えたので鍵を捨てる。
        if (sig) { delete opKeys[sig]; }
        settle(sig, own, 'fail', gasError(errorText(data) || INTERNAL_MSG));
        return;
      }
      settle(sig, own, 'fail', gasError(INTERNAL_MSG));
    }).then(null, function (e) {
      if (handled) { report(e); return; }  // 返事は処理済み（ここに来るのはシムの中の例外だけ）
      settle(sig, own, 'fail', gasError(INTERNAL_MSG));
    });
  }

  function call(name, args, opts) {
    var waiter = { opts: opts, ctx: currentAction };
    var body = JSON.stringify(args);
    if (!IS_WRITE[name]) {
      send(name, body, null, 0, [waiter]);
      return;
    }
    var sig = name + '\u0000' + body;
    if (inflight[sig]) {  // 二度押し: 送らずに前の返事を待つ
      inflight[sig].push(waiter);
      return;
    }
    inflight[sig] = [waiter];
    if (!opKeys[sig]) { opKeys[sig] = newKey(); }
    send(name, body, sig, 0, null);
  }

  function makeRunner(opts) {
    var r = {};
    function withOpt(k, v) {
      var o = { ok: opts.ok, fail: opts.fail, user: opts.user, hasUser: opts.hasUser };
      o[k] = v;
      if (k === 'user') { o.hasUser = true; }
      return makeRunner(o);
    }
    r.withSuccessHandler = function (f) { return withOpt('ok', f); };
    r.withFailureHandler = function (f) { return withOpt('fail', f); };
    r.withUserObject = function (u) { return withOpt('user', u); };
    ALL_FNS.forEach(function (name) {
      r[name] = function () { call(name, Array.prototype.slice.call(arguments), opts); };
    });
    return r;
  }

  global.google = global.google || {};
  global.google.script = global.google.script || {};
  global.google.script.run = makeRunner({ ok: null, fail: null, user: undefined, hasUser: false });

  // ---- ページの名前と、テンプレートから渡された値 ----
  function templateData() {
    try {
      var el = doc.getElementById('gas-template-data');
      return el ? JSON.parse(el.textContent || '{}') : {};
    } catch (e) { return {}; }
  }
  function pageName() {
    var t = templateData();
    if (t && t.PAGE) { return String(t.PAGE); }
    var m = /[?&]page=([^&#]*)/.exec(global.location.search || '');
    return m ? decodeURIComponent(m[1]) : 'top';
  }

  // ---- 書きかけの退避と復元（仕様 3-5） ----
  var adapters = [];
  var currentEmail = '';
  var goingToLogin = false;

  function draftKey() { return 'overtime-draft:' + pageName(); }

  function saveDraft() {
    try {
      var data = {}, any = false;
      adapters.forEach(function (a) {
        try {
          var d = a.collect();
          if (d !== null && d !== undefined) { data[a.name] = d; any = true; }
        } catch (e) { /* 退避できなくてもログインへは進む */ }
      });
      if (!any) { global.sessionStorage.removeItem(draftKey()); return; }
      global.sessionStorage.setItem(draftKey(), JSON.stringify({ savedAt: Date.now(), email: currentEmail, data: data }));
    } catch (e) { /* sessionStorage が使えなくても進む */ }
  }

  function restoreDraft() {
    var draft = null;
    try {
      draft = JSON.parse(global.sessionStorage.getItem(draftKey()) || 'null');
      global.sessionStorage.removeItem(draftKey());
    } catch (e) { return; }
    if (!draft || !draft.data || !currentEmail || draft.email !== currentEmail) { return; }
    if (!(Date.now() - Number(draft.savedAt) < DRAFT_MAX_AGE_MS)) { return; }
    adapters.forEach(function (a) {
      if (Object.prototype.hasOwnProperty.call(draft.data, a.name)) {
        try { a.restore(draft.data[a.name]); } catch (e) { report(e); }
      }
    });
  }

  // 画面の初期化が済むのを待ってから fn を呼ぶ（cond が真になるまで 200ms おき、既定 15 秒まで）
  function when(cond, fn, timeoutMs) {
    var limit = Date.now() + (timeoutMs || 15000);
    (function tick() {
      var ok = false;
      try { ok = !!cond(); } catch (e) { ok = false; }
      if (ok) { try { fn(); } catch (e) { report(e); } return; }
      if (Date.now() < limit) { global.setTimeout(tick, 200); }
    })();
  }

  global.__vpsDraft = {
    register: function (adapter) { adapters.push(adapter); },
    when: when,
    save: saveDraft
  };

  // ---- beforeunload（ログインへ移る前に外す） ----
  var unloadListeners = [];
  var origAdd = global.addEventListener;
  if (typeof origAdd === 'function') {
    global.addEventListener = function (type, fn, opt) {
      if (type === 'beforeunload') { unloadListeners.push([fn, opt]); }
      return origAdd.apply(this, arguments);
    };
  }
  function removeBeforeUnload() {
    try { global.onbeforeunload = null; } catch (e) { /* 無視 */ }
    unloadListeners.forEach(function (x) {
      try { global.removeEventListener('beforeunload', x[0], x[1]); } catch (e) { /* 無視 */ }
    });
    unloadListeners = [];
  }

  // ---- 401: ボタンを戻し、書きかけを退避してログインへ ----
  function restoreButtons(ctxs) {
    (ctxs || []).forEach(function (ctx) {
      if (!ctx || !ctx.snap) { return; }
      ctx.snap.forEach(function (s) {
        var b = s[0];
        if (!b || (doc.contains && !doc.contains(b))) { return; }
        if (b.disabled !== s[1]) { b.disabled = s[1]; }
        if (b.innerHTML !== s[2]) { b.innerHTML = s[2]; }
      });
    });
  }

  function goToLogin(ctxs) {
    restoreButtons(ctxs);
    if (goingToLogin) { return; }
    goingToLogin = true;
    saveDraft();
    removeBeforeUnload();
    try { global.alert(LOGIN_MSG); } catch (e) { /* 無視 */ }
    global.location.href = '/auth/login?next=' + encodeURIComponent(global.location.pathname + global.location.search);
  }

  // ログインが続いているかを先に確かめる（申請モーダルを開くとき・画面に戻ったとき）。first は読み込み直後の 1 回。
  function checkSession(first) {
    var p;
    try {
      p = global.fetch('/auth/me', { method: 'GET', credentials: 'same-origin', headers: { 'Accept': 'application/json' } });
    } catch (e) { return; }
    p.then(function (res) {
      if (res.status === 401) { goToLogin([]); return; }
      if (res.status !== 200) { return; }
      return res.json().then(function (me) {
        var email = (me && (me.email || (me.user && me.user.email))) || '';
        if (email) { currentEmail = String(email); }
        if (first) { restoreDraft(); }
      });
    }).then(null, function () { /* 通信できないときは何もしない（操作の時に分かる） */ });
  }
  global.__vpsCheckSession = function () { checkSession(false); };

  // ---- on 属性の代わりの受け口（仕様 3-6） ----
  function snapshotButtons() {
    var list = doc.querySelectorAll ? doc.querySelectorAll('button') : [];
    var out = [];
    for (var i = 0; i < list.length; i++) { out.push([list[i], list[i].disabled, list[i].innerHTML]); }
    return out;
  }

  function resolveArg(a, el, ev) {
    if (a === '$this') { return el; }
    if (a === '$event') { return ev; }
    if (a === '$value') { return el.value; }
    return a;
  }

  function invoke(el, ev, type) {
    var name = el.getAttribute('data-on-' + type);
    var fn = global[name];
    if (typeof fn !== 'function') { report(new Error('画面の関数が見つかりません: ' + name)); return; }
    var raw = el.getAttribute('data-args-' + type);
    if (raw === null || raw === undefined) { raw = el.getAttribute('data-args'); }
    var args = [];
    if (raw) {
      try { args = JSON.parse(raw); } catch (e) { report(e); return; }
    }
    args = args.map(function (a) { return resolveArg(a, el, ev); });
    var prev = currentAction;
    currentAction = { snap: type === 'click' ? snapshotButtons() : null };
    try {
      fn.apply(global, args);  // インラインの onclick と同じく、関数の this は window
    } catch (e) {
      report(e);
    } finally {
      currentAction = prev;
    }
  }

  function pathOf(target) {
    var out = [];
    var el = target;
    if (el && el.nodeType !== 1) { el = el.parentNode; }
    for (; el && el.nodeType === 1; el = el.parentNode) { out.push(el); }
    return out;
  }

  function onBubbling(ev) {
    var type = ev.type;
    // インラインのハンドラと同じく、イベントが起きた時点の祖先の並びで、内側から順に呼ぶ
    pathOf(ev.target).forEach(function (el) {
      if (!el.hasAttribute || !el.hasAttribute('data-on-' + type)) { return; }
      if (type === 'click' && el.disabled) { return; }
      invoke(el, ev, type);
    });
  }

  // mouseenter・mouseleave は委譲できないので、mouseover・mouseout と relatedTarget で同じ動きにする
  function onHover(ev, type) {
    var rel = ev.relatedTarget;
    pathOf(ev.target).forEach(function (el) {
      if (!el.hasAttribute || !el.hasAttribute('data-on-' + type)) { return; }
      if (rel && (rel === el || (el.contains && el.contains(rel)))) { return; }
      invoke(el, ev, type);
    });
  }

  // JS の文字列で作るボタンの data-args（HTML の属性の中に入れてよい形）
  global.__vpsArgs = function () {
    return JSON.stringify(Array.prototype.slice.call(arguments))
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  };

  if (doc && doc.addEventListener) {
    ['click', 'input', 'change'].forEach(function (t) { doc.addEventListener(t, onBubbling, false); });
    doc.addEventListener('mouseover', function (ev) { onHover(ev, 'mouseenter'); }, false);
    doc.addEventListener('mouseout', function (ev) { onHover(ev, 'mouseleave'); }, false);
    doc.addEventListener('visibilitychange', function () {
      if (doc.visibilityState === 'visible') { checkSession(false); }
    }, false);
    // 画面の JS（ページごとの .js）が書きかけの受け口を登録し終えてから確かめる
    if (doc.readyState === 'loading') {
      doc.addEventListener('DOMContentLoaded', function () { checkSession(true); }, false);
    } else {
      global.setTimeout(function () { checkSession(true); }, 0);
    }
  }

  // 試験用（node の shim.test.js が読む）
  global.__vpsShim = { READ_FNS: READ_FNS, WRITE_FNS: WRITE_FNS, INTERNAL_MSG: INTERNAL_MSG, LOGIN_MSG: LOGIN_MSG };
})(window);
