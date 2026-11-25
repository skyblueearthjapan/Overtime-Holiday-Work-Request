// gas-shim.js の試験（仕様 3-4〜3-6・12-2 の U）。外部の npm パッケージを使わない。
//   node --test vps-app/web/dev/
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { webcrypto } = require('crypto');

const SHIM = fs.readFileSync(path.join(__dirname, '..', 'static', 'js', 'gas-shim.js'), 'utf8');
const EMAIL = 'imaizumi@lineworks-local.info';

// ---- 偽物のブラウザ ----
function makeEl(tag, attrs, parent) {
  const el = {
    nodeType: 1, tagName: tag, attrs: Object.assign({}, attrs || {}), parentNode: parent || null,
    disabled: false, innerHTML: '', value: '', connected: true,
    getAttribute(k) { return Object.prototype.hasOwnProperty.call(this.attrs, k) ? this.attrs[k] : null; },
    hasAttribute(k) { return Object.prototype.hasOwnProperty.call(this.attrs, k); },
    contains(o) { for (let n = o; n; n = n.parentNode) { if (n === this) { return true; } } return false; }
  };
  return el;
}

function makeEnv(opts) {
  opts = opts || {};
  const listeners = {};
  const store = new Map(Object.entries(opts.session || {}));
  const env = {
    alerts: [], fetches: [], buttons: [], errors: [],
    me: opts.me === undefined ? { email: EMAIL } : opts.me,
    queue: [],        // RPC の返事（先頭から 1 つずつ使う）。無ければ hold（返事しない）
    held: []
  };
  const location = { pathname: '/', search: opts.search || '?page=top', href: '' };
  const document = {
    readyState: 'complete', visibilityState: 'visible',
    addEventListener(t, f) { (listeners[t] = listeners[t] || []).push(f); },
    getElementById(id) {
      if (id === 'gas-template-data') { return { textContent: JSON.stringify({ PAGE: opts.page || 'top' }) }; }
      return null;
    },
    querySelectorAll() { return env.buttons; },
    contains(el) { return el.connected !== false; }
  };
  function response(status, body) {
    const text = typeof body === 'string' ? body : JSON.stringify(body);
    return { status: status, text: () => Promise.resolve(text), json: () => Promise.resolve(JSON.parse(text)) };
  }
  const win = {
    document: document, location: location,
    alert: (m) => env.alerts.push(m),
    sessionStorage: {
      getItem: (k) => (store.has(k) ? store.get(k) : null),
      setItem: (k, v) => store.set(k, String(v)),
      removeItem: (k) => store.delete(k)
    },
    setTimeout: (f, ms) => setTimeout(() => { try { f(); } catch (e) { env.errors.push(e); } }, ms),
    crypto: webcrypto,
    addEventListener() {}, removeEventListener() {},
    fetch: (url, init) => {
      env.fetches.push({ url: url, init: init || {} });
      if (url === '/auth/me') {
        return Promise.resolve(env.me === 401 ? response(401, { ok: false }) : response(200, env.me));
      }
      const next = env.queue.shift();
      if (!next) { return new Promise((resolve) => env.held.push((s, b) => resolve(response(s, b)))); }
      if (next === 'network') { return Promise.reject(new TypeError('Failed to fetch')); }
      return Promise.resolve(response(next[0], next[1]));
    }
  };
  win.window = win;
  if (opts.retryMs !== undefined) { win.__VPS_SHIM_RETRY_MS = opts.retryMs; }
  vm.createContext(win);
  vm.runInContext(SHIM, win, { filename: 'gas-shim.js' });
  env.win = win;
  env.global = vm.runInContext('this', win);
  env.store = store;
  env.location = location;
  env.fire = (type, ev) => (listeners[type] || []).forEach((f) => f(ev));
  env.rpcFetches = () => env.fetches.filter((f) => f.url.startsWith('/api/rpc/'));
  return env;
}

const flush = (n) => new Promise((r) => setTimeout(r, n || 20));

// ---- 呼び出しの形 ----

test('ハンドラの付け方 4 通りが動き、呼ぶたびに別の runner になる', async () => {
  const env = makeEnv();
  const got = [];
  const run = env.win.google.script.run;
  env.queue.push([200, { ok: true, result: 1 }], [200, { ok: true, result: 2 }], [400, { ok: false, error: 'だめ' }], [200, { ok: true, result: 4 }]);
  run.withSuccessHandler((r) => got.push(['s1', r])).api_getDeptList();
  run.withFailureHandler((e) => got.push(['f2', e.message])).withSuccessHandler((r) => got.push(['s2', r])).api_getOrderList();
  run.withSuccessHandler((r) => got.push(['s3', r])).withFailureHandler((e) => got.push(['f3', e.message])).api_getDashboard();
  run.api_getDeptList();  // どちらも無し（何も起きない）
  await flush();
  assert.deepEqual(got.sort(), [['f3', 'だめ'], ['s1', 1], ['s2', 2]].sort());
  assert.equal(typeof run.withSuccessHandler(() => {}).api_getDeptList, 'function');
  assert.notEqual(run.withSuccessHandler(() => {}), run);
  assert.equal(run.api_getFormUrl, undefined, '閉じたリストに無い関数は生えない');
});

test('省略した引数は送らない・配列とオブジェクトはそのまま送る', async () => {
  const env = makeEnv();
  env.queue.push([200, { ok: true, result: {} }], [200, { ok: true, result: {} }], [200, { ok: true, result: {} }]);
  env.win.google.script.run.api_markOvertimeDone('id-1');
  env.win.google.script.run.api_markOvertimeDone('id-1', '20:00');
  env.win.google.script.run.api_getApproverDashboard(['製造1課', '品質管理課']);
  await flush();
  const f = env.rpcFetches();
  assert.equal(f[0].url, '/api/rpc/api_markOvertimeDone');
  assert.equal(f[0].init.body, '["id-1"]');
  assert.equal(f[1].init.body, '["id-1","20:00"]');
  assert.equal(f[2].init.body, '[["製造1課","品質管理課"]]');
  assert.equal(f[0].init.method, 'POST');
  assert.equal(f[0].init.credentials, 'same-origin');
  assert.equal(f[0].init.headers['Content-Type'], 'application/json');
});

test('200 の {error}・{ok:false}・null は成功ハンドラに届く', async () => {
  const env = makeEnv();
  const got = [];
  env.queue.push([200, { ok: true, result: { error: 'x', overtime: [] } }], [200, { ok: true, result: { ok: false, message: 'm' } }], [200, { ok: true, result: null }]);
  const r = env.win.google.script.run.withSuccessHandler((v) => got.push(v)).withFailureHandler(() => got.push('FAIL'));
  r.api_getDashboard();
  r.api_rejectRequest('a');
  r.api_getWorkerInfo();
  await flush();
  assert.deepEqual(JSON.parse(JSON.stringify(got)), [{ error: 'x', overtime: [] }, { ok: false, message: 'm' }, null]);
});

test('400 は失敗ハンドラに err.message と String(err) の両方で文が出る', async () => {
  const env = makeEnv();
  let err = null;
  env.queue.push([400, { ok: false, error: '既にキャンセル済みです。' }]);
  env.win.google.script.run.withFailureHandler((e) => { err = e; }).api_deleteRequest('a');
  await flush();
  assert.equal(err.message, '既にキャンセル済みです。');
  assert.equal(String(err), '既にキャンセル済みです。');
  assert.equal('エラー: ' + (err.message || err), 'エラー: 既にキャンセル済みです。');
});

test('5xx・通信エラーは決まった文で失敗ハンドラへ', async () => {
  const env = makeEnv();
  const msgs = [];
  env.queue.push([500, { ok: false, error: '内部の詳細' }], 'network', [502, '<html>bad gateway</html>']);
  const r = env.win.google.script.run.withFailureHandler((e) => msgs.push(e.message));
  r.api_getDeptList(); r.api_getOrderList(); r.api_getDashboard();
  await flush();
  const M = env.win.__vpsShim.INTERNAL_MSG;
  assert.equal(M, 'サーバーでエラーが起きました。時間をおいてもう一度お試しください。');
  assert.deepEqual(msgs, [M, M, M]);
});

// ---- 冪等キー ----

test('書き込みにだけ冪等キー。5xx の後は同じキー、200 の後は新しいキー', async () => {
  const env = makeEnv();
  const r = env.win.google.script.run.withFailureHandler(() => {}).withSuccessHandler(() => {});
  env.queue.push([200, { ok: true, result: [] }]);
  r.api_getDeptList();
  await flush();
  assert.equal(env.rpcFetches()[0].init.headers['X-Idempotency-Key'], undefined);

  env.queue.push([500, {}], [200, { ok: true, result: { ok: true } }], [200, { ok: true, result: { ok: true } }]);
  r.api_approveRequest('req-1');
  await flush();
  r.api_approveRequest('req-1');
  await flush();
  r.api_approveRequest('req-1');
  await flush();
  const keys = env.rpcFetches().slice(1).map((f) => f.init.headers['X-Idempotency-Key']);
  assert.ok(keys[0]);
  assert.equal(keys[1], keys[0], '5xx の後の押し直しは同じキー');
  assert.notEqual(keys[2], keys[0], '200 で捨てたので新しいキー');
});

test('400 でキーを捨て、引数が違えば別のキー', async () => {
  const env = makeEnv();
  const r = env.win.google.script.run.withFailureHandler(() => {});
  env.queue.push([400, { ok: false, error: 'x' }], [400, { ok: false, error: 'x' }], [400, { ok: false, error: 'x' }]);
  r.api_deleteRequest('a'); await flush();
  r.api_deleteRequest('a'); await flush();
  r.api_deleteRequest('b'); await flush();
  const keys = env.rpcFetches().map((f) => f.init.headers['X-Idempotency-Key']);
  assert.notEqual(keys[0], keys[1]);
  assert.notEqual(keys[1], keys[2]);
});

test('同じ書き込みの二度押しは 1 回だけ送り、両方のハンドラに同じ返事', async () => {
  const env = makeEnv();
  const got = [];
  const payload = { requestType: 'overtime', hours: '1.0' };
  env.win.google.script.run.withSuccessHandler((v) => got.push(['a', v])).api_submitRequest(payload);
  env.win.google.script.run.withSuccessHandler((v) => got.push(['b', v])).api_submitRequest(payload);
  await flush();
  assert.equal(env.rpcFetches().length, 1);
  env.held.shift()(200, { ok: true, result: { ok: true, success: true, requestId: 'R1' } });
  await flush();
  assert.deepEqual(JSON.parse(JSON.stringify(got)), [['a', { ok: true, success: true, requestId: 'R1' }], ['b', { ok: true, success: true, requestId: 'R1' }]]);
  // 読む関数は二度押しをまとめない
  env.win.google.script.run.api_getDashboard();
  env.win.google.script.run.api_getDashboard();
  await flush();
  assert.equal(env.rpcFetches().length, 3);
});

test('409 は同じキーで送り直し、通れば成功', async () => {
  const env = makeEnv({ retryMs: 1 });
  let ok = null;
  env.queue.push([409, { ok: false, error: '処理中です' }], [409, { ok: false, error: '処理中です' }], [200, { ok: true, result: { ok: true } }]);
  env.win.google.script.run.withSuccessHandler((v) => { ok = v; }).api_secondaryApprove('r');
  await flush(60);
  const f = env.rpcFetches();
  assert.equal(f.length, 3);
  assert.equal(new Set(f.map((x) => x.init.headers['X-Idempotency-Key'])).size, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(ok)), { ok: true });
});

test('409 が 10 回続いたらサーバーの文で失敗', async () => {
  const env = makeEnv({ retryMs: 1 });
  let err = null;
  for (let i = 0; i < 11; i++) { env.queue.push([409, { ok: false, error: '同じ操作を処理中です。' }]); }
  env.win.google.script.run.withFailureHandler((e) => { err = e; }).api_approveRequest('r');
  await flush(150);
  assert.equal(env.rpcFetches().length, 11);
  assert.equal(err.message, '同じ操作を処理中です。');
});

// ---- 401（罠 20・仕様 3-5） ----

test('401 は失敗ハンドラを呼ばず、書きかけを退避し、ボタンを戻して、ログインへ', async () => {
  const env = makeEnv({ search: '?page=approve2&date=2026-11-25', page: 'approve2' });
  await flush();  // 読み込み直後の /auth/me でメールが分かる
  env.win.__vpsDraft.register({ name: 'p', collect: () => ({ typed: '書きかけ' }), restore() {} });
  env.win.__unloadRemoved = 0;
  env.win.onbeforeunload = () => 'x';

  // 受け口から押したボタンが、処理中に無効になったまま 401 になる
  const btn = makeEl('BUTTON', { 'data-on-click': 'doIt', 'data-args': '["R1","$this"]' });
  btn.innerHTML = '二次承認';
  env.buttons = [btn];
  let failed = false;
  env.win.doIt = function (id, b) {
    b.disabled = true; b.innerHTML = '処理中...';
    env.win.google.script.run.withFailureHandler(() => { failed = true; }).api_secondaryApprove(id);
  };
  env.queue.push([401, { ok: false, error: 'ログインが必要です。', login: '/auth/login' }]);
  env.fire('click', { type: 'click', target: btn });
  assert.equal(btn.disabled, true);
  await flush();

  assert.equal(failed, false, '失敗ハンドラは呼ばない');
  assert.deepEqual(env.alerts, ['ログインが切れました。ログイン後にもう一度押してください。']);
  assert.equal(env.location.href, '/auth/login?next=' + encodeURIComponent('/?page=approve2&date=2026-11-25'));
  assert.equal(btn.disabled, false);
  assert.equal(btn.innerHTML, '二次承認');
  assert.equal(env.win.onbeforeunload, null);
  const saved = JSON.parse(env.store.get('overtime-draft:approve2'));
  assert.equal(saved.email, EMAIL);
  assert.deepEqual(saved.data, { p: { typed: '書きかけ' } });

  // 2 つ目の 401 で alert を重ねない。冪等キーは捨てていない
  env.queue.push([401, {}]);
  env.win.google.script.run.api_secondaryApprove('R1');
  await flush();
  assert.equal(env.alerts.length, 1);
  const keys = env.rpcFetches().map((f) => f.init.headers['X-Idempotency-Key']);
  assert.equal(keys[0], keys[1]);
});

test('ログインから戻ったら、同じメールで 24 時間以内の書きかけだけ戻す', async () => {
  const draft = (email, age) => JSON.stringify({ savedAt: Date.now() - age, email: email, data: { top: { v: 1 } } });
  const cases = [
    [draft(EMAIL, 1000), true],
    [draft('other@lineworks-local.info', 1000), false],
    [draft(EMAIL, 25 * 3600 * 1000), false]
  ];
  for (const [saved, expected] of cases) {
    const env = makeEnv({ session: { 'overtime-draft:top': saved } });
    let restored = null;
    env.win.__vpsDraft.register({ name: 'top', collect: () => null, restore: (d) => { restored = d; } });
    await flush();
    assert.equal(restored !== null, expected, saved);
    assert.equal(env.store.has('overtime-draft:top'), false, '読んだら消す');
  }
});

test('画面に戻ったとき /auth/me が 401 ならその場でログインへ', async () => {
  const env = makeEnv();
  await flush();
  env.me = 401;
  env.fire('visibilitychange', {});
  await flush();
  assert.equal(env.location.href, '/auth/login?next=' + encodeURIComponent('/?page=top'));
  assert.equal(env.alerts.length, 1);
});

// ---- on 属性の代わりの受け口（仕様 3-6） ----

test('受け口は data-on-* の関数を、内側から順に、$this・$value・数・真偽値の引数で呼ぶ', () => {
  const env = makeEnv();
  const calls = [];
  env.win.outer = function () { calls.push(['outer', this === env.global]); };
  env.win.inner = function (a, b, c, d) { calls.push(['inner', a, b === inner, c, d]); };
  env.win.onVal = function (v) { calls.push(['val', v]); };
  const div = makeEl('DIV', { 'data-on-click': 'outer' });
  const inner = makeEl('BUTTON', { 'data-on-click': 'inner', 'data-args': '["x","$this",3,true]' }, div);
  const text = { nodeType: 3, parentNode: inner };
  env.fire('click', { type: 'click', target: text });
  const sel = makeEl('SELECT', { 'data-on-change': 'onVal', 'data-args': '["$value"]' });
  sel.value = '1.5';
  env.fire('change', { type: 'change', target: sel });
  // 無効なボタンの click は呼ばない
  inner.disabled = true;
  env.fire('click', { type: 'click', target: inner });
  assert.deepEqual(calls, [['inner', 'x', true, 3, true], ['outer', true], ['val', '1.5'], ['outer', true]]);
});

test('mouseenter・mouseleave は要素の外から入ったとき・外へ出たときだけ', () => {
  const env = makeEnv();
  const calls = [];
  env.win.__vpsSetHover = (el, t, s) => calls.push([t, s]);
  const btn = makeEl('BUTTON', {
    'data-on-click': 'x',
    'data-on-mouseenter': '__vpsSetHover', 'data-args-mouseenter': '["$this","translateY(-1px)","s"]',
    'data-on-mouseleave': '__vpsSetHover', 'data-args-mouseleave': '["$this","",""]'
  });
  const span = makeEl('SPAN', {}, btn);
  const outside = makeEl('DIV', {});
  env.fire('mouseover', { type: 'mouseover', target: span, relatedTarget: outside });
  env.fire('mouseover', { type: 'mouseover', target: btn, relatedTarget: span });   // 中で動いただけ
  env.fire('mouseout', { type: 'mouseout', target: span, relatedTarget: btn });     // 中で動いただけ
  env.fire('mouseout', { type: 'mouseout', target: btn, relatedTarget: outside });
  assert.deepEqual(calls, [['translateY(-1px)', 's'], ['', '']]);
});

test('__vpsArgs は属性から抜けられない形で、元の値に戻る（XSS の直し 3-8）', () => {
  const env = makeEnv();
  const evil = ["x');alert(1)//", '"><img src=x onerror=alert(1)>', "a&b<c>'d"];
  const attr = env.win.__vpsArgs(evil[0], evil[1], evil[2], 3);
  assert.ok(!/["<>']/.test(attr));
  const decoded = attr.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  assert.deepEqual(JSON.parse(decoded), evil.concat([3]));
});

test('google.script.run の関数は仕様 4-3 の 29 関数と 4-2 の 9 関数', () => {
  const env = makeEnv();
  const names = JSON.parse(JSON.stringify(env.win.__vpsShim.READ_FNS.concat(env.win.__vpsShim.WRITE_FNS))).sort();
  assert.deepEqual(names, require('./rpc-spec').ALL.slice().sort());
});
