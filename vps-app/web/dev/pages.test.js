// 5 ページの JS の試験（仕様 3-6・3-8・12-2 の U）。jsdom は使わず、何でも受け付ける偽物の DOM で読み込む。
//   node --test vps-app/web/dev/
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { spawnSync } = require('child_process');
const { webcrypto } = require('crypto');
const SPEC = require('./rpc-spec');

const VPS = path.resolve(__dirname, '..', '..');
const JS_DIR = path.join(VPS, 'web', 'static', 'js');
const TPL_DIR = path.join(VPS, 'api', 'app', 'templates', 'pages');
const PAGES = ['top', 'approver', 'approve2', 'admin', 'staffSummary'];
const SHIM = fs.readFileSync(path.join(JS_DIR, 'gas-shim.js'), 'utf8');

// 何を聞かれても「それらしい空の物」を返す偽物（関数としても呼べる）
function makeAny() {
  const store = {};
  const target = function () {};
  return new Proxy(target, {
    get(t, k) {
      if (Object.prototype.hasOwnProperty.call(store, k)) { return store[k]; }
      if (k === Symbol.toPrimitive) { return (hint) => (hint === 'number' ? 0 : ''); }
      if (k === Symbol.iterator) { return function* () {}; }
      if (k === 'then') { return undefined; }
      if (k === 'length') { return 0; }
      if (k === 'options' || k === 'children' || k === 'childNodes') { return []; }
      if (['value', 'textContent', 'innerHTML', 'innerText', 'className', 'id'].includes(k)) { return ''; }
      if (k === 'checked' || k === 'disabled' || k === 'open') { return false; }
      return makeAny();
    },
    set(t, k, v) { store[k] = v; return true; },
    apply() { return makeAny(); },
    construct() { return makeAny(); }
  });
}

function loadPage(page, templateData) {
  const calls = [];
  const els = {};
  const document = makeAny();
  const docStore = {
    readyState: 'loading',
    getElementById(id) {
      if (id === 'gas-template-data') { return { textContent: JSON.stringify(templateData) }; }
      if (!els[id]) { els[id] = makeAny(); }
      return els[id];
    },
    addEventListener() {},
    querySelector() { return makeAny(); },
    querySelectorAll() { return []; },
    createElement() { return makeAny(); },
    contains() { return true; }
  };
  const doc = new Proxy(document, {
    get(t, k) { return Object.prototype.hasOwnProperty.call(docStore, k) ? docStore[k] : t[k]; }
  });
  const win = {
    document: doc, location: { pathname: '/', search: '?page=' + page, href: '' },
    alert() {}, confirm() { return false; }, prompt() { return null; },
    setTimeout() { return 0; }, clearTimeout() {}, setInterval() { return 0; }, clearInterval() {},
    addEventListener() {}, removeEventListener() {}, scrollTo() {},
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    crypto: webcrypto, devicePixelRatio: 1, console: { log() {}, error() {}, warn() {} },
    URL: { createObjectURL() { return ''; }, revokeObjectURL() {} }, Blob: function () {},
    fetch(url, init) {
      if (url.startsWith('/api/rpc/')) { calls.push({ name: url.slice('/api/rpc/'.length), body: init.body }); }
      return new Promise(() => {});
    }
  };
  win.window = win;
  vm.createContext(win);
  vm.runInContext(SHIM, win, { filename: 'gas-shim.js' });
  const code = fs.readFileSync(path.join(JS_DIR, page + '.js'), 'utf8');
  vm.runInContext(code, win, { filename: page + '.js' });
  return { win, calls };
}

const DATA = {
  PAGE: 'top', DATE_PARAM: '2026-11-25',
  REASONS: ['急ぎ作業を要するため', 'その他（→補足理由必須）'], OT_HOURS: ['0.5', '1.0'], HD_HOURS: ['半日', '1日']
};

test('画面は src/ から port-pages.js で作った物と同じ（手で直していない）', () => {
  const r = spawnSync(process.execPath, [path.join(__dirname, 'port-pages.js'), '--check'], { encoding: 'utf8' });
  assert.equal(r.status, 0, r.stderr);
});

for (const page of PAGES) {
  test(page + '.js: 構文の誤りが無く、偽物の DOM で最後まで読み込める', () => {
    const code = fs.readFileSync(path.join(JS_DIR, page + '.js'), 'utf8');
    assert.doesNotThrow(() => new vm.Script(code, { filename: page + '.js' }));
    const { calls } = loadPage(page, Object.assign({}, DATA, { PAGE: page }));
    // 読み込み時に呼ぶ RPC はすべて閉じたリストの中（無い名前ならシムに生えておらず TypeError になる）
    for (const c of calls) { assert.ok(SPEC.ALL.includes(c.name), c.name); }
  });

  test(page + '.js: data-on-* で呼ぶ関数がすべてある', () => {
    const { win } = loadPage(page, Object.assign({}, DATA, { PAGE: page }));
    const html = fs.readFileSync(path.join(TPL_DIR, page + '.html'), 'utf8');
    const js = fs.readFileSync(path.join(JS_DIR, page + '.js'), 'utf8');
    const names = new Set();
    for (const m of (html + '\n' + js).matchAll(/data-on-(?:click|input|change|mouseenter|mouseleave)=\\?"(\w+)/g)) { names.add(m[1]); }
    if (page !== 'staffSummary') { assert.ok(names.size > 0); }  // スタッフ集計は元から on 属性が無い
    for (const n of names) { assert.equal(vm.runInContext('typeof ' + n, win), 'function', n); }
  });
}

test('escHtml（4 ページ）と esc（スタッフ集計）が引用符を逃がす（仕様 3-8）', () => {
  const s = '<a href="x">\'&';
  for (const page of ['top', 'approver', 'approve2', 'admin']) {
    const { win } = loadPage(page, DATA);
    assert.equal(win.escHtml(s), '&lt;a href=&quot;x&quot;&gt;&#39;&amp;', page);
  }
  const { win } = loadPage('staffSummary', DATA);
  assert.equal(win.esc(s), '&lt;a href=&quot;x&quot;&gt;&#39;&amp;');
  assert.equal(win.esc(null), '');
});

test('テンプレートの値が今の変数名で受け取られる', () => {
  const top = loadPage('top', DATA).win;
  assert.deepEqual(JSON.parse(JSON.stringify(vm.runInContext('[REASONS, OT_HOURS_LIST, HD_HOURS_LIST]', top))),
    [DATA.REASONS, DATA.OT_HOURS, DATA.HD_HOURS]);
  const admin = loadPage('admin', DATA).win;
  assert.equal(vm.runInContext('URL_DATE_PARAM', admin), '2026-11-25');
  // approve2 は IIFE の中の urlDate。?date= の日で一覧を読む
  const a2 = loadPage('approve2', DATA);
  assert.deepEqual(a2.calls.map((c) => [c.name, c.body]), [['api_getSecondaryApprovalList', '["2026-11-25"]']]);
  // admin は ?date= があれば日次一覧も自動で読む（9-1 #22）
  const ad = loadPage('admin', DATA);
  assert.ok(ad.calls.some((c) => c.name === 'api_adminDailyDetail' && c.body.startsWith('["2026-11-25"')));
  const ad2 = loadPage('admin', Object.assign({}, DATA, { DATE_PARAM: '' }));
  assert.ok(!ad2.calls.some((c) => c.name === 'api_adminDailyDetail'));
});

test('画面が呼ぶ RPC の名前は仕様の 29 関数＋管理者の 9 関数（と使われていないコードの 1 つ）だけ', () => {
  const used = new Set();
  const usedBy = {};
  for (const page of PAGES) {
    const js = fs.readFileSync(path.join(JS_DIR, page + '.js'), 'utf8');
    for (const m of js.matchAll(/\.((?:api|ops)_\w+)\s*\(/g)) {
      used.add(m[1]);
      (usedBy[m[1]] = usedBy[m[1]] || new Set()).add(page);
    }
  }
  const unknown = [...used].filter((n) => !SPEC.ALL.includes(n) && !SPEC.DEAD_CODE_ONLY.includes(n));
  assert.deepEqual(unknown, [], '仕様に無い関数');
  const missing = SPEC.ALL.filter((n) => !used.has(n));
  assert.deepEqual(missing, [], '画面から呼ばれない仕様の関数');
  // 仕様に無い api_getFormUrl は top の使われていないコード（doOpenForm）からだけ
  assert.deepEqual([...usedBy.api_getFormUrl], ['top']);
  const top = fs.readFileSync(path.join(JS_DIR, 'top.js'), 'utf8');
  const doOpenForm = top.slice(top.indexOf('function doOpenForm('), top.indexOf('function showFormLink('));
  assert.ok(doOpenForm.includes('.api_getFormUrl('));
  assert.equal(top.split('.api_getFormUrl(').length - 1, 1);
  assert.equal(top.split('doOpenForm(').length - 1, 2, 'doOpenForm は定義と使われていない submitOrderPicker からだけ');
});
