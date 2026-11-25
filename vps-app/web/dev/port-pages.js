#!/usr/bin/env node
// GAS の画面（src/*.html）から VPS 版の画面を機械的に作る（仕様 3 章）。
//
//   node vps-app/web/dev/port-pages.js          作り直して書き出す
//   node vps-app/web/dev/port-pages.js --check  書き出した物が src/ から作った物と同じかだけ確かめる（違えば終了コード 1）
//
// 作る物:
//   vps-app/api/app/templates/pages/<page>.html  … HTML と CSS は元のまま。テンプレートの差し込みを Jinja2 の {{ }} に、
//                                                   on 属性を data-on-<event> に、<script> を外部の .js の読み込みにする
//   vps-app/web/static/js/<page>.js              … 元の <script> の中身。変えるのは下の PATCHES に書いた所だけ
// 手で直さないこと（直すときはこのファイルか port-additions/ を直して作り直す）。
'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..', '..');          // リポジトリの直下
const SRC = path.join(ROOT, 'src');
const TPL_DIR = path.join(ROOT, 'vps-app', 'api', 'app', 'templates', 'pages');
const JS_DIR = path.join(ROOT, 'vps-app', 'web', 'static', 'js');
const ADD_DIR = path.join(__dirname, 'port-additions');

const PAGES = ['top', 'approver', 'approve2', 'admin', 'staffSummary', 'no_auth'];
const TITLE = '残業・休日出勤申請';

// ---- 置き換えの道具（数が合わなければ止める） ----
function replaceExact(text, from, to, count, label) {
  const n = text.split(from).length - 1;
  if (n !== count) { throw new Error(label + ': 「' + from.slice(0, 60) + '」が ' + n + ' 件（' + count + ' 件のはず）'); }
  return text.split(from).join(to);
}
function replaceRegex(text, re, fn, count, label) {
  let n = 0;
  const out = text.replace(re, function () { n++; return fn.apply(null, arguments); });
  if (n !== count) { throw new Error(label + ': ' + re + ' が ' + n + ' 件（' + count + ' 件のはず）'); }
  return out;
}

// ---- HTML の on 属性 → data-on-<event>・data-args（仕様 3-6） ----
// 関数の呼び出しでない文は、ページの .js の「VPS 版で足した部分」の関数にする
const SPECIAL_HANDLERS = {
  "this.style.transform='translateY(-1px)';this.style.boxShadow='0 12px 28px rgba(14,165,233,.30)'":
    ['__vpsSetHover', ['$this', 'translateY(-1px)', '0 12px 28px rgba(14,165,233,.30)']],
  "this.style.transform='translateY(-1px)';this.style.boxShadow='0 12px 28px rgba(31,157,85,.30)'":
    ['__vpsSetHover', ['$this', 'translateY(-1px)', '0 12px 28px rgba(31,157,85,.30)']],
  "this.style.transform='';this.style.boxShadow=''":
    ['__vpsSetHover', ['$this', '', '']],
  "selectReqDate(document.getElementById('reqManualDate').value)":
    ['__vpsSelectReqManualDate', []]
};
const ON_ATTR = /\son(click|input|change|mouseenter|mouseleave)="([^"]*)"/g;

function parseHandler(code) {
  if (Object.prototype.hasOwnProperty.call(SPECIAL_HANDLERS, code)) { return SPECIAL_HANDLERS[code]; }
  const m = /^\s*([A-Za-z_$][\w$]*)\((.*)\)\s*;?\s*$/.exec(code);
  if (!m || /\bthis\b|\bdocument\b|\bevent\b|;/.test(m[2])) { throw new Error('機械的に置き換えられない on 属性: ' + code); }
  // 引数は文字・数・真偽値の並びだけ（src は信頼できる元なので、ここで評価してよい）
  const args = new Function('return [' + m[2] + '];')();
  return [m[1], args];
}
function attrJson(v) {
  return JSON.stringify(v).replace(/&/g, '&amp;').replace(/'/g, '&#39;');
}
function convertOnAttributes(html, label) {
  let total = 0;
  const out = html.replace(/<[A-Za-z][^>]*>/g, function (tag) {
    const found = [];
    tag.replace(ON_ATTR, function (all, ev, code) { found.push([ev, code]); return all; });
    if (found.length === 0) { return tag; }
    total += found.length;
    const single = found.length === 1;
    return tag.replace(ON_ATTR, function (all, ev, code) {
      const h = parseHandler(code);
      let s = ' data-on-' + ev + '="' + h[0] + '"';
      if (h[1].length > 0) { s += ' ' + (single ? 'data-args' : 'data-args-' + ev) + "='" + attrJson(h[1]) + "'"; }
      return s;
    });
  });
  if (/<[A-Za-z][^>]*\son[a-z]+\s*=/.test(out)) { throw new Error(label + ': on 属性が残っている'); }
  return { html: out, count: total };
}

// ---- JS の文字列で作る onclick（top・approver の同じ形のもの） ----
// '...onclick="FN(\'' + EXPR + '\')"'           → data-on-click="FN" data-args="' + __vpsArgs(EXPR) + '"
// '...onclick="FN(\'' + EXPR + '\',this)"'      → __vpsArgs(EXPR, '$this')
// '...onclick="FN(\'' + EXPR + '\',\'ot-end\')"' → __vpsArgs(EXPR, 'ot-end')
const JS_ONCLICK = /onclick="(\w+)\(\\'' \+ ([\w.]+) \+ '\\'(,this|,\\'[\w-]+\\')?\)"/g;
function jsOnclick(all, fn, expr, extra) {
  let args = expr;
  if (extra === ',this') { args += ", '$this'"; }
  else if (extra) { args += ", '" + extra.slice(3, -2) + "'"; }
  return 'data-on-click="' + fn + '" data-args="\' + __vpsArgs(' + args + ') + \'"';
}

const ESC_OLD = ".replace(/\"/g,'&quot;');";
const ESC_NEW = ".replace(/\"/g,'&quot;').replace(/'/g,'&#39;');";
const TEMPLATE_DATA = "var __T = JSON.parse(document.getElementById('gas-template-data').textContent);\n";

// ---- ページごとの直し ----
const PATCHES = {
  top: {
    html: [
      // 3-6 の機械的な置き換えだけ
    ],
    js: [
      (t) => replaceExact(t, 'var REASONS = <?!= REASONS_JSON ?>;', 'var REASONS = __T.REASONS;', 1, 'top REASONS'),
      (t) => replaceExact(t, 'var OT_HOURS_LIST = <?!= OT_HOURS_JSON ?>;', 'var OT_HOURS_LIST = __T.OT_HOURS;', 1, 'top OT_HOURS'),
      (t) => replaceExact(t, 'var HD_HOURS_LIST = <?!= HD_HOURS_JSON ?>;', 'var HD_HOURS_LIST = __T.HD_HOURS;', 1, 'top HD_HOURS'),
      (t) => replaceExact(t, ESC_OLD, ESC_NEW, 1, 'top escHtml'),
      (t) => replaceRegex(t, JS_ONCLICK, jsOnclick, 15, 'top onclick'),
      // 使われていないコード（工番選択モーダル）。CSP で on 属性を使えないので同じ形に直す
      (t) => replaceExact(t,
        "'<div class=\"order-item\" onclick=\"selectOrder(\\'' + o.code.replace(/'/g, \"\\\\'\") + '\\')\">'",
        "'<div class=\"order-item\" data-on-click=\"selectOrder\" data-args=\"' + __vpsArgs(o.code) + '\">'",
        1, 'top selectOrder'),
      // 使われていないコード（フォームへのリンク）
      (t) => replaceExact(t,
        'onclick="this.parentElement.parentElement.remove()"',
        'data-on-click="__vpsRemoveGrandparent" data-args="[&quot;$this&quot;]"',
        1, 'top showFormLink'),
      // 3-8: PDF リンクの pdfFileId も escHtml を通す
      (t) => replaceExact(t,
        "return '<a href=\"https://drive.google.com/file/d/' +\n      it.pdfFileId + '/view\"",
        "return '<a href=\"https://drive.google.com/file/d/' +\n      escHtml(it.pdfFileId) + '/view\"",
        1, 'top pdfFileId')
    ]
  },
  approver: {
    html: [
      // 3-6・gap-2 S5: 一括承認のボタンは onclick の文字ではなく id で探す
      (h) => replaceExact(h, "onclick=\"approveAllInTable('overtime')\"", "id=\"btnApproveAll-overtime\" onclick=\"approveAllInTable('overtime')\"", 1, 'approver btn overtime'),
      (h) => replaceExact(h, "onclick=\"approveAllInTable('holiday')\"", "id=\"btnApproveAll-holiday\" onclick=\"approveAllInTable('holiday')\"", 1, 'approver btn holiday')
    ],
    js: [
      (t) => replaceExact(t, ESC_OLD, ESC_NEW, 1, 'approver escHtml'),
      (t) => replaceRegex(t, JS_ONCLICK, jsOnclick, 2, 'approver onclick'),
      (t) => replaceExact(t,
        "var batchBtn = document.querySelector('[onclick*=\"approveAllInTable(\\'' + type + '\\')\"]');",
        "var batchBtn = document.getElementById('btnApproveAll-' + type);",
        1, 'approver batchBtn')
    ]
  },
  approve2: {
    html: [],
    js: [
      (t) => replaceExact(t, "var urlDate = '<?= DATE_PARAM ?>';", 'var urlDate = __T.DATE_PARAM;', 1, 'approve2 DATE_PARAM'),
      (t) => replaceExact(t, ESC_OLD, ESC_NEW, 1, 'approve2 escHtml'),
      (t) => replaceExact(t,
        "'<input type=\"checkbox\" class=\"row-cb\" data-idx=\"'+i+'\" onchange=\"updateSelectedCount()\">'",
        "'<input type=\"checkbox\" class=\"row-cb\" data-idx=\"'+i+'\" data-on-change=\"updateSelectedCount\">'",
        1, 'approve2 row-cb'),
      (t) => replaceExact(t,
        "'<button class=\"btn-sm\" onclick=\"singleApprove(\\''+escHtml(d.requestId)+'\\', '+i+', this)\">二次承認</button>'",
        "'<button class=\"btn-sm\" data-on-click=\"singleApprove\" data-args=\"'+__vpsArgs(d.requestId, i, '$this')+'\">二次承認</button>'",
        1, 'approve2 singleApprove')
    ]
  },
  admin: {
    html: [
      // 3-10: 運用（管理者）の欄を一番下（未承認滞留監視の後ろ）に足す
      (h) => replaceExact(h,
        '      <div class="sub" style="margin-top:8px">※status=submitted のまま放置されている申請を検出。</div>\n    </div>\n',
        '      <div class="sub" style="margin-top:8px">※status=submitted のまま放置されている申請を検出。</div>\n    </div>\n' +
          fs.readFileSync(path.join(ADD_DIR, 'admin-ops.html'), 'utf8').replace(/^\n/, '').replace(/\r\n/g, '\n'),
        1, 'admin ops panel')
    ],
    js: [
      (t) => replaceExact(t, "var URL_DATE_PARAM = '<?= DATE_PARAM ?>';", 'var URL_DATE_PARAM = __T.DATE_PARAM;', 1, 'admin DATE_PARAM'),
      (t) => replaceExact(t, ESC_OLD, ESC_NEW, 1, 'admin escHtml'),
      // 3-8・gap-2 S8: 手動入力ボタン。氏名などを JS の文字列に入れず、JSON（HTML の属性として逃がしたもの）で渡す
      (t) => replaceExact(t,
        "'onclick=\"openManualEntry(\\''+escHtml(it.requestId)+'\\',\\''+escHtml(it.workerName)+'\\',\\''+\n" +
        "          escHtml(it.requestType)+'\\',\\''+escHtml(res.date)+'\\',\\''+escHtml(it.actualStartAt||'')+'\\')\">手動入力</button>'",
        "'data-on-click=\"openManualEntry\" data-args=\"'+__vpsArgs(it.requestId, it.workerName,\n" +
        "          it.requestType, res.date, it.actualStartAt||'')+'\">手動入力</button>'",
        1, 'admin openManualEntry')
    ]
  },
  staffSummary: {
    html: [
      (h) => replaceExact(h, '<title>スタッフビュー画面 — 月次 残業・休日出勤 集計</title>', '<title>' + TITLE + '</title>', 1, 'staffSummary title')
    ],
    js: [
      (t) => replaceExact(t,
        "function esc(s){ return String(s==null?'':s).replace(/[&<>\"]/g, function(c){\n    return {'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[c]; }); }",
        "function esc(s){ return String(s==null?'':s).replace(/[&<>\"']/g, function(c){\n    return {'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',\"'\":'&#39;'}[c]; }); }",
        1, 'staffSummary esc')
    ]
  },
  no_auth: { html: [], js: [] }
};

const USES_TEMPLATE_DATA = { top: true, approve2: true, admin: true };

function buildPage(page) {
  const src = fs.readFileSync(path.join(SRC, page + '.html'), 'utf8').replace(/\r\n/g, '\n');
  const patch = PATCHES[page];

  // <script> は各ページに 1 つだけ（最後）。no_auth には無い
  const scripts = src.match(/<script>[\s\S]*?<\/script>/g) || [];
  if (scripts.length > 1) { throw new Error(page + ': <script> が ' + scripts.length + ' 個'); }
  let html = src;
  let js = null;
  if (scripts.length === 1) {
    js = scripts[0].slice('<script>'.length, -'</script>'.length);
    html = html.replace(scripts[0], '\u0000SCRIPT\u0000');
  }

  // ---- HTML ----
  for (const f of patch.html) { html = f(html); }
  html = replaceExact(html, '<?= APP_URL ?>', '{{ APP_URL }}', (src.split('<?= APP_URL ?>').length - 1), page + ' APP_URL');
  html = replaceExact(html, '<?= PORTAL_URL ?>', '{{ PORTAL_URL }}', 1, page + ' PORTAL_URL');
  if (page === 'no_auth') { html = replaceExact(html, '<?= message ?>', '{{ message }}', 1, 'no_auth message'); }
  if (/<\?/.test(html)) { throw new Error(page + ': GAS のテンプレートの書き方が残っている'); }
  const conv = convertOnAttributes(html, page);
  html = conv.html;
  // 題名（GAS の setTitle の代わり。仕様 3-2）
  if (!/<title>/.test(html)) {
    const t = page === 'no_auth' ? '権限エラー' : TITLE;
    html = replaceRegex(html, /(<meta name="viewport"[^>]*>\n)/, (all, m) => m + '  <title>' + t + '</title>\n', 1, page + ' title');
  }
  // スマホ向けの CSS（仕様 3-9。すべて 820px 以下の @media の中）
  html = replaceExact(html, '</head>', '<link rel="stylesheet" href="/static/css/mobile.css">\n</head>', 1, page + ' mobile.css');
  html = replaceExact(html, '<body>', '<body data-page="' + page + '">', 1, page + ' body');
  if (js !== null) {
    html = html.replace('\u0000SCRIPT\u0000',
      '<script type="application/json" id="gas-template-data">{{ data_json|safe }}</script>\n' +
      '<script src="/static/js/gas-shim.js"></script>\n' +
      '<script src="/static/js/' + page + '.js"></script>');
  }

  // ---- JS ----
  if (js !== null) {
    for (const f of patch.js) { js = f(js); }
    if (/<\?/.test(js)) { throw new Error(page + ': JS に GAS のテンプレートの書き方が残っている'); }
    if (/\son(click|input|change|mouseenter|mouseleave)=/.test(js)) { throw new Error(page + ': JS の文字列に on 属性が残っている'); }
    const head = '/* ' + page + '.js — GAS 版 src/' + page + '.html の <script> をそのまま写したもの（仕様 3-6）。\n' +
      ' * vps-app/web/dev/port-pages.js が作る。手で直さない。 */\n' +
      (USES_TEMPLATE_DATA[page] ? TEMPLATE_DATA : '');
    const addFile = path.join(ADD_DIR, page + '.js');
    const tail = fs.existsSync(addFile) ? fs.readFileSync(addFile, 'utf8').replace(/\r\n/g, '\n') : '';
    js = head + js.replace(/^\n/, '') + tail;
  }
  return { html: html, js: js, onAttrs: conv.count };
}

function main() {
  const check = process.argv.includes('--check');
  const diffs = [];
  let totalOn = 0;
  for (const page of PAGES) {
    const out = buildPage(page);
    totalOn += out.onAttrs;
    const targets = [[path.join(TPL_DIR, page + '.html'), out.html]];
    if (out.js !== null) { targets.push([path.join(JS_DIR, page + '.js'), out.js]); }
    for (const [file, text] of targets) {
      if (check) {
        const cur = fs.existsSync(file) ? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n') : null;
        if (cur !== text) { diffs.push(path.relative(ROOT, file)); }
      } else {
        fs.mkdirSync(path.dirname(file), { recursive: true });
        fs.writeFileSync(file, text);
      }
    }
  }
  if (check) {
    if (diffs.length) { console.error('src/ から作り直した物と違う: ' + diffs.join(', ')); process.exit(1); }
    console.log('ok: 書き出した画面は src/ から作った物と同じ（on 属性の置き換え ' + totalOn + ' 個）');
  } else {
    console.log('書き出した（on 属性の置き換え ' + totalOn + ' 個）');
  }
}

module.exports = { buildPage, PAGES };
if (require.main === module) { main(); }
