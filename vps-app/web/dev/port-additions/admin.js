
/* ===================================================================
 * VPS 版で足した部分（ここから下だけ。上は GAS 版 src/admin.html のまま）
 * =================================================================== */

// ====== 運用（管理者）の欄（仕様 3-10・4-2） ======
// 関数名をそのまま書く（google.script.run に生えている名前と照らし合わせる試験のため）
var OPS_CALLS = {
  ops_syncSheetConfigNow: function(r){ r.ops_syncSheetConfigNow(); },
  ops_runMasterSync: function(r){ r.ops_runMasterSync(); },
  ops_rebuildLedger: function(r){ r.ops_rebuildLedger(); },
  ops_recalcBreaks: function(r){ r.ops_recalcBreaks(); },
  ops_runSheetIntake: function(r){ r.ops_runSheetIntake(); },
  ops_regeneratePdf: function(r, v){ r.ops_regeneratePdf(v); },
  ops_batchPdfForDate: function(r, v){ r.ops_batchPdfForDate(v); },
  ops_revokeSessions: function(r, v){ r.ops_revokeSessions(v); },
  ops_retryJob: function(r, v){ r.ops_retryJob(v); }
};

function opsLoadStatus(){
  var pre = document.getElementById('opsStatus');
  pre.textContent = '読込中...';
  google.script.run.withSuccessHandler(function(st){
    pre.textContent = JSON.stringify(st, null, 2);
    opsRenderFailed(st);
  }).withFailureHandler(function(err){
    pre.textContent = 'エラー: ' + (err.message || err);
  }).ops_getStatus();
}

// 状態の中の「失敗した仕事」の一覧（id を持つ配列）を探して、行ごとに［再実行］を付ける
function opsFindFailedJobs(st){
  var found = [];
  (function walk(o, depth){
    if (!o || typeof o !== 'object' || depth > 3) return;
    Object.keys(o).forEach(function(k){
      var v = o[k];
      if (Array.isArray(v) && /fail/i.test(k)) {
        v.forEach(function(j){ if (j && typeof j === 'object' && j.id !== undefined && j.id !== null) found.push(j); });
      } else if (v && typeof v === 'object' && !Array.isArray(v)) {
        walk(v, depth + 1);
      }
    });
  })(st, 0);
  return found;
}

function opsRenderFailed(st){
  var box = document.getElementById('opsFailed');
  var jobs = opsFindFailedJobs(st);
  if (jobs.length === 0) { box.innerHTML = '<div class="sub">失敗した仕事はありません</div>'; return; }
  var html = '<table><thead><tr><th>id</th><th>申請 ID</th><th>仕事</th><th>理由</th><th>回数</th><th></th></tr></thead><tbody>';
  jobs.forEach(function(j){
    html += '<tr>' +
      '<td>' + escHtml(j.id) + '</td>' +
      '<td>' + escHtml(j.request_id || j.requestId || '') + '</td>' +
      '<td>' + escHtml(j.action || j.kind || '') + '</td>' +
      '<td>' + escHtml(j.last_error || j.lastError || j.reason || '') + '</td>' +
      '<td>' + escHtml(j.attempts !== undefined ? j.attempts : '') + '</td>' +
      '<td><button class="ghost" data-on-click="opsRetryJob" data-args="' + __vpsArgs(j.id, '$this') + '">再実行</button></td>' +
      '</tr>';
  });
  box.innerHTML = html + '</tbody></table>';
}

function opsCall(fn, value, label){
  if (!confirm('「' + label + '」を実行しますか？')) return;
  var r = google.script.run.withSuccessHandler(function(res){
    alert('「' + label + '」を受け付けました。\n' + JSON.stringify(res));
    opsLoadStatus();
  }).withFailureHandler(function(err){
    alert('エラー: ' + (err.message || err));
  });
  OPS_CALLS[fn](r, value);
}

function opsRun(fn, btn){
  opsCall(fn, undefined, btn.textContent);
}

function opsRunWithInput(fn, inputId, btn){
  var v = document.getElementById(inputId).value.trim();
  if (!v) { alert('値を入力してください。'); return; }
  opsCall(fn, v, btn.textContent + '（' + v + '）');
}

function opsRetryJob(jobId, btn){
  opsCall('ops_retryJob', jobId, '失敗した仕事の再実行（id ' + jobId + '）');
}

(function(){
  var panel = document.getElementById('opsPanel');
  if (panel) panel.addEventListener('toggle', function(){ if (panel.open) opsLoadStatus(); });
})();

// ====== 操作中にログインが切れたときの書きかけ（手動実績入力）の退避と復元（仕様 3-5） ======
(function(){
  if (!window.__vpsDraft) return;
  window.__vpsDraft.register({
    name: 'admin',
    collect: function(){
      var wrap = document.getElementById('manualEntryWrap');
      if (!wrap || wrap.style.display === 'none' || !ME_REQUEST_ID) return null;
      return {
        requestId: ME_REQUEST_ID,
        workerName: document.getElementById('meWorker').innerText,
        requestType: document.getElementById('meType').innerText === '残業' ? 'overtime' : 'holiday',
        dateYmd: document.getElementById('meDate').innerText,
        start: document.getElementById('meStart').value,
        end: document.getElementById('meEnd').value
      };
    },
    restore: function(d){
      if (!d || !d.requestId) return;
      openManualEntry(d.requestId, d.workerName, d.requestType, d.dateYmd, '');
      document.getElementById('meStart').value = d.start || '';
      document.getElementById('meEnd').value = d.end || '';
    }
  });
})();
