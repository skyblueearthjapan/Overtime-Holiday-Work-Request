/* approve2.js — GAS 版 src/approve2.html の <script> をそのまま写したもの（仕様 3-6）。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
var __T = JSON.parse(document.getElementById('gas-template-data').textContent);
var DATA = [];

function z(n){ return String(n).padStart(2,'0'); }
function todayYMD(){
  var d = new Date();
  return d.getFullYear() + '-' + z(d.getMonth()+1) + '-' + z(d.getDate());
}
function fmtMin(min){
  min = Math.max(0, Math.round(Number(min||0)));
  var h = Math.floor(min/60), r = min % 60;
  if(h===0) return r+'分';
  if(r===0) return h+'時間';
  return h+'時間'+r+'分';
}
function escHtml(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

// 初期化: URLパラメータまたは今日の日付
(function(){
  var dateInput = document.getElementById('targetDate');
  var urlDate = __T.DATE_PARAM;
  if (urlDate && /^\d{4}-\d{2}-\d{2}$/.test(urlDate)) {
    dateInput.value = urlDate;
  } else {
    dateInput.value = todayYMD();
  }
  // 初回ロード
  loadList();
})();

function loadList(){
  var dateVal = document.getElementById('targetDate').value;
  if (!dateVal) { alert('日付を選択してください'); return; }

  document.getElementById('loadingMsg').style.display = '';
  google.script.run
    .withSuccessHandler(function(res){
      document.getElementById('loadingMsg').style.display = 'none';
      DATA = res.items || [];
      renderAll();
    })
    .withFailureHandler(function(err){
      document.getElementById('loadingMsg').style.display = 'none';
      alert(err.message || err);
    })
    .api_getSecondaryApprovalList(dateVal);
}

function renderAll(){
  // KPI（二次承認済みはサーバー側で除外済み）
  var total = DATA.length;
  var approvedReady = DATA.filter(function(d){ return d.status === 'approved' && !d.approvedBy2; }).length;
  var pending = DATA.filter(function(d){ return d.status === 'submitted'; }).length;
  var overdue = DATA.filter(function(d){ return d.isOverdue; }).length;

  document.getElementById('kpiTotal').textContent = total;
  document.getElementById('kpiApproved').textContent = approvedReady;
  document.getElementById('kpiPending').textContent = pending;
  document.getElementById('kpiOverdue').textContent = overdue;

  // Table
  var tbody = document.getElementById('tbody');
  tbody.innerHTML = '';

  if (total === 0) {
    tbody.innerHTML = '<tr><td colspan="10" class="muted">申請データがありません</td></tr>';
    updateSelectedCount();
    return;
  }

  DATA.forEach(function(d, i){
    var isReady = (d.status === 'approved' && !d.approvedBy2);
    var isDone = !!d.approvedBy2;
    var isPending = (d.status === 'submitted');

    var tr = document.createElement('tr');
    if (isPending) tr.className = 'disabled';
    if (isDone) tr.className = 'done';
    // 期限超過の行は薄赤背景
    if (d.isOverdue) tr.style.background = 'rgba(255,90,95,0.06)';

    var typeJa = d.requestType === 'overtime' ? '残業' : '休日出勤';

    // チェックボックス
    var cbHtml = isReady
      ? '<input type="checkbox" class="row-cb" data-idx="'+i+'" data-on-change="updateSelectedCount">'
      : '';

    // 対象日（期限超過時は⚠マーク付き）
    var targetDateHtml = d.isOverdue
      ? '<span style="color:var(--danger)">⚠ ' + escHtml(d.targetDate) + '</span>'
      : escHtml(d.targetDate);

    // 一次承認
    var approvedLabel = d.status === 'approved'
      ? '<span class="pill ok">承認済</span>'
      : '<span class="pill warn">未承認</span>';

    // 二次承認
    var approved2Label = isDone
      ? '<span class="pill ok">承認済</span>'
      : (isReady ? '<span class="pill info">未承認</span>' : '<span class="pill">-</span>');

    // PDF
    var pdfLabel = d.pdfFileId
      ? '<a href="https://drive.google.com/file/d/'+escHtml(d.pdfFileId)+'/view" target="_blank" class="pill ok">開く</a>'
      : '<span class="pill muted">-</span>';

    // 操作
    var actionHtml = '';
    if (isDone) {
      actionHtml = '<span class="pill ok">承認済</span>';
    } else if (isReady) {
      actionHtml = '<button class="btn-sm" data-on-click="singleApprove" data-args="'+__vpsArgs(d.requestId, i, '$this')+'">二次承認</button>';
    } else {
      actionHtml = '<span class="pill muted">一次承認待ち</span>';
    }

    tr.innerHTML =
      '<td>'+cbHtml+'</td>'+
      '<td>'+escHtml(d.dept)+'</td>'+
      '<td><b>'+escHtml(d.workerName)+'</b></td>'+
      '<td>'+targetDateHtml+'</td>'+
      '<td>'+typeJa+'</td>'+
      '<td>'+fmtMin(d.approvedMinutes)+'</td>'+
      '<td>'+approvedLabel+'</td>'+
      '<td>'+approved2Label+'</td>'+
      '<td>'+pdfLabel+'</td>'+
      '<td>'+actionHtml+'</td>';

    tbody.appendChild(tr);
  });

  document.getElementById('selectAll').checked = false;
  updateSelectedCount();
}

function getSelectedIds(){
  var cbs = document.querySelectorAll('.row-cb:checked');
  var ids = [];
  cbs.forEach(function(cb){
    var idx = Number(cb.dataset.idx);
    if (DATA[idx]) ids.push(DATA[idx].requestId);
  });
  return ids;
}

function updateSelectedCount(){
  var ids = getSelectedIds();
  document.getElementById('selectedCount').textContent = ids.length;
  document.getElementById('batchBtn').disabled = (ids.length === 0);
}

function toggleSelectAll(){
  var checked = document.getElementById('selectAll').checked;
  document.querySelectorAll('.row-cb').forEach(function(cb){ cb.checked = checked; });
  updateSelectedCount();
}

function singleApprove(requestId, idx, btn){
  btn.disabled = true;
  btn.textContent = '処理中...';

  // 楽観的UI更新
  if (DATA[idx]) {
    DATA[idx].approvedBy2 = '(処理中)';
  }

  google.script.run
    .withSuccessHandler(function(res){
      if (res.ok) {
        if (DATA[idx]) {
          DATA[idx].approvedBy2 = res.approvedBy2 || '承認済';
          DATA[idx].approvedAt2 = res.approvedAt2 || '';
          DATA[idx].pdfFileId = ''; // PDF再生成中
        }
        renderAll();
      } else {
        alert(res.message || 'エラーが発生しました');
        loadList();
      }
    })
    .withFailureHandler(function(err){
      alert(err.message || err);
      loadList();
    })
    .api_secondaryApprove(requestId);
}

function batchApprove(){
  var ids = getSelectedIds();
  if (ids.length === 0) return;
  if (!confirm(ids.length + '件を一括二次承認しますか？')) return;

  document.getElementById('batchBtn').disabled = true;
  document.getElementById('batchBtn').textContent = '処理中...';

  google.script.run
    .withSuccessHandler(function(res){
      document.getElementById('batchBtn').textContent = '一括二次承認';
      if (res.ok) {
        var failCount = 0;
        (res.results || []).forEach(function(r){
          if (!r.ok) failCount++;
        });
        if (failCount > 0) {
          alert('一部エラー: ' + failCount + '件が失敗しました');
        }
        loadList(); // リロード
      } else {
        alert('エラーが発生しました');
        loadList();
      }
    })
    .withFailureHandler(function(err){
      document.getElementById('batchBtn').textContent = '一括二次承認';
      alert(err.message || err);
      loadList();
    })
    .api_secondaryApproveBatch(ids);
}
