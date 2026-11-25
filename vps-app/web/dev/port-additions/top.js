
/* ===================================================================
 * VPS 版で足した部分（ここから下だけ。上は GAS 版 src/top.html のまま）
 * =================================================================== */

// on 属性の中に書かれていた文（関数の呼び出しでないもの）を関数にしたもの（仕様 3-6）
function __vpsSetHover(el, transform, shadow){ el.style.transform = transform; el.style.boxShadow = shadow; }
function __vpsSelectReqManualDate(){ selectReqDate(document.getElementById('reqManualDate').value); }
function __vpsRemoveGrandparent(el){ el.parentElement.parentElement.remove(); }

// 申請モーダルを開く前に、ログインが続いているかを確かめる（仕様 3-5）
(function(){
  var orig = showRequestModal;
  showRequestModal = function(){
    if (window.__vpsCheckSession) window.__vpsCheckSession();
    return orig.apply(this, arguments);
  };
})();

// 操作中にログインが切れたときの書きかけ（申請モーダル・編集モーダル）の退避と復元（仕様 3-5）
(function(){
  if (!window.__vpsDraft) return;
  var when = window.__vpsDraft.when;
  function val(id){ var el = document.getElementById(id); return el ? el.value : ''; }
  function setVal(id, v){ var el = document.getElementById(id); if (el && v !== undefined && v !== null) el.value = v; }
  function shown(id){ var el = document.getElementById(id); return !!el && el.style.display !== 'none'; }

  function restoreRequest(r){
    // 申請カードの部署・作業員（部署一覧が読めてから）
    when(function(){ return document.getElementById('deptSelect').options.length > 1; }, function(){
      if (!r.dept) return;
      document.getElementById('deptSelect').value = r.dept;
      onDeptChange();
      when(function(){ var w = document.getElementById('workerSelect'); return w.options.length > 1 && w.options[0].value === ''; }, function(){
        document.getElementById('workerSelect').value = r.worker || '';
        SELECTED_WORKER_CODE = document.getElementById('workerSelect').value;
      });
    });
    showRequestModal(r.ctx.type, r.ctx.dept, r.ctx.workerLabel);
    if (r.ctx.type === 'holiday' && r.ctx.targetDate) {
      selectReqDate(r.ctx.targetDate);
      // 候補日のボタンが届いたら、選んだ日に印を付け直す
      when(function(){ return document.querySelectorAll('#reqDateButtons .req-date-btn').length > 0; }, function(){
        if (REQ_CTX.targetDate === r.ctx.targetDate) selectReqDate(r.ctx.targetDate);
      });
    }
    REQ_CTX.orders = (r.ctx.orders || ['', '', '']).slice(0, 3);
    while (REQ_CTX.orders.length < 3) REQ_CTX.orders.push('');
    updateReqOrderSlots();
    when(function(){ return ORDER_CACHE.length > 0; }, function(){ updateReqOrderSlots(); });
    if (r.retro) {
      toggleRetroPanel();
      var p = retroParseYmd_(r.retro);
      retroViewYear = p.getFullYear();
      retroViewMonth = p.getMonth();
      selectRetroDate(r.retro);
    }
    setVal('reqReason', r.reason);
    onReqReasonChange();
    setVal('reqReasonDetail', r.reasonDetail);
    setVal('reqHours', r.hours);
    setVal('reqWorkContent', r.workContent);
    if (r.step2) reqGoStep2();
  }

  function restoreEdit(e){
    function findItem(){
      return DASHBOARD_DATA.overtime.concat(DASHBOARD_DATA.holiday).some(function(it){ return it.requestId === e.requestId; });
    }
    when(findItem, function(){
      openEditModal(e.requestId);
      setVal('editHours', e.hours);
      if (EDIT_CTX.requestType === 'holiday') setVal('editTargetDate', e.targetDate);
    });
  }

  window.__vpsDraft.register({
    name: 'top',
    collect: function(){
      var out = {};
      if (shown('requestModalWrap') && REQ_CTX && REQ_CTX.type) {
        out.request = {
          ctx: JSON.parse(JSON.stringify(REQ_CTX)),
          retro: selectedRetroactiveDate,
          step2: shown('reqStep2'),
          reason: val('reqReason'), reasonDetail: val('reqReasonDetail'),
          hours: val('reqHours'), workContent: val('reqWorkContent'),
          dept: val('deptSelect'), worker: val('workerSelect')
        };
      }
      if (shown('editModalWrap') && EDIT_CTX && EDIT_CTX.requestId) {
        out.edit = { requestId: EDIT_CTX.requestId, hours: val('editHours'), targetDate: val('editTargetDate') };
      }
      return (out.request || out.edit) ? out : null;
    },
    restore: function(d){
      if (d.request) restoreRequest(d.request);
      else if (d.edit) restoreEdit(d.edit);
    }
  });
})();
