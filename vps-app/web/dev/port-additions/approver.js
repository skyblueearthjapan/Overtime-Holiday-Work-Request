
/* ===================================================================
 * VPS 版で足した部分（ここから下だけ。上は GAS 版 src/approver.html のまま）
 * =================================================================== */

// 操作中にログインが切れたとき、管理者が選んだ承認者を退避し、戻ったら選び直す（仕様 3-5）
(function(){
  if (!window.__vpsDraft) return;
  window.__vpsDraft.register({
    name: 'approver',
    collect: function(){
      if (!APPROVER_DATA || !APPROVER_DATA.isAdmin || !SELECTED_DEPTS || SELECTED_DEPTS.length === 0) return null;
      return { name: SELECTED_APPROVER_NAME, depts: SELECTED_DEPTS.slice() };
    },
    restore: function(d){
      if (!d || !d.depts || d.depts.length === 0) return;
      window.__vpsDraft.when(function(){ return APPROVER_DATA && APPROVER_DATA.isAdmin; }, function(){
        selectApprover(d.name, d.depts);
      });
    }
  });
})();
