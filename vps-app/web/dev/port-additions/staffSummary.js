
/* ===================================================================
 * VPS 版で足した部分（ここから下だけ。上は GAS 版 src/staffSummary.html のまま）
 * =================================================================== */

// 操作中にログインが切れたとき、月・部署・スタッフを退避し、戻ったら選び直す（仕様 3-5）
(function(){
  if (!window.__vpsDraft) return;
  window.__vpsDraft.register({
    name: 'staffSummary',
    collect: function(){
      return { ym: { y: state.ym.y, m: state.ym.m }, dept: state.dept, staff: state.staff };
    },
    restore: function(d){
      if (!d || !d.ym) return;
      window.__vpsDraft.when(function(){ return deptSel.options.length > 1; }, function(){
        if (!isFuture(d.ym)) state.ym = { y: Number(d.ym.y), m: Number(d.ym.m) };
        state.dept = d.dept || '';
        deptSel.value = state.dept;
        state.staff = d.staff || '';
        loadStaff(function(){ staffSel.value = state.staff; render(); });
      });
    }
  });
})();
