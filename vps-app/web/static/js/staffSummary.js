/* staffSummary.js — GAS 版 src/staffSummary.html の <script> をそのまま写したもの（仕様 3-6）。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
  var deptSel = document.getElementById('deptSel');
  var staffSel = document.getElementById('staffSel');
  var resultBody = document.getElementById('resultBody');

  // ---- 締め月の状態 ----
  function curClosingYm(){
    var d = new Date(), y = d.getFullYear(), m = d.getMonth()+1;
    if(d.getDate() <= 15){ m -= 1; if(m===0){ m=12; y-=1; } }
    return { y:y, m:m };
  }
  var CUR = curClosingYm();
  var state = { ym:{ y:CUR.y, m:CUR.m }, dept:'', staff:'' };

  function ymStr(o){ return o.y + '-' + ('0'+o.m).slice(-2); }
  function addMonth(o, delta){ var idx = o.y*12 + (o.m-1) + delta; return { y:Math.floor(idx/12), m:(idx%12)+1 }; }
  function isFuture(o){ return (o.y*12 + o.m) > (CUR.y*12 + CUR.m); }
  function monthLabel(o){ return o.y + '年 ' + o.m + '月分'; }
  function periodStr(o){ var n = addMonth(o,1); return o.m + '/16 〜 ' + n.m + '/15'; }

  function setMonthUI(){
    document.getElementById('monthM').textContent = monthLabel(state.ym);
    document.getElementById('monthP').textContent = '締め期間 ' + periodStr(state.ym);
    document.getElementById('periodLabel').textContent = periodStr(state.ym);
  }

  function fmtMin(min){ min=Math.round(min); var h=Math.floor(min/60), m=min%60;
    if(h===0) return m+'分'; if(m===0) return h+'時間'; return h+'時間'+m+'分'; }
  function hm(min){ min=Math.round(min); var h=Math.floor(min/60), m=min%60; return h+':'+('0'+m).slice(-2); }
  function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g, function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }

  // ---- セレクタ ----
  function renderDeptOptions(depts){
    deptSel.innerHTML = '<option value="">全部署</option>' +
      depts.map(function(d){ return '<option value="'+esc(d)+'">'+esc(d)+'</option>'; }).join('');
    deptSel.value = state.dept;
  }
  function renderStaffOptions(list){
    var opts = '<option value="">全員</option>';
    for(var i=0;i<list.length;i++){
      var tok = String(list[i]); var sp = tok.indexOf(' ');
      var name = sp>=0 ? tok.slice(sp+1) : tok;
      opts += '<option value="'+esc(tok)+'">'+esc(name)+'</option>';
    }
    staffSel.innerHTML = opts;
    staffSel.value = state.staff;
  }
  function loadStaff(cb){
    if(!state.dept){
      state.staff=''; staffSel.innerHTML='<option value="">全員</option>'; staffSel.value='';
      if(cb) cb(); return;
    }
    google.script.run
      .withSuccessHandler(function(list){
        list = list || []; renderStaffOptions(list);
        var ok=false; for(var i=0;i<list.length;i++){ if(String(list[i])===state.staff){ ok=true; break; } }
        if(!ok){ state.staff=''; staffSel.value=''; }
        if(cb) cb();
      })
      .withFailureHandler(function(){ state.staff=''; staffSel.innerHTML='<option value="">全員</option>'; if(cb) cb(); })
      .api_getWorkersByDept(state.dept);
  }

  // ---- 検索・描画 ----
  function render(){
    setMonthUI();
    resultBody.classList.remove('fade'); void resultBody.offsetWidth; resultBody.classList.add('fade');
    resultBody.innerHTML = '<div class="loading">読み込み中...</div>';
    google.script.run
      .withSuccessHandler(function(res){
        if(!res){ resultBody.innerHTML = '<div class="loading">データがありません</div>'; return; }
        resultBody.innerHTML = (res.mode==='personal') ? renderPerson(res) : renderList(res);
      })
      .withFailureHandler(function(err){
        resultBody.innerHTML = '<div class="loading">読み込みに失敗しました：' + esc(err && err.message || err) + '</div>';
      })
      .api_getClosingSummary(ymStr(state.ym), state.dept, state.staff);
  }

  function renderPerson(res){
    var w=res.worker, k=res.kpi, meta=res.meta;
    var initials = (w.workerName||'—').slice(0,1);
    var LIM40 = meta.limit40, LIM60 = meta.limit60, MAX = 4200;
    var total = k.totalNet;
    var pct = Math.min(100, Math.round((total/MAX)*1000)/10);
    var p40 = (LIM40/MAX*100).toFixed(2);
    var p60 = (LIM60/MAX*100).toFixed(2);
    var zoneBg = 'linear-gradient(90deg,#e9f7ef 0 '+p40+'%,#fff4e2 '+p40+'% '+p60+'%,#ffeceb '+p60+'% 100%)';
    var over = total>LIM60 ? 'danger' : (total>LIM40 ? 'warn' : 'ok');
    var overTxt = total>LIM60 ? '60時間 超過' : (total>LIM40 ? '40時間 超（年6回のうち）' : '40時間内');
    var yr = res.annualOver, rest = 6 - yr;
    var dots=''; for(var di=0; di<6; di++){ dots += '<span class="'+(di<yr?'on':'off')+'">●</span>'; }

    var rows = res.detail.map(function(r){
      var pdf = r.pdfFileId
        ? '<a class="pdflink" target="_blank" href="https://drive.google.com/file/d/'+esc(r.pdfFileId)+'/view">PDF</a>'
        : '<span class="pdfnone">—</span>';
      return '<tr>'+
        '<td class="tabnum">'+esc(r.date)+'</td>'+
        '<td><span class="typeChip '+(r.type==='overtime'?'ot':'hol')+'">'+(r.type==='overtime'?'残業':'休日出勤')+'</span></td>'+
        '<td class="num">'+(r.approvedMinutes?hm(r.approvedMinutes):'—')+'</td>'+
        '<td class="num">'+(r.netMinutes?hm(r.netMinutes):'—')+'</td>'+
        '<td class="num">'+(r.breakMinutes?hm(r.breakMinutes):'—')+'</td>'+
        '<td>'+(r.status==='approved'?'<span class="pill ok">承認済</span>':'<span class="pill warn">申請中</span>')+'</td>'+
        '<td>'+pdf+'</td>'+
      '</tr>';
    }).join('');
    if(!rows) rows = '<tr><td colspan="7" class="loading">この月のデータはありません</td></tr>';

    var pendTxt = k.pendMin ? ('＋'+fmtMin(k.pendMin)) : '—';

    return ''+
    '<div class="targetHead">'+
      '<div class="avatar">'+esc(initials)+'</div>'+
      '<div><div class="nm">'+esc(w.workerName||'—')+'</div><div class="dp">'+esc(w.dept||'')+'　/　'+esc(meta.monthLabel)+'</div></div>'+
      '<div style="flex:1"></div>'+
      '<span class="pill '+over+'">'+overTxt+'</span>'+
    '</div>'+
    '<div class="kpiGrid">'+
      '<div class="kpi total"><div class="k-label">合計（実働）</div><div class="k-val tabnum">'+fmtMin(total)+'</div><div class="k-sub">承認済 ・ 上限40時間（超過は年6回・60時間まで）</div></div>'+
      '<div class="kpi ot"><span class="rail"></span><div class="k-label">残業</div><div class="k-val tabnum">'+fmtMin(k.otNet)+'</div><div class="k-sub">netMinutes</div></div>'+
      '<div class="kpi hol"><span class="rail"></span><div class="k-label">休日出勤</div><div class="k-val tabnum">'+fmtMin(k.holNet)+'</div><div class="k-sub">netMinutes</div></div>'+
      '<div class="kpi pend"><span class="rail"></span><div class="k-label">見込み（実績待ち/申請中）</div><div class="k-val tabnum">'+pendTxt+'</div><div class="k-sub">承認予定・参考値</div></div>'+
    '</div>'+
    '<div class="gauge">'+
      '<div class="g-top"><span>合計（残業＋休日出勤）の上限基準</span><span><span class="big tabnum">'+fmtMin(total)+'</span> ／ 上限 40時間</span></div>'+
      '<div class="gaugeBar">'+
        '<span class="gl g40" style="left:'+p40+'%">40時間</span>'+
        '<span class="gl g60" style="left:'+p60+'%">60時間</span>'+
        '<div class="track" style="background:'+zoneBg+'"><div class="fill '+over+'" style="width:'+pct+'%"></div></div>'+
        '<div class="line l40" style="left:'+p40+'%"></div>'+
        '<div class="line l60" style="left:'+p60+'%"></div>'+
        '<div class="curtick" style="left:'+pct+'%">▲ '+hm(total)+'</div>'+
      '</div>'+
      '<div class="g-legend">'+
        '<span><i class="sw green"></i>40時間まで（上限内）</span>'+
        '<span><i class="sw amber"></i>40〜60時間（年6回まで）</span>'+
        '<span><i class="sw red"></i>60時間 超過（不可）</span>'+
      '</div>'+
      '<div class="annual">'+
        '<span class="lbl">'+esc(meta.fyLabel)+' 40時間超：</span>'+
        '<span class="dots">'+dots+'</span>'+
        '<span><span class="cnt">'+yr+'</span> / 6 回</span>'+
        '<span class="rest'+(rest<=0?' zero':'')+'">残り '+Math.max(0,rest)+' 回</span>'+
      '</div>'+
    '</div>'+
    '<div class="tableWrap"><table>'+
      '<thead><tr><th>対象日</th><th>種別</th><th class="num">承認</th><th class="num">実績(net)</th><th class="num">休憩</th><th>状態</th><th>個票</th></tr></thead>'+
      '<tbody>'+rows+'</tbody>'+
    '</table></div>';
  }

  function renderList(res){
    var people = res.people, meta = res.meta;
    var max = 1;
    for(var i=0;i<people.length;i++){ if(people[i].totalNet>max) max = people[i].totalNet; }
    var sum = { ot:0, hol:0, tot:0, pend:0 };
    var rows = people.map(function(p){
      var w = Math.round((p.totalNet/max)*100);
      sum.ot+=p.otNet; sum.hol+=p.holNet; sum.tot+=p.totalNet; sum.pend+=p.pendMin;
      var flag = p.totalNet>meta.limit60 ? '<span class="pill danger">60時間超</span>'
               : (p.totalNet>meta.limit40 ? '<span class="pill warn">40時間超</span>'
               : '<span class="pill ok">上限内</span>');
      flag += ' <span class="yrcnt">年'+p.annualOver+'/6</span>';
      return '<tr class="rowlink" data-dept="'+esc(p.dept)+'" data-code="'+esc(p.workerCode)+'" data-name="'+esc(p.workerName)+'" title="クリックで個人詳細を表示">'+
        '<td><div class="rowname">'+esc(p.workerName)+'</div><div class="rowdept">'+esc(p.dept)+'</div></td>'+
        '<td class="num tabnum">'+hm(p.otNet)+'</td>'+
        '<td class="num tabnum">'+hm(p.holNet)+'</td>'+
        '<td class="num"><div class="mini"><span class="tabnum">'+hm(p.totalNet)+'</span><span class="minibar"><span style="width:'+w+'%"></span></span></div></td>'+
        '<td class="num tabnum">'+(p.pendMin?'＋'+hm(p.pendMin):'—')+'</td>'+
        '<td>'+flag+'</td>'+
      '</tr>';
    }).join('');
    if(!rows) rows = '<tr><td colspan="6" class="loading">この条件のデータはありません</td></tr>';

    var scope = res.dept ? res.dept : '全部署';
    return ''+
    '<div class="sectionTitle" style="margin-bottom:8px">'+
      '<div class="lead"><span>スタッフ別 一覧 — '+esc(scope)+'</span></div>'+
      '<div class="section-count">'+people.length+'名 ・ '+esc(meta.monthLabel)+'</div>'+
    '</div>'+
    '<div class="tableWrap"><table>'+
      '<thead><tr><th>氏名 / 部署</th><th class="num">残業</th><th class="num">休日出勤</th><th class="num">合計（実働）▼</th><th class="num">見込み</th><th>状況</th></tr></thead>'+
      '<tbody>'+rows+'</tbody>'+
      '<tfoot><tr class="totrow">'+
        '<td>合計（'+people.length+'名）</td>'+
        '<td class="num tabnum">'+hm(sum.ot)+'</td>'+
        '<td class="num tabnum">'+hm(sum.hol)+'</td>'+
        '<td class="num tabnum">'+hm(sum.tot)+'</td>'+
        '<td class="num tabnum">'+(sum.pend?'＋'+hm(sum.pend):'—')+'</td>'+
        '<td></td>'+
      '</tr></tfoot>'+
    '</table></div>'+
    '<div class="hint" style="margin-top:12px">💡 行をクリックすると、その人の部署・氏名に絞り込みが切り替わり、個人詳細（日別明細）が開きます。</div>';
  }

  // 一覧の行クリック → その人の部署・氏名に切替えて個人詳細へ
  function openPerson(dept, code, name){
    state.dept = dept || '';
    deptSel.value = state.dept;
    if(!state.dept){ state.staff=''; render(); return; }
    google.script.run
      .withSuccessHandler(function(list){
        list = list || [];
        var tok = '';
        for(var i=0;i<list.length;i++){
          var t = String(list[i]); var c = t.split(/\s+/)[0];
          if((code && c === code) || (!code && name && t.indexOf(name) >= 0)){ tok = t; break; }
        }
        state.staff = tok;
        renderStaffOptions(list);
        render();
        try { window.scrollTo({ top:0, behavior:'smooth' }); } catch(e){ window.scrollTo(0,0); }
      })
      .withFailureHandler(function(){ state.staff=''; renderStaffOptions([]); render(); })
      .api_getWorkersByDept(state.dept);
  }

  // ---- イベント ----
  resultBody.addEventListener('click', function(e){
    var el = e.target;
    var tr = (el && el.closest) ? el.closest('tr.rowlink') : null;
    if(!tr) return;
    openPerson(tr.getAttribute('data-dept'), tr.getAttribute('data-code'), tr.getAttribute('data-name'));
  });
  document.getElementById('prevM').addEventListener('click', function(){ state.ym = addMonth(state.ym,-1); render(); });
  document.getElementById('nextM').addEventListener('click', function(){ var n = addMonth(state.ym,1); if(!isFuture(n)){ state.ym = n; render(); } });
  deptSel.addEventListener('change', function(){ state.dept = deptSel.value; loadStaff(function(){ render(); }); });
  staffSel.addEventListener('change', function(){ state.staff = staffSel.value; render(); });

  // ---- 初期化：部署一覧→本人（初期選択）→描画 ----
  function init(){
    google.script.run
      .withSuccessHandler(function(depts){
        depts = depts || []; renderDeptOptions(depts);
        google.script.run
          .withSuccessHandler(function(me){
            if(me && me.workerName){ document.getElementById('whoName').textContent = me.workerName + '（' + (me.dept||'') + '）'; }
            if(me && me.dept){
              state.dept = me.dept; deptSel.value = me.dept;
              google.script.run
                .withSuccessHandler(function(list){
                  list = list || []; renderStaffOptions(list);
                  var tok = '';
                  for(var i=0;i<list.length;i++){ if(String(list[i]).split(/\s+/)[0] === me.workerCode){ tok = String(list[i]); break; } }
                  state.staff = tok; staffSel.value = tok;
                  render();
                })
                .withFailureHandler(function(){ render(); })
                .api_getWorkersByDept(me.dept);
            } else { render(); }
          })
          .withFailureHandler(function(){ render(); })
          .api_getWorkerInfo();
      })
      .withFailureHandler(function(){ renderDeptOptions([]); render(); })
      .api_getDeptList();
  }
  init();

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
