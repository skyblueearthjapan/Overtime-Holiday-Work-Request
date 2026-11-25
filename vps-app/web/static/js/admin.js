/* admin.js — GAS 版 src/admin.html の <script> をそのまま写したもの（仕様 3-6）。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
var __T = JSON.parse(document.getElementById('gas-template-data').textContent);
var STATE = {
  monthly: null,
  fiscal: null,
  sortKey: 'totalNet',
  sortDir: 'desc',
  visiblePeople: []
};

function z(n){ return String(n).padStart(2,'0'); }
function todayYMD(){
  var d=new Date();
  return d.getFullYear()+'-'+z(d.getMonth()+1)+'-'+z(d.getDate());
}
function todayYM(){
  var d=new Date();
  return d.getFullYear()+'-'+z(d.getMonth()+1);
}
function fmtMin(min){
  min = Math.max(0, Math.round(Number(min||0)));
  var h = Math.floor(min/60), r = min % 60;
  if(h===0) return r+'分';
  if(r===0) return h+'時間';
  return h+'時間'+r+'分';
}
function clamp(v,a,b){ return Math.max(a, Math.min(b, v)); }
function escHtml(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

// URLパラメータから date を取得（メールリンクからの遷移用）
var URL_DATE_PARAM = __T.DATE_PARAM;

document.getElementById('ym').value = todayYM();
document.getElementById('day').value = URL_DATE_PARAM || todayYMD();
document.getElementById('fyBase').value = todayYMD();

// 部署候補
google.script.run.withSuccessHandler(function(list){
  var sel=document.getElementById('dept');
  sel.innerHTML='';
  list.forEach(function(d){
    var opt=document.createElement('option');
    opt.value=d;
    opt.textContent=(d==='ALL'?'全体':d);
    sel.appendChild(opt);
  });
}).withFailureHandler(function(err){alert(err.message||err);}).api_adminDeptOptions();

function applyFiltersPeople(people){
  var q = document.getElementById('q').value.trim();
  var fOver40 = document.getElementById('fOver40').checked;
  var fOver60 = document.getElementById('fOver60').checked;
  var fProj40 = document.getElementById('fProj40').checked;
  var fPdfMiss = document.getElementById('fPdfMiss').checked;

  var out = people.slice();

  if(q) out = out.filter(function(p){ return (p.workerName||'').indexOf(q) >= 0; });
  if(fOver40) out = out.filter(function(p){ return p.totalNet >= 2400; });
  if(fOver60) out = out.filter(function(p){ return p.totalNet >= 3600; });
  if(fProj40) out = out.filter(function(p){ return p.projectedTotal >= 2400; });
  if(fPdfMiss) out = out.filter(function(p){ return (p.pdfMissing||0) > 0; });

  // sort
  var k = STATE.sortKey;
  var dir = STATE.sortDir;
  out.sort(function(a,b){
    var av = (k==='totalNetH') ? a.totalNet : (a[k] != null ? a[k] : '');
    var bv = (k==='totalNetH') ? b.totalNet : (b[k] != null ? b[k] : '');
    if(typeof av === 'number' && typeof bv === 'number'){
      return dir==='asc' ? av-bv : bv-av;
    }
    return dir==='asc'
      ? String(av).localeCompare(String(bv),'ja')
      : String(bv).localeCompare(String(av),'ja');
  });

  return out;
}

function renderMonthly(){
  var res = STATE.monthly;
  if(!res) return;

  // KPI
  document.getElementById('kpi40').innerText = res.kpi.over40;
  document.getElementById('kpi60').innerText = res.kpi.over60;
  document.getElementById('kpiP40').innerText = res.kpi.projectedOver40;
  document.getElementById('kpiPdf').innerText = res.kpi.pdfMissingTotal;

  var mi = res.monthInfo;
  document.getElementById('kpi40sub').innerText = '対象人数：'+res.kpi.totalPeople+' / 経過：'+mi.dayOfMonth+'/'+mi.daysInMonth+'日';
  document.getElementById('kpi60sub').innerText = '実績60h超は特別条項レベル';
  document.getElementById('kpiP40sub').innerText = '予測は"ペース換算"';

  // フィルタ＆ソート後の可視データ
  var visible = applyFiltersPeople(res.people);
  STATE.visiblePeople = visible;

  // Table
  var tb = document.querySelector('#tblMonthly tbody');
  tb.innerHTML='';
  visible.forEach(function(p){
    var remain40 = Math.max(0, 2400 - p.totalNet);

    var badge40 = remain40===0 ? '<span class="pill danger">超過</span>' :
      (remain40<=300 ? '<span class="pill warn">'+fmtMin(remain40)+'残</span>' : '<span class="pill">'+fmtMin(remain40)+'残</span>');
    var proj = p.projectedTotal >= 3600 ? '<span class="pill danger">60超予測</span>' :
                 p.projectedTotal >= 2400 ? '<span class="pill warn">40超予測</span>' :
                 '<span class="pill ok">範囲内</span>';

    var pdf = (p.pdfMissing||0) > 0 ? '<span class="pill warn">'+p.pdfMissing+'件</span>' : '<span class="pill ok">0</span>';

    var tr=document.createElement('tr');
    tr.innerHTML =
      '<td>'+escHtml(p.dept)+'</td>'+
      '<td><b>'+escHtml(p.workerName)+'</b></td>'+
      '<td>'+p.totalNet+'</td>'+
      '<td><b>'+fmtMin(p.totalNet)+'</b></td>'+
      '<td>'+fmtMin(p.overtimeNet)+'</td>'+
      '<td>'+fmtMin(p.holidayNet)+'</td>'+
      '<td>'+badge40+'</td>'+
      '<td>'+fmtMin(p.projectedTotal)+' '+proj+'</td>'+
      '<td>'+pdf+'</td>';
    tb.appendChild(tr);
  });

  // Watch
  var wb = document.querySelector('#tblWatch tbody');
  wb.innerHTML='';
  (res.watch||[]).forEach(function(w){
    var pdf = (w.pdfMissing||0) > 0 ? '<span class="pill warn">'+w.pdfMissing+'件</span>' : '<span class="pill ok">0</span>';
    var projBadge = w.projectedOver60 ? '<span class="pill danger">60超</span>' :
                    (w.projectedOver40 ? '<span class="pill warn">40超</span>' : '<span class="pill ok">範囲内</span>');
    var tr=document.createElement('tr');
    tr.innerHTML =
      '<td>'+escHtml(w.dept)+'</td>'+
      '<td><b>'+escHtml(w.workerName)+'</b></td>'+
      '<td><b>'+fmtMin(w.totalNet)+'</b></td>'+
      '<td>'+fmtMin(w.projectedTotal)+' '+projBadge+'</td>'+
      '<td>'+fmtMin(w.remain40)+'</td>'+
      '<td>'+fmtMin(w.remain60)+'</td>'+
      '<td>'+pdf+'</td>';
    wb.appendChild(tr);
  });

  // Charts
  drawPeopleBar('cPeople', visible, res.charts.people.limit40, res.charts.people.limit60);
  drawDeptPie('cDept', res.charts.dept.labels, res.charts.dept.values);

  document.getElementById('chartInfo').innerText =
    '棒：表示中 '+visible.length+'名（実績=濃、予測=薄）。40h/60hラインは固定。';
}

function loadMonthly(){
  var ym = document.getElementById('ym').value.trim();
  var dept = document.getElementById('dept').value;

  google.script.run.withSuccessHandler(function(res){
    STATE.monthly = res;
    renderMonthly();
  }).withFailureHandler(function(err){
    alert(err.message || err);
  }).api_adminMonthlySummary(ym, dept);
}

function loadDaily(){
  var day = document.getElementById('day').value.trim();
  var dept = document.getElementById('dept').value;

  google.script.run.withSuccessHandler(function(res){
    var tb = document.querySelector('#tblDaily tbody');
    tb.innerHTML='';
    res.items.forEach(function(it){
      var typeJa = it.requestType==='overtime'?'残業':'休日出勤';
      var pdf = it.pdfFileId
        ? '<a href="https://drive.google.com/file/d/'+escHtml(it.pdfFileId)+'/view" target="_blank" class="pill ok">開く</a>'
        : '<span class="pill warn">未作成</span>';

      // 操作列
      var action = '';
      if(it.status === 'approved' && !it.actualEndAt){
        action = '<button class="ghost" style="font-size:12px;padding:4px 10px;border-radius:8px;" '+
          'data-on-click="openManualEntry" data-args="'+__vpsArgs(it.requestId, it.workerName,
          it.requestType, res.date, it.actualStartAt||'')+'">手動入力</button>';
      } else if(it.actualEndAt){
        action = '<span class="pill ok">完了</span>';
      }

      // 過去申請バッジ
      var retroBadge = '';
      if(it.isRetroactive === true){
        var days = (typeof it.daysAgo === 'number' && it.daysAgo >= 0) ? it.daysAgo : null;
        var daysLabel = (days !== null) ? ' <span style="color:#7a1b1b;font-size:10px;font-weight:700;">('+days+'日前)</span>' : '';
        retroBadge = ' <span class="pill" style="background:#ffe5e5;color:#b91c1c;border:1px solid #fca5a5;font-size:11px;font-weight:800;">&#128281; 過去申請</span>' + daysLabel;
      }

      var tr=document.createElement('tr');
      if(it.isRetroactive === true){
        tr.style.background = 'rgba(255,235,150,0.35)';
      }
      tr.innerHTML=
        '<td>'+escHtml(it.dept)+'</td>'+
        '<td><b>'+escHtml(it.workerName)+'</b>'+retroBadge+'</td>'+
        '<td>'+typeJa+'</td>'+
        '<td>'+fmtMin(it.approvedMinutes)+'</td>'+
        '<td>'+fmtMin(it.actualMinutes)+'</td>'+
        '<td>'+fmtMin(it.breakMinutes)+'</td>'+
        '<td><b>'+fmtMin(it.netMinutes)+'</b></td>'+
        '<td>'+pdf+'</td>'+
        '<td>'+action+'</td>';
      tb.appendChild(tr);
    });
  }).withFailureHandler(function(err){
    alert(err.message || err);
  }).api_adminDailyDetail(day, dept);
}

// ====== 手動実績入力モーダル ======
var ME_REQUEST_ID = '';

function openManualEntry(requestId, workerName, requestType, dateYmd, actualStartAt){
  ME_REQUEST_ID = requestId;
  document.getElementById('meWorker').innerText = workerName;
  document.getElementById('meType').innerText = requestType === 'overtime' ? '残業' : '休日出勤';
  document.getElementById('meDate').innerText = dateYmd;

  // 開始時刻プリセット
  var startInput = document.getElementById('meStart');
  if(requestType === 'overtime'){
    startInput.value = '17:20';
  } else if(actualStartAt){
    // "yyyy-MM-dd HH:mm:ss" → HH:mm
    var m = String(actualStartAt).match(/(\d{2}:\d{2})/);
    startInput.value = m ? m[1] : '';
  } else {
    startInput.value = '';
  }
  document.getElementById('meEnd').value = '';
  document.getElementById('manualEntryWrap').style.display = 'block';
}

function closeManualEntry(){
  document.getElementById('manualEntryWrap').style.display = 'none';
  ME_REQUEST_ID = '';
}

function submitManualEntry(){
  var startTime = document.getElementById('meStart').value;
  var endTime = document.getElementById('meEnd').value;
  if(!endTime){
    alert('終了時刻を入力してください。');
    return;
  }
  if(!startTime){
    alert('開始時刻を入力してください。');
    return;
  }
  if(!confirm('手動実績を登録しますか？\n開始: '+startTime+' → 終了: '+endTime)) return;

  google.script.run.withSuccessHandler(function(res){
    closeManualEntry();
    alert('登録完了\n実働: '+fmtMin(res.actualMinutes)+'\n休憩控除: '+fmtMin(res.breakMinutes)+'\n実残業/実働: '+fmtMin(res.netMinutes));
    loadDaily();
  }).withFailureHandler(function(err){
    alert('エラー: '+(err.message || err));
  }).api_adminManualComplete(ME_REQUEST_ID, endTime, startTime);
}

function loadFiscal(){
  var base = document.getElementById('fyBase').value.trim();
  var dept = document.getElementById('dept').value;

  google.script.run.withSuccessHandler(function(res){
    STATE.fiscal = res;
    drawLine('cFiscal', res.labels, res.values, 2400);
    document.getElementById('fyInfo').innerText =
      '年度：'+res.fiscalYearStart+' 〜 '+res.fiscalYearEnd+'（'+(dept==='ALL'?'全体':dept)+'）';
  }).withFailureHandler(function(err){
    alert(err.message || err);
  }).api_adminFiscalTrend(base, dept);
}

// ====== Sorting ======
document.querySelectorAll('#tblMonthly thead th').forEach(function(th){
  th.addEventListener('click', function(){
    var k = th.dataset.k;
    if(!k) return;
    if(STATE.sortKey === k){
      STATE.sortDir = (STATE.sortDir === 'asc') ? 'desc' : 'asc';
    } else {
      STATE.sortKey = k;
      STATE.sortDir = 'desc';
    }
    renderMonthly();
  });
});

// ====== Export visible to CSV ======
function exportVisible(){
  var rows = [];
  rows.push(['部署','氏名','合算(分)','合算','残業','休日','40残(分)','予測(分)','PDF未作成']);
  STATE.visiblePeople.forEach(function(p){
    rows.push([
      p.dept,
      p.workerName,
      p.totalNet,
      fmtMin(p.totalNet),
      fmtMin(p.overtimeNet),
      fmtMin(p.holidayNet),
      Math.max(0,2400-p.totalNet),
      p.projectedTotal,
      p.pdfMissing||0
    ]);
  });
  var csv = rows.map(function(r){return r.map(function(v){
    var s=String(v!=null?v:'');
    return /[",\n]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s;
  }).join(',');}).join('\n');

  var blob = new Blob([csv], {type:'text/csv'});
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = '総務_月次_'+document.getElementById('ym').value+'_'+document.getElementById('dept').value+'.csv';
  a.click();
  URL.revokeObjectURL(url);
}

// ====== CANVAS CHARTS (lightweight, no external libs) ======
function clearCanvas(id){
  var c=document.getElementById(id);
  var ctx=c.getContext('2d');
  var rect=c.getBoundingClientRect();
  c.width = Math.floor(rect.width * devicePixelRatio);
  c.height = Math.floor(rect.height * devicePixelRatio);
  ctx.scale(devicePixelRatio, devicePixelRatio);
  ctx.clearRect(0,0,rect.width,rect.height);
  return {c:c,ctx:ctx,w:rect.width,h:rect.height};
}

function drawPeopleBar(canvasId, people, limit40, limit60){
  var o = clearCanvas(canvasId);
  var ctx=o.ctx, w=o.w, h=o.h;
  var padL=44, padR=14, padT=14, padB=46;
  var innerW = w - padL - padR;
  var innerH = h - padT - padB;

  var maxVal = Math.max(limit60, 4000);
  for(var i=0;i<people.length;i++){
    if(people[i].totalNet > maxVal) maxVal = people[i].totalNet;
    if((people[i].projectedTotal||0) > maxVal) maxVal = people[i].projectedTotal;
  }
  var n = Math.max(1, people.length);
  var gap = 6;
  var barW = clamp((innerW - gap*(n-1)) / n, 6, 20);

  // grid lines
  ctx.font = '12px system-ui';
  ctx.fillStyle = '#5c6b7a';
  ctx.strokeStyle = '#d7e7f5';

  var ticks = [0, limit40, limit60, maxVal];
  ticks.forEach(function(t){
    var y = padT + innerH - (t/maxVal)*innerH;
    ctx.beginPath(); ctx.moveTo(padL,y); ctx.lineTo(padL+innerW,y); ctx.stroke();
    ctx.fillText(Math.round(t/60)+'h', 8, y+4);
  });

  // limit lines emphasized
  function drawLimitLine(yVal, label){
    ctx.save();
    var y = padT + innerH - (yVal/maxVal)*innerH;
    ctx.strokeStyle = '#1f9d55';
    ctx.setLineDash([6,4]);
    ctx.beginPath(); ctx.moveTo(padL,y); ctx.lineTo(padL+innerW,y); ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = '#1f9d55';
    ctx.fillText(label, padL+6, y-6);
    ctx.restore();
  }
  drawLimitLine(limit40, '40h');
  drawLimitLine(limit60, '60h');

  // bars
  people.forEach(function(p,i){
    var x = padL + i*(barW+gap);
    var v = p.totalNet;
    var pv = p.projectedTotal || 0;

    // projected (light)
    var ph = (pv/maxVal)*innerH;
    ctx.fillStyle = 'rgba(14,165,233,0.18)';
    ctx.fillRect(x, padT + innerH - ph, barW, ph);

    // actual (solid)
    var ah = (v/maxVal)*innerH;
    var over60 = v>=3600;
    var over40 = v>=2400;

    ctx.fillStyle = over60 ? 'rgba(255,90,95,0.68)' : (over40 ? 'rgba(255,176,32,0.68)' : 'rgba(31,157,85,0.62)');
    ctx.fillRect(x, padT + innerH - ah, barW, ah);

    // x label
    if(n<=40){
      var name = (p.workerName||'').slice(0,6);
      ctx.save();
      ctx.translate(x+barW/2, padT+innerH+22);
      ctx.rotate(-0.7);
      ctx.fillStyle='#5c6b7a';
      ctx.fillText(name, -10, 0);
      ctx.restore();
    }
  });

  // axis frame
  ctx.strokeStyle = '#d7e7f5';
  ctx.strokeRect(padL, padT, innerW, innerH);
}

function drawDeptPie(canvasId, labels, values){
  var o = clearCanvas(canvasId);
  var ctx=o.ctx, w=o.w, h=o.h;
  var cx=w/2, cy=h/2, r=Math.min(w,h)*0.32;
  var total = 0;
  for(var i=0;i<values.length;i++) total += (values[i]||0);
  if(total===0) total=1;

  var start = -Math.PI/2;
  labels.forEach(function(lab,i){
    var v=values[i]||0;
    var ang = (v/total)*Math.PI*2;
    var end = start + ang;

    var hue = (hashCode(lab) % 360 + 360) % 360;
    ctx.fillStyle = 'hsla('+hue+', 65%, 55%, 0.55)';

    ctx.beginPath();
    ctx.moveTo(cx,cy);
    ctx.arc(cx,cy,r,start,end);
    ctx.closePath();
    ctx.fill();

    start=end;
  });

  // legend
  ctx.font='12px system-ui';
  var y=18;
  labels.slice(0,10).forEach(function(lab,i){
    var hue = (hashCode(lab) % 360 + 360) % 360;
    ctx.fillStyle = 'hsla('+hue+', 65%, 55%, 0.55)';
    ctx.fillRect(10,y-10,10,10);
    ctx.fillStyle='#5c6b7a';
    var pct = Math.round((values[i]/total)*100);
    ctx.fillText(lab+' '+pct+'%', 26, y);
    y += 16;
  });
}

function drawLine(canvasId, labels, values, hintLineMin){
  var o = clearCanvas(canvasId);
  var ctx=o.ctx, w=o.w, h=o.h;
  var padL=44, padR=14, padT=14, padB=32;
  var innerW=w-padL-padR;
  var innerH=h-padT-padB;

  var maxVal = Math.max(hintLineMin||0, 3000);
  for(var i=0;i<values.length;i++){
    if(values[i]>maxVal) maxVal=values[i];
  }
  var n = Math.max(1, labels.length);

  // grid
  ctx.font='12px system-ui';
  ctx.fillStyle='#5c6b7a';
  ctx.strokeStyle='#d7e7f5';

  var ticks=[0, hintLineMin||0, maxVal];
  ticks.forEach(function(t){
    var y = padT + innerH - (t/maxVal)*innerH;
    ctx.beginPath(); ctx.moveTo(padL,y); ctx.lineTo(padL+innerW,y); ctx.stroke();
    ctx.fillText(Math.round(t/60)+'h', 8, y+4);
  });

  // line
  ctx.strokeStyle='rgba(31,157,85,0.75)';
  ctx.lineWidth=2;
  ctx.beginPath();
  values.forEach(function(v,i){
    var x = padL + (i/(n-1||1))*innerW;
    var y = padT + innerH - (v/maxVal)*innerH;
    if(i===0) ctx.moveTo(x,y);
    else ctx.lineTo(x,y);
  });
  ctx.stroke();

  // points
  ctx.fillStyle='rgba(31,157,85,0.75)';
  values.forEach(function(v,i){
    var x = padL + (i/(n-1||1))*innerW;
    var y = padT + innerH - (v/maxVal)*innerH;
    ctx.beginPath(); ctx.arc(x,y,3,0,Math.PI*2); ctx.fill();
  });

  // x labels
  ctx.fillStyle='#5c6b7a';
  labels.forEach(function(lab,i){
    var x = padL + (i/(n-1||1))*innerW;
    if(i%2===0){
      ctx.fillText(lab.slice(5), x-8, padT+innerH+22);
    }
  });

  // frame
  ctx.strokeStyle='#d7e7f5';
  ctx.strokeRect(padL,padT,innerW,innerH);
}

function hashCode(str){
  var h=0;
  for(var i=0;i<str.length;i++) h = ((h<<5)-h) + str.charCodeAt(i) | 0;
  return h;
}

// ====== Special Clause (60h) ======
function loadSpecial(){
  var base = document.getElementById('fyBase').value.trim();
  var dept = document.getElementById('dept').value;

  google.script.run.withSuccessHandler(function(res){
    var tb=document.querySelector('#tblSpecial tbody');
    tb.innerHTML='';
    if(!res.list || res.list.length===0){
      tb.innerHTML='<tr><td colspan="4" class="muted">60h超の月はありません</td></tr>';
      return;
    }
    res.list.forEach(function(r){
      var status = r.isDanger ? '<span class="pill danger">6回到達</span>'
                   : r.isWarn ? '<span class="pill warn">5回</span>'
                   : '<span class="pill ok">'+r.count60+'回</span>';
      var tr=document.createElement('tr');
      tr.innerHTML=
        '<td>'+escHtml(r.dept)+'</td>'+
        '<td><b>'+escHtml(r.workerName)+'</b></td>'+
        '<td>'+r.count60+'</td>'+
        '<td>'+status+'</td>';
      tb.appendChild(tr);
    });
  }).withFailureHandler(function(err){
    alert(err.message || err);
  }).api_adminSpecialClauseCount(base, dept);
}

// ====== Pending Watch ======
function loadPending(){
  google.script.run.withSuccessHandler(function(list){
    var tb=document.querySelector('#tblPending tbody');
    tb.innerHTML='';
    if(!list || list.length===0){
      tb.innerHTML='<tr><td colspan="6" class="muted">未承認の滞留はありません</td></tr>';
      return;
    }
    list.forEach(function(r){
      var status = r.isOver48h
        ? '<span class="pill danger">48h超</span>'
        : '<span class="pill warn">滞留中</span>';
      var typeJa = r.requestType === 'holiday' ? '休日出勤' : (r.requestType === 'overtime' ? '残業' : escHtml(r.requestType || ''));
      // 対象日を過ぎた未承認は赤字＋⚠（承認漏れ）
      var dateLabel = escHtml(r.targetDateLabel || r.targetDate || '');
      var dateCell = r.isPastDate
        ? '<span style="color:var(--danger);font-weight:700;white-space:nowrap;">&#9888; '+dateLabel+'</span>'
        : '<span style="white-space:nowrap;">'+dateLabel+'</span>';
      var tr=document.createElement('tr');
      tr.innerHTML=
        '<td>'+dateCell+'</td>'+
        '<td>'+typeJa+'</td>'+
        '<td>'+escHtml(r.dept)+'</td>'+
        '<td><b>'+escHtml(r.workerName)+'</b></td>'+
        '<td>'+r.hoursPending+'時間</td>'+
        '<td>'+status+'</td>';
      tb.appendChild(tr);
    });
  }).withFailureHandler(function(err){
    alert(err.message || err);
  }).api_adminPendingWatch();
}

// Initial load
loadMonthly();
loadFiscal();
loadSpecial();
loadPending();
// メールリンクからの遷移時は日次詳細を自動ロード
if(URL_DATE_PARAM){
  loadDaily();
}

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
