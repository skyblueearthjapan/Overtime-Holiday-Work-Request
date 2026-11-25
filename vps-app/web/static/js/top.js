/* top.js — GAS 版 src/top.html の <script> をそのまま写したもの（仕様 3-6）。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
var __T = JSON.parse(document.getElementById('gas-template-data').textContent);
var REASONS = __T.REASONS;
var OT_HOURS_LIST = __T.OT_HOURS;
var HD_HOURS_LIST = __T.HD_HOURS;

var SELECTED_DEPT = '';
var SELECTED_WORKER_CODE = '';

function fmtMin(min){
  min = Math.max(0, Math.round(Number(min||0)));
  var h = Math.floor(min/60), r = min % 60;
  if(h===0) return r+'分';
  if(r===0) return h+'時間';
  return h+'時間'+r+'分';
}

// ====== 初期ロード：部署一覧 ======
google.script.run
  .withSuccessHandler(function(depts){
    var sel = document.getElementById('deptSelect');
    sel.innerHTML = '<option value="">-- 部署を選択 --</option>';
    (depts || []).forEach(function(d){
      var opt = document.createElement('option');
      opt.value = d;
      opt.textContent = d;
      sel.appendChild(opt);
    });
    // 部署一覧ロード後にユーザー自動判定
    google.script.run
      .withSuccessHandler(function(worker){
        if(worker && worker.dept){
          sel.value = worker.dept;
          onDeptChange();
          // 作業員一覧ロード完了後に本人を自動選択
          var checkWorker = setInterval(function(){
            var wSel = document.getElementById('workerSelect');
            if(wSel.options.length > 1 && wSel.options[0].value !== ''){
              clearInterval(checkWorker);
              if(worker.workerCode){
                wSel.value = worker.workerCode;
                SELECTED_WORKER_CODE = worker.workerCode;
              }
            }
          }, 200);
          // 5秒でタイムアウト
          setTimeout(function(){ clearInterval(checkWorker); }, 5000);
        }
      })
      .withFailureHandler(function(){})
      .api_getWorkerInfo();
  })
  .withFailureHandler(function(){})
  .api_getDeptList();

// ====== 初期ロード：全社ダッシュボード ======
loadDashboard();

// ====== 部署変更時（申請フォーム用） ======
document.getElementById('deptSelect').addEventListener('change', onDeptChange);

function onDeptChange(){
  var dept = document.getElementById('deptSelect').value;
  SELECTED_DEPT = dept;
  SELECTED_WORKER_CODE = '';

  var wSel = document.getElementById('workerSelect');
  wSel.innerHTML = '<option value="">読込中...</option>';

  if(!dept) {
    wSel.innerHTML = '<option value="">-- 部署を先に選択 --</option>';
    return;
  }

  google.script.run
    .withSuccessHandler(function(workers){
      wSel.innerHTML = '<option value="">-- 作業員を選択 --</option>';
      (workers || []).forEach(function(w){
        var code = w.split(/\s+/)[0] || '';
        var opt = document.createElement('option');
        opt.value = code;
        opt.textContent = w;
        wSel.appendChild(opt);
      });
    })
    .withFailureHandler(function(){
      wSel.innerHTML = '<option value="">取得エラー</option>';
    })
    .api_getWorkersByDept(dept);
}

// ====== 作業員変更時（申請フォーム用） ======
document.getElementById('workerSelect').addEventListener('change', function(){
  var wSel = document.getElementById('workerSelect');
  SELECTED_WORKER_CODE = wSel.value;
});

// ====== 申請ボタン → 新モーダルフロー ======
function openForm(type){
  var dept = SELECTED_DEPT || document.getElementById('deptSelect').value;
  if(!dept){
    alert('部署を選択してください。');
    return;
  }

  var wSel = document.getElementById('workerSelect');
  var workerLabel = '';
  if(wSel.selectedIndex > 0){
    workerLabel = wSel.options[wSel.selectedIndex].textContent;
  }
  if(!workerLabel){
    alert('作業員を選択してください。');
    return;
  }

  showRequestModal(type, dept, workerLabel);
}

// ====== 休日日付選択モーダル ======
var PICKER_DEPT = '';
var PICKER_WORKER = '';

function showDatePicker(dept, workerLabel){
  PICKER_DEPT = dept;
  PICKER_WORKER = workerLabel;
  var wrap = document.getElementById('datePickerWrap');
  var btns = document.getElementById('dateButtons');
  btns.innerHTML = '<div class="loading">候補日を取得中...</div>';
  document.getElementById('manualDateInput').value = '';
  wrap.style.display = 'block';

  google.script.run
    .withSuccessHandler(function(candidates){
      btns.innerHTML = '';
      if(!candidates || candidates.length === 0){
        btns.innerHTML = '<div class="muted">候補日が見つかりませんでした</div>';
        return;
      }
      candidates.forEach(function(c){
        var btn = document.createElement('button');
        if(c.type === 'holiday'){
          btn.className = 'btn-warn';
        }
        btn.style.cssText = 'width:100%; padding:14px; font-size:15px; border-radius:12px;';
        var label = c.label;
        if(c.holidayName) label += ' (' + c.holidayName + ')';
        btn.textContent = label;
        btn.onclick = function(){
          closeDatePicker();
          showOrderPicker('holiday', PICKER_DEPT, PICKER_WORKER, c.date);
        };
        btns.appendChild(btn);
      });
    })
    .withFailureHandler(function(err){
      btns.innerHTML = '<div class="muted">取得エラー: ' + escHtml(err.message || err) + '</div>';
    })
    .api_getHolidayCandidateDates();
}

function closeDatePicker(){
  document.getElementById('datePickerWrap').style.display = 'none';
}

function onManualDateSelect(){
  var val = document.getElementById('manualDateInput').value;
  if(!val){
    alert('日付を選択してください。');
    return;
  }
  closeDatePicker();
  showOrderPicker('holiday', PICKER_DEPT, PICKER_WORKER, val);
}

// ====== 工番選択モーダル ======
var ORDER_CACHE = [];
var ORDER_SELECTED = ['','',''];
var ORDER_CONTEXT = {};

function showOrderPicker(type, dept, workerLabel, targetDateStr){
  ORDER_CONTEXT = { type: type, dept: dept, workerLabel: workerLabel, targetDateStr: targetDateStr };
  ORDER_SELECTED = ['','',''];
  updateOrderSlots();
  document.getElementById('orderSearchInput').value = '';
  document.getElementById('orderPickerWrap').style.display = 'block';

  if(ORDER_CACHE.length > 0){
    filterOrders();
    return;
  }

  document.getElementById('orderResults').innerHTML =
    '<div class="loading" id="orderLoading">工番マスタを読込中...</div>';
  google.script.run
    .withSuccessHandler(function(list){
      ORDER_CACHE = list || [];
      filterOrders();
    })
    .withFailureHandler(function(err){
      document.getElementById('orderResults').innerHTML =
        '<div class="muted" style="padding:16px;">読込エラー: ' + escHtml(err.message || err) + '</div>';
    })
    .api_getOrderList();
}

function closeOrderPicker(){
  document.getElementById('orderPickerWrap').style.display = 'none';
}

function filterOrders(){
  var q = (document.getElementById('orderSearchInput').value || '').toLowerCase().trim();
  var container = document.getElementById('orderResults');

  if(!q){
    container.innerHTML =
      '<div class="muted" style="padding:16px; text-align:center;">工番・品名・納入先を入力して検索</div>';
    return;
  }

  var terms = q.split(/\s+/).filter(function(t){ return t.length > 0; });

  var matched = ORDER_CACHE.filter(function(o){
    var haystack = (o.code + ' ' + o.product + ' ' + o.dest).toLowerCase();
    return terms.every(function(t){ return haystack.indexOf(t) >= 0; });
  });

  if(matched.length === 0){
    container.innerHTML =
      '<div class="muted" style="padding:16px; text-align:center;">該当する工番がありません</div>';
    return;
  }

  var html = '';
  var limit = Math.min(matched.length, 100);
  for(var i = 0; i < limit; i++){
    var o = matched[i];
    var detail = '';
    if(o.product) detail += o.product;
    if(o.product && o.dest) detail += ' ／ ';
    if(o.dest) detail += o.dest;
    html += '<div class="order-item" data-on-click="selectOrder" data-args="' + __vpsArgs(o.code) + '">' +
      '<div class="order-code">' + escHtml(o.code) + '</div>' +
      (detail ? '<div class="order-detail">' + escHtml(detail) + '</div>' : '') +
      '</div>';
  }
  if(matched.length > 100){
    html += '<div class="muted" style="padding:10px; text-align:center;">他 ' + (matched.length - 100) + ' 件…検索を絞り込んでください</div>';
  }
  container.innerHTML = html;
}

function escHtml(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function selectOrder(code){
  for(var i = 0; i < 3; i++){
    if(!ORDER_SELECTED[i]){
      ORDER_SELECTED[i] = code;
      updateOrderSlots();
      return;
    }
  }
  alert('工番は最大3件まで選択できます。\n不要な工番を×で削除してください。');
}

function clearOrderSlot(idx){
  ORDER_SELECTED[idx] = '';
  var compact = ORDER_SELECTED.filter(function(c){ return c !== ''; });
  while(compact.length < 3) compact.push('');
  ORDER_SELECTED = compact;
  updateOrderSlots();
}

function updateOrderSlots(){
  for(var i = 0; i < 3; i++){
    var el = document.getElementById('orderSlot' + (i+1));
    if(ORDER_SELECTED[i]){
      el.style.display = 'flex';
      var info = ORDER_CACHE.find(function(o){ return o.code === ORDER_SELECTED[i]; });
      var label = ORDER_SELECTED[i];
      if(info){
        if(info.product) label += ' ｜ ' + info.product;
        if(info.dest) label += ' ／ ' + info.dest;
      }
      el.querySelector('.slot-label').textContent = label;
    } else {
      el.style.display = 'none';
    }
  }
}

function submitOrderPicker(skipAll){
  closeOrderPicker();
  var codes = skipAll ? [] : ORDER_SELECTED.filter(function(c){ return c !== ''; });
  doOpenForm(
    ORDER_CONTEXT.type,
    ORDER_CONTEXT.dept,
    ORDER_CONTEXT.workerLabel,
    ORDER_CONTEXT.targetDateStr,
    codes
  );
}

// ====== 共通：フォームを開く ======
function doOpenForm(type, dept, workerLabel, targetDateStr, orderCodes){
  var label = (type === 'overtime') ? '残業' : '休日出勤';

  // ポップアップブロッカー対策
  var newWin = window.open('about:blank', '_blank');
  if(newWin){
    newWin.document.write(
      '<html><body style="font-family:system-ui,sans-serif;display:flex;align-items:center;' +
      'justify-content:center;height:100vh;margin:0;color:#5c6b7a;">' +
      '<div style="text-align:center;"><p style="font-size:18px;font-weight:700;">' +
      label + '申請フォームを準備中...</p>' +
      '<p style="font-size:13px;">しばらくお待ちください</p></div></body></html>'
    );
  }

  google.script.run
    .withSuccessHandler(function(url){
      if(url && newWin){
        newWin.location.href = url;
      } else if(url && !newWin){
        showFormLink(url, label);
      } else {
        if(newWin) newWin.close();
        alert(label + '申請フォームを作成できませんでした。管理者に連絡してください。');
      }
    })
    .withFailureHandler(function(err){
      if(newWin) newWin.close();
      alert(label + 'フォーム取得エラー: ' + (err.message || err));
    })
    .api_getFormUrl(type, dept, workerLabel, targetDateStr, orderCodes || []);
}

function showFormLink(url, label){
  var div = document.createElement('div');
  div.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);' +
    'z-index:300;background:#fff;border:2px solid var(--accent2);border-radius:18px;' +
    'padding:24px;text-align:center;box-shadow:0 10px 30px rgba(0,0,0,.15);max-width:360px;width:90%;';
  div.innerHTML =
    '<div style="font-weight:900;font-size:16px;margin-bottom:8px;">ポップアップがブロックされました</div>' +
    '<div class="sub" style="margin-bottom:14px;">下のリンクをクリックしてフォームを開いてください</div>' +
    '<a href="' + url + '" target="_blank" style="display:inline-block;padding:12px 24px;' +
    'background:linear-gradient(135deg,var(--accent2),#38bdf8);color:#fff;border-radius:12px;' +
    'font-weight:800;text-decoration:none;">' + label + 'フォームを開く</a>' +
    '<div style="margin-top:12px;"><button class="ghost btn-sm" data-on-click="__vpsRemoveGrandparent" data-args="[&quot;$this&quot;]">閉じる</button></div>';
  document.body.appendChild(div);
}

// ====== 全社ダッシュボード読込 ======
function loadDashboard(){
  var otTb = document.querySelector('#tblOvertime tbody');
  var hdTb = document.querySelector('#tblHoliday tbody');
  otTb.innerHTML = '<tr><td colspan="9" class="loading">読込中...</td></tr>';
  hdTb.innerHTML = '<tr><td colspan="10" class="loading">読込中...</td></tr>';

  google.script.run
    .withSuccessHandler(function(res){
      console.log('api_getDashboard 結果:', JSON.stringify(res));
      if(!res){
        otTb.innerHTML = '<tr><td colspan="9" class="muted" style="text-align:center;padding:24px">データを取得できませんでした（res=null）</td></tr>';
        hdTb.innerHTML = '<tr><td colspan="10" class="muted" style="text-align:center;padding:24px">デプロイの更新を確認してください</td></tr>';
        return;
      }
      if(res.error){
        otTb.innerHTML = '<tr><td colspan="9" class="muted" style="text-align:center;padding:24px">サーバーエラー: ' + escHtml(res.error) + '</td></tr>';
        hdTb.innerHTML = '<tr><td colspan="10"></td></tr>';
        return;
      }
      renderDashboard(res);
    })
    .withFailureHandler(function(err){
      console.error('api_getDashboard 失敗:', err);
      otTb.innerHTML = '<tr><td colspan="9" class="muted" style="text-align:center;padding:24px">API失敗: ' + escHtml(err.message || err) + '</td></tr>';
      hdTb.innerHTML = '<tr><td colspan="10" class="muted" style="text-align:center;padding:24px">GASエディタのログも確認してください</td></tr>';
    })
    .api_getDashboard();
}

// ====== ダッシュボード描画 ======
function renderDashboard(res){
  // 編集モーダル用にデータをキャッシュ
  DASHBOARD_DATA = { overtime: res.overtime || [], holiday: res.holiday || [] };

  document.getElementById('todayDate').innerText = res.today;

  // 残業テーブル
  var otTb = document.querySelector('#tblOvertime tbody');
  otTb.innerHTML = '';
  document.getElementById('otCount').innerText = res.overtime.length + '名';

  // 残業の期限超過件数を集計してバナー表示
  var otOverdueCount = 0;
  (res.overtime || []).forEach(function(it){ if(it.isOverdue) otOverdueCount++; });
  var otBanner = document.getElementById('otOverdueBanner');
  if(!otBanner){
    otBanner = document.createElement('div');
    otBanner.id = 'otOverdueBanner';
    otBanner.className = 'overdue-banner';
    var otCard = document.getElementById('tblOvertime').parentElement;
    otCard.insertBefore(otBanner, document.getElementById('tblOvertime'));
  }
  if(otOverdueCount > 0){
    otBanner.innerHTML = '<span style="font-size:18px;flex-shrink:0;">&#9888;</span>' +
      '<span>作業完了が未記録の申請があります（' + otOverdueCount + '件）。手入力で記録してください。</span>';
    otBanner.style.display = 'flex';
  } else {
    otBanner.style.display = 'none';
  }

  if(!res.overtime || res.overtime.length === 0){
    otTb.innerHTML =
      '<tr><td colspan="9" class="muted" style="text-align:center;padding:24px">' +
      '本日の残業申請はありません</td></tr>';
  } else {
    res.overtime.forEach(function(it){ otTb.appendChild(buildOvertimeRow(it)); });
  }

  // 休日テーブル
  var hdTb = document.querySelector('#tblHoliday tbody');
  hdTb.innerHTML = '';
  document.getElementById('hdCount').innerText = res.holiday.length + '名';

  // 休日の期限超過件数を集計してバナー表示
  var hdOverdueCount = 0;
  (res.holiday || []).forEach(function(it){ if(it.isOverdue) hdOverdueCount++; });
  var hdBanner = document.getElementById('hdOverdueBanner');
  if(!hdBanner){
    hdBanner = document.createElement('div');
    hdBanner.id = 'hdOverdueBanner';
    hdBanner.className = 'overdue-banner';
    var hdCard = document.getElementById('tblHoliday').parentElement;
    hdCard.insertBefore(hdBanner, document.getElementById('tblHoliday'));
  }
  if(hdOverdueCount > 0){
    hdBanner.innerHTML = '<span style="font-size:18px;flex-shrink:0;">&#9888;</span>' +
      '<span>作業完了が未記録の申請があります（' + hdOverdueCount + '件）。手入力で記録してください。</span>';
    hdBanner.style.display = 'flex';
  } else {
    hdBanner.style.display = 'none';
  }

  if(!res.holiday || res.holiday.length === 0){
    hdTb.innerHTML =
      '<tr><td colspan="10" class="muted" style="text-align:center;padding:24px">' +
      '休日出勤申請はありません</td></tr>';
  } else {
    res.holiday.forEach(function(it){ hdTb.appendChild(buildHolidayRow(it)); });
  }
}

// ====== 残業行 ======
function buildOvertimeRow(it){
  var statusPill = it.status === 'submitted'
    ? '<span class="pill warn">承認待ち</span>'
    : '<span class="pill ok">承認済</span>';

  var pdfCell = buildPdfCell(it);
  var actionCell = buildActionCell(it);

  // 過去日の期限超過アイテムには日付ラベルを表示
  var nameCell = '<b>' + escHtml(it.workerName) + '</b>';
  if(it.isOverdue && it.targetDateLabel){
    nameCell += ' <span class="pill danger" style="font-size:11px;">' + escHtml(it.targetDateLabel) + '</span>';
  }

  var tr = document.createElement('tr');
  tr.setAttribute('data-request-id', it.requestId);
  if(it.isOverdue) tr.style.background = 'rgba(255,90,95,0.05)';
  tr.innerHTML =
    '<td>' + escHtml(it.dept) + '</td>' +
    '<td>' + nameCell + '</td>' +
    '<td>' + statusPill + '</td>' +
    '<td>' + fmtMin(it.approvedMinutes) + '</td>' +
    '<td>' + (it.actualMinutes > 0 ? fmtMin(it.actualMinutes) : '<span class="muted">-</span>') + '</td>' +
    '<td>' + (it.breakMinutes > 0 ? fmtMin(it.breakMinutes) : '<span class="muted">-</span>') + '</td>' +
    '<td>' + (it.netMinutes > 0 ? '<b>' + fmtMin(it.netMinutes) + '</b>' : '<span class="muted">-</span>') + '</td>' +
    '<td>' + pdfCell + '</td>' +
    '<td>' + actionCell + '</td>';
  return tr;
}

// ====== 休日行 ======
function buildHolidayRow(it){
  var statusPill = it.status === 'submitted'
    ? '<span class="pill warn">承認待ち</span>'
    : '<span class="pill ok">承認済</span>';

  var pdfCell = buildPdfCell(it);
  var actionCell = buildActionCell(it);

  var tr = document.createElement('tr');
  tr.setAttribute('data-request-id', it.requestId);
  if(it.isOverdue) tr.style.background = 'rgba(255,90,95,0.05)';
  tr.innerHTML =
    '<td>' + escHtml(it.targetDateLabel || it.targetDate) + '</td>' +
    '<td>' + escHtml(it.dept) + '</td>' +
    '<td><b>' + escHtml(it.workerName) + '</b></td>' +
    '<td>' + statusPill + '</td>' +
    '<td>' + fmtMin(it.approvedMinutes) + '</td>' +
    '<td>' + (it.actualMinutes > 0 ? fmtMin(it.actualMinutes) : '<span class="muted">-</span>') + '</td>' +
    '<td>' + (it.breakMinutes > 0 ? fmtMin(it.breakMinutes) : '<span class="muted">-</span>') + '</td>' +
    '<td>' + (it.netMinutes > 0 ? '<b>' + fmtMin(it.netMinutes) + '</b>' : '<span class="muted">-</span>') + '</td>' +
    '<td>' + pdfCell + '</td>' +
    '<td>' + actionCell + '</td>';
  return tr;
}

// ====== PDF列 ======
function buildPdfCell(it){
  if(it.pdfFileId){
    return '<a href="https://drive.google.com/file/d/' +
      escHtml(it.pdfFileId) + '/view" target="_blank" class="pill ok">PDF</a>';
  } else if(it.netMinutes > 0 && it.status === 'approved'){
    return '<span class="pill warn">未作成</span>';
  }
  return '<span class="muted">-</span>';
}

// ====== 4時間超過アラーム判定 ======
function check4hAlert(it){
  if(it.status !== 'approved' || it.actualEndAt) return '';
  var now = new Date();
  var elapsed = 0;

  if(it.requestType === 'overtime'){
    // 残業: targetDate 17:20 JST からの経過
    var startTime = new Date(it.targetDate + 'T17:20:00+09:00');
    elapsed = (now.getTime() - startTime.getTime()) / 60000;
  } else if(it.requestType === 'holiday' && it.actualStartAt){
    // 休日: actualStartAt からの経過
    var startStr = String(it.actualStartAt).replace(' ','T');
    if(startStr.indexOf('+') < 0 && startStr.indexOf('Z') < 0) startStr += '+09:00';
    var startTime2 = new Date(startStr);
    if(!isNaN(startTime2.getTime())){
      elapsed = (now.getTime() - startTime2.getTime()) / 60000;
    }
  }

  if(elapsed > 240){
    return ' <span class="pill danger">4h超過</span>';
  }
  return '';
}

// ====== 操作ボタン構築 ======
function buildEditDeleteBtns(it){
  var rid = it.requestId;
  return ' <button class="ghost btn-sm" data-on-click="openEditModal" data-args="' + __vpsArgs(rid) + '" style="font-size:11px;">編集</button>' +
    ' <button class="ghost btn-sm" data-on-click="confirmDelete" data-args="' + __vpsArgs(rid) + '" style="font-size:11px; color:var(--danger);">削除</button>';
}

function buildActionCell(it){
  var alert4h = check4hAlert(it);
  var rid = it.requestId;
  var overdueBadge = '';

  var pendingLabel = (it.status === 'submitted') ? '<span class="pill warn">承認待ち</span> ' : '';

  if(it.requestType === 'overtime'){
    if(it.actualEndAt){
      if(it.status === 'submitted') return '<span class="pill warn">承認待ち</span> <span class="pill ok">完了済</span>';
      return '<span class="pill ok">完了済</span>';
    }
    if(it.isOverdue) overdueBadge = '<span class="pill danger">&#9888; 期限超過</span> ';
    return pendingLabel + overdueBadge +
      '<button class="btn-blue btn-sm" data-on-click="doOvertimeDone" data-args="' + __vpsArgs(rid, '$this') + '">残業完了</button>' +
      ' <button class="ghost btn-sm" data-on-click="toggleManualInput" data-args="' + __vpsArgs(rid, 'ot-end') + '">手入力</button>' +
      buildEditDeleteBtns(it) +
      alert4h +
      '<div class="manual-input-area" id="manual-area-' + rid + '">' +
        '<span style="font-size:12px;">終了時刻：</span>' +
        buildTimeSelect('manual-end-' + rid, 'ot-end', it.isOverdue ? '20:00' : '') +
        ' <button class="btn-blue btn-sm" data-on-click="doManualOvertimeDone" data-args="' + __vpsArgs(rid) + '">記録</button>' +
      '</div>';
  }

  if(it.requestType === 'holiday'){
    if(it.actualEndAt){
      if(it.status === 'submitted') return '<span class="pill warn">承認待ち</span> <span class="pill ok">完了済</span>';
      return '<span class="pill ok">完了済</span>';
    }
    if(it.actualStartAt){
      if(it.isOverdue) overdueBadge = '<span class="pill danger">&#9888; 期限超過</span> ';
      return pendingLabel + overdueBadge +
        '<button class="btn-sm" data-on-click="doHolidayDone" data-args="' + __vpsArgs(rid, '$this') + '">休日完了</button>' +
        ' <button class="ghost btn-sm" data-on-click="toggleManualInput" data-args="' + __vpsArgs(rid, 'hd-end') + '">手入力</button>' +
        buildEditDeleteBtns(it) +
        alert4h +
        '<div class="manual-input-area" id="manual-area-' + rid + '">' +
          '<span style="font-size:12px;">終了時刻：</span>' +
          buildTimeSelect('manual-end-' + rid, 'hd-end', it.isOverdue ? '17:00' : '') +
          ' <button class="btn-sm" data-on-click="doManualHolidayDone" data-args="' + __vpsArgs(rid) + '">記録</button>' +
        '</div>';
    }
    if(it.isOverdue) overdueBadge = '<span class="pill danger">&#9888; 未開始</span> ';
    return pendingLabel + overdueBadge +
      '<button class="btn-warn btn-sm" data-on-click="doHolidayStart" data-args="' + __vpsArgs(rid, '$this') + '">開始</button>' +
      ' <button class="ghost btn-sm" data-on-click="toggleManualInput" data-args="' + __vpsArgs(rid, 'hd-start') + '">手入力</button>' +
      buildEditDeleteBtns(it) +
      '<div class="manual-input-area" id="manual-area-' + rid + '">' +
        '<span style="font-size:12px;">開始時刻：</span>' +
        buildTimeSelect('manual-start-' + rid, 'hd-start', it.isOverdue ? '08:00' : '') +
        ' <button class="btn-warn btn-sm" data-on-click="doManualHolidayStart" data-args="' + __vpsArgs(rid) + '">記録</button>' +
      '</div>';
  }

  return '-';
}

function updateRowAction(requestId, newHtml) {
  var row = document.querySelector('tr[data-request-id="' + requestId + '"]');
  if (!row) return;
  var cells = row.querySelectorAll('td');
  cells[cells.length - 1].innerHTML = newHtml;
  row.style.transition = 'background 0.4s';
  row.style.background = 'rgba(14,165,233,0.08)';
  setTimeout(function(){ row.style.background = ''; }, 1200);
}

// ====== 手入力フィールドの展開/折りたたみ ======
function toggleManualInput(requestId, type){
  var area = document.getElementById('manual-area-' + requestId);
  if(!area) return;
  area.style.display = area.style.display === 'block' ? 'none' : 'block';
}

// ====== 10分刻みの時刻プルダウン生成 ======
// type: 'ot-end'(残業終了17:30~翌03:00), 'hd-start'(休日開始06:00~18:00), 'hd-end'(休日終了08:00~23:00)
function buildTimeSelect(id, type, defaultVal){
  var ranges = {
    'ot-end':   { startH:17, startM:30, endH:27, endM:0 },
    'hd-start': { startH:6,  startM:0,  endH:18, endM:0 },
    'hd-end':   { startH:8,  startM:0,  endH:23, endM:0 }
  };
  var r = ranges[type] || { startH:6, startM:0, endH:23, endM:0 };
  var html = '<select id="' + id + '">';
  html += '<option value="">--:--</option>';
  for(var t = r.startH * 60 + r.startM; t <= r.endH * 60 + r.endM; t += 10){
    var h = Math.floor(t / 60) % 24;
    var m = t % 60;
    var val = (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
    var sel = (val === defaultVal) ? ' selected' : '';
    html += '<option value="' + val + '"' + sel + '>' + val + '</option>';
  }
  html += '</select>';
  return html;
}

// ====== 残業完了 ======
function doOvertimeDone(requestId, btn){
  if(!confirm('残業完了を記録しますか？\n（17:20から現在時刻まで自動計算されます）')) return;
  btn.disabled = true;
  updateRowAction(requestId, '<span class="pill ok">完了済</span>');

  google.script.run
    .withSuccessHandler(function(res){
      var row = document.querySelector('tr[data-request-id="' + requestId + '"]');
      if(row){
        var c = row.querySelectorAll('td');
        c[4].innerHTML = fmtMin(res.actualMinutes);
        c[5].innerHTML = fmtMin(res.breakMinutes);
        c[6].innerHTML = '<b>' + fmtMin(res.netMinutes) + '</b>';
      }
      alert('残業完了を記録しました。\n実働: ' + fmtMin(res.actualMinutes) +
        '\n休憩控除: ' + fmtMin(res.breakMinutes) + '\n実残業: ' + fmtMin(res.netMinutes));
    })
    .withFailureHandler(function(err){
      alert('エラー: ' + (err.message || err));
      loadDashboard();
    })
    .api_markOvertimeDone(requestId);
}

// ====== 手入力で残業完了 ======
function doManualOvertimeDone(requestId){
  var timeStr = document.getElementById('manual-end-' + requestId).value;
  if(!timeStr){ alert('終了時刻を入力してください。'); return; }
  if(!confirm('終了時刻 ' + timeStr + ' で残業完了を記録しますか？')) return;
  updateRowAction(requestId, '<span class="pill ok">完了済</span>');

  google.script.run
    .withSuccessHandler(function(res){
      var row = document.querySelector('tr[data-request-id="' + requestId + '"]');
      if(row){
        var c = row.querySelectorAll('td');
        c[4].innerHTML = fmtMin(res.actualMinutes);
        c[5].innerHTML = fmtMin(res.breakMinutes);
        c[6].innerHTML = '<b>' + fmtMin(res.netMinutes) + '</b>';
      }
      alert('残業完了を記録しました。\n実働: ' + fmtMin(res.actualMinutes) +
        '\n休憩控除: ' + fmtMin(res.breakMinutes) + '\n実残業: ' + fmtMin(res.netMinutes));
    })
    .withFailureHandler(function(err){
      alert('エラー: ' + (err.message || err));
      loadDashboard();
    })
    .api_markOvertimeDone(requestId, timeStr);
}

// ====== 休日開始 ======
function doHolidayStart(requestId, btn){
  if(!confirm('休日出勤を開始しますか？\n（現在時刻が開始時刻として記録されます）')) return;
  btn.disabled = true;
  updateRowAction(requestId,
    '<button class="btn-sm" data-on-click="doHolidayDone" data-args="' + __vpsArgs(requestId, '$this') + '">休日完了</button>');

  google.script.run
    .withSuccessHandler(function(){})
    .withFailureHandler(function(err){
      alert('エラー: ' + (err.message || err));
      loadDashboard();
    })
    .api_markHolidayStart(requestId);
}

// ====== 手入力で休日開始 ======
function doManualHolidayStart(requestId){
  var timeStr = document.getElementById('manual-start-' + requestId).value;
  if(!timeStr){ alert('開始時刻を入力してください。'); return; }
  if(!confirm('開始時刻 ' + timeStr + ' で休日出勤開始を記録しますか？')) return;
  updateRowAction(requestId,
    '<button class="btn-sm" data-on-click="doHolidayDone" data-args="' + __vpsArgs(requestId, '$this') + '">休日完了</button>' +
    ' <button class="ghost btn-sm" data-on-click="toggleManualInput" data-args="' + __vpsArgs(requestId, 'hd-end') + '">手入力</button>' +
    '<div class="manual-input-area" id="manual-area-' + requestId + '">' +
      '<span style="font-size:12px;">終了時刻：</span>' +
      buildTimeSelect('manual-end-' + requestId, 'hd-end', '') +
      ' <button class="btn-sm" data-on-click="doManualHolidayDone" data-args="' + __vpsArgs(requestId) + '">記録</button>' +
    '</div>');

  google.script.run
    .withSuccessHandler(function(){})
    .withFailureHandler(function(err){
      alert('エラー: ' + (err.message || err));
      loadDashboard();
    })
    .api_markHolidayStart(requestId, timeStr);
}

// ====== 休日完了 ======
function doHolidayDone(requestId, btn){
  if(!confirm('休日出勤を完了しますか？\n（開始時刻から現在時刻まで自動計算されます）')) return;
  btn.disabled = true;
  updateRowAction(requestId, '<span class="pill ok">完了済</span>');

  google.script.run
    .withSuccessHandler(function(res){
      var row = document.querySelector('tr[data-request-id="' + requestId + '"]');
      if(row){
        var c = row.querySelectorAll('td');
        c[5].innerHTML = fmtMin(res.actualMinutes);
        c[6].innerHTML = fmtMin(res.breakMinutes);
        c[7].innerHTML = '<b>' + fmtMin(res.netMinutes) + '</b>';
      }
      alert('休日出勤を完了しました。\n実働: ' + fmtMin(res.actualMinutes) +
        '\n休憩控除: ' + fmtMin(res.breakMinutes) + '\n実残業/実働: ' + fmtMin(res.netMinutes));
    })
    .withFailureHandler(function(err){
      alert('エラー: ' + (err.message || err));
      loadDashboard();
    })
    .api_markHolidayDone(requestId);
}

// ====== 手入力で休日完了 ======
function doManualHolidayDone(requestId){
  var timeStr = document.getElementById('manual-end-' + requestId).value;
  if(!timeStr){ alert('終了時刻を入力してください。'); return; }
  if(!confirm('終了時刻 ' + timeStr + ' で休日出勤完了を記録しますか？')) return;
  updateRowAction(requestId, '<span class="pill ok">完了済</span>');

  google.script.run
    .withSuccessHandler(function(res){
      var row = document.querySelector('tr[data-request-id="' + requestId + '"]');
      if(row){
        var c = row.querySelectorAll('td');
        c[5].innerHTML = fmtMin(res.actualMinutes);
        c[6].innerHTML = fmtMin(res.breakMinutes);
        c[7].innerHTML = '<b>' + fmtMin(res.netMinutes) + '</b>';
      }
      alert('休日出勤を完了しました。\n実働: ' + fmtMin(res.actualMinutes) +
        '\n休憩控除: ' + fmtMin(res.breakMinutes) + '\n実残業/実働: ' + fmtMin(res.netMinutes));
    })
    .withFailureHandler(function(err){
      alert('エラー: ' + (err.message || err));
      loadDashboard();
    })
    .api_markHolidayDone(requestId, timeStr);
}

// ====== ダッシュボードデータキャッシュ（編集モーダル用） ======
var DASHBOARD_DATA = { overtime:[], holiday:[] };

// ====== 編集モーダル ======
var EDIT_CTX = {};

function openEditModal(requestId){
  var item = null;
  DASHBOARD_DATA.overtime.forEach(function(it){ if(it.requestId === requestId) item = it; });
  if(!item) DASHBOARD_DATA.holiday.forEach(function(it){ if(it.requestId === requestId) item = it; });
  if(!item){ alert('申請データが見つかりません。画面を更新してください。'); return; }

  EDIT_CTX = { requestId: requestId, requestType: item.requestType, item: item };

  var label = (item.requestType === 'overtime') ? '残業申請 編集' : '休日出勤申請 編集';
  document.getElementById('editModalTitle').textContent = label;

  // 対象者表示
  document.getElementById('editInfo').textContent = item.dept + ' ／ ' + item.workerName;

  // 予定時間プルダウン
  var hoursList = (item.requestType === 'overtime') ? OT_HOURS_LIST : HD_HOURS_LIST;
  var hSel = document.getElementById('editHours');
  hSel.innerHTML = '';
  hoursList.forEach(function(h){
    var opt = document.createElement('option');
    opt.value = h;
    opt.textContent = (item.requestType === 'overtime') ? h + ' 時間' : h;
    hSel.appendChild(opt);
  });
  // 現在の値を選択（approvedMinutes → hours ラベルに逆変換）
  if(item.requestType === 'overtime'){
    var curHours = (item.approvedMinutes / 60).toFixed(1);
    hSel.value = curHours;
  } else {
    hSel.value = (item.approvedMinutes >= 480) ? '1日' : '半日';
  }

  // 作業実施日（休日のみ表示）
  var dateWrap = document.getElementById('editDateWrap');
  if(item.requestType === 'holiday'){
    dateWrap.style.display = '';
    document.getElementById('editTargetDate').value = item.targetDate;
  } else {
    dateWrap.style.display = 'none';
  }

  // ボタンリセット
  var btn = document.getElementById('editBtnSave');
  btn.disabled = false;
  btn.textContent = '保存';

  document.getElementById('editModalWrap').style.display = '';
}

function closeEditModal(){
  document.getElementById('editModalWrap').style.display = 'none';
}

function saveEdit(){
  var hours = document.getElementById('editHours').value;
  if(!hours){ alert('予定時間を選択してください。'); return; }

  var patch = { hours: hours };
  if(EDIT_CTX.requestType === 'holiday'){
    var td = document.getElementById('editTargetDate').value;
    if(td) patch.targetDate = td;
  }

  var msg = '申請内容を変更しますか？\n予定時間: ' + hours;
  if(patch.targetDate) msg += '\n作業実施日: ' + patch.targetDate;
  if(!confirm(msg)) return;

  var btn = document.getElementById('editBtnSave');
  btn.disabled = true;
  btn.textContent = '保存中...';

  google.script.run
    .withSuccessHandler(function(res){
      closeEditModal();
      if(res && res.ok){
        alert('申請を更新しました。');
        loadDashboard();
      } else {
        alert('更新に失敗しました。');
      }
    })
    .withFailureHandler(function(err){
      btn.disabled = false;
      btn.textContent = '保存';
      alert('エラー: ' + (err.message || err));
    })
    .api_editRequest(EDIT_CTX.requestId, patch);
}

function confirmDelete(requestId){
  var item = null;
  DASHBOARD_DATA.overtime.forEach(function(it){ if(it.requestId === requestId) item = it; });
  if(!item) DASHBOARD_DATA.holiday.forEach(function(it){ if(it.requestId === requestId) item = it; });

  var label = item ? (item.workerName + ' の' + (item.requestType === 'overtime' ? '残業' : '休日出勤') + '申請') : '選択した申請';
  if(!confirm(label + 'を削除しますか？\n\nこの操作は取り消せません。')) return;

  google.script.run
    .withSuccessHandler(function(res){
      if(res && res.ok){
        alert('申請を削除しました。');
        loadDashboard();
      } else {
        alert('削除に失敗しました。');
      }
    })
    .withFailureHandler(function(err){
      alert('削除エラー: ' + (err.message || err));
    })
    .api_deleteRequest(requestId);
}

// ====== 申請モーダル制御 ======
var REQ_CTX = { type:'', dept:'', workerLabel:'', targetDate:'', orders:['','',''] };

// ====== 過去申請カレンダー 状態 ======
// 選択範囲: 現在の年度開始日（3/16） 〜 今日
// バックエンドの getFiscalYearStartDate_ と同じロジックで動的算出する。
function retroGetFyStart_(){
  var t = new Date();
  var y = t.getFullYear(), m = t.getMonth()+1, d = t.getDate();
  // 3/16以降ならその年の3/16、それ以前なら前年の3/16
  var sy = (m > 3 || (m === 3 && d >= 16)) ? y : (y - 1);
  return new Date(sy, 2, 16); // 月は0始まり
}
var selectedRetroactiveDate = null; // 'yyyy-MM-dd' or null
var retroViewYear = 0;   // カレンダー表示中の年
var retroViewMonth = 0;  // カレンダー表示中の月 (0-11)

function retroToday_(){
  var t = new Date();
  return new Date(t.getFullYear(), t.getMonth(), t.getDate());
}
function retroFmtYmd_(d){
  var y = d.getFullYear();
  var m = String(d.getMonth()+1).padStart(2,'0');
  var dd = String(d.getDate()).padStart(2,'0');
  return y + '-' + m + '-' + dd;
}
function retroParseYmd_(s){
  if(!s) return null;
  var p = s.split('-');
  return new Date(Number(p[0]), Number(p[1])-1, Number(p[2]));
}

function toggleRetroPanel(){
  var wrap = document.getElementById('retroPanelWrap');
  var btn  = document.getElementById('retroToggleBtn');
  var open = wrap.classList.toggle('open');
  btn.classList.toggle('open', open);
  if(open){
    // 初回展開時、今日を含む月を表示
    var t = retroToday_();
    if(retroViewYear === 0){
      retroViewYear = t.getFullYear();
      retroViewMonth = t.getMonth();
    }
    renderRetroCalendar();
  }
}

function resetRetroCalendar_(){
  selectedRetroactiveDate = null;
  var wrap = document.getElementById('retroPanelWrap');
  var btn  = document.getElementById('retroToggleBtn');
  if(wrap) wrap.classList.remove('open');
  if(btn) btn.classList.remove('open');
  var lbl = document.getElementById('retroSelectedLabel');
  if(lbl) lbl.textContent = '選択中: （未選択）';
  var t = retroToday_();
  retroViewYear = t.getFullYear();
  retroViewMonth = t.getMonth();
}

function clearRetroSelection(){
  selectedRetroactiveDate = null;
  document.getElementById('retroSelectedLabel').textContent = '選択中: （未選択）';
  renderRetroCalendar();
}

function retroPrevMonth(){
  var m = retroViewMonth - 1;
  var y = retroViewYear;
  if(m < 0){ m = 11; y -= 1; }
  retroViewYear = y; retroViewMonth = m;
  renderRetroCalendar();
}
function retroNextMonth(){
  var m = retroViewMonth + 1;
  var y = retroViewYear;
  if(m > 11){ m = 0; y += 1; }
  retroViewYear = y; retroViewMonth = m;
  renderRetroCalendar();
}

function renderRetroCalendar(){
  var today = retroToday_();
  var minD  = retroGetFyStart_();

  // ラベル
  document.getElementById('retroYmLabel').textContent =
    retroViewYear + '年 ' + (retroViewMonth + 1) + '月';

  // ◀▶ 活性/非活性判定
  // 前月: 表示中の月の1日が 年度開始月 より前の月ならdisable
  var firstOfView = new Date(retroViewYear, retroViewMonth, 1);
  var firstOfMin  = new Date(minD.getFullYear(), minD.getMonth(), 1);
  document.getElementById('retroPrevBtn').disabled = (firstOfView <= firstOfMin);
  // 翌月: 表示中の月が今日の月以降ならdisable
  var firstOfToday = new Date(today.getFullYear(), today.getMonth(), 1);
  document.getElementById('retroNextBtn').disabled = (firstOfView >= firstOfToday);

  // グリッド描画
  var grid = document.getElementById('retroGrid');
  grid.innerHTML = '';
  var dowLabels = ['日','月','火','水','木','金','土'];
  for(var i=0; i<7; i++){
    var dowEl = document.createElement('div');
    dowEl.className = 'retro-dow' + (i===0?' sun':'') + (i===6?' sat':'');
    dowEl.textContent = dowLabels[i];
    grid.appendChild(dowEl);
  }

  // 1日の曜日
  var first = new Date(retroViewYear, retroViewMonth, 1);
  var startDow = first.getDay();
  // 月末日
  var lastDate = new Date(retroViewYear, retroViewMonth + 1, 0).getDate();

  // 6行x7列 = 42マス
  for(var c=0; c<42; c++){
    var cell = document.createElement('div');
    var dayNum = c - startDow + 1;
    if(dayNum < 1 || dayNum > lastDate){
      cell.className = 'retro-cell other-month';
      cell.textContent = '';
      grid.appendChild(cell);
      continue;
    }
    var cellDate = new Date(retroViewYear, retroViewMonth, dayNum);
    var ymd = retroFmtYmd_(cellDate);
    var disabled = (cellDate < minD) || (cellDate > today);
    var cls = 'retro-cell';
    if(disabled) cls += ' disabled';
    if(ymd === retroFmtYmd_(today)) cls += ' today';
    if(selectedRetroactiveDate === ymd) cls += ' selected';
    cell.className = cls;
    cell.textContent = String(dayNum);
    if(!disabled){
      (function(ds){
        cell.onclick = function(){ selectRetroDate(ds); };
      })(ymd);
    }
    grid.appendChild(cell);
  }
}

function selectRetroDate(ymd){
  selectedRetroactiveDate = ymd;
  document.getElementById('retroSelectedLabel').textContent = '選択中: ' + ymd;
  renderRetroCalendar();
}

function showRequestModal(type, dept, workerLabel){
  REQ_CTX = { type:type, dept:dept, workerLabel:workerLabel, targetDate:'', orders:['','',''] };

  // タイトル設定
  var label = (type === 'overtime') ? '残業申請' : '休日出勤申請';
  document.getElementById('reqModalTitle').textContent = label;

  // ステップ1表示
  document.getElementById('reqStep1').style.display = '';
  document.getElementById('reqStep2').style.display = 'none';
  document.getElementById('reqBtnNext').style.display = '';
  document.getElementById('reqBtnSubmit').style.display = 'none';
  document.getElementById('reqBtnBack').style.display = 'none';

  // 日付設定
  if(type === 'overtime'){
    var today = new Date();
    var yyyy = today.getFullYear();
    var mm = String(today.getMonth()+1).padStart(2,'0');
    var dd = String(today.getDate()).padStart(2,'0');
    REQ_CTX.targetDate = yyyy + '-' + mm + '-' + dd;
    document.getElementById('reqDateDisplay').textContent = REQ_CTX.targetDate + '（本日）';
    document.getElementById('reqDateDisplay').style.display = '';
    document.getElementById('reqDatePicker').style.display = 'none';
  } else {
    // 休日: 候補日を取得して表示
    document.getElementById('reqDateDisplay').textContent = '日付を選択してください';
    document.getElementById('reqDateDisplay').style.display = 'none';
    document.getElementById('reqDatePicker').style.display = '';
    loadHolidayDates();
  }

  // 工番リセット
  REQ_CTX.orders = ['','',''];
  updateReqOrderSlots();
  document.getElementById('reqOrderSearch').value = '';
  document.getElementById('reqOrderResults').innerHTML =
    '<div class="muted" style="padding:14px; text-align:center;">工番・品名・納入先を入力して検索</div>';

  // 理由・時間フォームリセット
  document.getElementById('reqReason').innerHTML = '<option value="">-- 選択してください --</option>';
  REASONS.forEach(function(r){
    var opt = document.createElement('option');
    opt.value = r; opt.textContent = r;
    document.getElementById('reqReason').appendChild(opt);
  });
  document.getElementById('reqReasonDetailWrap').style.display = 'none';
  document.getElementById('reqReasonDetail').value = '';
  document.getElementById('reqWorkContent').value = '';

  // 予定時間プルダウン
  var hoursList = (type === 'overtime') ? OT_HOURS_LIST : HD_HOURS_LIST;
  var hSel = document.getElementById('reqHours');
  hSel.innerHTML = '<option value="">-- 選択してください --</option>';
  hoursList.forEach(function(h){
    var opt = document.createElement('option');
    opt.value = h; opt.textContent = (type === 'overtime') ? h + ' 時間' : h;
    hSel.appendChild(opt);
  });

  // 工番キャッシュがなければ取得
  if(ORDER_CACHE.length === 0){
    google.script.run
      .withSuccessHandler(function(list){ ORDER_CACHE = list || []; })
      .withFailureHandler(function(){})
      .api_getOrderList();
  }

  // 過去申請カレンダーを初期化（閉じた状態・選択解除）
  resetRetroCalendar_();

  document.getElementById('requestModalWrap').style.display = '';
}

function closeRequestModal(){
  document.getElementById('requestModalWrap').style.display = 'none';
  // ボタン状態をリセット（エラー後の再オープン対策）
  var btn = document.getElementById('reqBtnSubmit');
  btn.disabled = false;
  btn.textContent = '申請送信';
}

// --- 休日候補日取得 ---
function loadHolidayDates(){
  var container = document.getElementById('reqDateButtons');
  container.innerHTML = '<div class="loading" style="padding:10px;">候補日を取得中...</div>';
  google.script.run
    .withSuccessHandler(function(candidates){
      container.innerHTML = '';
      if(!candidates || candidates.length === 0){
        container.innerHTML = '<div class="muted">候補日が見つかりませんでした</div>';
        return;
      }
      candidates.forEach(function(c){
        var btn = document.createElement('button');
        btn.className = 'req-date-btn';
        var lbl = c.label;
        if(c.holidayName) lbl += ' (' + c.holidayName + ')';
        btn.textContent = lbl;
        // 土曜=緑、日曜・祝日=青
        var isSat = c.label.indexOf('(土)') >= 0;
        var isSunOrHoliday = c.label.indexOf('(日)') >= 0 || c.type === 'holiday';
        if(isSat){
          btn.style.borderLeft = '4px solid var(--accent)';
          btn.style.color = '#0b5a2b';
          btn.style.background = '#f0fff6';
        } else if(isSunOrHoliday){
          btn.style.borderLeft = '4px solid var(--accent2)';
          btn.style.color = '#0b2a5a';
          btn.style.background = '#f0f7ff';
        }
        btn.onclick = function(){ selectReqDate(c.date); };
        container.appendChild(btn);
      });
    })
    .withFailureHandler(function(err){
      container.innerHTML = '<div class="muted">取得エラー: ' + escHtml(err.message || err) + '</div>';
    })
    .api_getHolidayCandidateDates();
}

function selectReqDate(dateStr){
  if(!dateStr){ alert('日付を選択してください。'); return; }
  REQ_CTX.targetDate = dateStr;
  // 全ボタンのselected解除 + 選択中表示
  var btns = document.querySelectorAll('#reqDateButtons .req-date-btn');
  btns.forEach(function(b){ b.classList.remove('selected'); });
  btns.forEach(function(b){
    if(b.textContent.indexOf(dateStr) >= 0) b.classList.add('selected');
  });
  // 日付表示を更新
  document.getElementById('reqDateDisplay').textContent = dateStr;
  document.getElementById('reqDateDisplay').style.display = '';
}

// --- 工番選択（モーダル内） ---
function filterReqOrders(){
  var q = (document.getElementById('reqOrderSearch').value || '').toLowerCase().trim();
  var container = document.getElementById('reqOrderResults');
  if(!q){
    container.innerHTML = '<div class="muted" style="padding:14px; text-align:center;">工番・品名・納入先を入力して検索</div>';
    return;
  }
  var terms = q.split(/\s+/).filter(function(t){ return t.length > 0; });
  var matched = ORDER_CACHE.filter(function(o){
    var haystack = (o.code + ' ' + o.product + ' ' + o.dest).toLowerCase();
    return terms.every(function(t){ return haystack.indexOf(t) >= 0; });
  });
  if(matched.length === 0){
    container.innerHTML = '<div class="muted" style="padding:14px; text-align:center;">該当する工番がありません</div>';
    return;
  }
  container.innerHTML = '';
  var limit = Math.min(matched.length, 80);
  for(var i = 0; i < limit; i++){
    var o = matched[i];
    var div = document.createElement('div');
    div.className = 'order-item';
    var codeDiv = document.createElement('div');
    codeDiv.className = 'order-code';
    codeDiv.textContent = o.code;
    div.appendChild(codeDiv);
    var detail = '';
    if(o.product) detail += o.product;
    if(o.product && o.dest) detail += ' ／ ';
    if(o.dest) detail += o.dest;
    if(detail){
      var detDiv = document.createElement('div');
      detDiv.className = 'order-detail';
      detDiv.textContent = detail;
      div.appendChild(detDiv);
    }
    (function(code){ div.onclick = function(){ selectReqOrder(code); }; })(o.code);
    container.appendChild(div);
  }
  if(matched.length > 80){
    var msg = document.createElement('div');
    msg.className = 'muted';
    msg.style.cssText = 'padding:8px; text-align:center;';
    msg.textContent = '他 ' + (matched.length - 80) + ' 件…';
    container.appendChild(msg);
  }
}

function selectReqOrder(code){
  for(var i = 0; i < 3; i++){
    if(!REQ_CTX.orders[i]){
      REQ_CTX.orders[i] = code;
      updateReqOrderSlots();
      return;
    }
  }
  alert('工番は最大3件まで選択できます。\n不要な工番を×で削除してください。');
}

function clearReqOrderSlot(idx){
  REQ_CTX.orders[idx] = '';
  var compact = REQ_CTX.orders.filter(function(c){ return c !== ''; });
  while(compact.length < 3) compact.push('');
  REQ_CTX.orders = compact;
  updateReqOrderSlots();
}

function updateReqOrderSlots(){
  for(var i = 0; i < 3; i++){
    var el = document.getElementById('reqOrderSlot' + (i+1));
    if(REQ_CTX.orders[i]){
      el.style.display = 'flex';
      var info = ORDER_CACHE.find(function(o){ return o.code === REQ_CTX.orders[i]; });
      var label = REQ_CTX.orders[i];
      if(info){
        if(info.product) label += ' ｜ ' + info.product;
        if(info.dest) label += ' ／ ' + info.dest;
      }
      el.querySelector('.slot-label').textContent = label;
    } else {
      el.style.display = 'none';
    }
  }
}

// --- ステップ遷移 ---
function reqGoStep2(){
  // 休日出勤: 日付必須チェック
  if(REQ_CTX.type === 'holiday' && !REQ_CTX.targetDate){
    alert('作業実施日を選択してください。');
    return;
  }

  // サマリー生成
  var orderLabels = REQ_CTX.orders.filter(function(c){ return c !== ''; });
  var summaryText = '日付: ' + REQ_CTX.targetDate;
  if(orderLabels.length > 0){
    summaryText += ' ／ 工番: ' + orderLabels.join(', ');
  } else {
    summaryText += ' ／ 工番: なし';
  }
  document.getElementById('reqSummary').textContent = summaryText;

  document.getElementById('reqStep1').style.display = 'none';
  document.getElementById('reqStep2').style.display = '';
  document.getElementById('reqBtnNext').style.display = 'none';
  document.getElementById('reqBtnSubmit').style.display = '';
  document.getElementById('reqBtnBack').style.display = '';
}

function reqGoStep1(){
  document.getElementById('reqStep1').style.display = '';
  document.getElementById('reqStep2').style.display = 'none';
  document.getElementById('reqBtnNext').style.display = '';
  document.getElementById('reqBtnSubmit').style.display = 'none';
  document.getElementById('reqBtnBack').style.display = 'none';
}

function onReqReasonChange(){
  var val = document.getElementById('reqReason').value;
  document.getElementById('reqReasonDetailWrap').style.display =
    (val.indexOf('その他') >= 0) ? '' : 'none';
}

// --- 送信 ---
function submitRequest(){
  var reason = document.getElementById('reqReason').value;
  var reasonDetail = document.getElementById('reqReasonDetail').value.trim();
  var hours = document.getElementById('reqHours').value;
  var workContent = document.getElementById('reqWorkContent').value.trim();

  // クライアントバリデーション
  if(!reason){ alert('理由を選択してください。'); return; }
  if(reason.indexOf('その他') >= 0 && !reasonDetail){
    alert('理由が「その他」の場合、補足理由を入力してください。');
    return;
  }
  if(!hours){ alert('予定時間を選択してください。'); return; }

  var label = (REQ_CTX.type === 'overtime') ? '残業' : '休日出勤';

  // 過去申請判定：カレンダーで日付選択されていれば上書き
  var isRetro = !!selectedRetroactiveDate;
  var finalTargetDate = isRetro ? selectedRetroactiveDate : REQ_CTX.targetDate;

  var confirmMsg = label + '申請を送信しますか？\n\n'
    + '日付: ' + finalTargetDate + (isRetro ? '（過去申請）' : '') + '\n'
    + '理由: ' + reason + '\n'
    + '予定時間: ' + hours;
  if(isRetro){
    confirmMsg += '\n\n⚠ 過去日の申請です。承認者・総務への事前連絡はお済みでしょうか？';
  }
  if(!confirm(confirmMsg)){ return; }

  // ボタン無効化
  var btnSubmit = document.getElementById('reqBtnSubmit');
  btnSubmit.disabled = true;
  btnSubmit.textContent = '送信中...';

  var payload = {
    requestType: REQ_CTX.type,
    dept: REQ_CTX.dept,
    workerLabel: REQ_CTX.workerLabel,
    targetDate: finalTargetDate,
    workNo1: REQ_CTX.orders[0] || '',
    workNo2: REQ_CTX.orders[1] || '',
    workNo3: REQ_CTX.orders[2] || '',
    reason: reason,
    reasonDetail: reasonDetail,
    workContent: workContent,
    hours: hours,
    isRetroactive: isRetro,
  };

  google.script.run
    .withSuccessHandler(function(res){
      // 正常系判定: success または ok が true の場合のみ成功扱い
      var isOk = res && (res.success === true || res.ok === true);
      if(isOk){
        closeRequestModal();
        alert(label + '申請が完了しました。');
        loadDashboard();
        return;
      }
      // バリデーション失敗など: モーダルは閉じずにエラー表示し、ボタンを再活性化
      btnSubmit.disabled = false;
      btnSubmit.textContent = '申請送信';
      var errMsg = (res && res.error) ? res.error : '申請に失敗しました。管理者に連絡してください。';
      alert(errMsg);
    })
    .withFailureHandler(function(err){
      btnSubmit.disabled = false;
      btnSubmit.textContent = '申請送信';
      alert('送信エラー: ' + (err.message || err));
    })
    .api_submitRequest(payload);
}

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
