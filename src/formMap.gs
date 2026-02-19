// ====== FormMap の自動作成/取得 ======

function ensureFormMapSheet_() {
  const ss = getDb_();
  let sh = ss.getSheetByName(SHEET.FORM_MAP);
  if (sh) return sh;

  // FormMap が無ければ自動作成（ヘッダ付き）
  sh = ss.insertSheet(SHEET.FORM_MAP);
  sh.appendRow(['type', 'dept', 'formId', 'formUrl', 'updatedAt', 'isActive']);
  return sh;
}

// ====== FormMap の検索 ======

function findFormMapRow_(type, dept) {
  const sh = ensureFormMapSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null; // ヘッダのみ → 該当なし

  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idx = {
    type: H.indexOf('type'),
    dept: H.indexOf('dept'),
    formId: H.indexOf('formId'),
    formUrl: H.indexOf('formUrl'),
    updatedAt: H.indexOf('updatedAt'),
    isActive: H.indexOf('isActive'),
  };
  if (idx.type < 0 || idx.dept < 0 || idx.formId < 0) {
    throw new Error('FormMapのヘッダが想定と違います（type/dept/formId/formUrl...）');
  }

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (normalize_(row[idx.type]) === type && normalize_(row[idx.dept]) === dept) {
      const isActive = idx.isActive >= 0 ? row[idx.isActive] : true;
      return { rowIndex: r + 1, idx, row, isActive };
    }
  }
  return null;
}

// ====== FormMap の書き込み ======

function upsertFormMap_(type, dept, formId, formUrl) {
  const sh = ensureFormMapSheet_();
  const found = findFormMapRow_(type, dept);
  const now = new Date();
  if (found) {
    const r = found.rowIndex;
    sh.getRange(r, found.idx.formId + 1).setValue(formId);
    if (found.idx.formUrl >= 0) sh.getRange(r, found.idx.formUrl + 1).setValue(formUrl);
    if (found.idx.updatedAt >= 0) sh.getRange(r, found.idx.updatedAt + 1).setValue(now);
    if (found.idx.isActive >= 0) sh.getRange(r, found.idx.isActive + 1).setValue(true);
  } else {
    sh.appendRow([type, dept, formId, formUrl, now, true]);
  }
}

// ====== テンプレフォームID取得 ======
// Settings → FormTemplates シートの順で探す

function getTemplateFormId_(type) {
  // 1) Settings から取得（OVERTIME_TEMPLATE_FORM_ID / HOLIDAY_TEMPLATE_FORM_ID）
  const settingsKey = (type === 'overtime')
    ? 'OVERTIME_TEMPLATE_FORM_ID'
    : 'HOLIDAY_TEMPLATE_FORM_ID';
  try {
    const settings = getSettings_();
    const id = normalize_(settings[settingsKey]);
    if (id) return id;
  } catch (_) { /* Settings取得失敗は無視して次へ */ }

  // 2) FormTemplates シートから取得（従来方式）
  const ss = getDb_();
  const sh = ss.getSheetByName(SHEET.FORM_TEMPLATES);
  if (!sh) {
    throw new Error(
      'テンプレフォームIDが見つかりません。\n' +
      `Settings に ${settingsKey} を設定するか、FormTemplates シートを用意してください。`
    );
  }
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idxType = H.indexOf('type');
  const idxId = H.indexOf('templateFormId');
  if (idxType < 0 || idxId < 0) throw new Error('FormTemplatesのヘッダが想定と違います（type/templateFormId）。');

  for (let r = 1; r < values.length; r++) {
    if (normalize_(values[r][idxType]) === type) {
      const id = normalize_(values[r][idxId]);
      if (!id) throw new Error(`FormTemplatesにtemplateFormIdがありません: type=${type}`);
      return id;
    }
  }
  throw new Error(`FormTemplatesにtypeがありません: ${type}`);
}

// ====== フォームURL取得（TOP画面用：未作成なら自動生成） ======

function api_getFormUrl(type, dept, workerLabel, targetDateStr) {
  let formId = '';

  // 既存を探す
  const found = findFormMapRow_(type, dept);
  if (found) {
    const isActive = found.isActive;
    if (isActive !== false && String(isActive).toLowerCase() !== 'false') {
      formId = normalize_(found.row[found.idx.formId]);
    }
  }

  // 未作成 → 自動生成を試みる
  if (!formId) {
    const result = getOrCreateDeptForm_(type, dept);
    if (!result) return null;
    formId = result.formId;
  }

  if (!formId) return null;

  // プリフィルURLを生成（申請種別・部署・作業員・作業日を事前入力）
  return buildPrefillUrl_(formId, type, dept, workerLabel, targetDateStr);
}

/**
 * プリフィル（事前入力）付きURLを生成する。
 * 申請種別・部署・作業員を自動セットした状態でフォームを開く。
 */
function buildPrefillUrl_(formId, requestType, dept, workerLabel, targetDateStr) {
  const form = FormApp.openById(formId);
  let formResponse = form.createResponse();

  // 申請種別をプリフィル（"残業" or "休日"）
  const typeLabel = (requestType === 'overtime') ? '残業' : '休日';
  formResponse = addPrefill_(formResponse, form, Q.TYPE, typeLabel);

  // 部署をプリフィル
  formResponse = addPrefill_(formResponse, form, Q.DEPT, dept);

  // 作業員をプリフィル
  if (workerLabel) {
    formResponse = addPrefill_(formResponse, form, Q.WORKER, workerLabel);
  }

  // 作業実施日をプリフィル（指定日 or 今日）
  formResponse = addDatePrefill_(formResponse, form, Q.DATE, targetDateStr);

  return formResponse.toPrefilledUrl();
}

/**
 * フォーム回答オブジェクトにプリフィルを追加するヘルパー。
 * 質問が見つからない場合やタイプ不一致の場合はスキップ。
 */
function addPrefill_(formResponse, form, questionTitle, value) {
  const item = findItemByTitleOrNull_(form, questionTitle);
  if (!item || !value) return formResponse;

  const type = item.getType();
  let itemResponse;
  if (type === FormApp.ItemType.LIST) {
    itemResponse = item.asListItem().createResponse(value);
  } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    itemResponse = item.asMultipleChoiceItem().createResponse(value);
  } else if (type === FormApp.ItemType.TEXT) {
    itemResponse = item.asTextItem().createResponse(value);
  } else {
    return formResponse; // 未対応タイプはスキップ
  }

  return formResponse.withItemResponse(itemResponse);
}

/**
 * 日付フィールドをプリフィルするヘルパー。
 * targetDateStr が指定されればその日付、なければ今日。
 */
function addDatePrefill_(formResponse, form, questionTitle, targetDateStr) {
  const item = findItemByTitleOrNull_(form, questionTitle);
  if (!item) return formResponse;

  if (item.getType() !== FormApp.ItemType.DATE) return formResponse;

  let d;
  if (targetDateStr) {
    // 'yyyy-MM-dd' 形式を想定
    const parts = String(targetDateStr).split('-');
    d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  } else {
    d = new Date();
  }

  const year = Number(Utilities.formatDate(d, TZ, 'yyyy'));
  const month = Number(Utilities.formatDate(d, TZ, 'M'));
  const day = Number(Utilities.formatDate(d, TZ, 'd'));

  const itemResponse = item.asDateItem().createResponse(year, month, day);
  return formResponse.withItemResponse(itemResponse);
}

// ====== 休日出勤：今週の候補日（土日＋祝日）を取得 ======

function api_getHolidayCandidateDates() {
  const today = new Date();
  const dow = today.getDay(); // 0=Sun..6=Sat

  // 今週の月曜日を算出
  const monday = new Date(today);
  monday.setDate(today.getDate() - (dow === 0 ? 6 : dow - 1));
  monday.setHours(0, 0, 0, 0);

  // 今週の土曜・日曜
  const saturday = new Date(monday);
  saturday.setDate(monday.getDate() + 5);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);

  const dayNames = ['日','月','火','水','木','金','土'];
  const candidates = [];

  candidates.push({
    date: fmtDate_(saturday),
    label: fmtDate_(saturday, 'M/d') + '(土)',
    type: 'weekend'
  });
  candidates.push({
    date: fmtDate_(sunday),
    label: fmtDate_(sunday, 'M/d') + '(日)',
    type: 'weekend'
  });

  // Googleカレンダーの日本の祝日を検索
  try {
    const calId = 'ja.japanese#holiday@group.v.calendar.google.com';
    const cal = CalendarApp.getCalendarById(calId);
    if (cal) {
      const rangeEnd = new Date(sunday);
      rangeEnd.setDate(rangeEnd.getDate() + 1);
      const events = cal.getEvents(monday, rangeEnd);
      events.forEach(function(ev) {
        const d = ev.getStartTime();
        const ds = fmtDate_(d);
        // 既に候補にある日（土日が祝日の場合）は祝日名を追加
        const existing = candidates.find(function(c) { return c.date === ds; });
        if (existing) {
          existing.holidayName = ev.getTitle();
          return;
        }
        // 平日の祝日を追加
        candidates.push({
          date: ds,
          label: fmtDate_(d, 'M/d') + '(' + dayNames[d.getDay()] + ') ' + ev.getTitle(),
          type: 'holiday'
        });
      });
    }
  } catch (_) {
    // カレンダーアクセス失敗時は土日のみ表示
  }

  // 日付順でソート
  candidates.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  return candidates;
}
