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

function api_getFormUrl(type, dept, workerLabel, targetDateStr, orderCodes) {
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

  // プリフィルURLを生成（申請種別・部署・作業員・作業日・工番を事前入力）
  return buildPrefillUrl_(formId, type, dept, workerLabel, targetDateStr, orderCodes || []);
}

/**
 * プリフィル（事前入力）付きURLを生成する。
 * 申請種別・部署・作業員を自動セットした状態でフォームを開く。
 */
function buildPrefillUrl_(formId, requestType, dept, workerLabel, targetDateStr, orderCodes) {
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

  // 工番をプリフィル（最大3件、コードをそのままテキスト欄へ）
  if (orderCodes && orderCodes.length > 0) {
    const orderQs = [Q.ORDER_1, Q.ORDER_2, Q.ORDER_3];
    for (let i = 0; i < Math.min(orderCodes.length, 3); i++) {
      const code = normalize_(orderCodes[i] || '');
      if (!code) continue;
      formResponse = addPrefill_(formResponse, form, orderQs[i], code);
    }
  }

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

  const itemResponse = item.asDateItem().createResponse(d);
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
  const calId = 'ja.japanese#holiday@group.v.calendar.google.com';
  let holidays = getHolidaysFromCalendar_(calId, monday, sunday);

  // カレンダーAPIで取得できなければ振替休日ロジックで補完
  if (holidays.length === 0) {
    holidays = getKnownJapaneseHolidays_(monday, sunday);
  }

  holidays.forEach(function(h) {
    const existing = candidates.find(function(c) { return c.date === h.date; });
    if (existing) {
      existing.holidayName = h.title;
      return;
    }
    const hd = new Date(h.date + 'T00:00:00');
    candidates.push({
      date: h.date,
      label: fmtDate_(hd, 'M/d') + '(' + dayNames[hd.getDay()] + ') ' + h.title,
      type: 'holiday'
    });
  });

  // 日付順でソート
  candidates.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  return candidates;
}

/**
 * Googleカレンダーから祝日を取得するヘルパー。
 * サブスクライブを試みてからイベント取得。
 */
function getHolidaysFromCalendar_(calId, startDate, endDate) {
  const results = [];
  try {
    // まずサブスクライブを試みる（既に登録済みならエラーを無視）
    try { CalendarApp.subscribeToCalendar(calId); } catch (_) {}

    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) return results;

    const rangeEnd = new Date(endDate);
    rangeEnd.setDate(rangeEnd.getDate() + 1);
    const events = cal.getEvents(startDate, rangeEnd);
    events.forEach(function(ev) {
      results.push({
        date: fmtDate_(ev.getStartTime()),
        title: ev.getTitle()
      });
    });
  } catch (_) {}
  return results;
}

/**
 * カレンダーAPI不可時のフォールバック：日本の祝日（法定）を算出。
 * 指定期間内に該当する祝日を返す。
 */
function getKnownJapaneseHolidays_(startDate, endDate) {
  const year = Number(Utilities.formatDate(startDate, TZ, 'yyyy'));
  const years = [year - 1, year, year + 1];

  // 固定祝日リスト
  const fixed = [
    [1, 1, '元日'],
    [1, -2, '成人の日'],       // 1月第2月曜（特殊処理）
    [2, 11, '建国記念の日'],
    [2, 23, '天皇誕生日'],
    [3, 0, '春分の日'],        // 算出（特殊処理）
    [4, 29, '昭和の日'],
    [5, 3, '憲法記念日'],
    [5, 4, 'みどりの日'],
    [5, 5, 'こどもの日'],
    [7, -3, '海の日'],         // 7月第3月曜
    [8, 11, '山の日'],
    [9, -3, '敬老の日'],       // 9月第3月曜
    [9, 0, '秋分の日'],        // 算出（特殊処理）
    [10, -2, 'スポーツの日'],  // 10月第2月曜
    [11, 3, '文化の日'],
    [11, 23, '勤労感謝の日'],
  ];

  const rangeStart = fmtDate_(startDate);
  const rangeEndD = new Date(endDate);
  rangeEndD.setDate(rangeEndD.getDate() + 1);
  const rangeEnd = fmtDate_(rangeEndD);

  const holidays = {}; // date -> title

  years.forEach(function(y) {
    fixed.forEach(function(f) {
      const month = f[0], dayOrType = f[1], title = f[2];
      let d;

      if (dayOrType < 0) {
        // 第N月曜日（-2 = 第2月曜, -3 = 第3月曜）
        d = getNthMonday_(y, month, Math.abs(dayOrType));
      } else if (dayOrType === 0 && month === 3) {
        d = new Date(y, 2, getVernalEquinox_(y));
      } else if (dayOrType === 0 && month === 9) {
        d = new Date(y, 8, getAutumnalEquinox_(y));
      } else {
        d = new Date(y, month - 1, dayOrType);
      }

      const ds = fmtDate_(d);
      if (ds >= rangeStart && ds < rangeEnd) {
        holidays[ds] = title;
      }

      // 振替休日：祝日が日曜なら翌月曜
      if (d.getDay() === 0) {
        const sub = new Date(d);
        sub.setDate(sub.getDate() + 1);
        // 翌日も祝日ならさらに翌日へ
        while (holidays[fmtDate_(sub)]) sub.setDate(sub.getDate() + 1);
        const subDs = fmtDate_(sub);
        if (subDs >= rangeStart && subDs < rangeEnd) {
          holidays[subDs] = title + '（振替休日）';
        }
      }
    });
  });

  return Object.keys(holidays).map(function(ds) {
    return { date: ds, title: holidays[ds] };
  });
}

function getNthMonday_(year, month, n) {
  var d = new Date(year, month - 1, 1);
  var count = 0;
  while (count < n) {
    if (d.getDay() === 1) count++;
    if (count < n) d.setDate(d.getDate() + 1);
  }
  return d;
}

function getVernalEquinox_(y) {
  // 春分日の近似式（1980-2099）
  return Math.floor(20.8431 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
}

function getAutumnalEquinox_(y) {
  // 秋分日の近似式（1980-2099）
  return Math.floor(23.2488 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
}
