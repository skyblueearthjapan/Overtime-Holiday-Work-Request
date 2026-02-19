// ====== DeptApprovers シート対応（部署×承認者メール） ======
// DeptApprovers ヘッダ：A=dept, B=approverEmails（カンマ区切り可）

/**
 * DeptApprovers シートから { dept → Set<email> } を構築する。
 * DeptApprovers が無い場合は既存 ApproverMap にフォールバック。
 */
function getDeptApproverMap_() {
  const ss = getDb_();

  // 1) DeptApprovers シート（新方式）
  const sh = ss.getSheetByName(SHEET.DEPT_APPROVERS);
  if (sh && sh.getLastRow() > 1) {
    const values = sh.getDataRange().getValues();
    const map = {};
    for (let r = 1; r < values.length; r++) {
      const dept = normalize_(values[r][0]);
      const raw = normalize_(values[r][1]);
      if (!dept) continue;
      const emails = raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
      if (!map[dept]) map[dept] = new Set();
      for (const em of emails) map[dept].add(em);
    }
    return map;
  }

  // 2) ApproverMap にフォールバック（従来方式）
  const amSh = ss.getSheetByName('ApproverMap');
  if (!amSh) return {};
  const values = amSh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idx = {
    dept: H.indexOf('部署'),
    mail: H.indexOf('承認者メール'),
    enabled: H.indexOf('有効フラグ'),
  };
  if (idx.dept < 0 || idx.mail < 0) return {};

  const map = {};
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const enabled = idx.enabled >= 0 ? row[idx.enabled] : true;
    if (enabled === false || String(enabled).toLowerCase() === 'false') continue;
    const dept = normalize_(row[idx.dept]);
    const mail = normalize_(row[idx.mail]).toLowerCase();
    if (!dept || !mail) continue;
    if (!map[dept]) map[dept] = new Set();
    map[dept].add(mail);
  }
  return map;
}

/**
 * 総務（管理者）メールリストを取得する。
 * Settings の GENERAL_AFFAIRS_EMAILS → ADMIN_EMAILS の順で探す。
 */
function getGeneralAffairsEmails_() {
  try {
    const settings = getSettings_();
    // GENERAL_AFFAIRS_EMAILS を優先
    const gaRaw = normalize_(settings['GENERAL_AFFAIRS_EMAILS']);
    if (gaRaw) {
      return gaRaw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    }
    // ADMIN_EMAILS にフォールバック
    const adminRaw = normalize_(settings['ADMIN_EMAILS']);
    if (adminRaw) {
      return adminRaw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    }
  } catch (_) {}
  return [];
}

/**
 * 総務部（管理者）権限を持つかチェック（例外を投げない版）。
 */
function isGeneralAffairs_(email) {
  if (!email) return false;
  const list = getGeneralAffairsEmails_();
  return list.includes(email.toLowerCase());
}

/**
 * 部署一覧を返す API（TOPの部署選択用・誰でも呼べる）。
 */
function api_getDeptList() {
  return loadDeptList_();
}

/**
 * 指定部署の作業員一覧を返す API（TOPの作業員選択用）。
 * 戻り値: ["A001 今泉雄二", "A002 田中太郎", ...]
 */
function api_getWorkersByDept(dept) {
  if (!dept) return [];
  const workersByDept = loadWorkersByDept_();
  return workersByDept.get(dept) || [];
}
