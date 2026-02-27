// ====== DeptApprovers シート対応（部署×承認者メール） ======
// DeptApprovers ヘッダ：A=dept, B=approverEmails（カンマ区切り可）

/**
 * DeptApprovers シートから { dept → Set<email> } を構築する。
 * DeptApprovers が無い場合は既存 ApproverMap にフォールバック。
 */
var _deptApproverMapCache = null;

function getDeptApproverMap_() {
  if (_deptApproverMapCache) return _deptApproverMapCache;
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
    _deptApproverMapCache = map;
    return map;
  }

  // 2) ApproverMap にフォールバック（従来方式）
  const amSh = ss.getSheetByName('ApproverMap');
  if (!amSh) { _deptApproverMapCache = {}; return {}; }
  const values = amSh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idx = {
    dept: H.indexOf('部署'),
    mail: H.indexOf('承認者メール'),
    enabled: H.indexOf('有効フラグ'),
  };
  if (idx.dept < 0 || idx.mail < 0) { _deptApproverMapCache = {}; return {}; }

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
  _deptApproverMapCache = map;
  return map;
}

/**
 * DeptApproversシートから承認者プロファイル一覧を構築する。
 * 戻り値: [{ email: 'tanaka@...', name: '田中太郎', depts: ['製造1課','品質管理課'] }]
 */
function getApproverProfiles_() {
  const sh = getDb_().getSheetByName(SHEET.DEPT_APPROVERS);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  var map = {};  // email → { name, depts }
  for (var r = 1; r < values.length; r++) {
    var dept = normalize_(values[r][0]);
    var rawEmails = normalize_(values[r][1]);
    var name = (values.length > 0 && values[r].length > 2) ? normalize_(values[r][2]) : '';
    if (!dept || !rawEmails) continue;
    // カンマ区切りのメール対応（既存形式の後方互換）
    var emails = rawEmails.split(',').map(function(s){ return s.trim().toLowerCase(); }).filter(Boolean);
    for (var e = 0; e < emails.length; e++) {
      var email = emails[e];
      if (!map[email]) map[email] = { email: email, name: name || email, depts: [] };
      if (map[email].depts.indexOf(dept) < 0) map[email].depts.push(dept);
      if (name && name !== email) map[email].name = name;
    }
  }
  return Object.values(map);
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
