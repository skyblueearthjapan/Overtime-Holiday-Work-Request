// ====== マスタ読込（部署/作業員/業務/工番） ======

function loadDeptList_() {
  // 作業員マスタの部署列からユニーク値を取得（部署マスタ不要）
  const workersByDept = loadWorkersByDept_();
  const depts = Array.from(workersByDept.keys()).sort();
  return depts;
}

function loadWorkersByDept_() {
  const sh = requireSheet_(SHEET.WORKERS);
  const values = sh.getDataRange().getValues();
  // 想定ヘッダ：作業員コード, 氏名, 部署（在籍フラグは任意）
  const H = values[0].map(h => normalize_(h));
  const idx = {
    code: H.indexOf('作業員コード'),
    name: H.indexOf('氏名'),
    dept: H.indexOf('部署'),
    active: H.indexOf('在籍フラグ'),   // 列が無ければ -1 → 全員有効扱い
  };
  if (idx.code < 0 || idx.name < 0 || idx.dept < 0) {
    throw new Error('作業員マスタのヘッダが想定と違います（作業員コード/氏名/部署）。');
  }

  const map = new Map(); // dept -> ["A001 今泉雄二", ...]
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dept = normalize_(row[idx.dept]);
    const code = normalize_(row[idx.code]);
    const name = normalize_(row[idx.name]);
    if (!dept || !code || !name) continue;

    // 在籍フラグ列がある場合のみフィルタ（無ければ全員有効）
    if (idx.active >= 0) {
      const active = normalize_(row[idx.active]);
      if (active && active !== '1' && active.toLowerCase() !== 'true') continue;
    }

    const label = `${code} ${name}`;
    if (!map.has(dept)) map.set(dept, []);
    map.get(dept).push(label);
  }
  // ソート（コード順）
  for (const [k, arr] of map.entries()) arr.sort();
  return map;
}

function loadJobsByDept_() {
  const sh = requireSheet_(SHEET.JOBS);
  const values = sh.getDataRange().getValues();
  // 想定ヘッダ：業務ID, 業務NO, 業務名, 部署, 説明
  const H = values[0].map(h => normalize_(h));
  const idx = {
    id: H.indexOf('業務ID'),
    no: H.indexOf('業務NO'),
    name: H.indexOf('業務名'),
    dept: H.indexOf('部署'),
  };
  if (idx.id < 0 || idx.no < 0 || idx.name < 0 || idx.dept < 0) {
    throw new Error('業務NOマスタのヘッダが想定と違います（業務ID/業務NO/業務名/部署）。');
  }

  const map = new Map(); // dept -> ["J01 001:設計", ...]
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dept = normalize_(row[idx.dept]);
    const id = normalize_(row[idx.id]);
    const no = normalize_(row[idx.no]);
    const name = normalize_(row[idx.name]);
    if (!dept || !id) continue;
    const label = `${id} ${no}:${name}`; // 表示は自由に変更OK
    if (!map.has(dept)) map.set(dept, []);
    map.get(dept).push(label);
  }
  for (const [k, arr] of map.entries()) arr.sort();
  return map;
}

function loadOrderChoices_() {
  const sh = requireSheet_(SHEET.ORDERS);
  const values = sh.getDataRange().getValues();
  // 想定ヘッダ：工番, 受注先, 納入先, 納入先住所, 品名, 数量, 取込日時
  const H = values[0].map(h => normalize_(h));
  const idx = {
    order: H.indexOf('工番'),
    dest: H.indexOf('納入先'),
    product: H.indexOf('品名'),
  };
  if (idx.order < 0) throw new Error('工番マスタのヘッダが想定と違います（工番）。');

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const order = normalize_(row[idx.order]);
    if (!order) continue;
    const dest = idx.dest >= 0 ? normalize_(row[idx.dest]) : '';
    const product = idx.product >= 0 ? normalize_(row[idx.product]) : '';
    // 工番プルダウン表示（長すぎると見づらいので最小限）
    const label = product || dest ? `${order}｜${product || ''}${product && dest ? '／' : ''}${dest || ''}` : order;
    out.push(label);
  }

  // 重複排除＋ソート
  const uniq = Array.from(new Set(out));
  uniq.sort();
  return uniq;
}

// ====== 作業者本人情報取得（Googleアカウント列があれば照合、無ければ null） ======

function api_getWorkerInfo() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return null;

  const sh = requireSheet_(SHEET.WORKERS);
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const idx = {
    code: H.indexOf('作業員コード'),
    name: H.indexOf('氏名'),
    dept: H.indexOf('部署'),
    email: H.findIndex(h => h.startsWith('Googleアカウント')),  // 列が無ければ -1
    active: H.indexOf('在籍フラグ'),                             // 列が無ければ -1
  };
  // Googleアカウント列が存在しない場合は照合不可 → null
  if (idx.email < 0) return null;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    // 在籍フラグ列がある場合のみフィルタ（無ければ全員有効）
    if (idx.active >= 0) {
      const active = normalize_(row[idx.active]);
      if (active && active !== '1' && active.toLowerCase() !== 'true') continue;
    }

    if (normalize_(row[idx.email]) === email) {
      return {
        workerCode: normalize_(row[idx.code]),
        workerName: normalize_(row[idx.name]),
        dept: normalize_(row[idx.dept]),
        email: email,
      };
    }
  }
  return null;
}
