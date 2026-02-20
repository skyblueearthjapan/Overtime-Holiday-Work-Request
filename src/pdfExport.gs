// ====== PDF_MAP：申請書フォーム テンプレートのセル位置マッピング ======
// テンプレートSSの '申請書フォーム' シートに直接値を書き込む場合に使用
const PDF_MAP = {
  createdDate: 'G1',       // 作成日
  dept: 'B4',              // 部署
  name: 'D4',              // 氏名
  typeLabelBig: 'F4',      // 残業/休日出勤
  kubun: 'C6',             // 区分（残業/半日/1日）
  targetDate: 'C7',        // 作業実施日
  startAt: 'C10',          // 開始時刻
  endAt: 'F10',            // 終了時刻
  breakMin: 'C12',         // 休憩時間（分）
  netMin: 'F12',           // 実残業/実働時間（分）
  detail: [
    { workNo: 'B18', customer: 'D18', product: 'F18' },
    { workNo: 'B19', customer: 'D19', product: 'F19' },
    { workNo: 'B20', customer: 'D20', product: 'F20' },
  ],
  workContent: 'A23',      // 業務内容
  reason: 'A29',           // 理由
  approverBox: 'F34',      // 承認者
  approverBox2: 'G34',     // 2次承認者
};

// ====== 印鑑画像ヘルパー ======

/**
 * Drive上の画像ファイルIDからBlobを取得（存在しなければnull）
 */
function getStampBlob_(fileId) {
  if (!fileId) return null;
  try {
    return DriveApp.getFileById(fileId).getBlob();
  } catch (e) {
    Logger.log('印鑑画像取得エラー (fileId=' + fileId + '): ' + e.message);
    return null;
  }
}

/**
 * メールアドレスから作業員マスタの stampFileId を引く
 */
function lookupStampFileIdByEmail_(email) {
  if (!email) return '';
  const sh = requireSheet_(SHEET.WORKERS);
  const values = sh.getDataRange().getValues();
  const H = values[0].map(h => normalize_(h));
  const emailIdx = H.findIndex(h => h.startsWith('googleアカウント') || h.startsWith('Googleアカウント'));
  const stampIdx = H.indexOf('stampfileid');
  if (emailIdx < 0 || stampIdx < 0) return '';

  const target = email.toLowerCase();
  for (let r = 1; r < values.length; r++) {
    if (normalize_(values[r][emailIdx]).toLowerCase() === target) {
      return normalize_(values[r][stampIdx]);
    }
  }
  return '';
}

// ====== Drive フォルダ作成（YYYY.MM.DD） ======

function getOrCreateDateFolder_(rootFolderId, dateObj) {
  const root = DriveApp.getFolderById(rootFolderId);
  const folderName = Utilities.formatDate(dateObj, TZ, 'yyyy.MM.dd');

  const it = root.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();

  return root.createFolder(folderName);
}

// ====== PDF生成本体（1件） ======
// テンプレSSをコピー → 操作!B3 に requestId セット → 申請書フォームシートだけPDF出力
// → Drive日付フォルダに保存 → Requests に pdfGeneratedAt/pdfFileId/pdfFolderId を記録

function generatePdfForRequest_(requestId) {
  const req = getRequestById_(requestId);
  if (!req) throw new Error('申請が見つかりません。');
  if (req.status !== 'approved') throw new Error('未承認のためPDF生成できません。');
  if (req.pdfGeneratedAt && req.pdfFileId) {
    return { already:true, pdfFileId:req.pdfFileId };
  }

  const settings = getSettings_();
  const rootFolderId = normalize_(settings['PDF_ROOT_FOLDER_ID']);
  const templateSsid = normalize_(settings['TEMPLATE_SSID']);
  if (!rootFolderId) throw new Error('Settingsに PDF_ROOT_FOLDER_ID が未設定です。');
  if (!templateSsid) throw new Error('Settingsに TEMPLATE_SSID が未設定です。');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // テンプレSSをコピー（Driveでコピー）
    const templateFile = DriveApp.getFileById(templateSsid);
    const tmpName = `TMP_${requestId}_${fmtDate_(new Date(), 'yyyyMMdd_HHmmss')}`;
    const tmpFile = templateFile.makeCopy(tmpName);
    const tmpSs = SpreadsheetApp.openById(tmpFile.getId());

    // 操作!B3 に requestId をセット
    const op = tmpSs.getSheetByName('操作');
    if (!op) throw new Error('テンプレに「操作」シートがありません。');
    op.getRange('B3').setValue(requestId);

    // 再計算待ち（XLOOKUP反映待ち）
    SpreadsheetApp.flush();
    Utilities.sleep(600);
    SpreadsheetApp.flush();

    // 出力対象シート
    const formSheet = tmpSs.getSheetByName('申請書フォーム');
    if (!formSheet) throw new Error('テンプレに「申請書フォーム」シートがありません。');

    // 保存先フォルダ（targetDate基準で日付フォルダ）
    const targetDate = req.targetDate instanceof Date ? req.targetDate : new Date(req.targetDate);
    const dateFolder = getOrCreateDateFolder_(rootFolderId, targetDate);

    // ファイル名
    const ymd = Utilities.formatDate(targetDate, TZ, 'yyyyMMdd');
    const typeLabel = req.requestType === 'overtime' ? '残業' : '休日出勤';
    const safeDept = (req.dept || '').replace(/[\\\/\:\*\?\"\<\>\|]/g, '_');
    const safeName = (req.workerName || '').toString().replace(/[\\\/\:\*\?\"\<\>\|]/g, '_');
    const pdfName = `${ymd}_${safeDept}_${safeName}_${typeLabel}.pdf`;

    // PDFエクスポート（対象シートのみ）
    const pdfBlob = exportSheetToPdfBlob_(tmpSs.getId(), formSheet.getSheetId(), pdfName);

    // Drive保存
    const pdfFile = dateFolder.createFile(pdfBlob).setName(pdfName);

    // Requestsに記録
    const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
    const now = new Date();
    sh.getRange(req.rowNo, idx['pdfGeneratedAt']+1).setValue(now);
    sh.getRange(req.rowNo, idx['pdfFileId']+1).setValue(pdfFile.getId());
    sh.getRange(req.rowNo, idx['pdfFolderId']+1).setValue(dateFolder.getId());

    // 一時コピー削除（推奨：ゴミ箱へ）
    tmpFile.setTrashed(true);

    return { ok:true, pdfFileId: pdfFile.getId(), pdfName, folderId: dateFolder.getId() };
  } finally {
    lock.releaseLock();
  }
}

// ====== シート1枚をPDF化するユーティリティ（GAS定番） ======

function exportSheetToPdfBlob_(spreadsheetId, sheetId, filename) {
  const url =
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export` +
    `?format=pdf` +
    `&gid=${sheetId}` +
    `&portrait=true` +
    `&size=A4` +
    `&fitw=true` +
    `&sheetnames=false&printtitle=false` +
    `&pagenumbers=false` +
    `&gridlines=false` +
    `&fzr=false`;

  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error(`PDFエクスポートに失敗しました: HTTP ${code} / ${res.getContentText().slice(0,200)}`);
  }

  const blob = res.getBlob().setName(filename);
  return blob;
}

// ====== Requestsの全列データ取得（PDF直接書込用） ======

function getRequestFullData_(requestId) {
  const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (normalize_(row[idx['requestId']]) === requestId) {
      const data = { rowNo: i + 2 };
      for (const [key, col] of Object.entries(idx)) {
        data[key] = row[col];
      }
      return data;
    }
  }
  return null;
}

// ====== PDF直接書込方式（XLOOKUP不要、PDF_MAPでセルに直接値を書く） ======

function generatePdfDirect_(requestId) {
  const req = getRequestFullData_(requestId);
  if (!req) throw new Error('申請が見つかりません。');

  const status = normalize_(req['status(submitted/approved/canceled)']);
  if (status !== 'approved') throw new Error('未承認のためPDF生成できません。');

  const existingPdfId = normalize_(req['pdfFileId']);
  const existingPdfAt = req['pdfGeneratedAt'];
  if (existingPdfAt && existingPdfId) {
    return { already: true, pdfFileId: existingPdfId };
  }

  const settings = getSettings_();
  const rootFolderId = normalize_(settings['PDF_ROOT_FOLDER_ID']);
  const templateSsid = normalize_(settings['TEMPLATE_SSID']);
  if (!rootFolderId) throw new Error('Settingsに PDF_ROOT_FOLDER_ID が未設定です。');
  if (!templateSsid) throw new Error('Settingsに TEMPLATE_SSID が未設定です。');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // テンプレSSをコピー
    const templateFile = DriveApp.getFileById(templateSsid);
    const tmpName = 'TMP_' + requestId + '_' + fmtDate_(new Date(), 'yyyyMMdd_HHmmss');
    const tmpFile = templateFile.makeCopy(tmpName);
    const tmpSs = SpreadsheetApp.openById(tmpFile.getId());

    // 申請書フォーム シートに直接書込み
    const formSheet = tmpSs.getSheetByName('申請書フォーム');
    if (!formSheet) throw new Error('テンプレに「申請書フォーム」シートがありません。');

    fillPdfTemplate_(formSheet, req, requestId);
    SpreadsheetApp.flush();

    // 保存先フォルダ（targetDate基準で日付フォルダ）
    const targetDateVal = req['targetDate'];
    const targetDate = targetDateVal instanceof Date ? targetDateVal : new Date(targetDateVal);
    const dateFolder = getOrCreateDateFolder_(rootFolderId, targetDate);

    // ファイル名
    const requestType = normalize_(req['requestType(overtime/holiday)']);
    const dept = normalize_(req['dept']);
    const workerName = normalize_(req['workerName']);
    const ymd = fmtDate_(targetDate, 'yyyyMMdd');
    const typeLabel = requestType === 'overtime' ? '残業' : '休日出勤';
    const safeDept = dept.replace(/[\\\/\:\*\?\"\<\>\|]/g, '_');
    const safeName = workerName.replace(/[\\\/\:\*\?\"\<\>\|]/g, '_');
    const pdfName = ymd + '_' + safeDept + '_' + safeName + '_' + typeLabel + '.pdf';

    // PDFエクスポート（対象シートのみ）
    const pdfBlob = exportSheetToPdfBlob_(tmpSs.getId(), formSheet.getSheetId(), pdfName);

    // Drive保存
    const pdfFile = dateFolder.createFile(pdfBlob).setName(pdfName);

    // Requestsに記録
    const { sh, idx } = getSheetHeaderIndex_('Requests', 1);
    const now = new Date();
    if (idx['pdfGeneratedAt'] !== undefined) sh.getRange(req.rowNo, idx['pdfGeneratedAt'] + 1).setValue(now);
    if (idx['pdfFileId'] !== undefined) sh.getRange(req.rowNo, idx['pdfFileId'] + 1).setValue(pdfFile.getId());
    if (idx['pdfFolderId'] !== undefined) sh.getRange(req.rowNo, idx['pdfFolderId'] + 1).setValue(dateFolder.getId());

    // 一時コピー削除
    tmpFile.setTrashed(true);

    return { ok: true, pdfFileId: pdfFile.getId(), pdfName: pdfName, folderId: dateFolder.getId() };
  } finally {
    lock.releaseLock();
  }
}

// ====== 申請書フォームへのセル直接書込 ======

function fillPdfTemplate_(sheet, reqData, requestId) {
  const now = new Date();
  const requestType = normalize_(reqData['requestType(overtime/holiday)']);
  const targetDateVal = reqData['targetDate'];
  const targetDate = targetDateVal instanceof Date ? targetDateVal : new Date(targetDateVal);
  const typeLabel = requestType === 'overtime' ? '残業' : '休日出勤';

  // WorkLogs データ取得
  const wlMap = buildWorkLogsMapByRequestId_();
  const wl = wlMap.get(requestId) || {};

  const settings = getSettings_();

  // 基本情報
  sheet.getRange(PDF_MAP.createdDate).setValue(fmtDate_(now, 'yyyy/MM/dd'));
  sheet.getRange(PDF_MAP.dept).setValue(normalize_(reqData['dept']));
  sheet.getRange(PDF_MAP.dept).setHorizontalAlignment('center');  // 部署名を中央揃え
  sheet.getRange(PDF_MAP.name).setValue(normalize_(reqData['workerName']));

  // F4: 区分印鑑（残業/休日出勤のPNG画像を挿入）
  sheet.getRange(PDF_MAP.typeLabelBig).setValue(typeLabel);
  const stampTypeKey = requestType === 'overtime' ? 'STAMP_OVERTIME_FILE_ID' : 'STAMP_HOLIDAY_FILE_ID';
  const stampTypeId = normalize_(settings[stampTypeKey]);
  if (stampTypeId) {
    const stampBlob = getStampBlob_(stampTypeId);
    if (stampBlob) {
      const img = sheet.insertImage(stampBlob, 6, 4); // F4
      img.setWidth(60).setHeight(60);
    }
  }

  // 区分
  if (requestType === 'overtime') {
    sheet.getRange(PDF_MAP.kubun).setValue('残業');
  } else {
    const mins = Number(reqData['approvedMinutes'] || 0);
    sheet.getRange(PDF_MAP.kubun).setValue(mins <= 240 ? '半日' : '1日');
  }

  // 日付
  sheet.getRange(PDF_MAP.targetDate).setValue(fmtDate_(targetDate, 'yyyy/MM/dd'));

  // 実績時刻（Date or JST文字列 "yyyy-MM-dd HH:mm:ss" どちらにも対応）
  const startTime = extractHHmm_(wl.actualStartAt);
  const endTime   = extractHHmm_(wl.actualEndAt);
  if (startTime) sheet.getRange(PDF_MAP.startAt).setValue(startTime);
  if (endTime)   sheet.getRange(PDF_MAP.endAt).setValue(endTime);

  // 休憩・実残業
  sheet.getRange(PDF_MAP.breakMin).setValue(Number(wl.breakMinutes || 0));
  sheet.getRange(PDF_MAP.netMin).setValue(Number(wl.netMinutes || 0));

  // 明細行（最大3行）
  for (let i = 0; i < PDF_MAP.detail.length; i++) {
    const d = PDF_MAP.detail[i];
    const suffix = String(i + 1);
    const workNo = normalize_(reqData['workNo' + suffix] || reqData['orderNo' + suffix]);
    const customer = normalize_(reqData['customer' + suffix]);
    const product = normalize_(reqData['product' + suffix]);
    if (workNo) sheet.getRange(d.workNo).setValue(workNo);
    if (customer) sheet.getRange(d.customer).setValue(customer);
    if (product) sheet.getRange(d.product).setValue(product);
  }

  // 業務内容
  const workContent = normalize_(reqData['workContent']);
  if (workContent) sheet.getRange(PDF_MAP.workContent).setValue(workContent);

  // 理由（定型理由 or 「その他: 補足理由」）
  const reason = normalize_(reqData['reason']);
  const reasonDetail = normalize_(reqData['reasonDetail']);
  let reasonText = reason;
  if (reason && reason.indexOf('その他') >= 0 && reasonDetail) {
    reasonText = 'その他: ' + reasonDetail;
  }
  if (reasonText) sheet.getRange(PDF_MAP.reason).setValue(reasonText);

  // F34: 承認者（名前＋印鑑画像）
  const approvedBy = normalize_(reqData['approvedBy']);
  if (approvedBy) {
    sheet.getRange(PDF_MAP.approverBox).setValue(approvedBy);
    const approverStampId = lookupStampFileIdByEmail_(approvedBy);
    if (approverStampId) {
      const blob = getStampBlob_(approverStampId);
      if (blob) {
        const img = sheet.insertImage(blob, 6, 34); // F34
        img.setWidth(50).setHeight(50);
      }
    }
  }

  // G34: 2次承認者（印鑑画像）※将来の2次承認フロー実装時に有効化
  const approvedBy2 = normalize_(reqData['approvedBy2'] || '');
  if (approvedBy2) {
    sheet.getRange(PDF_MAP.approverBox2).setValue(approvedBy2);
    const approver2StampId = lookupStampFileIdByEmail_(approvedBy2);
    if (approver2StampId) {
      const blob = getStampBlob_(approver2StampId);
      if (blob) {
        const img = sheet.insertImage(blob, 7, 34); // G34
        img.setWidth(50).setHeight(50);
      }
    }
  }
}

// ====== PDF一括生成（本日の承認済み＆PDF未生成を処理） ======

function batchGeneratePdfs_(dateObj) {
  const target = dateObj || new Date();
  const items = listApprovedRequestsByDate_(target);
  const results = { ok: 0, skip: 0, fail: 0, errors: [] };

  // Settings で直接書込方式かXLOOKUP方式か判定
  const settings = getSettings_();
  const useDirect = normalize_(settings['PDF_MODE']) === 'direct';

  for (const it of items) {
    if (it.pdfFileId) {
      results.skip++;
      continue;
    }
    try {
      if (useDirect) {
        generatePdfDirect_(it.requestId);
      } else {
        generatePdfForRequest_(it.requestId);
      }
      results.ok++;
    } catch (e) {
      results.fail++;
      results.errors.push(it.requestId + ' (' + it.dept + '/' + it.workerName + '): ' + e.message);
    }
    // API負荷軽減
    if (items.length > 5) Utilities.sleep(500);
  }

  return results;
}

// ====== BatchLogs 記録 ======

function logBatchResult_(batchName, dateObj, result) {
  const ss = getDb_();
  let sh = ss.getSheetByName(SHEET.BATCH_LOGS);
  if (!sh) {
    sh = ss.insertSheet(SHEET.BATCH_LOGS);
    sh.appendRow(['実行日時', 'バッチ名', '対象日', '成功', 'スキップ', '失敗', 'エラー詳細']);
  }

  const now = new Date();
  const targetYmd = fmtDate_(dateObj, 'yyyy-MM-dd');
  const errText = (result.errors && result.errors.length > 0)
    ? result.errors.join('\n')
    : '';

  sh.appendRow([
    now,
    batchName,
    targetYmd,
    result.ok || 0,
    result.skip || 0,
    result.fail || 0,
    errText,
  ]);
}

// ====== Date or JST文字列から "HH:mm" を抽出 ======

function extractHHmm_(val) {
  if (!val) return '';
  if (val instanceof Date) return fmtDate_(val, 'HH:mm');
  // JST文字列 "yyyy-MM-dd HH:mm:ss" や "HH:mm" 形式
  const s = String(val);
  const m = s.match(/(\d{2}:\d{2})/);
  return m ? m[1] : '';
}
