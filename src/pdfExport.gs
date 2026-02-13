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
