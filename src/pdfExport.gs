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
 * 大きい画像はDriveサムネイルAPIで縮小して取得（2MB/100万px制限回避）
 */
function getStampBlob_(fileId) {
  if (!fileId) return null;
  try {
    var blob = DriveApp.getFileById(fileId).getBlob();
    // 2MB未満ならそのまま返す
    if (blob.getBytes().length < 2 * 1024 * 1024) return blob;
  } catch (e) {
    Logger.log('印鑑画像取得エラー (fileId=' + fileId + '): ' + e.message);
    return null;
  }
  // 大きい場合はDriveサムネイルで縮小取得
  try {
    Logger.log('印鑑画像が大きいためサムネイル取得: fileId=' + fileId);
    var url = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=400';
    var res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() === 200) return res.getBlob();
    Logger.log('サムネイル取得失敗: HTTP ' + res.getResponseCode());
  } catch (e) {
    Logger.log('サムネイル取得エラー: ' + e.message);
  }
  return null;
}

/**
 * メールアドレスから StampMap シートの stampFileId を引く
 * StampMap ヘッダ想定: メール, stampFileId, 備考
 */
function lookupStampFileIdByEmail_(email) {
  if (!email) return '';
  try {
    const sh = requireSheet_(SHEET.STAMP_MAP);
    const values = sh.getDataRange().getValues();
    const H = values[0].map(h => normalize_(h));
    const HL = H.map(h => h.toLowerCase());
    const emailIdx = H.indexOf('メール');
    const stampIdx = HL.indexOf('stampfileid');
    if (emailIdx < 0 || stampIdx < 0) {
      Logger.log('StampMap: ヘッダ不一致 — emailIdx=' + emailIdx + ' stampIdx=' + stampIdx + ' ヘッダ=[' + H.join(', ') + ']');
      return '';
    }

    const target = email.toLowerCase();
    for (let r = 1; r < values.length; r++) {
      if (normalize_(values[r][emailIdx]).toLowerCase() === target) {
        return normalize_(values[r][stampIdx]);
      }
    }
  } catch (e) {
    Logger.log('StampMap 参照エラー: ' + e.message);
  }
  return '';
}

// ====== 理由テキスト生成（共通ヘルパー） ======

function buildReasonText_(reason, reasonDetail) {
  var r = reason || '';
  var d = reasonDetail || '';
  if (!r && !d) return '';
  // 「その他」系の理由 → 補足理由を結合
  if (r.indexOf('その他') >= 0 && d) {
    return r + '\n' + d;
  }
  // 通常理由だが補足もある場合 → 補足も付与
  if (r && d) {
    return r + '\n' + d;
  }
  return r || d;
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

    // 理由を直接書込（XLOOKUPでは理由列を正しく参照できない場合の保険）
    const reasonVal = normalize_(req.reason || '');
    const reasonDetailVal = normalize_(req.reasonDetail || '');
    Logger.log('[PDF/XLOOKUP] reason=[' + reasonVal + '] reasonDetail=[' + reasonDetailVal + ']');
    var reasonText = buildReasonText_(reasonVal, reasonDetailVal);
    formSheet.getRange(PDF_MAP.reason).setValue(reasonText);

    // 残業の開始時刻を直接書込（WorkLogsのTZバグ回避）
    if (req.requestType === 'overtime') {
      formSheet.getRange(PDF_MAP.startAt).setValue('17:20');
    }

    // 印鑑画像挿入（XLOOKUPでは画像挿入不可のため直接処理）
    const requestType = req.requestType;
    const stampTypeKey = requestType === 'overtime' ? 'STAMP_OVERTIME_FILE_ID' : 'STAMP_HOLIDAY_FILE_ID';
    const stampTypeId = normalize_(settings[stampTypeKey]);
    if (stampTypeId) {
      const stampBlob = getStampBlob_(stampTypeId);
      if (stampBlob) {
        try {
          const stampSize = 270; // 区分印鑑を大きく（3倍）
          const img = formSheet.insertImage(stampBlob, 6, 4, 0, 0); // F4: 区分印鑑
          img.setWidth(stampSize).setHeight(stampSize);
        } catch (e) { Logger.log('区分印鑑挿入エラー: ' + e.message); }
      }
    } else {
      // 印鑑画像がない場合、テキストのフォントを大きくする
      try {
        formSheet.getRange('F4').setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
      } catch (e) { /* ignore */ }
    }
    // 承認者印鑑（メールアドレスは書き込まない）
    const approvedBy = req.approvedBy || '';
    if (approvedBy) {
      const approverStampId = lookupStampFileIdByEmail_(approvedBy);
      if (approverStampId) {
        const blob = getStampBlob_(approverStampId);
        if (blob) {
          try {
            const aStampSize = 60;
            const aColW = formSheet.getColumnWidth(6);
            // 承認欄が複数行にまたがる場合を考慮（F34〜F37の合計高さで中央計算）
            var aTotalH = 0;
            for (var ar = 34; ar <= 37; ar++) aTotalH += formSheet.getRowHeight(ar);
            const aOffX = Math.max(0, Math.floor((aColW - aStampSize) / 2));
            const aOffY = Math.max(0, Math.floor((aTotalH - aStampSize) / 2));
            const img = formSheet.insertImage(blob, 6, 34, aOffX, aOffY); // F34 中央配置
            img.setWidth(aStampSize).setHeight(aStampSize);
          } catch (e) { Logger.log('承認者印鑑挿入エラー: ' + e.message); }
        }
      }
    }

    SpreadsheetApp.flush();

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
  Logger.log('[PDF] Requests headers: ' + Object.keys(idx).join(', '));
  Logger.log('[PDF] "reason" in idx? ' + (idx['reason'] !== undefined) + ', "reasonDetail" in idx? ' + (idx['reasonDetail'] !== undefined));
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
      Logger.log('[PDF] Found request row=' + (i+2) + ', reason=' + JSON.stringify(data['reason']) + ', reasonDetail=' + JSON.stringify(data['reasonDetail']));
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
      const stampSize = 270; // 区分印鑑を大きく（3倍）
      const img = sheet.insertImage(stampBlob, 6, 4, 0, 0); // F4: 区分印鑑
      img.setWidth(stampSize).setHeight(stampSize);
    }
  } else {
    // 印鑑画像がない場合、テキストのフォントを大きくする
    sheet.getRange(PDF_MAP.typeLabelBig).setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
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

  // 実績時刻
  if (requestType === 'overtime') {
    // 残業の開始は常に17:20固定（WorkLogsのTZバグ回避）
    sheet.getRange(PDF_MAP.startAt).setValue('17:20');
  } else {
    const startTime = extractHHmm_(wl.actualStartAt);
    if (startTime) sheet.getRange(PDF_MAP.startAt).setValue(startTime);
  }
  const endTime = extractHHmm_(wl.actualEndAt);
  if (endTime) sheet.getRange(PDF_MAP.endAt).setValue(endTime);

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
  Logger.log('[PDF/Direct] reason=[' + reason + '] reasonDetail=[' + reasonDetail + ']');
  // 空でもsetValueして、テンプレのXLOOKUP数式をクリアする
  sheet.getRange(PDF_MAP.reason).setValue(buildReasonText_(reason, reasonDetail));

  // F34: 承認者（印鑑画像のみ、メールアドレスは書き込まない）
  const approvedBy = normalize_(reqData['approvedBy']);
  if (approvedBy) {
    const approverStampId = lookupStampFileIdByEmail_(approvedBy);
    if (approverStampId) {
      const blob = getStampBlob_(approverStampId);
      if (blob) {
        const aStampSize = 60;
        const aColW = sheet.getColumnWidth(6);
        // 承認欄が複数行にまたがる場合を考慮（F34〜F37の合計高さで中央計算）
        var aTotalH = 0;
        for (var ar = 34; ar <= 37; ar++) aTotalH += sheet.getRowHeight(ar);
        const aOffX = Math.max(0, Math.floor((aColW - aStampSize) / 2));
        const aOffY = Math.max(0, Math.floor((aTotalH - aStampSize) / 2));
        const img = sheet.insertImage(blob, 6, 34, aOffX, aOffY); // F34 中央配置
        img.setWidth(aStampSize).setHeight(aStampSize);
      }
    }
  }

  // G34: 2次承認者（印鑑画像のみ）※将来の2次承認フロー実装時に有効化
  const approvedBy2 = normalize_(reqData['approvedBy2'] || '');
  if (approvedBy2) {
    const approver2StampId = lookupStampFileIdByEmail_(approvedBy2);
    if (approver2StampId) {
      const blob = getStampBlob_(approver2StampId);
      if (blob) {
        const a2StampSize = 60;
        const a2ColW = sheet.getColumnWidth(7);
        var a2TotalH = 0;
        for (var a2r = 34; a2r <= 37; a2r++) a2TotalH += sheet.getRowHeight(a2r);
        const a2OffX = Math.max(0, Math.floor((a2ColW - a2StampSize) / 2));
        const a2OffY = Math.max(0, Math.floor((a2TotalH - a2StampSize) / 2));
        const img = sheet.insertImage(blob, 7, 34, a2OffX, a2OffY); // G34 中央配置
        img.setWidth(a2StampSize).setHeight(a2StampSize);
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
  const useDirect = normalize_(settings['PDF_MODE']).indexOf('direct') >= 0;

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

// ====== 診断: PDF理由フィールドのデバッグ（手動実行用） ======

function debugPdfReason() {
  Logger.log('===== PDF理由フィールド診断 =====');

  // 1. Requests シートのヘッダー一覧
  Logger.log('\n--- 1. Requests ヘッダー ---');
  const reqSh = requireSheet_('Requests');
  const reqHeaders = reqSh.getRange(1, 1, 1, reqSh.getLastColumn()).getValues()[0];
  reqHeaders.forEach(function(h, i) {
    Logger.log('  col ' + (i+1) + ': [' + h + '] → normalize: [' + normalize_(h) + ']');
  });

  const normHeaders = reqHeaders.map(function(h) { return normalize_(h); });
  var reasonColIdx = normHeaders.indexOf('reason');
  var reasonDetailColIdx = normHeaders.indexOf('reasonDetail');
  Logger.log('  reason 列インデックス: ' + reasonColIdx + ' (-1 = 存在しない!)');
  Logger.log('  reasonDetail 列インデックス: ' + reasonDetailColIdx + ' (-1 = 存在しない!)');

  // 理由っぽい列を探す（日本語ヘッダーの可能性）
  Logger.log('\n--- 1b. 理由に関連しそうなヘッダー ---');
  normHeaders.forEach(function(h, i) {
    if (h.indexOf('理由') >= 0 || h.indexOf('reason') >= 0 || h.indexOf('Reason') >= 0) {
      Logger.log('  col ' + (i+1) + ': [' + reqHeaders[i] + '] → normalize: [' + h + ']');
    }
  });

  // 2. 最新の承認済み申請の reason データ
  Logger.log('\n--- 2. 最新承認済み申請の理由データ ---');
  var lastRow = reqSh.getLastRow();
  if (lastRow >= 2) {
    var allData = reqSh.getRange(2, 1, lastRow - 1, reqSh.getLastColumn()).getValues();
    var statusColIdx = normHeaders.indexOf('status(submitted/approved/canceled)');
    var found = false;
    // 最新から逆順で探す
    for (var r = allData.length - 1; r >= 0; r--) {
      var st = normalize_(allData[r][statusColIdx]);
      if (st === 'approved') {
        Logger.log('  行番号: ' + (r + 2));
        if (reasonColIdx >= 0) {
          Logger.log('  reason 列の値: [' + allData[r][reasonColIdx] + '] (type=' + typeof allData[r][reasonColIdx] + ')');
        } else {
          Logger.log('  reason 列が存在しません！');
        }
        if (reasonDetailColIdx >= 0) {
          Logger.log('  reasonDetail 列の値: [' + allData[r][reasonDetailColIdx] + '] (type=' + typeof allData[r][reasonDetailColIdx] + ')');
        }
        // 全列ダンプ（該当行）
        Logger.log('  全列ダンプ:');
        for (var c = 0; c < reqHeaders.length; c++) {
          var v = allData[r][c];
          if (v !== '' && v !== null && v !== undefined) {
            Logger.log('    [' + reqHeaders[c] + '] = [' + v + ']');
          }
        }
        found = true;
        break;
      }
    }
    if (!found) Logger.log('  承認済み申請が見つかりません。');
  }

  // 3. テンプレートの A29 セル構造
  Logger.log('\n--- 3. テンプレート 申請書フォーム A29 の構造 ---');
  try {
    var settings = getSettings_();
    var templateSsid = normalize_(settings['TEMPLATE_SSID']);
    if (!templateSsid) {
      Logger.log('  TEMPLATE_SSID が未設定です。');
    } else {
      var tmpSs = SpreadsheetApp.openById(templateSsid);
      var formSheet = tmpSs.getSheetByName('申請書フォーム');
      if (!formSheet) {
        Logger.log('  申請書フォーム シートが見つかりません。');
      } else {
        var a29 = formSheet.getRange('A29');
        Logger.log('  A29 現在の値: [' + a29.getValue() + ']');
        Logger.log('  A29 数式: [' + a29.getFormula() + ']');
        Logger.log('  A29 表示値: [' + a29.getDisplayValue() + ']');
        Logger.log('  A29 isPartOfMerge: ' + a29.isPartOfMerge());

        // マージ範囲の確認
        if (a29.isPartOfMerge()) {
          var merges = formSheet.getRange('A25:H35').getMergedRanges();
          for (var m = 0; m < merges.length; m++) {
            var mr = merges[m];
            if (mr.getRow() <= 29 && mr.getLastRow() >= 29 &&
                mr.getColumn() <= 1 && mr.getLastColumn() >= 1) {
              Logger.log('  A29 を含むマージ範囲: ' + mr.getA1Notation());
              Logger.log('  マージ左上セル: 行' + mr.getRow() + ' 列' + mr.getColumn());
            }
          }
        }

        // A29 周辺のセルも確認（A27〜A31）
        Logger.log('\n  A27〜A31 のセル値:');
        for (var row = 27; row <= 31; row++) {
          var cell = formSheet.getRange(row, 1);
          var val = cell.getValue();
          var formula = cell.getFormula();
          var merged = cell.isPartOfMerge();
          Logger.log('  A' + row + ': 値=[' + val + '] 数式=[' + formula + '] マージ=' + merged);
        }
      }
    }
  } catch (e) {
    Logger.log('  テンプレート読取エラー: ' + e.message);
  }

  Logger.log('\n===== 診断完了 =====');
}

// ====== PDF強制再生成（最新の承認済み申請、手動実行用） ======

function forceRegeneratePdf() {
  Logger.log('===== PDF強制再生成 =====');

  // 最新の承認済み申請を探す
  var reqSh = requireSheet_('Requests');
  var headers = reqSh.getRange(1, 1, 1, reqSh.getLastColumn()).getValues()[0];
  var normH = headers.map(function(h) { return normalize_(h); });
  var idx = {};
  normH.forEach(function(h, i) { if (h) idx[h] = i; });

  var statusCol = idx['status(submitted/approved/canceled)'];
  var ridCol = idx['requestId'];
  var pdfAtCol = idx['pdfGeneratedAt'];
  var pdfIdCol = idx['pdfFileId'];
  var pdfFolderCol = idx['pdfFolderId'];

  var lastRow = reqSh.getLastRow();
  if (lastRow < 2) { Logger.log('Requestsにデータなし'); return; }

  var data = reqSh.getRange(2, 1, lastRow - 1, reqSh.getLastColumn()).getValues();
  var targetRow = -1;
  var requestId = '';

  // 最新から逆順で承認済みを探す
  for (var r = data.length - 1; r >= 0; r--) {
    if (normalize_(data[r][statusCol]) === 'approved') {
      targetRow = r + 2;
      requestId = normalize_(data[r][ridCol]);
      Logger.log('対象: 行' + targetRow + ' requestId=' + requestId);
      Logger.log('  reason=[' + data[r][idx['reason']] + ']');
      Logger.log('  reasonDetail=[' + data[r][idx['reasonDetail']] + ']');
      Logger.log('  approvedBy=[' + data[r][idx['approvedBy']] + ']');
      Logger.log('  既存pdfFileId=[' + data[r][pdfIdCol] + ']');
      break;
    }
  }

  if (targetRow < 0) { Logger.log('承認済み申請が見つかりません'); return; }

  // WorkLogs の actualStartAt 時刻バグ修復（残業で"02:20"等のJSTずれを修正）
  var requestType = idx['requestType(overtime/holiday)'] !== undefined
    ? normalize_(data[targetRow - 2][idx['requestType(overtime/holiday)']]) : '';
  if (requestType === 'overtime') {
    try {
      var wlSh = requireSheet_('WorkLogs');
      var wlLastRow = wlSh.getLastRow();
      if (wlLastRow >= 3) {
        var wlHeader = wlSh.getRange(2, 1, 1, wlSh.getLastColumn()).getValues()[0].map(function(h) { return normalize_(h); });
        var wlIdx = buildHeaderIndex_(wlHeader);
        var wlData = wlSh.getRange(3, 1, wlLastRow - 2, wlSh.getLastColumn()).getValues();
        for (var w = 0; w < wlData.length; w++) {
          if (normalize_(wlData[w][wlIdx['requestId']]) === requestId) {
            var wlRow = w + 3;
            // actualStartAt を正しい17:20 JSTに修正
            var targetDateVal = data[targetRow - 2][idx['targetDate']];
            var td = targetDateVal instanceof Date ? targetDateVal : new Date(targetDateVal);
            var ymdStr = fmtDate_(td, 'yyyy-MM-dd');
            var correctStart = ymdStr + ' 17:20:00';
            var currentStart = wlData[w][wlIdx['actualStartAt']];
            Logger.log('WorkLogs actualStartAt: 現在=[' + currentStart + '] → 修正=[' + correctStart + ']');
            wlSh.getRange(wlRow, wlIdx['actualStartAt'] + 1).setValue(correctStart);
            SpreadsheetApp.flush();
            break;
          }
        }
      }
    } catch (wlErr) {
      Logger.log('WorkLogs修正スキップ: ' + wlErr.message);
    }
  }

  // reasonDetail が空の場合、フォーム回答から補完を試みる
  var reasonDetailCol = idx['reasonDetail'];
  var currentReasonDetail = reasonDetailCol !== undefined ? data[targetRow - 2][reasonDetailCol] : '';
  var reasonCol = idx['reason'];
  var currentReason = reasonCol !== undefined ? normalize_(data[targetRow - 2][reasonCol]) : '';
  if ((!currentReasonDetail || String(currentReasonDetail).trim() === '') && currentReason.indexOf('その他') >= 0) {
    Logger.log('reasonDetail が空 & 理由が「その他」→ フォーム回答から補完を試みます');
    try {
      var supplemented = tryFillReasonDetailFromForm_(requestId, targetRow, reqSh, idx);
      if (supplemented) {
        Logger.log('reasonDetail 補完成功: [' + supplemented + ']');
      } else {
        Logger.log('reasonDetail 補完失敗 — Requestsシートの reasonDetail 列に手動で入力してください');
      }
    } catch (rdErr) {
      Logger.log('reasonDetail 補完エラー: ' + rdErr.message);
    }
  }

  // pdfGeneratedAt / pdfFileId / pdfFolderId をクリア（再生成を許可）
  if (pdfAtCol !== undefined) reqSh.getRange(targetRow, pdfAtCol + 1).setValue('');
  if (pdfIdCol !== undefined) reqSh.getRange(targetRow, pdfIdCol + 1).setValue('');
  if (pdfFolderCol !== undefined) reqSh.getRange(targetRow, pdfFolderCol + 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('pdfGeneratedAt/pdfFileId/pdfFolderIdをクリア → 再生成開始');

  // PDF_MODE に応じて再生成
  var settings = getSettings_();
  var useDirect = normalize_(settings['PDF_MODE']).indexOf('direct') >= 0;
  Logger.log('PDF_MODE: [' + (settings['PDF_MODE'] || '(未設定→XLOOKUP)') + '] → ' + (useDirect ? 'Direct方式' : 'XLOOKUP方式'));

  try {
    var result;
    if (useDirect) {
      result = generatePdfDirect_(requestId);
    } else {
      result = generatePdfForRequest_(requestId);
    }
    Logger.log('再生成成功: ' + JSON.stringify(result));
  } catch (e) {
    Logger.log('再生成エラー: ' + e.message + '\n' + e.stack);
  }

  Logger.log('===== 完了 =====');
}

// ====== 診断: 印鑑データのデバッグ（手動実行用） ======

function debugStampData() {
  Logger.log('===== 印鑑データ診断 =====');

  // 1. Settings の印鑑関連キー
  Logger.log('\n--- 1. Settings 印鑑キー ---');
  try {
    var settings = getSettings_();
    var keys = ['STAMP_OVERTIME_FILE_ID', 'STAMP_HOLIDAY_FILE_ID'];
    for (var k = 0; k < keys.length; k++) {
      var val = settings[keys[k]];
      Logger.log('  ' + keys[k] + ': [' + (val || '(未設定)') + ']');
      if (val) {
        try {
          var file = DriveApp.getFileById(normalize_(val));
          Logger.log('    → ファイル存在: ' + file.getName() + ' (' + file.getMimeType() + ')');
        } catch (e) {
          Logger.log('    → ファイル取得エラー: ' + e.message);
        }
      }
    }
  } catch (e) {
    Logger.log('  Settings読取エラー: ' + e.message);
  }

  // 2. StampMap シートの構造
  Logger.log('\n--- 2. StampMap シート ---');
  try {
    var sh = requireSheet_(SHEET.STAMP_MAP);
    var values = sh.getDataRange().getValues();
    var headers = values[0];
    Logger.log('  ヘッダー(生): [' + headers.join(', ') + ']');
    var normH = headers.map(function(h) { return normalize_(h); });
    Logger.log('  ヘッダー(normalize): [' + normH.join(', ') + ']');
    var lowerH = normH.map(function(h) { return h.toLowerCase(); });
    Logger.log('  ヘッダー(lowercase): [' + lowerH.join(', ') + ']');

    var emailIdx = normH.indexOf('メール');
    var stampIdx_exact = normH.indexOf('stampfileid');
    var stampIdx_lower = lowerH.indexOf('stampfileid');
    Logger.log('  メール列: idx=' + emailIdx);
    Logger.log('  stampfileid(完全一致): idx=' + stampIdx_exact + ' ← バグ原因！');
    Logger.log('  stampfileid(小文字比較): idx=' + stampIdx_lower + ' ← 修正後');

    // データ行ダンプ
    Logger.log('  データ行数: ' + (values.length - 1));
    for (var r = 1; r < values.length; r++) {
      var email = emailIdx >= 0 ? values[r][emailIdx] : '(列なし)';
      var stampId = stampIdx_lower >= 0 ? values[r][stampIdx_lower] : '(列なし)';
      Logger.log('  行' + (r+1) + ': メール=[' + email + '] stampFileId=[' + stampId + ']');
      if (stampId && String(stampId).trim()) {
        try {
          var f = DriveApp.getFileById(normalize_(stampId));
          Logger.log('    → ファイル存在: ' + f.getName() + ' (' + f.getMimeType() + ')');
        } catch (e) {
          Logger.log('    → ファイル取得エラー: ' + e.message);
        }
      }
    }
  } catch (e) {
    Logger.log('  StampMap読取エラー: ' + e.message);
  }

  // 3. 最新承認済み申請の approvedBy
  Logger.log('\n--- 3. 最新承認済み申請の承認者 ---');
  try {
    var reqSh = requireSheet_('Requests');
    var reqHeaders = reqSh.getRange(1, 1, 1, reqSh.getLastColumn()).getValues()[0];
    var normReqH = reqHeaders.map(function(h) { return normalize_(h); });
    var approvedByIdx = normReqH.indexOf('approvedBy');
    var statusIdx = normReqH.indexOf('status(submitted/approved/canceled)');
    Logger.log('  approvedBy 列: idx=' + approvedByIdx);

    if (approvedByIdx >= 0 && statusIdx >= 0) {
      var lastRow = reqSh.getLastRow();
      if (lastRow >= 2) {
        var data = reqSh.getRange(2, 1, lastRow - 1, reqSh.getLastColumn()).getValues();
        for (var i = data.length - 1; i >= 0; i--) {
          if (normalize_(data[i][statusIdx]) === 'approved') {
            var approver = data[i][approvedByIdx];
            Logger.log('  最新承認者: [' + approver + ']');
            // StampMap で引けるか
            var stampId = lookupStampFileIdByEmail_(String(approver));
            Logger.log('  → lookupStampFileIdByEmail_ 結果: [' + stampId + ']');
            break;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('  Requests読取エラー: ' + e.message);
  }

  Logger.log('\n===== 印鑑診断完了 =====');
}

// ====== Date or JST文字列から "HH:mm" を抽出 ======

function extractHHmm_(val) {
  if (!val) return '';
  if (val instanceof Date) return fmtDate_(val, 'HH:mm');
  // JST文字列 "yyyy-MM-dd HH:mm:ss" や "HH:mm" 形式
  const s = String(val);
  const m = s.match(/(\d{1,2}:\d{2})/);
  return m ? m[1] : '';
}

// ====== フォーム回答からreasonDetailを補完 ======
// Requests に reasonDetail が空で reason が「その他」の場合、
// 元のフォーム回答から補足理由を取得してシートに書き戻す

function tryFillReasonDetailFromForm_(requestId, rowNo, reqSh, idx) {
  // submittedAt からフォーム回答のタイムスタンプを特定
  var submittedAtCol = idx['submittedAt'];
  if (submittedAtCol === undefined) return null;

  var submittedAt = reqSh.getRange(rowNo, submittedAtCol + 1).getValue();
  if (!submittedAt) return null;

  var ts = submittedAt instanceof Date ? submittedAt : new Date(submittedAt);

  // FormMap から全アクティブフォームを取得して探索
  var fmSh = requireSheet_('FormMap');
  var fmValues = fmSh.getDataRange().getValues();
  var fmH = fmValues[0].map(function(h) { return normalize_(h); });
  var fmFormIdCol = fmH.indexOf('formId');
  if (fmFormIdCol < 0) return null;

  // submittedAt の前後2分で検索
  var since = new Date(ts.getTime() - 120000);

  for (var r = 1; r < fmValues.length; r++) {
    var formId = normalize_(fmValues[r][fmFormIdCol]);
    if (!formId) continue;

    try {
      var form = FormApp.openById(formId);
      var responses = form.getResponses(since);

      for (var i = 0; i < responses.length; i++) {
        var resp = responses[i];
        var respTs = resp.getTimestamp();
        // タイムスタンプが近い回答を探す（前後60秒以内）
        if (Math.abs(respTs.getTime() - ts.getTime()) > 60000) continue;

        var irs = resp.getItemResponses();
        for (var j = 0; j < irs.length; j++) {
          try {
            var title = normalize_(irs[j].getItem().getTitle());
            if (title === normalize_('補足理由')) {
              var detail = irs[j].getResponse();
              if (detail && String(detail).trim()) {
                // Requests シートに書き戻し
                var rdCol = idx['reasonDetail'];
                if (rdCol !== undefined) {
                  reqSh.getRange(rowNo, rdCol + 1).setValue(detail);
                  SpreadsheetApp.flush();
                }
                return detail;
              }
            }
          } catch (e) { /* 削除済みアイテムのスキップ */ }
        }
      }
    } catch (e) {
      // フォーム読み取りエラーは無視して次へ
    }
  }
  return null;
}
