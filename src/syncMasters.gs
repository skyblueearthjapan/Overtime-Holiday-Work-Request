// ====== マスタ自動転記（元スプレッドシート → DB） ======

/**
 * 転記対象マスタ定義
 * src: 元スプレッドシートのシート名
 * dst: 転記先（DB）のシート名
 *
 * ※ 元シート名が実際と異なる場合は src を修正してください。
 *   転記先は config.gs の SHEET 定数を参照。
 */
const SYNC_MAP = [
  { src: '部署マスタ',     dst: SHEET.DEPTS   },   // 部署マスタ → 部署マスタ
  { src: '作業員マスタ',   dst: SHEET.WORKERS },   // 作業員マスタ → 作業員マスタ
  { src: '業務NO.マスタ',  dst: SHEET.JOBS    },   // 業務NO.マスタ → 業務NOマスタ
  { src: '工番マスタ',     dst: SHEET.ORDERS  },   // 工番マスタ → 工番マスタ
];

/**
 * 全マスタを元スプレッドシート → DB へ同期（全置換）。
 *
 * 前提：Settings シートに以下を設定
 *   SOURCE_SSID = 元スプレッドシート（作業日報_全従業員用）の ID
 *
 * 手動実行 or トリガーから呼び出す。
 */
function syncAllMasters() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const settings = getSettings_();
    const sourceId = normalize_(settings['SOURCE_SSID']);
    if (!sourceId) {
      throw new Error(
        'Settings に SOURCE_SSID が未設定です。\n' +
        '元スプレッドシート（作業日報_全従業員用）の ID を Settings シートに追加してください。\n' +
        'URL: https://docs.google.com/spreadsheets/d/【このID部分】/edit'
      );
    }

    const srcSS = SpreadsheetApp.openById(sourceId);
    const dstSS = getDb_();
    const log = [];

    for (const m of SYNC_MAP) {
      try {
        const srcSh = srcSS.getSheetByName(m.src);
        if (!srcSh) { log.push('SKIP ' + m.src + ': 元シートなし'); continue; }

        const dstSh = dstSS.getSheetByName(m.dst);
        if (!dstSh) { log.push('SKIP ' + m.dst + ': 転記先シートなし'); continue; }

        var data = readSheetData_(srcSh);
        if (!data || !data.length) { log.push('SKIP ' + m.src + ': 空'); continue; }

        replaceMasterData_(dstSh, data);
        SpreadsheetApp.flush();
        log.push('OK ' + m.src + ' -> ' + m.dst + ' (' + data.length + ' rows, ' + data[0].length + ' cols)');
      } catch (e) {
        log.push('ERROR ' + m.src + ': ' + e.message);
      }
    }

    const summary = 'syncAllMasters:\n' + log.join('\n');
    Logger.log(summary);
    console.log(summary);
  } finally {
    lock.releaseLock();
  }
}

/**
 * シートデータを読み取る。
 * Service error 対策：リトライ → バッチ読み取りフォールバック。
 */
function readSheetData_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return null;

  // 1st: 通常読み取り
  try {
    return sh.getRange(1, 1, lastRow, lastCol).getValues();
  } catch (e) {
    console.warn('readSheetData_ 1st try failed (rows=' + lastRow +
      ', cols=' + lastCol + '): ' + e.message);
  }

  // 2nd: 2秒待ってリトライ
  Utilities.sleep(2000);
  try {
    return sh.getRange(1, 1, lastRow, lastCol).getValues();
  } catch (e) {
    console.warn('readSheetData_ 2nd try failed: ' + e.message);
  }

  // 3rd: バッチ読み取り（50行ずつ）
  Utilities.sleep(2000);
  try {
    var BATCH = 50;
    var all = [];
    for (var r = 1; r <= lastRow; r += BATCH) {
      var n = Math.min(BATCH, lastRow - r + 1);
      var batch = sh.getRange(r, 1, n, lastCol).getValues();
      for (var i = 0; i < batch.length; i++) all.push(batch[i]);
    }
    return all;
  } catch (e) {
    throw new Error(
      'Read failed after all retries (rows=' + lastRow +
      ', cols=' + lastCol + '): ' + e.message
    );
  }
}

/**
 * シートを全置換（値のみ）。
 * - clearContents で既存データをクリア
 * - 行列数を元データに合わせて拡張/縮小
 * - 値を一括貼り付け
 */
function replaceMasterData_(sh, values) {
  var rows = values.length;
  var cols = values[0].length;

  sh.clearContents();

  // 行列を確保（不足分を追加）
  if (sh.getMaxRows() < rows) {
    sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  }
  if (sh.getMaxColumns() < cols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  }

  // 余剰行列を削除（シート肥大化防止、最低1行は残す）
  var excessRows = sh.getMaxRows() - rows;
  if (excessRows > 5) {
    sh.deleteRows(rows + 1, excessRows);
  }
  var excessCols = sh.getMaxColumns() - cols;
  if (excessCols > 2) {
    sh.deleteColumns(cols + 1, excessCols);
  }

  // 値貼り付け
  sh.getRange(1, 1, rows, cols).setValues(values);
}
