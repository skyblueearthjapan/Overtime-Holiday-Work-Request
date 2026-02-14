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
  { src: '工番',           dst: SHEET.ORDERS  },   // 工番 → 工番マスタ
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

        const data = srcSh.getDataRange().getValues();
        if (!data.length) { log.push('SKIP ' + m.src + ': 空'); continue; }

        replaceMasterData_(dstSh, data);
        log.push('OK ' + m.src + ' -> ' + m.dst + ' (' + data.length + ' rows)');
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
