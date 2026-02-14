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
  { src: '部署マスタ',     dst: SHEET.DEPTS   },
  { src: '作業員マスタ',   dst: SHEET.WORKERS, maxRows: 81 },  // ヘッダ+80人
  { src: '業務NO.マスタ',  dst: SHEET.JOBS    },
  { src: '工番マスタ',     dst: SHEET.ORDERS  },
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

        var dstSh = dstSS.getSheetByName(m.dst);
        if (!dstSh) { log.push('SKIP ' + m.dst + ': 転記先シートなし'); continue; }

        var data = readSheetData_(srcSh);
        if (!data || !data.length) { log.push('SKIP ' + m.src + ': 空'); continue; }

        // 空行除外（1列目が空の行を除去、ヘッダは維持）
        data = compactRows_(data);

        // maxRows 制限（ヘッダ含む行数）
        if (m.maxRows && data.length > m.maxRows) {
          data = data.slice(0, m.maxRows);
        }

        // 通常の置換を試行、失敗時はシート再作成で回避
        try {
          replaceMasterData_(dstSh, data);
        } catch (writeErr) {
          console.warn('replaceMasterData_ failed for ' + m.dst +
            ', recreating sheet: ' + writeErr.message);
          dstSS.deleteSheet(dstSh);
          dstSh = dstSS.insertSheet(m.dst);
          dstSh.getRange(1, 1, data.length, data[0].length).setValues(data);
        }

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

/**
 * 空行を除外（1列目が空の行を除去、ヘッダ行は維持）。
 */
function compactRows_(values) {
  var header = values[0];
  var body = [];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() !== '') {
      body.push(values[i]);
    }
  }
  return [header].concat(body);
}

// ====== 診断用（切り分けテスト） ======

/**
 * 作業員マスタの1行だけ転記テスト。
 * - 成功 → データ量/シート状態が原因
 * - 失敗 → 保護/権限が原因
 */
function testSyncWorkerOneRow() {
  var settings = getSettings_();
  var sourceId = normalize_(settings['SOURCE_SSID']);
  if (!sourceId) throw new Error('SOURCE_SSID 未設定');

  var srcSS = SpreadsheetApp.openById(sourceId);
  var dstSS = getDb_();

  var srcSh = srcSS.getSheetByName('作業員マスタ');
  if (!srcSh) throw new Error('元シート「作業員マスタ」が見つかりません');

  var dstSh = dstSS.getSheetByName(SHEET.WORKERS);
  if (!dstSh) throw new Error('転記先シート「' + SHEET.WORKERS + '」が見つかりません');

  // 元からヘッダ＋1行だけ読む
  var lastCol = srcSh.getLastColumn();
  Logger.log('元シート: lastRow=' + srcSh.getLastRow() + ', lastCol=' + lastCol);

  var v = srcSh.getRange(1, 1, 2, lastCol).getValues();
  Logger.log('読み取り成功: ' + JSON.stringify(v));

  // 転記先に書く
  dstSh.clearContents();
  dstSh.getRange(1, 1, v.length, v[0].length).setValues(v);
  Logger.log('書き込み成功: 1行テスト OK');
}
