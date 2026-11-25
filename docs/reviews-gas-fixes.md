# レビュー指摘の GAS 側の修正結果（team G 2 回目）

- 日付: 2026-11-26
- ブランチ: `team-g-gas`
- 対象: `docs/reviews/review-security.md`・`review-data-safety.md`・`review-correctness.md` のうち、`src/`（GAS）に場所がある指摘
- 試験: `node tools/gas-harness/run.js` → 70/70 passed（前回 49 件。21 件を足し、4 件を新しい動きに合わせて直した）
- 移転モードが OFF のときの動きは、ふつうの入力では今と同じです。変わるのは、先頭が `=`・`+`・`-`・`@` の文字を台帳・PDF のひな形へ書くときだけです（`'` を前に付けて文字として入れる。`'` はセルにも PDF にも出ない）。

## 結果

| 指摘 | 結果 | 内容・理由 |
|---|---|---|
| セキュリティ 高-1（台帳・PDF のセルに式が入る） | 直した | `config.gs` に `sheetLiteral_` を足し、`mail.gs` の `buildAccRow_`（upsert・前日の書き直しで使う）と `rebuildAccumulationSS` の行、`pdfExport.gs` の部署・氏名・工番・得意先・品名・業務内容・理由（直接方式と XLOOKUP 方式の理由）を通す。画面の経路・Sebastian の経路・内部用の入口のどれも同じ関数を通るので全部に効く。VPS の入口の検証は VPS 側の担当 |
| セキュリティ 中-2 ／ データ安全 高-1（どの Drive ファイルでもゴミ箱へ・ステージングが本番の PDF を捨てる） | 直した | `vpsTryTrash_` は、Settings の `PDF_ROOT_FOLDER_ID` の下（親を 3 段までたどる）にある `application/pdf` だけを捨てる。違えば捨てずに `trash.status='skipped'`・`reason`（`outside_pdf_root`／`not_pdf`／`pdf_root_not_set`）・`trashSkipped:true` を返し、返事の最上段にも `trashSkipped:true` を付ける。シートの PDF の 3 列を空にする動きは今のまま。ステージングの DB の `pdf_file_id` を空にする件は VPS 側 |
| セキュリティ 中-2 の後半（`workLog`／`wlOverride` の値） | 直した | 開始・終了は `yyyy-MM-dd HH:mm(:ss)` で終了 >= 開始、分の 3 項目は 0〜1440 の数（空は 0）。合わなければ `bad_request`（再試行しない）。PDF の関数へは 5 項目だけの写しを渡す |
| データ安全 高-3 の連携（ping の `migrated`） | 直した | ping は毎回スクリプトプロパティを読み直して `migrated`・`migrationMode`（いつも true か false）を返す。読めないときは `migrated:false`・`migratedReadError:true`。`light:true` を付けると DB を開かずに `migrated`・`scriptId`・版だけを返す。VPS 側の「本番は migrated=true でなければ書かない」は VPS 側の担当 |
| token の決まり（32 字未満・Sebastian と同じ値を拒む） | 確かめた（もとから） | `dispatchVpsInternal_` が何も読む前に拒む。31 字・32 字の試験を足した |
| Sebastian の返事に例外の文を出さない | 確かめた（今の GAS と同じ） | GAS が自分で答える Sebastian の返事は変えていない。転送の失敗の形は次の行（仕様 4-4 が `upstream unavailable: <message>` と決めているので、その文だけは仕様どおり返す） |
| team C の追加 1（転送の失敗の形を仕様 4-4 に合わせる） | 直した | 例外は `{"error":"upstream unavailable: <message>"}`、200 以外は `{"error":"upstream error: <code>"}`（`action` は付けない）。200 でも JSON でない返事（Cloudflare の画面など）は `upstream error: 200` にして HTML を Sebastian に渡さない（この 1 点は仕様に書かれていない判断） |
| team C の追加 2（転送のヘッダー） | 直した | 仕様 4-4 のとおり `X-Overtime-Forwarded: gas` を付ける |
| team C の追加 3（VPS_INTERNAL_TOKEN・VPS_ORIGIN の設定） | 直した | `Cutover.gs` に `setInternalToken(token?)`（引数なしなら 64 字の乱数を作る。32 字以上・Sebastian と違う値だけ受ける。ログと返り値は長さだけ）と `setVpsOrigin(origin)`（`https://ホスト名` だけ。移転モードには入らない）を足した。どちらも持ち主がエディタから実行したときだけ通る。スクリプト プロパティの画面で直接入れてもよい |
| データ安全 中-3（冪等の記録がプロパティの上限に当たる） | 直した | 記録は 100 件まで、1 件 3,000 字まで（長い結果は要点だけ）。記録に失敗したら返事に `idempotencyStored:false`（成功は `true`）。VPS 側で警告に出すのは VPS 側の担当。本番のプロパティの合計の大きさを測る件は「確かめること」 |
| データ安全 中-4（ledgerRebuild の重なり） | 直した（GAS 側の分） | `internal.ledgerRebuild`・`vps.ledgerRebuild` は、一度始めた jobId を途中で止まっても同じ jobId で動かし直さない（`previous_attempt_unfinished`、再試行しない）。失敗しても「動いている」の印を消さない。再試行しない設定・待ち 0 の確認・docstring は VPS 側 |
| データ安全 中-6（台帳の upsert の失敗を握りつぶす） | 直した | `upsertAccumulationRow_` の後に台帳の年度タブの 1 列目を読み直し、行が無ければ `ledger.status='error'`・`error='ledger_upsert_failed'`・`retryable:true`。`vps.upsertLedger` は再試行できる失敗、PDF の経路は `ok:false`・`error='ledger_failed'`・`retryable:true`（PDF の 3 列は返すので VPS は先に記録できる。送り直すと PDF は already になる） |
| データ安全 低-4（`vps.morningBatch` が残っている） | 直した | 呼ばれても `bad_request` を返し、BatchLogs に行を足さない。`vpsMorningBatch_` は手で使うときのために残した |
| データ安全 高-2（PDF の記録漏れの後に照合が 3 列を空にする） | GAS 側は変更なし | GAS は `skipped/has_pdf` のときも最上段に `pdfFileId` を返している（今のまま）。直すのは VPS の `apply_response` と照合 |
| データ安全 中-1・中-2・中-5・中-7、低-1〜3・5〜9 | 対象外 | 場所が VPS（`vps-app/`）だけ |
| セキュリティ 中-1・中-3・中-4・中-5、低-1〜7・10〜12 | 対象外 | 場所が VPS（`vps-app/`）だけ |
| セキュリティ 低-8（内部用の入口のリプレイ） | 直していない | 低。VPS が `ts` を送る変更と一緒に入れる必要がある |
| セキュリティ 低-9（`api.gs` の token の比較が定数時間でない・転送の前に token を確かめない） | 直していない | 低。`api.gs` の比較は今の本番の動きなので変えない。転送の前の token の確認は「本文をそのまま転送する」決まり（試験あり）を変えるため、決めてから入れる |
| 正しさのレビュー（中-1、低-1〜19） | 対象外 | GAS のコードを直す指摘は無い（どれも Python 側か仕様書の直し） |

## 気を付けること

- 作り直し（`internal.regenerate`）で台帳だけが失敗して送り直したとき、VPS が `oldPdfFileId` を送らないと、1 回目に作った新しい PDF をゴミ箱へ移して作り直します。結果のシートと返事は正しいままです。VPS は作り直しのときにいつも `oldPdfFileId` を送っています。
- `PDF_ROOT_FOLDER_ID` が空の環境では、内部用の入口は PDF を 1 つもゴミ箱へ移しません（`pdf_root_not_set`）。

## 確かめること

- 本番の GAS のスクリプトプロパティの今の合計の大きさ（データ安全 中-3）
- 本番の PDF が `PDF_ROOT_FOLDER_ID` の直下の日付フォルダにあること（3 段より深い所に置いた PDF は捨てない）
