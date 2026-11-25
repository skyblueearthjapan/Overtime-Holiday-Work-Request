# 移行計画の穴さがし 3: メール・PDF・蓄積台帳

対象の計画: `docs/vps-migration-plan.md`（以下「計画」）。読んだコード（全行）: `src/pdfExport.gs`・`src/mail.gs`。PDF と台帳を呼ぶ部分: `src/approval.gs:278-476, 692-883`・`src/worklog.gs:1-130, 141-342, 535-700`・`src/api.gs`（全行）・`src/adminApi.gs:428-518`・`src/menu.gs:1-43`・`src/config.gs:60-100`。2026-11-26 時点の作業ツリー。Google には触れていません。

「確認済み」はコードで確かめたこと、「推測」はコードだけでは決まらないことです。

---

## 0. 先に結論（高の指摘）

| ID | 重さ | 一言 |
|---|---|---|
| G3-H1 | 高 | 朝バッチは「PDF を全部作り終えてから」BatchLogs と台帳を書く。PDF を outbox に積むだけだと、BatchLogs の数と台帳の PDF リンクが今と変わる |
| G3-H2 | 高 | 台帳の upsert は「PDF の後」に、経路ごとに違う条件で走る。PDF と台帳を別の仕事にすると、PDF 列の中身と「書くか書かないか」が今と変わる |
| G3-H3 | 高 | 案 1 の「違いは Requests を読まない・書かないの 2 つだけ」は誤り。PDF のコードは Settings・StampMap・`getDb_()`（`SpreadsheetApp.getActive()`）も読むので、別プロジェクトに写すとそのままでは動かない |
| G3-H4 | 高 | 直接方式は「値が空なら書かない」セルがあり、テンプレの数式（`操作!B3` を引く XLOOKUP の見込み）がそのまま残る。今でも別の申請の値が出ている恐れがある。切り替え後は Requests が毎晩の写しになり、さらにずれる |

中・低の指摘は 9 章、計画の記述で正しいと確かめたものは 10 章にあります。

---

## 1. PDF を作る経路（本番 `PDF_MODE`=direct の前提）

| きっかけ | 関数（場所） | 方式 | 実績の渡し方 | ロック | PDF の後の台帳 | PDF が失敗したとき |
|---|---|---|---|---|---|---|
| 一次承認（個別） | `api_approveRequest` `approval.gs:328-341` | `PDF_MODE`→direct | WorkLogs を読み直す | 承認のロックの**外**（PDF 自身のロックは取る） | 同じ `try` の中で PDF の**後**。PDF が例外なら**台帳も書かない** | `console.warn` のみ。承認は成功 |
| 一次承認（一括） | `api_approveRequestsBatch` `approval.gs:454-473` | 同上 | 同上 | 同上 | 1 件ごとに PDF の後。PDF が例外ならその件の台帳は書かない | 同上 |
| 残業の完了（画面） | `api_markOvertimeDone` `worklog.gs:193-207` | direct | `wlOverride`（計算済みの値） | 完了のロックの**中**で入れ子 | PDF の成否に関係なく upsert | `console.warn`。完了は成功 |
| 休日の完了（画面） | `api_markHolidayDone` `worklog.gs:322-336` | direct | `wlOverride` | 入れ子 | 同上 | 同上 |
| 総務の手入力 | `api_adminManualComplete` `adminApi.gs:498-512` | direct | `wlOverride` | 入れ子 | 同上 | `Logger.log` |
| Sebastian 残業完了 | `sebastian_markOvertimeDone_` `api.gs:102-113` | **いつも XLOOKUP**（`generatePdfForRequest_`） | WorkLogs を読み直す（override なし） | 入れ子 | 同上 | 返す JSON の `pdf` に `{error}` |
| Sebastian 休日完了 | `sebastian_markHolidayDone_` `api.gs:229-239` | いつも XLOOKUP | 同上 | 入れ子 | 同上 | 同上 |
| 二次承認（個別） | `api_secondaryApprove` `approval.gs:735-760` | **いつも direct** | 読み直す | ロック内で旧 PDF をゴミ箱＋3 列を空に → ロック外で作り直し | PDF の成否に関係なく upsert | 旧 PDF は消えたまま、3 列は空のまま |
| 二次承認（一括） | `api_secondaryApproveBatch` `approval.gs:833-865` | いつも direct | 同上 | 同上。6 件以上は 1 件ごと 500ms 待ち | 二次承認した全件 upsert | 同上 |
| 申請の編集 | `api_editRequest` `worklog.gs:610-634` | `PDF_MODE`→direct | 読み直す | 入れ子（ロック内で旧 PDF をゴミ箱＋3 列を空に） | PDF が**あった**ときだけ upsert | 消えたまま |
| 申請の削除 | `api_deleteRequest` `worklog.gs:668-694` | 作らない | — | — | `removeAccumulationRow_` | — |
| 朝バッチ | `morningBatchNoMail_` `mail.gs:1454-1468` → `batchGeneratePdfs_` `pdfExport.gs:570-600` | `PDF_MODE`→direct | 読み直す | 1 件ごとに PDF のロック | PDF を全部作った**後**に前日分を書き直し | `BatchLogs` の失敗数とエラー文 |
| 手動メニュー「PDF一括生成（本日分）」 | `manualBatchPdfs_` `menu.gs:33-43` | 同上（**今日**の分） | 同上 | 同上 | 台帳は書かない | `BatchLogs` に `manualBatchPdfs` |
| 手動メニュー「朝バッチ」 | `morningBatch_` `mail.gs:1427-1448` | 同上 | 同上 | 同上 | 前日分を書き直し | `BatchLogs` に `morningBatch`（メールはコメントアウト済み。コミット `fcf8e31` 2026-03-11） |
| 手動「PDF 強制再生成」 | `forceRegeneratePdf` `pdfExport.gs:743-871` | `PDF_MODE` | 読み直す | — | 書かない | ログのみ。**旧 PDF はゴミ箱へ送らない**（3 列を空にするだけ）。WorkLogs の残業開始を 17:20 に書き換え、`reasonDetail` 列が無ければ足し、フォーム回答から補足理由を補う |

- 作る条件は「`status==='approved'`」と「`pdfGeneratedAt` と `pdfFileId` の両方が無い」（`pdfExport.gs:167-170, 370-377`）。朝バッチの「作成済み」判定は `pdfFileId` だけ（`pdfExport.gs:580`）。確認済み。
- 一次承認で PDF を作るのは「WorkLogs に `actualEndAt` がある」ときだけ（`approval.gs:319-321, 441-448`）。確認済み。
- 二次承認は**既に PDF があるときだけ**作り直す（`approval.gs:721, 736, 747`）。完了済みでも PDF が無ければ作らない。確認済み。
- ロックの入れ子: 完了系は自分のロックを持ったまま `generatePdf*_` が同じスクリプトロックを `waitLock(20000)` で取り、最後に `releaseLock()` する。GAS のロックが同じ実行の中で取り直せるなら、PDF の後の台帳 upsert はロックが外れた状態で走る（推測。今も PDF が作られているので「取り直せる」側の見込み）。

## 2. テンプレに書くセル（`申請書フォーム`、`PDF_MAP` `pdfExport.gs:3-23`）

直接方式 `fillPdfTemplate_`（`pdfExport.gs:440-566`）:

| セル | 書く値 | いつも書くか |
|---|---|---|
| G1 | **PDF を作った時刻**の `yyyy/MM/dd`（`new Date()`） | いつも |
| B4 | `dept`＋中央揃え | いつも |
| D4 | `workerName` | いつも |
| F4 | 区分印が読めたら `''`。印の ID はあるが画像が読めなければ `残業/休日出勤`（書式なし）。印の ID が無ければ文字＋24pt 太字・中央・上下中央 | いつも |
| C6 | 残業 `残業`、休日 `approvedMinutes<=240` で `半日`、他 `1日` | いつも |
| C7 | `targetDate` の `yyyy/MM/dd` | いつも |
| C10 | 残業は `17:20` 固定。休日は `actualStartAt` の `HH:mm` | 休日は**値があるときだけ** |
| F10 | `actualEndAt` の `HH:mm` | **値があるときだけ** |
| C12 | `fmtMinToHM_(breakMinutes)`（空でも `0分`） | いつも |
| F12 | `sanitizeNetMinutes_` 後を `fmtMinToHM_`（残業 720 分・休日 960 分超は開始〜終了−休憩で計算し直す） | いつも |
| B18/D18/F18〜B20/D20/F20 | `workNoN`（空なら `orderNoN`）・`customerN`・`productN` | **値があるときだけ** |
| A23 | `workContent` | **値があるときだけ** |
| A29 | `buildReasonText_`（「（…補足理由…）」を取り、補足理由を改行でつなぐ） | いつも（注釈「空でも setValue して、テンプレの XLOOKUP 数式をクリアする」`pdfExport.gs:525`） |
| F34 / G34 | 文字は書かない。画像だけ | — |

XLOOKUP 方式 `generatePdfForRequest_`（Sebastian だけ）: コピーの `操作!B3` に requestId → flush → 600ms 待ち → flush。直接書くのは A29・C10・F10・C12・F12・C6（休日だけ）と画像。F4 は印が読めたら `''`、印の ID が無ければ書式だけ（文字は数式まかせ）。G1・B4・D4・C6（残業）・C7・明細・A23 は**テンプレの数式まかせ**。G34（二次承認印）は入れない（`pdfExport.gs:187-273`）。確認済み。

値の書き方はどちらも `setValue`（人が入力したのと同じ扱い）。`17:20` は時刻、`2026/11/25` は日付、工番の `12-3` のような文字は日付に化けうる。表示はテンプレのセルの表示形式しだい。確認済み（書き方）／推測（化け方はテンプレの書式による）。

## 3. 印鑑

| 印 | 画像 | 置き方（`Sheet.insertImage`＝セルの上に重ねる画像） |
|---|---|---|
| 区分印 | Settings `STAMP_OVERTIME_FILE_ID`／`STAMP_HOLIDAY_FILE_ID` | `insertImage(blob, 6, 1, 0, 0)`＝F1 起点、270×270px（`pdfExport.gs:467-468`） |
| 一次承認印 | `approvedBy` を小文字にして `StampMap`（見出し `メール` と `stampfileid` を小文字で比較）で引く | `insertImage(blob, 6, 34, x, y)`。x＝(F 列幅−60)/2、y＝(34〜37 行の高さの合計−60)/2、60×60px（`pdfExport.gs:535-543`） |
| 二次承認印 | `approvedBy2` で同じく | G 列で同じ計算（`pdfExport.gs:555-562`）。直接方式だけ |

- 画像は 2MB 未満ならそのまま、以上なら `drive.google.com/thumbnail?id=…&sz=400` を Bearer 付きで取る（`pdfExport.gs:73-97`）。
- 直接方式では `insertImage` が `try` で囲まれていない（`pdfExport.gs:461-565`）。画像の差し込みで例外が出ると **PDF 全体が失敗し、一時コピーが残る**。XLOOKUP 方式は `try` で囲み、印なしで続ける（`pdfExport.gs:239-245, 260-270`）。確認済み。

## 4. 書き出し・名前・置き場所・後片付け

- URL: `https://docs.google.com/spreadsheets/d/{コピーID}/export?format=pdf&gid={申請書フォームのgid}&portrait=true&size=A4&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`、`Authorization: Bearer ScriptApp.getOAuthToken()`。200 以外は例外（`pdfExport.gs:312-338`）。余白・中央寄せの指定は無い。
- 一時コピー: `DriveApp.getFileById(TEMPLATE_SSID).makeCopy('TMP_{requestId}_{yyyyMMdd_HHmmss}')`。テンプレ＝DB なので DB 全体（Requests・WorkLogs・マスタ類）の写し。置き場所は `makeCopy` の既定（推測: DB と同じフォルダ）。DB に紐づく GAS も一緒に写る（推測: コンテナに紐づくスクリプトの性質）。
- 名前: `{targetDate yyyyMMdd}_{dept}_{workerName}_{残業|休日出勤}.pdf`。`\ / : * ? " < > |` を `_` に。同じ人・同じ日・同じ種別の申請が 2 件あれば**同名の PDF が同じフォルダに 2 つ**できる（Drive は同名を許す）。確認済み。
- フォルダ: `PDF_ROOT_FOLDER_ID` 直下の `targetDate` の `yyyy.MM.dd`。`getFoldersByName` の最初の 1 つ、無ければ `createFolder`（`pdfExport.gs:150-158`）。今はスクリプトロックの中なので同時に 2 つ作られない。確認済み。
- 記録: Requests の `pdfGeneratedAt`（Date）・`pdfFileId`・`pdfFolderId`。
- 後片付け: 成功したときだけ一時コピーをゴミ箱へ（`pdfExport.gs:302, 430`）。`try/finally` はロックの解放だけ。PDF を保存した後で Requests への記録が失敗すると、**PDF は残り Requests は空**のまま（再実行で同名の 2 つ目ができる）。確認済み。
- 旧 PDF: 二次承認・編集・削除で `setTrashed(true)`（`approval.gs:873-883`）。`forceRegeneratePdf` は旧 PDF を消さない。確認済み。

## 5. Requests を読む数式（切り替え後に壊れうるもの）

コードから分かるのは次のとおりです。テンプレの中身は読んでいないので、一覧は候補です。

| セル | 根拠 | 確かさ |
|---|---|---|
| A29 | `pdfExport.gs:525` の注釈「テンプレの XLOOKUP 数式をクリアする」、`debugPdfReason` が A29 の数式を調べる（`pdfExport.gs:704-731`） | 確認済み（注釈） |
| `操作!B3` を起点に Requests を引く数式群 | XLOOKUP 方式が B3 に requestId を入れて「XLOOKUP 反映待ち」をする（`pdfExport.gs:187-195`） | 確認済み（仕組み） |
| G1・B4・D4・F4・C6・C7・B18〜F20・A23 | XLOOKUP 方式で直接書かないのに値が出ている | 推測 |
| C10・F10・C12・F12 | XLOOKUP 方式が「TZ バグ回避」で上書きしている＝元は数式 | 推測 |
| F34（承認者のメール） | コミット `01646a3`「メール非表示」と、直接方式の注釈「メールアドレスは書き込まない」 | 推測（今は数式を消した可能性もある） |

直接方式でも書かないセル（F10・休日の C10・明細・A23・F34・G34 の文字）に数式が残っていれば、その値はコピーの `操作!B3` が指す申請から来ます。直接方式は B3 を書かないので、**DB の `操作!B3` に入っている固定の requestId の申請**が引かれます（G3-H4）。

## 6. 蓄積台帳（`ACCUMULATION_SS_ID`）

- シート: `残業_{年度}年度`／`休日出勤_{年度}年度`。年度は 3/16 始まり（`mail.gs:465-472`）。新しいシートは先頭に作り、`Sheet1`／`シート1` を消す。見出し 27 列・列幅・見出しの書式は `mail.gs:343-382, 433-450`（棚卸し 03 §3.2 の表は正しい）。
- 書く値はすべて `setValues`（入力と同じ扱い）。`yyyy/MM/dd` は日付、`HH:mm` は時刻になる。理由・業務内容の自由入力も同じ扱いなので、`1/2` は日付、`=` で始まる文は数式になる。PDF 列だけ後から `setFormula('=HYPERLINK(…,"PDF")')`。確認済み。
- 行の書式: 書いた行だけ 10pt・上下中央、行番号が偶数なら `#f8fafc`、奇数なら `#ffffff`（`mail.gs:775-795`）。行を消した後も他の行は塗り直さないので、縞は少しずつずれる。確認済み。

書き手:

| 関数 | 呼ばれる所 | 動き | 埋まる列の癖 |
|---|---|---|---|
| `upsertAccumulationRow_` `mail.gs:651-710` | 1 章の表のとおり | 年度シートの 1 列目で requestId を探し、あれば上書き、無ければ末尾。例外は飲み込む | 1次承認日が空（`getRequestById_` に `approvedAt` が無い）。状態は今の status（完了が承認より先なら `申請中`） |
| `removeAccumulationRow_` `mail.gs:716-741` | 削除だけ（`worklog.gs:691`） | 下から探して最初の 1 行を消す | 却下（`api_rejectRequest`）では呼ばれない → 完了後に却下された申請の `申請中` 行は、その日の朝バッチが走るまで（走った後なら永久に）残る |
| `appendToAccumulationSS_(前日)` `mail.gs:489-514` | 朝バッチ（PDF の後） | 3 列目が前日の行を**全部**消し、前日の approved を末尾に足す | 理由・補足・業務内容・1次承認者・2次承認日・2次承認者・状態が空。1次承認日は入る。作業員 ID は Requests を 1 件ごとに全行読み直す |
| `rebuildAccumulationSS` `mail.gs:943-1127` | メニュー（手動） | 年度シートを `_backup_{時刻}` に名前替えし、canceled 以外の全件で作り直す | 全列が埋まる。`submitted` も入る |
| `migrateAccumulationTimeColumns` `mail.gs:806-927` | メニュー（1 回限り） | 旧 23 列に表記列 4 つを差し込む | 移行不要の見込み |
| `getAccumulationSS_` `mail.gs:390-419` | 上の全部 | 開けなければ新しく作り、直書きのフォルダへ移し、Settings の末尾に `ACCUMULATION_SS_ID` を足す | `getSettings_` は同じキーが 2 行あれば**下の行**が勝つ（`config.gs:83-87`） |

- `getRequestFullData_` は 2 つある（`mail.gs:746`・`pdfExport.gs:342`）。直接方式は `status(...)` などの全列を要るので、本番で直接方式の PDF ができている以上、`pdfExport.gs` 版が勝っている（推測。根拠は結果から）。移行先では 1 つにする。

## 7. メール

- 送る所は `sendEveningMail_`（`mail.gs:326-336`）と `sendMorningMail_`（`mail.gs:1316-1420`）の 2 つだけ。どちらも `GmailApp.sendEmail(HR_MAIL_TO, 件名, 本文)`。トリガーからは呼ばれず、メニューからだけ届く。`morningBatch_` の送信行はコミット `fcf8e31`（2026-03-11）でコメントアウト。確認済み。
- Settings の `HR_MAIL_TO_EVENING_EXTRA` はコードのどこにも出てこない（grep で 0 件）。統合メール通知システムが Settings を読んでいる手がかり（推測）。
- 朝メールの xlsx 添付（`exportRowsToXlsxBlob_`）と CSV（`makeCsvBlob_`）は、今どこからも呼ばれていない（grep で呼び元なし）。確認済み。

## 8. 案 1（PDF だけ GAS に残す）で今を再現できるか

| 項目 | 再現できるか | 条件 |
|---|---|---|
| 描画・印の置き方・書き出し URL | できる | コードをそのまま写す |
| 二次承認印 G34 | できる | VPS が `approvedBy2` を渡す。StampMap の引き方（小文字）を同じに |
| Settings・StampMap・TZ | **このままでは動かない** | `getDb_()` が `SpreadsheetApp.getActive()`（`config.gs:69-72`）。スタンドアロンでは null。印の ID を VPS が渡すか、DB を `openById` で読む形に直す（G3-H3） |
| XLOOKUP 方式（Sebastian） | このままでは別物になる | コピーの Requests は毎晩の写し。続けるなら、コピーの Requests にその申請の行を書いてから B3 を入れる。直接方式にそろえるなら Q3 |
| ロック | 注意が要る | 写したコードの `waitLock(20000)` は新しいプロジェクトのロック。worker が並べて呼ぶと 20 秒待ちで失敗しうる（G3-M1） |
| 失敗時の後片付け | 直す（計画どおり） | `finally` で一時コピーをゴミ箱へ。`insertImage` の例外も同じく |
| G1（作成日） | できる | 実行した日になる。worker の遅れ・再試行で日が変わりうる（G3-L4） |
| 朝バッチの数と台帳 | 呼び方しだい | worker が 1 件ずつ待って結果を数える（G3-H1） |

案 1 そのものは成り立ちます。ただし「違いは 2 つだけ」ではなく、入力の契約（下の 9 章 G3-H3）と同時実行の決め（G3-M1）が要ります。

---

## 9. 指摘の一覧

### G3-H1（高）朝バッチの順序と BatchLogs の数

- 根拠（確認済み）: `morningBatchNoMail_` は `batchGeneratePdfs_` で前日の PDF を**同期で 1 件ずつ**作り、`{ok, skip, fail, errors}` を数え（`pdfExport.gs:570-600`）、`logBatchResult_` で BatchLogs に 1 行（`pdfExport.gs:604-627`）、その後で `appendToAccumulationSS_`（`mail.gs:1458-1467`）。台帳の行は PDF の後に読むので、朝に作った PDF のリンクが入る。
- 計画: 4-6「PDF は worker の outbox で作る」、4-9「`morning_batch`: 前日の approved で PDF なし → PDF、BatchLogs に 1 行、台帳の前日分を書き直し」。PDF を outbox に積むのか、その場で作るのかは書いていない。
- なぜ困るか: 積むだけなら、BatchLogs の「成功・失敗」が分からず、台帳の前日分は PDF 列が空のまま書かれる。朝バッチの後に台帳を書き直す経路は無いので、総務の台帳から朝バッチ分の PDF リンクが消える。
- 直し方: 4-9 に「`morning_batch` は PDF を outbox に積まず、同じ仕事の中で 1 件ずつ作り終えて数え、BatchLogs（`実行日時, morningBatchNoMail, 対象日 yyyy-MM-dd, 成功, スキップ, 失敗, エラー詳細`）を書き、最後に台帳の前日分を書き直す」と書く。エラー詳細の形 `{requestId} ({dept}/{workerName}): {メッセージ}` も同じに。6 件以上のとき 1 件ごとに 500ms 待つ（`pdfExport.gs:596`）も写す。

### G3-H2（高）台帳 upsert の順序と条件が経路ごとに違う

- 根拠（確認済み）: 1 章の表。一次承認は PDF と同じ `try` の中で PDF の後に upsert（`approval.gs:330-340, 459-468`）→ PDF が例外なら台帳を書かない。完了系・総務の手入力・Sebastian は PDF の成否に関係なく upsert。二次承認は PDF の後に全件 upsert。編集は PDF があったときだけ。どの経路も upsert は PDF の**後**に `pdfFileId` を読み直して PDF 列に入れる（`mail.gs:673-675`）。
- 計画: 4-7「今の動き（upsert・削除・…）を同じ結果で再現」。PDF と台帳を同じ仕事にするか、順番をどう守るかは書いていない。棚卸し 03 §3.3 も「一次承認で PDF を作ったとき」とだけ書き、「PDF が失敗したら書かない」は抜けている。
- なぜ困るか: 台帳の仕事が PDF より先に走ると PDF 列が空になる。PDF が失敗したときの「書く／書かない」も変わる。
- 直し方: 4-7 に経路ごとの表（きっかけ・PDF の後か・PDF 失敗時に書くか）を載せ、「台帳の upsert は PDF の仕事の最後の手順として同じ仕事で行う（PDF が無い経路は単独の仕事）」と決める。PDF を再試行して後から成功したときに upsert をもう 1 回走らせるか（今は走らない）も Q5 に足す。

### G3-H3（高）案 1 の「違いは 2 つだけ」は誤り

- 根拠（確認済み）: `fillPdfTemplate_` は `getSettings_()`（`pdfExport.gs:450, 459-460`）、`lookupStampFileIdByEmail_`（`requireSheet_(SHEET.STAMP_MAP)`、`pdfExport.gs:103-127`）を呼ぶ。どちらも `getDb_()`＝`SpreadsheetApp.getActive()`（`config.gs:69-80, 92-97`）。`extractHHmm_` も Date のとき `getSafeTimeZone_(null)`→`getDb_()`（`pdfExport.gs:978`）。`generatePdfDirect_` は `PDF_ROOT_FOLDER_ID`・`TEMPLATE_SSID` も Settings から読む（`pdfExport.gs:379-381`）。
- 計画: 4-6 案 1「違いは『Requests を読まず、VPS から渡された値で書く』ことと『Requests に書かず ID を返す』ことだけ」。
- なぜ困るか: 新しい別のプロジェクト（スタンドアロン）では `getActive()` が null になり、Settings を読む所で落ちる。直し方しだいで印の出方が変わる。
- 直し方: 案 1 の「呼び口の入力」を 1 つの表で決める。要る値は `requestId`・`requestType`・`targetDate`（`yyyy-MM-dd`）・`dept`・`workerName`・`approvedMinutes`・`workNo1〜3`（無ければ `orderNo1〜3`）・`customer1〜3`・`product1〜3`・`workContent`・`reason`・`reasonDetail`・`approvedBy`・`approvedBy2`・`actualStartAt`・`actualEndAt`（JST 文字列）・`breakMinutes`・`netMinutes`・区分印の ID・一次と二次の印の ID（または StampMap を DB から `openById` で読む）・`PDF_ROOT_FOLDER_ID`・`TEMPLATE_SSID`・作成日（G1）。返す値は `pdfFileId`・`folderId`・`pdfName`・生成時刻。新しいプロジェクトの `timeZone` は `Asia/Tokyo`。

### G3-H4（高）テンプレの数式が書かないセルに残る

- 根拠: 直接方式は F10・休日の C10・明細 9 セル・A23 を「値があるときだけ」書き、F34・G34 の文字は触らない（2 章、確認済み）。テンプレに XLOOKUP 数式があることは A29 の注釈で確認済み（`pdfExport.gs:525`）。他のセルの数式は推測（5 章）。直接方式はコピーの `操作!B3` を書かないので、数式は DB の `操作!B3` に入っている requestId を引く（推測。B3 の今の値は未確認）。
- 計画: 4-6「難しい点」3 つ目と 8 章のリスク表で「数式が残っていると、切り替えの後は値がずれる（推測）」とだけ書く。**今の本番でも**別の申請の値が出ている恐れと、B3 の値に触れていない。
- なぜ困るか: 工番が 1 件だけの申請の 2・3 行目、業務内容が空の申請、朝バッチで作る未完了の申請の F10 に、別の人の工番・得意先・時刻が出うる。同じに移すと、この誤りも写す。切り替え後は Requests が毎晩の写しなので、出る値がさらに変わる。
- 直し方: スパイクの手順 1 に「`申請書フォーム` の全セルの数式（`getFormulas`）と `操作` タブの全数式、DB の `操作!B3` の今の値を控える」を足す。本番の PDF（明細 1 件・業務内容なし・未完了の朝バッチ分）を 3 枚開き、空欄が本当に空かを見る。結果で「書かないセルに空文字を書いて数式を消す（見た目の変更としてオーナーに聞く）」か「今のまま」かを Q に足す。

### G3-M1（中）案 1 の同時実行とフォルダの重複

- 根拠（確認済み）: `generatePdfDirect_` は `waitLock(20000)`（`pdfExport.gs:385-386`）。日付フォルダは「名前で探す→無ければ作る」（`pdfExport.gs:150-158`）。今は全部が 1 本のスクリプトロックで直列。
- 計画: 4-6「申請 1 件の仕事は同時に 1 本、積んだ順」。別の申請どうしの同時実行は書いていない。
- なぜ困るか: worker が別の申請の PDF を並べて呼ぶと、GAS 側のロック待ちが 20 秒を超えて失敗しうる（1 件で DB 全体の複製と書き出し）。ロックを外せば、同じ日付のフォルダが 2 つできうる。
- 直し方: 「PDF の仕事は全体で同時に 1 本」と書く（今のスクリプトロックと同じ）。GAS 側のロックは残す。スパイクで 1 件の所要時間を測る。

### G3-M2（中）失敗したときの戻り方（25 件の原因の候補）

- 根拠（確認済み）: 二次承認と編集は、旧 PDF をゴミ箱へ送り 3 列を空にしてから作り直す（`approval.gs:735-754, 833-857`、`worklog.gs:611-626`）。作り直しに失敗すると PDF が無いまま。朝バッチは**前日**分しか拾わない（`mail.gs:1455-1458`）。完了を targetDate の翌日以降に記録した（総務の手入力など）PDF の失敗も、二度と拾われない。一次承認で PDF が失敗すると台帳も書かない（G3-H2）。
- 計画: 2-1・Q13 で「完了済みで PDF なし 25 件（要確認）」。原因の見当は無い。outbox の再試行で動きが変わることも書いていない。
- なぜ困るか: VPS の outbox は失敗を再試行するので、今なら PDF なしで終わる申請に PDF ができる（良い方向の差）。「今と同じ」と言い切ると受け入れで食い違う。
- 直し方: Q13 に原因の候補（二次承認・編集の作り直し失敗、前日以外の完了、ロック待ちの失敗）を書く。「再試行で後から作られる」は改善として 3-4 に載せる。旧 PDF をゴミ箱へ送るのは新しい PDF ができた後にするかも 3-4 に足す（見た目は同じ、失敗に強くなる）。

### G3-M3（中）再試行で PDF が 2 つできる

- 根拠（確認済み）: Drive に保存（`pdfExport.gs:420`）してから Requests に記録（`423-427`）。間で失敗すると PDF だけ残る。同名は Drive が許す。
- 計画: 書いていない。
- 直し方: 仕事の中で「作った `pdfFileId` を先に jobs に記録 → Requests（PostgreSQL）へ」の順にし、再試行のときは記録済みの ID を使う。または保存時に Drive の `appProperties` に requestId を付け、再試行の前に探す。同じ人・同じ日・同じ種別の 2 件が同名になるのは今どおり（名前で上書き・重複除去をしない）と書く。

### G3-M4（中）値の書き方（入力と同じ扱い）

- 根拠（確認済み）: PDF も台帳も `setValue`／`setValues`（6 章）。
- 計画: 4-7 は HYPERLINK と日付だけ USER_ENTERED と書く。5-3 と罠 11 は「自由入力は RAW」。台帳でどちらを使うかがぶつかる。
- なぜ困るか: 台帳の理由・業務内容を RAW で書くと、今は日付や数式に変わっていた入力が文字のまま出る（見た目が変わる）。案 2 で PDF を Sheets API で書くなら、RAW だと `17:20`・日付の表示が変わる。
- 直し方: 4-7 に「台帳は全列 USER_ENTERED（GAS の `setValues` と同じ）。罠 11 の RAW は DB の Requests の書き出しだけ」と書く。案 2 を試すなら、PDF のセルも USER_ENTERED。

### G3-M5（中）書き出しで Requests タブを作り直すとテンプレの数式が壊れる

- 根拠: テンプレ（DB の中のタブ）の数式が Requests を引く（5 章、A29 は確認済み）。タブを消して作り直すと参照が `#REF!` になる（推測: Sheets の性質）。
- 計画: 5-3「名前で列を合わせて書く」。タブを消さないとは書いていない。
- 直し方: 5-3 に「Requests・WorkLogs のタブは消さず、値だけを書き換える。列の挿入・削除もしない」と書く。Sebastian を XLOOKUP 方式のまま残すなら必須。

### G3-L1（低）`insertImage` の例外で PDF 全体が落ちる

`pdfExport.gs:461-565` は `try` なし（3 章）。案 1 で写すなら同じ動きになる。直すなら「印が入らなくても PDF は作る」になり、見た目が変わる。今のままを既定にし、計画に一行書く。

### G3-L2（低）朝バッチが 1 行の壊れた targetDate で止まる

`listApprovedRequestsByDate_` は `targetDate` の変換を `try` で囲まない（`mail.gs:62-66`）。approved の行に変換できない `targetDate` があると、朝バッチ全体が例外で止まり、BatchLogs も書かない。VPS は日付を検証して入れるので起きにくい。移行の照合（5-4）で「approved で targetDate が日付にならない行」を数える項目を足す。

### G3-L3（低）手動の PDF 関数の扱いが計画に無い

`manualBatchPdfs_`（今日の分、BatchLogs に `manualBatchPdfs`）と `forceRegeneratePdf`（最新の approved 1 件。WorkLogs の開始を書き換え、Requests に列を足し、旧 PDF を消さない）。計画 4-4 は「保守の処理は管理者だけ」とだけ書く。移すなら「指定した申請 1 件を作り直す（旧 PDF はゴミ箱へ）」と「指定日の一括」の 2 つに整理し、WorkLogs の書き換えは移さない、と書く。

### G3-L4（低）G1 の作成日

G1 は PDF を作った日（`pdfExport.gs:453`）。二次承認の作り直しでも変わる。worker の遅れ・再試行で日をまたぐと変わる。worker が実際に作った JST の日付を渡す、と 4-6 に書く。

### G3-L5（低）台帳の細かい癖

- 却下は台帳の行を消さない（`approval.gs:348-384`）。完了後に却下された申請の行が `申請中` のまま残りうる。
- 行を消した後の縞のずれ（6 章）。VPS で「全行を塗り直す」にすると見た目が変わる。書いた行だけ塗る、と書く。
- `getAccumulationSS_` の自動作成をやめる（計画 4-7）のは良い。Settings に `ACCUMULATION_SS_ID` が 2 行あれば下が勝つ（`config.gs:83-87`）ので、取り込み（4-8）も「同じキーは下の行」にする。
- 年度をまたぐ休日の日付編集では、upsert が新しい年度のシートに行を足し、古い年度の行が残る（`mail.gs:670-702`）。

### G3-L6（低）ロックの記述

計画 3-1「PDF と台帳は承認のロックの外」は、一次・二次承認では正しい。完了系・総務の手入力・編集・Sebastian では PDF と台帳は呼び元のロックの中（入れ子）。VPS には影響しないが、表を直しておく。

### G3-L7（低）一時コピーに DB の GAS も写る

DB はコンテナに紐づく GAS を持つので、`makeCopy` のたびにスクリプトの写しもできる（推測）。残った一時コピーの中にコードの写しも残る。後片付け（計画 3-4）で一緒に消えるので、置き場所を決めるときに一行書く。

### G3-L8（低）`HR_MAIL_TO_EVENING_EXTRA` はこのコードで使っていない

grep で 0 件。統合メール通知システムが Settings を読んでいる手がかり。8-2 の「読み方を確かめる」に「Settings タブも読むか」を足す。

---

## 10. 計画の記述で正しいと確かめたもの

- 4-6 の今の作り（複製 → セル → `insertImage` → 書き出し URL → 日付フォルダ → 3 列 → ゴミ箱）。`pdfExport.gs:366-436, 440-566, 312-338`。
- 二次承認はいつも直接方式で G34 入り（`approval.gs:750, 851`）。Sebastian はいつも XLOOKUP 方式で G34 なし（`api.gs:105, 232`、`pdfExport.gs:164-308`）。
- ファイル名 `yyyyMMdd_部署_氏名_残業|休日出勤.pdf`、フォルダは対象日の `yyyy.MM.dd`。
- 朝バッチは未完了でも PDF を作る（`batchGeneratePdfs_` は WorkLogs を見ない）。
- 例外のとき一時コピーが残る。
- upsert の「1次承認日がいつも空」、朝バッチの「薄い行で置き換え」。
- Q8（PDF が無い申請の編集は台帳を更新しない、`worklog.gs:611-634`）。Q9（完了の再実行は実績を上書きし、PDF は `already` で作り直さない、台帳は upsert する）。
- メールは `sendEveningMail_`・`sendMorningMail_` だけで、トリガーからは呼ばれない。`morningBatch_` の送信行は 2026-03-11 にコメントアウト。「3/10 以降メールを送っていない」と合う。
- 直書きの Drive フォルダ `1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B`（`mail.gs:393, 1327`）。
- 印鑑の大きさと位置（F1 270px、F34・G34 60px で中央）と、2MB 以上の画像は縮小版を取ること。
- XLOOKUP 方式は切り替え後に正しく動かない見込み（4-5）。コピーの Requests が毎晩の写しになるため、筋は正しい（推測のまま）。

## 確かめること

- DB の `申請書フォーム` と `操作` の全数式、`操作!B3` の今の値（G3-H4）。
- 本番の PDF 3 枚（明細 1 件、業務内容なし、朝バッチで作った未完了分）の空欄（G3-H4）。
- PDF 1 件の所要時間（DB 全体の複製を含む）（G3-M1）。
- Settings に `ACCUMULATION_SS_ID` が何行あるか（G3-L5）。
