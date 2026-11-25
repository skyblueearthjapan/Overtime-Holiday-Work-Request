# 移行棚卸し 03: メール・PDF・蓄積スプレッドシート・Google フォーム

作成日: 2026-11-26 ／ 対象: `src/` の GAS コード（読み取りのみ。clasp・Google 側は見ていない）

この資料は、VPS 版（FastAPI + PostgreSQL + worker）で「メール・PDF・Drive のファイルが今とまったく同じに出る」ことを確かめる受け入れ表の元です。

- **確認済み** … コードを読んで確かめたこと（`ファイル:行` を付けています）
- **推測** … コードからの推論。本番の設定・トリガー・Drive の中身を見て確かめる必要があります
- **提案** … VPS 側の作り方の案

---

## 0. 先に結論（受け入れで外せない点）

1. **申請の入口は Web 画面のモーダル（`api_submitRequest`）が本流です。** Google フォームは 2026-03-03 のコミット `82cc3e0「フォーム廃止→モーダル直接送信化」` で画面からは外れました（確認済み: `top.html:733-751` の `openForm` は `showRequestModal` を呼び、`top.html:1959-2031` の `submitRequest` が `api_submitRequest` を呼ぶ）。ただし **フォームの仕組みは裏口として生きています**（`?page=redirect` の URL、1 分ごとのポーリング、毎朝 6:30 の選択肢更新）。詳しくは 4 章。
2. **このアプリからのメール送信は、今は止まっている可能性が高いです（推測）。** 2026-03-11 のコミット `fcf8e31「メール送信を統合メール通知システムに移行」` で、メールのトリガーを作る処理が外されました（`menu.gs:204-206`）。夕方メール・朝メールの関数はコードに残っており、メニューからは手で送れます。**「統合メール通知システム」はこのリポジトリの外にあります。** 総務部に届いている通知の本当の送り元を、移行前に確かめる必要があります。
3. **PDF はスプレッドシートのテンプレート（`TEMPLATE_SSID`）を複製し、Google の「スプレッドシート→PDF 書き出し URL」で作っています。** 同じ見た目を VPS で出すには、Google 側の仕組み（Sheets の描画）をそのまま使い続けるのが一番確実です（6 章の提案）。
4. **PDF の作り方が呼び出し元によって 2 通りに分かれています。** `PDF_MODE` 設定で決まる所、二次承認（いつも direct）、Sebastian 経由（いつも XLOOKUP 方式）が混在しています。本番の `PDF_MODE` の値は【要確認】です。
5. **蓄積スプレッドシートは書き込み経路で中身が違います。** リアルタイムの upsert と、朝バッチの「同じ日を消して書き直し」では、埋まる列が違います（3 章）。今の出力と同じにするなら、この差も再現するか、直すかを決める必要があります。

---

## 1. メール

### 1.1 送信している箇所（全部で 2 か所）

| # | 関数 | 送信の呼び出し | 確かさ |
|---|---|---|---|
| M1 | `sendEveningMail_()` 夕方メール | `GmailApp.sendEmail(to, subject, body)` `mail.gs:335` | 確認済み |
| M2 | `sendMorningMail_(targetDateObj)` 朝メール | `GmailApp.sendEmail(to, subject, bodyLines.join('\n'))` `mail.gs:1419` | 確認済み |

- `MailApp` は使っていません。`src/` 全体で `GmailApp.sendEmail` はこの 2 か所だけです（確認済み: grep）。
- 承認者・作業員への個別メール（申請・承認・却下の通知）は**ありません**（確認済み）。
- `Requests.hrMailSentAt` 列は `appendRequestRow_` で空文字を入れるだけで、どこからも書かれていません（確認済み: `formSubmit.gs:572,641,793`）。

### 1.2 いつ送られるか（トリガー）

| 予定 | 関数 | 作る処理 | 今の状態 |
|---|---|---|---|
| 毎日 16 時台（`nearMinute(45)`）＝16:45 頃 | `sendEveningMail_` | `setupMailTriggers_` `menu.gs:121-126` | `setupAllTriggers_` から外されています（`menu.gs:204-205`）。**過去に作ったトリガーが本番に残っているかは【要確認】** |
| 毎日 18 時台（`nearMinute(10)`）＝18:10 頃 | `sendEveningMail_` | `menu.gs:129-134` | 同上 |
| 毎日 7 時台（`nearMinute(10)`） | `morningBatch_`（PDF＋蓄積＋メール…だったが、メール行はコメントアウト `mail.gs:1446-1447`） | `menu.gs:138-145` | 同上。動いていても今はメールを送りません |
| 毎日 6 時台（`nearMinute(50)`） | `morningBatchNoMail_`（PDF＋蓄積のみ） | `setupMorningBatchTrigger_` `menu.gs:212-221` | `setupAllTriggers_` に入っている（`menu.gs:206`）。**動いていると推測** |
| 手動 | `sendEveningMail_` / `morningBatch_` / `sendMorningMail_` | メニュー `menu.gs:10-12` | スプレッドシートのメニュー「残業休日アプリ」から押せる |

- 補足: `setupAllTriggers_` は古いメールトリガーを**消しません**。3/11 より前に `setupMailTriggers_` を動かしていれば、`sendEveningMail_`×2 と `morningBatch_` が本番に残っている可能性があります（推測）。この場合、夕方メールは今も GAS から送られています。**GAS のトリガー一覧画面で確かめてください。**
- `docs/form-integration-spec.md` の「17:10, 18:10, 7:10、CSV/Excel 添付」は古い記述です。今のコードとは違います（確認済み）。

### 1.3 宛先・差出人

| 項目 | 内容 | 確かさ |
|---|---|---|
| To | Settings の `HR_MAIL_TO`（`mail.gs:328`, `mail.gs:1318`）。そのまま第 1 引数に渡すので、カンマ区切りで複数可（GmailApp の仕様） | 確認済み |
| Cc / Bcc / replyTo / name | 指定なし（第 4 引数なし） | 確認済み |
| 未設定時 | `HR_MAIL_TO` が空なら例外 `Settingsに HR_MAIL_TO が未設定です。` で送らない | 確認済み |
| 差出人 | トリガーを作った人（時間トリガー）または、メニューを押した人の Gmail。表示名は Gmail アカウントの既定 | 推測（GAS の仕様） |
| 形式 | テキストのみ（HTML 本文なし） | 確認済み |
| 添付 | **なし**。CSV・XLSX を作る関数（`makeCsvBlob_` `mail.gs:1262`、`exportRowsToXlsxBlob_` `mail.gs:1276`、`buildMorningReportRows_` `mail.gs:1199`）は残っていますが、**どこからも呼ばれていません** | 確認済み（grep） |
| 送信上限への対処 | なし（残り回数の確認・再送・記録なし） | 確認済み |
| スコープ | `appsscript.json` の `oauthScopes` は `gmail.send` のみ。`GmailApp.sendEmail` は通常 `https://mail.google.com/` を要るため、明示スコープのままでは送れない可能性あり | 推測【要確認】 |

### 1.4 夕方メール（M1）の形

件名（`mail.gs:332`。日付は送る瞬間の JST）:

```
【残業・休日出勤】申請状況報告 yyyy/MM/dd
```

本文の作り方 `buildEveningMailBody_(now)` `mail.gs:236-320`:

- 残業: 対象日（=今日）の申請のうち、`canceled` 以外すべて（`listAllRequestsByDate_` `mail.gs:3-44`、並びは「部署+氏名」の日本語順）。
- 休日: `submitted` は**日付を問わず全件**、`approved` は**今日の分だけ**（`listHolidayRequestsForWeekend_` `mail.gs:172-232`、並びは対象日→部署+氏名）。
- 状態ラベル: `approvedBy2` があれば `二次承認済`、なければ `approved` で `承認済`、それ以外 `未承認`（`mail.gs:276-281`, `300-305`）。
- 時間表記 `fmtMinutesJa_`（`mail.gs:157-164`）: 0 時間なら `N分`、端数 0 なら `N時間`、それ以外 `N時間M分`。

0 件のとき（`mail.gs:250-258`）。`.filter(Boolean)` で空行も消えるので、**空行の入らない 3 行（APP_URL が無ければ 2 行）**になります。末尾の「※本メールは自動送信です。」もありません:

```
【残業・休日出勤 申請状況報告】yyyy/MM/dd
本日分の残業申請および休日出勤申請（未承認）はありません。
詳細（アプリ）：{APP_URL}
```

1 件以上のとき（`mail.gs:260-319`）:

```
【残業・休日出勤 申請状況報告】yyyy/MM/dd

【本日の残業申請】

■ {部署}
- {氏名}：残業 {承認時間}（{状態}）

【休日出勤申請（未承認＋当日承認済み）】

■ {部署}
- {氏名}：休日出勤 {M/d(曜)} {承認時間}（{状態}）

詳細（アプリ）：{APP_URL}
二次承認ページ：{APP_URL}?page=approve2&date=yyyy-MM-dd

※本メールは自動送信です。
```

- 残業・休日の節は、その種類が 0 件なら見出しごと出ません。部署ごとに最後に空行 1 つ。
- 休日の日付ラベルは `new Date(targetYmd + 'T00:00:00')` から作るので、実行環境の時刻帯に依存します（GAS は Asia/Tokyo）。
- APP_URL が空なら「詳細」「二次承認ページ」と直後の空行を出しません。

### 1.5 朝メール（M2）の形

対象日: 引数が無ければ**前日**（`mail.gs:1322-1324`）。件名 `【残業・休日出勤】実績報告 yyyy/MM/dd`（`mail.gs:1335`）。

本文 `mail.gs:1337-1417`:

```
【⚠ 作業完了未記録のお知らせ】
以下の作業員は作業完了ボタンが押されておらず、実働時間が記録されていません。
該当の作業員に連絡し、TOP画面から手入力で実績を記録するよう依頼してください。
※ 総務部画面（日次詳細）からも手動入力が可能です。

▼ yyyy/MM/dd の未記録
  ■ {部署}
    - {氏名}：{残業|休日出勤} 承認{承認時間}
  → 作業員TOP画面：{APP_URL}?page=top
  → 総務部手動入力：{APP_URL}?page=admin&date=yyyy-MM-dd

---

【残業・休日出勤 実績報告】yyyy/MM/dd

■ {部署}
- {氏名}：{残業|休日出勤} 承認{X} / 実働{X} / 休憩{X} / 実残業{X} [PDF作成済|PDF未作成]

PDF作成状況：残業 N件 / 休日 N件 / 合計 N件

PDFフォルダ：{PDF_DRIVE_FOLDER_URL}
データ蓄積：https://docs.google.com/spreadsheets/d/{ACCUMULATION_SS_ID}/edit
詳細（アプリ）：{APP_URL}

※本メールは自動送信です。
```

- 未記録の節は、`approved` で対象日が「今日より前・7 日前以降」、WorkLogs の `actualEndAt` が空のもの（`listIncompleteRequests_` `mail.gs:1133-1191`）。無ければ節ごと出ません。
- 実績の対象は前日の `approved` のみ。0 件のときは `本日の承認済み申請はありません。` と空行（前日分なのに「本日」と書く。今の文言どおり）。
- `PDF_DRIVE_FOLDER_URL` が未設定なら既定値 `https://drive.google.com/drive/folders/1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B?usp=sharing`（`mail.gs:1327`、コードに直書き）。
- 「⚠」は絵文字ではなく記号ですが、文字コードはそのまま再現が要ります。

### 1.6 使われていない CSV / XLSX の仕様（参考・確認済み）

今は添付されませんが、移行後に「元に戻す」話が出たときのために残します。

- 列（`mail.gs:1204-1221`）: `日付, 部署, 氏名, 種別, 承認時間(分), 承認時間, 開始, 終了, 実働(分), 実働, 休憩(分), 休憩, 実残業/実働(分), 実残業/実働, PDF作成, requestId`
- CSV: カンマ区切り、改行 `\n`、`"` `,` `\n` を含むときだけ `"` で囲む、BOM なし、`text/csv`、文字コードは UTF-8（`Utilities.newBlob` の既定、推測）。
- XLSX: 一時スプレッドシート `TMP_EXPORT_{filename}_{ミリ秒}` の `Report` シートに書き、`/export?format=xlsx` で取り出してゴミ箱へ。
- 不具合: 開始・終了は `wl.actualStartAt instanceof Date` のときだけ埋まりますが、`buildWorkLogsMapByRequestId_` は文字列を返すので**常に空**になります（`mail.gs:1229-1230` と `mail.gs:110-111`）。

---

## 2. PDF（申請書）

### 2.1 2 つの方式

| | XLOOKUP 方式 `generatePdfForRequest_` `pdfExport.gs:164-308` | 直接書き込み方式 `generatePdfDirect_` `pdfExport.gs:366-436` |
|---|---|---|
| 選ばれる条件 | `PDF_MODE` に `direct` を含まない（未設定を含む） | `PDF_MODE` に `direct` を含む |
| テンプレート | どちらも `TEMPLATE_SSID` を Drive で複製 `TMP_{requestId}_{yyyyMMdd_HHmmss}` | 同じ |
| 値の入れ方 | 複製の `操作!B3` に requestId を入れ、テンプレート内の数式（XLOOKUP）で埋める。flush → 600ms 待ち → flush。そのうえで理由・時刻・休憩・実残業・区分（休日のみ）を直接上書き | `fillPdfTemplate_` `pdfExport.gs:440-566` で全項目をセルに直接書く |
| 2 次承認印（G34） | **入れない** | 入れる |
| 作成日（G1）・部署・氏名・明細・業務内容 | テンプレートの数式まかせ | 直接書く（G1 = PDF を作った日 `yyyy/MM/dd`） |

- XLOOKUP の参照先（テンプレート内のどのシート・どのデータか、IMPORTRANGE か）はコードからは分かりません【要確認: テンプレートの中身】。
- 本番の `PDF_MODE` の値は【要確認】。

呼び出し元ごとの方式（確認済み）:

| 呼び出し元 | 方式 |
|---|---|
| 一次承認 `api_approveRequest` `approval.gs:329-341`、一括承認 `approval.gs:455-473` | `PDF_MODE` に従う |
| 残業完了 `api_markOvertimeDone` `worklog.gs:193-207`、休日完了 `worklog.gs:322-336` | `PDF_MODE` に従う（計算済みの実績 `wlOverride` を渡す） |
| 総務の手入力 `api_adminManualComplete` `adminApi.gs:498-512` | `PDF_MODE` に従う（`wlOverride` あり） |
| 申請の編集 `api_editRequest` `worklog.gs:610-634` | PDF があれば消して `PDF_MODE` で作り直す |
| 朝バッチ・手動一括 `batchGeneratePdfs_` `pdfExport.gs:570-600` | `PDF_MODE` に従う |
| 手動再生成 `forceRegeneratePdf` `pdfExport.gs:743-871` | `PDF_MODE` に従う |
| **二次承認** `api_secondaryApprove` `approval.gs:745-757`、一括 `approval.gs:846-858` | **いつも direct** |
| **Sebastian 経由の完了** `api.gs:102-110`、`api.gs:229-236` | **いつも XLOOKUP** |

### 2.2 いつ作られるか

| きっかけ | 条件 | 出典 |
|---|---|---|
| 一次承認（個別・一括） | WorkLogs の `actualEndAt` がある（作業完了済み） | `approval.gs:318-341`, `441-473` |
| 作業完了（残業・休日・Sebastian・総務手入力） | `status === 'approved'` | `worklog.gs:195`, `324`, `api.gs:103`, `230`, `adminApi.gs:498` |
| 二次承認 | 既に PDF がある場合だけ、古い PDF をゴミ箱へ → 作り直し（2 次承認印入り） | `approval.gs:735-757`, `873-883` |
| 申請の編集 | PDF がある場合だけ、ゴミ箱へ → 作り直し | `worklog.gs:610-634` |
| 申請の削除（キャンセル） | PDF があればゴミ箱へ。作り直しはしない | `worklog.gs:668-672` |
| 朝バッチ `morningBatchNoMail_`（6:50 頃）/`morningBatch_` | 前日の `approved` で `pdfFileId` が空のもの。**作業完了かどうかは見ない**（実績が空の PDF もできうる） | `mail.gs:1427-1468`, `pdfExport.gs:570-600` |
| メニュー「PDF一括生成（本日分）」 | 今日の分で同じ条件 | `menu.gs:33-43` |

- 既に `pdfGeneratedAt` と `pdfFileId` が両方あれば作らずに `{already:true}` を返します（`pdfExport.gs:168-170`, `373-377`）。
- 一括では 6 件以上のとき 1 件ごとに 500ms 待ちます（`pdfExport.gs:596`）。
- 失敗は警告ログのみで、本体の操作（承認・完了）は成功扱い。朝バッチの結果は `BatchLogs` シートに `実行日時, バッチ名, 対象日, 成功, スキップ, 失敗, エラー詳細` で追記（`pdfExport.gs:604-627`）。

### 2.3 テンプレートのセル（`PDF_MAP` `pdfExport.gs:3-23`）

シート名は `申請書フォーム`（XLOOKUP 方式は加えて `操作` シートの B3）。

| セル | 内容 | direct 方式で入れる値 |
|---|---|---|
| G1 | 作成日 | PDF を作った日 `yyyy/MM/dd` |
| B4 | 部署 | 部署名（中央揃えにする） |
| D4 | 氏名 | `workerName` |
| F4 | 種別（大） | 区分印の画像があれば空にする。無ければ `残業`/`休日出勤` を 24pt 太字・中央 |
| C6 | 区分 | 残業は `残業`。休日は `approvedMinutes <= 240` なら `半日`、それ以外 `1日` |
| C7 | 作業実施日 | `yyyy/MM/dd` |
| C10 | 開始 | 残業は**いつも `17:20`**。休日は `actualStartAt` の `HH:mm` |
| F10 | 終了 | `actualEndAt` の `HH:mm`（あれば） |
| C12 | 休憩 | `fmtMinToHM_(breakMinutes)`（例 `10分`、`1時間`、`0分`） |
| F12 | 実残業 | `sanitizeNetMinutes_` 後の分を `fmtMinToHM_`。残業 720 分・休日 960 分を超えたら開始〜終了−休憩で計算し直す（`pdfExport.gs:40-65`） |
| B18/D18/F18〜B20/D20/F20 | 明細 3 行（工番・得意先・品名） | `workNoN`（無ければ `orderNoN`）、`customerN`、`productN`。空なら書かない |
| A23 | 業務内容 | `workContent`（空なら書かない） |
| A29 | 理由 | `buildReasonText_`（`pdfExport.gs:131-146`）: 理由から「（…補足理由…）」を取り除き、補足理由があれば改行でつなぐ。空でも書く（数式を消すため） |
| F34 | 一次承認印 | 画像のみ（メールアドレスは書かない） |
| G34 | 二次承認印 | 画像のみ（direct 方式だけ） |

- 「休憩 0 分」のときも `0分` と書きます（`fmtMinToHM_` は 0 → `0分`）。
- XLOOKUP 方式で直接上書きするのは A29・C10・F10・C12・F12・C6（休日）と画像（`pdfExport.gs:201-273`）。

### 2.4 印鑑画像

| 印 | 画像の決まり方 | 位置・大きさ |
|---|---|---|
| 区分印（残業／休日出勤） | Settings `STAMP_OVERTIME_FILE_ID` / `STAMP_HOLIDAY_FILE_ID`（Drive の画像 ID） | F1 起点、オフセット 0、270×270px（`pdfExport.gs:242-244`, `466-468`） |
| 一次承認印 | `approvedBy`（メール）で `StampMap` シートを引く（見出し `メール` と `stampFileId`（大文字小文字は無視）、メールは小文字で比較）`pdfExport.gs:103-127` | F34 起点、60×60px。横は F 列幅の中央、縦は 34〜37 行の高さの合計の中央 |
| 二次承認印 | `approvedBy2` で同じく StampMap | G34 起点、60×60px、同じ中央寄せ |

- 画像は 2MB 未満ならそのまま、2MB 以上なら `https://drive.google.com/thumbnail?id={id}&sz=400` で縮小版を取ります（`pdfExport.gs:73-97`）。
- 承認者が StampMap に無い・画像が読めないときは、印なしで続けます。

### 2.5 書き出しの設定・ファイル名・置き場所

- 書き出し URL（`pdfExport.gs:312-338`）:
  `https://docs.google.com/spreadsheets/d/{複製ID}/export?format=pdf&gid={申請書フォームのgid}&portrait=true&size=A4&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`
  余白・拡大率・水平位置は指定なし（テンプレートの印刷設定と Google の既定に従う）。`Authorization: Bearer {ScriptApp.getOAuthToken()}`。200 以外は例外。
- ファイル名: `{対象日 yyyyMMdd}_{部署}_{氏名}_{残業|休日出勤}.pdf`。部署と氏名の `\ / : * ? " < > |` は `_` に置き換え（`pdfExport.gs:282-286`, `410-414`）。
- 置き場所: `PDF_ROOT_FOLDER_ID` の直下に、**対象日**の `yyyy.MM.dd` フォルダ（年・月・部署の階層は無し）。同名フォルダがあれば最初の 1 つを使い、無ければ作る（`pdfExport.gs:150-158`）。
- 記録: `Requests` の `pdfGeneratedAt`（Date）、`pdfFileId`、`pdfFolderId`（`pdfExport.gs:295-299`, `423-427`）。
- 複製は作業後にゴミ箱へ。**途中で例外が出ると複製はゴミ箱へ行かず残ります**（`try/finally` はロック解放だけ。確認済み）。複製の置き場所は `makeCopy(name)` の既定（推測: テンプレートと同じフォルダ）。
- 二次承認・編集・削除では古い PDF を `setTrashed(true)` でゴミ箱へ（完全削除ではない）`approval.gs:873-883`。
- 排他: どちらの方式も `LockService.getScriptLock().waitLock(20000)` を取り、最後に `releaseLock()` します。呼び出し元（完了 API など）が同じロックを持ったまま呼ぶため、PDF の後はロックが外れた状態で続きます（推測: GAS のロックの仕様）。

### 2.6 手動・診断用の関数（移行では再現不要の見込み）

`forceRegeneratePdf`（最新の承認済み 1 件を作り直す。WorkLogs の残業開始を 17:20 に直し、補足理由をフォーム回答から補う `tryFillReasonDetailFromForm_` `pdfExport.gs:994-1093` も呼ぶ）、`debugPdfReason`、`debugStampData`。

---

## 3. 蓄積スプレッドシート（総務確認用台帳）

### 3.1 ファイル

- ID は Settings `ACCUMULATION_SS_ID`。開けなければ**新しく作り**、名前は `残業・休日出勤 データ蓄積（総務確認用）`、フォルダ `1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B`（コードに直書き）へ移し、Settings の末尾に `ACCUMULATION_SS_ID, {新ID}` を追記します（`mail.gs:390-419`）。移行後に ID が無効になると勝手に新しいファイルができる点に注意。

### 3.2 シートと列

- シート名: `残業_{年度}年度` / `休日出勤_{年度}年度`。年度は **3/16 始まり**（3/15 までは前年度）`mail.gs:465-472`。新しいシートは先頭（位置 0）に作り、`Sheet1`/`シート1` を消します（`mail.gs:428-460`）。
- 見出し（1 行目、27 列。`ACC_HEADERS_` `mail.gs:343-351`）と列幅（`mail.gs:354-382`）:

| # | 見出し | 幅 | 値（`buildAccRow_` `mail.gs:519-618`） |
|---|---|---|---|
| 1 | 管理番号 | 140 | requestId（UUID） |
| 2 | 申請種別 | 80 | `残業` / `休日出勤` |
| 3 | 作業実施日 | 100 | `yyyy/MM/dd`（文字列で書くのでシート側で日付になる見込み・推測） |
| 4 | 曜日 | 50 | `日`〜`土` |
| 5 | 部署 | 100 | dept |
| 6 | 作業員ID | 90 | workerCode |
| 7 | 氏名 | 100 | workerName |
| 8 | 理由 | 160 | reason |
| 9 | 補足理由 | 160 | reasonDetail |
| 10 | 業務内容 | 160 | workContent |
| 11 | 承認時間(分) | 80 | approvedMinutes（数値） |
| 12 | 承認時間 | 100 | `fmtMinutesJa_` |
| 13 | 開始時刻 | 80 | `HH:mm` |
| 14 | 終了時刻 | 80 | `HH:mm` |
| 15 | 実働(分) | 80 | actualMinutes |
| 16 | 実働 | 100 | 表記 |
| 17 | 休憩(分) | 80 | breakMinutes |
| 18 | 休憩 | 100 | 表記 |
| 19 | 実残業/実働(分) | 80 | netMinutes |
| 20 | 実残業/実働 | 100 | 表記 |
| 21 | 申請日 | 100 | submittedAt `yyyy/MM/dd` |
| 22 | 1次承認日 | 100 | approvedAt `yyyy/MM/dd` |
| 23 | 1次承認者 | 120 | approvedBy（メール） |
| 24 | 2次承認日 | 100 | approvedAt2 `yyyy/MM/dd` |
| 25 | 2次承認者 | 120 | approvedBy2（メール） |
| 26 | 承認状態 | 80 | `承認済` / `申請中` / `取消` |
| 27 | PDF | 60 | `=HYPERLINK("https://drive.google.com/file/d/{id}/view","PDF")`（ID が空なら空） |

- 書式: 見出しは太字・背景 `#1a365d`・文字白・10pt・中央・上下中央、1 行目固定、行の高さ 32（`mail.gs:437-445`）。データ行は 10pt・上下中央、**行番号が偶数なら背景 `#f8fafc`、奇数なら `#ffffff`**（`formatAccRow_` `mail.gs:775-795`）。

### 3.3 書き込み経路と、経路ごとの違い

| 経路 | いつ | キー | 確かさ |
|---|---|---|---|
| `upsertAccumulationRow_(requestId)` `mail.gs:651-710` | 作業完了（全経路）、一次承認で PDF を作ったとき、二次承認、編集で PDF があったとき | 1 列目＝requestId。あれば上書き、無ければ末尾に追記 | 確認済み |
| `removeAccumulationRow_` `mail.gs:716-741` | 申請の削除（キャンセル）`worklog.gs:690-694` | 1 列目＝requestId、下から探して最初の 1 行を消す | 確認済み |
| `appendToAccumulationSS_(前日)` `mail.gs:489-514` | 朝バッチ（6:50 頃） | 3 列目（作業実施日）が同じ日の行を**全部消して**、前日の `approved` だけ書き直す | 確認済み |
| `rebuildAccumulationSS()` `mail.gs:943-1127` | メニュー（手動） | 既存の年度シートを `{名前}_backup_yyyyMMdd_HHmmss` に名前替えし、Requests 全件（canceled 以外）から作り直す。作業実施日順 | 確認済み |
| `migrateAccumulationTimeColumns()` `mail.gs:806-927` | メニュー（手動・1 回限り） | 旧 23 列のシートに表記列 4 つを差し込む | 確認済み |

経路ごとに埋まる列が違います（確認済み。受け入れで「どちらに合わせるか」を決める必要あり）:

- **upsert**: `getRequestById_`（`worklog.gs:51-81`）が `approvedAt` を返さないため、**22 列目「1次承認日」はいつも空**。承認状態は今の status（`申請中` もありうる）。
- **朝バッチ**: 元データ `listApprovedRequestsByDate_`（`mail.gs:48-85`）に reason・reasonDetail・workContent・approvedBy・approvedAt2・approvedBy2・status が無いため、**8〜10 列・23〜26 列が空**になります（「承認状態」も空）。しかも同じ日の行を消して書き直すので、**upsert で入っていた詳しい行が、翌朝この薄い行に置き換わります**。作業員 ID は Requests を 1 行ずつ探し直します（`mail.gs:525-545`）。
- **再構築**: 全列が埋まります。
- `getRequestFullData_` が `mail.gs:746` と `pdfExport.gs:342` の 2 か所で定義されています。GAS では後から読み込まれた方が勝ちます（推測: ファイル順で `pdfExport.gs` の版）。どちらでも upsert に要る項目は取れます。

---

## 4. Google フォーム

### 4.1 今の入口（まとめ）

| 入口 | 生きているか | 根拠 |
|---|---|---|
| Web 画面の申請ボタン → モーダル → `api_submitRequest` `formSubmit.gs:678-812` | **本流（使われている）** | `top.html:338-342` → `openForm` `top.html:733-751` → `showRequestModal` → `submitRequest` `top.html:1959-2031`。コミット `82cc3e0`（2026-03-03） |
| Web 画面からフォームを開く `doOpenForm` → `api_getFormUrl` `top.html:940-971` | 画面からは届かない（死んだ経路） | 呼び元の `showDatePicker` を呼ぶ所が無い（grep で確認） |
| `?page=redirect&type=overtime&dept=…` → `handleFormRedirect_` → `api_getFormUrl` → フォームへ移動 | **URL を知っていれば今も使える** | `adminApi.gs:539-541`, `611-657` |
| フォームの公開 URL（ブックマーク等）から直接回答 | **使える**（回答を受け付けたまま） | `formGenerate.gs:83`、ポーリングが動いていれば取り込まれる |
| 1 分ごとのポーリング `pollNewResponses_` | `setupAllTriggers_` で作る（`menu.gs:207`, `182-196`）。動いていると推測 | 【要確認: トリガー一覧と各フォームの最近の回答数（`debugPollStatus_`）】 |

**移行で決めること:** フォームからの回答が今も来ているなら、VPS でも取り込む（Forms API で読む）必要があります。来ていなければ、フォームを「回答を受け付けない」にして入口を閉じるのが簡単です。**判断の材料として、FormMap の全フォームの直近の回答数を見てください。**

### 4.2 フォームの作り方

- テンプレート ID: Settings `OVERTIME_TEMPLATE_FORM_ID` / `HOLIDAY_TEMPLATE_FORM_ID` を先に見て、無ければ `FormTemplates` シート（見出し `type`, `templateFormId`）`formMap.gs:65-99`。
- `getOrCreateDeptForm_(type, dept)` `formGenerate.gs:3-95`:
  1. `FormMap`（見出し `type, dept, formId, formUrl, updatedAt, isActive`、無ければ自動作成 `formMap.gs:3-12`）で (type, dept) を探し、あれば公開 URL を返す。
  2. その部署の作業員が 0 人なら作らない。
  3. テンプレートを Drive で複製。名前 `【残業申請】{部署}` / `【休日申請】{部署}`。Settings `FORMS_PARENT_FOLDER_ID` があればそこへ。
  4. 共有を「リンクを知っている全員・閲覧」、だめなら「ドメイン内・リンク・閲覧」。
  5. `setPublished(true)`、タイトル、説明文（下記）、選択肢の設定、`setAcceptingResponses(true)`、FormMap に登録。
- 説明文（`formGenerate.gs:60-65`）:

```
【運用案内】
・送信後の修正は原則できません。
・トップ画面に申請が表示され、承認後に「承認済み」ラベルに変わります。
・実績はトップ画面のボタン（残業=完了のみ／休日=開始・完了）で記録します。
```

- 一括作成 `buildAllDeptForms`（メニュー）`menu.gs:57-82`。

### 4.3 質問と選択肢（タイトルで探す。`Q` `config.gs:18-40`）

| 質問タイトル | 種類 | 選択肢・中身 |
|---|---|---|
| 申請種別 | プルダウン | 1 択固定 `残業` または `休日` |
| 部署 | プルダウン | 1 択固定 `{部署}` |
| 作業員 | プルダウン | 作業員マスタのその部署・在籍の人を `{作業員コード} {氏名}` で、コード順（`masters.gs:10-46`） |
| 作業実施日 | 日付 | ― |
| 工番1 / 工番2 / 工番3 | 短文（任意） | 説明「WEBアプリのモーダルから選択するか、工番コードを直接入力してください（任意）」。作業実施日の直後に並べ直す（`formUtils.gs:30-149`） |
| 業務内容 | ― | 回答があれば `workContent` に入れる（選択肢は触らない） |
| 理由 | プルダウン | `REASONS`（`config.gs:42-49`）6 つ |
| 補足理由 | 短文 | 「その他」のとき必須（サーバでも確認） |
| 予定時間 | プルダウン | 残業 `0.5`〜`4.0`（0.5 刻み 8 つ）、休日 `半日` / `1日` |

- 旧形式の工番質問（`工番`、`工番プレフィックス_N`、`工番番号_N`、「追加の工番…」「さらに…工番」、工番を含む節見出し）は `ensureOrderItems_` が**本番フォームから消します**（`formUtils.gs:30-121`）。`docs/holiday-form-template-manual.md` の 16 問構成・節分岐はこの処理で崩れる前提の古い手順です。
- テンプレートの「業務ID」質問（手引きでは必須・仮の選択肢）をコードは触らず、回答も読みません（`jobId1` は常に空）。本番フォームに残っていると回答者が仮の選択肢を選ぶことになります【要確認】。
- 毎朝の更新 `nightlyUpdateAllForms_` `formBatch.gs:3-41`: 6 時台（`nearMinute(30)`）`menu.gs:86-99`。FormMap の有効な全フォームに `updateDeptFormChoices_`（作業員・工番欄の整備・理由・予定時間・種別/部署の固定）`formGenerate.gs:99-131`。マスタ同期 `syncAllMasters`（6:00）の後に走る想定。

### 4.4 事前入力 URL

`api_getFormUrl(type, dept, workerLabel, targetDateStr, orderCodes)` `formMap.gs:103-126` → `buildPrefillUrl_` `formMap.gs:132-169`: 申請種別・部署・作業員・作業実施日（指定日または今日、正午で作る）・工番 1〜3 を入れた `toPrefilledUrl()` を返します。毎回 `ensureOrderItems_` も走ります。

### 4.5 回答の取り込み

- `onFormSubmit` トリガーは使っていません（`addFormSubmitTrigger_` は何もしない `formSubmit.gs:6-8`。トリガー 20 個の上限を避けるため）。
- `pollNewResponses_` `formSubmit.gs:14-114`（1 分ごと）:
  - スクリプトのロックを 10 秒で取れなければ何もしない。
  - Script Properties `POLL_LAST_TS` 以降の回答を、FormMap の有効な全フォームから `getResponses(since)` で取る（初回は 24 時間前から）。
  - Script Cache `POLL_PROCESSED_IDS`（直近 100 件、10 分で消える）にある回答 ID は飛ばす。
  - 1 件でも失敗があれば `POLL_LAST_TS` を進めない（全件を次回やり直し）。失敗が 10 分以上続くと、キャッシュが消えて**成功済みの回答も二重に取り込まれます**（確認済み: 論理上）。
  - Requests に「どの回答から来たか」を残さないので、後から重複を見分ける手がかりはありません。
- `handleFormSubmit_` `formSubmit.gs:420-593` の変換:

| フォーム | Requests 列 | 変換 |
|---|---|---|
| 申請種別 | `requestType(overtime/holiday)` | `残業`→`overtime`、それ以外→`holiday` |
| 部署 | `dept` | そのまま（マスタでの上書きなし） |
| 作業員 | `workerCode` / `workerName` | 最初の空白で分ける。メールは作業員マスタから `lookupWorkerEmail_` |
| 作業実施日 | `targetDate` | `yyyy-MM-dd` の文字列 |
| 工番1〜3 | `workNo1〜3`、`orderNo1〜3` | `工番マスタ` で `customerN`（納入先、無ければ受注先）・`productN`（品名）を補う。無ければ `exportError` に `工番マスタ未登録: …` |
| 業務内容 | `workContent` | ― |
| 理由 / 補足理由 | `reason` / `reasonDetail` | `その他` で始まり補足が空なら例外（取り込まない） |
| 予定時間 | `approvedMinutes` | 残業は時間×60、休日 `半日`=240 / `1日`=480 |
| （自動） | `requestId`=UUID、`status`=`submitted`、`submittedAt`=取り込んだ時刻（回答時刻ではない） | ― |

- 過去日・年度外でもログに警告を出すだけで受け付けます（`formSubmit.gs:471-499`）。モーダル経路（`api_submitRequest`）は同じ条件で**拒否**します。経路で決まりが違います。
- WorkLogs に空の行（`updatedBy`=`formSubmit`）を作ります（`formSubmit.gs:816-837`）。
- 手動の救済: `debugPollStatus_`、`debugProcessLatest_`、`reprocessAllResponses_`（48 時間前から全部取り込み直す。**重複を防がない**）`formSubmit.gs:120-297`。
- 回答者の体験: テンプレートの確認メッセージ（手引きでは「申請を受け付けました。トップ画面に反映されます。」）が出て、最大およそ 1 分後に TOP 画面に出ます。

---

## 5. 外部の ID・設定

### 5.1 Settings シートのキー（この領域で使うもの）

`getSettings_` は Settings の 3 行目以降の A 列＝キー、B 列＝値を読みます（`config.gs:76-90`）。値はコードに無いので、本番シートから控えてください。

| キー | 用途 | 出典 |
|---|---|---|
| `HR_MAIL_TO` | 夕方・朝メールの宛先 | `mail.gs:328`, `1318` |
| `APP_URL` | メール本文のリンク | `mail.gs:238`, `1326` |
| `PDF_DRIVE_FOLDER_URL` | 朝メールの PDF フォルダリンク（既定値あり） | `mail.gs:1327` |
| `ACCUMULATION_SS_ID` | 蓄積スプレッドシート | `mail.gs:392`, `1328` |
| `PDF_ROOT_FOLDER_ID` | PDF の保存先ルート | `pdfExport.gs:173`, `380` |
| `TEMPLATE_SSID` | 申請書テンプレート | `pdfExport.gs:174`, `381`, `695` |
| `PDF_MODE` | `direct` を含めば直接書き込み方式 | `pdfExport.gs:577` ほか 6 か所 |
| `STAMP_OVERTIME_FILE_ID` / `STAMP_HOLIDAY_FILE_ID` | 区分印の画像 | `pdfExport.gs:234`, `459` |
| `OVERTIME_TEMPLATE_FORM_ID` / `HOLIDAY_TEMPLATE_FORM_ID` | フォームのテンプレート | `formMap.gs:67-69` |
| `FORMS_PARENT_FOLDER_ID` | 複製フォームの置き場所 | `formGenerate.gs:32` |

（ほかの領域のキー: `ADMIN_EMAILS`、`GENERAL_AFFAIRS_EMAILS`、`SOURCE_SSID`。Script Properties: `SEBASTIAN_API_TOKEN`）

### 5.2 シート（DB スプレッドシート内）

`StampMap`（メール→印鑑画像 ID）、`FormMap`、`FormTemplates`、`BatchLogs`、`Settings`、`Requests`、`WorkLogs`（2 行目が見出し）、`作業員マスタ`、`工番マスタ`。

### 5.3 その他の保存場所

- Script Properties `POLL_LAST_TS`（ISO 文字列）、Script Cache `POLL_PROCESSED_IDS`。

### 5.4 コードに直書きの値

| 値 | 場所 | 用途 |
|---|---|---|
| Drive フォルダ `1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B` | `mail.gs:393`, `mail.gs:1327` | 蓄積 SS を新しく作るときの置き場所／PDF フォルダリンクの既定値 |
| `https://drive.google.com/file/d/{id}/view` | `mail.gs:792` | 蓄積 SS の PDF リンク |
| `https://drive.google.com/thumbnail?id={id}&sz=400` | `pdfExport.gs:86` | 大きい印鑑画像の縮小 |
| スプレッドシート書き出し URL（pdf / xlsx） | `pdfExport.gs:313-323`, `mail.gs:1288` | PDF・XLSX |
| 固定の時刻 `17:20` | `pdfExport.gs:212`, `492` | 残業の開始時刻 |
| 祝日カレンダー `ja.japanese#holiday@group.v.calendar.google.com` | `formMap.gs:331` | 休日候補日（予備の経路） |

---

## 6. GAS の外で再現しにくい所と、VPS 側の案

| 項目 | 難しさ | 提案 |
|---|---|---|
| メールの差出人・送信済みフォルダ | GAS は実行者本人の Gmail から送る | 提案: 貯蔵品アプリと同じく、GAS の持ち主 1 人の OAuth 委任トークンで **Gmail API `users.messages.send`**。テキスト本文・UTF-8・件名は RFC 2047 でエンコード。`mail_log` に 1 通ずつ記録して送り直せるようにする。**ただし、まず「このアプリが今メールを送っているか」を確かめ、送っていなければ VPS でも送らない（統合メール通知システムが Requests を読んでいるなら、DB からの書き出し先を変えない）** |
| 統合メール通知システムとのつながり | リポジトリの外。Requests シートを直接読んでいる可能性（推測） | 提案: 何を読んでいるか（シート名・列）を調べ、VPS から本番シートへの書き出しで同じ列を保つ。【要確認】 |
| PDF の見た目 | Sheets の描画エンジンが作る PDF。HTML→Chromium では同じにならない | 提案 A（同じ見た目を優先）: 委任トークンで **Drive API でテンプレートを複製 → Sheets API `values.batchUpdate` で値を書く → 書き出し URL（`/export?format=pdf&gid=…` と同じ引数）を Bearer 付きで取る → Drive に保存 → 複製を削除**。印鑑画像は Sheets API に「セルの上に置く画像」を入れる命令が無いので、案 A-1 テンプレートに画像を置かず `=IMAGE()` のセルで出す（見た目が少し変わる）、案 A-2 小さな GAS（PDF だけを作る関数）を残して Apps Script API で呼ぶ。提案 B: HTML→Chromium で作り直す（見た目は変わる。オーナーの了承が要る） |
| XLOOKUP 方式 | テンプレート内の数式と、そこから参照するデータの置き場所が分からない | 提案: VPS では direct 方式 1 本にまとめる（Sebastian 経由も含めて）。ただし今の本番が XLOOKUP 方式なら、出力の差を見本で比べてオーナーに確かめる |
| 日付フォルダと Drive の持ち主 | `PDF_ROOT_FOLDER_ID/yyyy.MM.dd` を実行者の権限で作る | 提案: 委任トークン（同じ持ち主）で Drive API。フォルダは名前で探して最初の 1 つ、無ければ作る（今と同じ）。安全装置で書いてよいフォルダを本番ルートの中だけにする |
| 蓄積スプレッドシートの書式 | 行ごとの背景色・HYPERLINK 数式・見出しの書式 | 提案: Sheets API `batchUpdate`（`repeatCell` で背景、`USER_ENTERED` で HYPERLINK 数式と日付）。朝バッチの「薄い行で上書き」をそのまま再現するか、全列を埋める形に直すかは、オーナーに決めてもらう |
| Google フォーム | フォームの複製・選択肢の更新・回答の読み取りは FormApp 前提 | 提案: 回答が来ていないならフォームは「回答を受け付けない」にして廃止（VPS で再現しない）。来ているなら **Forms API**（`forms.responses.list` を `filter=timestamp >= …` で 1 分ごとに読み、回答 ID を DB に保存して重複を防ぐ）と、選択肢更新は `forms.batchUpdate`。フォームの新規作成は Drive API の複製＋`batchUpdate` |
| 事前入力 URL | `toPrefilledUrl()` は FormApp だけ | 提案: `entry.{questionId}=値` の URL を自分で組み立てる（質問 ID は Forms API で取れる）。画面から使われていないので、作らない選択も可 |
| XLSX 書き出し（未使用） | 一時スプレッドシート経由 | 提案: 使っていないので作らない。要るなら openpyxl |
| ロック | `LockService` のスクリプトロック | 提案: PostgreSQL の行ロック／advisory lock と、PDF は worker の outbox で 1 件ずつ |

---

## 7. 受け入れの確かめ項目（この領域）

- [ ] 本番のトリガー一覧を控えた（`sendEveningMail_`、`morningBatch_`、`morningBatchNoMail_`、`nightlyUpdateAllForms_`、`pollNewResponses_`、`syncAllMasters` の有無と時刻）
- [ ] 総務部に今届いている通知メールの送り元（このアプリか、統合メール通知システムか）を確かめた
- [ ] Settings の 5.1 のキーの値を控えた（`PDF_MODE` を含む）
- [ ] 夕方メール・朝メールを VPS で再現する場合、0 件・1 件以上・APP_URL なしの 3 通りで、GAS と文字単位で一致
- [ ] PDF: 残業・休日（半日／1 日）、印鑑あり／なし、一次のみ／二次あり、工番 0〜3 件、補足理由あり／なしで、GAS の PDF と並べて見比べた
- [ ] PDF のファイル名・日付フォルダ名・`pdfFileId`/`pdfFolderId`/`pdfGeneratedAt` が今と同じ形
- [ ] 二次承認・編集・削除で、古い PDF がゴミ箱へ行き、必要なら作り直される
- [ ] 蓄積 SS: シート名・27 列・書式・HYPERLINK が同じ。upsert／朝バッチ／削除の動きを決めたとおりに再現
- [ ] フォーム: 直近の回答数を見て、残すか閉じるかを決めた。閉じるなら全フォームを「回答を受け付けない」にした

## 確かめること

- 本番の `PDF_MODE`、テンプレート `TEMPLATE_SSID` の XLOOKUP の参照先
- 古いメールトリガーが残っているか、`gmail.send` だけのスコープで送れているか
- 統合メール通知システムがこのアプリのどのシート・列を読んでいるか
- 各部署フォームに最近の回答があるか、「業務ID」質問が残っているか
