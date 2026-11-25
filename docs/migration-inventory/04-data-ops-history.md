# 移行棚卸し 04：データ・マスタ・トリガー／運用・履歴

対象：残業・休日出勤申請（GAS、`scriptId=1GPNBR9U9mw7TZPnl1pxtdVjFagUPxt9S6qSYlbNIEuHNvRLVspNIQNc7`、`.clasp.json:2`）
目的：VPS（FastAPI＋PostgreSQL）へデータを欠けなく移し、GAS 版へいつでも戻せる状態を保つための、データ移行と運用切り替えのチェックリスト。
調査日：2026-11-26。リポジトリ HEAD = `ca2ecac`（2026-11-07）。Google 側（実シート・トリガー・デプロイ設定）は一切触っていません。

表記のきまり

- **確認済み**：コードや git の記録で確かめたこと。根拠を `ファイル:行` で付けています。
- **推測**：コードから読み取った見込みです。本番シートや Apps Script の画面で確かめる必要があります。
- 行番号は HEAD（`ca2ecac`）時点のものです。

---

## 0. 先に結論（移行担当が最初に読むところ）

- 本番データの置き場所は、スクリプトが紐づいた 1 冊のスプレッドシート（以下「DB」）です。`getDb_()` は `SpreadsheetApp.getActive()` を返します（確認済み、`src/config.gs:69-72`）。DB の ID はコードにありません（推測：Apps Script 画面の「コンテナ」から確かめる）。
- コードが読み書きするシートは DB 内に 15 種類、外部スプレッドシートは 3 冊（マスタ元・蓄積台帳・PDF テンプレ）です。ほかに Google フォーム（部署×種別）と Drive の PDF・印鑑画像があります。
- 列は「ヘッダ名」で引いています。列順は本番シートごとに違う可能性があり、後から自動追加された列（`reasonDetail`、`approvedBy2`、`approvedAt2`、`modifiedAt`、`modifiedBy`）は末尾にあります（確認済み、`src/formSubmit.gs:648-660`、`src/approval.gs:560-578`、`src/worklog.gs:593-608`）。移行スクリプトは **列番号ではなくヘッダ名で** 読んでください。
- 日時は「Date 型」と「JST 文字列」が混在しています。特に `WorkLogs` の日時列と `Requests.targetDate` は時期によって型が違います（第 4 章）。
- 時間で動くトリガーは最大 4 本（＋古いものが残っている可能性）です。1 分ごとのフォーム回答ポーリングがあり、Google フォームは今も回答を受け付けている可能性があります（第 3 章・第 5 章）。
- 外部から `doPost` を呼ぶ仕組み（Sebastian／Hermès Agent）があります。切り替え時に呼び先 URL とトークンの付け替えが要ります（第 6 章）。
- GitHub Actions の自動デプロイは 2026-02-20 から失敗し続けています（認証切れ `invalid_rapt`）。本番 GAS のコードがリポジトリと一致しているかは未確認です（第 7 章）。

---

## 1. DB（紐づきスプレッドシート）のシート一覧

共通の前提（確認済み）

- ヘッダ名は `normalize_()` で前後空白を取り、全角英数を半角にしてから照合します（`src/config.gs:99-103`）。
- `getSheetHeaderIndex_(sheetName, headerRowNo=1)` はヘッダ行を読み、`{ヘッダ名: 0 始まり列番号}` を作ります。結果は実行中だけ `_headerCache` に置きます（`src/config.gs:232-244`、`buildHeaderIndex_` は `src/formSubmit.gs:301-308`）。
- `requireSheet_(name)` はシートが無いと `Sheet not found` で止まります（`src/config.gs:92-97`）。
- `getSettings_()` は実行中だけ `_settingsCache` に置きます（`src/config.gs:74-90`）。

### 1-1. Requests（申請本体、1 行 = 1 申請）

- ヘッダ行：**1 行目**、データは 2 行目から（確認済み、`getSheetHeaderIndex_('Requests', 1)` を全箇所で使用。例 `src/worklog.gs:52`、`src/approval.gs:289`）。
- 書き込む関数（確認済み）
  - 新規行：`appendRequestRow_()`（`src/formSubmit.gs:598-672`）。呼び出し元は `api_submitRequest()`（画面のモーダル送信、`src/formSubmit.gs:678-812`）と `handleFormSubmit_()`（フォーム回答、`src/formSubmit.gs:420-593`）。
  - 承認：`api_approveRequest`（`src/approval.gs:280-344`）、一括 `api_approveRequestsBatch`（`src/approval.gs:388-`）。
  - 却下：`api_rejectRequest`（`src/approval.gs:348-386`）。status を `canceled`、approvedBy を `メール(却下)` にします。
  - 二次承認：`api_secondaryApprove`／`api_secondaryApproveBatch`（`src/approval.gs:696-`、`775-`）。
  - 編集：`api_editRequest`（`src/worklog.gs:540-647`）。削除：`api_deleteRequest`（`src/worklog.gs:652-700`）。
  - PDF 列：`generatePdfForRequest_`（`src/pdfExport.gs:295-299`）、`generatePdfDirect_`（`src/pdfExport.gs:423-427`）、クリアは `clearExistingPdfForRegeneration_`（`src/approval.gs:873-883`）。
  - 補足理由の後埋め：`tryFillReasonDetailFromForm_`（`src/pdfExport.gs:994-`、`forceRegeneratePdf` から呼ぶ、`src/pdfExport.gs:835`）。
  - 列の自動追加：`appendRequestRow_`（不足ヘッダを末尾に追加、`src/formSubmit.gs:648-660`）、`ensureSecondaryApprovalColumns_`（`src/approval.gs:560-578`）、`api_editRequest`（`modifiedAt`/`modifiedBy`、`src/worklog.gs:596-608`）。

列（ヘッダ名はコードが使う正式名。確認済み、`src/formSubmit.gs:605-646`）

| ヘッダ名（完全一致） | 型・値の形 | 書くところ／備考 |
|---|---|---|
| `requestId` | UUID 文字列（`Utilities.getUuid()`、例 `xxxxxxxx-xxxx-4xxx-…` 36 字） | 新規時（`src/formSubmit.gs:535`、`758`） |
| `requestType(overtime/holiday)` | `overtime` / `holiday` | 新規時。フォーム経由は「残業」以外をすべて `holiday` にします（`src/formSubmit.gs:455`） |
| `status(submitted/approved/canceled)` | `submitted` / `approved` / `canceled`（＋古い `modified` が残る可能性、第 4 章） | `canceled` は「本人削除」と「承認者の却下」の両方です。見分けは `approvedBy` 末尾の `(却下)` だけです（`src/approval.gs:376-377`、`src/worklog.gs:666`） |
| `dept` | 部署名（作業員マスタの「部署」と一致） | |
| `workerCode` | 作業員コード（ラベル「コード␣氏名」の先頭、`src/formSubmit.gs:392-399`） | |
| `workerName` | 氏名（ラベルの 2 語目以降） | |
| `workerEmail` | 作業員マスタ「Googleアカウント…」列から補完。空のことが多い（推測） | `lookupWorkerEmail_`（`src/masters.gs:180-197`） |
| `targetDate` | 書き込み値は `'yyyy-MM-dd'` 文字列（`src/formSubmit.gs:469`、`711`、`src/worklog.gs:586`）。`appendRow`/`setValue` で書くため、シート側で日付に自動変換されて **Date 型で読める行と文字列の行が混在** する見込み（推測）。読む側は両対応（`src/mail.gs:17-23`、`src/worklog.gs:369-372`） | 編集で休日のみ変更可 |
| `submittedAt` | Date（`new Date()`） | |
| `approvedAt` | Date（承認・却下時刻）または空 | 却下でも入ります |
| `approvedBy` | 承認者メール／`メール(却下)`／`unknown`（メールが取れない時、`src/approval.gs:281`） | |
| `approvedMinutes` | 数値（分）。残業 30〜240（0.5h 刻み、`src/config.gs:51`、`src/formSubmit.gs:403-408`）、休日 240（半日）/480（1日）（`src/formSubmit.gs:410-415`） | 「承認された時間」ではなく **申請時の予定時間** です（名前に注意） |
| `reason` | 理由の固定文言（`src/config.gs:42-49`） | |
| `reasonDetail` | 補足理由（その他の時は必須） | 後から自動追加された列（`c45ac6c`） |
| `workContent` | 業務内容（任意） | |
| `jobId1` | 常に空文字（業務NO 連携は廃止、`3d6214b`） | ドキュメントにある `jobId2`/`jobId3` はコードが書きません |
| `workNo1`〜`workNo3` | 工番コード（任意、最大 3 件） | |
| `orderNo1`〜`orderNo3` | **workNo と同じ値**（`src/formSubmit.gs:365-368`） | 文書 `docs/spreadsheet-structure.md:192-194` の「受注先」という説明はコードと違います |
| `customer1`〜`customer3` | 工番マスタの **納入先**（無ければ受注先）（`src/formSubmit.gs:376`） | 文書の「納入先」と一致 |
| `product1`〜`product3` | 工番マスタの品名 | |
| `approvedBy2` | 二次承認者メール／空 | 後から自動追加 |
| `approvedAt2` | Date／空 | 後から自動追加 |
| `hrMailSentAt` | 常に空（書くコードなし） | 旧仕様の名残 |
| `pdfGeneratedAt` | Date／空 | |
| `pdfFileId` | Drive ファイル ID／空 | PDF 実体は Drive（第 2-4 節） |
| `pdfFolderId` | Drive フォルダ ID（`PDF_ROOT_FOLDER_ID` 配下の `yyyy.MM.dd`、`src/pdfExport.gs:150-158`） | |
| `exportError` | `工番マスタ未登録: X / …` などの注意文 | 「エラー」ですが申請は有効です |
| `modifiedAt` | `'yyyy-MM-dd HH:mm:ss'` 文字列で書く（自動で Date 化されうる、推測） | 編集時だけ。列は初回編集時に追加 |
| `modifiedBy` | 編集者メール／`unknown` | 同上 |

注意（確認済み）：`api_editRequest` は `modifiedAt` がある・`modifiedBy` が無い場合に `modifiedBy` の位置を `modifiedAt + 1` と決め打ちします（`src/worklog.gs:604-607`）。その列に別のヘッダがあると上書きされます。実シートで `modifiedAt` の右隣が `modifiedBy` か確かめてください。

### 1-2. WorkLogs（実績、1 行 = 1 申請）

- ヘッダ行：**2 行目**（1 行目は注記）。データは 3 行目から（確認済み、`src/worklog.gs:9`、`14`、`src/formSubmit.gs:820-822`）。
- 書く関数（確認済み）：`ensureWorkLogRow_`（プレースホルダ、`src/formSubmit.gs:816-837`）、`updateWorkLog_`（`src/worklog.gs:21-47`）。呼び出し元は `api_markOvertimeDone`／`api_markHolidayStart`／`api_markHolidayDone`（`src/worklog.gs:143-342`）、`api_adminManualComplete`（`src/adminApi.gs:432-`）、Sebastian 用 3 関数（`src/api.gs:73-257`）、削除時のクリア（`src/worklog.gs:675-685`）、修復 `recalcWorkLogsBreakMinutes`（`src/worklog.gs:712-806`）。

| ヘッダ名 | 型・値の形 |
|---|---|
| `requestId` | Requests.requestId |
| `actualStartAt` | 2026-02-27 以降は **テキスト書式の JST 文字列** `'yyyy-MM-dd HH:mm:ss'`（`setNumberFormat('@')`、`src/worklog.gs:35-45`）。それより前の行は Date 型の可能性（第 4 章）。残業は常に「targetDate 17:20」 |
| `actualEndAt` | 同上 |
| `actualMinutes` | 数値（分）＝終了−開始。削除時は空文字 |
| `breakMinutes` | 数値（分）。時刻ベース計算 `calcBreakByClockTime_`（`src/worklog.gs:111-129`、休憩窓は `src/config.gs:55-62`） |
| `netMinutes` | 数値（分）＝actual−break（0 未満にしない） |
| `updatedAt` | 多くは JST 文字列。ただし `ensureWorkLogRow_` は Date を書くので混在（確認済み、`src/formSubmit.gs:833`） |
| `updatedBy` | メール／`formSubmit`／`sebastian-agent`／`admin-manual`／`unknown`／`delete`／`break-recalc`（`src/formSubmit.gs:834`、`src/api.gs:99`、`src/adminApi.gs:486`、`src/worklog.gs:683`、`788`） |

- 1 行目の注記文（文書）：「※実績。残業は actualEndAt のみ入力。（開始は 17:20 固定）」（`docs/spreadsheet-structure.md:228`）。
- 実績は承認前でも記録できます（`128643b`）。PDF は「承認済み かつ 完了済み」になった時点で作ります。

### 1-3. Settings（Key / Value）

- 1 行目：説明、2 行目：ヘッダ（key, value, memo）、3 行目から key/value（確認済み、`src/config.gs:81-87`）。A 列 = key、B 列 = value。
- `getAccumulationSS_` は蓄積台帳を新しく作ったとき `['ACCUMULATION_SS_ID', id]` を末尾に追記します（確認済み、`src/mail.gs:415-416`）。

コードが読むキー（すべて確認済み）

| キー | 用途 | 参照箇所 |
|---|---|---|
| `SOURCE_SSID` | マスタ元 SS（作業日報_全従業員用）の ID | `src/syncMasters.gs:34` |
| `PDF_ROOT_FOLDER_ID` | PDF 保存先の親フォルダ | `src/pdfExport.gs:173`、`381` |
| `TEMPLATE_SSID` | PDF テンプレ SS | `src/pdfExport.gs:174`、`382` |
| `PDF_MODE` | `direct` を含めば直接書き込み方式、それ以外は XLOOKUP 方式 | `src/pdfExport.gs:577`、`src/worklog.gs:197` ほか |
| `ADMIN_EMAILS` | 総務（管理者）メール、カンマ区切り | `src/approval.gs:19`、`src/deptAuth.gs:97` |
| `GENERAL_AFFAIRS_EMAILS` | 同上（優先） | `src/approval.gs:20`、`src/deptAuth.gs:92` |
| `ACCUMULATION_SS_ID` | 蓄積台帳 SS の ID | `src/mail.gs:392` |
| `HR_MAIL_TO` | 夕方・朝メールの宛先（今は手動実行時のみ） | `src/mail.gs:328`、`1318` |
| `APP_URL` | メール本文に入れるアプリ URL | `src/mail.gs:238`、`1326` |
| `PDF_DRIVE_FOLDER_URL` | 朝メール本文の PDF フォルダリンク（既定値は `1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B`） | `src/mail.gs:1327` |
| `FORMS_PARENT_FOLDER_ID` | 部署別フォームの複製先フォルダ | `src/formGenerate.gs:32` |
| `OVERTIME_TEMPLATE_FORM_ID` / `HOLIDAY_TEMPLATE_FORM_ID` | テンプレフォーム ID（無ければ FormTemplates シート） | `src/formMap.gs:67-69` |
| `STAMP_OVERTIME_FILE_ID` / `STAMP_HOLIDAY_FILE_ID` | PDF 右上の区分印（画像の Drive ID） | `src/pdfExport.gs:234` |

Settings 以外の保存場所（確認済み）

- スクリプト プロパティ：`SEBASTIAN_API_TOKEN`（`src/api.gs:18`）、`POLL_LAST_TS`（ISO 文字列、`src/formSubmit.gs:23`、`91`）。**移行時に値を控える**こと（Apps Script 画面の「プロジェクトの設定」）。
- スクリプト キャッシュ：`POLL_PROCESSED_IDS`（10 分で消える、`src/formSubmit.gs:49`、`105`）。移行不要。

### 1-4. 作業員マスタ（同期で毎朝全置換）

- ヘッダ行：**1 行目**（確認済み。`values[0]` をヘッダとして読む `src/masters.gs:12-14`。同期は元シートを 1 行目からそのまま写す `src/syncMasters.gs:117-153`、`219-228`）。`docs/spreadsheet-structure.md:79`「Row 2」は古い記述です。
- コードが使う列（確認済み、`src/masters.gs:14-20`、`147-154`）：`作業員コード`、`氏名`、`部署`、`在籍フラグ`（任意。空・`1`・`true` 以外は除外）、`Googleアカウント…`（前方一致、任意）。ほかの列（例：`担当業務ID（カンマ区切り）`）も元のまま写ります。
- **上限 81 行（ヘッダ＋80 人）で切り捨て**（確認済み、`src/syncMasters.gs:12`、`61-64`）。81 人目以降は DB に入りません。
- 部署一覧はこのシートの「部署」列の重複なしで作ります（部署マスタは 2026-02-16 に廃止、`88e58bb`、`src/masters.gs:3-8`）。

### 1-5. 業務NOマスタ（同期のみ、アプリは使わない）

- 元シート名は **`業務NO.マスタ`**（ドットあり）→ DB では `業務NOマスタ`（確認済み、`src/syncMasters.gs:13`、`src/config.gs:10`）。
- アプリでは使いません。同期は「日報アプリへの影響を考えて意図的に残す」とされています（確認済み、`3d6214b` のコミット文）。推測：別のアプリがこの DB の `業務NOマスタ` を読んでいる可能性があります。切り替え後も DB と同期を止めるかどうか、別途確かめてください。

### 1-6. 工番マスタ（同期で毎朝全置換）

- ヘッダ行：1 行目（確認済み、`src/masters.gs:51-52`）。列：`工番`、`受注先`、`納入先`、`納入先住所`、`品名`、`数量`、`取込日時`（`src/masters.gs:50`）。
- `取込日時` は同期時に `'yyyy-MM-dd HH:mm:ss'` 文字列へ変換されます（2026-06-23 以降、`src/syncMasters.gs:204-214`）。
- 申請時に工番から納入先・品名を引いて Requests に写します（`src/formSubmit.gs:320-388`）。申請後に工番マスタが変わっても Requests の値は変わりません。

### 1-7. 休憩マスタ（ほぼ未使用）

- ヘッダ行：1 行目（確認済み、`src/worklog.gs:88-95`）。列：`適用区分(overtime/holiday)`、`実働分_下限`、`実働分_上限`（空＝上限なし）、`休憩分`。
- 使うのは **Sebastian 経由の残業完了だけ** です（確認済み、`src/api.gs:89`）。画面からの完了は時刻ベース（`calcBreakByClockTime_`）で、2026-02-28 に切り替え済みです（`4a84703`）。このため Sebastian 経由と画面経由で休憩の計算方法が違います（第 4 章）。

### 1-8. ApproverMap（旧方式の承認者表）

- ヘッダ行：1 行目（確認済み、`src/approval.gs:37-47`、`src/deptAuth.gs:33-40`）。文書の「Row 2」とは違います。
- 列：`部署`、`承認者メール`、`role(approver/admin)`、`有効フラグ`（`false` で無効）。
- 用途：`role=admin` の判定（`src/approval.gs:34-62`）。部署の承認者判定は DeptApprovers が 2 行以上あればそちらを使い、無いときだけ ApproverMap に戻ります（`src/deptAuth.gs:14-55`）。

### 1-9. DeptApprovers（現行の承認者表）

- ヘッダ行：1 行目、データは 2 行目から（確認済み、`src/deptAuth.gs:15-27`）。
- 列は位置で読みます：A = 部署、B = 承認者メール（カンマ区切り可）、C = 氏名（任意、`src/deptAuth.gs:66-70`）。メールは小文字化して比べます。

### 1-10. StampMap（承認者の印鑑画像）

- ヘッダ行：1 行目。列：`メール`、`stampFileId`（大文字小文字は無視）、`備考`（確認済み、`src/pdfExport.gs:99-128`）。
- 値は Drive の画像ファイル ID です。PDF の承認欄に貼ります（`src/pdfExport.gs:253-271`）。マスタ同期で消えないよう専用シートにした経緯があります（`73dc59e`）。**画像の実体（Drive）も移行対象** です。

### 1-11. BatchLogs（朝バッチの記録）

- 無ければ作り、ヘッダを 1 行目に書きます：`実行日時`、`バッチ名`、`対象日`、`成功`、`スキップ`、`失敗`、`エラー詳細`（確認済み、`src/pdfExport.gs:604-627`）。
- `実行日時` は Date、`対象日` は `'yyyy-MM-dd'` 文字列（自動で Date 化されうる、推測）。`バッチ名` は `morningBatch`／`morningBatchNoMail`／`manualBatchPdfs`。
- 行数のめやす（推測）：1 日 1 行 × 運用日数（2026-02 から）＝数百行。

### 1-12. FormMap（部署×種別の Google フォーム）

- ヘッダ行：1 行目。列：`type`、`dept`、`formId`、`formUrl`、`updatedAt`（Date）、`isActive`（true/false）（確認済み、`src/formMap.gs:3-12`、`47-60`）。
- 行数のめやす（推測）：部署数 × 2。

### 1-13. FormTemplates（テンプレフォーム、旧方式）

- ヘッダ行：1 行目。列：`type`、`templateFormId`（コードが使う）、`templateUrl`、`note`（確認済み、`src/formMap.gs:85-98`、`docs/spreadsheet-structure.md:255-260`）。Settings にテンプレ ID があればそちらが優先です。

### 1-14. 社内カレンダーマスタ（同期で毎朝全置換）

- ヘッダ行：1 行目。列：`日付`、`区分`（`休日`／`祝日`／`出勤土曜`）、`曜日`、`備考`（確認済み、`src/menu.gs:236`、読むのは A〜D 列固定 `src/formMap.gs:245-251`）。
- 正本はマスタ元 SS の同名シートです（`src/menu.gs:224-227`）。`ensureCompanyCalendarSheet` は器（ヘッダと条件付き書式）だけを用意し、データは消しません（`src/menu.gs:228-267`）。
- `日付` は 2026-06-23 以降 `'yyyy-MM-dd'` 文字列で書きます（`src/syncMasters.gs:204-214`）。
- 2026 年度の行数：124 行（休日 93、祝日 19、出勤土曜 12。`docs/社内カレンダーマスタ_2026年度_新.tsv` を数えて確認済み。TSV はヘッダ込み 125 行、タブ区切り、同じ 4 列）。
- 用途：休日出勤の候補日（今週月曜〜来週金曜）。データが無いと旧方式（Google カレンダーの祝日）に戻りますが、`appsscript.json` にカレンダーの権限が無いため旧方式は空振りし、法定祝日の計算に落ちます（推測、`src/formMap.gs:310-378`、`src/appsscript.json:4-12`）。

### 1-15. 文書にだけ出てくるシート（コードは読み書きしない）

| シート | 状況 |
|---|---|
| `操作` | DB 側にあると文書に記載（`docs/spreadsheet-structure.md:28-30`）。コードが読むのは **テンプレ SS 側** の `操作!B3`（`src/pdfExport.gs:188-190`） |
| `README` | 説明用（文書のみ） |
| `部署マスタ` | 2026-02-16 に廃止（`88e58bb`）。残っていても未使用 |
| `ExportHistory` | 文書に定義あり（`docs/spreadsheet-structure.md:232-247`）。**書くコードはありません**（確認済み、grep で該当なし）。空か古いテストの可能性 |
| `申請書フォーム` / `休憩計算_フォーム` | 文書では DB 内のシート（`docs/spreadsheet-structure.md:292-319`）。コードが使う `申請書フォーム` はテンプレ SS 側です。推測：`TEMPLATE_SSID` が DB 自身、または DB の複製を指している可能性があります（XLOOKUP 方式は `操作!B3` の requestId で同じ SS 内の Requests を引く作りのため）。その場合、PDF を作るたびに **DB 全体の一時コピー**（`TMP_<requestId>_yyyyMMdd_HHmmss`）が作られてゴミ箱に入っています（`src/pdfExport.gs:183-185`、`302`） |
| フォームの回答シート | コードに `setDestination` はありません（確認済み）。回答はフォーム側に溜まり、`pollNewResponses_` が Requests に写します。DB に「フォームの回答 N」シートがあるかは要確認 |

---

## 2. 外部スプレッドシート・Drive

### 2-1. マスタ元：`SOURCE_SSID`（作業日報_全従業員用、「LW／作業日報_全従業員用」）

- **読むだけ** です（確認済み。`openById` → `getValues` のみ、`src/syncMasters.gs:43-55`、`readSheetData_` `115-153`）。ただし診断関数 `testSyncWorkerOneRow` も元は読むだけで、書くのは DB 側です（`src/syncMasters.gs:237-262`）。
- 対応表（確認済み、`src/syncMasters.gs:11-18`）

| 元シート名 | DB 側シート名 | 制限 |
|---|---|---|
| `作業員マスタ` | `作業員マスタ` | ヘッダ＋80 行まで |
| `業務NO.マスタ` | `業務NOマスタ` | なし |
| `工番マスタ` | `工番マスタ` | なし |
| `社内カレンダーマスタ` | `社内カレンダーマスタ` | なし |

- 処理の流れ（確認済み、`src/syncMasters.gs:28-94`）
  1. スクリプトロックを 30 秒待つ。
  2. 最終列は「ヘッダ行の最後の空でない列」で決める（書式だけ残った列を無視、`detectLastColByHeader_` `100-108`）。
  3. 読み取りは 3 段構え（通常 → 2 秒後に再試行 → 50 行ずつ）（`115-153`）。
  4. 1 列目が空の行を捨てる（ヘッダは残す、`219-228`）。
  5. 上限行で切る。
  6. Date セルを JST 文字列にする（0 時ちょうどは `yyyy-MM-dd`、それ以外は `yyyy-MM-dd HH:mm:ss`、`204-214`）。
  7. DB 側を `clearContents` して値だけ貼る。余分な行（5 行超）・列（2 列超）は削除（`161-187`）。失敗したら **シートを削除して作り直す**（`70-79`）。この場合、条件付き書式や保護が消えます。
- 移行で気をつけること：VPS ではこの同期をどこで動かすか（マスタ元は Google に残る見込み）を決める必要があります。DB 側のマスタは毎朝上書きされるので、移行データとしては「その時点の写し」で足ります。

### 2-2. 蓄積台帳：`ACCUMULATION_SS_ID`（「残業・休日出勤 データ蓄積（総務確認用）」）

- **書き込みます**（確認済み）。無い・開けないときは新しく作り、フォルダ `1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B` に入れて Settings に ID を追記します（`src/mail.gs:390-426`）。開けないだけで新規作成するため、権限の問題で **台帳が二重にできる** おそれがあります。
- シート：`残業_YYYY年度`／`休日出勤_YYYY年度`（年度は 3/16 始まり、`src/mail.gs:465-472`）。修復時のバックアップ `…_backup_yyyyMMdd_HHmmss` も残ります（`src/mail.gs:958-972`）。
- ヘッダ（1 行目、27 列、確認済み `src/mail.gs:343-351`）：`管理番号`（=requestId）、`申請種別`、`作業実施日`（`yyyy/MM/dd`）、`曜日`、`部署`、`作業員ID`、`氏名`、`理由`、`補足理由`、`業務内容`、`承認時間(分)`、`承認時間`、`開始時刻`、`終了時刻`、`実働(分)`、`実働`、`休憩(分)`、`休憩`、`実残業/実働(分)`、`実残業/実働`、`申請日`、`1次承認日`、`1次承認者`、`2次承認日`、`2次承認者`、`承認状態`（`承認済`／`申請中`／`取消`）、`PDF`（`=HYPERLINK(...,"PDF")` 数式）。
- 書くタイミング（確認済み）：作業完了（`src/worklog.gs:207`、`336`）、承認時に完了済みなら（`src/approval.gs:337`、`465`）、二次承認（`src/approval.gs:760`、`863`）、編集で PDF 再生成時（`src/worklog.gs:630`）、管理者の手入力完了（`src/adminApi.gs:512`）、Sebastian（`src/api.gs:113`、`239`）、朝バッチ（前日分を日単位で消して書き直し、`src/mail.gs:489-517`、`623-648`）。削除は本人削除のときだけ（`src/worklog.gs:691`）。
- 中身は Requests＋WorkLogs から作り直せる **派生データ** です（`rebuildAccumulationSS` が実際に作り直す、`src/mail.gs:943-1127`）。ただし総務が人の目で使う台帳なので、手で書き足した内容が無いかは要確認です。

### 2-3. PDF テンプレ：`TEMPLATE_SSID`

- 本体は **読むだけ**（`makeCopy` で一時コピーを作り、コピーに書く。確認済み、`src/pdfExport.gs:182-185`、`389-392`）。診断 `debugPdfReason` は本体を開いて読むだけです（`src/pdfExport.gs:699-720`）。
- 使うシート：`操作`（B3 に requestId、XLOOKUP 方式のみ）、`申請書フォーム`（セル位置は `PDF_MAP`、`src/pdfExport.gs:3-23`）。
- 一時コピーはゴミ箱へ移すだけで、完全削除はしていません（`src/pdfExport.gs:302`、`430`）。

### 2-4. Drive 上の実体（移行・参照の対象）

- PDF：`PDF_ROOT_FOLDER_ID/yyyy.MM.dd/` に `yyyyMMdd_部署_氏名_残業|休日出勤.pdf`（確認済み、`src/pdfExport.gs:276-288`）。Requests の `pdfFileId` が指します。編集・二次承認・削除で古い PDF はゴミ箱に入ります（`src/approval.gs:873-878`）。
- 印鑑画像：StampMap の `stampFileId`、Settings の `STAMP_*_FILE_ID`。
- 部署別 Google フォーム：FormMap の `formId`。テンプレフォーム 2 つ。
- フォルダ `1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B`：蓄積台帳の置き場兼 PDF フォルダの既定リンク（`src/mail.gs:393`、`1327`）。

---

## 3. トリガー・メニュー・手動実行の関数

### 3-1. コードが作る時間トリガー

トリガーの持ち主は **セットアップを実行したアカウント** です。そのアカウントの権限で動き、そのアカウントの実行枠を使います（Apps Script の仕様）。どのアカウントで作られているかは Apps Script 画面の「トリガー」で確かめてください（推測：今泉オーナーのアカウント）。

| 関数 | 種類・時刻 | 作る関数 | 中身 |
|---|---|---|---|
| `syncAllMasters` | 毎日 6 時台（nearMinute 0） | `setupSyncTrigger_`（`src/menu.gs:150-161`） | マスタ 4 種の同期 |
| `nightlyUpdateAllForms_` | 毎日 6 時台（nearMinute 30） | `setupTriggers_`（`src/menu.gs:86-99`） | 全フォームの選択肢更新（`src/formBatch.gs:3-41`） |
| `morningBatchNoMail_` | 毎日 6 時台（nearMinute 50） | `setupMorningBatchTrigger_`（`src/menu.gs:212-221`） | 前日分の PDF 一括作成＋BatchLogs＋蓄積台帳（`src/mail.gs:1454-1468`） |
| `pollNewResponses_` | 1 分ごと | `setupFormPollTrigger_`（`src/menu.gs:182-196`） | 全フォームの新しい回答を Requests へ（`src/formSubmit.gs:14-114`） |

現行の一括セットアップは `setupAllTriggers_`（`src/menu.gs:200-209`）で、上の 4 本を作り、古いフォーム送信トリガーを消します。

**古いトリガーが残っている可能性**（推測、要確認）

- `sendEveningMail_` ×2（16 時台・18 時台）と `morningBatch_`（7 時台）は `setupMailTriggers_` が作ります（`src/menu.gs:103-146`）。2026-03-11 にメールを「統合メール通知システム」へ移したとき、一括セットアップから外しただけで **既存のトリガーを消す処理はありません**（`fcf8e31`、`src/menu.gs:204-205`）。残っていれば、夕方メールの二重送信や、`morningBatch_` と `morningBatchNoMail_` の二重実行が起きています。
- `handleFormSubmit_` のフォーム送信トリガー（2026-02-19 以前の方式）。`cleanupFormSubmitTriggers_` で消す作りです（`src/menu.gs:167-178`）。
- トリガーは 1 プロジェクト 20 本までです（`15697d5` の経緯）。

### 3-2. スプレッドシートのメニュー「残業休日アプリ」（`onOpen`、`src/menu.gs:3-30`）

| 表示名 | 関数 | 中身 |
|---|---|---|
| フォームを全更新（マスタ反映） | `nightlyUpdateAllForms_` | フォームの選択肢を作り直す |
| フォーム生成テスト（残業×先頭部署） | `debugCreateFirstDeptOvertime_` | 先頭部署の残業フォームを作る／返す |
| 全部署フォーム一括生成 | `buildAllDeptForms` | 全部署×2 のフォームを作る（`src/menu.gs:57-82`） |
| 夕方メール送信（手動） | `sendEveningMail_` | `HR_MAIL_TO` へメール（`src/mail.gs:326-336`） |
| 朝バッチ（PDF生成+朝メール） | `morningBatch_` | PDF＋BatchLogs＋蓄積台帳（メールはコメントアウト済み、`src/mail.gs:1427-1448`） |
| 朝メールのみ（手動） | `sendMorningMail_` | 前日分の実績メール |
| PDF一括生成（本日分） | `manualBatchPdfs_` | 当日分の PDF（`src/menu.gs:33-43`） |
| マスタ転記（手動実行） | `syncAllMasters` | マスタ同期 |
| 全トリガー初期セットアップ | `setupAllTriggers_` | 3-1 の 4 本 |
| 旧フォームトリガー削除 | `cleanupFormSubmitTriggers_` | |
| 診断：ポーリング状態確認 | `debugPollStatus_` | ログのみ（`src/formSubmit.gs:120-193`） |
| 診断：最新回答を手動処理 | `debugProcessLatest_` | 1 件処理して `POLL_LAST_TS` を今に進める（`src/formSubmit.gs:198-238`） |
| 全回答を一括処理（初回/リカバリ用） | `reprocessAllResponses_` | 48 時間分を **重複確認なしで** 再処理（`src/formSubmit.gs:245-297`） |
| 社内カレンダーマスタ シート初期化（書式のみ） | `ensureCompanyCalendarSheet` | 器だけ用意 |
| 蓄積SS：時間表記列を追加（マイグレーション） | `migrateAccumulationTimeColumns` | 台帳に「X時間Y分」列を差し込む（`src/mail.gs:806-927`） |
| WorkLogs：休憩分を再計算（1分バグ修復） | `recalcWorkLogsBreakMinutes` | 全行の break/net を再計算（`src/worklog.gs:712-806`） |
| 蓄積SS：完全再構築（バックアップ付き） | `rebuildAccumulationSS` | 台帳を作り直す（`src/mail.gs:943-1127`） |

### 3-3. エディタから手で動かす前提の関数（メニュー外）

- `testSyncWorkerOneRow`（同期の切り分け、DB の作業員マスタを 2 行だけにする。**本番で動かすと作業員マスタが消える**、`src/syncMasters.gs:237-262`）
- `debugPdfReason`（`src/pdfExport.gs:631-`）、`forceRegeneratePdf`（最新の承認済み 1 件の PDF を作り直し、Requests を書き換える、`src/pdfExport.gs:743-871`）、`debugStampData`（`src/pdfExport.gs:875-968`）
- `setupTriggers_`／`setupMailTriggers_`／`setupSyncTrigger_`／`setupFormPollTrigger_`／`setupMorningBatchTrigger_`（個別のトリガー作成）

---

## 4. 移行で気をつけるデータの癖（移行スクリプトの必須対応）

### 4-1. 日付・日時の型ずれ（最重要）

| 対象 | 状況 | 移行での扱い |
|---|---|---|
| `Requests.targetDate` | 書く値は `'yyyy-MM-dd'`。シートが自動で日付にした行は Date で返る（推測）。2026-02-27 20:22 より前のフォーム経由の行は Date をそのまま書いており、**休日出勤で 1 日ずれた可能性**（確認済み、`72b8eaf`） | Date なら DB の TZ で日付を取り出す。2026-02-27 以前の休日出勤は、作業者・PDF 名（`yyyyMMdd_…`）と照らし合わせる |
| `WorkLogs.actualStartAt/EndAt` | 2026-02-19 13:07 UTC より前は Date で書いており、**9 時間ずれ**の例あり（`3c55164`）。2026-02-27 15:09 より前は文字列で書いても自動で Date 化（`74b87b4`）。以降はテキスト書式の JST 文字列 | 文字列はそのまま JST。Date は「DB の TZ で書式化した文字列を JST とみなす」（コードと同じ扱い、`src/worklog.gs:273-283`、`src/mail.gs:101-130`） |
| `WorkLogs.updatedAt`、`Requests.modifiedAt`、`BatchLogs.対象日` | Date と文字列が混在 | 両方を受ける |
| `submittedAt`／`approvedAt`／`approvedAt2`／`pdfGeneratedAt`／`FormMap.updatedAt` | Date | タイムスタンプとして移す |
| マスタの日付（カレンダー・工番取込日時） | 2026-06-23 より前の同期では、元と DB のファイル TZ が違うと 1 日ずれた（`26f81a5`）。今は毎朝上書きされるので解消済みの見込み | 移行時点の値をそのまま使う |

- DB スプレッドシートのタイムゾーン設定は未確認です。コードは「DB の TZ が東京でない」ことを前提に回避策を重ねています（`getSafeTimeZone_`、`src/mail.gs:137-153`）。**エクスポート前に「ファイル → 設定 → タイムゾーン」を控えてください。**
- 安全なやり方：Sheets API で `UNFORMATTED_VALUE`（シリアル値）と `FORMATTED_VALUE`（表示文字列）を両方取って突き合わせる。

### 4-2. 値の意味のずれ・名残

- `status` に `modified` が残っている可能性（推測）：2026-03-28 16:44〜17:12 の約 30 分だけ「編集すると `modified`」の仕様でした（`0230f7f` → `5aa4621`）。承認者画面は今も `modified` を未承認として数えます（`src/approver.html:775`）。
- `canceled` は「本人削除」と「却下」の両方です。却下は `approvedBy` が `…(却下)` で終わります（`src/approval.gs:377`）。新システムで区別するなら移行時に分けてください。
- 却下した申請は蓄積台帳から消しません（`api_rejectRequest` は `removeAccumulationRow_` を呼ばない、確認済み）。完了済みの申請を却下すると、台帳に `申請中` のまま残る可能性があります（推測）。
- 削除した申請は WorkLogs の行を残し、中身だけ空にします（`src/worklog.gs:675-685`）。
- `approvedMinutes` は「予定時間」です。承認で変わる値ではありません。
- `orderNo1-3` は工番の写しで、受注先ではありません（1-1 節）。
- `jobId1`、`hrMailSentAt` は常に空、`jobId2/3` はコードが書かない列です。
- `approvedBy`／`updatedBy` が `unknown` の行（推測）：2026-04-28 に Web アプリ設定を `USER_DEPLOYING`＋`ANYONE_ANONYMOUS` に変えています（`953cb72`）。この設定のデプロイでは `Session.getActiveUser().getEmail()` が空になりやすく、承認者や本人確認が記録されない可能性があります。実データで `unknown` の件数を数えてください。

### 4-3. 計算の不一致・誤データ

- 休憩 1〜3 分の誤データ：休憩窓に少しだけ重なった完了で「休憩 1 分」などが 14 件出ていました（確認済み、`4ce6f25` のコミット文）。修復関数 `recalcWorkLogsBreakMinutes` を本番で実行したかは不明です。移行時に `breakMinutes ∈ {1,2,…,9}` の行を洗い出し、開始・終了から計算し直して突き合わせてください。
- Sebastian 経由の残業完了は休憩マスタ方式（`src/api.gs:89`）、画面からは時刻方式です。`updatedBy='sebastian-agent'` の残業行は計算方式が違います。また Sebastian 経由は `PDF_MODE` を見ずに XLOOKUP 方式で PDF を作ります（`src/api.gs:105`、`232`）。
- PDF 側は `netMinutes` が残業 720 分・休日 960 分を超えると開始・終了から計算し直します（`src/pdfExport.gs:40-65`）。シートには 1440 分などの異常値が残っている可能性があります（`0e94991`）。

### 4-4. 重複・テストデータ

- フォーム回答の重複（推測）：`reprocessAllResponses_` は重複確認をしません（`src/formSubmit.gs:241-243` のコメント）。`pollNewResponses_` の重複防止は 10 分・100 件のキャッシュだけで、エラーが続くと同じ回答をもう一度処理しうる作りです（`src/formSubmit.gs:88-106`）。`workerCode＋targetDate＋requestType＋submittedAt（数分以内）` で重複を探してください。
- `requestId` の重複は作りの上では起きませんが、移行前に一意性を確かめてください。WorkLogs も `requestId` ごとに 1 行の前提です。
- テストデータ（推測）：2026-02-13〜02-20 の開発期間（作者 `Claude` のコミットが集中）に作られた試験申請が残っている可能性があります。`dept`・`workerName`・`submittedAt` で洗い出してください。
- 蓄積台帳の列順ずれ：`migrateAccumulationTimeColumns` の列挿入バグ（`77ef76f`、`4ce6f25` で修正）で、古い年度シートの列がずれていた時期があります。`rebuildAccumulationSS` を実行していれば `_backup_` シートが残っています。台帳は Requests＋WorkLogs から作り直せるため、移行では **台帳を正本にしない** でください。

### 4-5. 構造の癖

- Requests の列順は固定ではありません（自動追加列が末尾）。ヘッダ名で読むこと。
- ヘッダ行の位置がシートで違います：WorkLogs は 2 行目、Settings はデータが 3 行目から、ほかは 1 行目。
- 作業員マスタは 80 人で切れます。実人数が 80 人を超えていないか、マスタ元で確かめてください。
- `getRequestFullData_` が 2 か所で定義されています（`src/mail.gs:746` と `src/pdfExport.gs:342`）。GAS では後から読まれた方が有効になり、ファイル順から `pdfExport.gs` 版（全列を返す）が使われている見込みです（推測）。
- `src/.omc/`（作業メモ）が clasp の rootDir 内にあります。`.clasprc.json` 以外の除外設定（`.claspignore`）はありません。clasp は `.gs/.html/appsscript.json` 以外を送らないため実害は無い見込みです（推測）。

### 4-6. 行数のめやす

実シートは見ていないため、数はすべて推測です。移行前に実測してください。

- 作業員：80 人以下（同期の上限から）。
- Requests／WorkLogs：2026-02 から運用。全行を毎回読む作り（例 `src/worklog.gs:56`）なので、数千行規模と見込みます。
- 社内カレンダー：2026 年度 124 行（確認済み）。
- FormMap：部署数 × 2。BatchLogs：運用日数ぶん。

---

## 5. 切り替え時に止める GAS の入口

| 入口 | 場所 | 止め方・注意 |
|---|---|---|
| `doGet`（画面：`top`／`approver`／`admin`／`approve2`／`staffSummary`／`redirect`） | `src/adminApi.gs:534-598` | デプロイを差し替える or アーカイブ。戻せるように **アーカイブではなく案内ページへの差し替え** が安全 |
| `doPost`（Sebastian API：`ping`／`completeOvertime`／`markHolidayStart`／`markHolidayDone`） | `src/api.gs:10-61` | Sebastian 側の URL・トークンを VPS に付け替えてから止める |
| `google.script.run` で呼ばれる `api_*` 関数群 | 各 `.gs`（画面側は別の棚卸し担当） | `doGet` を止めれば呼ばれなくなる |
| 時間トリガー 4 本（＋古いもの） | 第 3-1 節 | Apps Script 画面で一覧を控えてから削除。戻すときは `setupAllTriggers_` で作り直せる |
| `pollNewResponses_` と Google フォーム | `src/formSubmit.gs:14-114`、FormMap | 画面からの申請は 2026-03-03 にモーダル送信へ移りました（`82cc3e0`）。ただし **フォーム自体は回答受付のまま** の可能性があります（推測）。古いリンクから送られた回答は、ポーリングを止めると取り込まれなくなります。切り替え前に各フォームの回答受付を止めるか、VPS 側に取り込み口を用意してください |
| `onOpen` メニュー | `src/menu.gs:3-30` | 単純トリガーなので消せません。切り替え後に誰かが「マスタ転記」「全回答を一括処理」などを押すと DB が書き換わります。シートを閲覧専用にするか、メニューに注意書きを出してください |
| フォーム送信トリガー `handleFormSubmit_` | 古い方式 | 残っていれば削除 |
| ライブラリ利用 | `appsscript.json` の `dependencies` は空（確認済み、`src/appsscript.json:3`） | 他プロジェクトからライブラリとして使われているかは未確認 |
| 外部からの読み取り | 推測 | 「統合メール通知システム」（`fcf8e31`、`src/menu.gs:204`）と日報アプリ（`業務NOマスタ`、`3d6214b`）が DB を直接読んでいる可能性。切り替え後も DB を更新し続けるか、読み先を VPS に変える必要があります |

---

## 6. 履歴のまとめ（git log、149 コミット）

### 6-1. 時系列（主なもの）

| 日付 | コミット | 内容 |
|---|---|---|
| 2026-02-13 | `5b3571a`、`9d640f8` | シート構造の定義書・フォーム連携仕様を追加 |
| 2026-02-13 | `cdc5cd3`、`bf4c30b`、`0bd714f`、`8c4e85a` | フォーム生成・承認・実績・PDF・メール・総務画面の雛形 |
| 2026-02-14 | `2116ec5`、`b889406`、`60e4bd4`、`512c28e` | マスタ同期を追加。読み取りの再試行、作り直しの保険、空行除外と 81 行上限 |
| 2026-02-15 | `60ea917` | 書式だけの列を無視して最終列を判定 |
| 2026-02-15 | `1a53baa` | PDF 一括作成・直接書き込み方式・朝バッチ・BatchLogs |
| 2026-02-16 | `88e58bb` | 部署マスタを廃止し、作業員マスタから部署を作る |
| 2026-02-19 | `15697d5` | フォームごとのトリガーをやめ、1 分ポーリングへ（トリガー 20 本上限対策） |
| 2026-02-19 | `b49fcf1` | Google アカウントでの本人判定をやめ、部署→作業員の選択式へ（アカウント共有のため） |
| 2026-02-19 | `3d6214b` | 業務NO 連携を廃止（同期は残す） |
| 2026-02-19 | `f9d200b`、`cb9c40b` | 工番入力を簡素化、旧フォームの工番欄を自動移行 |
| 2026-02-19 | `3c55164` | WorkLogs の 9 時間ずれを修正 |
| 2026-02-20 | `3d81ab4`、`16ba1dd` | GitHub Actions の自動デプロイと `appsscript.json` を追加 |
| 2026-02-20 | `73dc59e` | 印鑑 ID を StampMap シートへ（同期で消えないように） |
| 2026-02-27 | `74b87b4`、`1fd0ec8`、`2eafaea` | WorkLogs の日時をテキスト書式の JST 文字列に固定、TZ 対策 |
| 2026-02-27 | `e16c4e4`、`20fc361` | 二次承認（`approvedBy2`／`approvedAt2` を自動追加） |
| 2026-02-27 | `c45ac6c` | `reasonDetail` 列の自動追加 |
| 2026-02-27 | `ef704a8` | 蓄積台帳・却下機能・メール改修 |
| 2026-02-27 | `72b8eaf` | 休日出勤 `targetDate` の 1 日ずれ修正 |
| 2026-02-28 | `4a84703` | 休憩を時刻ベースに変更（休憩マスタ方式から） |
| 2026-02-28 | `6b0a0f0` | 蓄積台帳への即時書き込み |
| 2026-03-03 | `82cc3e0` | **フォームをやめ、画面のモーダルから直接送信へ**（フォームのコードは残置） |
| 2026-03-03 | `69f9600` | 社内カレンダーマスタ導入 |
| 2026-03-03 | `0e94991` | PDF の 1440 分異常値対策 |
| 2026-03-04 | `8068536` | 朝バッチの対象日を前日に修正（0 件通知バグ） |
| 2026-03-07 | `e57a6d4` | `oauthScopes` を明記 |
| 2026-03-11 | `fcf8e31` | メール送信を統合メール通知システムへ移す（`morningBatchNoMail_` 新設） |
| 2026-03-12 | `6d61d0e` | 蓄積台帳を年度別・総務向け台帳に |
| 2026-03-28 | `0230f7f` → `5aa4621` | 編集・削除機能。`modified` ステータスは 30 分で廃止 |
| 2026-04-10 | `525c831`、`77ef76f` | 台帳に時間表記列。列挿入のずれ修正 |
| 2026-04-10 | `128643b`、`10f68a4` | 承認前でも完了を記録できるように。承認時の PDF 作成をロックの外へ |
| 2026-04-15 | `787a5aa` | 過去日の遡及申請 |
| 2026-04-15〜16 | `162f84e`、`0d1ec7c` | Sebastian 用 `doPost` API（残業完了・休日開始／完了） |
| 2026-04-28 | `953cb72` | Web アプリ設定を GAS 側に合わせて `USER_DEPLOYING`／`ANYONE_ANONYMOUS` に |
| 2026-05-15 | `4ce6f25` | 休憩 1 分バグ（14 件）修正、`recalcWorkLogsBreakMinutes`、`rebuildAccumulationSS` 追加 |
| 2026-05-25 | `fbae912` | Sebastian 経由でも台帳へ即時反映 |
| 2026-06-22 | `e11990d`、`e398e95` | 社内カレンダーを同期対象に、2026 年度新版 |
| 2026-06-23 | `26f81a5` | マスタ同期の 1 日ずれ修正 |
| 2026-07-07 | `a6eb4e6`、`2e7d4d1`、`a18e0d0` | スタッフビュー（締め月集計）追加 |
| 2026-11-07 | `ca2ecac` | 承認漏れに対象日を表示（HEAD） |

ブランチ：`master`（本番）、`claude/overtime-request-app-eH3CM`、リモートに `claude/document-spreadsheet-structure-6FIgO`。リモートは `github.com/skyblueearthjapan/Overtime-Holiday-Work-Request`。

### 6-2. Sebastian（Hermès Agent）連携

- 呼ぶ側：Hermès Agent「Sebastian」。URL とトークンを Hermès 側の `.env` に登録する作りです（確認済み、`src/api.gs:1-8`、`162f84e`）。
- 呼び方：`POST <Web アプリ URL>`、本文は JSON `{ "token": "...", "action": "ping|completeOvertime|markHolidayStart|markHolidayDone", "requestId": "...", "manualStartTime": "HH:mm", "manualEndTime": "HH:mm" }`。応答は JSON（確認済み、`src/api.gs:10-67`）。
- 認証：スクリプト プロパティ `SEBASTIAN_API_TOKEN` との一致だけ（`src/api.gs:18-24`）。本人確認（`assertSelf_`）は省きます。
- このため Web アプリを「全員（匿名可）」で公開しています（`953cb72`）。どのデプロイ URL を Sebastian が使っているかはリポジトリにありません（要確認：Hermès 側 `.env`）。
- 移行で必要なこと：VPS に同じ JSON 契約のエンドポイントを用意し、Hermès 側の URL とトークンを付け替える。休憩計算の違い（4-3 節）をそろえるかどうか決める。

---

## 7. 自動デプロイとデプロイの状況

### 7-1. GitHub Actions（`.github/workflows/deploy-gas.yml`）

- 内容（確認済み）：`master` の `src/**` への push か手動実行で、clasp を入れて `secrets.CLASPRC_JSON` を `~/.clasprc.json` に書き、`clasp push --force` を実行します。**`clasp deploy` はしません**。つまり成功しても HEAD（開発版）が更新されるだけで、本番デプロイの版は手で上げる作りです。
- 実績（確認済み、`gh run list` で 81 件）：成功 8 件はすべて 2026-02-20（06:08〜10:22 UTC）。2026-02-20 02:54 UTC 以降は失敗 73 件で、最新は 2026-11-07（`ca2ecac`）です。
- 失敗の理由（確認済み、最新ログ）：`invalid_grant` ／ `reauth related error (invalid_rapt)`。Google の再認証が要る状態で、保存した認証情報が使えません。
- 意味すること（推測）：2026-02-20 以降の変更は、手元から clasp で送るかエディタで直接直しています。`953cb72` の「GAS 現状に同期」のように **GAS 側だけで変えた設定がある** ため、本番のコードがリポジトリ HEAD と同じとは限りません。移行前に、本番デプロイの版（v131）の中身をエディタの版履歴か `clasp pull` で取り出し、`ca2ecac` と突き合わせてください。

### 7-2. デプロイ（依頼文で伝えられた情報。リポジトリでは確かめられません）

- 本番 Web アプリ：`AKfycbwHnaE…`、版 131（2026-11-07）。ほかに `@122`、`@121`、`@115`、`@HEAD` のデプロイがあります。
- コード・文書の中に、このアプリ自身のデプロイ URL は書かれていません（確認済み、grep）。画面内のリンクは `ScriptApp.getService().getUrl()` で動的に作ります（`src/adminApi.gs:536`）。
- 今も使われていそうな URL（推測、要確認）
  - 本番 `AKfycbwHnaE…`：社員のブックマーク、社内ポータル、LINE WORKS の案内など。
  - 社内ポータル（別アプリ）：`https://script.google.com/a/macros/lineworks-local.info/s/AKfycbx2eyJMOYP9o--GPBuhY-pj071IIR6Kqb_0xALwwNzdLQZux0dIAlL3P9EoCucnzXA/exec`。このアプリから戻るリンク先です（`src/adminApi.gs:554`、`578`、`589`）。逆にポータルからこのアプリへのリンクがあるはずなので、切り替え時はポータル側のリンクも付け替えが要ります。
  - Settings の `APP_URL`：メール本文に入ります（今は手動送信のみ）。どのデプロイかは Settings を見て確かめてください。
  - Sebastian の `.env`：`doPost` の呼び先。匿名アクセスを許すデプロイのはずです。`@122` などの古い版が Sebastian 用として残されている可能性があります。
  - 統合メール通知システムが本文に入れているアプリ URL。
  - Google フォームの URL（FormMap の `formUrl`）。古い案内に残っている可能性。
- Web アプリ設定はデプロイごとに版の `appsscript.json` に従います。リポジトリの設定は `USER_DEPLOYING`／`ANYONE_ANONYMOUS`（`src/appsscript.json:15-18`）ですが、社員向けと Sebastian 向けでデプロイを分けて設定を変えている可能性があります。各デプロイの「次のユーザーとして実行」「アクセスできるユーザー」を控えてください。

---

## 8. 移行チェックリスト（このファイルの範囲）

移行前に控えること

- [ ] DB スプレッドシートの ID・タイムゾーン設定・全シート名と行数
- [ ] Settings の全キーと値（第 1-3 節）、スクリプト プロパティ `SEBASTIAN_API_TOKEN`・`POLL_LAST_TS`
- [ ] Apps Script のトリガー一覧（関数・種類・時刻・持ち主）
- [ ] 全デプロイの ID・版・実行ユーザー・アクセス範囲
- [ ] 本番版（v131）のコードと `ca2ecac` の差分
- [ ] FormMap の全フォームの回答受付状態と、`POLL_LAST_TS` 以降の未取り込み回答
- [ ] 蓄積台帳の ID とシート一覧（`_backup_` を含む）
- [ ] StampMap・Settings の画像 ID が指す Drive ファイル、PDF フォルダの場所と権限

移行スクリプトで扱うこと

- [ ] Requests・WorkLogs はヘッダ名で読む（列順に頼らない）
- [ ] 日付・日時の Date／文字列の混在を両方受ける（第 4-1 節）
- [ ] `status=modified`、`approvedBy` の `(却下)`・`unknown` を分類する
- [ ] `breakMinutes` が 1〜9 分の行、`netMinutes` の異常値（720/960 分超）を洗い出す
- [ ] 重複申請・テストデータを洗い出し、残すか除くかを今泉オーナーに確認する
- [ ] `requestId` の一意性、Requests と WorkLogs の 1 対 1 を確かめる
- [ ] 移行後に件数・合計分数（`approvedMinutes`、`netMinutes`）を GAS 側と突き合わせる

切り替え当日

- [ ] フォームの回答受付を止める（または VPS 側で受ける）→ 最後のポーリングを待つ
- [ ] 時間トリガーを控えてから止める（戻すときは `setupAllTriggers_`）
- [ ] Sebastian の URL・トークンを VPS に付け替える
- [ ] 社内ポータル・統合メール・ブックマークのリンクを付け替える
- [ ] DB を閲覧専用にし、メニュー操作で書き換わらないようにする
- [ ] マスタ同期（元 SS → VPS）と朝バッチ（PDF・台帳）の担当を VPS 側に移す

確かめること（今泉オーナーへ）

- 統合メール通知システムと日報アプリが、この DB を直接読んでいるか
- 蓄積台帳に総務が手で書き足した内容があるか
- `recalcWorkLogsBreakMinutes` と `rebuildAccumulationSS` を本番で実行済みか
- 作業員が 80 人を超えていないか
