# 残業・休日出勤申請アプリ VPS 移行計画書（下書き）

作成開始: 2026-11-26（段階 0 下調べ中）。スキル `gas-to-vps-migration` に沿って進める。
棚卸しの元資料: `docs/migration-inventory/01-business-logic.md`・`02-screens.md`・`03-mail-pdf-forms.md`・`04-data-ops-history.md`。
改訂: 2026-11-26 抜け調査（`docs/vps-migration-gap-audit.md`、`docs/migration-inventory/gap-1`〜`gap-5`）とディレクター判断 A〜D（9-1 #18〜#21）を反映。1・4・5-5・5-6・6・7・8 章と 9-2 を書き直し、指摘ごとの手当ては 8-4 の表にまとめた。

> 書き方: 「確認済み」はコード・実物で確かめた事実。「推測」「見込み」は未確認。資料に無い値は 〇〇 か【要確認】。
> 「提案」はディレクターの案でオーナーの了承前。inventory の参照は `[01 §1.2]`（＝`01-business-logic.md` の 1.2 節）のように書く。

## 1. 結論

- 今の GAS 版（本番 @131）を、機能・画面・帳票・癖を変えずに VPS（FastAPI＋PostgreSQL＋worker＋Cloudflare Tunnel）へ移す。業務データの正は VPS の DB。今のスプレッドシートには、**VPS が書き込むたびに該当の行だけを数秒〜1 分以内に反映する**（決定 A、9-1 #18。毎晩のタブ置き換えはしない）。行の位置と並び・日付の型と表示形式を保ち、行もタブも消さない。Sebastian・統合メール通知システム・日報アプリ・PDF テンプレの数式・戻した GAS が、今と同じシートを読んで動き続ける。
- **PDF と蓄積台帳は、今の GAS（DB に紐づいた同じプロジェクト）のコードをそのまま使う**（決定 B、9-1 #19）。VPS の worker が行をシートに反映した後で、GAS の内部用の入口（`doPost` の新しい action、token 認証）を 1 件ずつ呼ぶ（4-6a）。見た目・台帳の中身・順序・失敗時の動きが今と同じになり、利用者の画面は待たない。切り替えた後の GAS は、この入口と Sebastian の転送役だけが動き、時刻のトリガーは 1 本も持たない。
- サーバーの日付の計算はすべて Asia/Tokyo で行う（決定 C、9-1 #20）。マスタ同期は VPS の worker が今と同じ規則で行い、DB の 4 タブ（`業務NOマスタ` ほか）にも今の形で書く（決定 D、9-1 #21）。
- 土台（ログイン・RPC・安全装置・outbox・バックアップ・Tunnel）は貯蔵品 `vps-app/` を写す。マスタ同期と、業務 ID を `.env` 1 か所で読む作りは工番部品 `vps-app/` を写す。業務（申請・承認・二次承認・実績・集計・スタッフ集計・Sebastian API）はこのアプリの GAS から 1 関数ずつ移す（提案）。行単位の反映と GAS の内部用の入口は、前例に無い新しい部品。
- **一番の山は「行単位の反映」と「GAS の内部用の入口」。** 段階 2 でスパイク（本番 DB のコピーで、行の反映 → GAS の PDF の入口 → 10 件の見た目の比較 → 台帳への書き込み）を行い、確かめてから段階 3 に入る（7 章の 2b）。
- このアプリは 2026-03-10 以降メールを送っていない（確認済み・BatchLogs）。VPS でもメールは送らない（提案。9-2 Q6）。
- 進行中の案件は、承認待ち 0・承認済みで完了未記録 1・二次承認待ち 6（2026-11-26 確認済み）。どれも状態のまま移せば VPS で続けられる。送り直すメールは無い。
- Sebastian は切り替えの時点では設定を変えず、GAS の `doPost` を VPS への転送役として残す（決定 #15 の案）。
- 抜け調査（5 観点、`docs/vps-migration-gap-audit.md`）の指摘は、すべて 8-4 の表で手当ての場所を決めた。
- 戻しは「VPS を止めて GAS の移転モードを外す」が中心になる（シートが常に最新のため。6-6）。
- 工数の見込みは作業 18〜26 時間（作業日で 3 日前後）＋見守り 5 営業日（推測。7 章）。

## 2. なぜ移行するか（実測）

目的（オーナー 2026-11-26）: 申請・承認のたびに出る、GAS とスプレッドシート連携による送信待ちをなくす。機能は一つも落とさない。
### 2-1. 実データと実行の記録（2026-11-26 確認済み。読むだけ）

- 申請（Requests）2,334 件（2026-02-28〜2026-11-25）。approved 2,176・canceled 158・**承認待ち 0 件**。残業 2,080・休日 254。
- 月 230〜370 件、直近 30 日で 303 件（平日 1 日あたり約 14 件）。**提出の 7 割以上が朝 6〜8 時**。
- 承認済みで作業完了が未記録 1 件（targetDate 2026-11-25）。二次承認待ち 6 件。完了済みで PDF なし 25 件（要確認）。
- 承認者は部署の共有アカウント（`admin-01@`・`prod-assem01@`・`machining@`・`imaizumi@`・`tsc-02@`・`sales-01@`）。`workerEmail` は 95% が空。
- `ApproverMap` は見出しだけ（空）。承認の判定は `DeptApprovers`（13 部署）だけで動いている。
- 作業員マスタ 50 人（同期の上限 80 人に余裕あり）、工番マスタ 560 件、業務 NO 47 件、社内カレンダー 130 行。
- WorkLogs・ApproverMap・休憩マスタは **1 行目が説明文、2 行目が見出し**。
- BatchLogs: `morningBatch`（メールあり）は 2026-03-10 が最後、`morningBatchNoMail` が毎日。**このアプリは 3/10 以降メールを送っていない**（統合メール通知システムへ移った記録と合う）。
- 本番のトリガー 4 本（持ち主 imaizumi@）: `morningBatchNoMail_`（毎朝 6:55 ごろ）、`syncAllMasters`（直近 0:32）、`nightlyUpdateAllForms_`（毎朝 6:28 ごろ、失敗率 16%）、`pollNewResponses_`（**1 分ごと、1 回 8〜20 秒**）。古いメール送信のトリガーは残っていない。
- FormMap: 有効なフォーム 21（残業 10・休日 11）。
- GAS の HEAD は 2026-11-07 04:06Z 以降変わっていない（Apps Script API の updateTime）＝ v131 と同じ中身。

### 2-2. 遅さの実測

- `pollNewResponses_` が 1 分ごとに 8〜20 秒動き、その間スクリプト全体のロックを持つ。申請・承認の書き込みは同じロックを最大 20 秒待つので、その間の操作は待たされる（確認済み・コード。gap-1 G19・gap-4 F-18）。操作ごとの秒数は利用の多い時間帯に実行数画面で取る【未実施】。

## 3. 現行機能の棚卸し

細則は `docs/migration-inventory/` の 4 文書。ここは受け入れ確認表の目次として要点だけ置く。とくに断りの無いものは確認済み（コードの全行読み）。

### 3-1. 領域ごとの要約

| 領域 | 中身の要点 | 詳細 |
|---|---|---|
| 画面 | `?page=` で 5 ページ（`top` 既定・`approver`・`approve2`・`admin`・`staffSummary`）＋権限エラー `no_auth`＋旧フォーム用 `?page=redirect`。不明な page は top。`?date=` は approve2（対象日）と admin（日次一覧の自動表示）が使う。モーダルは申請（2 ステップ）・編集・手動実績入力・承認者選択。グラフは canvas の自前描画、外部 CDN なし | [02 §1・§2] |
| 画面の入口判定 | admin・approve2 は管理者だけ。approver は管理者か DeptApprovers の承認者（**メールが空なら判定せず通す**）。top・staffSummary は誰でも | [02 §1.3]、[01 §1.1] |
| テンプレート変数 | `APP_URL`・`PORTAL_URL`（3 か所に直書き）・`DATE_PARAM`（検証なし）・`REASONS_JSON`・`OT_HOURS_JSON`・`HD_HOURS_JSON`・`message` | [02 §1.2] |
| 画面から呼ぶ関数 | `google.script.run` 40 か所・**30 関数**（申請者 13・承認者 6・二次承認 3・総務 7・スタッフ集計 1）。`api_getFormUrl` はデッドコードからだけ。返事を待たずに画面を変える楽観更新が 10 か所 | [02 §3]、[01 §1.3] |
| 画面から呼ばれないが公開されている関数 | `api_getTodayRequestsForDept`・`api_getTodayRequestsForWorker`・`recalcWorkLogsBreakMinutes`・`rebuildAccumulationSS`・`migrateAccumulationTimeColumns`・`syncAllMasters`・`testSyncWorkerOneRow`・`buildAllDeptForms`・`ensureCompanyCalendarSheet`・`forceRegeneratePdf` ほか。権限の確認なし | [01 §1.3] |
| 外部 API（Sebastian） | `doPost`。本文 JSON の `token` を `SEBASTIAN_API_TOKEN` と比べるだけ。action は `ping`・`completeOvertime`・`markHolidayStart`・`markHolidayDone`。常に HTTP 200、失敗も `{error}`、英語のエラー文あり。`targetDate` はシートの生値。`pdf` に PDF 作成の結果を入れて返す | [01 §1.2]、[04 §6-2] |
| 状態の遷移 | status は `submitted`・`approved`・`canceled` の 3 つ。却下も削除も `canceled`（却下は `approvedBy` の末尾 `(却下)` だけで区別）。二次承認は status を変えず `approvedBy2/approvedAt2` を書く。作業完了は WorkLogs の `actualEndAt`。**承認前でも完了を記録できる**。編集は完了前だけ、削除は誰でもいつでも | [01 §2] |
| 実績と休憩 | 残業の開始は対象日 17:20 固定、終了は手入力か今の時刻。休憩は**時刻窓方式**（6 つの窓を「まるごと含むときだけ」数える。日をまたぐと 0）。**Sebastian の残業完了だけ休憩マスタ方式** | [01 §3] |
| 休憩マスタ方式の実態 | コードは休憩マスタの 1 行目を見出しとして読むが、本番は 2 行目が見出し（2-1）。見出しが見つからないので、**Sebastian の残業完了の休憩は常に 0 分**になっている（確認済み＝コードと実シートの組み合わせ。実データの件数は【要確認】） | [01 §3.4・Q2] |
| 集計 | `buildJoinedRecords_` が土台（submitted も含む）。月次サマリー（40h=2400 分・60h=3600 分・注意 1800 分、すべて `>=`、予測＝ペース換算）、年度推移と特別条項（**4/1 始まり**）、未承認滞留（48h）、スタッフ集計（**締め月 16 日〜翌 15 日、3/16 年度、approved だけ**、画面の判定は `>`） | [01 §5]、[02 §5] |
| PDF | 一時コピー → セルに書く → 印鑑を画像で重ねる → 書き出し URL で PDF → `PDF_ROOT_FOLDER_ID/yyyy.MM.dd/` に保存 → 一時コピーをゴミ箱へ。方式は 2 つ（XLOOKUP 方式・直接書き込み方式）。本番の `PDF_MODE` は `direct`。**二次承認はいつも直接方式（二次承認印 G34 入り）で作り直す。Sebastian はいつも XLOOKUP 方式（G34 なし）**。テンプレは DB 自身（`TEMPLATE_SSID`＝DB）なので、**PDF 1 枚ごとに DB 全体の一時コピーができる**（確認済み＝コードと Settings の値） | [03 §2]、4-0 |
| PDF を作る場面 | 一次承認（完了済みなら）、作業完了（承認済みなら。画面 2・総務 1・Sebastian 2 の経路）、二次承認（PDF があれば作り直し）、編集（PDF があれば作り直し）、削除（ゴミ箱へ）、朝バッチ（前日の approved で PDF なし。**完了していなくても作る**） | [03 §2.2] |
| 印鑑 | 区分印（Settings `STAMP_*_FILE_ID`、F1 起点 270×270px）、一次・二次承認印（`StampMap` をメールで引く、F34・G34 起点 60×60px、セルの中央に寄せる）。2MB 以上の画像は縮小版を取る | [03 §2.4] |
| 蓄積台帳 | 別スプレッドシート。年度（3/16 始まり）ごとの `残業_YYYY年度`・`休日出勤_YYYY年度`、**27 列**、見出しの書式・1 行おきの背景色・`=HYPERLINK()`。即時の upsert（1 次承認日がいつも空）と削除、**朝バッチが前日分を消して薄い行（理由・承認者・状態が空）で書き直す癖**、手動の作り直し。開けないときは新しく作って Settings に追記する | [03 §3]、[04 §2-2] |
| マスタ同期 | `SOURCE_SSID`（作業日報_全従業員用）から 4 タブ（作業員・業務NO.・工番・社内カレンダー）を DB のタブへ全置換。作業員は 80 人で切る。読み取りの再試行、Date を JST 文字列に | [04 §1-4〜1-7・§2-1] |
| 承認者の表 | `DeptApprovers`（A 部署・B メール（カンマ区切り）・C 氏名、位置で読む）。2 行以上あればこれだけを使う。貯蔵品の VPS 版もこのタブを読む | [01 §4.3]、[04 §1-9] |
| 管理者 | Settings `ADMIN_EMAILS`＋`GENERAL_AFFAIRS_EMAILS`（大文字小文字を無視）か、ApproverMap の `role=admin`（本番は空で見出しも 2 行目なので効いていない） | [01 §4.2] |
| トリガー | 本番 4 本: `morningBatchNoMail_`・`syncAllMasters`・`nightlyUpdateAllForms_`・`pollNewResponses_`（2-1） | [04 §3-1] |
| 旧フォーム | 部署別 21 本。画面からは 2026-03-03 に外れたが、`?page=redirect`・公開 URL・1 分ごとの取り込み・毎朝の選択肢更新が生きている。取り込みは種別「残業」以外を休日にし、日付の検査をしない | [03 §4]、[01 §2.2] |
| メール | 夕方メール・朝メール（テキスト、宛先 `HR_MAIL_TO`）の関数は残るが、トリガーは無く 3/10 以降送っていない（2-1） | [03 §1] |
| メニュー | DB スプレッドシートの `onOpen` メニュー「残業休日アプリ」に 17 項目（マスタ転記・全回答の再処理・台帳の作り直しなど） | [04 §3-2] |
| ロック | すべてスクリプト全体のロック 1 種類（20 秒、一括は 30 秒）。PDF と台帳は、一次・二次承認ではロックの外、完了・編集・削除・総務の手動入力・Sebastian では呼び元のロックの中（入れ子。gap-1 G19・gap-3 G3-L6） | [01 §7] |

### 3-2. 今のまま移す癖（決定 #5「社員を信頼する作り」と「機能を落とさない」から）

- 申請の作成・編集・削除・TOP（全社）・スタッフ集計に権限の確認が無い。自分の申請を自分で承認できる（[01 §4.4] H4〜H8・H16）。
- 残業の開始 17:20 固定、休憩の「窓をまるごと含むときだけ」、日をまたぐと休憩 0（Q11）。
- 境界の `>=`（40h・60h・30h・48h・5 回・6 回）と、スタッフ集計の画面だけ `>`。
- 年度の始まりが 2 通り（3/16 と 4/1、Q16）。月次は submitted の実績も数え、スタッフ集計は approved だけ（Q14）。未来月は月末まで経過とみなす（Q15）。
- 承認者画面が当日の canceled も出す（Q4）。過去の submitted 残業は TOP に出ない（Q5）。
- 休日の日付編集で年度・未来日を確かめない（Q7）。PDF が無い申請の編集は台帳を更新しない（Q8）。完了の再実行は実績を上書きし PDF は作り直さない（Q9）。
- 申請時は残業の予定時間を一覧と照合しない（編集時だけ、Q13）。通常の休日は未来日の上限なし（Q26）。
- 朝バッチは未完了でも前日の approved の PDF を作る。
- 画面: 作業員の自動選択が効かない、承認者画面の KPI の数え方が初回と再計算で違う、STEP 2 のサマリーに過去日が出ない、admin のフィルタは［月次を更新］まで反映しない、CSV は BOM なし、approve2 の注記「翌朝バッチで…」（実際は即時）、単票の二次承認に confirm なし（[02 §2・§4.3]）。
- メールの大文字小文字の扱い（`api_getWorkerInfo`・`assertSelf_` は区別、承認者は小文字に揃える、Q24）。
- 作業員マスタ 80 人の上限（今 50 人）。

### 3-3. 直すと決めたもの（決定事項から来るもの）

| 直すこと | 元の決定 |
|---|---|
| 全画面で Google ログイン。メールが空のときの「管理者扱い」「判定を飛ばす」（H1〜H3）と `approvedBy=unknown` が起きなくなる | #3 |
| 保守の処理（台帳の作り直し・マスタ同期・休憩の再計算・PDF の作り直し）は管理者だけ。画面から呼ばれない公開関数は RPC に出さない。`?date=` は形式を確かめ、JS の文字列として安全に差し込む | #13 |
| 旧フォーム 21 本の受付を閉じる。`pollNewResponses_`・`nightlyUpdateAllForms_`・`api_getFormUrl`（フォームの自動生成）は移さない | #11 |
| 申請・承認などの書き込みに冪等キー（二重押しで 2 件できない）と状態の照合 | #5 |
| PDF は今と同じ Drive のフォルダに、今の GAS のコードで作る。シートの業務データのタブへは、行単位で今の形のまま反映する | #6・#7、#18・#19 |
| マスタ同期で、元が読めない・空のときは DB のタブを上書きしない（今の GAS は見出しだけに置き換える） | #21（4-8） |

### 3-4. ディレクターの提案（了承待ち。9-2 の質問に載せる）

- 権限エラー画面のデバッグ情報（管理者メールの一覧・ApproverMap の見出し、H13）を出さない。文言の本文は残す（Q8）。
- `?page=redirect` は top へ案内する（Q9）。
- エスケープの抜け（admin の手動入力ボタンの氏名など [02 §6]）を、見た目を変えずに直す。
- PDF の一時コピーを、失敗したときも必ず片付ける（今は例外のときに残る [03 §2.5]）。決定 B の下では GAS の関数を直す変更（`finally` 1 か所）になるので、入れるかはスパイクの後で決める（4-6）。
- 蓄積台帳が開けないときに新しい台帳を作る動きをやめ、失敗として知らせる（[04 §2-2]）。GAS の関数は直さず、内部用の入口が先に開けるかを確かめる形で行う（4-7）。
- PDF の作成に失敗した仕事を再試行する。今なら PDF なしで終わる申請に、後から PDF ができる（gap-1 G7・gap-3 G3-M2。Q20）。
- 画面のエラー文の「GASエディタのログも確認してください」などは文言のまま残す（利用者の目にはエラー時だけ）。
- 承認者の複数部署サマリーで `charts.dept` に全部署が残る（H14・Q27）は今のまま。
- 夕方・朝メールの手動送信（メニュー）は移さない（Q6）。

### 3-5. オーナーに聞くもの

9-2 の Q2・Q6〜Q10・Q12〜Q20（Q1・Q3・Q4・Q5・Q11 は決定 A〜D で不要になった）。主なものは Sebastian の休憩、共有アカウントのログインの持ち時間、総務の日次一覧（本番で空）、スマホの表示、ポータルの iframe 埋め込み、朝バッチの未完了分の PDF、PDF の再試行。

## 4. 移行先の構成

### 4-0. 確定値表

| 値 | このアプリ | 状態 |
|---|---|---|
| アプリ名（フォルダ・コンテナ・DB 名） | `overtime` | 決定（2026-11-26） |
| 公開 URL | `https://overtime.lineworks-local.net` | 決定（オーナー 2026-11-26） |
| Tunnel 名 | `overtime-vps` | 案 |
| VPS | `srv1508169`（`ssh lineworks-vps-user`） | 確認済み（2026-11-26 hostname） |
| VPS の置き場 | `/opt/overtime` | 案（2026-11-26 時点で未使用を確認） |
| docker ネットワークのサブネット | `10.93.0.0/24` | 案（2026-11-26 `docker network inspect` で空きを確認） |
| OAuth クライアント名 | `Overtime Holiday Work Request (VPS)`（GCP `sebastian01`） | 案 |
| GAS scriptId | `1GPNBR9U9mw7TZPnl1pxtdVjFagUPxt9S6qSYlbNIEuHNvRLVspNIQNc7` | 確認済み（`.clasp.json`） |
| GAS 本番 deploymentId・版 | `AKfycbwHnaECeni3FmmgtVzBZIM_xDMH3bw9Qx5sQMrW9J3XoRCOs_RmdyBensZX5Xta6sfIxA` @131 | 確認済み（2026-11-07）。切り替え前に再確認 |
| ほかの公開版（使われていないか要確認） | @HEAD `AKfycbxWkkR0…`、@122 `AKfycbwYDXF_…`、@121 `AKfycbzkQGqd…`、@115 `AKfycbxHUAZt…` | 【要確認】利用の有無 |
| GAS の実行ユーザー（委任の持ち主） | `imaizumi@lineworks-local.info`（webapp executeAs USER_DEPLOYING） | 確認済み（2026-11-07 デプロイ画面） |
| 本番スプレッドシート（DB、コンテナ） | `1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ`「残業・休日出勤申請app」（imaizumi@ 所有） | 確認済み（2026-11-26 Apps Script API の parentId） |
| **ほかのアプリが読んでいるタブ** | 貯蔵品の VPS 版がこの DB の `DeptApprovers`（5 列、見出しの文字で）を読む（スキル 00 の確定値表、gap-4 F-5）。日報アプリが `業務NOマスタ` を読む見込み（履歴 `3d6214b` に意図の記録。gap-5 高-3）。Sebastian・統合メール通知システムが Requests などを読む見込み。工番部品も要確認 | 確認済み（貯蔵品）・推測（ほか）。決定 A・D でタブを今の形で保つ |
| 使ってはいけない古いスプレッドシート | なし（名前に「残業」「休日出勤」を含むのは DB と蓄積台帳の 2 冊だけ） | 確認済み（2026-11-26 Drive 検索） |
| マスタ元（`SOURCE_SSID`、作業日報_全従業員用） | `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ`（貯蔵品と同じマスタ元） | 確認済み（Settings 2026-11-26） |
| 蓄積台帳（`ACCUMULATION_SS_ID`） | `1tKlT57SJxJijgfDVGafZ434sdVjscJJqN2uQ7P3sUyU`「残業・休日出勤 データ蓄積」 | 確認済み（Settings 2026-11-26） |
| PDF テンプレ（`TEMPLATE_SSID`） | DB と同じ `1Knx_kaQ…`（DB の中のタブがテンプレ）。`PDF_MODE` の値は「`direct で直接書込方式（未設定=XLOOKUP方式）`」（`direct` を含むので直接方式。GAS は含むかで判定する。gap-1 G4・gap-5 低-3） | 確認済み（Settings 2026-11-26） |
| DB のファイルの時間帯・地域 | 〇〇（GMT の見込み。gap-5 高-2）。段階 2 で「ファイル → 設定」から読んで記録し、`.env` に持つ（決定 C） | 【要確認】 |
| GAS の内部用の入口 | 同じプロジェクトの内部用のデプロイ（`GAS_INTERNAL_URL`）と、スクリプトのプロパティ `VPS_INTERNAL_TOKEN`（4-6a） | 案（段階 3 で作る） |
| Sebastian のトークン | GAS のスクリプトのプロパティ `SEBASTIAN_API_TOKEN` の値を VPS の `.env` へ写す（gap-4 F-16） | 確認済み（キーの所在）。値は段階 4 で写す |
| Settings の `APP_URL` | GAS の `/exec` のまま変えない（4-11） | 提案 |
| 本番 PDF フォルダ（`PDF_ROOT_FOLDER_ID`） | `1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B` | 確認済み（Settings 2026-11-26） |
| 印鑑画像 | `STAMP_OVERTIME_FILE_ID` `1zfnDZOz…`、`STAMP_HOLIDAY_FILE_ID` `1tDYK_IF…`、StampMap 6 行 | 確認済み（Settings 2026-11-26） |
| 旧フォーム | 残業テンプレ `1cFl7C_t…`、休日テンプレ `1ftXoDoW…`、置き場 `16S-phsA…`、部署別 21 本 | 確認済み（Settings・FormMap 2026-11-26） |
| メール宛先（Settings） | `HR_MAIL_TO` 3 人・`HR_MAIL_TO_EVENING_EXTRA` 1 人（すべて `@lineworks.co.jp`）。管理者 `ADMIN_EMAILS` = imaizumi@・admin-01@ | 確認済み（Settings 2026-11-26） |
| 本番データの控え（読み取り時点） | `C:\Users\imaizumi.LINEWORKS-NET\.secrets\overtime\snapshot-20261126\`（gviz CSV。集計用。移行には使わない） | 2026-11-26 |
| コードに直書きの Drive フォルダ | `1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B`（`PDF_DRIVE_FOLDER_URL` の既定値） | 確認済み（`mail.gs`） |
| 外部から呼ぶシステム | Sebastian（`doPost`、token 認証）、統合メール通知システム（DB を読む可能性） | 【要確認】置き場所・持ち主 |
| ステージングの DB（本番のコピー） | `1I0dPlH6w14UaiJVbO9i6wFCXlvGyPEJvDus8ekGptWY`「【ステージング】残業・休日出勤申請app（VPS移行試験用・消してよい）」。Settings を試験用に書き換え済み（PDF・テンプレ・台帳を試験用へ、旧フォームの ID と夕方の追加宛先は `STAGING-DISABLED`、`HR_MAIL_TO` は imaizumi@lineworks.co.jp）。FormMap タブは `FormMap_STAGING_DISABLED` に改名 | 作成済み（2026-11-26） |
| ステージングの GAS（DB に紐づいたコピー） | scriptId `1y027cB_59N1P2o3toawpCyPZ2k_W6mPETtgiDpPFtpY7TKgBWslnyLNM`。トリガーなし。初回の実行にはオーナーの「許可」が要る | 作成済み（2026-11-26） |
| ステージングの台帳（コピー） | `16M5PxyozS_tTEpQmR6jizICgZov4SoK2M1IO4TG1oHI` | 作成済み（2026-11-26） |
| ステージングの PDF フォルダ | `1KsWanVEQr-gMsb_N9mxzplEjKFibsnrc` | 作成済み（2026-11-26） |
| ステージングの GAS の公開版 | `AKfycbw9uPC4C5zlY5S3YqR9Yo0TECZETvUUgGY2ih447eL8E8XS0_VhVpwtftinTr4Nj5A` @2（team-g-gas `8524708` を clasp push）。オーナーの許可と `VPS_INTERNAL_TOKEN` の設定が要る | 作成済み（2026-11-26） |
| Cloudflare | Tunnel `overtime-vps`（id `1e31b48e…`）、DNS `overtime.lineworks-local.net` → `http://web:80`。TUNNEL_TOKEN は `~/.secrets/overtime/tunnel-token.env` | 作成済み（2026-11-26） |
| VPS の置き場 | `/opt/overtime`（lineworks 所有、750） | 作成済み（2026-11-26） |
| OAuth クライアント | GCP `sebastian01` で作る途中。GCP の画面が本人確認のパスワードを求めたため、オーナーの手が要る | 未（2026-11-26） |
| 手元の秘密の置き場 | `C:\Users\imaizumi.LINEWORKS-NET\.secrets\overtime\` | 案 |
| 社内ポータルの行 | APP_MASTER の key 〇〇 | 【要確認】 |

### 4-1. 全体の形（決定 A〜D を反映。名前と細部は提案）

貯蔵品・工番部品と同じ 6 コンテナ（決定 #1）。名前・ボリューム・ネットワークはすべて `overtime-` で始め、ホストにポートを出さない。これに加えて、**今の GAS（DB に紐づいたプロジェクト）を「PDF と台帳を作る係」として残す**（決定 B）。

| 部品 | 役目 | 写す元 |
|---|---|---|
| `overtime-db` | PostgreSQL 16。ボリューム `overtime-pgdata` | 貯蔵品 compose |
| `overtime-api` | FastAPI。ページの HTML（入口判定とテンプレート変数）、`/api/rpc/{関数名}`、`/api/sebastian`、`/auth/*`、`/healthz`。書き込みは DB の取引の中で、「シートの行の反映」と「GAS の仕事」を outbox（`jobs`）に積む。起動時に `alembic upgrade head` | 貯蔵品 `auth.py`・`rpc.py`・`google_delegate.py`・`safety.py`・`db.py`・`timeutil.py`、工番部品 `config.py` |
| `overtime-worker` | 同じイメージで `python -m app.worker`。3 つの列（レーン）を持つ。①**シートの行の反映**（`sheet_row_sync`。全体で 1 本ずつ）②**GAS の内部用の入口を呼ぶ仕事**（`gas_call`。全体で 1 本ずつ）③定期処理（4-9）。`healthcheck: disable` | 貯蔵品 `jobs.py`・`worker.py`、工番部品の `jobs.register_kind` と `sync_runs`。レーンと行の反映は新しく作る |
| `overtime-web` | nginx。ページの JS・CSS と `/api/`・`/auth/`・`/healthz` の中継。固定 IP `10.93.0.10`（案） | 貯蔵品 `web/nginx.conf` |
| `overtime-cloudflared` | Tunnel `overtime-vps`（案）。転送先 `http://web:80` | 貯蔵品 `cloudflare/setup_tunnel.py`（定数を直し、同名の Tunnel があれば止める行を足す） |
| `overtime-backup` | 毎日 3 時に `pg_dump -Fc`、1KB 未満は FAILED | 貯蔵品 `backup/`（接頭辞を直す） |
| **今の GAS**（scriptId `1GPNBR9U…`、DB に紐づく） | 切り替えた後は、内部用の入口（4-6a）と Sebastian の転送役（4-5）だけが動く。画面・RPC・トリガーは移転モードで止める（6-3） | 今のコードに `Internal.gs`（新）と `Cutover.gs`（工番部品の型）を足す |

データの流れ（1 件の承認の例）:

1. 利用者が承認を押す。api は DB の `requests` の行をロックして書き、同じ取引で `sheet_row_sync`（その申請の Requests・WorkLogs の行）と `gas_call`（`internal.afterApprove`）を積み、画面へすぐ返す。
2. worker のレーン①が、その行をシートの同じ位置に書く（数秒〜1 分以内。5-5）。
3. worker のレーン②が、行が反映済みであることを確かめてから GAS の入口を呼ぶ。GAS は今の関数で PDF と台帳を作り、Requests の PDF の 3 列をシートに書き、結果を返す。
4. VPS は返ってきた `pdfFileId` などを DB に入れる。シートの値と同じなので、レーン①は書き直さない（列ごとの差分だけを書くため）。

Google の API は GAS の持ち主 imaizumi@ の委任トークンで呼ぶ。VPS が使うのは Sheets（DB の行の反映・マスタ元の読み取り・DB のマスタ 4 タブへの書き込み）と、切り替えの控えを作る道具の Drive（読み取りと複製）だけ。PDF と台帳は GAS が作るので、VPS は PDF フォルダにも台帳にも書かない。メールを送らないので `gmail.*` のスコープは要らない（提案）。スコープは最初に決める（後から足すと許可のやり直しが要る＝罠 7）。GAS の内部用の入口は Web アプリ（匿名可・デプロイした人として実行）なので、呼ぶのに Google のスコープは要らない（token で守る）。

### 4-2. DB の表とシートのタブの対応（決定 A・C・D）

タブを 5 つに分け、「どちらが正か」と「VPS が書くか」で扱いを変える。このアプリの肝になる点。

| 区分 | タブ | DB の表 | 正 | VPS の扱い |
|---|---|---|---|---|
| A. 業務データ（**VPS が行単位で書く**） | `Requests`（見出し 1 行目） | `requests` | VPS | 移行で移す。VPS が書くたびに、その行だけを同じ位置へ書く（5-5）。行もタブも消さない |
| A | `WorkLogs`（1 行目説明・2 行目見出し） | `work_logs`（主キーは別に持つ。`request_id` で引くが、孤児と欠けを許す） | VPS | 同上。1 行目の説明文に触れない |
| A | `BatchLogs` | `batch_logs` | VPS | 同上（朝バッチの記録を VPS が末尾に足す） |
| B. 人がシートで直す設定（VPS は書かない） | `Settings`（1 行目説明・2 行目見出し・3 行目から） | `settings` | シート | 読むだけ。数分ごとに取り込む。同じキーは下の行が勝つ（GAS と同じ） |
| B | `DeptApprovers`（5 列） | `dept_approvers` | シート | 同上。**貯蔵品の VPS 版が見出しの文字で 5 列（D `役職`・E `通知メール` まで）を読む約束**。VPS は A・B・C を位置で読み、D・E は持つだけ。見出しの空白（` approverName`）とタブ名も今のまま（gap-4 F-5） |
| B | `StampMap` | `stamp_map` | シート | 読むだけ。GAS の内部用の入口も今と同じく直接読む |
| B | `ApproverMap`（1 行目説明・2 行目見出し、今は空） | `approver_map` | シート | 読むだけ。GAS は 1 行目を見出しとして読むので今は効いていない（確認済み）。同じく効かない形のまま移すかは仕様で決める |
| B | `休憩マスタ`（1 行目説明・2 行目見出し） | `break_master` | シート | 読むだけ。Q2 が決まるまでは「見出しが見つからず 0 分」を再現できる形にする |
| C. マスタ元の写し（**VPS が同期で書く**） | `作業員マスタ`・`工番マスタ`・`社内カレンダーマスタ`・`業務NOマスタ` | `workers`・`orders`・`company_calendar`（業務 NO は表を持たない） | マスタ元 | VPS の `sync_masters` が `SOURCE_SSID` を読み、今の `syncAllMasters` と同じ規則・同じ形で DB のタブに書く（決定 D。4-8）。日報アプリなどが読み続けられる。タブを消して作り直す動きは写さない |
| D. GAS が使う（VPS は触らない） | `申請書フォーム`（PDF のテンプレ）・`操作`・`休憩計算_フォーム` | なし | — | GAS の内部用の入口が PDF を作るときに今と同じく使う |
| E. 移さない | `FormMap`・`FormTemplates`・`ExportHistory`・`部署マスタ` ほか | なし | — | 触らない |

土台の表: `sessions`・`google_credentials`・`jobs`・`idempotency_keys`・`sync_runs`（貯蔵品と同じ。メールを送らないので `mail_log` は作らない）。行の反映の版を持つ `sheet_sync_state`（表・主キー・反映済みの版・シートの行番号）を足す（提案）。

- VPS が書いてよいタブは A の 3 つと C の 4 つだけ（4-10 の許可リスト）。どの区分にも入らないタブ（人が作ったメモのタブなど）にも触らない。
- 日報アプリが DB の `業務NOマスタ` を読む見込みは、履歴に意図の記録がある（`3d6214b`「日報アプリへの影響を考慮して意図的に残している」。読み手がいるかは未確認。gap-5 高-3）。決定 D で、読み手がいてもいなくても今と同じに保つ。
- Requests は自動で足された列（`reasonDetail`・`approvedBy2`・`approvedAt2`・`modifiedAt`・`modifiedBy`）が末尾にある（確認済み [04 §1-1]）。見出しは `requestType(overtime/holiday)`・`status(submitted/approved/canceled)` のように括弧まで含む（gap-4 F-13）。表は意味で列を持ち、シートとは見出しの文字で対応させる。全件空の 5 列（`jobId1`〜`jobId3`・`hrMailSentAt`・`exportError`）も列として持つ。
- 各表に `sheet_row`（シートの行番号）を持つ。移行で元の行番号を入れ、VPS で足した行は書いた位置を入れる（5-5。gap-4 F-3）。
- WorkLogs は Requests と 1 対 1 ではない（確認済み・控え: WorkLogs 2,349 行、Requests に無い ID 32 行、WorkLogs の行が無い申請 17 件。gap-4 F-4）。移行では行を作らない。申請のときに下書きの行を足し、完了などの操作で行が無ければ作る（GAS の `ensureWorkLogRow_` と同じ）。
- `updatedBy` の決まった値（`formSubmit`・`sebastian-agent`・`unknown`・`delete`・`admin-manual`・`break-recalc`）は GAS と同じ文字で書く（gap-5 低-2）。
- `isRetroactive`（過去申請）は列を作らない。`submitted_at` を JST の日付にして `target_date` と比べる（GAS の `computeIsRetroactive_` と同じ。gap-5 中-7）。
- 日時は UTC の timestamptz、日付は date。JSON では GAS が関数ごとに返していた形（`yyyy-MM-dd`、`yyyy-MM-dd HH:mm`、`yyyy-MM-dd HH:mm:ss` など。JST、オフセットなし）で返す（gap-2 S6）。
- **時刻の決まり（決定 C）**: サーバーの「今日」・年度（3/16 と 4/1）・締め月（16 日〜翌 15 日）・月の経過日数・30 日と 7 日の窓・日時の解釈は、すべて `Asia/Tokyo` で行う。`date.today()`・`datetime.now()` は使わず、`timeutil` の JST の関数だけを使う（gap-1 G1）。受け入れ試験に「06:30 JST に通常残業を出すと targetDate が当日になる」「3/16 08:00 JST に 3/15 の遡及残業が拒まれる」「15 日 23:30 と 16 日 00:30 の締め月」を入れる。DB のスプレッドシートの時間帯は段階 2 で読んで記録する（5-5）。
- 行ロック: 申請 1 件ごとの `SELECT … FOR UPDATE`。GAS の全体ロックの代わりで、別の申請は待たない。**`work_logs` だけを変える処理（完了・開始・総務の手動入力・Sebastian の 3 action・削除）も、先に `requests` の同じ行をロックする**（gap-1 G3。承認と完了が同時でも、どちらかが必ず PDF の仕事を積む）。一括の処理は ID を並べ替えてからロックする（デッドロックを防ぐ）。
- 丸めは JS と同じにする。`Math.round` は `floor(x+0.5)`、`Math.floor` は `floor`。休憩の窓は JST の時・分で比べて秒を捨てる（gap-1 G5。境界の試験は仕様）。

### 4-3. 画面の移植（提案）

- 5 ページと `no_auth` の HTML・CSS・文言・画面幅ごとの配置は今のまま。ページは api が返す（`GET /?page=…&date=…`）。admin・approve2・approver の入口判定（`no_auth` を返す）をサーバーで行うため。不明な page は top。
- **`google.script.run` の代わりの部品（シム）を 1 つ作り、各ページの呼び出しは書き換えない。** `google.script.run.withSuccessHandler(f).withFailureHandler(g).api_xxx(…)` を `POST /api/rpc/api_xxx`（本文は引数の配列）に置き換える。失敗のときは `message` を持つ Error を渡す（画面が `err.message` と `err` の両方を使うため [02 §3.6]）。書き込み系には操作ごとの冪等キーを付ける（決定 #5）。401 のときは貯蔵品と同じく書きかけを退避してログインへ。
- テンプレート変数: `APP_URL`＝新しい URL、`PORTAL_URL`＝今の固定値、`REASONS_JSON` など＝今の定数。`DATE_PARAM` は `yyyy-MM-dd` の形だけ通し、`json.dumps` で JS の文字列として出す（GAS の自動エスケープの代わり。[02 §6] の「高」）。
- CSP: 今の画面は script タグの中の JS と、行を作るときの `onclick="…"` を多く使う（確認済み）。案 1（おすすめ）: ページごとの JS を外のファイルへ移し、`onclick` は 1 つの受け口（`data-action`）に機械的に置き換え、CSP は貯蔵品と同じ `script-src 'self'`。案 2: 置き換えが大きすぎるときは、このアプリだけ `'unsafe-inline'` を許す（見た目は同じ、安全は下がる）。段階 2 で数を数えて決める。
- 公開する関数は閉じたリスト: 画面の 29 関数（`api_getFormUrl` を除く）＋管理用（保守の処理・失敗した仕事の再実行・共有タブの即時取り込み・委任の許可し直し・セッションの取り消し）。`api_getFormUrl` と旧 API 2 つは出さない（決定 #11・#13）。
- 保守の処理は今は DB のメニューから押している。VPS では管理者だけが使える呼び口と `python -m app.worker run <名前>` にする。新しい画面を作るかは仕様で決める（提案: 総務画面の下に管理者だけの小さな欄）。
- iframe: 今は `ALLOWALL`。社内ポータルが埋め込んでいなければ `frame-ancestors 'self'`（Q10）。
- スマホ: GAS は viewport を無視していたので、移すと初めて 640px 以下の配置が効く（罠 16）。現場の端末で実機確認。
- 「今日」はブラウザの時計で決まる箇所が多い（[02 §5]）。そのまま。サーバー側の日付は決定 C（4-2）。
- 抜け調査（gap-2）で足すこと:
  - 公開する画面の関数は `api_getFormUrl` を除いた **29 関数**（S9）。
  - シムは HTTP 200 の本文をそのまま成功ハンドラに渡し、例外（HTTP 400/500）のときだけ `message` 付きの Error を失敗ハンドラに渡す。成功の返事の中で失敗を返す 5 つの形（`{error}`・`{ok:false,error}`・`{ok:false,message}`・`results[].ok`・`null`）を変えない。日時の文字の形は関数ごとに今と同じ（S6。仕様に関数ごとの表）。
  - 401 のとき、楽観更新した行を元に戻すか一覧を読み直してから「ログインが切れました。ログイン後にもう一度押してください」と出して移動する。書き込みは自動で送り直さず、冪等キーは 401 で捨てない。`next` には `?page=…&date=…` を含む今の URL。申請モーダルを開くときと画面に戻ったときにセッションを確かめる（S3。受け入れに 401 を注入する試験 9 件）。
  - CSP 案 1 の対象は on 属性すべて（onclick・oninput・onchange・onmouseenter・onmouseleave、計 81 個）と JS の文字列で作る onclick。`approver.html:876` の `[onclick*=…]` で探す箇所は id で探す形に直す。style 属性は書き換えない（S5）。
  - 総務の手動入力ボタンの氏名の XSS（`'` を逃がさない）は段階 3 の必須で直す。`data-*` 属性と 1 つの受け口にし、Jinja2 の自動エスケープ、JS に入れる値は `json.dumps` のあと `<`・`>`・`&` を逃がす（S8）。日次一覧を直すか（Q16）に関係なく直す。
  - 全ページの `<title>` は GAS の `setTitle` と同じ「残業・休日出勤申請」（権限エラーは「権限エラー」）。権限エラーの文は改行しない今の見た目のまま。CSV は BOM なしのまま（S10）。
  - スマホの配置は Q17 で決める（S7）。

### 4-4. ログインと権限（提案。決定 #3・#4・#5・#12・#13）

- Google OIDC（`hd=lineworks-local.info`）。部署の共有アカウントでそのままログインする。人を見分けない（#4）。
- 管理者＝Settings の `ADMIN_EMAILS`＋`GENERAL_AFFAIRS_EMAILS`（シートから取り込み）＋`.env` の `BOOTSTRAP_ADMINS`（imaizumi@）。一覧が空でも全員を管理者にしない。
- 承認者＝DeptApprovers（シートから取り込み、小文字で比べる。今と同じ）。
- 作業者・承認者の操作の権限は今と同じ（申請・編集・削除・TOP・スタッフ集計は誰でも、`assertSelf_` は `workerEmail` が空なら素通り）。保守の処理は管理者だけ（#13）。
- **共有アカウントと朝のピーク**: 提出の 7 割が朝 6〜8 時（2-1）。貯蔵品の既定（無操作 12 時間で切れる）だと、毎朝のピークに部署の端末でログインし直しが起きる見込み（推測）。無操作の期限を延ばすかを Q7 で聞く。
- 貯蔵品の土台には絶対期限（既定 14 日）もある。無操作 12 時間なら、金曜の夕方から月曜の朝（約 60 時間）で必ず切れ、毎週月曜と連休明けの 6〜8 時に全端末がログインし直しになる（確認済み＝期限の仕組み。gap-2 S2）。切り替え翌朝は全端末が初めてのログインになるので、前日に各部署の端末で 1 回開いてもらう（6-2）。ログインに `prompt=consent` や `select_account` を付けない。段階 1 で WorkLogs の `updatedBy`、Requests の `approvedBy`・`modifiedBy` の `unknown` の割合を数え、Google にログインしていない端末の多さの見当にする。
- 管理者の判定は `ADMIN_EMAILS`＋`GENERAL_AFFAIRS_EMAILS` の和（キーの前後の空白を取る。gap-1 G10。9-1 #13 の文は `ADMIN_EMAILS` だけを書いているが、今の本番では 2 つが同じ 2 人なので差は出ない）。

### 4-5. 外部 API（Sebastian）（決定 #14・#15、決定 A・B）

- VPS に `POST /api/sebastian` を作る。本文・action・返す JSON・英語のエラー文・**常に HTTP 200** を GAS と同じにする。トークンは定数時間で比べる（外から見た動きは同じ）。Google ログインと Origin の確認は通さない（提案）。`SEBASTIAN_API_TOKEN` の値は GAS のスクリプトのプロパティから VPS の `.env` へ写す（転送は本文をそのまま渡すので、同じ値が要る。gap-4 F-16）。
- **Sebastian 用の処理は画面用と別に持つ**（gap-1 G6）。エラー文は `api.gs` の文字どおり（末尾に `: <requestId>` が付く）。失敗は `{error, action}` で `ok` は無い。JSON の誤り・トークン未設定・不一致の 3 つは `action` も無い。本文は Content-Type に関係なく生の文字列を `json.loads` する（Pydantic の本文モデルを使うと `text/plain` で 422 になる）。`completeOvertime` は `manualEndTime` を無視する。返す `status` は更新前に読んだ値。4 action × 成功・各エラーの返事が 01 §1.2 の表と一致することを受け入れ条件にする。
- 切り替えの時点では Sebastian の設定を変えない。GAS の `doPost` は移転モードのとき、受けた本文をそのまま VPS の `/api/sebastian` へ転送し、返事をそのまま返す（`UrlFetchApp`）。転送先はスクリプトプロパティで持ち、戻すときは消すだけ（提案）。内部用の入口の本文（4-6a）は転送せず、GAS の中で処理する。
- **Sebastian がどの公開版（deploymentId）を呼んでいるか**を w2403c で確かめる（#14）。使っている版にも移転モード入りの版を当てる（8-2）。当てないと、その版の GAS がシートに直接書き、VPS の DB に入らない（毎晩の照合で差として見つかるが、DB の値で上書きされる）。
- **requestId の読み先**（gap-5 高-1）: Sebastian は requestId を受け取るだけで、探す action は無い（確認済み）。シートの Requests を読んでいる見込み（推測）。決定 A でシートの行は数秒〜1 分で反映されるので、読み先がシートでも今と同じく当日の申請を見つけられる。読み先と頻度は w2403c で確かめる（9-2）。
- `completeOvertime` の休憩は、今は見出し行のずれで常に 0 分（3-1）。移行で直すかは Q2。
- PDF: Sebastian の完了は、今と同じく GAS の XLOOKUP 方式で作る（内部用の action `internal.afterComplete`、`source:"sebastian"`）。決定 A でシートの Requests・WorkLogs の行が先に反映されるので、XLOOKUP の数式も今と同じ値を引く（旧 Q3 は不要になった）。
- 返す `pdf`: GAS は PDF を作り終えてから返していた。VPS は DB に書いて行の反映と GAS の仕事を積み、**その仕事の結果を最大 〇〇 秒待ってから**、今と同じ 4 つの形（`null`・`{ok:true,pdfFileId,pdfName,folderId}`・`{already:true,pdfFileId}`・`{error}`）で返す（提案）。Sebastian の仕事は GAS の仕事の列の先頭に入れる。待ち切れたときは `{error: "…"}` の形で返し、仕事はそのまま続ける。上限は GAS の転送（`UrlFetchApp`）の待ち時間より短くし、値は段階 2 のスパイクで PDF 1 件の所要時間を測って決める。
- `targetDate`: GAS はシートの生値（文字なら `yyyy-MM-dd`、日付のセルなら UTC の ISO 文字列）を返す。決定 A で行の型を保つので、VPS も「シートの型どおりの形」で返せる。どちらにするかは Sebastian 側の使い方を確かめて仕様で決める（提案: シートの型どおり）。

### 4-6. PDF（決定 B: 今の GAS のコードで作る）

今の作り（確認済み [03 §2]・`pdfExport.gs:366-436`）: テンプレ（＝DB スプレッドシート全体）を `makeCopy` → `申請書フォーム` タブのセルに値を書く（部署の中央揃え、区分の 24pt 太字など書式も）→ `insertImage` で区分印（F1）と承認印（F34・G34、セルの中央）を重ねる → `/export?format=pdf&gid=…&portrait=true&size=A4&fitw=true&…` を Bearer 付きで取る → `PDF_ROOT_FOLDER_ID/yyyy.MM.dd/` に保存 → Requests に 3 列を書く → 一時コピーをゴミ箱へ。

**決定 B（ディレクター判断 2026-11-26、9-1 #19）**: PDF は今の GAS（DB に紐づく同じプロジェクト）が、今の関数（`generatePdfDirect_`・`fillPdfTemplate_`・`generatePdfForRequest_`・印鑑・書き出し・保存）で作る。VPS の worker は、申請の行をシートに反映した後で、GAS の内部用の入口（4-6a）を 1 件ずつ呼ぶ。GAS は今と同じく Requests の PDF の 3 列をシートに書き、結果（`pdfFileId` など）を返す。VPS は結果を DB に入れる。

これで今と同じになるもの（確認済みのコードをそのまま使うため）:

- 描画・印鑑の置き方（`insertImage`）・書き出し URL・ファイル名（`yyyyMMdd_部署_氏名_残業|休日出勤.pdf`）・フォルダ名（**対象日**の `yyyy.MM.dd`）・一時コピーの作り方・`netMinutes` の異常値の補正。
- 方式の選び方。`PDF_MODE` は「`direct` を含むか」で判定し（本番の値は説明の文つき。gap-1 G4）、二次承認はいつも直接方式で G34 入り、Sebastian はいつも XLOOKUP 方式。
- テンプレの数式（`操作!B3` を起点に Requests を引くもの）。決定 A でシートの行が先に反映されるので、今と同じ値を引く（gap-3 G3-H4・G3-M5）。ただし、直接方式で書かないセルの数式が DB の `操作!B3` の申請を引いている恐れは、今の本番にもある（推測）。スパイクで全セルの数式と `操作!B3` の値を控え、本番の PDF 3 枚（明細 1 件・業務内容なし・朝バッチの未完了分）を見る。直すかは Q19。
- Settings・StampMap・`getActive()` を読む作り。同じプロジェクトなのでそのまま動く（別のプロジェクトに切り出す案の問題 gap-3 G3-H3 が起きない）。

作る場面（きっかけ × 方式 × 既にある PDF。gap-1 G8・G9、gap-3 1 章。どれも確認済みのコードの動き）:

| きっかけ | 内部用の action | 方式 | PDF を作る条件 | 既に PDF があるとき | PDF の後の台帳 |
|---|---|---|---|---|---|
| 一次承認（個別、一括は 1 件ずつ） | `internal.afterApprove` | `PDF_MODE`（本番は直接） | WorkLogs に `actualEndAt` がある | 作らない（`already`） | 同じ `try` の中で PDF の後に upsert。PDF が例外なら書かない |
| 残業・休日の完了（画面） | `internal.afterComplete`（`source:"screen"`） | `PDF_MODE` | status が approved | `already` | PDF の成否に関係なく upsert |
| 総務の手動入力 | `internal.afterComplete`（`source:"admin"`） | `PDF_MODE` | approved | `already` | 同上 |
| Sebastian の完了（残業・休日） | `internal.afterComplete`（`source:"sebastian"`） | いつも XLOOKUP | approved | `already` | 同上 |
| 二次承認（個別・一括） | `internal.regenerate`（`reason:"secondary"`） | いつも直接（G34 入り） | 前に PDF があったときだけ | 前の PDF をゴミ箱へ → 作り直す | PDF の成否に関係なく upsert（単票は完了前でもいつも upsert。一括は新しく二次承認したものだけ） |
| 編集 | `internal.regenerate`（`reason:"edit"`） | `PDF_MODE` | 前に PDF があったときだけ | 同上 | PDF があったときだけ upsert |
| 削除 | `internal.afterDelete` | 作らない | — | ゴミ箱へ | 台帳の行を消す |
| 朝バッチ（前日の approved で `pdfFileId` なし。**未完了でも作る**） | `internal.batchPdf`（1 件ずつ）→ 最後に `internal.ledgerRewriteDay` | `PDF_MODE` | `pdfFileId` が無い | 数えるだけ（スキップ） | 全件の後で前日分を書き直す（4-7・4-9） |
| 管理者の作り直し（1 件・指定日の一括） | `internal.regenerate`（`reason:"admin"`）・`internal.batchPdf` | `PDF_MODE` | 管理者が指定 | ゴミ箱へ → 作り直す | 書かない（今の `forceRegeneratePdf`・`manualBatchPdfs_` と同じ）。`forceRegeneratePdf` の WorkLogs の開始の書き換えと列の追加は移さない（gap-3 G3-L3） |

- 二次承認・編集の「前の PDF をゴミ箱へ移し 3 列を空にする」は、今は利用者の操作の中で行う。VPS では、操作の中で DB の 3 列を空にして前の `pdfFileId` を仕事に記録し、ゴミ箱へ移すのは GAS の仕事の中で行う（順序「捨ててから作る」は同じ）。

どの経路でも守ること:

- **GAS の仕事は全体で同時に 1 本**（gap-1 G2・gap-3 G3-M1。今の全体ロックと同じ直列。台帳の行と日付フォルダの取り合いを防ぐ）。同じ申請の仕事は積んだ順に動かし、前の仕事が終わるまで次は動かさない。
- 呼ぶ前に、その申請の Requests・WorkLogs の行が「仕事を積んだ時点の版」まで反映済みであることを確かめる。まだなら待つ。GAS 側でも `expect` で確かめる（4-6a）。
- 呼ぶ前に DB の状態を読み直す。取消済みなら作らない（GAS の関数も status を見る）。
- 前の PDF をゴミ箱へ移せなくても、記録だけ残して作り直しを続ける（GAS と同じ。4 月中旬より前の PDF は部署のアカウントが持ち主の見込み。gap-5 中-3）。受け入れ試験に「持ち主が別のアカウントの PDF を持つ申請の編集・二次承認」を 1 件入れる。
- 直接方式で印鑑の画像の差し込みに例外が出ると、PDF 全体が失敗する。今のまま（gap-3 G3-L1）。
- G1（作成日）は GAS が実際に作った日。worker の遅れや再試行で日をまたぐと変わる（今も同じ性質。gap-3 G3-L4）。
- **再試行**: VPS は失敗した仕事を再試行するので、今なら「PDF なしで終わる」申請に後から PDF ができる（良い方向の差。gap-1 G7・gap-3 G3-M2）。再試行で台帳の upsert も走る。既定は「再試行する」とし、Q20 でオーナーに確かめる。変える／変えない一覧に載せる。
- PDF が 2 つできない: 再試行のとき、GAS の関数は `pdfFileId` があれば `already` を返す（確認済み）。Drive へ保存した後で Requests への記録が失敗した場合だけ、今と同じく 2 つ目ができうる（gap-3 G3-M3）。内部用の入口は保存した `pdfFileId` を返事に入れ、VPS は返事を受けたら先に DB へ記録する。同じ人・同じ日・同じ種別の 2 件が同名の PDF になるのは今どおり。
- 一時コピー（DB 全体の写し。紐づく GAS の写しも一緒にできる見込み、gap-3 G3-L7）は、今は成功したときだけ片付く。3-4 の「失敗しても必ず片付ける」は、決定 B の下では GAS の関数を直す変更（`finally` 1 か所、見た目は変わらない）になる。入れるかはスパイクの後で決める（提案）。
- 安全装置: ステージングは GAS のステージング用の写しを呼ぶ（4-10）。本番の PDF フォルダには書かない。

採らなかった案（決定 B の前の案。記録として残す）:

| 案 | 作り | 採らなかった理由 |
|---|---|---|
| 1. PDF 専用の小さな別の GAS | 新しい別の Apps Script プロジェクトに PDF の関数を写し、VPS から値を渡して呼ぶ | 別のプロジェクトでは `getActive()` が null で、Settings・StampMap・時間帯を読めない。入力の取り決めを全部作り直す必要がある（gap-3 G3-H3）。ロックが本番と別になり、日付フォルダの重複を防げない（G3-M1）。コピーの Requests が古いので XLOOKUP 方式が別物になる。台帳の順序（PDF の後に upsert）を 2 つのプロジェクトにまたがって守る必要がある（G3-H2） |
| 2. VPS だけで作り、印鑑を PDF に後から重ねる | 委任トークンで複製 → Sheets API で値と書式 → 書き出し URL → Python で印鑑を重ねる | `fitw=true` の縮みで印の位置がずれうる。合わせ込みと見比べに時間がかかる。テンプレの数式と XLOOKUP 方式を再現できない。台帳も VPS で作り直す必要がある |
| 3. HTML→Chromium（貯蔵品の方式） | テンプレを HTML で作り直す | 見た目が変わる。方針「変えない」に合わない |

スパイク（段階 2。7 章の 2b。手順の案）:

1. 本番 DB のコピー（サンドボックス。紐づく GAS の写しも入る見込み）と台帳のコピーを作り、写しの Settings の ID をサンドボックス用に直す（4-10）。`申請書フォーム`・`操作` の全セルの数式・書式・印刷設定・結合と、`操作!B3` の値を控える。
2. 写しの GAS に `Internal.gs` を入れてデプロイし、VPS の行の反映の試作で代表 10 件（残業／休日の半日・1 日、印鑑あり／なし、一次のみ／二次あり、工番 0〜3 件、補足理由あり／なし、異常値の補正 [03 §2.3 F12]、Sebastian の XLOOKUP 方式）を書いてから入口を呼ぶ。
3. 今の本番の PDF（同じ申請）と `pdftoppm` で画像にして画素ごとに比べ、差の画像を残す（G1 の作成日は除く）。オーナーが実物を見る。
4. 台帳の行を、今の GAS が書いた行と 1 セルずつ比べる（値・書式・数式）。
5. PDF 1 件の所要時間を測り、4-5 の待ち時間と 4-9 の朝バッチの時間を決める。

### 4-6a. GAS の内部用の入口の取り決め（決定 B。中身は提案）

| 項目 | 決め |
|---|---|
| 置き場 | 今の GAS（scriptId `1GPNBR9U…`、DB に紐づく）の `doPost`。新しいファイル `Internal.gs` に受け口を置き、本文に `"internal": true` があるときだけここへ回す（Sebastian の 4 action とは別の道） |
| 呼ぶ URL | 本番: 同じプロジェクトに内部用のデプロイを新しく作り、その `/exec`（`.env` の `GAS_INTERNAL_URL`）。画面用・Sebastian 用のデプロイと分け、版の当て替えで巻き込まない。版は移転モードと同じ版。ステージング: ステージング用の写しのデプロイ（4-10） |
| 認証 | 本文の `token` を、スクリプトのプロパティ `VPS_INTERNAL_TOKEN`（32 バイト以上の乱数。`SEBASTIAN_API_TOKEN` とは別の値）と比べる。GAS の `doPost` は要求のヘッダーを読めないので本文に入れる。値は VPS の `.env`（`GAS_INTERNAL_TOKEN`、権限 600）と手元の `.secrets\overtime\` だけに置く。一致しなければ何もせずに `{"ok":false,"error":"unauthorized"}` |
| 本文 | `{"internal":true,"token":"…","action":"internal.afterApprove","jobId":"<uuid>","requestId":"…","expect":{"status":"approved","actualEndAt":"…","approvedBy2":"…","modifiedAt":"…"},"args":{…}}`。`expect` は「シートの行がこの値になっているはず」という確かめ。`args` は action ごと（`source`・`reason`・`oldPdfFileId`・`targetDate` など） |
| action | `internal.ping`（版・移転モードか・紐づく DB の ID・Settings の `TEMPLATE_SSID`・`PDF_ROOT_FOLDER_ID`・`ACCUMULATION_SS_ID`・DB のファイルの時間帯を返す。何も書かない）、`internal.afterApprove`、`internal.afterComplete`、`internal.regenerate`、`internal.afterDelete`、`internal.batchPdf`、`internal.ledgerRewriteDay`（朝バッチの台帳の前日分）。中身は 4-6 の表の経路を、今の関数をそのまま呼んで今と同じ順序・条件で再現する。台帳の作り直し（`rebuildAccumulationSS`）を管理者の操作として `internal.ledgerRebuild` に置くかは仕様で決める |
| 返事 | いつも HTTP 200（GAS の Web アプリは状態コードを選べない）。成功: `{"ok":true,"action","jobId","requestId","pdf":{"status":"created｜already｜skipped｜error","pdfFileId","pdfName","folderId","pdfGeneratedAt":"yyyy-MM-dd HH:mm:ss（JST）","error"},"ledger":{"status":"upserted｜removed｜rewritten｜skipped｜error","sheetName","error"},"trash":{"status","error"},"durationMs"}`。朝バッチの `internal.batchPdf` は、BatchLogs のエラー詳細に使う `dept`・`workerName` も返す。失敗: `{"ok":false,"error":"<コード>","message":"…","retryable":true｜false}`。コードは `unauthorized`・`bad_request`・`row_not_synced`（`expect` と合わない。再試行）・`busy`（ロック待ち切れ。再試行）・`ledger_unavailable`（台帳が開けない。新しく作らない。4-7）・`ledger_header_mismatch`（4-7）・`internal`（そのほかの例外） |
| 冪等キー | `jobId`。GAS は処理した `jobId` と返事を CacheService に 6 時間残し、同じ `jobId` が来たら処理せずに同じ返事を返す（`"replayed":true`）。6 時間を過ぎた再送でも、PDF は `already`、台帳は requestId で上書きの upsert なので二重にならない |
| 直列 | VPS は内部用の入口を全体で同時に 1 本だけ呼ぶ（4-6）。GAS の関数の中のスクリプトのロック（`waitLock(20000)`）は今のまま使う。ロックが取れなければ `busy` |
| 時間 | GAS の 1 回の実行の上限は 6 分。1 回の呼び出しで 1 件だけ処理する（朝バッチも 1 件ずつ）。VPS の待ち時間は 1 回 330 秒（提案）。時間切れは「結果が分からない」として、同じ `jobId` で送り直す（冪等なので安全） |
| GAS が止まっている・失敗したとき | 再試行の間隔は 1・2・4・8・16・30 分、その後は 30 分ごとに 24 時間まで（提案）。その間も利用者の操作は止めない（画面は DB で動く）。`retryable:false` の失敗と、24 時間を過ぎた仕事は `failed` にする。管理者の欄に「GAS の仕事: 待ちと失敗の件数、最後に成功した時刻、失敗した申請と理由」と［再実行］を出す。`internal.ping` を 5 分ごとに呼び、続けて失敗したら管理者の欄と `/healthz` の詳細に出す。外への通知の先は 8-3 で決める |
| 順序の守り | 同じ申請の仕事は積んだ順。前の仕事が `failed` のままなら、同じ申請の次の仕事も止めて管理者の欄に出す（台帳を古い値で上書きしないため）。別の申請の仕事は先へ進める |
| 移転モード | 移転モードでも内部用の入口は動く（6-3）。`getDb_()` は、token を確かめた実行の中でだけ通す |
| 書くもの | 今の関数が書くものだけ: Requests の PDF の 3 列（シート）、PDF フォルダ、一時コピー、蓄積台帳。VPS はシートの PDF の 3 列を、返事の値で DB に入れるだけ |

### 4-7. 蓄積台帳（決定 B）

- 台帳への書き込み（upsert・削除・朝バッチの前日分の書き直し・新しい年度タブ）は、今の GAS の関数（`upsertAccumulationRow_`・`removeAccumulationRow_`・`appendToAccumulationSS_`）が内部用の入口の中で行う。VPS は台帳に直接書かない。書式・`=HYPERLINK()`・値の書き方（`setValues`＝入力と同じ扱い）・1 行おきの背景・書いた行だけ塗る、が今と同じになる（gap-3 G3-M4・G3-L5 の書き方の問題が起きない）。
- 書く順序と条件は、今の経路ごとと同じ（4-6 の表の右端。gap-1 G8・gap-3 G3-H2）。PDF と台帳は 1 回の呼び出しの中で、今と同じ順（PDF の後に台帳）で行う。PDF の無い経路（削除）は単独で行う。
- 癖は今のまま残る（決定 B の結果。旧 Q5 は不要になった）: upsert の 1 次承認日がいつも空、朝バッチが前日分を薄い行（理由・承認者・状態が空）で置き換える、却下では台帳の行を消さない、年度をまたぐ休日の日付編集で古い年度の行が残る（gap-3 G3-L5）。
- **書く前の見出しの確かめ**（gap-5 中-1。内部用の入口に足す新しい守り）: 書く年度タブの 1 行目が `ACC_HEADERS_`（27 列）と 1 字ずつ同じかを確かめ、違えば書かずに `ledger_header_mismatch` を返す（管理者の欄に出る）。年度タブは名前の完全一致（`^(残業|休日出勤)_\d{4}年度$`）で探し、`_backup_` 付きのタブには触れない。段階 2 で本番台帳のタブ一覧と見出し、`rebuildAccumulationSS` を実行したかを控え、今の本番で見出しが違っていればオーナーと扱いを決める。
- **台帳が開けないとき**: 今の `getAccumulationSS_` は新しい台帳を作り、Settings に行を足す。内部用の入口は、関数を呼ぶ前に `ACCUMULATION_SS_ID` を開けるかを確かめ、開けなければ `ledger_unavailable` を返して何も作らない（3-4 の提案を、GAS の関数を直さずに実現する）。
- Settings に `ACCUMULATION_SS_ID` が 2 行以上あれば、GAS は下の行を使う（gap-5 中-2・gap-3 G3-L5）。段階 2 で数え、2 行以上ならどれが本物かをオーナーと決める。
- 台帳は Requests＋WorkLogs から作り直せる派生データ（確認済み）。移行の正にはしない。

### 4-8. 共有のタブ（DeptApprovers など）とマスタ（決定 D）

- B 区分（Settings・DeptApprovers・StampMap・ApproverMap・休憩マスタ）: 今の GAS は操作のたびにシートから読むので、人が直すとすぐ効く（確認済み）。VPS は worker が 5 分ごと（案）に取り込み、管理画面に「今すぐ取り込む」を置く。直してから VPS の判定に効くまで最大 5 分のずれが出る（Q12）。GAS の内部用の入口は今と同じくシートを直接読むので、PDF の印鑑の ID は直せばすぐ効く。
- 取り込みの守り（工番部品 `sync.py` の型）: 読めない・空・有効な行 0 のときは上書きしない。キーと見出しは前後の空白を取って比べる（Settings のキーの空白は `GENERAL_AFFAIRS_EMAILS`・`STAMP_HOLIDAY_FILE_ID`。gap-4 F-14）。Settings の同じキーは下の行が勝つ（GAS と同じ。重なりは管理者の欄と報告に出す。gap-5 中-2）。DeptApprovers は位置（A・B・C）で読み、D・E は持つだけ（gap-4 F-5）。
- 値を比べる前に、`normalize_` と同じ変換（前後の空白を取り、全角英数を半角に）を、マスタ・DeptApprovers・StampMap の値にも当てる（Settings はキーの前後の空白だけ。gap-4 F-11）。同じ値の組で GAS と VPS の結果を比べる試験を作る。
- **マスタ同期（決定 D）**: VPS の worker の `sync_masters` が、毎日 0 時台（今の `syncAllMasters` と同じ時刻）と手動で動く。
  - 読む: `SOURCE_SSID` の 4 タブ（作業員・業務NO.・工番・社内カレンダー）。読み取りの再試行は今と同じ。
  - 規則: 1 列目が空の行を捨てた後の「見出し＋先頭 80 行」で作業員を切る（元の並び順）。Date は「JST の 0 時ちょうど → `yyyy-MM-dd`、それ以外 → `yyyy-MM-dd HH:mm:ss`」。UNFORMATTED で読むときは、マスタ元のファイルの時間帯で日時に直してから同じ規則で文字にする。書式だけの列は無視する（`60ea917`。gap-4 F-12）。
  - 書く: DB の `作業員マスタ`・`工番マスタ`・`社内カレンダーマスタ`・`業務NOマスタ` へ、今と同じ形（1 行目から、`setValues` と同じ USER_ENTERED）で書き、余りの行は値だけ消す。**タブを消して作り直す動きは写さない**（gap-4 F-2）。
  - 守り: 元が読めない・見出しだけ・有効な行 0 のときは上書きしない（今の GAS は見出しだけに置き換える。改善として変える／変えない一覧に載せる）。タブごとに失敗を握り、ほかのタブは続ける。
  - VPS の表（`workers`・`orders`・`company_calendar`）も同じ読み取りから作る。
  - 照合試験: サンドボックスで「GAS の `syncAllMasters` の後の DB のタブ」と「VPS の `sync_masters` の後のタブ」を 1 セルずつ比べる（値と表示）。
- 社内カレンダーの休日の候補日: A〜D 列を位置で読み、曜日は C 列の文字そのまま（今は `Sat` などの英語の略号。gap-4 F-9）、`出勤土曜` を除き、範囲は今週月曜〜来週金曜かつ今日以降。マスタが読めない・候補 0 件のときは、土日＋計算の祝日（`getKnownJapaneseHolidays_` を写す）に切り替える。Calendar API は使わない（今も権限が無く効いていない見込み。gap-5 中-6）。
- 作業員マスタの今の列にメールと在籍フラグが無いので、`api_getWorkerInfo` はいつも `null` を返す（gap-4 F-12）。列が足されたら効く作りは残す。

### 4-9. 定期処理（GAS のトリガーとの対応。決定 A〜D）

切り替えた後、**GAS の時刻のトリガーは 1 本も残さない**（4 本とも消す。6-4 B-6）。GAS に残る仕事は、VPS が呼ぶ内部用の入口（4-6a）と Sebastian の転送役（4-5）だけ。

| GAS（本番の時刻） | 切り替え後 | 中身 |
|---|---|---|
| `morningBatchNoMail_`（6:55 ごろ） | VPS `morning_batch`（6:50 ごろ）が GAS の入口を呼ぶ | ①前日の approved で `pdfFileId` が無いものを DB から拾う。②1 件ずつ `internal.batchPdf` を呼び、終わりを待つ（outbox に積んで先へ進まない。6 件以上は 1 件ごとに 500ms 待つ。gap-3 G3-H1）。③成功・スキップ・失敗を数え、BatchLogs に `実行日時・morningBatchNoMail・対象日 yyyy-MM-dd・成功・スキップ・失敗・エラー詳細（{requestId} ({dept}/{workerName}): {メッセージ}）` の 1 行を書く（名前も今と同じ。gap-4 F-15）。④最後に `internal.ledgerRewriteDay`（前日分の書き直し）。⑤行の反映が済んだことを確かめる（統合メール通知システムの朝の通知の前に、`pdfGeneratedAt` と BatchLogs をシートへ。gap-5 中-8）。未完了でも PDF を作る癖は今のまま（Q18） |
| `syncAllMasters`（0:32 ごろ） | VPS `sync_masters`（0 時台）＋手動 | 4-8。DB の 4 タブにも書く（決定 D） |
| `nightlyUpdateAllForms_`（6:28 ごろ） | 移さない | 決定 #11 |
| `pollNewResponses_`（1 分ごと） | 移さない | 決定 #11 |
| — | VPS `sheet_row_sync`（常時。積まれた仕事を数秒ごとに見る） | 5-5。決定 A |
| — | VPS `sheet_reconcile`（毎晩 2 時台） | 全行の照合と直し（5-5） |
| — | VPS `sync_sheet_config`（5 分ごと） | Settings・DeptApprovers・StampMap・ApproverMap・休憩マスタの取り込み（4-8） |
| — | VPS `gas_ping`（5 分ごと） | GAS の内部用の入口の死活と、組み合わせの確かめ（4-6a・4-10） |
| — | VPS `purge_idempotency_keys`（毎日） | 貯蔵品の積み残し（スキル 02 の 17 節） |

- 決定 #7（毎晩の書き出し）と #8（戻せる期間は毎時）は、行単位の反映と毎晩の照合で満たす（より細かい）。`EXPORT_HOURLY_UNTIL` は使わない。
- 統合メール通知システムの夕方（16:50 ごろ・18:10 ごろ）と朝（7:10 ごろ）の通知（gap-5 中-8）は、行が 1 分以内に反映されるので、今と同じシートを読める見込み（推測。w2403c で読む時刻と列を確かめる）。
- 最終の実行日時と成否は `sync_runs` に残し、管理画面に出す（GAS ではトリガーが黙って止まる前例があった）。

### 4-10. 安全装置（提案。スキル 02 の 5 節）

- `APP_ENV`（staging／production）と許可リスト。起動時と書き込みの直前の両方で確かめる。
- **ステージングは本番の DB・蓄積台帳・PDF フォルダに決して書かない。** ステージングは、本番 DB のコピー（紐づく GAS の写しも一緒にできる見込み）、台帳のコピー、ステージング用の PDF フォルダを使う。GAS は、その写しのプロジェクト（またはそこに作ったサンドボックスのデプロイ）を呼ぶ。写しの Settings の `TEMPLATE_SSID`・`PDF_ROOT_FOLDER_ID`・`ACCUMULATION_SS_ID` はステージング用の ID に直す。写しのコードの直書き（`mail.gs` の Drive フォルダ `1Bs1Fv…`、`PDF_DRIVE_FOLDER_URL` の既定値）は、台帳を作らない守り（4-7）と Settings の値で使われないことを確かめる。
- VPS は起動時と 5 分ごとに `internal.ping` を呼び、返った DB の ID と Settings の 3 つの ID が、自分の `APP_ENV` の許可リストと合うかを確かめる。staging なのに本番の ID が返れば、GAS の仕事をすべて止めて知らせる（本番の GAS を呼んでしまう事故を防ぐ）。
- 書いてよいスプレッドシート: production は DB `1Knx_kaQ…` だけ（蓄積台帳は GAS が書く。VPS は書かない）。staging は DB のコピーだけ。
- 書いてよいタブ: A 区分の `Requests`・`WorkLogs`・`BatchLogs`（行単位の反映と照合だけ）と、C 区分の 4 タブ（`sync_masters` だけ）。それ以外のタブ（とくに `DeptApprovers`・`Settings`・`申請書フォーム`・`操作`）は拒む。
- 書き方の制限: 行・列・タブの挿入・削除・作り直し・並べ替えは使わない（5-5）。
- 書いてはいけない: マスタ元 `1iu5Hoak…`、社内ポータルのシート、貯蔵品・工番部品のシート、蓄積台帳、PDF フォルダ（VPS は Drive に書かない。切り替えの控えを作る道具は別で、複製だけ）。
- GAS の内部用の入口の URL と token は `APP_ENV` ごとに別の値。staging の `.env` に本番の URL を書けない（許可リスト）。
- 業務の ID は `.env` の 1 か所で持ち（工番部品の型）、Settings の値と食い違えば知らせる。DB のファイルの時間帯も `.env` に持ち、シートの設定と食い違えば行の反映を止める（gap-5 高-2）。

### 4-11. メール（提案）

- 3/10 以降メールを送っていない（確認済み）ので、VPS でも送らない（Q6）。夕方・朝メールの関数は GAS に残す（戻したときのため）。
- 統合メール通知システムは**このシートを読んで送っている見込み**（推測 [03 §0]、gap-5 中-8）。決定 A で行が 1 分以内に反映され、列・型・並びも保つので、読み先が変わらない。Settings の `HR_MAIL_TO_EVENING_EXTRA` はこのコードで使われていない（gap-3 G3-L8）ので、統合メール側が Settings も読んでいる見込み（推測）。Settings は人が直すタブのままで、VPS は書かない。
- Settings の `APP_URL` は GAS の `/exec` のままにする（提案。統合メールの本文のリンクは移転ページを通って新しいアプリの同じ画面へ届く。戻したときも直さなくてよい。gap-4 F-16）。

## 5. データ移行

### 5-1. 移す元と読み方

- 移す元は DB `1Knx_kaQ…` だけ（古い DB は無い。4-0 確認済み）。道具は ID が違えば止まる（照合は `--source sheets` のときだけ働く。罠 12）。
- **Sheets API で `UNFORMATTED_VALUE`（シリアル値）と `FORMATTED_VALUE` を一緒に読む。gviz の CSV は使わない**（型の混ざった列が空で出た。罠 11、8-1）。2026-11-26 の gviz の控えは件数の見当にだけ使う。
- 列は見出しの文字で照合する（前後の空白を取り、全角英数を半角に＝GAS の `normalize_` と同じ）。欠け・重なりで止める。決まりに無い見出しは `droppedColumns` に出す。
- 移す前に DB の「ファイル → 設定」のタイムゾーンと地域を控える（[04 §4-1]。【要確認】）。

### 5-2. タブごとの決まり

| タブ | 見出し行 | 決まり | 確かさ |
|---|---|---|---|
| `Requests` | 1 行目 | 見出し名で全列。末尾に足された列も名前で取る。`requestId` の重なりがあれば止める | 確認済み（形）、件数は 2-1 |
| `WorkLogs` | **2 行目**（1 行目は説明） | `requestId` で引くが 1 対 1 ではない（孤児 32 行・行の無い申請 17 件。gap-4 F-4）。Requests に無い ID は止めずに移して `orphans` に載せる（外部キーは張らない。工番部品の前例）。行の無い申請に行を作らない。行番号と、型が混ざる列（`updatedAt`）の行ごとの元の型も移す | 確認済み（2-1・控え） |
| `Settings` | 2 行目、データは 3 行目から | **キーの前後の空白を取って読む**（`GENERAL_AFFAIRS_EMAILS ` など）。同じキーが 2 行以上あれば下の行が勝つ（GAS と同じ）。重なりは報告に出す（gap-5 中-2）。取り込むだけ（正はシート）。ID のキーは `settingsForEnv` に出して `.env` と突き合わせる | 確認済み（2-1・8-1） |
| `ApproverMap`・`休憩マスタ` | **2 行目**（1 行目は説明） | 取り込むだけ（正はシート）。ApproverMap は空 | 確認済み（2-1） |
| `DeptApprovers` | 1 行目 | A・B・C を位置で。B はカンマで分けて小文字に。13 部署 | 確認済み |
| `StampMap` | 1 行目 | `メール`（小文字）・`stampFileId`。6 行 | 確認済み |
| `BatchLogs` | 1 行目 | 全行を移す（`実行日時` は Date、`対象日` は文字と Date の混在） | 確認済み（形）。件数は gviz の見当で 279 行（`morningBatchNoMail` 259・`morningBatch` 20。gap-4 F-14） |
| マスタ 4 タブ | 1 行目 | 移さない。切り替えでマスタ元から同期し、DB のタブにも今の形で書く（決定 D） | 提案（工番部品の前例） |

### 5-3. 型と意味の決まり（[04 §4]）

- 日時の型の混在: `targetDate` は文字と Date の両方を受け、Date はシートの時間帯で日付を取る。WorkLogs の `actualStartAt/EndAt` は文字なら JST、Date ならシートの時間帯で書式化した文字を JST とみなす（GAS の読み方と同じ）。`updatedAt`・`modifiedAt`・BatchLogs の `対象日` も両方を受ける。
- 2026-02-19 より前の WorkLogs の 9 時間ずれ、2026-02-27 より前の休日の 1 日ずれは**値を直さずに移す**（今の画面・集計と同じ数字にするため）。一覧を報告に出す。Requests は 2026-02-28 からなので、該当は WorkLogs の孤児だけの見込み（推測。dry-run で数え、0 件ならそう書く。gap-5 低-6）。
- `isRetroactive` は列を作らずに求める（4-2）。移行で `submittedAt` の時刻を 1 秒も変えない（gap-5 中-7）。
- `approvedMinutes` は「申請時の予定時間」（承認で変わらない）。`orderNo1-3` は工番の写し（受注先ではない）。`canceled` は削除と却下の両方（却下は `approvedBy` の末尾 `(却下)`）。どれもそのまま移す。
- `status` は `submitted/approved/canceled` だけを許す（値で照合する＝罠 13）。`modified`（2026-03-28 の 30 分だけあった値）が 1 件でもあれば止めて報告し、扱いをオーナーと決める。
- 分の列は整数。空は 0 にせず空のまま持ち、JSON では今と同じ形で返す。`jobId1`・`hrMailSentAt` はいつも空だが列として持つ（シートへの行の反映で同じ列を保つため）。
- 洗い出して報告するだけのもの（直さない）: 休憩 1〜9 分の行（2026-05 の不具合の名残。14 件の記録あり）、`netMinutes` が残業 720 分・休日 960 分を超える行、`approvedBy=unknown`、同じ作業員・同じ日・同じ種別で数分以内に出た重なり、2026-02-13〜02-20 の試験らしい行、完了済みで PDF なしの 25 件、Sebastian 経由の残業（`updatedBy='sebastian-agent'`）の件数と休憩。
- テスト行の除外: 決まった接頭辞は無い（ID は UUID）。既定は「全部移す」（提案。Q14）。

### 5-4. 照合（一致しなければ切り替えない）

| 項目 | 中身 |
|---|---|
| 件数 | Requests・WorkLogs・BatchLogs の行数（シートの行数−除外） |
| 状態・種別・月 | status×requestType の件数、targetDate の月×種別の件数（2-1 と同じ形） |
| 分の合計 | `approvedMinutes`・`actualMinutes`・`breakMinutes`・`netMinutes` の合計（全体・月別） |
| 進行中 | submitted、approved で `actualEndAt` なし、approved で `approvedBy2` なし、の件数 |
| PDF | `pdfFileId` が入っている件数、approved かつ net>0 で PDF なしの件数（月次の「PDF 未作成」と同じ数） |
| その他 | `approvedBy2` 入り、却下（`(却下)`）、`workerEmail` 入り、WorkLogs の孤児（今 32）、WorkLogs の行が無い申請（今 17）、の件数 |
| 抜け調査で足した項目 | 月×種別の過去申請（`isRetroactive`）の件数を新旧で比べる（gap-5 中-7）。canceled で WorkLogs に実績が残る件数（04-10 より前の取消。gap-5 低-1）。approved で `targetDate` が日付にならない行（gap-3 G3-L2）。PDF があり、PDF を作った時刻が完了より前の件数（朝バッチの未完了分。gap-5 中-4、Q18） |
| 型 | `verifyProblems: []`・`droppedColumns: {}` |
| 総当たり | サンドボックス（DB のコピー）で GAS の集計（月次サマリー全月×全部署、年度推移、特別条項、スタッフ集計の全締め月×全員、TOP・承認者・二次承認の一覧）を JSON に書き出し、VPS の同じ関数の結果と比べる（提案。読むだけの診断関数をサンドボックスにだけ足す） |
| 見比べ | 任意の 20 件を新旧の画面で。PDF は 4-6 のスパイクの 10 件 |

### 5-5. シートへの行単位の反映（決定 A。決定 #7・#8 を満たす）

目的: VPS の DB を正にしたまま、今のシートの A 区分のタブを「GAS がその場で書いたのと同じ状態」に数秒〜1 分で保つ。シートを読むもの（Sebastian・統合メール通知システム・日報アプリ・PDF のテンプレの数式・GAS の内部用の入口・戻した GAS）が、今と同じに動くようにする。設計は提案。型の扱いは段階 2 のスパイクで確かめる。

流れ:

1. api は DB の取引の中で、変えた行ごとに `sheet_row_sync` の仕事（タブ・表の主キー・行の版）を積む。同じ行の仕事はまとめる（最後の版だけを書けばよい）。
2. worker の行の反映のレーン（全体で 1 本）が、DB の**今の**値を読んでシートに書く（積んだときの値ではない。順序が入れ替わっても最後は正しくなる）。
3. 書いたら、その行の「反映済みの版」を DB に記録する。GAS の仕事はこれを見て待つ（4-6）。

行の位置と並び（gap-4 F-3）:

- 移行で各行に元の行番号（`sheet_row`）を入れる。今ある行はその位置に書く。VPS で足した行は、作った順にシートの末尾（最後のデータ行の次）に足す（GAS と同じ「新しい行は末尾」）。
- 書く前に、その行の ID の列（Requests と WorkLogs は `requestId`、BatchLogs は `実行日時` と `バッチ名`）を読み、DB の値と同じかを確かめる。違えば（人が並べ替えた・行を挿入した）、ID の列を全部読んで位置を探し直し、`sheet_row` を直してから書く。見つからなければ書かずに止め、管理者に知らせる。
- 行の並べ替え・挿入・削除はしない。

消さない（gap-4 F-2、gap-3 G3-M5）:

- 行・列・タブの削除・挿入・作り直しはしない。申請の削除は今と同じく、status を `canceled` にし、WorkLogs の 5 列を空にする「値の書き換え」で表す（GAS も行を消さない）。
- 余りの行を消す処理は無い（行単位なので要らない）。毎晩の照合でも行は消さない。

書く列と型（gap-4 F-1、gap-5 高-2・中-9）:

- シートの今の見出しを読み、見出しの文字で列を合わせる（列の並びを変えない）。決まりの列が無ければ止める（GAS が列を足して並びが変わるのを防ぐ。gap-1 G15・gap-4 F-13）。決まりに無い列（右側）には触れない。WorkLogs は 1 行目の説明文に触れず、2 行目の見出しを確かめる。
- 書くのは DB で変わった列だけ（列ごとの差分）。GAS の入口が書いた PDF の 3 列などを、古い値で上書きしないため。
- **型の表を段階 2 で作る**: Requests・WorkLogs・BatchLogs の全列を `UNFORMATTED_VALUE` で読み、列ごと（混ざる列は行ごと）に「日付・日時のセル（数値）か文字か」と表示形式（`numberFormat` の型とパターン）を控える。書くときは、その型に合わせる。
  - 日付・日時のセルの列（`submittedAt`・`approvedAt`・`approvedAt2`・`pdfGeneratedAt`、BatchLogs の `実行日時`。`targetDate`・`modifiedAt` は型の表で決める）: **DB のファイルの時間帯の壁時計**でシリアル値を作り、RAW で書く。書いたセルに、同じ列の今の表示形式を `repeatCell` で毎回当てる（新しく足した行に表示形式が無いと、戻した GAS が数値として読み、未承認滞留 48h などが狂う）。DB のファイルの時間帯は段階 2 で「ファイル → 設定」から読んで記録し（決定 C）、`.env` に持つ。地域設定も控える（gviz が `m/d/yyyy` で出るので en_US の見込み。推測）。
  - 文字の列（WorkLogs の `actualStartAt/EndAt` の `yyyy-MM-dd HH:mm:ss`（JST、書式 `@`）、自由入力）: RAW で書く（罠 11。`'` を付けない）。自由入力に `1/2` のような文字があると、GAS の `setValues` は日付に変えていたが、VPS は文字のまま書く（受け入れる差。変える／変えない一覧に載せる）。
  - 行ごとに型が混ざる列（WorkLogs の `updatedAt` は文字 約 2,325 行・日時のセル 23 行、`modifiedAt`）: 移行で行ごとの元の型を持ち、その型で書き戻す。VPS で足す行は、GAS がその経路で書く型にする（例: 申請の下書き行の `updatedAt` は日時のセル、完了は文字）。
  - 分の列は整数。空は空のまま。`jobId1`〜`jobId3`・`hrMailSentAt`・`exportError` は空のまま。
- 受け入れ条件: サンドボックスで、VPS が足した行・書き換えた行を GAS の `getValues()` で読み、日時の列が `Date` で返ること。`getRequestById_`・`computeIsRetroactive_`・承認者画面の提出列・未承認滞留が元の値と同じになること（gap-4 F-1、gap-5 高-2）。

ほかのタブの決まり:

- WorkLogs は Requests と 1 対 1 ではない（gap-4 F-4）。孤児 32 行と、行の無い申請 17 件はそのまま。VPS は GAS が行を作る場面でだけ作る。行数の照合は「移行時の行数＋VPS で足した行」と比べる。
- BatchLogs は末尾に足すだけ。
- B 区分は書かない。とくに DeptApprovers は 5 列・見出しの文字・タブ名が貯蔵品との約束（gap-4 F-5）。Settings を取り込むときは、同じキーは下の行が勝つ（GAS と同じ。gap-5 中-2）。
- 蓄積台帳は VPS が書かない（GAS が書く。書く前に見出しを確かめる＝4-7、gap-5 中-1）。

毎晩の照合（`sheet_reconcile`、2 時台）:

- A 区分の 3 タブを UNFORMATTED と FORMATTED で全行読み、DB と 1 セルずつ比べる（日時は値で、表示形式も見る）。
- 差があれば DB の値でその行を書き直し、直した行・列・前後の値を `sync_runs` と管理者の欄に残す。シートにあって DB に無い行（GAS や人が書いた行）は、消さずに報告する。
- 歯止め（罠 49）: DB の Requests か WorkLogs が 0 行のとき、またはシートより DB の行が少ないときは、1 行も書かずに止めて知らせる。
- 書かずに数えるだけの形（`--check`）も持つ。切り替えの当日（B-16）と戻しの前（6-6）に使う。
- 同じ列で差が続けて出るなら、型の表か書き方の誤り。リハーサルの間に 0 にする。

運用:

- 「A 区分のタブを手で直しても、その行の次の書き込みか毎晩の照合で DB の値に戻る」ことを管理者に伝える。B 区分は今までどおり人が直す。
- 反映の遅れ（積んでから書くまで）と待ちの件数を管理者の欄に出す。目安は 1 分以内。5 分を超えたら知らせる。

### 5-6. リハーサル

- ステージングで本番からの移行を 2 回以上。数字（件数・分の合計・状態別）を控え、当日の dry-run と比べる（当日は同じか増えている）。
- 移行の直後に毎晩の照合を `--check` で流し、差が 0 であること（移行した値を書き戻しても、型・表示形式・位置が変わらないこと）。
- 画面からの操作（申請・承認・完了・二次承認・編集・削除・Sebastian の 3 action・一括承認 20 件）の後、行が 1 分以内にシートに出て、5-5 の型の受け入れ条件を満たすこと。
- GAS の内部用の入口（写し）で PDF と台帳ができ、完了済み 20 件の一括承認で台帳に 20 行が重ならずに入ること（gap-1 G2）。写しの GAS を止めた状態で操作し、再試行と管理者の欄の表示を確かめること。
- 戻しの練習: VPS を止め、写しの GAS の移転モードを外し、画面・集計・PDF・Sebastian が今と同じに動くこと（6-6）。書き込んだ行を GAS の `getRequestById_`・`computeIsRetroactive_`・承認者画面の提出列で読み、元の値と一致すること（gap-5 高-2）。

## 6. 切り替え計画

### 6-1. 方針

- 並行稼働はしない。受け入れテストと JEV ハンズの実機テストが済んだ日の、業務の後の夜に行う（決定 #16）。始める前に声をかける（#17）。
- 夜にする理由: 提出の 7 割が朝 6〜8 時（2-1）。朝バッチ（6:55 ごろ）より前に終えれば、翌朝の利用者は新しいアプリしか見ない。
- 手順書は工番部品の `docs/cutover-runbook.md` を写して値を直し、このリポジトリの `docs/cutover-runbook.md` に書く（段階 5 で読むだけのリハーサルを 1 回）。
- 決定 A・B の後は、シートが常に最新で、PDF と台帳は初めから GAS が作る。切り替え後の GAS は「移転モード＋内部用の入口＋Sebastian の転送役」だけが動く。戻しは「VPS を止めて移転モードを外す」が中心になる（6-6）。

### 6-2. 前もって済ませること

- ステージングの受け入れテスト（29 関数・5 画面・PDF と台帳の見比べ・Sebastian の 4 action・行の反映と型・GAS が止まったときの再試行・06:30 JST の試験）と、JEV ハンズでオーナーが主な流れと管理画面を押したこと。
- 本番からの移行のリハーサル 2 回以上、照合がすべて一致、移行直後の照合（`--check`）で差 0。
- GAS に移転モードと内部用の入口（`Internal.gs`）を入れた版を作り、**本番のデプロイ、Sebastian が使っているデプロイ、新しい内部用のデプロイ**に当てる。`MIGRATED_TO_URL` が無い間は今と同じ動きであることを画面で確かめる。版の `webapp` の設定（USER_DEPLOYING・ANYONE_ANONYMOUS）を `clasp pull --versionNumber` で確かめる（罠 45）。戻す版は @131。
- `VPS_INTERNAL_TOKEN` を本番の GAS のプロパティと VPS の `.env` に置き、本番の DB に向けた VPS から `internal.ping` が通ることを確かめる（ping は何も書かない）。
- DB のファイルの時間帯と地域設定、型の表（5-5）、蓄積台帳のタブと見出し（4-7）、Settings の重なり（4-8）を控えてある。
- トリガーの画面で「すべての人のトリガー」を見て、imaizumi@ 以外の持ち物が無いこと（gap-4 F-17）。DB の編集者の一覧と、`操作` タブのボタンに割り当てたスクリプトの有無を控える（gap-4 F-7）。
- 部署の共有アカウント（13 部署）が新しい URL（ステージング）でログインできること。翌朝のピークに初めてのログインが重ならないよう、切り替えの前日の業務中に各部署の端末で 1 回開いてもらうか、Q7 の答えに合わせて決める（gap-2 S2）。
- 旧フォーム 21 本の直近の回答が 0 件であること（決定 #11 の条件）を確かめ、`POLL_LAST_TS` を控える。

### 6-3. GAS の移転モード（`Cutover.gs` の型。工番部品の版を写して直す）

| 入口 | 移転モードでの動き |
|---|---|
| `doGet` | 「新しいアドレスへ移りました」のページ。**`page` と `date` を引き継いだボタン**（`https://overtime.lineworks-local.net/?page=approve2&date=…`）を出す。統合メール通知システムなどの本文にある `?page=…&date=…` のリンクのため（貯蔵品は `requestId` だけだった）。ページは `<base target="_top">`、ボタンも `target="_top"`（GAS の iframe の中で新しいアプリやログイン画面を開かないため。gap-2 S4）。自動の移動は「動けばよい」扱い（罠 29） |
| 画面から呼ぶ関数と、公開されている保守の関数 | DB を開く 1 か所 `getDb_()`（`config.gs:69`）の先頭で止める（罠 47）。例外の文は「新しいアドレスへ移りました」。**ただし内部用の入口が token を確かめた実行の中だけは通す**（実行の中だけの印を見る） |
| `doPost`（内部用の入口） | **止めない**。`"internal": true` の本文は token を確かめ、4-6a の action を処理する |
| `doPost`（Sebastian） | 止めずに転送役にする（4-5）。転送先のプロパティが無い間（当日の数分）は `{"error":"service is moving, retry later"}`（HTTP 200）を返す（提案） |
| トリガーの 4 関数（`morningBatchNoMail_`・`syncAllMasters`・`nightlyUpdateAllForms_`・`pollNewResponses_`） | 先頭で何もせずに戻る。トリガー自体も 4 本とも消す。決定 D で `syncAllMasters` も VPS に移るので残さない（旧 Q11 の「GAS に残す」案と手順の食い違い＝gap-5 高-3 はこれで解消） |
| `getDb_()` を通らない入口（gap-4 F-7） | `setupAllTriggers_`・`setup*Trigger_`・`cleanupFormSubmitTriggers_`・`reprocessAllResponses_`・`debugProcessLatest_`・`debugPollStatus_` の先頭にも移転モードの判定を置く（トリガーが作られる・`POLL_LAST_TS` が 48 時間前に戻されるのを防ぐ）。`api_getHolidayCandidateDates` の従来方式（`CalendarApp`）への逃げ道も先頭で止める |
| メニュー（`onOpen`） | 移転モードのときは「移転済み（新しいアドレス）」の 1 項目だけを出す |
| 運用関数 | `enterMigrationModeForOvertime()`（引数なしの包み）、`setSebastianForwardUrl()`、`setInternalToken()`、`closeAllFormsForMigration()`（21 本の受付を閉じて ID を記録。フォームと回答は消さない。gap-4 F-19）、`removeAllTriggersForMigration()`（4 本）、`exitMigrationMode()`、`restoreTriggersAfterRollback()`（作る前に同じ関数のトリガーが無いことを確かめる）、`showMigrationStatus()`（読むだけ。移転モードでも動くよう `getDb_()` の判定を通さない。gap-4 F-8）。持ち主（imaizumi@）が操作しているときだけ通す（罠 48） |

### 6-4. 当日の流れ（案）

| # | 手順 | 確かめ | 戻す範囲 |
|---|---|---|---|
| B-0 | `ssh lineworks-vps-user hostname` | `srv1508169` | — |
| B-1 | 進行中の件数を数える（submitted、approved で完了未記録、二次承認待ち、PDF なし） | 数を控える（2026-11-26 は 0・1・6・25） | — |
| B-2 | GAS の実行数画面で動いている実行が無いこと。`showMigrationStatus` で旧フォームの未取り込みの回答 0、版、トリガー 4 本 | 0 件 | — |
| B-3 | **先にフォームを閉じる**: `closeAllFormsForMigration`（gap-4 F-8。移転モードの後は取り込みが止まるため） | 21 本が「回答を受け付けない」。記録が残る | 記録から受付を戻す |
| B-4 | 2 分待ち、1 分ごとの取り込みが最後に動いたことを実行数画面で見る。`showMigrationStatus` で `POLL_LAST_TS` 以降の回答 0 件 | 0 件。あれば取り込みが済むまで待つ | — |
| B-5 | `enterMigrationModeForOvertime` | `/exec` が移転ページ。`/exec?page=approve2&date=<日付>` のボタンで、新しいアプリの同じ画面が**最上位の窓で**開く | `exitMigrationMode` |
| B-6 | `removeAllTriggersForMigration` | 4 本を消した。「すべての人のトリガー」に残りなし | `restoreTriggersAfterRollback` |
| B-7 | 1〜2 分待つ（入口を通り過ぎた書き込みと PDF 作成が終わるのを待つ）。もう一度 `showMigrationStatus`（罠 30） | 実行数画面に動いている実行が無い。未取り込み 0 | — |
| B-8 | **控えを作る**: DB と蓄積台帳を Drive API で複製（「【切替前の控え】… 日付」、共有しない） | 控えの Requests・WorkLogs・BatchLogs・マスタ 4 タブ・台帳の行数が元と同じ | — |
| B-9 | VPS の api・worker を止め、`pg_dump -Fc` を `/opt/overtime/cutover/` へ | exited、ダンプ 1KB 超 | `docker compose start api worker` |
| B-10 | 表を空にする: 業務の全表＋`jobs, idempotency_keys, sync_runs, sessions, sheet_sync_state`（残すのは `google_credentials` と `alembic_version`） | 0 件、委任 1 件 | `pg_restore` |
| B-11 | 移行 `--source sheets --dry-run` → 本実行（`sheet_row` と行ごとの元の型も入れる） | `verifyProblems: []`、`droppedColumns: {}`、件数がリハーサル以上 | B-10 から |
| B-12 | 照合 SQL（5-4）。Settings の ID と `.env` の一致（PDF フォルダ・台帳・テンプレ・マスタ元）。DB のファイルの時間帯と `.env` の一致（gap-5 高-2） | 全部一致 | B-10 からか C へ |
| B-13 | `.env` を控えて `APP_ENV=production`、`GAS_INTERNAL_URL`・`GAS_INTERNAL_TOKEN` が本番の値。`APP_IMAGE_TAG` が最新 | 値と権限 600 | 控えを戻す |
| B-14 | **api だけ**起動して安全装置を確かめる（工番部品の手順書 B-11 の python を写す。ステージングのシート・フォルダ、マスタ元、ポータル、蓄積台帳、PDF フォルダへの書き込みが拒まれ、DB の A 区分 3 タブと C 区分 4 タブだけ通る）。`internal.ping` が本番の DB の ID と移転モード on を返す | `NG` 0 行 | ログを見て直すか C へ |
| B-15 | マスタ同期と共有タブの取り込みを手で 1 回（`docker compose run --rm --no-deps -T api python -m app.worker run sync_masters`。worker はまだ起動しない） | 作業員 50・工番 560・業務 NO 47・カレンダー 130 前後。DB の 4 タブの行数と先頭の行が控えと同じ。DeptApprovers 13 部署、管理者 2 人 | 控えから 4 タブの値を写し戻す |
| B-16 | worker を起動し、照合を `--check` で 1 回 | A 区分の差 0、`sheet_row_sync` の待ち 0、B・D 区分のタブ（特に DeptApprovers）が変わっていない、`gas_ping` 成功 | worker を止め、控えから写し戻して C へ |
| B-17 | `setSebastianForwardUrl` → GAS の `/exec` へ `ping` を送る | VPS のログに届き、返事が `{ok:true,pong}` | プロパティを消す |
| B-18 | 画面で確認（オーナーのアカウントと部署の共有アカウント 1 つ）: TOP、承認者画面、二次承認の 6 件、総務の月次と日次、スタッフ集計、完了未記録 1 件の表示。古いリンク → 移転ページ → 新しいアプリの同じ画面。書き込みはしない | 期待どおり | — |
| B-19 | 社内ポータルの APP_MASTER の url 列を新しい URL に（前に声をかける） | ポータルから新しい URL が開く | 元の `/exec` に戻す |
| B-20 | 最終確認: 全コンテナ up、jobs の pending・failed 0、`gas_ping` 成功、外からの healthz 200、エラーログなし | — | — |
| B-21 | 引継ぎメモ・メモリ・スキルを更新し、オーナーへ報告 | — | — |

- B-8 の控えから写し戻す道具（A 区分と C 区分のタブを値と表示形式ごと本番へ写す。Sheets API の `copyPaste` か、控えのタブを `copyTo` で写して値を移す）を段階 3 で作り、段階 5 のリハーサルで 1 回使う（gap-4 F-10）。決定 A では行の反映がタブを置き換えないので、使う場面は B-15・B-16 の失敗だけの見込み。

### 6-5. 進行中の案件

- **送り直すメールは無い**（このアプリはメールを送っていない＝確認済み）。送り直しの道具は作らない（提案）。
- 承認待ち・承認済みで完了未記録・二次承認待ちは、状態と実績をそのまま移せば VPS の画面で続けられる（TOP の期限超過、二次承認の一覧に出る）。
- 翌朝に確かめること（提案）: 最初の申請が 1 分以内にシートの Requests の末尾に出たこと（型と表示形式も）。最初の PDF が GAS の入口で作られ、DB とシートの 3 列が同じこと。VPS の朝バッチ（6:50 ごろ）が前日分の PDF を作り、BatchLogs に `morningBatchNoMail` の 1 行、台帳に前日分を書いたこと。本番で最初に作られた PDF 1 件をオーナーが見る。統合メール通知システムの通知がいつもどおり出たこと。Sebastian の完了が届き、PDF が XLOOKUP 方式で作られたこと。
- 完了済みで PDF なしの 25 件は切り替えでは触らない（VPS の再試行は新しく積んだ仕事だけで、この 25 件には動かない）。原因を調べて作るかを Q13 で決める。調べ方: 25 件の `approvedAt2`・`modifiedAt`、完了日と対象日の差を見る（二次承認・編集の作り直しの失敗、前日以外の完了、ロック待ちの失敗が候補。gap-1 G7・gap-3 G3-M2）。

### 6-6. 元に戻す（C。戻すと決めるのはオーナー）

決定 A でシートは常に最新、決定 B で PDF と台帳は初めから GAS が作っている。そのため、書き出しを待つ・写し戻す作業は要らず、「VPS を止めて、GAS の移転モードを外す」が中心になる。

1. **C-1** VPS の api を止める（新しい書き込みを止める）。worker は動かしたまま。
2. **C-2** worker の待ちを流しきる: `sheet_row_sync` の待ち 0、`gas_call` の待ち 0（GAS が動いていれば数分）。照合を `--check` で 1 回流し、差 0 を確かめる。`failed` や流しきれない GAS の仕事があれば、申請 ID と action を控える。**worker を止める**（止めないと、GAS が戻ってから書いた行を、毎晩の照合が DB の値で書き直してしまう）。
3. **C-3** GAS で `exitMigrationMode`（Sebastian の転送も外れる。内部用の入口は token が無ければ何もしないので、残してよい）→ `restoreTriggersAfterRollback`（`morningBatchNoMail_`・`syncAllMasters` を作る。同じ関数のトリガーが既にあれば作らない。フォームの 2 本は決定 #11 のとおり戻さない。オーナーが戻すと言えば戻す）。
4. **C-4** `syncAllMasters` を手で 1 回動かし、ログの `OK` 4 行を確かめる（決定 D で DB のマスタ 4 タブは VPS が書いていたので、GAS の同期が今の形で書けることを確かめる。gap-4 F-6）。
5. **C-5** C-2 で控えた仕事を、GAS で手で処理する（メニューの PDF の作り直し。台帳は管理の関数で upsert）。
6. **C-6** 本番 `/exec` で一覧が出て、VPS で最後に作られた申請が見えること。GAS の `/exec` に Sebastian の `ping` を送る。
7. **C-7** ポータルの url 列を元の `/exec` に戻す。Settings の `APP_URL` は変えていないので何もしない。Sebastian の設定も変えていないので何もしない。
8. 旧フォームは閉じたまま（決定 #11）。フォームに戻すには `openForm` の行の変更と受付の再開が要る（gap-5 低-4）。
9. 利用者への案内はオーナーと相談。

- その日の朝バッチを VPS が既に流していれば、GAS の朝バッチは翌朝から動かす。同じ日に GAS の朝バッチを手で流すと、BatchLogs に 2 行目ができ、台帳の前日分をもう一度書き直す（結果は同じ）。
- 決定 A で型と表示形式を保つので、戻した後に日時の見え方が変わる行は出ない見込み（罠 54 は当てはまらない。リハーサルの戻しの練習で確かめる）。
- 戻した後は VPS の DB を使わない。もう一度切り替えるときは移行からやり直す。

## 7. 段階と工数

AI エージェントが作り、オーナーが確かめる前提の見込み（推測）。前例は貯蔵品（関数 23・PDF・メール・承認。相談から切り替えまで約 6.5 時間）と工番部品（関数 16・xlsx。半日）。このアプリは関数 29＋Sebastian 4・画面 5・集計が多い・行単位の反映・GAS の内部用の入口があり、前例より大きい。

| 段階 | やること | 完了の条件 | 見込み |
|---|---|---|---|
| 0. 下調べ | 確定値表・実データ・トリガー | 4-0 が埋まり始めている | 済 |
| 1. 計画 | この計画書、抜け調査（5 観点）、9-2 の質問への回答 | `docs/vps-migration-plan.md` と `docs/vps-migration-gap-audit.md` があり、9 章が埋まる | 2〜3 時間（抜け調査と決定 A〜D まで済） |
| 2. 仕様 | RPC の閉じたリスト（29＋管理）、関数ごとの引数と返事の形の表（gap-2 S6）、担当をまたぐ関数の表、変える／変えない一覧、安全装置、表の定義、型の表（5-5）、GAS の内部用の入口の取り決め（4-6a） | `docs/vps-implementation-spec.md` | 2〜3 時間 |
| 2b. スパイク（決定 A＋B） | 本番 DB のコピー（サンドボックス。紐づく GAS の写し）と台帳のコピーで: DB のファイルの時間帯と型の表を取る → 行の反映の試作で 10 件を書く → 写しの GAS に `Internal.gs` を入れて入口を呼ぶ → PDF 10 件を本番の PDF と画素ごとに比べる → 台帳の行を 1 セルずつ比べる → GAS の `getValues()` で型を確かめる → PDF 1 件の所要時間を測る（4-6 のスパイクの手順） | 10 件の見比べの結果、型の受け入れ条件の合否、所要時間をオーナーに報告。問題があれば計画を直してから段階 3 | 3〜4 時間（段階 2 と並行） |
| 3. 作る | F 土台 → U 画面（シム・5 ページ）・W 業務（申請・承認・実績・集計・Sebastian）・S 行の反映と照合・G GAS 側（`Internal.gs`・`Cutover.gs`）・D 移行と写し戻しの道具 を並行 → つなぎ → 正しさと安全のレビュー → 修正 | 自動テストが全件通る、レビューの指摘が片付く、貯蔵品・工番部品の固有の文字列が grep で 0 件 | 7〜10 時間 |
| 4. 置く | OAuth クライアント、Tunnel・DNS、VPS の `.env`、委任の許可、ステージング用の GAS の写しのデプロイと token | 外から `/healthz` が `env: staging`、委任が許可済み、`internal.ping` がステージングの ID を返す | 1 時間＋オーナーの操作待ち |
| 5. 試す | 本番の写しを移す、受け入れ（API・画面・PDF・台帳・Sebastian・行の反映と型・GAS が止まったとき・06:30 JST）、集計の総当たり比較、JEV ハンズの実機、スマホ実機、手順書のリハーサル 2 回、戻しの練習 | 受け入れ表がすべて合格、照合が一致、「実機で押した／自動テストだけ／まだ」を分けて報告済み | 4〜6 時間 |
| 6. 切り替え | 6-4 の流れ | 照合一致、安全装置、古いリンク、Sebastian の転送、`gas_ping`、報告済み | 1 時間 |
| 7. 見守り | 5 営業日。毎日、行の反映の遅れ・GAS の仕事の待ちと失敗・毎晩の照合の差・同期・バックアップを見る。翌朝の朝バッチ | 期間を問題なく過ぎる。ステージングの片付け | 毎日 15 分 × 5 日 |

合計の見込み: 作業 18〜26 時間（作業日で 3 日前後）＋見守り 5 営業日（推測）。前の見込み（15〜22 時間）から、行単位の反映・GAS の内部用の入口・スパイク 2b の分が増え、PDF と台帳を別の GAS か VPS で作り直す分が減った。スパイクで型の扱いや GAS の所要時間に問題が見つかると、2〜4 時間増える見込み。

## 8. リスクと運用で決めること

### 8-1. 棚卸しで見つかった主なもの

棚卸しで見つかった主なもの（詳細は inventory 各文書の末尾）:

- 公開設定が「誰でも（匿名）」で、メールが空のとき管理者扱いになる箇所がある（`approval.gs:135`、`adminApi.gs:569`）。
- 外部 API（Sebastian `doPost`）は token だけで認証し、常に HTTP 200 を返す。休憩の計算が画面と違う。
- 旧 Google フォームの経路（1 分ごとの回答取り込み・毎朝の選択肢更新）が残っている可能性。
- ~~古いトリガー（夕方メール・朝バッチ）が本番に残り、二重送信している可能性。~~ → 残っていない（2026-11-26 トリガー画面で確認）。
- DB は貯蔵品の VPS 版が `DeptApprovers` を読む共有のシート。切り替え後も DeptApprovers は今のシートで人が編集し、両アプリが読む形を保つ。
- Settings のキー名に前後の空白がある（`GENERAL_AFFAIRS_EMAILS`（後ろ 2 つ）と `STAMP_HOLIDAY_FILE_ID`（前 1 つ・後ろ 2 つ）。gap-4 F-14）。移行ツールは空白を許す。
- gviz の CSV は型の混ざった列を空にすることがある（WorkLogs の見出しが一部空で出た）。移行は Sheets API（UNFORMATTED と FORMATTED）で読む（スキル 07 の 11 番）。
- 日付・日時の型が時期によって混在（2026-02-19・02-27 の修正前の行）。
- 年度の始まりが 2 通り（3/16 と 4/1）、集計の数え方が画面で違う（今のまま移す）。

### 8-2. 追加のリスクと手当て（2026-11-26 抜け調査と決定 A〜D で改訂）

| リスク | 起きると | 手当て | 確かさ |
|---|---|---|---|
| PDF の見た目が変わる | 総務の書類として使えない | 決定 B で今の GAS のコードで作る。スパイクで 10 件を画素ごとに比べる | 確認済み（同じコードを使う）。比べるのは【未実施】 |
| テンプレの書かないセルの数式が別の申請を引く | 今の本番でも PDF に別の人の値が出ている恐れ | 決定 A で行が先に反映されるので、切り替えで悪くはならない。今の誤りを直すかは Q19 | 推測（gap-3 G3-H4） |
| 統合メール通知システムが Requests などを読んで通知している | 行の反映が遅れると通知に出ない | 決定 A で 1 分以内に反映。朝バッチの後の反映を確かめる（4-9）。w2403c で読み方と時刻を確かめる | 推測（[03 §0]・gap-5 中-8） |
| Sebastian が古い公開版を呼んでいる | 転送が効かず、GAS が直接書いた結果が DB に入らない（毎晩の照合で DB の値に戻る） | 使っている deploymentId を確かめ、移転モード入りの版を当てる。切り替え後に GAS の実行数で `doPost` を見る | 推測 |
| Sebastian の返事の形が変わる | Sebastian 側の判定が狂う | 4-5 の返事の表との一致試験。`pdf` は結果を待って今の形で返す | 推測 |
| Cloudflare が GAS からの転送（Google の IP）を止める | Sebastian の完了が届かない | ステージングで GAS → VPS の転送を試す。止められたらそのパスだけ例外にする | 推測 |
| Sebastian の残業の休憩が常に 0 分 | 画面から押した場合と実残業が違う | Q2 で決める。`updatedBy='sebastian-agent'` の行（gviz の見当で 97 行）の内訳を Sheets API で数え直す | 確認済み（コード×実シート）、件数【要確認】 |
| 日報アプリなどが DB の `業務NOマスタ` ほかのタブを読む | 切り替え後に古いままになる | 決定 D で VPS が今と同じ形で書き続ける | 履歴に意図の記録あり、読み手は推測（gap-5 高-3） |
| 共有タブの取り込みの遅れ | DeptApprovers を直してから最大 5 分、VPS の判定に効かない | 「今すぐ取り込む」ボタン。運用で伝える（Q12） | 提案 |
| 行の反映で共有タブを壊す | 貯蔵品の承認判定まで止まる | 書いてよいタブを A 区分 3 つと C 区分 4 つに限る安全装置と試験。B-16 で DeptApprovers が変わっていないことを見る | 提案 |
| 人が A 区分のタブを並べ替える・行を挿入する | 行の位置がずれ、別の行を上書きする | 書く前に ID を確かめ、違えば探し直し、見つからなければ止める（5-5） | 提案 |
| 書き込んだ日時が数値として読まれる・時間帯がずれる | 戻した GAS・統合メールで滞留 48h・過去申請・提出時刻が狂う | 型の表、DB のファイルの時間帯の壁時計、表示形式を毎回当てる、受け入れ条件（5-5） | 確認済み（仕組み。gap-4 F-1・gap-5 高-2）。DB の時間帯は推測（GMT の見込み） |
| 行の反映が遅れる | Sebastian・統合メール・GAS の入口が古い値を読む | 行の反映と GAS の仕事のレーンを分ける。GAS の仕事は反映済みの版を待つ。遅れを管理者の欄に出し、5 分超で知らせる | 推測 |
| GAS が止まる・遅い（Google の障害・実行の上限） | PDF と台帳が遅れる（利用者の操作は止まらない） | 再試行と管理者の欄、`gas_ping`（4-6a）。1 日約 14 件＋朝バッチなので実行の上限には余裕がある見込み | 推測 |
| 内部用の token が漏れる | 外から PDF と台帳の処理を呼ばれる（書けるのは今の関数の範囲） | Sebastian と別の値。`.env` は権限 600。漏れたら GAS のプロパティと `.env` を同時に入れ替える | 提案 |
| ステージングが本番の GAS を呼ぶ | 本番の PDF フォルダ・台帳に試験の PDF・行ができる | `internal.ping` で ID を確かめ、合わなければ止める（4-10） | 提案 |
| 戻した後も worker が動いている | 毎晩の照合が、GAS が書いた値を DB の値で書き直す | 戻しで worker を必ず止める（6-6 C-2） | 提案 |
| 共有アカウントのログイン切れ | 毎週月曜と連休明けの朝 6〜8 時に全端末でログインし直し。切り替え翌朝は全端末が初めてのログイン | 期限を延ばす（Q7）。前日に各部署の端末で 1 回開く（6-2） | 確認済み（期限の仕組み）／推測（端末の実態）。gap-2 S2 |
| 画面の JS が CSP で動かない | ボタンが効かない | 4-3 の案 1（on 属性 81 個と JS の文字列で作る onclick、`approver.html:876` の探し方も）。全ボタンを JEV ハンズで押す。だめなら案 2 | 推測（gap-2 S5） |
| `?date=` と手動入力ボタンの氏名からの XSS | 画面で任意の JS が動き、その人の権限で RPC を呼べる | `?date=` は形式を確かめ `json.dumps` で差し込む。手動入力ボタンは `data-*` 属性にする（段階 3 の必須。gap-2 S8）。試験を書く | 確認済み（[02 §6]） |
| 権限エラー画面の情報漏れ | 管理者のメール一覧が見える | デバッグ情報を出さない（Q8） | 確認済み |
| 蓄積台帳の見出しが 27 列の並びと違う | 列のずれた行を書く | 内部用の入口が書く前に見出しを確かめる（4-7） | 推測（gap-5 中-1） |
| 台帳が開けないとき新しい台帳ができる | 台帳が二重になる | 内部用の入口が先に開けるかを確かめ、作らずに失敗として知らせる（4-7） | 確認済み（コード） |
| PDF の一時コピーが残る | DB 全体の写し（個人の実績）が Drive に残る | 今は失敗時に残る。直すかはスパイクの後で決める（4-6） | 確認済み（コード） |
| 移行時の型ずれ・異常値 | 画面の数字が変わる | 値を直さずに移し、一覧を報告。集計の総当たり比較（5-4） | 確認済み（[04 §4]） |
| `status=modified` の残り | 移行が止まる | dry-run で数え、あればオーナーと扱いを決める | 推測 |
| 本番の版とリポジトリの差 | 移した動きと GAS の内部用の入口の土台が本番と違う | `clasp pull`（@131）と `ca2ecac` を突き合わせる（2-1 では HEAD の更新日時から同じ中身と判断） | 推測（突き合わせは【未実施】） |
| 旧フォームに回答が来ていた | 申請を取りこぼす | 先にフォームを閉じ、取り込みが最後に動いた後で未取り込み 0 を確かめる（6-4 B-3・B-4） | 当日に再確認 |
| 切り替えた夜から翌朝の朝バッチ | 切り替えた日の PDF・台帳が抜ける | VPS の朝バッチの初回を翌朝に確かめる（6-5） | 提案 |
| 移転モードの間に誰かが保守の処理やトリガーの作成を押す | DB が書き換わる・持ち主の違うトリガーができる | `getDb_()` と、`getDb_()` を通らない入口の先頭で止める。メニューは 1 項目だけ（6-3） | 確認済み（入口。gap-4 F-7） |
| 承認・完了の直後に PDF のリンクが出ない | 今は操作の中で作るので、押した直後に出ていた | GAS が裏で作る（数十秒〜で出る見込み）。見え方の差をオーナーに伝える | 推測 |
| 写した貯蔵品・工番部品の固有値が残る | 利用者を別アプリへ案内する、別の Tunnel を上書きする | grep で 0 件を段階 3 の完了条件にする（スキル 02 の 17 節） | 確認済み（前例） |

### 8-3. 運用で決めること

- VPS の外へのバックアップの置き場と、外からの死活監視・通知先（9-2。貯蔵品・工番部品でも未決）。GAS の仕事の失敗と行の反映の遅れを、管理者の欄のほかにどこへ知らせるか（通知先は 〇〇）。
- GAS の内部用の入口の token の置き場（`.secrets\overtime\`）と入れ替えの手順。
- シートの B 区分（Settings・DeptApprovers・StampMap）を直す人と、直してから最大 5 分で VPS に効くこと。
- DeptApprovers は 5 列・見出しの文字・タブ名が貯蔵品との約束。列を消したり見出しを直したりしない（gap-4 F-5）。
- A 区分のタブ（Requests・WorkLogs・BatchLogs）と C 区分のマスタ 4 タブは、手で直しても VPS の値に戻ること（管理者に伝える）。並べ替え・行の挿入もしない。

### 8-4. 指摘の手当て

抜け調査 5 文書（`docs/migration-inventory/gap-1`〜`gap-5`）の全指摘と、手当ての場所。「仕様」は `docs/vps-implementation-spec.md` の受け入れ条件に入れるもの。Q は 9-2 の質問。

| 指摘 | 重さ | 中身（短く） | 手当ての場所 |
|---|---|---|---|
| gap-1 G1 | 高 | サーバーの「今日」・年度・締め月を JST で | 決定 C（9-1 #20）、4-2 の時刻の決まり、仕様（06:30 JST ほか 3 本） |
| gap-1 G2 | 高 | worker の並行で台帳の行・日付フォルダが壊れる | 4-6（GAS の仕事は全体で 1 本）、4-6a の直列、5-6（一括承認 20 件） |
| gap-1 G3 | 中 | 承認と完了が同時だと PDF ができない | 4-2 の行ロック、仕様 |
| gap-1 G4 | 中 | `PDF_MODE` は「含むか」で判定 | 4-0 の値を直した。決定 B で GAS の判定をそのまま使う |
| gap-1 G5 | 中 | 丸めが JS と Python で違う | 4-2、仕様（境界の試験） |
| gap-1 G6 | 中 | Sebastian の返事の細部 | 4-5、仕様（4 action × 返事の表） |
| gap-1 G7 | 中 | 二次承認・編集の作り直しの失敗で PDF が消えたまま | 4-6 の再試行、6-5、Q13・Q20 |
| gap-1 G8 | 中 | 台帳に書く・消すきっかけ | 4-6 の表、4-7 |
| gap-1 G9 | 中 | きっかけ × 方式 × 既存の PDF | 4-6 の表 |
| gap-1 G10〜G19 | 低 | 管理者の定義（G10）、並び順（G11）、削除の副作用（G12）、時刻の形（G13）、開始のやり直し（G14）、見出しの自動追加（G15）、スタッフ集計の氏名合わせ（G16）、一括承認の順番（G17）、滞留の返事（G18）、ロックの記述（G19） | G10 は 4-4（2 キーの和。9-1 #13 の文は既存の行なので変えず、仕様で和にそろえる）、G12 は 4-6 の表、G15 は 5-5、G19 は 3-1 と 2-2。そのほかは仕様 |
| gap-2 S1 | 高 | 総務の日次一覧が本番で空（2026-11-26 確認） | Q16、仕様（変えるならその形） |
| gap-2 S2 | 高 | 共有アカウントのログイン切れと朝のピーク | Q7、4-4、6-2 |
| gap-2 S3 | 中 | 楽観更新 9 か所と 401 | 4-3、仕様（401 を注入する試験 9 件） |
| gap-2 S4 | 中 | 移転ページ・ログインは iframe の中で開けない | 6-3 の `doGet`、6-4 B-5、Q10 |
| gap-2 S5 | 中 | CSP 案 1 の置き換えの数 | 4-3、段階 2 で案 1／案 2 を決める |
| gap-2 S6 | 中 | シムの取り決めと返事の形 | 4-3、4-2、仕様（関数ごとの引数と返事の表） |
| gap-2 S7 | 中 | スマホの配置を今と同じにするか | Q17、4-3 |
| gap-2 S8 | 中 | 手動入力ボタンの XSS | 4-3（段階 3 の必須）、8-2、仕様の試験 |
| gap-2 S9・S10 | 低 | 公開する関数は 29、題名・改行・CSV などの細かな差 | 4-3・6-2・7 章を 29 に、S10 は 4-3 と仕様 |
| gap-3 G3-H1 | 高 | 朝バッチの順序と BatchLogs の数 | 4-9 の `morning_batch`、4-6a |
| gap-3 G3-H2 | 高 | 台帳の upsert の順序と条件が経路ごとに違う | 決定 B（同じ呼び出しの中で今の順）、4-6 の表、4-7 |
| gap-3 G3-H3 | 高 | 別の GAS に切り出すと `getActive()` が使えない | 決定 B で不要（同じプロジェクト）。4-6 の採らなかった案 |
| gap-3 G3-H4 | 高 | テンプレの書かないセルの数式 | 4-6、スパイク 2b、Q19 |
| gap-3 G3-M1 | 中 | 同時実行と日付フォルダの重複 | 4-6、4-6a の直列 |
| gap-3 G3-M2 | 中 | 失敗したときの戻り方（25 件の原因の候補） | 4-6 の再試行、6-5、Q13・Q20 |
| gap-3 G3-M3 | 中 | 再試行で PDF が 2 つできる | 4-6、4-6a の冪等キー |
| gap-3 G3-M4 | 中 | 値の書き方（入力と同じ扱い） | 決定 B で台帳は GAS が書くので不要。5-5 の自由入力は RAW（受け入れる差） |
| gap-3 G3-M5 | 中 | タブを作り直すとテンプレの数式が壊れる | 5-5（行もタブも消さない）、4-10 |
| gap-3 G3-L1〜L8 | 低 | 印鑑の例外（L1）、壊れた targetDate（L2）、手動の PDF 関数（L3）、作成日（L4）、台帳の細かい癖（L5）、ロックの記述（L6）、一時コピーの GAS（L7）、`HR_MAIL_TO_EVENING_EXTRA`（L8） | L1・L3・L4・L7 は 4-6、L2 は 5-4 の照合、L5 は 4-7、L6 は 3-1、L8 は 4-11 |
| gap-4 F-1 | 高 | 日時のセルが新しい行で数値として読まれる | 5-5 の型、受け入れ条件、5-6 |
| gap-4 F-2 | 中 | 行・タブを消さない約束 | 5-5、4-8（マスタも作り直さない）、4-10 |
| gap-4 F-3 | 中 | 行の並びの保存 | 5-5 の行の位置、4-2 の `sheet_row` |
| gap-4 F-4 | 中 | WorkLogs は 1 対 1 ではない | 4-2、5-5、5-4 の照合 |
| gap-4 F-5 | 中 | DeptApprovers は 5 列で貯蔵品が見出しで読む | 4-2、5-5、8-3 |
| gap-4 F-6 | 中 | マスタのタブの鮮度と移転モードの衝突 | 決定 D、6-3、6-6 C-4 |
| gap-4 F-7 | 中 | `getDb_()` を通らない入口 | 6-3、6-2 |
| gap-4 F-8 | 中 | 切り替えの順番で旧フォームの回答を取りこぼす | 6-4 B-3〜B-7、6-3 の `showMigrationStatus` |
| gap-4 F-9 | 中 | カレンダーの曜日の文字 | 4-8 |
| gap-4 F-10 | 中 | 控えから写し戻す道具が無い | 6-4 の注（段階 3 で作り、段階 5 で 1 回使う） |
| gap-4 F-11 | 中 | `normalize_` を値にも当てる | 4-8、仕様の試験 |
| gap-4 F-12〜F-19 | 低 | マスタ同期の細部（F-12）、Requests の見出しと列（F-13）、件数の穴埋め（F-14）、BatchLogs の名前（F-15）、秘密と `APP_URL`（F-16）、トリガーの持ち主（F-17）、取り込みのロック（F-18）、フォームの回答を残す（F-19） | F-12 は 4-8、F-13 は 4-2・5-5、F-14 は 5-2・8-1、F-15 は 4-9、F-16 は 4-5・4-11・4-0、F-17 は 6-2・6-4 B-6、F-18 は 2-2、F-19 は 6-3 |
| gap-5 高-1 | 高 | Sebastian が requestId をどこから得るか | 決定 A（シートが 1 分以内に最新）、4-5、9-2 の確かめ |
| gap-5 高-2 | 高 | DB のシートの時間帯とシリアル値 | 決定 C、5-5、4-10、6-4 B-12、6-2 |
| gap-5 高-3 | 高 | 業務NOマスタの転記は日報アプリのため | 決定 D、4-2、4-8、6-3 |
| gap-5 中-1 | 中 | 台帳の書く前の見出しの確認 | 4-7、6-2 |
| gap-5 中-2 | 中 | Settings のキーの重なり（後の行が勝つ） | 4-7、4-8、5-5、5-2 |
| gap-5 中-3 | 中 | 旧 PDF をゴミ箱へ移せないときに止めない | 4-6、仕様の試験 |
| gap-5 中-4 | 中 | 朝バッチは未完了でも PDF を作り、後から直らない | Q18、5-4 の照合 |
| gap-5 中-5 | 中 | Sebastian の休憩 0 分は作り漏れ | Q2 |
| gap-5 中-6 | 中 | 休日の候補日の逃げ道 | 4-8、仕様の試験 |
| gap-5 中-7 | 中 | 過去申請は毎回求める | 4-2、5-4 の照合 |
| gap-5 中-8 | 中 | 統合メールは朝バッチの結果を読む見込み | 4-9、4-11 |
| gap-5 中-9 | 中 | `targetDate`・`modifiedAt` の型を変えない | 5-5 の型の表 |
| gap-5 低-1〜低-6 | 低 | 04-10 より前の取消の実績（低-1）、`updatedBy` の値（低-2）、`PDF_MODE` の判定（低-3）、旧フォームのコード（低-4）、画面に返す日時は文字（低-5）、2 月の型ずれの行（低-6） | 低-1 は 5-4、低-2 は 4-2、低-3 は 4-0、低-4 は 6-6、低-5 は仕様（S1 と同じ洗い出し）、低-6 は 5-3 |

## 9. オーナーに決めてもらう事項と決定事項

### 9-1. 決定事項

| # | 項目 | 決定 | 日付・根拠 |
|---|---|---|---|
| 1 | 構成 | 貯蔵品・工番部品と同じ（FastAPI＋PostgreSQL＋worker＋Cloudflare Tunnel、今の画面を移植） | 2026-11-26 オーナー「共通の進め方は同じでよい」 |
| 2 | 公開 URL | `overtime.lineworks-local.net` | 2026-11-26 オーナー |
| 3 | ログイン | 全画面で Google ログイン（今は匿名でも開ける） | 2026-11-26 オーナー |
| 4 | Google アカウントの使われ方 | **部署ごとに 1 つを共有**（加工班・組立班・機械設計・電気設計・営業・資材購買・総務など）。上席も一般社員も同じアカウント | 2026-11-26 オーナー |
| 5 | 権限の作り | 今の GAS と同じ「社員を信頼する」作り。アカウントで人を見分けない（#4 のため）。状態の照合と二重押しの防止だけ足す | 2026-11-26（#4 と共通の決定から） |
| 6 | PDF の保存先 | 今と同じ Google ドライブのフォルダ | 2026-11-26 オーナー（共通の決定） |
| 7 | 業務データの控え | 毎晩、今のスプレッドシートの今のタブに、今と同じ列・形で書き出す（正は VPS の DB） | 2026-11-26 オーナー（共通の決定） |
| 8 | 戻せる期間 | 5 営業日（その間は毎時書き出し） | 2026-11-26 オーナー（共通の決定） |
| 9 | Cloudflare Access | 置かない（アプリの Google ログインだけ） | 2026-11-26 オーナー（共通の決定） |
| 10 | メールの差出人 | 今と同じアカウント名義 | 2026-11-26 オーナー（共通の決定） |
| 11 | 旧 Google フォーム（部署別 21 本）と 1 分ごとの取り込み | 切り替えで受付を閉じる（直近の回答が 0 件であることを確かめてから）。毎朝の選択肢更新もやめる | 2026-11-26 オーナー |
| 12 | 上席・一般社員の区別 | 作らない（部署で 1 アカウントを共有し、上席も同じアカウントのため分けられない） | 2026-11-26 オーナー |
| 13 | 保守用の処理（台帳の作り直し・マスタ同期・休憩の再計算・フォーム生成など） | 今の管理者一覧（Settings `ADMIN_EMAILS`。imaizumi@・admin-01@）のアカウントだけが使える。`?date=` などの差し込みは防ぐ。作業者・承認者の操作は今と同じ | 2026-11-26 ディレクター判断（#12 の回答を受けて。今の GAS でも管理画面はこの一覧で守っている） |
| 14 | Sebastian の置き場所 | w2403c の PC（オーナー「多分」）。PC が開いているときに確かめる | 2026-11-26 オーナー |
| 15 | Sebastian の付け替え | 切り替えの時点では Sebastian の設定を変えない。GAS の `doPost` を残し、受けた書き込みを VPS の API へ転送する（案）。Sebastian の送り先を直接 VPS に変えるのは、w2403c で確かめた後 | 2026-11-26 ディレクター案（未確定） |
| 16 | 切り替えの日時 | 受け入れテストと JEV ハンズの実機テストが済んだ日の、業務の後の夜。始める前に声をかける | 2026-11-26 オーナー |
| 17 | 前もって任せる範囲 | 声かけなしで進めてよい: 計画・仕様書・試験版の開発と GitHub への push、VPS（`/opt/overtime`）への試験版の配置、Cloudflare（`overtime.lineworks-local.net`・Tunnel `overtime-vps`）と GCP `sebastian01` の OAuth クライアントの登録、w2403c を読むだけで調べること。**本番の切り替え・本番 GAS の変更・本物の利用者へのメールは必ず事前に聞く** | 2026-11-26 オーナー |
| 18 | シートへの反映（決定 A） | 毎晩のタブ置き換えをやめ、VPS の書き込みのたびに該当の行だけをシートへ反映する（行単位の upsert、数秒〜1 分以内）。行の位置と並び・日付の型と表示形式を保ち、行もタブも消さない。夜に全行の照合を行い、差があれば直して記録する。#7・#8 の「書き出し」はこれで満たす（5-5） | 2026-11-26 ディレクター判断（抜け調査 gap-5 高-1・高-2・中-8・中-9、gap-4 F-1〜F-3、gap-3 高-4・中-5） |
| 19 | PDF と蓄積台帳（決定 B） | 今の GAS（同じプロジェクト、DB に紐づいたまま）のコードをそのまま使う。VPS の worker が、行をシートに反映した後で GAS の内部用の入口（`doPost` の新しい action、token 認証）を 1 件ずつ呼ぶ。GAS は今の関数を今と同じ順序・条件で動かし、`pdfFileId` などを返す。PDF の仕事は全体で同時に 1 本（4-6・4-6a・4-7） | 2026-11-26 ディレクター判断（gap-3 高-1〜高-3・中-1〜中-3、gap-1 G2・G3） |
| 20 | 時刻（決定 C） | サーバーの「今日」・年度・締め月・日時の解釈は、すべて Asia/Tokyo。06:30 JST の受け入れ試験を入れる。DB スプレッドシートの時間帯は段階 2 で読んで記録する（4-2・5-5） | 2026-11-26 ディレクター判断（gap-1 G1、gap-5 高-2） |
| 21 | マスタ同期（決定 D） | `syncAllMasters` は VPS の worker が今と同じ規則で行い、DB の同じタブ（`業務NOマスタ` など）にも今と同じ形で書く（日報アプリのため。4-8） | 2026-11-26 ディレクター判断（gap-5 高-3、gap-4 F-6） |
| 22 | 総務の「日次：申請一覧」（本番では表が空。2026-11-26 確認） | 直して表示する。一緒に出る「手動入力」ボタンの差し込みの穴（gap-2 S8）もふさぐ | 2026-11-26 オーナー |
| 23 | Sebastian 経由の残業完了の休憩（休憩マスタの見出し行のずれで常に 0 分） | 移行と一緒に直す。画面から完了したときと同じ計算にそろえる（切り替えの日から実残業の数字が変わる。過去の行は直さない） | 2026-11-26 オーナー |
| 24 | 共有アカウントのログインの持ち時間 | 90 日（朝のピークに一斉のログインし直しを起こさない） | 2026-11-26 オーナー |
| 25 | スマホの表示 | スマホ向けの配置にする（今は PC の画面を縮めた表示）。PC での見た目は今と同じ | 2026-11-26 オーナー |
| 26 | Sebastian がシートへ直接足す申請（`sheet_intake`） | Sebastian（同じ VPS の `/root/.hermes`。こちらからは変えない）は Requests へ Sheets API の `values.append`（USER_ENTERED・INSERT_ROWS）で新しい申請を直接足し、Requests を読んで requestId を探す（既存 2,334 行のうち約 100 行）。切り替えの後は VPS の DB が正なので、worker が 60 秒ごとに Requests の requestId の列を読み、DB に無い行を移行の道具と同じ決まりで読んで申請として取り込む（行番号を記録し、その後の承認などは同じ行へ反映）。GAS と同じく取り込むときに特別なことはしない（WorkLogs の行は実績の記録のときに作る）。requestId の空・重なり・読めない値の行は入れずに管理者の欄に出し、その行へは書かない。VPS が作る新しい行は `values.append`（OVERWRITE）で行を取り、Sebastian の追記と取り合わない。取り込みはシートを読むだけで、本番では GAS が移転モード（切り替えの後）のときだけ動く（実装仕様 6-10） | 2026-11-26 ディレクター判断（Sebastian の追記が見つかったため） |

### 9-2. 未決（オーナーに聞くこと）

- 統合メール通知システムの置き場所と、この DB の読み方（どのタブ・列を、何時に読むか。Settings・BatchLogs も読むか）。w2403c で Sebastian と一緒に確かめる。決定 A で読み先は変わらないので、切り替えの条件ではない。
- Sebastian が requestId をどこから得ているか（シートのどのタブ・列か、何分おきか。gap-5 高-1）、使っている deploymentId、返事の `status`・`pdf`・エラー文を使っているか。決定 A でシートが 1 分以内に最新になるので、読み先は切り替えの条件ではなくなった。deploymentId は切り替えの前に確かめる。
- 申請に使う端末（部署の PC・タブレット・個人のスマホ）と、それぞれの Google ログインの状態（gap-2 S2。Q7 と一緒に）。
- 社内ポータルの差し替え担当、VPS 外のバックアップの置き場、死活監視の通知先。

2026-11-26 計画書の下書きで足した質問（おすすめを先に書く）。決定 A〜D で不要になったものは消し線を引き、理由を添えた。Q16〜Q20 は抜け調査で足した。

| # | 質問 | おすすめ | ほかの選択肢 | 関係する章 |
|---|---|---|---|---|
| ~~Q1~~ | ~~PDF の作り方~~ | — | — | **不要（決定 B、9-1 #19）**: 今の GAS のコードで作り、VPS から内部用の入口を呼ぶ。案 1〜3 は 4-6 に採らなかった案として残した |
| Q2 | Sebastian の残業完了の休憩（今は見出し行のずれで常に 0 分）。履歴では、作者の意図は「画面と同じ」で 0 分は作り漏れ（Sebastian の休日完了は時刻窓方式）。当初の設計では残業の休憩は 0 だった（gap-5 中-5）。該当は `updatedBy='sebastian-agent'` の行（gviz の見当で 97 行、残業と休日の合計）のうちの残業 | 移行では今のまま 0 分（数字を変えない）。直すなら移行の後に別に行う | 画面と同じ時刻窓方式にそろえる／休憩マスタを正しく読む | 3-1・4-5 |
| ~~Q3~~ | ~~Sebastian 経由の PDF の方式~~ | — | — | **不要（決定 A・B）**: シートの行が先に反映され、今の GAS が XLOOKUP 方式で作るので、今と同じに動く（4-5） |
| ~~Q4~~ | ~~Sebastian の返事の `pdf`~~ | — | — | **不要（決定 B）**: 今と同じ GAS で作り、結果を待って今と同じ 4 つの形で返す（4-5）。待つ上限は段階 2 のスパイクで決める |
| ~~Q5~~ | ~~蓄積台帳の朝バッチの癖~~ | — | — | **不要（決定 B）**: 台帳は今の GAS の関数が今と同じ順序・条件で書くので、癖も今のまま残る（4-7） |
| Q6 | 夕方・朝メールの手動送信（メニュー）を VPS に移すか | 移さない（3/10 以降送っていない） | 管理者だけが押せる形で移す | 4-11 |
| Q7 | 共有アカウントのログインの持ち時間。貯蔵品の既定（無操作 12 時間・絶対 14 日）だと、毎週月曜と連休明けの朝 6〜8 時に全部署の端末でログインし直しになる。切り替え翌朝は全端末が初めてのログインになる（gap-2 S2） | 無操作は 7 日以上（使うたびに延びる）、絶対期限は 〇〇 日（30〜90 日の間で）。切り替えの前日に各部署の端末で 1 回開いてもらう | 貯蔵品と同じ（12 時間・14 日） | 4-4・6-2 |
| Q8 | 権限エラー画面のデバッグ情報（管理者のメール一覧など） | 出さない（文言の本文は残す） | 今のまま出す | 3-4 |
| Q9 | `?page=redirect`（旧フォームへ飛ばす URL） | top へ案内する | 今のまま（フォームは閉じるので意味が無い） | 3-4 |
| Q10 | 社内ポータルはこのアプリを iframe で埋め込んでいるか。埋め込んでいるなら、ログイン画面は iframe の中では出せないので、新しいタブか最上位の窓で開く形に変える必要がある（gap-2 S4） | 埋め込みが無ければ、ほかのサイトへの埋め込みを禁止 | 今のまま許す | 4-3 |
| ~~Q11~~ | ~~日報アプリなど、この DB の `業務NOマスタ`・作業員マスタなどのタブを読むアプリはあるか~~ | — | — | **不要（決定 D、9-1 #21）**: VPS が今と同じ規則・同じ形で DB の 4 タブに書き続けるので、読み手がいてもいなくても今と同じ（4-8） |
| Q12 | DeptApprovers・Settings を直してから VPS に効くまでの時間（決定 A〜D では決まらない。GAS の内部用の入口は今と同じくすぐ効く） | 最大 5 分＋「今すぐ取り込む」ボタン | 毎分取り込む | 4-8 |
| Q13 | 完了済みで PDF なしの 25 件。原因の候補は、二次承認・編集の作り直しの失敗、対象日の翌日以外に記録した完了での失敗、ロック待ちの失敗（gap-1 G7・gap-3 G3-M2） | 原因を調べて一覧を見せ、作るかを決める | 今のまま | 6-5 |
| Q14 | 移さないテスト行 | 全部移す（決まった接頭辞が無いため） | 一覧を見て除外する | 5-3 |
| Q15 | `staffSummary` のブラウザのタブの題名（今は「残業・休日出勤申請」と出る見込み） | 今と同じ「残業・休日出勤申請」 | ページの `<title>`「スタッフビュー画面 …」 | [02 §1.5] |
| Q16 | 総務画面の「日次：申請一覧」。2026-11-26 に本番で、申請のある日（11/25）を指定しても表が空と確かめた。原因は返り値に Date が入っていることの見込み。日次表の「手動入力」と、`?page=admin&date=` の自動表示も同じく使えていない見込み（gap-2 S1） | VPS で表示されるように直す。手動入力ボタンの XSS を先に直すのが条件（gap-2 S8） | 今のまま空にする（今と同じ見た目） | 3-3・4-3 |
| Q17 | スマホの表示。今の GAS は viewport を無視しているので、スマホは PC 幅の画面を縮めて表示している。VPS では初めてスマホ向けの配置が効く（gap-2 S7） | スマホ向けの配置を効かせ、日付・時刻の欄に罠 16 の CSS を当てる。現場の機種で実機確認 | 今のまま（PC 幅の縮小表示。viewport の meta を外す） | 4-3 |
| Q18 | 朝バッチは未完了の申請でも PDF を作り、後から完了しても PDF は実績なしのまま。履歴の決定「承認済み＋完了済みでだけ作る」と逆（gap-5 中-4） | 移行では今のまま再現する（件数は 5-4 の照合で数えて見せる） | 朝バッチで未完了を飛ばす（GAS の小さな変更） | 4-9 |
| Q19 | PDF テンプレの、直接方式で書かないセル（終了時刻・明細の 2〜3 行目・業務内容など）の数式。DB の `操作!B3` の申請の値が出ている恐れがある（gap-3 G3-H4） | スパイクで本番の PDF 3 枚を見てから決める。別の申請の値が出ていれば、書かないセルに空を書いて数式を消す（GAS の小さな変更） | 今のまま | 4-6 |
| Q20 | PDF の作成が失敗したときの自動の再試行。今は 1 回だけ試し、失敗すると PDF なしのまま（gap-1 G7・gap-3 G3-M2） | 再試行する（今なら PDF なしで終わる申請に後から PDF ができ、台帳も書かれる）。前の PDF を捨ててから作る順は今のまま | 今と同じく 1 回だけ試す | 4-6 |
