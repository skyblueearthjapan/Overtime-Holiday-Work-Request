# 02 画面の棚卸し（移行の受け入れチェックリスト）

対象：`src/top.html`、`src/approver.html`、`src/approve2.html`、`src/admin.html`、`src/staffSummary.html`、`src/no_auth.html`、`src/adminApi.gs` の `doGet` / `include_` / `handleFormRedirect_`（adminApi.gs:534-657）。

表記のきまり：

- 「確認済み」…コードを読んで確かめたこと（ファイル:行 を付けます）。
- 「推測」…GAS の仕様や実行時の挙動から考えたことで、コードだけでは確かめられないこと。
- 行番号は 2026-11-26 時点の `src/` のものです。

---

## 0. 先に結論

- 画面は 5 ページ＋権限エラー画面＋リダイレクト用の簡易応答 3 種です（確認済み、adminApi.gs:534-657）。
- `google.script.run` の呼び出しは **40 か所、サーバー関数は 30 種類** です（確認済み）。うち `api_getFormUrl` は top.html の使われていないコード（デッドコード）からしか呼ばれません。
- 作業員の申請は、TOP 画面のモーダル（2 ステップ）から `api_submitRequest(payload)` を直接呼びます。Google フォームは開きません（確認済み、top.html:733-751, 1959-2030）。Google フォームを開く旧経路のコードは残っていますが、どこからも呼ばれません。`?page=redirect` だけがサーバー側でフォームへ飛ばす旧経路として生きています。
- GAS 依存は `google.script.run`、テンプレートの `<?= ?>` / `<?!= ?>`、`setXFrameOptionsMode(ALLOWALL)`、`<base target="_top">`、`Session.getActiveUser()`、`ScriptApp.getService().getUrl()` です。`google.script.host` / `google.script.url` は使っていません（確認済み）。

---

## 1. 画面の一覧

### 1.1 ルーティング（doGet）

| 項目 | 内容 | 根拠 |
|---|---|---|
| 既定ページ | `page` 未指定なら `top` | 確認済み adminApi.gs:535 |
| ホワイトリスト | `top / approver / admin / approve2 / staffSummary`。それ以外は `top` に落とす | 確認済み adminApi.gs:544-545 |
| 特別ページ | `page=redirect` は `handleFormRedirect_` へ（ホワイトリストより前に判定） | 確認済み adminApi.gs:539-541 |
| タイトル | 5 ページとも `setTitle('残業・休日出勤申請')`。権限エラーは `'権限エラー'` | 確認済み adminApi.gs:561, 580, 596 |
| XFrameOptions | 全応答で `XFrameOptionsMode.ALLOWALL`（他サイトの iframe に埋め込み可） | 確認済み adminApi.gs:562, 581, 597, 616, 636, 649, 656 |
| 実行ユーザー / 公開範囲 | `executeAs: USER_DEPLOYING`、`access: ANYONE_ANONYMOUS` | 確認済み src/appsscript.json |
| `include_(name)` | 定義だけあり、どの HTML からも呼ばれていない | 確認済み adminApi.gs:601-603、HTML 内に `include_` なし |

注意（推測）：`ANYONE_ANONYMOUS` かつ `USER_DEPLOYING` なので、`Session.getActiveUser().getEmail()` は、同じ Workspace ドメイン（lineworks-local.info）にログインしている人でないと空文字になります。画面の権限判定と「本人の自動選択」はこのメールアドレス頼みです。

### 1.2 テンプレート変数（doGet が差し込む値）

| 変数 | 値 | 使っているページと使い方 | 根拠 |
|---|---|---|---|
| `APP_URL` | `ScriptApp.getService().getUrl()`（/exec の URL） | 全ページのヘッダーリンク `<?= APP_URL ?>?page=...`、no_auth の「トップ画面へ戻る」 | 確認済み adminApi.gs:536, 588 |
| `PORTAL_URL` | 社内ポータル（別 GAS）の固定 URL `https://script.google.com/a/macros/lineworks-local.info/s/AKfycbx2ey.../exec` | 全ページの「社内ポータルへ移動」ボタン | 確認済み adminApi.gs:554, 578, 589 |
| `DATE_PARAM` | `?date=` の値（そのまま、検証なし） | approve2：`var urlDate = '<?= DATE_PARAM ?>'` で、`^\d{4}-\d{2}-\d{2}$` に合えば対象日に使う。admin：`URL_DATE_PARAM` を「日」欄に入れ、値があれば日次一覧を自動で読む（形式チェックなし） | 確認済み adminApi.gs:590、approve2.html:256-261、admin.html:466-469, 1030-1032 |
| `REASONS_JSON` | `JSON.stringify(REASONS)`（6 項目、最後が「その他（→補足理由必須）」） | top：`var REASONS = <?!= REASONS_JSON ?>;`（無エスケープ出力）→ 理由プルダウン | 確認済み adminApi.gs:591、top.html:637、config.gs:42-49 |
| `OT_HOURS_JSON` | `['0.5','1.0',...,'4.0']` | top：残業の予定時間プルダウンと編集モーダル（表示は「h 時間」） | 確認済み adminApi.gs:592、top.html:638、config.gs:51 |
| `HD_HOURS_JSON` | `['半日','1日']` | top：休日の予定時間プルダウンと編集モーダル | 確認済み adminApi.gs:593、top.html:639、config.gs:52 |
| `message` | 権限エラーの文（改行 `\n` とデバッグ情報を含む） | no_auth：`<div class="msg"><?= message ?></div>` | 確認済み adminApi.gs:555-559, 579、no_auth.html:73 |

- 3 つの JSON 変数と `DATE_PARAM` は全ページに渡されますが、使うのは上の表のページだけです（確認済み）。
- no_auth の `message` には `\n` が入っていますが、`.msg` に `white-space:pre` がないため、画面では 1 行につながって表示されます（確認済み no_auth.html:50-55。表示結果は推測）。移行先でも同じ見え方にするか決める必要があります。

### 1.3 ページごとの開ける人と権限チェック

| ページ | 開ける人 | 権限チェック（doGet） | 読む URL パラメータ | 根拠 |
|---|---|---|---|---|
| `?page=top`（既定） | 誰でも | なし | なし | 確認済み adminApi.gs:587-597 |
| `?page=approver` | 管理者、または DeptApprovers / ApproverMap のどこかの部署の承認者 | メールが取れたときだけ判定。**メールが空なら判定せず通す** | なし | 確認済み adminApi.gs:567-585（`if (email)` で囲まれている） |
| `?page=admin` | 管理者（ApproverMap の role=admin、または Settings の ADMIN_EMAILS） | `isAdminWithDebug_(email)`。NG なら no_auth（デバッグ情報つき） | `date` | 確認済み adminApi.gs:548-564 |
| `?page=approve2` | admin と同じ | admin と同じ | `date` | 確認済み adminApi.gs:548-564 |
| `?page=staffSummary` | 誰でも | なし（どの部署・どの人の集計も見られる設計） | なし | 確認済み adminApi.gs:587、staffSummary.html:234 の案内文 |
| no_auth | 上の権限 NG のとき | — | — | 確認済み adminApi.gs:552-562, 576-581 |
| `?page=redirect` | 誰でも | なし | `type`（overtime/holiday）、`dept` | 確認済み adminApi.gs:611-657 |

- admin / approve2 の権限エラー文には、`isAdminWithDebug_` のデバッグログ（`adminResult.debug`）と、現在のアカウントのメールアドレスがそのまま出ます（確認済み adminApi.gs:555-559）。インターネットに公開すると内部情報が見えるので、移行時に出すかどうか決める必要があります。
- サーバー関数側でも権限を見ているかは、この文書の範囲外です（別の棚卸し担当の範囲）。

### 1.4 `?page=redirect`（handleFormRedirect_）

- `type` が `overtime` / `holiday` 以外なら `<p>エラー: type パラメータが不正です（overtime / holiday）</p>` を返します（確認済み adminApi.gs:612-617）。
- `dept` がなければ `api_getWorkerInfo()` で本人の部署を探します（確認済み adminApi.gs:620-628）。
- 部署が決まらなければ「部署が指定されていません」画面（タイトル `部署未指定`、トップへ戻るリンク）を返します（確認済み adminApi.gs:630-637）。
- `api_getFormUrl(type, dept)` でフォーム URL を取り、なければ「〇〇申請フォームを準備できませんでした」画面（タイトル `フォーム未設定`）です（確認済み adminApi.gs:640-650）。
- 取れたら `<script>window.top.location.href="<formUrl>";</script>` と手動リンクを返します（確認済み adminApi.gs:653-656）。`formUrl` はエスケープなしで埋め込みます。
- 現在の画面からこの URL へのリンクはありません（確認済み、HTML に `page=redirect` なし）。外部（ポータルやブックマーク）から使われているかは不明です（推測。要確認）。

### 1.5 ページ間リンクと外部リンク

| ページ | ヘッダーのリンク（すべて `target="_top"`） | 根拠 |
|---|---|---|
| top | スタッフ集計 / 承認者画面 / 総務部画面 / 社内ポータルへ移動 | 確認済み top.html:315-318 |
| approver | トップへ / 総務部画面 / 社内ポータルへ移動 / 「承認者変更」ボタン（管理者のみ表示）。承認者選択モーダル内に「申請TOPへ戻る」 | 確認済み approver.html:197-200, 381 |
| approve2 | 総務部画面 / 承認者画面 / トップへ / 社内ポータルへ移動 | 確認済み approve2.html:162-165 |
| admin | トップへ / 承認者画面 / 二次承認 / 社内ポータルへ移動 | 確認済み admin.html:175-178 |
| staffSummary | トップへ / 社内ポータルへ移動 | 確認済み staffSummary.html:201-202 |
| no_auth | トップ画面へ戻る（`?page=top`） / 社内ポータルへ移動 | 確認済み no_auth.html:74-76 |

- `<base target="_top">` は top / approver / approve2 / admin / no_auth にあります。staffSummary にはありません（各リンクに `target="_top"` を個別指定）（確認済み top.html:4、approver.html:4、approve2.html:4、admin.html:4、no_auth.html:4、staffSummary.html:1-6）。
- staffSummary だけ `<title>スタッフビュー画面 — 月次 残業・休日出勤 集計</title>` と `lang="ja"` を持ちます（確認済み staffSummary.html:2, 6）。GAS では `setTitle` が外枠のタイトルになるため、ブラウザのタブは「残業・休日出勤申請」になります（推測）。移行先でどちらのタイトルにするか決める必要があります。
- Drive の PDF リンク：`https://drive.google.com/file/d/<pdfFileId>/view` を `target="_blank"` で開きます（確認済み top.html:1150-1151、approve2.html:341、admin.html:612、staffSummary.html:338）。
- メールから来るリンク（参考、画面の外）：`?page=approve2&date=...`（mail.gs:314）、`?page=top`、`?page=admin&date=...`（mail.gs:1372-1373）。移行後も同じ URL 形式で開ける必要があります（確認済み）。

---

## 2. ページごとの中身

### 2.1 top.html（申請TOP）

ヘッダー：タイトル「残業・休日出勤 申請TOP」、サブ「全社の本日残業 / 休日出勤申請 一覧」（sticky、半透明＋blur）（確認済み top.html:309-320）。

申請カード（top.html:325-352）：

- 「部署：」プルダウン `#deptSelect`（初期「読込中...」）と「作業員：」プルダウン `#workerSelect`（初期「-- 部署を先に選択 --」）。
- 「残業申請」（青）と「休日出勤申請」（緑）の 2 ボタン。`openForm('overtime'|'holiday')`。
- 注意文（固定）：「4時30分までに申請を出さなかった場合、申請が下りない場合がございます。」

本日の残業申請カード（top.html:355-381）：

- 見出し「本日の残業申請（`res.today`）」＋件数「N名」、右に「更新」ボタン（`loadDashboard()`）。
- 列：部署 / 氏名 / 承認 / 予定 / 実績 / 休憩 / 実残業 / PDF / 操作（9 列）。
- 期限超過バナー：`isOverdue` の件数が 1 以上のとき、表の上に「作業完了が未記録の申請があります（N件）。手入力で記録してください。」を JS で差し込みます（確認済み top.html:1030-1047）。
- 氏名欄：期限超過なら赤い日付ピル（`targetDateLabel`）を付けます。行の背景は薄赤です（確認済み top.html:1100-1107）。

休日出勤申請カード（top.html:384-410）：

- 見出し「休日出勤申請（未承認＋承認済み）」＋件数。更新ボタンはありません。
- 列：日付 / 部署 / 氏名 / 承認 / 予定 / 実績 / 休憩 / 実働 / PDF / 操作（10 列）。
- 期限超過バナーは残業と同じ文言です（確認済み top.html:1062-1079）。

表示のルール（確認済み）：

- 承認列：`submitted` → 「承認待ち」（黄）、それ以外 → 「承認済」（緑）（top.html:1092-1094）。
- 実績・休憩・実残業は 0 なら「-」、実残業は太字（top.html:1113-1115）。
- PDF 列：`pdfFileId` があれば「PDF」リンク、なければ「承認済み かつ net>0」で「未作成」、それ以外「-」（top.html:1148-1156）。
- 操作列（top.html:1191-1247）：
  - 残業：完了済みなら「完了済」（未承認なら「承認待ち」も）。未完了なら［残業完了］［手入力］［編集］［削除］と、4h 超過バッジ、手入力欄（終了時刻プルダウン、期限超過なら既定 20:00）。
  - 休日・開始済み：［休日完了］［手入力］［編集］［削除］、手入力欄（終了時刻、期限超過なら既定 17:00）。
  - 休日・未開始：期限超過なら「⚠ 未開始」、［開始］［手入力］［編集］［削除］、手入力欄（開始時刻、期限超過なら既定 08:00）。
  - 未承認でも完了・開始ボタンは押せます（「承認待ち」ピルが前に付くだけ）。
- 4h 超過バッジ（top.html:1159-1182）：承認済みで未完了のとき。残業は対象日 17:20（+09:00）から、休日は `actualStartAt` から、240 分を超えたら「4h超過」。画面を開いた時点で一度だけ判定します（自動更新なし）。
- 時刻プルダウン（top.html:1268-1286）：10 分刻み。`ot-end` は 17:30〜翌 03:00（27:00 → 「03:00」表示）、`hd-start` は 06:00〜18:00、`hd-end` は 08:00〜23:00。

モーダル：

| モーダル | 内容 | 根拠 |
|---|---|---|
| 申請モーダル（2 ステップ、`#requestModalWrap`） | 4 章で詳しく書きます | 確認済み top.html:463-583 |
| 編集モーダル（`#editModalWrap`） | 対象者（部署 ／ 氏名）、予定時間（必須）、作業実施日（休日のみ、`type=date`）。［キャンセル］［保存］ | 確認済み top.html:607-634, 1435-1522 |
| 工番選択モーダル（`#orderPickerWrap`） | **デッドコード**（開く関数 `showOrderPicker` の呼び出し元が死んでいる） | 確認済み top.html:415-460, 814 |
| 休日日付選択モーダル（`#datePickerWrap`） | **デッドコード**（`showDatePicker` はどこからも呼ばれない） | 確認済み top.html:586-604, 757 |

- 編集モーダルの現在値：残業は `(approvedMinutes/60).toFixed(1)` を選択、休日は 480 分以上なら「1日」、それ以外「半日」（確認済み top.html:1460-1465）。
- オーバーレイのクリックで、申請・編集モーダルは閉じます（確認済み top.html:464, 608）。

### 2.2 approver.html（承認者画面）

ヘッダー：「承認者画面」、サブ `#approverSub`（「〇〇 の担当（部署, 部署）」＋管理者なら「  [管理者]」）（確認済み approver.html:191-202, 553-556）。

起動の流れ（確認済み approver.html:407-487）：

- まず承認者選択モーダルを出し「承認者情報を確認中...」。`api_getApproverDepts()` を呼びます。
- 承認者が 0 件なら「承認可能な部署がありません。」で止まります（モーダルは閉じない）。
- 管理者でなければモーダルを出さずに、`currentEmail` と一致する承認者（なければ先頭の承認者）の担当部署で表示します。
- 管理者なら「全部署（一括表示）」ボタン＋承認者ごとのボタン（名前＋担当部署）＋注記「管理者: 承認者を選択するか、全部署を一括表示できます」を出します。
- モーダルには閉じるボタンがなく、オーバーレイをクリックしても閉じません。選ぶか「申請TOPへ戻る」しかありません（確認済み approver.html:372-384）。

KPI カード 3 枚（確認済み approver.html:207-220, 573-582）：申請件数（残業+休日）／未承認（黄）／承認済（緑）。

- 初回描画は `submitted` だけを未承認に数えます。承認・却下後の再計算（`updateKpiCounters`）は `submitted` と `modified` を未承認に数えます（確認済み approver.html:577, 775）。数え方が少し違います。

残業カード（approver.html:223-250）：見出し「残業申請（本日分＋未承認）（日付）」＋件数、［残業 全件承認］［更新］。
休日カード（approver.html:253-277）：見出し「休日出勤申請（未承認＋当日承認済み）」＋件数、［休日 全件承認］。

- 列は両方とも：日付 / 部署 / 氏名 / 承認 / 予定時間 / 提出 / 操作（7 列）。
- 承認列：`submitted` → 「未承認」、`canceled` → 「却下」、それ以外「承認済」（approver.html:616-618）。
- 日付列：期限超過なら赤字＋⚠（approver.html:651-657）。
- 氏名列：期限超過なら「⚠ 未処理」、過去申請なら「🔙 過去申請」ピル＋「(N日前)」（approver.html:608-613, 623-628）。
- 行の背景：期限超過は薄赤、過去申請は薄黄（過去申請が優先）（approver.html:632-638）。
- 提出列：当日なら `HH:mm`、別の日なら `M/D HH:mm`（approver.html:722-730、5 章参照）。
- 操作列：未承認なら［承認］［却下］、却下済みなら「却下済」、それ以外「承認済」（approver.html:732-743）。

月次サマリー（折りたたみ、approver.html:280-367, 903-1175）：

- 見出し「月次：個人別サマリー」「実績・予測・40h/60h監視」。クリックで開閉。初めて開いたときだけ自動で読み込みます。
- フィルタ：月（`type=month`、初期値は今月）、氏名検索（部分一致）、チップ［40h超］［60h超］［予測40h超］［PDF未作成］、［更新］。月を変えても自動では読み直さず、［更新］が必要です（確認済み、`#approverYm` に onchange なし）。
- ミニ KPI：対象人数 / 40h超（実績/予測） / PDF未作成。
- 棒グラフ（canvas 自前描画、`#cApproverPeople`、高さ 260px）。
- 個人別一覧：氏名 / 合算(分) / 合算 / 残業 / 休日 / 40残 / 予測 / PDF。見出しクリックで並べ替え（同じ列なら昇順⇔降順、別の列なら降順から）。初期は `totalNet` 降順。
- 注意対象（30h超 or 予測40h超）：氏名 / 実績 / 予測 / 40残 / 60残 / PDF。並べ替えなし。
- 40残の色：0 で「超過」（赤）、300 分以下で黄、それ以外は無色。予測：3600 分以上「60超予測」、2400 分以上「40超予測」、それ以外「範囲内」。
- 承認者を選び直すと、月次の状態をリセットして折りたたみます（approver.html:503-509）。

### 2.3 approve2.html（二次承認）

- ヘッダー「二次承認」、サブ「総務部長による二次承認（一次承認済みの申請を確認・承認）」（確認済み approve2.html:156-167）。
- 対象日（`type=date`）＋［読込］＋「読み込み中...」表示。初期値は `?date=`（形式が合うとき）または今日。開いたらすぐ読み込みます（確認済み approve2.html:172-181, 254-264）。
- KPI 4 枚：表示件数 / 承認可能（一次承認済） / 一次未承認 / 期限超過（approve2.html:184-201, 285-294）。
- 一括バー：［全選択］チェック、「N件選択中」、［一括二次承認］（0 件なら無効）（approve2.html:205-209）。
- 表の列：（チェック） / 部署 / 氏名 / 対象日 / 種別 / 予定時間 / 一次承認 / 二次承認 / PDF / 操作（10 列）（approve2.html:210-225）。
  - チェックボックスは「一次承認済み かつ 二次未承認」の行だけ。
  - 一次未承認の行は薄く（`opacity .5`）、二次承認済みは `opacity .6`。期限超過は薄赤背景＋対象日に「⚠」。
  - 操作：二次承認済み「承認済」、可能なら［二次承認］、それ以外「一次承認待ち」。
- 注記：「※二次承認を行うとPDFが再生成されます（翌朝バッチで二次承認印付きPDFが作成されます）。」（approve2.html:227-229）。
- 単票の［二次承認］は **confirm なし** で即実行します（確認済み approve2.html:395-423）。一括は confirm あり。

### 2.4 admin.html（総務部管理）

ヘッダー「総務部管理」、サブ「月次40h監視（実績net）／予測（ペース）／部署配分／年度推移／PDF未作成抽出」（確認済み admin.html:169-180）。

フィルタカード（admin.html:185-222）：

- 月（yyyy-mm）：**ただのテキスト入力**（`type=month` ではない）。初期値は今月（暦月）。
- 部署：`api_adminDeptOptions()` の結果（先頭 `ALL` を「全体」と表示）。
- 検索（氏名、placeholder「例：今泉」）、チップ 4 つ（実績40h超のみ / 実績60h超のみ / 予測40h超のみ / PDF未作成のみ）。
- ［月次を更新］［表示中をCSV］、右側に日（yyyy-mm-dd、テキスト入力、初期値は `?date=` か今日）と［日次を表示］。
- 注記「※実績は WorkLogs.netMinutes を使用（残業＋休日を合算）。予測は「今月累計÷経過日×月日数」。」
- フィルタ（検索文字・チップ）を変えても自動では描き直しません。［月次を更新］でサーバーから取り直したときと、列の並べ替えのときに反映されます（確認済み、`#q` や `#fOver40` にイベントなし。approver とは違う挙動）。部署の変更も自動では読み直しません。

KPI 4 枚（admin.html:227-248, 520-529）：実績40h超（サブ「対象人数：N / 経過：d/D日」）、実績60h超（サブ「実績60h超は特別条項レベル」）、予測40h超（サブ「予測は"ペース換算"」）、PDF未作成（サブ「未作成は監査上の要注意」）。

グラフ（canvas 自前描画、外部ライブラリなし）：

| canvas | 内容 | 根拠 |
|---|---|---|
| `#cPeople` | 個人別の実績（濃）＋予測（薄）の棒、40h/60h の点線。目盛 0/40h/60h/最大。41 名以上は名前ラベルを省略 | 確認済み admin.html:259, 777-854 |
| `#cDept` | 部署別配分の円グラフ。色は部署名のハッシュから決める。凡例は先頭 10 部署＋% | 確認済み admin.html:268, 856-894 |
| `#cFiscal` | 年度推移（4/1〜3/31）の折れ線。2400 分の目盛線。x ラベルは偶数番目だけ `lab.slice(5)` | 確認済み admin.html:283, 896-953 |

- canvas の高さは CSS で 280px（`.small` は 220px）。`devicePixelRatio` で解像度を合わせます。描き直しは読み込み時だけで、ウィンドウの大きさを変えても描き直しません（確認済み admin.html:134-135, 766-775。`resize` リスナーなし）。

表：

| 表 | 列 | 根拠 |
|---|---|---|
| 月次：個人別 `#tblMonthly` | 部署 / 氏名 / 合算(分) / 合算 / 残業 / 休日 / 40まで残 / 予測（月末） / PDF未作成。見出しクリックで並べ替え（全 `th` に `cursor:pointer`） | 確認済み admin.html:295-310, 720-732 |
| 注意対象（DX抽出） `#tblWatch` | 部署 / 氏名 / 実績 / 予測 / 40残 / 60残 / PDF | 確認済み admin.html:321-331 |
| 日次 `#tblDaily` | 部署 / 氏名（過去申請バッジ） / 種別 / 承認 / 実績 / 休憩 / 実残業/実働 / PDF / 操作 | 確認済み admin.html:342-351, 602-652 |
| 特別条項（60h超）年度内回数 `#tblSpecial` | 部署 / 氏名 / 60h超回数 / 状況（6 回到達＝赤、5 回＝黄、それ以外「N回」）。0 件なら「60h超の月はありません」 | 確認済み admin.html:357-377, 962-988 |
| 未承認滞留監視 `#tblPending` | 対象日（過去日は赤字⚠） / 種別 / 部署 / 氏名 / 経過時間（N時間） / 状態（48h超＝赤、それ以外「滞留中」）。0 件なら「未承認の滞留はありません」 | 確認済み admin.html:382-404, 991-1022 |

- 日次の操作列：承認済みで未完了なら［手動入力］、完了済みなら「完了」（admin.html:616-623）。日次は 0 件でも空表で、「データなし」表示はありません。
- 月次・注意対象も 0 件時の案内文はありません（approver とは違う）。
- 特別条項と年度推移は「基準日」（`#fyBase`、テキスト入力、初期値は今日）を使います。特別条項の基準日は年度推移カードの入力欄を共用しています（admin.html:279, 963）。

手動実績入力モーダル（admin.html:409-433, 657-703）：

- 氏名 / 種別 / 日付の表示、開始時刻・終了時刻（`type=time`）。
- 開始時刻の初期値：残業は 17:20、休日は `actualStartAt` の `HH:mm`、なければ空。
- 終了・開始の両方が必須。オーバーレイのクリックで閉じます。

CSV 出力（admin.html:735-763）：

| 項目 | 内容 |
|---|---|
| 対象 | 画面に出ている行（フィルタ・並べ替え後の `STATE.visiblePeople`） |
| 列 | 部署, 氏名, 合算(分), 合算, 残業, 休日, 40残(分), 予測(分), PDF未作成 |
| 文字コード | UTF-8、**BOM なし**、`Blob type: 'text/csv'`（charset 指定なし） |
| 改行 | `\n` |
| クォート | `"` `,` 改行を含むときだけ `"..."` で囲み、`"` は `""` |
| ファイル名 | `総務_月次_<月入力欄の値>_<部署値>.csv`（全体なら `ALL`） |
| 注意 | BOM がないため、日本語 Windows の Excel でダブルクリックすると文字化けします（推測。現行も同じ挙動のはず）。`URL.revokeObjectURL` をクリック直後に呼んでいます |

### 2.5 staffSummary.html（スタッフビュー画面）

- ヘッダー：ロゴ「残」、「スタッフビュー画面」「残業・休日出勤 月次集計（社内）」、右に「ログイン：氏名（部署）」（確認済み staffSummary.html:190-205, 480）。
- 検索カード：対象月（締め）の ‹ › ボタンと「YYYY年 M月分」「締め期間 M/16 〜 M'/15」、部署（「全部署」）、スタッフ（「全員」）。右上に期間表示（確認済み staffSummary.html:208-235）。
  - 締め月の初期値：今日が 15 日以前なら前月、16 日以降なら今月（確認済み staffSummary.html:248-252）。
  - 未来の締め月へは進めません（› は無反応）（確認済み staffSummary.html:258, 469）。
  - 案内文「🔎 どの部署・どのスタッフでも検索できます。…」
- 結果カードは 2 種類（`res.mode`）：
  - 個人ビュー（`personal`、staffSummary.html:322-393）：イニシャルのアバター、氏名・部署・月ラベル、状態ピル（40時間内 / 40時間 超（年6回のうち） / 60時間 超過）。KPI 4 枚（合計（実働） / 残業 / 休日出勤 / 見込み（実績待ち/申請中））。上限ゲージ（最大 4200 分、40h・60h の線、現在値の ▲ 表示、3 色の背景帯、凡例）。年度の 40 時間超回数（● 6 個、残り回数、0 なら赤）。明細表：対象日 / 種別 / 承認 / 実績(net) / 休憩 / 状態（承認済 / 申請中） / 個票（PDF リンク）。
  - 一覧ビュー（staffSummary.html:395-437）：見出し「スタッフ別 一覧 — 部署」と「N名 ・ 月ラベル」。列：氏名 / 部署、残業、休日出勤、合計（実働）▼（ミニ棒つき）、見込み、状況（60時間超 / 40時間超 / 上限内＋「年N/6」）。合計行（tfoot）あり。行クリックでその人の個人ビューへ切り替え、ページ先頭へスムーズスクロールします。
  - 並べ替えボタンはありません（「▼」は表示だけ）。順番はサーバーの返した順です（確認済み）。
- 起動の流れ：部署一覧 → 本人情報 → 本人の部署の作業員一覧 → 本人を選んで描画（確認済み staffSummary.html:474-501）。
- 描画のたびにフェードのアニメーション（`prefers-reduced-motion` では無効）（確認済み staffSummary.html:172-174, 309）。
- CSS の `color-mix()` を使っています（確認済み staffSummary.html:150-152, 170）。古いブラウザでは行のホバー色などが出ません（推測）。

### 2.6 no_auth.html（権限エラー）

- 🔒 アイコン、「アクセス権限がありません」、`message`、［トップ画面へ戻る］［社内ポータルへ移動］（確認済み no_auth.html:70-78）。

### 2.7 全ページ共通

| 項目 | 内容 | 根拠 |
|---|---|---|
| localStorage / sessionStorage | 使っていない | 確認済み（grep で 0 件） |
| 自動更新タイマー | なし。`setInterval` は top の「本人の作業員自動選択」用だけ（200ms ごと、5 秒で打ち切り） | 確認済み top.html:670-681 |
| 外部 CDN・Web フォント | なし。フォントは system-ui 系の指定だけ。staffSummary の select の矢印は data URI の SVG | 確認済み（`<script src` `<link` なし）、staffSummary.html:75 |
| 画面幅の対応 | top / approver は 640px 以下、approve2 / admin は 768px と 480px、staffSummary は 820px と 520px で 1 列化など | 確認済み top.html:297-305、approver.html:178-187、approve2.html:141-152、admin.html:151-165、staffSummary.html:176-184 |
| top のスマホ特例 | `.card div[style*="grid-template-columns:1fr 1fr"]` を 1 列にする（申請ボタン 2 つが縦並びになる） | 確認済み top.html:304 |
| console 出力 | top / approver はダッシュボードの結果を `console.log(JSON.stringify(res))` で出す | 確認済み top.html:997、approver.html:536 |

### 2.8 confirm() / alert() の文言一覧

（`\n` は改行。`<...>` は差し込み値。）

top.html：

| 種類 | 文言 | 行 |
|---|---|---|
| alert | 部署を選択してください。 | 736 |
| alert | 作業員を選択してください。 | 746 |
| alert | 日付を選択してください。 | 802（デッド）, 1820 |
| alert | 工番は最大3件まで選択できます。\n不要な工番を×で削除してください。 | 898（デッド）, 1890 |
| alert | `<残業/休日出勤>`申請フォームを作成できませんでした。管理者に連絡してください。 / `<label>`フォーム取得エラー: `<msg>` | 963, 968（デッド） |
| confirm | 残業完了を記録しますか？\n（17:20から現在時刻まで自動計算されます） | 1290 |
| alert | 残業完了を記録しました。\n実働: `<>`\n休憩控除: `<>`\n実残業: `<>` | 1303, 1329 |
| alert | エラー: `<msg>` | 1307, 1333, 1349, 1372, 1397, 1423, 1519 |
| alert | 終了時刻を入力してください。 / 開始時刻を入力してください。 | 1316, 1406 / 1358 |
| confirm | 終了時刻 `<HH:mm>` で残業完了を記録しますか？ | 1317 |
| confirm | 休日出勤を開始しますか？\n（現在時刻が開始時刻として記録されます） | 1341 |
| confirm | 開始時刻 `<HH:mm>` で休日出勤開始を記録しますか？ | 1359 |
| confirm | 休日出勤を完了しますか？\n（開始時刻から現在時刻まで自動計算されます） | 1380 |
| alert | 休日出勤を完了しました。\n実働: `<>`\n休憩控除: `<>`\n実残業/実働: `<>` | 1393, 1419 |
| confirm | 終了時刻 `<HH:mm>` で休日出勤完了を記録しますか？ | 1407 |
| alert | 申請データが見つかりません。画面を更新してください。 | 1439 |
| alert | 予定時間を選択してください。 | 1490, 1971 |
| confirm | 申請内容を変更しますか？\n予定時間: `<>`（休日で日付ありなら \n作業実施日: `<>`） | 1498-1500 |
| alert | 申請を更新しました。 / 更新に失敗しました。 | 1510 / 1513 |
| confirm | `<氏名>` の`<残業/休日出勤>`申請を削除しますか？\n\nこの操作は取り消せません。（見つからなければ「選択した申請を削除しますか？…」） | 1529-1530 |
| alert | 申請を削除しました。 / 削除に失敗しました。 / 削除エラー: `<msg>` | 1535 / 1538 / 1542 |
| alert | 作業実施日を選択してください。 | 1923 |
| alert | 理由を選択してください。 | 1966 |
| alert | 理由が「その他」の場合、補足理由を入力してください。 | 1968 |
| confirm | `<label>`申請を送信しますか？\n\n日付: `<日付>`（過去なら「（過去申請）」）\n理由: `<>`\n予定時間: `<>`（過去なら \n\n⚠ 過去日の申請です。承認者・総務への事前連絡はお済みでしょうか？） | 1979-1986 |
| alert | `<label>`申請が完了しました。 | 2014 |
| alert | `<res.error>` または「申請に失敗しました。管理者に連絡してください。」 | 2021-2022 |
| alert | 送信エラー: `<msg>` | 2027 |

approver.html：

| 種類 | 文言 | 行 |
|---|---|---|
| confirm | この申請を承認しますか？\n\n`<M/D(曜) 残業>`\n`<部署> <氏名>（予定 <時間>）` | 786, 674-684 |
| alert | エラー: `<msg>` | 796 |
| confirm | この申請を却下しますか？\n\n`<同上>` | 804 |
| alert | エラー: `<res.message or 不明>` / 却下エラー: `<msg>` | 812 / 817 |
| alert | 未承認の申請はありません。 | 853 |
| confirm | `<残業/休日出勤> N件を一括承認しますか？`\n\n・`<日付>`：N件 …（過去日あれば \n※うち N件は過去日の承認漏れです。）\n\n（完了済みの申請がある場合、PDF生成に時間がかかることがあります） | 865-866 |
| alert | 一括承認完了（エラーN件）\n`<requestId: error>`… / 一括承認エラー: `<msg>` | 890 / 896 |
| alert | 月次サマリー取得エラー: `<msg>` | 955 |

approve2.html：`日付を選択してください`（268）、`<err.message>`（279, 419, 452）、`<res.message or エラーが発生しました>`（414）、confirm `N件を一括二次承認しますか？`（428）、`一部エラー: N件が失敗しました`（442）、`エラーが発生しました`（446）。

admin.html：`<err.message>`（482, 598, 650, 715, 986, 1020）、`終了時刻を入力してください。`（687）、`開始時刻を入力してください。`（691）、confirm `手動実績を登録しますか？\n開始: <> → 終了: <>`（694）、`登録完了\n実働: <>\n休憩控除: <>\n実残業/実働: <>`（698）、`エラー: <msg>`（701）。

staffSummary.html：alert / confirm なし（確認済み）。

---

## 3. google.script.run の呼び出し一覧（40 か所・30 関数）

「楽観」＝サーバーの返事を待たずに画面を先に変えるもの。

### 3.1 top.html（18 か所）

| # | 関数（引数） | 呼ぶ場面 | 成功時 | 失敗時 | 読み込み表示 / 楽観 | 行 |
|---|---|---|---|---|---|---|
| T1 | `api_getDeptList()` | 起動時 | 部署プルダウンを作り、続けて T2 | 何もしない（「読込中...」のまま） | — | 653-688 |
| T2 | `api_getWorkerInfo()` | T1 成功後 | `worker.dept` を選び `onDeptChange()`。作業員の自動選択を 200ms ごとに試す（後述の不具合あり） | 何もしない | — | 664-685 |
| T3 | `api_getWorkersByDept(dept)` | 部署変更時 | 「-- 作業員を選択 --」＋作業員（value=先頭トークンの作業員コード、表示=文字列全体） | 「取得エラー」 | 「読込中...」 | 709-723 |
| T4 | `api_getHolidayCandidateDates()` | **デッド**（showDatePicker） | 候補日ボタン | 「取得エラー: 」 | 「候補日を取得中...」 | 766-792 |
| T5 | `api_getOrderList()` | **デッド**（showOrderPicker） | `ORDER_CACHE` に入れ検索 | 「読込エラー: 」 | 「工番マスタを読込中...」 | 828-837 |
| T6 | `api_getFormUrl(type, dept, workerLabel, targetDateStr, orderCodes)` | **デッド**（doOpenForm） | 先に開いた空タブを Google フォーム URL へ。ポップアップが止められたらリンク表示 | タブを閉じ alert | 空タブに「申請フォームを準備中...」 | 940-971 |
| T7 | `api_getDashboard()` | 起動時、［更新］、各操作の後 | null なら「データを取得できませんでした（res=null）」、`res.error` なら「サーバーエラー: 」、それ以外は描画 | 「API失敗: 」「GASエディタのログも確認してください」 | 両表に「読込中...」 | 989-1016 |
| T8 | `api_markOvertimeDone(requestId)` | ［残業完了］ | 実績・休憩・実残業のセルを書き換え、alert | alert＋T7 | **楽観**：操作列を「完了済」に。ボタン無効化 | 1289-1311 |
| T9 | `api_markOvertimeDone(requestId, timeStr)` | 手入力［記録］ | 同上 | 同上 | **楽観**：「完了済」 | 1314-1337 |
| T10 | `api_markHolidayStart(requestId)` | ［開始］ | 何もしない | alert＋T7 | **楽観**：操作列を［休日完了］だけにする（手入力・編集・削除が消える） | 1340-1353 |
| T11 | `api_markHolidayStart(requestId, timeStr)` | 手入力［記録］ | 何もしない | alert＋T7 | **楽観**：［休日完了］［手入力］＋終了時刻欄（編集・削除は出ない） | 1356-1376 |
| T12 | `api_markHolidayDone(requestId)` | ［休日完了］ | セル書き換え＋alert | alert＋T7 | **楽観**：「完了済」 | 1379-1401 |
| T13 | `api_markHolidayDone(requestId, timeStr)` | 手入力［記録］ | 同上 | 同上 | **楽観**：「完了済」 | 1404-1427 |
| T14 | `api_editRequest(requestId, {hours, targetDate?})` | 編集［保存］ | モーダルを閉じ、`res.ok` なら alert＋T7、違えば「更新に失敗しました。」 | ボタンを戻して alert | ボタン「保存中...」で無効化 | 1488-1522 |
| T15 | `api_deleteRequest(requestId)` | ［削除］ | `res.ok` なら alert＋T7 | 「削除エラー: 」 | なし | 1524-1545 |
| T16 | `api_getOrderList()` | 申請モーダルを開いたとき（キャッシュが空なら） | `ORDER_CACHE` に入れるだけ | 何もしない | なし（読み込み前に検索すると「該当する工番がありません」） | 1759-1764 |
| T17 | `api_getHolidayCandidateDates()` | 休日の申請モーダルを開いたとき | 候補日ボタン（土＝緑、日・祝＝青、祝日名つき） | 「取得エラー: 」 | 「候補日を取得中...」 | 1781-1817 |
| T18 | `api_submitRequest(payload)` | ［申請送信］ | `res.success===true` か `res.ok===true` なら閉じて alert＋T7。違えばモーダルを開いたまま `res.error` を alert | ボタンを戻して「送信エラー: 」 | ボタン「送信中...」で無効化 | 1959-2030 |

### 3.2 approver.html（6 か所）

| # | 関数（引数） | 場面 | 成功時 | 失敗時 | 読み込み / 楽観 | 行 |
|---|---|---|---|---|---|---|
| A1 | `api_getApproverDepts()` | 起動時、［承認者変更］、部署未選択での更新 | 2.2 の流れ。結果は `{isAdmin, currentEmail, approvers:[{name,email,depts}]}` を想定 | モーダルに「エラー: 」 | 「承認者情報を確認中...」 | 409-487 |
| A2 | `api_getApproverDashboard(SELECTED_DEPTS)`（配列） | 承認者を選んだとき、［更新］、失敗時の復元 | サブ見出しを更新し描画。null / `res.error` の扱いは top と同様（文言は「デプロイの「新しいデプロイ」ではなく既存デプロイの更新が必要です」など GAS 前提） | サブ見出し「API呼出失敗」と表にエラー | 「読込中...」 | 523-568 |
| A3 | `api_approveRequest(requestId)` | ［承認］ | ボタン文字「承認済」（行は既に置き換え済み） | alert＋A2 | **楽観**：行を承認済みに、⚠を外す、緑フラッシュ、KPI 再計算 | 785-800 |
| A4 | `api_rejectRequest(requestId)` | ［却下］ | `res.ok` でなければ alert＋A2 | alert＋A2 | **楽観**：行を却下に、赤フラッシュ、KPI 再計算 | 803-821 |
| A5 | `api_approveRequestsBatch(requestIds)` | ［残業 全件承認］／［休日 全件承認］ | ボタン「一括承認完了」（無効のまま）。`res.results` に失敗があれば alert＋A2 | ボタンを「一括承認」に戻し alert＋A2 | **楽観**：対象行をすべて承認済みに。ボタン「承認処理中…」 | 849-900 |
| A6 | `api_approverMonthlySummary(ym, SELECTED_DEPTS)` | 月次を初めて開いたとき、［更新］ | 描画 | alert＋表にエラー | 「読込中...」 | 937-960 |

### 3.3 approve2.html（3 か所）

| # | 関数（引数） | 場面 | 成功時 | 失敗時 | 読み込み / 楽観 | 行 |
|---|---|---|---|---|---|---|
| S1 | `api_getSecondaryApprovalList(dateVal)` | 起動時、［読込］、各操作後 | `DATA = res.items` で描画 | alert | 「読み込み中...」 | 266-282 |
| S2 | `api_secondaryApprove(requestId)` | ［二次承認］ | `res.ok` なら `approvedBy2` 等を入れ `pdfFileId=''` にして再描画。違えば alert＋S1 | alert＋S1 | ボタン「処理中...」。`approvedBy2='(処理中)'` を先に入れる | 395-423 |
| S3 | `api_secondaryApproveBatch(ids)` | ［一括二次承認］ | `res.ok` なら失敗件数を alert（あれば）して S1 | alert＋S1 | ボタン「処理中...」 | 425-456 |

### 3.4 admin.html（7 か所）

| # | 関数（引数） | 場面 | 成功時 | 失敗時 | 読み込み表示 | 行 |
|---|---|---|---|---|---|---|
| D1 | `api_adminDeptOptions()` | 起動時 | 部署プルダウン（`ALL`→「全体」） | alert | なし | 473-482 |
| D2 | `api_adminMonthlySummary(ym, dept)` | 起動時、［月次を更新］ | KPI・表・グラフ | alert | なし | 590-600 |
| D3 | `api_adminDailyDetail(day, dept)` | ［日次を表示］、`?date=` ありの起動時、手動入力の後 | 日次表 | alert | なし | 602-652 |
| D4 | `api_adminManualComplete(ME_REQUEST_ID, endTime, startTime)` | 手動入力［登録］ | 閉じて alert＋D3 | 「エラー: 」 | なし | 683-703 |
| D5 | `api_adminFiscalTrend(base, dept)` | 起動時、年度推移［更新］ | 折れ線＋「年度：開始 〜 終了（部署）」 | alert | なし | 705-717 |
| D6 | `api_adminSpecialClauseCount(base, dept)` | 起動時、［更新］ | 特別条項表 | alert | なし | 962-988 |
| D7 | `api_adminPendingWatch()` | 起動時、［更新］ | 滞留表 | alert | なし | 991-1022 |

- 起動時に D1・D2・D5・D6・D7（＋`?date=` があれば D3）を同時に投げます。D1 の結果を待たないので、初回の D2・D5・D6 の部署は HTML の初期値 `ALL` です（確認済み admin.html:194, 1025-1032）。

### 3.5 staffSummary.html（6 か所）

| # | 関数（引数） | 場面 | 成功時 | 失敗時 | 行 |
|---|---|---|---|---|---|
| F1 | `api_getDeptList()` | 起動時 | 部署プルダウン → F2 | 空の部署で描画 | 474-499 |
| F2 | `api_getWorkerInfo()` | F1 の後 | 「ログイン：」表示、本人の部署なら F3 | そのまま描画 | 478-496 |
| F3 | `api_getWorkersByDept(me.dept)` | F2 の後 | 本人を選んで描画 | そのまま描画 | 483-492 |
| F4 | `api_getWorkersByDept(state.dept)` | 部署変更時（loadStaff） | スタッフプルダウン（前の選択が残っていれば維持） | 「全員」だけにする | 290-304 |
| F5 | `api_getWorkersByDept(state.dept)` | 一覧の行クリック（openPerson） | コード一致（なければ名前の部分一致）で本人を選び描画、先頭へスクロール | 「全員」で描画 | 440-459 |
| F6 | `api_getClosingSummary('YYYY-MM', dept, staffToken)` | 描画のたび | `mode==='personal'` なら個人、それ以外は一覧。null なら「データがありません」 | 「読み込みに失敗しました：」 | 307-320 |

- 読み込み中は「読み込み中...」を出します。staffToken は「作業員コード 氏名」の文字列そのものです（確認済み staffSummary.html:283-285）。

### 3.6 サーバー関数 30 種類（定義場所）

| 関数 | 定義 | 呼ぶ画面 |
|---|---|---|
| api_getDeptList | deptAuth.gs | top, staffSummary |
| api_getWorkersByDept | deptAuth.gs | top, staffSummary |
| api_getWorkerInfo | masters.gs | top, staffSummary（＋redirect でサーバー内から） |
| api_getOrderList | masters.gs | top |
| api_getHolidayCandidateDates | formMap.gs | top |
| api_getFormUrl | formMap.gs | top（デッド）、redirect（サーバー内） |
| api_submitRequest | formSubmit.gs | top |
| api_getDashboard | worklog.gs | top |
| api_markOvertimeDone | worklog.gs | top |
| api_markHolidayStart | worklog.gs | top |
| api_markHolidayDone | worklog.gs | top |
| api_editRequest | worklog.gs | top |
| api_deleteRequest | worklog.gs | top |
| api_getApproverDepts | approval.gs | approver |
| api_getApproverDashboard | approval.gs | approver |
| api_approveRequest | approval.gs | approver |
| api_rejectRequest | approval.gs | approver |
| api_approveRequestsBatch | approval.gs | approver |
| api_approverMonthlySummary | approval.gs | approver |
| api_getSecondaryApprovalList | approval.gs | approve2 |
| api_secondaryApprove | approval.gs | approve2 |
| api_secondaryApproveBatch | approval.gs | approve2 |
| api_adminDeptOptions | adminApi.gs | admin |
| api_adminMonthlySummary | adminApi.gs | admin |
| api_adminDailyDetail | adminApi.gs | admin |
| api_adminManualComplete | adminApi.gs | admin |
| api_adminFiscalTrend | adminApi.gs | admin |
| api_adminSpecialClauseCount | adminApi.gs | admin |
| api_adminPendingWatch | adminApi.gs | admin |
| api_getClosingSummary | staffSummaryApi.gs | staffSummary |

（確認済み：`grep "function api_..." src/*.gs`）

シムを作るときの注意：

- 引数・戻り値は JSON で往復できる形です（配列・オブジェクト・文字列）。失敗時は `err.message` を使う箇所と `err` をそのまま文字列にする箇所が混在しています（確認済み）。シムは `Error` 風の `{message}` を失敗ハンドラに渡すと、どちらでも動きます。
- 成功ハンドラが null を受けるケース（`res=null` 表示）は、GAS で Date などを返して直列化に失敗したときの名残と見られます（推測）。

---

## 4. 作業員の申請の流れ

### 4.1 今の流れ（モーダル → api_submitRequest）

確認済み（top.html:733-751, 1698-2030）：

1. 部署を選びます。起動時に `api_getWorkerInfo()` で本人の部署が自動で選ばれます。
2. 作業員を選びます。プルダウンの value は作業員コード、表示は「コード 氏名」の文字列です。
3. ［残業申請］または［休日出勤申請］を押します。部署・作業員が未選択なら alert です。
4. 申請モーダル STEP 1 / 2：
   - 作業実施日：残業は**ブラウザの今日**で固定（「yyyy-MM-dd（本日）」）。休日は `api_getHolidayCandidateDates()` の候補日ボタン（ラベルに日付を含む）か、「上記以外の日付：」の `type=date` ＋［決定］で選びます。
   - 過去日カレンダー（折りたたみ、「過去日に申請する場合はこちら」）：警告「過去申請をされる場合は、必ず事前に承認者および総務へご連絡ください。」、月送り ◀ ▶、7 列 × 6 行のマス。選べるのは**年度開始日（3/16）〜今日**です。今日のマスは青枠、選択は橙。「選択中: yyyy-MM-dd」と［選択を解除］。
   - 工番（任意、最大 3 件）：検索欄にスペース区切りで入力し、`code / product / dest` の AND 部分一致（大文字小文字無視）で最大 80 件を表示、超えたら「他 N 件…」。クリックで空きスロットに入れ、×で外すと前に詰めます。
   - ［次へ →］：休日で日付が未選択なら alert。
5. STEP 2 / 2：サマリー「日付: <STEP1 の日付> ／ 工番: <コード> or なし」、理由（必須、「その他」を含むと補足理由が必須）、予定時間（必須）、業務内容（任意）。［← 戻る］［キャンセル］［申請送信］。
6. confirm の後、`api_submitRequest(payload)` を呼びます。

payload（確認済み top.html:1993-2006）：

```
{ requestType, dept, workerLabel, targetDate, workNo1, workNo2, workNo3,
  reason, reasonDetail, workContent, hours, isRetroactive }
```

- `workerLabel` は「コード 氏名」の表示文字列です（コードだけではありません）。
- 過去日カレンダーで日付を選ぶと、STEP 1 の日付より**優先**され、`isRetroactive: true` になります（確認済み top.html:1976-1977）。STEP 2 のサマリーには過去日が出ず、STEP 1 の日付が出ます（確認済み top.html:1929）。気づきにくい挙動なので、移行先でも同じにするか決める必要があります。
- 過去日カレンダーで**今日**も選べます。その場合も `isRetroactive: true` になります（確認済み top.html:1676。`cellDate > today` だけを不可にしている）。
- 送信が成功するとモーダルを閉じ、「〇〇申請が完了しました。」の後にダッシュボードを読み直します。

### 4.2 Google フォームを使う旧経路

- top.html の `showDatePicker` → `showOrderPicker` → `submitOrderPicker` → `doOpenForm` → `api_getFormUrl(...)` で、新しいタブに Google フォーム（事前入力つきと推測）を開く経路が残っていますが、`showDatePicker` / `showOrderPicker` を呼ぶ処理がなく、使われていません（確認済み top.html:757-986。`openForm` は `showRequestModal` だけを呼ぶ、top.html:750）。
- サーバー側の `?page=redirect`（1.4）は生きていて、フォームへ飛ばします。フォーム送信のトリガー（onFormSubmit）がまだ動いているかは別の棚卸しの範囲です。
- 移行で旧経路を捨てるなら、`?page=redirect` を誰かが使っていないか確かめる必要があります。

### 4.3 申請まわりの不具合・癖

- **本人の作業員が自動で選ばれません**（確認済み、コードの論理から）。`setInterval` の条件 `wSel.options[0].value !== ''` は、先頭が必ず value 空のプレースホルダーなので一度も真になりません。部署だけが自動選択され、5 秒でタイマーが止まります（top.html:670-681, 711）。移行先で直すか、今の挙動を保つか決める必要があります。
- 本人の部署がプルダウンにない場合、`sel.value` が効かず、空の部署で `onDeptChange()` が走ります（確認済み top.html:666-668）。
- 申請モーダルの日付は、休日候補ボタンの `textContent` に日付文字列が含まれるかで選択表示を付けます（確認済み top.html:1825-1827）。ラベル形式が変わると選択表示が付かなくなります。

---

## 5. 日付・時刻と数値の扱い（ブラウザ側）

| 箇所 | 処理 | Safari などの心配 | 根拠 |
|---|---|---|---|
| top：4h 超過判定 | `new Date(targetDate + 'T17:20:00+09:00')`、`actualStartAt` は空白を `T` に置き換え、オフセットがなければ `+09:00` を足す | ISO 形式なので問題なし。`targetDate` が `yyyy-MM-dd` 以外だと `Invalid Date` → 経過 NaN → バッジなし | 確認済み top.html:1166-1175 |
| approver：提出時刻 | `'yyyy-MM-dd HH:mm:ss'` を正規表現で分解（`new Date` を使わない）。当日判定は `#todayDate` の表示文字列と比べる | 対策済み（コメントあり） | 確認済み approver.html:720-730 |
| admin：手動入力の開始時刻 | `actualStartAt` から `(\d{2}:\d{2})` を正規表現で取り出す | 問題なし | 確認済み admin.html:669-670 |
| 今日の日付 | top の残業日、過去カレンダー、approve2 / admin の初期日付、approver の月初期値、staffSummary の締め月は、すべて**ブラウザの時計とタイムゾーン**で決まる | 日本以外のタイムゾーンの端末や時計のずれで日付がずれる（推測） | 確認済み top.html:1714-1718, 1553-1567、approve2.html:238-241、admin.html:445-452、approver.html:912-917、staffSummary.html:248-252 |
| 年度開始（過去カレンダー） | 3/16 以降ならその年の 3/16、前なら前年の 3/16。サーバーの `getFiscalYearStartDate_` と同じと書かれている | — | 確認済み top.html:1551-1559 |
| admin の「年度推移」 | 見出しは「4/1〜3/31」（過去カレンダーの 3/16 とは別の定義） | 定義が 2 つある点に注意 | 確認済み admin.html:276, 376 |
| staffSummary の締め期間 | 毎月 16 日〜翌月 15 日 | — | 確認済み staffSummary.html:248-260 |
| `padStart` | ES2017。古いブラウザは不可 | 現実的には問題なし（推測） | 確認済み 各ページ |

数値の書式：

| 関数 | 中身 | ページ |
|---|---|---|
| `fmtMin(min)` | 負数は 0 に丸め、四捨五入。「N分」「N時間」「N時間M分」 | top:644、approver:392、approve2:242、admin:453 |
| `fmtMin(min)`（staffSummary 版） | 負数を 0 にしない。`Math.round` だけ（null なら `NaN分` になりうる） | staffSummary:268-269 |
| `hm(min)` | 「H:MM」 | staffSummary:270 |
| 予定時間の表示 | 残業は「0.5 時間」のように選択肢の文字列＋「 時間」 | top:1754, 1456 |
| しきい値 | 40h=2400 分、60h=3600 分、残り 300 分以下で黄（画面に直書き） | approver:971-973, 1017-1022、admin:494-496, 539-545 |
| staffSummary のしきい値 | サーバーの `meta.limit40` / `meta.limit60`、ゲージの最大は 4200 分（直書き）。判定は「より大きい」（`>`）で、admin / approver の「以上」（`>=`）と違う | staffSummary:325, 331, 403-404 |

---

## 6. 外部リソースとエスケープ（公開インターネットに出すときの心配）

外部リソース：CDN・Web フォント・外部スクリプトはありません。グラフはすべて canvas の自前描画です（確認済み）。読み込む外部 URL は Drive の PDF リンクと社内ポータルのリンクだけです。

エスケープ関数：

- `escHtml(s)`（top / approver / approve2 / admin）は `& < > "` だけを置き換え、**`'` は置き換えません**（確認済み top.html:886-888 ほか）。
- `esc(s)`（staffSummary）も `& < > "` だけです（確認済み staffSummary.html:271-272）。

心配な箇所（エスケープの抜け）：

| 箇所 | 内容 | 危険度 | 根拠 |
|---|---|---|---|
| admin 日次の［手動入力］ | `onclick="openManualEntry('<escHtml(workerName)>', ...)"`。`'` を置き換えないので、氏名に `'` が入ると JS が壊れる・任意コードが動く | 中（作業員マスタを書ける人が攻撃者になる場合） | 確認済み admin.html:618-620 |
| top の操作ボタン | `onclick="doOvertimeDone('<requestId>',this)"` など requestId を無エスケープで埋め込む。`data-request-id` も同じ | 低（requestId はサーバーが作る値。推測） | 確認済み top.html:1187-1188, 1205-1242, 1250 |
| approver の操作ボタン | `doApprove('<requestId>')` / `doReject` を無エスケープで埋め込む | 低（同上） | 確認済み approver.html:734-737, 746 |
| PDF リンク（top） | `pdfFileId` を `href` に無エスケープで埋め込む（他ページは escHtml 済み） | 低 | 確認済み top.html:1150-1151 |
| 工番選択（デッド） | `o.code.replace(/'/g,"\\'")` だけで `"` を逃がさず onclick に入れる | デッドコードなので現状なし | 確認済み top.html:875 |
| ポップアップのリンク（デッド） | `url` を無エスケープで `href` に入れる | デッドコード | 確認済み top.html:981 |
| `<?!= REASONS_JSON ?>` など | 無エスケープ出力。中身はサーバーの定数なので現状は安全。移行先でも `</script>` を含まないように JSON を出す必要がある | 低 | 確認済み top.html:637-639 |
| `'<?= DATE_PARAM ?>'` | GAS の `<?= ?>` は出力先に合わせて自動エスケープする（JS の文字列の中なら JS 用にエスケープ）。移行先でテンプレートに素直に差し込むと、`?date=` から XSS になる。admin は形式チェックもしない | **高（移行で作り込みやすい）** | 確認済み approve2.html:256、admin.html:466。GAS のエスケープ挙動は推測（公式ドキュメントの「contextual escaping」） |
| `<?= message ?>` | GAS が HTML エスケープ。内容にデバッグ情報とメールアドレスを含む | 情報漏れ（中） | 確認済み no_auth.html:73、adminApi.gs:555-559 |
| redirect の `formUrl` | `<script>window.top.location.href="<formUrl>"</script>` に無エスケープ | 低（Settings の値由来。推測） | 確認済み adminApi.gs:654-655 |
| サーバーのエラー文 | `err.message` を alert にそのまま出す（alert なので XSS ではない）。表の中は escHtml 済み | 情報漏れ（低） | 確認済み 各ページ |
| 数値の無エスケープ | `p.totalNet`、`r.hoursPending`、`r.count60`、`p.annualOver`、`yr` などをそのまま HTML に入れる | 低（サーバーが数値を返す前提） | 確認済み admin.html:553, 981, 1015、staffSummary.html:385, 406 |

その他、公開時に見直すこと：

- `ALLOWALL` で他サイトの iframe に入れられます。移行先では `frame-ancestors` を決める必要があります（ポータルが iframe で埋め込んでいるかは推測。要確認）。
- staffSummary は誰でも全員の月次実績を見られます（画面側に制限なし）（確認済み）。
- approver はメールが取れないと doGet の判定を素通りします（確認済み adminApi.gs:569）。

---

## 7. GAS に依存している部分

| 依存 | 使っている所 | 移行先での置き換え | 根拠 |
|---|---|---|---|
| `google.script.run.withSuccessHandler(..).withFailureHandler(..).fn(args)` | 全 5 ページ、40 か所 | 同じ書き方で使えるシム（fetch で `/api/<fn>` に JSON を POST し、成功・失敗ハンドラを呼ぶ）。ハンドラを付けない呼び方はないが、`withFailureHandler` だけ先に書く順序もある（admin.html:473-482 は success → failure の順、他も同じ） | 確認済み |
| `google.script.host` / `google.script.url` | **使っていない** | 不要 | 確認済み（grep 0 件） |
| テンプレートの `<?= ?>` | `APP_URL`、`PORTAL_URL`、`DATE_PARAM`、`message` | Jinja2 などで差し込む。`DATE_PARAM` は JS 文字列用のエスケープ（`json.dumps` で引用符ごと出す等）が必要 | 確認済み |
| テンプレートの `<?!= ?>` | `REASONS_JSON`、`OT_HOURS_JSON`、`HD_HOURS_JSON` | `json.dumps` の結果を `</` を逃がして出す | 確認済み top.html:637-639 |
| `HtmlService` の iframe（script.googleusercontent.com のサンドボックス） | 全ページ。`<base target="_top">` と各リンクの `target="_top"` はこのため | iframe でなくなるので `_top` は同じページ遷移として害なし | 確認済み（`<base>`）、iframe 構造は推測（GAS の仕様） |
| `setTitle` | タブのタイトル | `<title>` に入れる。staffSummary の独自 `<title>` をどうするか決める | 確認済み adminApi.gs:596 |
| `setXFrameOptionsMode(ALLOWALL)` | 全応答 | 応答ヘッダーで決める | 確認済み |
| `ScriptApp.getService().getUrl()` | `APP_URL`、redirect のトップへ戻るリンク | 設定値（新しいベース URL） | 確認済み adminApi.gs:536, 634, 647 |
| `Session.getActiveUser().getEmail()` | admin / approve2 / approver の入口判定、本人の自動選択（サーバー側） | Cloudflare Access などで取ったメールアドレス | 確認済み adminApi.gs:549, 568 |
| `window.top.location.href` | redirect の応答 | 302 リダイレクトでよい | 確認済み adminApi.gs:654 |
| `window.open` ＋ `document.write` | デッドコードだけ | 不要 | 確認済み top.html:944-953 |
| 画面の文言に出る GAS 前提の案内 | 「GASエディタのログも確認してください」「デプロイの「新しいデプロイ」ではなく既存デプロイの更新が必要です」「GASエディタのログ（実行数）も確認してください」 | 文言を変えるか決める（利用者の目に入るのはエラー時だけ） | 確認済み top.html:1013、approver.html:540, 565 |
| Drive の PDF URL | `https://drive.google.com/file/d/<id>/view` | PDF を Drive に置き続けるなら同じ。VPS に置くなら URL の作り方を変える | 確認済み |
| 社内ポータル URL（別 GAS） | 全ページ | 設定値に出す | 確認済み adminApi.gs:554, 578, 589 |

---

## 8. 受け入れチェック用の要点（移行先で同じになっているか見る項目）

- ページ 5 つ＋権限エラーが同じ URL 形式（`?page=...&date=...`）で開け、不明な `page` は top になる。
- admin / approve2 は管理者以外に権限エラー、approver は承認者か管理者だけ（メールが取れないときの扱いを決めておく）。
- `?page=approve2&date=yyyy-MM-dd` で対象日が入り、`?page=admin&date=...` で日次一覧が自動で出る。
- 40 か所の呼び出しが、成功・失敗・楽観更新を含めて同じ動きになる（3 章の表）。
- 申請モーダルの 2 ステップ、休日候補日、過去日カレンダー（3/16〜今日、今日も選べる、選ぶと優先）、工番 3 件、「その他」の補足理由必須、confirm 文言。
- top の操作列（完了・開始・手入力・編集・削除・4h 超過・期限超過バナー）と時刻プルダウンの範囲。
- approver の承認者選択モーダル（管理者だけ選べる）、一括承認の confirm（日付ごとの件数・過去日の件数）、月次の並べ替えとフィルタ。
- approve2 のチェックできる行の条件と、単票承認に confirm がないこと。
- admin の CSV（列・ファイル名・BOM なし）、3 つの canvas グラフ、手動入力モーダル、特別条項・滞留監視。
- staffSummary の締め月（16 日〜翌 15 日、15 日以前は前月）、個人ビューと一覧ビュー、行クリックで個人へ。
- エスケープの抜け（6 章）を、見た目を変えずに直すかどうか。
- 既知の不具合（作業員の自動選択が効かない、KPI の数え方の違い、STEP 2 のサマリーに過去日が出ない）を「そのまま」にするか「直す」かの判断。

## 確かめること

- `?page=redirect`（Google フォームへの旧経路）を、ポータルやブックマークから今も使っている人がいるか。
- 社内ポータルがこのアプリを iframe で埋め込んでいるか（`ALLOWALL` の必要性）。
- 権限エラー画面のデバッグ情報（メールアドレスと判定ログ）を、公開後も出すか。
- 作業員の自動選択が効かない不具合を直してよいか（利用者には「変わった」と見える）。
- staffSummary を誰でも全員分見られる今の設計を、公開後も続けるか。
