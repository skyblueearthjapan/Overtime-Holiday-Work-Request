# GAS → VPS 移植の正しさレビュー（残業・休日出勤申請）

実施: 2026-11-26。対象は `src/`（GAS 本番）と `vps-app/`（VPS 版）。コードは変えていません。

## 結論

ふつうの画面操作で、数字・状態・権限・返す形が GAS と変わる違いは見つかりませんでした。高は 0 件、中は 1 件、低は 19 件です。低のほとんどは、画面が送らない値（配列でない引数、0000 年、`0x` 付きの数など）か、シートに珍しい書き方があるときだけ起きます。判定は COMMENT（切り替えを止める差はありません）です。

| 重さ | 件数 |
|---|---|
| 高 | 0 |
| 中 | 1 |
| 低 | 19 |
| 推測（確かさが低く、判定に使わないもの） | 5 |
| 仕様書の食い違い | 3 |

## 確かめ方

- 38 関数（実装仕様 4-2・4-3）と Sebastian の入口（`src/api.gs` の `doPost` と `vps-app/api/app/domain/sebastian.py`）を、GAS と Python で 1 行ずつ読み比べました。4 つの領域（実績と Sebastian、申請・編集・削除・マスタ、承認・二次承認・ダッシュボード、集計・総務・運用）に分けて読みました。
- 日付の読み方（`new Date` の写し）、祝日の計算（2019〜2031 年の 678 週）、並べ替え（`localeCompare('ja')` 1500 組）は、node（V8・ICU）と Python の結果を突き合わせました。
- 画面の JS は GAS の `<script>` と diff を取り、画面が読むキーとサーバーが返すキーを照合しました。`gas-shim.js` の振り分け（200 は成功のハンドラ、400・403 は `Error(message)` で失敗のハンドラ）も GAS の `google.script.run` と同じでした。
- 試験は `904 passed, 1 skipped` でした。ただし、これは同じ動きの証拠には使っていません。
- 仕様（計画書 9-1 #11〜#25、抜け調査の決定 A〜D、実装仕様 4-1〜4-4・8 章）が認めている差は指摘に入れていません。

## 中

### 中-1 対象日が読めない申請の完了で、エラー文が違う（確かさ: 中。GAS 側の文は推測）

- GAS: `src/worklog.gs:157-158`・`232`・`292`、`src/adminApi.gs:443-444`、`src/api.gs:146-147`・`208-209`
- Python: `vps-app/api/app/domain/worklog.py:62-66`（`_target_ymd_or_raise`）
- 入力: Requests の `targetDate` が空か `abc` の申請で、残業完了・休日完了・総務の手動入力を押す。
  - GAS: `new Date('abc')` が Invalid Date になり、`Utilities.formatDate` が英語の例外文を投げる見込みです。
  - Python: 400 `作業実施日が無効です: abc`。
- 仕様 8-2 が受け入れる差としているのは時刻（`25:00` など）だけで、日付の側は書かれていません。
- 直し方: Python の文のほうが利用者に分かりやすいので、コードは今のままにして、仕様 8-2 に「対象日が読めないときは `作業実施日が無効です: X`（受け入れる差）」を一行足します。

## 低

### 低-1 手入力の時刻・日付の末尾に改行があると、Python だけ 400 になる（確かさ: 高）

- GAS: `src/worklog.gs:163`・`230`・`290`、`src/api.gs:145`・`207`
- Python: `vps-app/api/app/domain/jsdate.py:41-42`（`YMD_RE`・`HM_RE` を `re.match` と `$` で使う）
- 入力: Sebastian の `manualEndTime: "19:30\n"`
  - GAS: 正規表現に合わないので、黙って今の時刻で完了します（G13）。
  - Python: `$` が末尾の改行の前で合うため形は通り、その後の解釈で 400 `時刻が不正です: 19:30\n` になります。
- 直し方: `fullmatch` を使うか、`$` を `\Z` に変えます。

### 低-2 配列・オブジェクトを文字にする方法が JS の `String()` と違う（確かさ: 高）

- GAS: `src/api.gs:26`・`35`・`40-41`・`47-48`、`src/formSubmit.gs`（`String(data.isRetroactive)` など）
- Python: `vps-app/api/app/jsutil.py`（`js_string`）、`vps-app/api/app/domain/sebastian.py:139`・`147`・`151`・`153`、`vps-app/api/app/domain/requests.py:119`
- 入力と結果:
  - Sebastian `{"requestId":["R1"]}`: GAS は `"R1"` として処理します。Python は `申請が見つかりません: ['R1']`。
  - 申請 `isRetroactive: ["true"]`: GAS は遡及として扱います。Python は遡及になりません。
  - `requestType: {}`: エラー文の後ろが GAS は `[object Object]`、Python は `{}`。
- 直し方: `js_string` で list は要素を `,` でつなぎ（`None` は空）、dict は `[object Object]` にします。

### 低-3 Sebastian の本文が `null` のときの返事が違う（確かさ: 高）

- GAS: `src/api.gs:13`・`22` / Python: `vps-app/api/app/domain/sebastian.py:135-137`
- 入力: 本文 `null`。GAS は `body.token` で TypeError になり、HTML のエラー画面を返します。Python は `{"error":"unauthorized"}`。
- 直し方: Python のほうが安全なので、仕様 4-4 に受け入れる差として書きます。

### 低-4 Sebastian の返す `targetDate` が、文字のセル `yyyy/M/d` のとき違う（確かさ: 中）

- GAS: `src/api.gs:119`・`163`・`245`（`req.targetDate` を生のまま返す）
- Python: `vps-app/api/app/domain/sebastian.py:38-44`、`vps-app/api/app/domain/store.py:81-85`。原因は取り込みの `vps-app/api/app/services/sheet_codec.py:316-325` で、読める文字なら `target_date_text` を空にして日付だけを残すためです。
- 入力: `targetDate` のセルが文字の `2026/11/24`。GAS は `"2026/11/24"`、Python は `"2026-11-24"`。
- 仕様 4-4 は「文字ならその文字」と決めているので、仕様とも合っていません。
- 直し方: 文字のセルのときは、元の文字を `target_date_text` に残します。

### 低-5 一括承認・一括二次承認に配列でない値を渡すと、Python は承認してしまう（確かさ: 高）

- GAS: `src/approval.gs:389-390`・`408`、`776-777`・`794`
- Python: `vps-app/api/app/domain/approval.py:69-70`、`vps-app/api/app/domain/secondary.py:112-113`（配列でなければ `[x]` に包む）
- 入力: `api_approveRequestsBatch("<実在する requestId>")`
  - GAS: 文字列を 1 文字ずつ ID として扱うので、全件「見つかりません」になり、何も承認しません（51 文字以上なら `一括承認は50件までです。`）。
  - Python: その申請を承認します。数値 `5` では、GAS が `{ok:true, results:[]}`、Python が `見つかりません` の 1 件です。
- 画面は必ず配列を送るので、ふだんは起きません。
- 直し方: 包み直しをやめます。配列でなければ GAS と同じく文字列を 1 文字ずつに分けるか、少なくとも何も承認せずに返します。

### 低-6 一括承認の「見つかりません」の結果で、ID の書き方と並びが違う（確かさ: 高）

- GAS: `src/approval.gs:408`・`451`（JS のオブジェクトのキー）
- Python: `vps-app/api/app/domain/approval.py:77-78`・`114-115`
- 入力: `[null]` は GAS が `"null"`、Python が `"None"`。`[1.0]` は `"1"` と `"1.0"`。存在しない `["b","5"]` は、JS が数字だけのキーを先に並べるので `5, b`、Python は `b, 5`。
- 直し方: `str()` を `js_string` に替えます。厳密に合わせるなら数字だけのキーを先に並べます。

### 低-7 引数が多すぎる呼び出しを Python だけ 400 にする（確かさ: 高）

- GAS: 余った引数は黙って捨てます。
- Python: `vps-app/api/app/rpc.py:295-299`（`inspect.signature(...).bind` の失敗で `引数の数が正しくありません。`）
- 入力: `api_getApproverDepts("x")`。GAS は通常の結果、Python は 400。
- 画面は余分な引数を送らないので、直さなくてもよい差です。直すなら、引数の数が多いときは切り詰めてから呼びます。

### 低-8 0000 年の日付で Python だけ 500 になる（確かさ: 高）

- GAS: `src/formSubmit.gs:717-718` / Python: `vps-app/api/app/domain/jsdate.py:116`（`datetime(y, …)`）、呼び出し元 `requests.py:144`・`227`
- 入力: 休日の申請で `targetDate:'0000-01-01'`。GAS は 200 `{ok:false, error:'年度開始日（2026-03-16）より前の日付は申請できません。'}`。Python は `ValueError: year 0 is out of range` で 500。編集でも同じです。
- 直し方: `y < 1` のときは例外にせず「有効だが範囲の前」として扱います。

### 低-9 予定時間の `0x` などの読み方が `Number()` より甘い（確かさ: 高）

- GAS: `src/formSubmit.gs:403-407` / Python: `vps-app/api/app/domain/jsdate.py:70-74`（`js_number`）
- 入力: `hours:'0x1_0'`。GAS は `Number()` が NaN で 400 `残業の予定時間が不正です: 0x1_0`。Python は 16 と読んで 960 分で受け付けます。`'0x 1'` も同じ形の差です。
- 直し方: `int(text[2:], base)` の前に、その基数の数字だけかを `fullmatch` で確かめます。

### 低-10 予定時間が極端に大きいと Python だけ 500 になる（確かさ: 中）

- GAS: `src/formSubmit.gs:407` / Python: `vps-app/api/app/jsutil.py:23-25`（`js_round`）と DB の `approved_minutes` 列
- 入力: `hours:'1e308'`。GAS は `Math.round(Infinity)` のまま書き込みます。Python は `math.floor(inf)` が OverflowError で 500。`'1e20'` は整数の列に入らない見込みです（推測）。
- 直し方: 結果が有限で列の範囲に入るかを確かめ、外れたら `残業の予定時間が不正です: X` の 400 にします（GAS より厳しくなるので、受け入れる差として仕様に書く）。

### 低-11 作業員の文字列を分ける空白の種類が JS の `\s` と違う（確かさ: 高）

- GAS: `src/formSubmit.gs:395` / Python: `vps-app/api/app/domain/requests.py:57`（`re.split(r"\s+")`）
- 入力: `'001\ufeff山田'` は GAS が `["001","山田"]` に分け、Python は分けません。`'001\x1c山田'`・`'001\x85山田'` は逆に Python だけ分けます。
- 起きるのは作業員マスタに見えない文字が入っているときだけです。
- 直し方: JS の `\s` と同じ文字の組（`jsutil._JS_WHITESPACE` に `\u2028`・`\u2029` などを足す）を明示した正規表現で分けます。

### 低-12 `api_submitRequest(null)`・`api_editRequest(id, null)` の扱いが違う（確かさ: 高）

- GAS: `src/formSubmit.gs:678` 以降、`src/worklog.gs:540` 以降 / Python: `vps-app/api/app/domain/requests.py:105-106`・`196-197`
- 入力: `api_submitRequest(null)`。GAS は TypeError の文で失敗します。Python は `{}` として扱い、200 `{ok:false, error:'申請種別が不正です: '}`。
- 画面からは起きません。Python の動きのほうが穏当なので、受け入れる差として記録すれば足ります。

### 低-13 社内カレンダーの A 列が `yyyy-MM-dd` 以外の文字だと、候補日が変わる（確かさ: 低〜中）

- GAS: `src/formMap.gs:257-263`・`269` / Python: `vps-app/api/app/domain/masters.py:168-175`（`_calendar_date_text`）・`195-197`
- 入力: A 列が `'2026-1-5'` や `'2026-11-28 09:00:00'`。GAS では DB タブが USER_ENTERED で日付のセルになり、`fmtDate_` で候補日になります。Python は `'2026-1-5'` を捨て、時刻付きは読めずに一覧全体が従来方式（土日と計算の祝日）に切り替わります。
- 直し方: `_calendar_date_text` で、`-` 区切りの 1 桁の月日や時刻付きも `js_new_date` で `yyyy-MM-dd` に直します。本番のタブの A 列の型を一度確かめてください。

### 低-14 日本語の並び（`localeCompare(…,'ja')` の写し）が ICU と合わない組がある（確かさ: 高。node の ICU で測定）

- GAS: `src/adminApi.gs:122`（日次一覧）、`src/adminApi.gs:244`（部署の円グラフ）
- Python: `vps-app/api/app/domain/collation.py:25-84`、使う所は `vps-app/api/app/domain/summaries.py:169`・`106`
- 1500 組中 13 組が合いませんでした。例: `ケース` と `ケイ`（ICU は後ろ、Python は前）、`ﾃﾞｰﾀ` と `データ`、全角 `ＡＢ` と `AB`（ICU は同じ扱いでシートの順、Python は半角を先）。`佐々木`/`佐藤`、`渡邉`/`渡辺`、`髙橋`/`高橋` は一致しました。
- 仕様 8-4 は「合わない並びは受け入れる差として報告」と決めているので、報告のみです。直すなら、全角・半角だけの違いを同じ扱いにし、`ー` を前の字の母音に置き換える近似を足します。

### 低-15 スタッフ集計の個人明細で、同じ日の行の順が変わることがある（確かさ: 中）

- GAS: `src/staffSummaryApi.gs:112`（`a.targetDate - b.targetDate`）/ Python: `vps-app/api/app/domain/summaries.py:336`
- 入力: 同じ人・同じ日に 2 件あり、片方が文字 `'2026-11-05'`（JS では 9 時扱い）、もう片方が日付のセル（0 時）。GAS は日付のセルの行がいつも先、Python はシートの順のまま。
- 直し方: `target_date_text` がある行は 9 時として並べるキーを足します。混在が無ければ直さなくて構いません。

### 低-16 集計で年が範囲外だと Python だけ 500 になる（確かさ: 高）

- GAS: `src/adminApi.gs:142-147`、`src/staffSummaryApi.gs:26-31`
- Python: `vps-app/api/app/domain/summaries.py:55-57`・`280-282`、`vps-app/api/app/domain/fiscal.py:17-20`
- 入力: `api_adminMonthlySummary('9999-12','ALL')`・`'10000-01'` は GAS が空の集計、Python が 500。`api_getClosingSummary('0.5-01')` は GAS が 1900 年の空の一覧、Python が 500。
- 直し方: 年が 1〜9998 の外なら空の集計を返します。

### 低-17 `ops_retryJob` が整数でない ID も受け付ける（確かさ: 高。GAS に当たる関数は無い）

- Python: `vps-app/api/app/domain/ops.py:26`（`int(job_id)`）
- 入力: `ops_retryJob(1.9)`・`ops_retryJob(true)` は 1 番、`" 3 "` は 3 番を戻します。
- 直し方: `bool` でない `int` か、`^\d+$` の文字だけを受けます。

### 低-18 マスタに見出しが無いとき、GAS は例外、Python は空の一覧（確かさ: 中）

- GAS: `src/deptAuth.gs:117-129`（`api_getDeptList`・`api_getWorkersByDept`）、`src/masters.gs:80` 以降（`api_getOrderList`）
- Python: `vps-app/api/app/domain/masters.py:73-124`
- 入力: 作業員マスタに「作業員コード／氏名／部署」の見出しが無い。GAS は失敗ハンドラにエラー文、Python は `[]`。
- 表の作りが固定なので実際には起きない見込みです。受け入れる差として記録します。

### 低-19 Sebastian で、確定の後の例外が `{error, action}` になる恐れ（確かさ: 低）

- GAS: `src/api.gs`（PDF の失敗は `pdf:{error}` に入れて `ok:true`）
- Python: `vps-app/api/app/domain/sebastian.py:92-94`・`118-120`（WorkLogs を commit してから PDF の結果を待つ）
- 待っている間に DB の例外が出ると、保存済みなのに `{error, action}` が返り、Sebastian が再送する恐れがあります。`gas_jobs.wait_for_result` が例外を外へ出さない作りかは確かめていません。
- 直し方: 待つ部分を try で包み、例外は `pdf:{error:"…"}` にして `ok:true` で返します。

## 推測（確かさが低い。判定には使っていません）

- 申請 ID・状態の前後の空白: GAS は `getRequestById_`・`findWorkLogRowNo_`（`src/approval.gs:300`・`364`・`716`、`src/worklog.gs:59`・`64`・`591`）で行の値を `normalize_` してから比べます。Python は `vps-app/api/app/domain/store.py:51`・`118`・`156-163` で完全一致です。取り込み（`vps-app/api/app/services/sheet_codec.py:115`）で空白を取っていない見込みです。シートの ID が `"abc "` なら、GAS は見つけ、Python は `requestIdが見つかりません。` になります。`api_editRequest` の `newStatus`（`requests.py:248`）も生の値を返します。本番のデータに空白付きの ID・status があるかを確かめ、あれば取り込みで `normalize_gas` してください。
- ApproverMap の 1 行目が空のとき: GAS（`src/approval.gs:37-40`、`src/deptAuth.gs:34-35`）は 1 行目を見出しとして読みます。`vps-app/api/app/services/config_sync.py:78-88` は空行を捨てるので、2 行目が見出しに繰り上がり、GAS では効かない行が効く恐れがあります。
- 作業員マスタの見出しが全角の「Ｇｏｏｇｌｅアカウント」のとき: `master_sync.worker_rows` は前後の空白しか取らず、`masters._has_google_account_column` は `normalize_gas` で比べるので、列があるのに値が空になる恐れがあります。今のマスタにこの列は無いので影響はありません。
- `localeCompare` の測定は node の ICU で行いました。GAS 本番の ICU の版で同じ結果かは確かめていません。
- 中-1 の GAS の例外文は、実機で確かめていません。

## 仕様書の食い違い（コードの不具合ではありません）

- `docs/vps-implementation-spec.md:330`（4-3 の 17）: 同じ行に「`approval.gs:388-475` の形と順」と「結果の並びは渡された順」の両方があります。GAS と Python は Requests の行の順で、見つからない ID を最後に足します。画面（`vps-app/web/static/js/approver.js:500-507`）は並びに依存しません。文を「Requests の行の順、その後に見つからなかった ID」に直すことをおすすめします。
- `docs/vps-implementation-spec.md:342`（4-3 の 29）: 「データが無ければ `null`」とありますが、GAS（`src/staffSummaryApi.gs:57-193`）も Python も `null` を返さず、`{mode:'list', people:[]}` などを返します。仕様の文を消します。
- 実装仕様 8-2: 対象日が読めないときのエラー文（中-1）が受け入れる差に入っていません。

## 同じだと確かめた関数

- 実績: `api_markOvertimeDone`・`api_markHolidayStart`・`api_markHolidayDone`・`api_adminManualComplete`（確かめる順、17:20 JST、手入力の扱い、休憩の時刻窓、`js_round`、`updatedBy`、返す形）。`assertSelf_` は大文字小文字を区別（Q24）。WorkLogs の重複は「書くのは上の行、読むのは下の行」まで同じです。
- Sebastian: ping・3 つの action・エラー文（末尾の `: <id>`）・`{error, action}` の形・トークンの 3 つのエラーに action が無いこと・`status` は書く前の値・#23 の時刻窓方式。
- 申請: `api_submitRequest`（検査の順と文 1〜10、日付の V8 の読み方、3/15・3/16 の年度境界、0:30 JST、`'2.5'`→150 分、`半日`→240、`工番マスタ未登録: X`）、`api_editRequest`、`api_deleteRequest`。
- マスタ: `api_getDeptList`・`api_getWorkersByDept`・`api_getOrderList`・`api_getWorkerInfo`・`api_getHolidayCandidateDates`（祝日の計算は 678 週で一致）。
- 承認: `api_approveRequest`・`api_rejectRequest`・`api_approveRequestsBatch`（配列のとき）・`api_getApproverDepts`・`api_getApproverDashboard`・`api_getDashboard`・`api_getSecondaryApprovalList`・`api_secondaryApprove`・`api_secondaryApproveBatch`（配列のとき。件数の確かめが権限より先）、`permissions.py`。
- 集計: `api_approverMonthlySummary`・`api_adminMonthlySummary`（`>=` の判定、予測の 15.5→16、未来月）・`api_adminFiscalTrend`・`api_adminDailyDetail`（#22 は仕様どおり）・`api_adminSpecialClauseCount`・`api_adminPendingWatch`・`api_adminDeptOptions`・`api_getClosingSummary`、`ops_recalcBreaks` とほかの ops 8 関数（仕様 4-2 の形）。
- 画面: `top.js`・`approver.js`・`approve2.js`・`admin.js`・`staffSummary.js` が読むキーは、すべて Python の戻り値にあります。

## 直す順のおすすめ

1. 仕様書の 3 か所（中-1、4-3 の 17・29）を直す。手間が小さく、後の試験の基準がそろいます。
2. 低-5（配列でない一括承認で承認してしまう）と低-1（`$` と改行）。どちらも数行で直せます。
3. 推測の 1 つ目（ID の空白）は、本番のデータを 1 回数えてから決める。
4. 残りの低は、受け入れる差として仕様に書くか、手が空いたときに直す。
## 修正の結果 (team R)

2026-11-26、ブランチ `team-r`。VPS 側（`vps-app/`）だけを直しました。試験は `vps-app/api` で `1062 passed, 1 skipped`。

| 指摘 | 結果 | コミット・試験 |
|---|---|---|
| 中-1 対象日が読めないときの文 | 直した（コードは今のまま。仕様 8-2 に受け入れる差として追記） | e8a693b |
| 低-1 末尾の改行と `$` | 直した（入力を確かめる正規表現をすべて `\Z` に） | d2c12a5・`tests/test_review_fixes_r.py::test_regex_end_anchor_*`・`test_manual_time_with_newline_*` |
| 低-2 配列・オブジェクトの `String()` | 直した（`js_string` が配列を `,` でつなぎ、オブジェクトは `[object Object]`） | 889ccee・`test_js_helpers` |
| 低-3 Sebastian の本文 `null` | 直さない（Python のほうが安全。仕様 8-2a に受け入れる差として記録） | e8a693b |
| 低-4 文字のセルの `targetDate` を ISO で返す | 直さない（取り込みと行の反映の両方に関わり、行の反映がセルを書き直す恐れがあるため。本番に文字の `yyyy/M/d` があるかを先に数える） | — |
| 低-5 配列でない一括承認 | 直した（包まない。文字列は 1 字ずつ、数・オブジェクトは対象なし。二次承認も同じ） | 889ccee・`test_batch_string_argument_*`・`test_batch_non_array_*`・`test_secondary_batch_*` |
| 低-6 見つからない ID の書き方と並び | 直した（`js_property_key` と JS のキーの順） | 889ccee・`test_batch_missing_ids_use_js_keys_and_order` |
| 低-7 引数が多すぎる呼び出し | 直した（余りを捨てる） | ba7c9d0・`tests/test_rpc.py::test_optional_args_and_wrong_arity` |
| 低-8 0000 年 | 直した（500 にせず年度の検査で拒む） | d2c12a5・`test_year_zero_is_valid_but_before_fiscal_year` |
| 低-9 `0x1_0` など | 直した（基数ごとの数字だけを受ける） | d2c12a5・`test_js_number_rejects_loose_radix_literals` |
| 低-10 極端に大きい予定時間 | 直した（400。仕様 8-2a に受け入れる差） | d2c12a5・`test_huge_planned_hours_are_rejected_not_500` |
| 低-11 空白の種類 | 直した（JS の `\s` と同じ文字の組で分ける） | f88bb6a・`test_worker_label_split_uses_js_whitespace` |
| 低-12 `api_submitRequest(null)` | 直さない（Python のほうが穏当。仕様 8-2a に記録） | e8a693b |
| 低-13 社内カレンダーの A 列の書き方 | 直さない（本番のタブの A 列の型を確かめてから。マスタ同期の書き方にも関わる） | — |
| 低-14 `localeCompare('ja')` の 13 組 | 直さない（仕様 8-4 のとおり報告のみ） | — |
| 低-15 同じ日の行の順 | 直さない（文字と日付のセルが混ざるときだけ。混在の有無を先に確かめる） | — |
| 低-16 集計の年が範囲外 | 直した（500 にせず 400。仕様 8-2a に受け入れる差） | d2c12a5・`test_out_of_range_year_in_summaries_is_400_not_500` |
| 低-17 `ops_retryJob` の番号 | 直した（整数か数字だけの文字） | d2c12a5・`test_retry_job_accepts_only_integer_ids` |
| 低-18 マスタに見出しが無い | 直さない（表の作りが固定。仕様 8-2a に記録） | e8a693b |
| 低-19 確定の後の例外 | 直した（PDF の待ちの例外は `ok:true` のまま `pdf:{error}`） | 257b45e・`tests/test_sebastian.py::test_exception_while_waiting_for_pdf_keeps_ok_true` |
| 推測 1 ID の前後の空白 | 直した（GAS の `normalize_` と同じく、行の ID の空白を無視して探す。申請と WorkLogs） | 889ccee・`test_request_with_trailing_space_*`・`test_worklog_with_spaces_*` |
| 推測 2 ApproverMap の 1 行目が空 | 直さない（本番の表の形を確かめてから） | — |
| 推測 3 全角の見出し | 直さない（今のマスタに列が無い） | — |
| 仕様 4-3 の 17・29、8-2 | 直した（仕様の文を直した） | e8a693b |
