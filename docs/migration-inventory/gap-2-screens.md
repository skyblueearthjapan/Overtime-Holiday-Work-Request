# 抜け調査 2：画面（クライアント）

対象: `src/top.html`（2033 行）・`src/approver.html`（1178 行）・`src/approve2.html`（459 行）・`src/admin.html`（1035 行）・`src/staffSummary.html`（504 行）・`src/no_auth.html`（80 行）・`src/adminApi.gs` の `doGet`／`include_`／`handleFormRedirect_`（534-657 行）。全行を読んだ。照合した計画: `docs/vps-migration-plan.md`（主に 3-1・3-2・3-4・4-3・4-4・6-3・6-4・8-2・Q7・Q10・Q15）。サーバー側の返り値の形は、画面との取り決めに関わる範囲だけ `worklog.gs`・`approval.gs`・`formSubmit.gs`・`masters.gs`・`staffSummaryApi.gs` を読んで確かめた。

書き方: 「確認済み」はコードで確かめたこと。「推測」は本番の画面・実データで確かめていないこと。重さは 高／中／低。

## 要点（高を先に）

| # | 重さ | 指摘 | 確かさ |
|---|---|---|---|
| S1 | 高 | 総務画面の「日次：申請一覧」は、本番の GAS では 1 件でもあると表に何も出ていない見込み（返り値に Date が入り `google.script.run` が null を返す）。VPS で素直に移すと「直る」＝見た目が変わる。手動実績入力のモーダルと `?date=` の自動表示も同じ道を通る | 推測（根拠強） |
| S2 | 高 | 共有アカウントのログイン切れは「見込み」ではなく、貯蔵品の既定（無操作 12 時間・絶対 14 日）なら**毎週月曜と連休明けの朝 6〜8 時に全端末で必ず起きる**。切り替え翌朝は全端末が初めてのログインになる。今は匿名でも使えるので、Google に未ログインの端末・スマホがどれだけあるかが分かっていない | 確認済み（期限の仕組み）／推測（端末の実態） |
| S3 | 中 | 罠 20（操作中のログイン切れ）：この画面は返事を待たずに行を「完了済」「承認済」に変える箇所が 9 か所ある。401 のとき画面が嘘の状態のまま移動する。計画は「書きかけを退避」の 1 行だけ | 確認済み |
| S4 | 中 | 移転ページのボタン・ログイン画面は GAS の iframe の中では開けない（`target="_top"` が要る）。ポータルが iframe で埋め込んでいるとログインそのものができない | 確認済み（コード）／推測（ポータル） |
| S5 | 中 | CSP 案 1 の「onclick を機械的に置き換え」は、インラインの on 属性 81 個（onclick 以外に oninput・onchange・onmouseenter/leave）と、JS の文字列で作る onclick、`[onclick*=…]` で探すコード（`approver.html:876`）を含む | 確認済み |
| S6 | 中 | シムの取り決めが「失敗は Error」だけ。実際は成功の返事の中で失敗を返す形が 5 種類あり、日時の文字列の形も関数ごとに違い、画面が正規表現で時刻を切り出している | 確認済み |
| S7 | 中 | スマホの配置（罠 16）：計画は「実機確認」までで、「今と同じ見た目（PC 幅の縮小表示）を保つ」か「スマホ向け配置を初めて効かせる」かを決めていない | 確認済み（コード） |
| S8 | 中 | XSS：`admin.html:619` の手動入力ボタンは、作業員名を JS 文字列の中に `'` を逃がさずに入れる。作業員名は申請時に**画面から送った文字列**がそのまま保存される（マスタと照合しない）。S1 を直すとこの穴が初めて動く | 確認済み |
| S9 | 低 | 「画面の 30 関数」は `api_getFormUrl` を含む数。4-3 の公開リストは 29 関数になる | 確認済み |
| S10 | 低 | ページの `<title>`、権限エラーの文の改行、CSV の文字化け、GAS 由来の日本語エラー文など細かな差（本文 7 章） | 確認済み／一部推測 |

## 1. 画面から呼ぶ関数の全数（4-3 の移植表）

`grep google.script.run` で 40 か所（top 18・approver 6・approve2 3・admin 7・staffSummary 6・no_auth 0）。重なりを除いて 30 関数。計画 3-1 の「40 か所・30 関数（申請者 13・承認者 6・二次承認 3・総務 7・スタッフ集計 1）」は**正しい**。

| 画面 | 行 | 関数（引数） | 引数の型の注意 | 成功の返事で失敗を返す形 |
|---|---|---|---|---|
| top | 688 | `api_getDeptList()` | — | — |
| top | 685 | `api_getWorkerInfo()` | — | 一致しないと `null`（画面は何もしない） |
| top | 723 | `api_getWorkersByDept(dept)` | 返事は `"コード 氏名"` の文字列配列。画面が空白で割る | — |
| top | 792（死んだコード）・1816 | `api_getHolidayCandidateDates()` | — | — |
| top | 837（死んだコード）・1763 | `api_getOrderList()` | — | — |
| top | 970（死んだコード） | `api_getFormUrl(type, dept, workerLabel, targetDateStr, orderCodes)` | 移さない（決定 #11） | — |
| top | 1015 | `api_getDashboard()` | — | `{error, today:'', overtime:[], holiday:[]}`（`worklog.gs:532-536`）。null なら「res=null」表示 |
| top | 1310・1336 | `api_markOvertimeDone(id[, 'HH:mm'])` | 2 つ目は**省略あり** | — |
| top | 1352・1375 | `api_markHolidayStart(id[, 'HH:mm'])` | 同上 | — |
| top | 1400・1426 | `api_markHolidayDone(id[, 'HH:mm'])` | 同上 | — |
| top | 1521 | `api_editRequest(id, {hours, targetDate?})` | オブジェクト | `res.ok` が偽なら「更新に失敗しました。」 |
| top | 1544 | `api_deleteRequest(id)` | — | `res.ok` が偽なら「削除に失敗しました。」 |
| top | 2029 | `api_submitRequest(payload)` | オブジェクト。`isRetroactive` は真偽値 | `{ok:false, error}`。`res.success===true` も成功扱い |
| approver | 486 | `api_getApproverDepts()` | — | 非承認者は例外 |
| approver | 567 | `api_getApproverDashboard(depts[])` | **配列** | `{error, isAdmin:true …}`（`approval.gs:274-277`。失敗時に isAdmin が true になる癖） |
| approver | 799 | `api_approveRequest(id)` | — | 返事を見ない |
| approver | 820 | `api_rejectRequest(id)` | — | `{ok:false, message}`（`error` ではなく `message`） |
| approver | 899 | `api_approveRequestsBatch(ids[])` | **配列** | `results[].ok/error` |
| approver | 959 | `api_approverMonthlySummary(ym, depts[])` | **配列** | — |
| approve2 | 281 | `api_getSecondaryApprovalList(ymd)` | — | — |
| approve2 | 422 | `api_secondaryApprove(id)` | — | `{ok:false, message}` |
| approve2 | 455 | `api_secondaryApproveBatch(ids[])` | **配列** | `{ok:false}`／`results[].ok` |
| admin | 482 | `api_adminDeptOptions()` | — | — |
| admin | 599 | `api_adminMonthlySummary(ym, dept)` | — | — |
| admin | 651 | `api_adminDailyDetail(ymd, dept)` | — | **S1** |
| admin | 702 | `api_adminManualComplete(id, end, start)` | 引数の順が end→start | — |
| admin | 716 | `api_adminFiscalTrend(ymd, dept)` | — | — |
| admin | 987 | `api_adminSpecialClauseCount(ymd, dept)` | — | — |
| admin | 1021 | `api_adminPendingWatch()` | — | — |
| staffSummary | 303・458・492 | `api_getWorkersByDept(dept)` | — | — |
| staffSummary | 319 | `api_getClosingSummary(ym, dept, token)` | 空文字を送る | null なら「データがありません」 |
| staffSummary | 496・499 | `api_getWorkerInfo()`・`api_getDeptList()` | — | — |

死んだコード（確認済み）: `showDatePicker` はどこからも呼ばれない（grep で定義だけ）。そこから先の `showOrderPicker`・`onManualDateSelect`・`submitOrderPicker`・`doOpenForm`・`showFormLink`・工番選択モーダル（`top.html:415-460`）・休日日付モーダル（586-604）も画面からは届かない。`include_`（`adminApi.gs:601`）はどの HTML からも使われていない。

## 2. 指摘の詳細

### S1（高）総務の日次一覧は今の本番で表示されていない見込み

- 根拠（確認済み）: `api_adminDailyDetail`（`adminApi.gs:100-124`）は `buildJoinedRecords_` の行をそのまま返す。その `targetDate` は `toDate_()` の **Date オブジェクト**（`adminApi.gs:57・82`）、`pdfGeneratedAt` もシートの生値（Date）。
- `google.script.run` は Date を含む返り値を渡せず、成功ハンドラに `null` が届く（Apps Script の既知の制約）。同じコードベースのほかの関数は「Date→文字列変換（GASシリアライズ問題回避）」とコメントして文字列にしている（`worklog.gs:438`、`approval.gs:189`）。作者がこの制約を知っていて、日次一覧だけ漏れた形。
- その結果（推測）: 対象日に 1 件でも申請があると `res` が null → `res.items.forEach`（`admin.html:609`）で例外 → 表は空のまま、alert も出ない。0 件の日だけ「正常」に見える。これに連なって次の 3 つも本番では使えていない見込み。
  - 日次表の「手動入力」ボタンと手動実績入力モーダル（`admin.html:617-620・657-703`）→ `api_adminManualComplete` は画面から届かない。
  - `?page=admin&date=…` で開いたときの日次の自動表示（`admin.html:1030-1032`）。
  - 日次表の過去申請バッジ（`admin.html:627-636`）。そもそも `api_adminDailyDetail` は `isRetroactive`・`daysAgo` を返さないので、日次が動いてもバッジは出ない（確認済み）。
- 計画: 3-1 は日次一覧と手動入力を「ある機能」として扱い、6-4 B-17 で「総務の月次と日次」を確かめるとだけ書く。この癖に触れていない。4-2 の「JSON は GAS の形で返す」を守ると、VPS では日次が**初めて表示される**。
- なぜ大事か: 「機能も癖も変えない」の判断が分かれる。VPS で表示されると、総務が見たことのない表と、今まで使えなかった手動入力（WorkLogs と PDF と台帳を書き換える操作）が急に使えるようになる。
- 計画の直し方:
  1. 段階 1 で本番の総務画面（読むだけ）を開き、申請のある日で［日次を表示］を押して、表が空のままかを確かめる。ブラウザのコンソールに `Cannot read properties of null (reading 'items')` が出るかも見る。
  2. 9-2 に質問を足す。「日次一覧と手動入力は今動いていない見込み。VPS で動くようにしてよいか」。おすすめは「動くようにする」。ただし S8 の XSS を先に直すことを条件にする。
  3. 動くようにするなら、`targetDate`・`pdfGeneratedAt`・`actualStartAt/EndAt` を文字列で返す形を仕様に書き、「変えるもの」の一覧（3-3）に載せる。
  4. 5-4 の総当たり比較で、GAS 側の日次は null になることを前提に比較から外す。

### S2（高）共有アカウントのログイン切れと朝のピーク

- 計画 4-4・8-2・Q7 は「12 時間だとピークにログインし直しが起きる見込み（推測）」「無操作の期限を延ばす」とする。
- 確認済みの事実: スキル 02 の土台は `SESSION_IDLE_HOURS`（既定 12）と `SESSION_MAX_AGE_DAYS`（既定 14）の 2 つの期限を持つ。計画は絶対期限 14 日に触れていない。
- そこから言えること（計算）:
  - 夕方 18 時に最後に使った端末は、翌朝 6 時で 12 時間に達する。平日でも境目にある。**金曜の夕方から月曜の朝は約 60 時間**なので、無操作 12 時間なら毎週月曜の 6〜8 時に、全部署の端末が必ずログインし直しになる。連休明けも同じ。
  - 絶対 14 日の期限は、無操作の期限をいくら延ばしても 2 週間ごとに来る。来る時刻はログインした時刻で決まるので、朝のピークに当たりやすい（朝に最初にログインするため）。
  - 切り替えは夜に行う（6-1）。セッション表は切り替えで空にする（B-9）。**切り替え翌朝の 6〜8 時は、全端末が初めてのログイン**になる。新しい OAuth クライアントの初回は、アカウントの選択や同意の画面が出る見込み（推測）。
- 分かっていないこと（推測）: 今の公開設定は「誰でも（匿名）」。申請者（TOP）が使う端末やスマホが Google にログインしているかは記録が無い。手がかりはある。WorkLogs の `updatedBy` は操作した人のメールで、空なら `'unknown'`（`worklog.gs:180・241・309`）。Requests の `approvedBy`・`modifiedBy` も同じ（`approval.gs:281・349`、`worklog.gs:595`）。`'unknown'` が多ければ、Google にログインしていない端末で使われている。そうした端末では、共有アカウントのパスワード（と 2 段階認証）を現場の人が知らないと申請できない。
- 計画の直し方:
  1. 段階 1 で実データを数える（読むだけ）。WorkLogs の `updatedBy` と Requests の `approvedBy`・`modifiedBy` を値ごとに数え、`'unknown'` の割合と時間帯を出す。共有アカウント 6 つ以外のメールが出るかも見る。
  2. オーナーに、申請に使う端末（部署の PC・タブレット・個人のスマホ）と、それぞれがどのアカウントで Google にログインしているかを聞く（Q7 に足す）。
  3. Q7 の案を具体にする。無操作の期限は「7 日」以上か、使うたびに延びる形にする。絶対期限は 30〜90 日にする。値は〇〇日（オーナーが決める）。
  4. 切り替えの前日の業務中に、各部署の端末で新しい URL を 1 回開いてログインしてもらう手順を足す。セッション表は B-9 で空にするので、そのログインは「ステージングと同じ OAuth クライアントでの同意だけ済ませる」目的になる。翌朝のログインは押すだけで終わる形にする。または、切り替えを朝のピーク後の時間帯にずらし、当日の業務中に全端末でログインしてもらう。
  5. B-17 は「オーナーと共有アカウント 1 つ」だけ。部署の全アカウント（13 部署）がログインできることを切り替え前に確かめる。
  6. ログインに `prompt=consent` や `select_account` を付けない（毎回の選択画面を出さない）ことを仕様に書く。

### S3（中）操作中のログイン切れ（罠 20）とこの画面の楽観更新

- 確認済み: 返事を待たずに行を書き換える箇所は 9 か所。top の `doOvertimeDone`（1292）・`doManualOvertimeDone`（1318）・`doHolidayStart`（1343）・`doManualHolidayStart`（1360）・`doHolidayDone`（1382）・`doManualHolidayDone`（1408）、approver の `doApprove`（789）・`doReject`（807）・`approveAllInTable`（868-873）。approve2 の `singleApprove`（400-402）は内部のデータだけ先に変える。計画の「10 か所」はこれと 1 違うが、数え方の差で大きな問題ではない。
- 今の GAS では、失敗すると alert を出して一覧を読み直す（例 `top.html:1306-1309`）。
- 計画 4-3 は「401 のときは書きかけを退避してログインへ」だけ。このままだと、行が「完了済」「承認済」になった画面のまま移動する。ログインから戻ると一覧は「未完了」に戻る。利用者は押したつもりで、記録されていない。
- 守る書きかけ（このアプリ固有）: 申請モーダル（種別・部署・作業員・日付・過去日の選択・工番 3 件・理由・補足理由・予定時間・業務内容）、編集モーダル、総務の手動実績入力、承認者画面で管理者が選んだ承認者（`SELECTED_DEPTS`。移動すると選び直し）、staffSummary の月・部署・スタッフ。
- 計画の直し方:
  1. シムは 401 のとき、失敗ハンドラの alert を出さずに、楽観更新した行を元に戻すか一覧を読み直してから移動する。移動の前に「ログインが切れました。ログイン後にもう一度押してください」と出す。
  2. 書き込みは自動で送り直さない。冪等キーは操作ごとに 1 つで、401 では捨てない（罠 17）。
  3. `next` には今の URL（`?page=…&date=…` を含む）を入れる。
  4. 申請モーダルを開くときと、画面にフォーカスが戻ったとき（`visibilitychange`）にセッションを確かめ、切れていれば入力を始める前にログインへ送る。朝のピークに入力し終えてから切れる事故を減らす。
  5. 受け入れ表に「各楽観更新の操作で 401 を注入し、画面が嘘の状態で止まらない」を 9 件足す（スキル 02 の mock-server の `status=401`）。

### S4（中）移転ページとログインは iframe の中で開けない

- 確認済み: 4 画面は `<base target="_top">` を持ち、全リンクに `target="_top"` がある。GAS の画面は Google の iframe（googleusercontent.com）の中で動いている。`handleFormRedirect_` は `window.top.location.href` を使う（`adminApi.gs:654`）。罠 29 のとおり、クリックなしの `window.top` 移動は止められることがある。
- 計画 6-3 の移転ページは「`page` と `date` を引き継いだボタン」とだけ書く。ボタンに `target="_top"` が無いと、新しいアプリが GAS の iframe の中で開こうとする。VPS は `frame-ancestors 'self'` の予定（4-3）で、Google のログイン画面も iframe を拒むので、真っ白になる。
- ポータルがこのアプリを iframe で埋め込んでいる場合（Q10、未確認）も同じで、ログイン画面が出せない。計画 Q10 は「埋め込みを禁止するか」だけを聞いている。
- 計画の直し方: 移転ページは `<base target="_top">` とボタンの `target="_top"` を必須にし、自動移動は「動けばよい」扱いにする（罠 29）。Q10 に「埋め込んでいるなら、ログインは新しいタブか最上位の窓で開く必要がある」を足す。B-3 の確認に「GAS の `/exec` から押して、新しいアプリが最上位の窓で開く」を足す。

### S5（中）CSP 案 1 の置き換えの数

- 確認済み（grep）: HTML のインライン on 属性は 81 個。top は onclick 48・oninput 2・onchange 1・onmouseenter 2・onmouseleave 2、approver は onclick 8・onchange 4・oninput 1、approve2 は onclick 3・onchange 2、admin は onclick 10。このほか、JS の中で `onclick="…"` を文字列で組み立てる箇所が多い（`top.html:1187-1243・1344・1361-1366`、`approver.html:734-737`、`approve2.html:321・349`、`admin.html:618-620`）。
- `approver.html:876` は `document.querySelector('[onclick*="approveAllInTable(\'…\')"]')` で一括承認ボタンを探す。onclick を `data-action` に替えると見つからなくなり、「承認処理中…」「一括承認完了」の表示とボタンの無効化が消える（エラーにはならないので気づきにくい）。
- 計画の直し方: 4-3 案 1 の対象を「on 属性すべて（onclick・oninput・onchange・onmouseenter・onmouseleave）と、JS の文字列で作る onclick」と書く。`approver.html:876` は id で探す形に直す（見た目は同じ）。段階 2 で上の数を前提に案 1／案 2 を決める。style 属性は `style-src 'unsafe-inline'`（スキル 02）で今のまま動く。`top.html:304` の `div[style*="grid-template-columns:1fr 1fr"]` も style 属性の文字列に頼っているので、style 属性は書き換えない。

### S6（中）シムの取り決めと返事の形

- 失敗の返し方が 6 通りある（確認済み）。
  1. 例外（`withFailureHandler`）。画面は `err.message || err` を出す。
  2. `{error}` を成功で返す（`api_getDashboard`、`api_getApproverDashboard`）。
  3. `{ok:false, error}`（`api_submitRequest`）。
  4. `{ok:false, message}`（`api_rejectRequest`、`api_secondaryApprove`）。
  5. `results[].ok:false`（一括承認 2 種）。
  6. `null`（`api_getWorkerInfo` の不一致、`api_getClosingSummary`）。
  シムが 2〜6 を 400 や例外に変えると、画面の文言と動きが変わる。計画 4-3 は 1 しか書いていない。
- 日時の文字列の形（確認済み）: 承認者画面の `submittedAt` は `yyyy-MM-dd HH:mm:ss`、二次承認の `submittedAt/approvedAt/approvedAt2` は `yyyy-MM-dd HH:mm`、`targetDate` は `yyyy-MM-dd`。シートの値が文字のときは、その文字が**そのまま**返る（`approval.gs:193-194`）。画面はこれを正規表現で読む。
  - `approver.html:724` `fmtTime` は `yyyy-MM-dd[ T]HH:mm` を読み、`todayDate`（サーバーの今日）と比べて「HH:mm」か「M/D HH:mm」を出す。
  - `admin.html:669` と `adminApi.gs:459` は `actualStartAt` から最初の `HH:mm` を取る。UTC の ISO 文字列（`…T23:00:00Z`）で返すと 9 時間ずれた開始時刻が入る。
  - `top.html:1166` は `targetDate + 'T17:20:00+09:00'`、`top.html:1170-1171` は `actualStartAt` の空白を `T` にして `+09:00` を足す。
- 省略できる引数（`api_markOvertimeDone(id)` と `(id, '20:00')`）と、配列・オブジェクトの引数（表 1）がある。
- 計画の直し方: 仕様（段階 2）に「関数ごとの引数と返事の形の表」を作り、上の 6 通りの失敗の形と、日時の列ごとの書式（JST、オフセットなし）を書く。シムは HTTP 200 の本文をそのまま成功ハンドラに渡し、例外（HTTP 400/500）のときだけ失敗ハンドラに `message` 付きの Error を渡す。GAS と VPS で同じ入力の返事を突き合わせる試験を 29 関数ぶん作る（5-4 の総当たりに含める）。

### S7（中）スマホの配置を「今と同じ」にするか決めていない

- 確認済み: 6 つの HTML はすべて `<meta name="viewport" content="width=device-width, initial-scale=1">` を書いているが、`doGet` は `addMetaTag` を呼ばない（grep で 0 件）。GAS はこの meta を効かせないので、今のスマホは PC 幅の画面を縮めて表示している（罠 16）。
- 画面ごとの切り替え幅はばらばら（確認済み）: top・approver は 640px、admin・approve2 は 768px と 480px、staffSummary は 820px と 520px。VPS で meta が効くと、スマホでヘッダーが縦積みになり、表の文字が 11〜12px になり、admin と approve2 の表は折り返すようになる（`white-space:normal`）。
- 日付と時刻の入力欄（罠 16 の症状が出る場所）: `top.html:485・596・626`、`approve2.html:176`、`admin.html:421・425`（time）、`approver.html:292`（month）。
- 計画 4-3 は「初めて効く。実機確認」まで。どちらにするかを決めていない。
- 計画の直し方: 9-2 に質問を足す。おすすめは「スマホ向けの配置を効かせ、罠 16 の CSS を上の欄に当てる」。今と同じ見た目を保つなら、viewport の meta を外す（PC 幅で縮小表示＝今と同じ）。どちらでも、現場で使う機種で TOP の申請・完了と承認者画面を実機で押す。

### S8（中）XSS：GAS の守りが無くなる箇所

- 確認済み:
  - `admin.html:618-620`: `onclick="openManualEntry('…氏名…', …)"` に `escHtml()` の値を入れている。`escHtml` は `'` を逃がさない（`admin.html:461-463`）。氏名に `'` があると JS の文字列から抜けられる。
  - 作業員名は `api_submitRequest` の `workerLabel`（画面から送る文字列）を空白で割ったものを、マスタと照合せずに保存する（`formSubmit.gs:682・738-740`、`parseWorkerChoice_` 392-399）。部署も同じ。誰でも任意の名前で申請を作れる（決定 #5 の「信頼する作り」の範囲）。
  - 今は S1 のため日次表が描かれず、この穴は動いていない見込み。S1 を直すと動き出す。
  - GAS では画面が Google の別ドメインの iframe で動き、テンプレートの `<?= ?>` が文脈に合わせて逃がしていた（`admin.html:466`、`approve2.html:256`、`no_auth.html:73`）。VPS では画面とセッションが同じドメインになり、XSS が起きるとその人の権限で RPC を呼べる。
  - ほかの描画は安全側（確認済み）: top・approver・approve2・staffSummary は氏名・部署を `escHtml`／`esc` してから文の中に入れる。`top.html:875`（工番の onclick）は `"` を逃がさないが、死んだコードにある。
- 計画 3-4 は「admin の手動入力ボタンの氏名などを直す」と提案しているので方向は合っている。足りないのは、作業員名が画面から送った値だという根拠と、S1 との順番。
- 計画の直し方: 3-4 の提案を「段階 3 の必須」に上げる。手動入力ボタンは `data-*` 属性と 1 つの受け口にする（S5 の作業と一緒にできる）。テンプレートは Jinja2 の自動エスケープを有効にし、JS の中に入れる値は `json.dumps` のあと `<`・`>`・`&` を `<` などに変える（`</script>` 対策）。試験に「氏名が `x');alert(1)//` の申請を作り、総務の日次表で何も起きない」を足す。

### S9（低）公開する関数の数

- 計画 4-3「公開する関数は閉じたリスト: 画面の 30 関数＋管理用。`api_getFormUrl` と旧 API 2 つは出さない」。30 は `api_getFormUrl` を含む数なので、出すのは 29 関数になる。
- 計画の直し方: 4-3 と 6-2（受け入れ「30 関数」）、7 章（「関数 30」）を「29 関数（`api_getFormUrl` を除く）」に直す。

### S10（低）細かな差

- ページの題名（確認済み）: top・approver・admin・approve2 は `<title>` を持たない。GAS の `setTitle('残業・休日出勤申請')`（`adminApi.gs:596`）が題名になっている。権限エラーは「権限エラー」（561・580）。VPS では 6 ページすべてに `<title>` を書く。Q15（staffSummary）だけでなく全ページの題名を 4-3 に書く。
- 権限エラーの文（確認済み）: `message` は `\n` で区切るが、`.msg` に `white-space:pre-line` が無いので、今は改行されずに 1 行につながって見える（`no_auth.html:50-55・73`）。Q8 でデバッグ情報を消すときも、改行しない見た目は同じにする。
- CSV（確認済み）: `admin.html:756` は `type:'text/csv'`、BOM なし。日本語の Windows の Excel で開くと文字化けする（推測）。計画は「BOM なし」を癖として残すとしている。GAS の iframe の中で `a.download` が今動いているかは未確認（推測。iframe の制限でダウンロードが止まる場合がある）。本番の総務画面で［表示中をCSV］が保存できるかを確かめ、今の状態を受け入れ表に書く。
- GAS が出していた文言（推測）: ロックの待ち切れ（`lock.waitLock(20000)`）や Google の内部エラーのときに画面に出ていた日本語のエラー文は、VPS では出ない。利用者の目に触れるのはエラー時だけなので、受け入れの対象外と明記する。
- 癖の一覧（3-2）に足すもの（確認済み）:
  - 編集モーダルは予定時間を `approvedMinutes/60` の小数 1 桁で選ぶ（`top.html:1461`）。一覧に無い値（例 45 分、4.5 時間）だと何も選ばれず、保存で「予定時間を選択してください。」になる。休日は 480 分以上なら「1日」、それ以外は「半日」。
  - 承認者画面の一括承認は、成功するとボタンが「一括承認完了」のまま押せなくなる。失敗するとボタンの文字が「一括承認」に変わる（元は「残業 全件承認」）（`approver.html:877・882・895`）。
  - 二次承認の単票は、成功すると PDF のリンクを消す（`approve2.html:410`）。読み直すまで出ない。
  - 総務の KPI の見出しは、どの月を選んでも「（今月）」（`admin.html:229・234`）。
  - 総務の手動実績の［登録］はボタンを無効にしない（`admin.html:683-703`）。二度押しできる（VPS は冪等キーで 1 回になる）。
  - 承認者画面の非管理者は、自分のメールと一致する承認者が無いと先頭の承認者を表示する（`approver.html:439-441`）。メールの比較は小文字同士（`approval.gs:539・551`）。

## 3. 計画の記述で正しいと確かめたもの

- 3-1 画面: `?page=` の 5 ページと既定 top、`no_auth`、`?page=redirect`、不明な page は top（`adminApi.gs:535-545`）。`?date=` は approve2 と admin が使う（approve2 は画面で `^\d{4}-\d{2}-\d{2}$` を確かめ、admin は確かめない＝`admin.html:466-469`）。
- 3-1 入口判定: admin・approve2 は管理者だけ、approver は管理者か承認者で、メールが空なら判定しない（`adminApi.gs:548-585`）。top・staffSummary は判定なし。
- 3-1 テンプレート変数: `APP_URL`・`PORTAL_URL`（直書き 3 か所 554・578・589）・`DATE_PARAM`・`REASONS_JSON`・`OT_HOURS_JSON`・`HD_HOURS_JSON`・`message`。`<?!= ?>`（逃がさない）は 3 つの JSON だけ（`top.html:637-639`）。
- 3-1 呼び出しの数: 40 か所・30 関数・`api_getFormUrl` は死んだコードからだけ。
- 3-1 グラフ: canvas の自前描画（`admin.html:766-959`、`approver.html:1069-1160`）。外部の script も CDN も無い。staffSummary のゲージは CSS。
- 3-2 画面の癖: 作業員の自動選択が効かない（`top.html:672` は `options[0].value !== ''` を待つが、先頭はいつも空）。承認者画面の KPI は初回が submitted だけ、再計算が `modified` も数える（`approver.html:577` と `775`）。STEP 2 のサマリーに過去日が出ない（`top.html:1929`）。admin のフィルタは［月次を更新］まで反映しない（部署の select にもイベントが無い）。CSV は BOM なし。approve2 の注記「翌朝バッチで…」（228）。単票の二次承認に confirm なし（395-423）。
- 4-3「今日はブラウザの時計で決まる」: 残業の対象日（`top.html:1714-1718`、ただしサーバーが今日で上書きする＝`formSubmit.gs:710-711`）、過去日カレンダー、4 時間超の警告（1159-1182）、admin の月・日・基準日の初期値、approve2 の初期日、承認者の月、staffSummary の締め月（248-252）。
- Q15: staffSummary の `<title>` は GAS の `setTitle` で「残業・休日出勤申請」に置き換わる（コードで確認。実際のタブの表示は推測）。
- 8-2「`?date=` からの XSS」: `<?= ?>` の文脈エスケープが無くなると、admin の `URL_DATE_PARAM` は JS の文字列に生で入る。`yyyy-MM-dd` だけを通す案で足りる。

## 4. 確かめること（段階 1 で、読むだけ）

- 本番の総務画面で、申請のある日の［日次を表示］が空のままか（S1）。［表示中をCSV］で保存できるか（S10）。
- WorkLogs `updatedBy`、Requests `approvedBy`・`modifiedBy` の値ごとの件数と、`'unknown'` の時間帯（S2）。
- 申請に使う端末と、そのブラウザの Google ログインの状態（オーナーに聞く。S2）。
- 社内ポータルがこのアプリを iframe で埋め込んでいるか（Q10。S4）。
