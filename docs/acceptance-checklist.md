# 受け入れ表（残業・休日出勤申請 VPS 版）

作成: 2026-11-26（担当 C）。もと: `docs/vps-implementation-spec.md` 12-3（A-1〜A-42）と、JEV ハンズでオーナーが毎回確かめる実機の操作。
切り替え（`docs/cutover-runbook.md` の A-1）の前に、すべての行の「状態」が「済」になっていることを確かめます。

## 使い方

- 「確かめ方」の書き方:
  - **自動**: 手元の自動テスト（`cd vps-app/api && .venv/Scripts/python -m pytest -q`。担当と試験のファイル名を書く）。
  - **acceptance**: `vps-app/tools/qa/acceptance.py` をステージングで流したときの項目の番号（`H-`・`E-`・`P-`・`F-`・`O-`・`SB-`・`W-`）。
  - **JEV**: JEV ハンズでオーナーが実機で押して確かめる操作（スキル 05 §4-2）。
  - **SB**: サンドボックス（DB のコピーと紐づく GAS の写し）で新旧を突き合わせる。
  - **実機**: iPhone・Android の実物（JEV では動かせないもの）。
- 「状態」は 未／済／NG のどれかにし、済にした日と、確かめた人を「メモ」に書きます。
- 報告では「実機で押した／自動だけ／まだ」を分けて書きます。自動のテストが通っても、JEV の行は済になりません。

acceptance.py の動かし方（ステージング。読むだけ → 書き込みの順）:

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --mint-session --expect-env staging"
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --mint-session --expect-env staging --write"
EOF
```

## 1. 実装仕様 12-3 の受け入れ表

| # | 項目 | 確かめ方 | 状態 | メモ |
|---|---|---|---|---|
| A-1 | 画面 5 ページ＋権限エラー、`?page=`、不明は top、`redirect` | 自動 U `test_pages.py`。acceptance P-1〜P-5、E-1・E-2。JEV: 各ページを開き、ヘッダーのリンクを全部押す（J-1） | 未 | |
| A-2 | 入口の判定（admin・approve2・approver） | 自動 U `test_pages.py`。JEV: 承認者でない共有アカウントで approver・admin を開き、権限エラー（J-14） | 未 | |
| A-3 | テンプレート変数・`DATE_PARAM` | 自動 U。JEV: `?page=approve2&date=2026-11-25`・`?page=admin&date=2026-11-25` で、その日が自動で出る（J-5・J-6） | 未 | |
| A-4 | 29 関数の形・エラー文・失敗の 6 通り | 自動 W `test_domain_*.py`。acceptance F-*（39 関数すべて）、F-all、E-5〜E-9。SB: 同じ入力で GAS と VPS の返事を突き合わせる（日次一覧は GAS が null なので比べない） | 未 | |
| A-5 | Sebastian の 4 action | 自動 W `test_sebastian.py`。acceptance SB-1〜SB-5。SB: 写しの GAS の転送を通して `ping`・完了。切り替え後に本物の Sebastian の完了 1 件（手順書 D 章） | 未 | |
| A-6 | 状態の遷移・削除の副作用 | 自動 W。acceptance W-3〜W-16。JEV: 申請 → 承認 → 完了 → 二次承認 → 削除を 1 件通す（J-2〜J-9） | 未 | |
| A-7 | 実績と休憩、Sebastian の休憩（#23） | 自動 W `test_domain_breaks.py`・`test_sebastian.py`。acceptance W-6（17:20〜18:20 で休憩 0） | 未 | |
| A-8 | 集計（月次・年度推移・特別条項・滞留・スタッフ集計） | 自動 W `test_domain_summaries.py`。acceptance F-api_adminMonthlySummary ほか。SB: 総当たり比較（全月×全部署、全締め月×全員） | 未 | |
| A-9 | PDF の見た目 | SB: スパイクの 10 件を本番の PDF と画素ごとに比べる。オーナーが実物を見る | 未 | |
| A-10 | PDF を作る場面×方式×既存の PDF | 自動 G `internal.test.js`・M `test_gas_jobs.py`。acceptance W-8・W-11・W-17。SB: 各きっかけを 1 件ずつ通し、PDF と台帳の行を見る | 未 | |
| A-11 | 台帳 | SB: 台帳の行を GAS が書いた行と 1 セルずつ比べる。見出しを壊したコピーで `ledger_header_mismatch` | 未 | |
| A-12 | 朝バッチと BatchLogs | 自動 M `test_morning_batch.py`。SB: 前日分 3 件（1 件は PDF あり）で流し、BatchLogs の 1 行・台帳・7:05 の確かめ | 未 | |
| A-13 | マスタ同期と DB の 4 タブ（#21） | 自動 D `test_master_sync.py`。acceptance F-ops_runMasterSync（--write）。SB: GAS の `syncAllMasters` の後と VPS の `sync_masters` の後のタブを 1 セルずつ比べる | 未 | |
| A-14 | 承認者の表と管理者 | 自動 F `test_permissions.py`・D `test_config_sync.py`。acceptance F-api_getApproverDepts、F-ops_syncSheetConfigNow（--write）。手順書 B-16 で DeptApprovers が変わっていない | 未 | |
| A-15 | トリガーと旧フォーム（#11） | 自動 G `cutover.test.js`。手順書のリハーサル（B-2〜B-7） | 未 | |
| A-16 | ロックと同時実行 | 自動 W `test_domain_concurrency.py`（postgres）。SB: 完了済み 20 件の一括承認で台帳に 20 行が重ならない | 未 | |
| A-17 | 今のまま移す癖（計画書 3-2） | 自動 W（項目ごとに 1 試験）。JEV: 作業員の自動選択が効かない、STEP 2 のサマリー、CSV（J-1・J-15） | 未 | |
| A-18 | 直すと決めたもの（計画書 3-3、11-1） | 自動（各担当）。acceptance F-api_adminDailyDetail・W-7。JEV: 日次一覧が出る・手動入力ができる（J-10） | 未 | |
| A-19 | **06:30 JST の試験**（決定 C） | 自動 W `test_domain_fiscal_jst.py`。SB: 朝 6:30 台に staging で通常残業を 1 件出し、targetDate とシートの行を見る | 未 | |
| A-20 | **二重押し**（#5） | 自動 F `test_rpc.py`・U `shim.test.js`・W。acceptance W-2。JEV: 申請の送信を素早く 2 回押して 1 件（J-12） | 未 | |
| A-21 | **2 つのタブでの承認** | 自動 W `test_domain_concurrency.py`。acceptance W-4・W-14・W-17（承認の仕事が 1 つ）。JEV: 2 つのタブで同じ申請を承認し、2 回目が「既に承認済みです。」相当、PDF は 1 つ（J-13） | 未 | |
| A-22 | **Safari の日付** | 自動 U `test_static_assets.py`。実機: iPhone で TOP の申請・完了・期限超過の表示、承認者画面（オーナーの写真） | 未 | |
| A-23 | 401 の注入 9 件 | 自動 U `shim.test.js`。mock-server の `status=401` で楽観更新の 9 操作を手で押す。JEV（J-17） | 未 | |
| A-24 | 移転ページと iframe | 自動 G。リハーサル: GAS の `/exec?page=approve2&date=…` のボタンで、新しいアプリが最上位の窓で開く（手順書 B-5） | 未 | |
| A-25 | CSP で全ボタンが動く | 自動 F `test_web_config.py`・U。acceptance E-10・E-11・P-1〜P-5（インラインの script と on 属性が無い）。JEV: 5 ページの全ボタン・入力欄・マウスを載せる表示 | 未 | |
| A-26 | スマホ（#25） | 自動 F `test_web_config.py`（`mobile.css`）。JEV: 狭い幅（J-11）。実機: iPhone と現場の Android で TOP の申請・完了、承認者画面の承認、二次承認。PC（1366px）で今の画面と見比べて差が無い | 未 | |
| A-27 | XSS | 自動 U `test_static_assets.py`。JEV: 3-8 の氏名の申請を作り、5 画面を開く（J-16） | 未 | |
| A-28 | 題名・改行・CSV | 自動 U。acceptance P-1〜P-5（題名）。JEV: ［表示中をCSV］が保存できる（J-15） | 未 | |
| A-29 | 行の反映の型（決定 A） | SB: VPS が足した行・直した行を写しの GAS の `getValues()` で読み、日時の列が `Date`。`getRequestById_`・`computeIsRetroactive_`・承認者画面の提出列・未承認滞留が元の値と同じ | 未 | |
| A-30 | 行の位置と並び・消さない | 自動 D `test_row_sync.py`。SB: 人が行を挿入したコピーで、探し直しと止まり方 | 未 | |
| A-31 | 毎晩の照合（決定 A） | 自動 D `test_reconcile.py`。SB: 移行の直後に `--check` で差 0（手順書 B-16） | 未 | |
| A-32 | GAS が止まったとき（決定 B） | SB: 写しの GAS の token を外して操作し、画面は動き、仕事が再試行され、管理者の欄に出て、戻すと流れる。acceptance O-1・O-2 | 未 | |
| A-33 | ステージングから本番に書かない（9 章） | 自動 F `test_config_safety.py`。手順書 B-14 の安全装置の確かめ（NG 0 行） | 未 | |
| A-34 | Settings の重なり・時間帯 | 自動 D・F。段階 2 で数えた結果を `.env` と照合（手順書 A-4・B-12） | 未 | |
| A-35 | 印鑑の画像の例外・前の PDF の持ち主が違う | SB: 持ち主が部署のアカウントの PDF を持つ申請の編集・二次承認 | 未 | |
| A-36 | 再試行で PDF が 2 つできない | 自動 G（`jobId` の再生）・M `test_gas_jobs.py` | 未 | |
| A-37 | `PDF_MODE` を含むかで判定 | 自動 G（偽物の Settings に本番の文） | 未 | |
| A-38 | 候補日の逃げ道 | 自動 W `test_domain_masters.py`。acceptance F-api_getHolidayCandidateDates | 未 | |
| A-39 | `updatedBy` の決まった値・過去申請の求め方 | 自動 W・D。SB: 月×種別の過去申請の件数を新旧で比べる | 未 | |
| A-40 | 移行の照合（計画書 5-4） | 自動 D `test_migrate_from_sheets.py`。手順書 B-11・B-12 の SQL | 未 | |
| A-41 | 速さ（目安: 一覧 100ms 以下、承認 100ms 以下） | acceptance の最後の行「速さ（中央値 ms）」 | 未 | |
| A-42 | 手順書のリハーサルと戻しの練習 | ディレクターが SB と staging で 2 回（手順書 A-2・C 章） | 未 | |
| A-43 | Sebastian の直接の追記（決定 #26。実装仕様 6-10） | 自動 D2 `test_sheet_intake.py`・`test_reconcile.py`。SB と JEV は S-1〜S-4 | 未 | |

## 2. 入口と安全（acceptance.py が毎回見るもの）

| # | 項目 | 確かめ方 | 状態 | メモ |
|---|---|---|---|---|
| G-1 | `/healthz` が `{ok:true, env}` だけを返し、env が環境どおり | acceptance H-1・H-2・H-3 | 未 | |
| G-2 | ログインなしで `/` は 302 `/auth/login?next=…`、`page` と `date` を保つ | acceptance E-1・E-2 | 未 | |
| G-3 | ログインなしの RPC は 401、`/auth/me` も 401 | acceptance E-3・E-4 | 未 | |
| G-4 | JSON でない本文 415、別サイトの Origin 403、知らない関数 404、出さない GAS の関数 404、配列でない本文 400 | acceptance E-5〜E-9 | 未 | |
| G-5 | CSP・nosniff・X-Frame-Options・Referrer-Policy・HSTS・Permissions-Policy | acceptance E-10（nginx 経由） | 未 | |
| G-6 | 書き込みの試験は staging でだけ動く（production では何も書かずに終了コード 2） | 自動 `test_acceptance_runner.py` | 未 | |
| G-7 | 受け入れのセッションは staging でだけ作れ、最後に消える | 自動 `test_acceptance_runner.py`。acceptance の最後の `session removed: True` | 未 | |

## 3. JEV ハンズで押す主な流れと管理画面（オーナーが毎回聞くもの）

オーナーのアカウント（管理者）と、部署の共有アカウント（承認者）を使います。staging で行い、作る申請の業務内容に【受け入れテスト】と書きます。PC の幅（1366px）で押し、J-11 だけ狭い幅にします。

| # | 操作 | 期待すること | 対応する A | 状態 | メモ |
|---|---|---|---|---|---|
| J-1 | TOP を開き、残業と休日の申請を 1 件ずつ送信する（工番・理由・予定時間を選ぶ。休日は候補日から選ぶ） | 「申請しました」相当の表示。TOP の一覧に出る。1 分以内にステージングのシートの Requests の末尾に出る | A-1・A-6・A-17 | 未 | |
| J-2 | 承認者画面（部署の共有アカウント）で、J-1 の申請を承認する（単票と、2 件以上の一括承認） | 状態が承認済み。KPI の数が変わる | A-6 | 未 | |
| J-3 | TOP で残業の［完了］、休日の［開始］→［完了］を押す | 実績・休憩・正味の分が出る。数十秒〜で PDF のリンクが出る | A-6・A-7・A-10 | 未 | |
| J-4 | 二次承認の画面（`?page=approve2`）で、単票と一括の二次承認をする | 一覧から消える。PDF が作り直される（リンクが新しくなる） | A-6・A-10 | 未 | |
| J-5 | 二次承認の画面を `?page=approve2&date=2026-11-25` で開く | その日の一覧が自動で出る | A-3 | 未 | |
| J-6 | 総務の画面（`?page=admin`）で月次サマリー・年度推移・特別条項・滞留監視を見る。`?page=admin&date=…` も開く | 数字がステージングの GAS の写しと同じ。［月次を更新］まで絞り込みが反映しない（今のまま） | A-3・A-8・A-17 | 未 | |
| J-7 | 総務の**日次一覧**を開く | 一覧が表示される（GAS では出なかった。#22）。過去申請のバッジが出る | A-18 | 未 | |
| J-8 | 申請の編集（予定時間・休日の日付）と削除を、TOP から 1 件ずつ | 編集が反映する。削除した申請は取消になり、PDF と台帳の行が消える | A-6 | 未 | |
| J-9 | スタッフ集計（`?page=staffSummary`）で月・部署・スタッフを選ぶ。一覧の行を押して個人の詳細へ | 締め月の集計が出る。個人の詳細に切り替わる | A-8 | 未 | |
| J-10 | 総務の日次一覧の［手動入力］で、完了の時刻を入れる | 実績が入り、PDF が作られる | A-18 | 未 | |
| J-11 | ブラウザの幅を狭く（スマホ幅、390px 前後）して、TOP の申請・完了、承認者画面の承認、二次承認を押す | 横にはみ出さない。日付の欄が枠に収まる。ボタンが押せる大きさ。PC 幅に戻すと今と同じ見た目 | A-26 | 未 | |
| J-12 | 申請の［送信］を素早く 2 回押す | 申請は 1 件だけ | A-20 | 未 | |
| J-13 | 同じ申請を 2 つのタブで開き、両方で［承認］を押す | 2 回目は「既に承認済みです。」相当。PDF は 1 つ | A-21 | 未 | |
| J-14 | 承認者でない共有アカウントで `?page=approver` と `?page=admin` を開く | 権限エラーの画面（デバッグ情報は出ない） | A-2 | 未 | |
| J-15 | 総務の［表示中をCSV］、TOP の STEP 2 のサマリー、作業員の自動選択 | CSV が保存できる（BOM なし）。STEP 2 に過去日が出ない・自動選択が効かない（今のまま） | A-17・A-28 | 未 | |
| J-16 | 氏名が `x');alert(1)//` と `"><img src=x onerror=alert(1)>` の申請を作り、5 画面を開く | 何も実行されず、文字のまま見える | A-27 | 未 | |
| J-17 | 操作中にログインを切る（管理者の欄の［セッションの取り消し］か、別のタブでログアウト）→ 申請のモーダルで送信 | 「ログインが切れました…」の表示 → ログイン → 書きかけが戻る。二重に alert が出ない | A-23 | 未 | |
| J-18 | 総務の画面の一番下の「運用（管理者）」欄を開き、［共有タブを今すぐ取り込む］［マスタ同期］［シートの取り込み（Sebastian の追記）］を押す | 結果が出る。失敗した仕事・反映待ちの数、`sheetIntake` の最後の回が見える | A-13・A-14・A-32・A-43 | 未 | |
| J-19 | 5 ページのヘッダーのリンク・全ボタン・マウスを載せる表示を一通り押す | どれも動く（CSP で止まらない）。ブラウザのコンソールに CSP のエラーが出ない | A-1・A-25 | 未 | |
| J-20 | 古いリンク（GAS の `/exec?page=approve2&date=…`）を開き、移転ページのボタンを押す（切り替えのリハーサルで） | 新しいアプリの同じ画面が最上位の窓で開く | A-24 | 未 | |

## 3a. Sebastian の直接の追記（決定 #26。ステージングの写しで。切り替えの後に本番でも 1 件）

Sebastian と同じ呼び方（`values().append(range='Requests!A:A', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS')`、`status='submitted'`、`submittedAt` は JST の `M/D/YYYY H:MM:SS`）で、ステージングの写しの Requests に 1 行足して確かめます。本番では、切り替えの後に Sebastian が作った最初の申請で S-1〜S-3 を見ます。

| # | 操作 | 期待すること | 対応する A | 状態 | メモ |
|---|---|---|---|---|---|
| S-1 | Sebastian と同じ呼び方で Requests に 1 行足し、TOP（その部署の共有アカウント）と承認者画面を開く | **2 分以内に**一覧に出る（worker の 1 分ごとの取り込み。急ぐときは管理者の欄の［シートの取り込み（Sebastian の追記）］）。管理者の欄の `sheetIntake.ingested` にその requestId | A-43 | 未 | |
| S-2 | S-1 の申請を承認者画面で承認する | 状態が承認済み。シートの**同じ行**の status・approvedAt・approvedBy だけが変わり、新しい行はできない。ほかの列（submittedAt の表示など）は変わらない | A-43 | 未 | |
| S-3 | S-1 の申請の完了を Sebastian（または `/api/sebastian` の `completeOvertime`）で送る | 完了になり、実績が出る。WorkLogs に行ができる。PDF ができる（承認済みなら） | A-5・A-43 | 未 | |
| S-4 | requestId が空の行、同じ requestId の 2 行を足す | 取り込まれず、管理者の欄の `sheetIntake.skipped` に行番号と理由が出る。シートのその行は変わらない（次の照合の後も） | A-43 | 未 | |

## 4. 実機（JEV では動かせないもの）

| # | 操作 | 期待すること | 対応する A | 状態 | メモ |
|---|---|---|---|---|---|
| M-1 | iPhone（Safari）で TOP の申請・完了、期限超過の表示、承認者画面の承認 | 日付が正しく出る（Invalid Date にならない）。オーナーが写真を送る | A-22・A-26 | 未 | |
| M-2 | 現場の Android で TOP の申請・完了、二次承認 | 横にはみ出さず、押せる | A-26 | 未 | |
| M-3 | スマホで初めて開く | Google のアカウント選択が 1 回出て、会社のアカウントでそのまま使える | — | 未 | |
