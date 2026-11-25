# レビュー指摘の修正の確かめ（高・中）

- 日付: 2026-11-26
- 対象: VPS `github-repo` の `git diff 8db13b3..87dcfc5 -- vps-app docs`、GAS `wt-team-g` の `git diff bc7024a..93e9909 -- src tools`
- 見たもの: `review-correctness.md`・`review-security.md`・`review-data-safety.md` の高と中（17 件）と、各報告の末尾の「修正の結果 (team R)」、`wt-team-g/docs/reviews-gas-fixes.md`
- 読むだけの確かめです。コードは変えていません。古いコードでの試験は、作業用フォルダ（スクラッチ）に取り出した写しで動かしました。

## 結論

17 件のうち、直ったと確かめたものが 14 件、一部だけが 1 件、直していないものが 2 件です。直していない 2 件は、どちらも「後で決める」と報告に書かれており、報告との食い違いはありません。

ただし、GAS 側の修正に **GAS と動きが変わる退行が 1 件（中）** あります。完了の時刻が開始より前か、24 時間を超えると、PDF と台帳の行が作られなくなります（R-1）。切り替えの前に直すことをすすめます。

| 判定 | 件数 |
|---|---|
| VERIFIED（直った） | 14 |
| PARTIAL（一部） | 1 |
| NOT FIXED（直していない） | 2 |
| 見つけた退行 | 中 1・低 4 |

## 試験の結果（今回、手元で実行）

| 確かめ | 結果 | コマンド | 最後の行 |
|---|---|---|---|
| VPS の全試験 | 合格 | `cd vps-app/api && .venv/Scripts/python -m pytest -q` | `1062 passed, 1 skipped, 1 warning in 15.10s`（skip は前からある PostgreSQL の結合試験 1 件） |
| GAS の試験 | 合格 | `node tools/gas-harness/run.js`（wt-team-g） | `70/70 passed` |
| 新しい VPS の試験を古いコード（8db13b3）で | 53 件が失敗 | 古い `app`・`tools` に新しい `tests` を重ねて実行（conftest が読む `write_gate` だけ空の形で足した） | `53 failed, 914 passed, 70 skipped, 1 error` |
| 新しい GAS の試験を古いコード（bc7024a）で | 20 件が失敗 | 古い `src` に新しい `tools` を重ねて実行 | `50/70 passed, 20 FAILED` |

下の表の「落ちる試験」は、上の「古いコードで」の実行で実際に落ちた試験です。修正が無いと落ちることを確かめています。

## 指摘ごとの判定

### 正しさ

| 指摘 | 判定 | 根拠 |
|---|---|---|
| 中-1 対象日が読めないときの文 | VERIFIED | 直し方どおりコードは今のまま（`vps-app/api/app/domain/worklog.py:65`）。仕様に受け入れる差として追記（`docs/vps-implementation-spec.md:643`）。仕様書の直しなので試験は無し |

### セキュリティ

| 指摘 | 判定 | 根拠 |
|---|---|---|
| 高-1 申請の文字が台帳・PDF のセルで式になる | VERIFIED | 式になる場所（GAS）を直した。`src/config.gs` の `sheetLiteral_` を、`src/mail.gs` の `buildAccRow_`（upsert と前日の書き直しの両方が通る）と `rebuildAccumulationSS`、`src/pdfExport.gs:206`・`454`・`456`・`512-514`・`519`・`526` で通す。VPS はマスタ同期で `= + - @` の文字に `'` を付ける（`services/master_sync.py` の `formula_safe`）。Requests タブは前から `stringValue`。落ちる試験: GAS `sheetLiteral_: …`・`buildAccRow_: …`・`fillPdfTemplate_: …`・`XLOOKUP path …`、VPS `test_formula_safety.py::test_master_sync_quotes_formula_leads[*]`・`test_formula_safe_rules`。VPS の入口で拒む案は、GAS と同じ動きを保つため採っていない（書く側で止めているので効き目は足りる） |
| 中-1 Sebastian の token・例外の文 | VERIFIED | `config.py` で `sebastian_api_token` を 32 字以上にし、`GAS_INTERNAL_TOKEN` と同じ値を拒む。`domain/sebastian.py` の想定外の例外は `internal error` だけ返す。nginx で `/api/sebastian` を 1 IP 1 分 30 回。落ちる試験: `test_short_sebastian_token_refuses_to_start[staging/production]`・`test_sebastian_token_must_differ_from_gas_internal_token`・`test_unexpected_exception_does_not_leak_internal_text`・`test_nginx_sebastian_and_auth_limits_and_callback_not_logged`。起動を拒む条件が増えた点は R-2 |
| 中-2 GAS の内部入口で任意のファイルをゴミ箱へ・PDF の時間を偽れる | VERIFIED（退行あり） | `src/Internal.gs` の `vpsTryTrash_` は、`PDF_ROOT_FOLDER_ID` の下（親を 3 段まで）にある `application/pdf` だけを捨てる。違えば `trash.status='skipped'`。PDF は「根→日付フォルダ→PDF」に作られる（`pdfExport.gs:150-158`）ので、本番の PDF は今までどおり捨てられる。`workLog` は `vpsValidWorkLog_` で形と範囲を確かめる。落ちる試験: GAS `trash guard: …`（2 件）・`workLog override: …`。ただし範囲の確かめが VPS の正しい依頼まで拒む（R-1） |
| 中-3 委任トークンが Drive 全部の権限 | NOT FIXED | 報告のとおり後で決める（控えの道具が Drive を使うため）。`google_delegate.py` のスコープは変わっていない |
| 中-4 90 日のセッションを止める手段 | PARTIAL | ログインの成功・失敗をログへ（`auth.py` の `login ok`・`login rejected`）。全員の取り消し手順を `vps-app/docs/RUNBOOK.md` に追記。端末の一覧と、全員を消す運用関数は無い（報告どおり） |
| 中-5 依存の既知の脆弱性 | VERIFIED | `requirements.txt` とインストール済みの版が一致。fastapi 0.165.0・starlette 1.10.0・PyJWT 2.15.1・cryptography 50.0.2・requests 2.34.2・anyio 4.16.0 で、どれもレビューの「直る版」以上。全試験が合格。`pip-audit` は環境に無いため再実行していない（依存の依存は未確認） |

### データ安全

| 指摘 | 判定 | 根拠 |
|---|---|---|
| 高-1 ステージングの GAS が本番の PDF を捨てる | VERIFIED | 2 か所で止める。VPS: staging の移行は `pdf_file_id`・`pdf_folder_id` を空にし、元の値を `source_pdf_*` へ（`tools/migrate_from_sheets.py` の `blank_pdf_ids_for_staging`、alembic `0002`）。production は今のまま。GAS: 上の `vpsTryTrash_` のフォルダ確認（ステージングの `PDF_ROOT_FOLDER_ID` の下に本番の PDF は無い）。落ちる試験: `test_staging_migration_blanks_production_pdf_ids`・`test_production_migration_keeps_pdf_ids`・GAS `trash guard: …` |
| 高-2 照合が PDF の 3 列を空にする | VERIFIED | `services/gas_jobs.py` の `pdf_result_to_record` で、`skipped` でも `pdfFileId` があれば記録する（朝バッチも同じ）。`services/reconcile.py` の `split_protected_diffs` で、「シートに値・DB が空」の PDF の列は書かずに `pdfKeptOnSheet` へ出す。落ちる試験: `test_has_pdf_skip_from_a_job_is_recorded`・`test_has_pdf_skip_in_morning_batch_is_recorded_and_counted_as_skip`・`test_reconcile_never_blanks_pdf_columns_on_sheet`・`test_skipped_without_file_id_is_not_recorded` |
| 高-3 移転モードと VPS の書き込みが連動しない | VERIFIED | `services/write_gate.py`。行の反映（`row_sync.py:603`）・照合（`reconcile.py:285`）・マスタ同期（`master_sync.py:267`）が書く前に確かめる。本番は、10 分以内に合格した ping があり、`migrationMode is True` のときだけ書く。止めた行の反映は 30 秒ごとに確かめ直し、関所が開けば自動で戻る（`row_sync.py:629-631`、`GATE_RECHECK_SECONDS`）ので、止まったままにはならない。GAS の ping は毎回プロパティを読み直し、読めなければ `migrated:false`。落ちる試験: `test_write_gate.py` の 10 件、GAS `ping: migrated is always a boolean…`・`ping: script properties unreadable…` |
| 中-1 照合が GAS の仕事の途中の行を比べる | VERIFIED | `reconcile.py` の `gas_busy_request_ids` で、待ち・実行中の仕事がある申請を比べない。落ちる試験: `test_reconcile_skips_requests_with_running_gas_jobs` |
| 中-2 blocked の行が戻しで失われる | VERIFIED | `row_sync.reappend_blocked` と CLI の `reappend`。手順書 C-2 に「blocked も 0 に」（`docs/cutover-runbook.md:590`）。C-2 は GAS が移転モードのまま行うので、関所に止められずに流しきれる。落ちる試験: `test_reappend_blocked_row_writes_a_new_row_at_the_end`・`test_cli_status_and_reappend` |
| 中-3 冪等の記録がプロパティの上限に当たる | VERIFIED | GAS: 100 件・1 件 3,000 字まで。残せなければ `idempotencyStored:false`（`Internal.gs` の `vpsJobWrite_`）。VPS: `apply_response` が管理者の欄へ出す。落ちる試験: GAS `idempotency: …`（2 件）、VPS `test_idempotency_not_stored_is_visible_to_admin`。本番のプロパティの大きさを測る件は残る |
| 中-4 台帳の作り直しが重なる | VERIFIED | VPS: 時間切れ・通信の失敗・retryable でも送り直さない。反映待ちが 0 でなければ動かさない（`gas_jobs.py` の `_require_no_pending_rows`）。GAS: 始めた jobId は同じ jobId で動かし直さない（`noRerun`）。落ちる試験: `test_ledger_rebuild_is_never_retried[*]`・`test_ledger_rebuild_waits_for_pending_rows`・`test_ops_rebuild_ledger_refuses_while_rows_are_pending`・GAS `ledgerRebuild: …`（2 件） |
| 中-5 起動が毎日の予定の枠にかかる | VERIFIED | `SKIP_DAILY_ON_STARTUP=1`（`worker.py`・`docker-compose.yml`）。手順書 B-16 で使う。落ちる試験: `test_skip_daily_on_startup_skips_current_daily_slot_only` |
| 中-6 台帳の upsert の失敗を握りつぶす | VERIFIED | GAS: `upsertAccumulationRow_` の後に年度タブの 1 列目（管理番号＝`buildAccRow_` の先頭）を読み直す。無ければ `ledger_upsert_failed`・retryable。PDF の経路は `ledger_failed`・retryable で、送り直すと PDF は already になる。落ちる試験: GAS `ledger upsert silently failing -> …`・`ledger upsert success is verified by reading the row back` |
| 中-7 数字に見える文字が文字のセルで書かれる | NOT FIXED | 報告どおり、リハーサルで型を比べてから決める。`sheet_codec` の text の列は変わっていない |

## 見つけた退行（GAS と同じに動く、という決まりに照らして）

### R-1（中）完了の時刻が開始より前か 24 時間を超えると、PDF も台帳の行も作られない

- 場所: `wt-team-g/src/Internal.gs` の `vpsValidWorkLog_`。`n > 1440` と「終了 >= 開始」に合わないと `null` を返し、`vpsRenderPdf_` が `bad_request`・`retryable:false` を返します。afterComplete では、この確かめが PDF と台帳の前にあるので、台帳の行も書かれません。VPS の仕事は送り直されずに failed になります。
- VPS はこの形の依頼を、ふつうの操作で送ります。画面の残業の完了は、開始を対象日の 17:20 に固定し、終了は手入力か「今」です（`vps-app/api/app/domain/worklog.py:69-77`・`145-156`）。休日の完了も同じで、範囲の確かめはありません。拒むのは総務の手動入力の `end <= start` だけです（`worklog.py:248`）。
- 確かめた例（今回実行）:
  - VPS の `Completed.wl_override`: 17:10 に完了すると `actualEndAt 17:10:00`・`actualMinutes 0`。翌日 18:00 に完了すると `actualMinutes 1480`。
  - GAS の `vpsValidWorkLog_` に同じ値を渡すと、どちらも `null`（拒む）。ふつうの値（17:20〜19:00）は通ります。
  - GAS の試験 `workLog override: bad format / out of range -> bad_request …` の 6 つ目の例（終了 17:00・開始 17:20）は、この拒否を正しい動きとして確かめています。
- 今の GAS（`src/worklog.gs:142-209` の `api_markOvertimeDone`）は、同じ入力でも分を 0 に丸めて（または 1480 分のまま）PDF を作り、台帳にも書きます。週末をまたいで休日の完了を押す、押し忘れて翌日の夕方に押す、といった使い方で起きます。
- 直し方の案: `vpsValidWorkLog_` は形（日時の書き方・数であること・0 以上）だけを確かめ、「終了 >= 開始」と 1440 分の上限は外します。上の 2 例を、GAS の試験に「通る」側として足してください。

### R-2（低〜中）Sebastian の token が 32 字未満だと、VPS が起動しない

- `config.py` の `sebastian_api_token: Field(min_length=32)` は、staging と production の両方に効きます。今の GAS の `SEBASTIAN_API_TOKEN` が 32 字未満なら、`.env` にその値を写した時点で api・worker が起動しません。入れ替えには Hermès（Sebastian）側の変更も要ります。
- `.env.example` には書かれていますが、手順書 A-9（`docs/cutover-runbook.md:134`）は「同じ値か」を確かめるだけで、長さの確かめがありません。A-9 に「32 字以上か。短ければ Hermès 側と一緒に入れ替える」を足すことをすすめます。

### R-3（低）関所は、後の ping が不合格でも、前に合格した ping を使う

- `write_gate.last_ok_ping` は `ok=True` の最新だけを見ます（試験 `test_failed_latest_ping_falls_back_to_last_ok_one` のとおり、意図した動きです）。戻した直後に ping が不合格になると、最大 10 分は前の `migrationMode: true` のまま書けます。手順書 C-2 で worker を消すので、実害は小さいです。

### R-4（低）`PDF_ROOT_FOLDER_ID` の外にある古い PDF はゴミ箱へ移さない

- 過去に PDF の根のフォルダを変えていた場合、その頃の PDF は作り直しや削除で捨てられず、Drive に残ります（今の GAS は捨てます）。VPS は `old_pdf_file_id_pending` を残したままになります（`gas_jobs.py` の `_apply_trash` は `outside_pdf_root` の `skipped` を「済み」にしません）。データが消えない側の違いなので低です。

### R-5（低）GAS と VPS で「式にしない」決まりが少し違う

- GAS の `sheetLiteral_` は `-5`・`+3` のような数の文字にも `'` を付け、文字のセルにします。VPS の `formula_safe` は数の文字を除きます。対象は部署・氏名・理由・業務内容・工番などの文字の欄なので、影響はほぼありません。

## 確かめていないこと

- `pip-audit` の再実行（道具が環境に無いため）。依存の依存も未確認です。
- 動いている環境（ステージング・本番の GAS と Sheets）での確かめ。上の判定は、すべてコードと試験によるものです。
- 既存の台帳・DB に `= + - @` で始まる値が無いか、本番の GAS のスクリプトプロパティの合計の大きさ、今の `SEBASTIAN_API_TOKEN` の長さ。どれも各報告の「確かめること」に残っています。

## 判定

REQUEST_CHANGES（変更を求めます）。指摘の修正のうち 14 件は、試験つきで直ったと確かめました。ただ R-1 は「GAS と同じに動く」決まりに反し、PDF と台帳の行が作られない形でデータに響きます。R-1 を直し、手順書 A-9 に token の長さの確かめ（R-2）を足せば、切り替えを止める理由は残りません。