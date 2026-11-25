# データ安全のレビュー（GAS → VPS 移行、本番切り替えの前）

- 日付: 2026-11-26
- 対象: `vps-app/`（api・worker・tools・docker-compose・backup）、GAS の `team-g-gas` ブランチ `bc7024a`（`src/Internal.gs`・`src/Cutover.gs`）、`docs/vps-migration-plan.md`（決定 A〜D、5〜6 章）
- 見方: データが消える・重なる・壊れる・違う場所へ書かれる道筋を洗い出しました。読むだけのレビューで、コードは変えていません。
- テスト: `cd vps-app/api && .venv/Scripts/python -m pytest -q` → **904 passed, 1 skipped**（2026-11-26、手元で実行）。

## 結論

高が 3 件あります。切り替えの前に直すことをおすすめします。

1. ステージングの GAS が、本番の PDF をゴミ箱へ移せます（高-1）。
2. GAS が作った PDF を VPS が記録しそこねると、毎晩の照合がシートの PDF の 3 列を空で上書きします（高-2）。
3. GAS の移転モードと VPS の書き込みが連動していません。戻した後に worker が動くと、GAS の承認を照合が元に戻します（高-3）。

中は 7 件、低は 9 件、確かめが要る点が 2 件です。安全装置（書き込み先の許可リスト、DeptApprovers を書かない、行・列を消さない、空の DB で書かない）は、コードと試験の両方で守られていました。

| 重さ | 件数 |
|---|---|
| 高 | 3 |
| 中 | 7 |
| 低 | 9 |
| 確かめが要る点 | 2 |

---

## 高

### 高-1 ステージングの GAS が本番の PDF をゴミ箱へ移す

- 場所: `wt-team-g/src/Internal.gs:576-617`（`vpsTrashOldPdf_`・`vpsTryTrash_`）、`Internal.gs:651-662`（`vpsAfterDelete_`）、`vps-app/tools/migrate_from_sheets.py:126-193`（本番の `pdfFileId` をそのまま入れる）、`vps-app/api/app/services/gas_jobs.py:388-392`・`417-420`（`oldPdfFileId` を DB から作る）
- 確かさ: 高
- 起きること: ステージングの DB は本番から移した行を持ち、ステージングのシートも本番のコピーです。どちらにも本番の PDF の ID が入っています。ステージングで二次承認・編集・削除・管理者の作り直しを試すと、VPS は `oldPdfFileId` に本番の ID を入れてステージングの GAS を呼びます。ステージングの GAS は imaizumi@ として動くので、`DriveApp.getFileById(本番の ID).setTrashed(true)` が通ります。本番のシートは PDF を指したまま、ファイルだけがゴミ箱へ移ります。安全装置（`safety.py`・`verify_gas_ping`）はシートの書き込み先しか見ていないため、この道は止まりません。
- 直し方:
  - `vpsTryTrash_` で、ファイルの親フォルダ（日付のフォルダ）の親が Settings の `PDF_ROOT_FOLDER_ID` であることを確かめ、違えば捨てずに `trash.status='skipped', reason='outside_pdf_root'` を返す。
  - ステージングへの移行では、`pdf_file_id`・`pdf_folder_id` を空にするか印を付ける（例: `STAGING-` を前に付ける）。ステージングのシートのコピーにある本番の ID も同じ扱いにする。
  - 今までのステージングの試験で、本番の PDF フォルダからゴミ箱へ移ったファイルが無いかを確かめる。

### 高-2 PDF の記録漏れの後、毎晩の照合が GAS の書いた PDF の 3 列を空にする

- 場所: `vps-app/api/app/services/gas_jobs.py:483-488`（`apply_response` は `pdf.status` が `created`／`already` のときだけ記録する）、`Internal.gs:514-516`（`skipIfHasPdf` で `has_pdf` の skipped）、`Internal.gs:553-563`（skipped でも最上段に `pdfFileId` を返している）、`morning_batch.py:174-201`・`208-213`（通信の失敗 3 回で新しい jobId の送り直しを積む）、`reconcile.py:93-94`・`148-195`（差は DB の値で書き直す）
- 確かさ: 高（コードの流れとして確か。起きる頻度は中）
- 起きること:
  1. 朝バッチの 1 件で、GAS は PDF を作ってシートに 3 列を書いたが、返事が VPS に届かない（GAS の Web アプリの 5xx・時間切れ）。3 回続くと fail に数え、`enqueue_batch_pdf` が**新しい jobId** で送り直す。
  2. 送り直しは GAS の冪等の記録に当たらず、シートに `pdfFileId` があるので `skipped/has_pdf` を返す。VPS はこれを記録しないので、DB の 3 列は空のまま `done` になる。
  3. その夜 2:10 の照合は、シートの `pdfFileId`＝X と DB の空を「差」とみなし、DB の値（空）で書き直す。PDF は Drive に残るが、シートと DB のどこからも指されなくなる。対象日は前日だけなので、翌朝の朝バッチも作り直さない。
  - 仕事が 24 時間失敗し続けた場合（`jobs.py:403-417`）や、GAS の冪等の記録が失われた場合（中-3）も同じ結末になります。
- 直し方:
  - `apply_response` で、`pdf.status` に関係なく返事の最上段の `pdfFileId` が空でなければ `record_pdf_result` を呼ぶ（`has_pdf` は「シートに既にある」という確かな情報です）。`morning_batch._process_item` の skipped も同じにする。
  - 照合で、PDF の 3 列（GAS が書く列）が「シートに値・DB が空」のときは書き直さず、報告に回す（または DB へ取り込む）。列ごとに「どちらが正か」を `sheet_codec` の決まりに持たせると分かりやすくなります。

### 高-3 GAS の移転モードと VPS の書き込みが連動していない（戻した後に照合が GAS の書き込みを戻す）

- 場所: `vps-app/api/app/services/reconcile.py:198-261`、`row_sync.py:568-633`、`master_sync.py:216-273`（どれも GAS の状態を見ない）、`gas_jobs.py:724-728`（`migrationMode` は記録するだけ）、`worker.py:309-330`、`docker-compose.yml`（worker は `restart: unless-stopped`）、計画書 6-6 C-2
- 確かさ: 中（コードの流れは確か。人の手順が守られれば起きない）
- 起きること: 戻した後（C-3 の後）に、`docker compose up -d`（cloudflared の入れ直しなど）で worker が動き出すと、2:10 の照合が本番のシートを DB と比べます。GAS が戻ってから承認した行は、行の数が増えないので罠 49 の歯止め（`reconcile.py:228-246`）を通り抜け、`status`・`approvedAt`・`approvedBy` が DB の古い値に書き戻されます。切り替えの前に、誤って `APP_ENV=production` で worker を起動した場合も同じで、GAS と VPS が同じ行を書き合います。今の守りは計画書 C-2「worker を止める」という人の手順 1 つだけです。
- 直し方:
  - production では、row_sync・reconcile・master_sync が書く前に「最後の `gas_ping` が 10 分以内で、`migrationMode` が true」を必須にする（`gas_jobs.gas_ping_ok` と同じ形）。false なら書かずにレーンを止め、管理者の欄に出す。ステージングでは写しの GAS の移転モードを見る。
  - 戻しの手順に `docker compose rm -sf worker`（コンテナを消す）を入れ、`up -d` で勝手に戻らないようにする。

---

## 中

### 中-1 照合が「GAS の仕事の途中の行」を飛ばさない

- 場所: `reconcile.py:118-134`（skip は反映待ちと blocked だけ）、`reconcile.py:164-172`
- 確かさ: 中
- 起きること: 2:10 に GAS の仕事（PDF 作成）が動いている間は、GAS がシートに書いた 3 列がまだ DB に入っていません。行は「反映待ち」ではないので比べられ、照合が 3 列を空にします。直後に `record_pdf_result` が DB と `synced_values` に X を入れるため、行の反映は書き直さず、翌晩まで空のままになります。
- 直し方: jobs に pending／running の GAS の仕事がある申請は照合で飛ばす。高-2 の「PDF の 3 列は空で上書きしない」を入れれば、この件も防げます。

### 中-2 blocked の行は戻しで失われ、GAS の仕事は 24 時間で failed になる

- 場所: `row_sync.py:410-414`・`538-541`、`row_sync.py:656-663`（`unblock` は状態を戻すだけ）、`gas_jobs.py:222-231`、計画書 6-6 C-2
- 確かさ: 高
- 起きること: 人がシートの行を消すと、その行の次の書き込みは「ID が見つからない」で blocked になり、それきりシートに出ません。その申請の PDF・台帳の仕事は 24 時間待って failed になります。戻すとき（C-2）は「反映待ち 0」を確かめますが、blocked の行は流しきれないので、そのまま戻すと VPS で入った内容が失われます。
- 直し方: 管理者の操作に「blocked の行を末尾に新しい行として書き直す」を足す（`sheet_row` を NULL に戻して新しい行の扱いにする）。戻しの手順 C-2 に「blocked 0 を確かめる。あれば書き直してから」を入れる。

### 中-3 GAS の冪等の記録がスクリプトプロパティの上限に当たる

- 場所: `Internal.gs:29-32`（300 件・1 件 8.5KB まで）、`Internal.gs:110-128`（失敗を握りつぶす）、`Internal.gs:134-155`
- 確かさ: 中
- 起きること: スクリプトプロパティは 1 つのスクリプトで合計 500KB までです。1 件 1〜2KB の返事を 300 件持つと上限に近づきます。上限に当たると `vpsJobWrite_` は黙って失敗し、同じ jobId の送り直しが再生されずにもう一度動きます（高-2 の入口）。同じ上限を使う `POLL_LAST_TS`・`MIGRATION_CLOSED_FORMS`・`MIGRATED_TO_URL` の書き込みも失敗するおそれがあり、戻しで使う閉じたフォームの記録が残らないこともあり得ます。
- 直し方: 記録の上限を 100 件程度に下げ、返事は要点（ok・error・PDF の 3 列・ledger.status）だけを持つ。`vpsJobWrite_` が失敗したら返事に `idempotencyStored:false` を入れ、VPS 側で警告に出す。切り替えの前に、本番の GAS の今のプロパティの合計の大きさを測る。

### 中-4 台帳の作り直し（ledgerRebuild）が時間切れで重なる・行の反映を待たない

- 場所: `gas_jobs.py:384`・`609-614`（requires なし、ほかの仕事と同じ再試行）、`Internal.gs:728-743`、`mail.gs:945-`（`rebuildAccumulationSS` は年度タブを `_backup_<日時>` へ改名してから全件を書く）、`gas_jobs.py:17`（「G に未実装」の記述が古い）
- 確かさ: 中
- 起きること: 作り直しが GAS の 6 分の上限で止まると、VPS は時間切れとして 1 分後から送り直します。7 分を過ぎると GAS は同じ jobId をもう一度動かし、途中まで書いた年度タブをまた `_backup` へ改名して書き直します。これが 24 時間くり返されると、`_backup` のタブが増え続け、本物のタブは途中までのまま残ります。行の反映待ちがあるときに動くと、まだシートに出ていない申請が台帳から抜けます。
- 直し方: ledgerRebuild は再試行しない（時間切れは failed にして管理者に知らせる）。積む前に `row_sync.status()` の待ち 0 を必須にする。docstring の「未実装」を直す。

### 中-5 worker の起動が毎日の予定の枠にかかると、確かめの前に動く

- 場所: `worker.py:95-109`・`163-181`（枠の中で起動すれば最初の tick で動く）、`reconcile.py:271`（2:10）、`master_sync.py:292`（0:30）、`morning_batch.py:329-330`（6:50・7:05）、計画書 6-4 B-16
- 確かさ: 高
- 起きること: 切り替えが長引いて 2:10〜3:10 に B-16 で worker を起動すると、`--check` より前に照合が書き込みを始めます。6:50〜7:50 なら朝バッチが動き、GAS がその日の朝に流した朝バッチと重なって BatchLogs に 2 行目ができ、台帳の前日分をもう一度書き直します（罠 24 と同じ形）。
- 直し方: 起動から最初の 1 回の毎日の予定を飛ばす設定（例 `SKIP_DAILY_ON_STARTUP=1`）を足し、切り替えの夜はそれを付けて起動する。手順書に「0:30・2:10・6:50 の前後 1 時間は B-16 を行わない」を書く。

### 中-6 台帳の upsert の失敗を GAS が握りつぶし、VPS は成功として終える

- 場所: `Internal.gs:445-455`（`upsertAccumulationRow_` を呼んだら `upserted` を返す）、`mail.gs:652-715`（中の例外を `Logger.log` で握る）
- 確かさ: 高（今の GAS と同じ動き。ただし移行後は GAS のログを見る人がいない）
- 起きること: 台帳への書き込みが例外で失敗しても、VPS の仕事は done になり、管理者の欄にも出ません。台帳の行が欠けたまま気づけません。
- 直し方: 内部用の入口では、`upsertAccumulationRow_` の後に台帳のその行を読み直して requestId があるかを確かめ、無ければ `ledger.status='error'`・`retryable:true` を返す。

### 中-7 数字に見える文字が、VPS では「文字のセル」で書かれる

- 場所: `sheet_codec.py:434-435`（text の列はいつも `stringValue`）、`sheet_codec.py:490-493`（読むときは数のセルも文字として同じとみなす）
- 確かさ: 低（影響の大きさは未確認）
- 起きること: GAS の `appendRow`・`setValues` は `"12345"` を数のセルに変えていました。VPS は `workNo1`・`orderNo1`・`jobId1`・`workerCode` などを文字のセルで書きます。照合は両方を同じとみなすので差に出ません。戻した後の GAS、統合メール通知システム、日報アプリ、テンプレの数式（XLOOKUP）が型を見ていると、VPS が書いた行だけ違う結果になるおそれがあります。
- 直し方: リハーサルで、VPS が書いた行と GAS が書いた行を `getValues()` で読み、上の列の型を比べる。型が違って困る列があれば、その列だけ「数に読める文字は `numberValue`」で書く決まりを `sheet_codec` に足す。

---

## 低

### 低-1 照合の書き直しは、読んだ時点の行番号へ確かめずに書く

- 場所: `reconcile.py:156-184`（`id_rows={}`、`target_row=d["sheet_row"]`）
- 確かさ: 中
- 起きること: 読んでから書くまでの間に人が並べ替えると、別の申請の行に書きます。間は数秒なので起きにくいです。
- 直し方: `_fix_tab` で書く直前に ID の列を読み直し、行の反映と同じ `place_rows` を通す。

### 低-2 requestId が空の WorkLogs の行は位置を確かめられない

- 場所: `row_sync.py:150-151`・`404-409`
- 確かさ: 中
- 起きること: 空の ID どうしは「同じ」とみなすので、行がずれた後に空の ID の別の行へ書くことがあります。該当の行は少ないので影響は小さいです。
- 直し方: 空の ID の行は書かずに blocked にするか、ID 以外の列も一緒に照合する。

### 低-3 マスタ同期で、書いた後の余りの行の消去が失敗すると古い行が残る

- 場所: `master_sync.py:193-213`・`257-263`
- 確かさ: 中
- 起きること: 値は書けたが `batchClear` が失敗すると、シートの末尾に前の行が残ります。DB の表は置き換わらないので、シートと DB が食い違い、日報アプリや戻した GAS は古い行も読みます。
- 直し方: 消去も 3 回まで送り直す。失敗したらそのタブを `ok:false` にして管理者に知らせる。

### 低-4 GAS に `vps.morningBatch` の入口が残っている

- 場所: `Internal.gs:270`・`695-722`、`pdfExport.gs:604-618`（`logBatchResult_` が BatchLogs に `appendRow`）
- 確かさ: 高
- 起きること: VPS は使っていませんが、呼ばれると GAS が BatchLogs に直接行を足します。DB より行が多くなり、毎晩の照合が「DB の行が少ない」で止まり続けます。
- 直し方: 入口の表から外すか、`bad_request` で拒む。

### 低-5 `record_pdf_result` の食い違いの扱いで、作られた PDF がどこにも残らないことがある

- 場所: `gas_jobs.py:291-299`
- 確かさ: 低
- 起きること: GAS を呼んでいる間に編集で 3 列が変わり、しかも `old_pdf_file_id_pending` が既に埋まっていると、作られた PDF の ID を記録しません。Drive に同じ申請の PDF が 2 つ残ります。
- 直し方: pending が埋まっているときは、「捨てる候補」として別に残す（配列の列か `sync_runs`）。

### 低-6 本番の内部用デプロイの ID が、ステージングの禁止リストに無い

- 場所: `config.py:32-36`（`PROD_GAS_DEPLOYMENT_IDS_FIXED` は画面用 @131 だけ）
- 確かさ: 高
- 起きること: ステージングの `.env` の `GAS_INTERNAL_URL` が本番の内部用デプロイを指しても起動で止まりません。`gas_ping` の照合で GAS のレーンは止まりますが、ping そのものは本番の GAS で動きます（読むだけ）。
- 直し方: 本番の内部用デプロイを作ったら、その ID を定数に足す。切り替えの前提条件に入れる。

### 低-7 委任のトークンが Drive の許可を持っている

- 場所: `google_delegate.py:41-46`
- 確かさ: 高
- 起きること: api・worker は Drive を使いませんが、トークンが漏れたときの被害が Drive 全体に広がります。
- 直し方: 使っていなければ `drive` のスコープを外す（罠 7 のとおり、許可し直しの順序に気を付ける）。

### 低-8 朝バッチの枠を逃すと取り戻さない

- 場所: `worker.py:95-109`（枠は 60 分）、`morning_batch.py:113-122`
- 確かさ: 高
- 起きること: 6:50〜7:50 に worker が止まっていると、その日の朝バッチは動きません。7:05 の確かめが「仕事がありません」と出すので気づけます。
- 直し方: 手順書に `ops_batchPdfForDate` での取り戻し方を書く。

### 低-9 バックアップが同じ VPS の中にしかない

- 場所: `docker-compose.yml`（backup）、`backup/backup_entrypoint.sh`
- 確かさ: 高
- 起きること: 決定 A でシートが写しになるので申請の中身は残りますが、DB だけにあるもの（`audit_log`・jobs の記録・`old_pdf_file_id_pending`）は VPS のディスクと一緒に失われます。
- 直し方: 週 1 回でよいので、ダンプを VPS の外へ写す。

---

## 確かめが要る点（判定を止めない）

- **Sebastian の送り直し**: 移転モードから転送を入れるまで（B-5〜B-17）と、戻しの C-1〜C-3 の間、Sebastian の完了は `service is moving` か転送の失敗を受けます（`Cutover.gs:175-209`）。Sebastian が自分で送り直さないなら、完了の記録が抜けます。送り直しの動きを確かめ、送り直さないなら切り替えの時間を Sebastian が使われない時間に合わせてください。
- **ほかのシステムが A 区分のタブに書くか**: 計画書 4-0 では、統合メール通知システム・日報アプリは「読む見込み」です。もし Requests・WorkLogs に書くものがあれば、毎晩の照合がその書き込みを DB の値で戻します。DB の編集者の一覧と、各システムの書き込み先を切り替えの前に確かめてください。

---

## 頼まれた観点ごとの結果

| 観点 | 結果 |
|---|---|
| staging が本番の ID に書く／その逆 | Sheets への書き込みは `config.py:182-223` と `safety.py:76-95` で両方向とも止まる（試験あり）。Drive は高-1 の穴がある |
| 行の反映が違う行に書く（行のずれ・並べ替え・フィルタ・人の挿入） | 書く前に ID の列を読み、違えば探し直す（`row_sync.py:385-426`）。フィルタは API に効かない。残るのは低-1・低-2 と、行を消されたときの中-2 |
| 人の直し・GAS の PDF の列の上書き | 行の反映は差分だけ書くので守られている（`row_sync.py:266-285`、`gas_jobs.py:271-317`）。照合は高-2・中-1 で上書きする |
| 戻した後に日付・日時が数になる | 日付の列はいつも表示形式を一緒に書き、照合も表示形式を見る（`sheet_codec.py:447-474`、`reconcile.py:96-102`）。問題は見つからなかった。文字の列の型は中-7 |
| 照合が逆向きに直す | 高-2・高-3・中-1 |
| 空の DB がシートを消す（罠 49） | 照合（`reconcile.py:207-212`・`228-246`）、移行（`migrate_from_sheets.py:168-171`）、マスタ同期（`master_sync.py:251-256`）に歯止めがある。行の反映は待ちの行しか書かないので当たらない |
| 移行の道具が違うシートを読む・列を取り違える | 読む元は本番の ID に固定（`migrate_from_sheets.py:71-73`）。列は見出しの文字で照合し、欠け・重なりで止まる。問題なし |
| 切り替えの TRUNCATE | `jobs, idempotency_keys, sync_runs, sessions, sheet_sync_state` を含む（罠 24b）。連番も進める（罠 9）。問題なし |
| 冪等・PDF の二重 | jobId は積んだときに決め、送り直しは同じ jobId（`gas_jobs.py:433-438`）。作り直しも「新しい PDF は捨てない」（`Internal.gs:580-583`）。残るのは中-3・低-5 |
| 仕事の再試行で台帳が重なる | 台帳は requestId の upsert と日ごとの書き直しなので重ならない。中-4・中-6 は別の形の傷み |
| worker の起動が確認の前に本番へ書く（罠 24） | 中-5 と高-3 |
| 時間帯 | 予定は JST、シリアル値は `DB_SHEET_TIMEZONE`、シートの時間帯が違えば行の反映を止める（`row_sync.py:154-160`）。問題なし |
| 移転モードで古い入口が書く | 書く道は `getDb_()`（`config.gs:69-74`）を通り、移転モードで止まる。トリガーの関数も何もしない。抜けは低-4（内部用の入口の `vps.morningBatch`）だけ |
| 戻しで VPS のデータを失う | 中-2（blocked の行）、中-7（型）、低-9（DB だけの記録） |

## よくできている点

- 書き込みの口が `SheetsApi` の 4 つに絞られ、毎回スプレッドシート・タブ・目的・要求の種類・sheetId を照合しています。DeptApprovers・Settings など B 区分は書く目的に入っておらず、試験（`tests/test_config_safety.py`）でも拒むことを確かめています。
- 行の反映が「DB の今の値と、最後に書いた値の差分だけ」を書くので、GAS が書いた PDF の列を古い値で上書きしません。
- 書く前に ID を確かめ、見つからなければ書かずに止める作りで、人の並べ替えや挿入に強くなっています。
- 移行の道具が、本番の ID の固定、見出しの照合、0 行の歯止め、1 トランザクション、移した直後の 1 セルずつの照合まで持っています。
- GAS の内部用の入口が、token の定数時間比較、jobId の再生、`expect` の照合、台帳を新しく作らない事前の確かめを持っています。

## 判定

**REQUEST CHANGES**（変更を求めます）。高-1・高-2 は確かさが高く、本番のデータ（PDF のファイルとシートの 3 列）を壊します。高-3 は人の手順で防げますが、罠 24 と同じ形なので、コードでの連動を入れてから切り替えることをおすすめします。

## 確かめること

- 高-1: 今までのステージングの試験で、本番の PDF フォルダからゴミ箱へ移ったファイルが無いか。
- 中-3: 本番の GAS の今のスクリプトプロパティの合計の大きさ。
- 中-7・確かめが要る点: 統合メール通知システム・日報アプリが読む列と、書くかどうか。
## 修正の結果 (team R)

2026-11-26、ブランチ `team-r`。VPS 側（`vps-app/`）だけを直しました。GAS 側（`src/`）は GAS の担当（`team-g-gas`）が直します。試験は `1062 passed, 1 skipped`。

| 指摘 | 結果 | コミット・試験 |
|---|---|---|
| 高-1 ステージングの GAS が本番の PDF を捨てる | 直した（VPS 側。APP_ENV=staging の移行は `pdfFileId`・`pdfFolderId` を空にし、元の値を `source_pdf_file_id`・`source_pdf_folder_id` に残す＝仕事は使わない。行は反映待ちにして、行の反映がステージングのシートの 2 列も空にする。alembic 0002。`vpsTryTrash_` のフォルダの確かめは GAS の担当） | 09500c3・`test_staging_migration_blanks_production_pdf_ids` |
| 高-2 照合が PDF の 3 列を空にする | 直した（GAS が `skipped`・`has_pdf` で返した pdfFileId を記録する＝仕事と朝バッチ。照合はシートの PDF の列に値があり DB が空なら書かずに `pdfKeptOnSheet` に出す） | 36bf9c7・`tests/test_pdf_safety.py` |
| 高-3 移転モードと書き込みの連動 | 直した（`services/write_gate`。本番は合格した `gas_ping` が 10 分以内で `migrated=true` のときだけ、行の反映・照合・マスタ同期が書く。ステージングはステージングの GAS への合格した ping。止めている理由は `ops_getStatus.sheetWriteGate` に出る。戻しの手順は `docker compose rm -sf worker`） | f6899e1・`tests/test_write_gate.py` |
| 中-1 GAS の仕事の途中の行 | 直した（待ち・実行中の GAS の仕事がある申請は照合で比べない） | 36bf9c7・`test_reconcile_skips_requests_with_running_gas_jobs` |
| 中-2 blocked の行 | 直した（`row_sync.reappend_blocked` と `python -m app.services.row_sync reappend`。手順書 C-2 に blocked 0 を追加） | fd3a986・`test_reappend_blocked_row_*`・`test_cli_status_and_reappend` |
| 中-3 冪等の記録の上限 | 一部（VPS 側: `idempotencyStored:false` を管理者の欄に出す。記録の大きさは GAS の担当） | fedf3fa・`test_idempotency_not_stored_is_visible_to_admin` |
| 中-4 台帳の作り直し | 直した（送り直さない。反映待ちが 0 でなければ積まない・動かさない。docstring も直した） | 97b795a・`test_ledger_rebuild_*`・`test_ops_rebuild_ledger_refuses_*` |
| 中-5 起動が毎日の枠にかかる | 直した（`SKIP_DAILY_ON_STARTUP=1`。手順書 B-16 で使う） | 97b795a・`test_skip_daily_on_startup_*` |
| 中-6 台帳の upsert の失敗を握りつぶす | 直さない（GAS の担当） | — |
| 中-7 数字に見える文字の型 | 直さない（リハーサルで型を比べてから決める。コードの不具合かが未確認） | — |
| 低-1 照合が行番号を確かめずに書く | 直した（書く直前に ID の列を読み直し、位置が変わった行には書かない） | 90a17ca・`test_rows_moved_between_read_and_write_are_not_written` |
| 低-2 空の ID の WorkLogs の行 | 直した（位置を確かめられないので blocked） | 90a17ca・`test_existing_row_with_blank_id_is_blocked_not_written` |
| 低-3 マスタ同期の消去の失敗 | 直した（3 回まで送り直し、だめならそのタブを失敗にして DB の表も置き換えない） | fedf3fa・`test_clear_is_retried_and_failure_marks_tab_failed` |
| 低-4 GAS の `vps.morningBatch` | 直さない（GAS の担当） | — |
| 低-5 どこにも残らない PDF | 直した（`orphan_pdf` として管理者の欄 `orphanPdfs` に出す） | fedf3fa・`test_stray_pdf_is_reported_when_pending_slot_is_taken` |
| 低-6 本番の内部用デプロイの ID | 直した（`.env` の `PROD_GAS_INTERNAL_DEPLOYMENT_ID`。形とステージングの ID でないことを確かめ、production はこのデプロイだけを呼び、staging はこれを指すと起動しない。手順書 A-3） | 489cc22・`tests/test_config_safety.py::test_prod_internal_*`・`test_gas_client_refuses_other_prod_deployment` |
| 低-7 委任の drive の許可 | 直さない（控えの道具が Drive の複製に使う。セキュリティ 中-3 と同じ） | — |
| 低-8 朝バッチの取り戻し | 直した（RUNBOOK に `ops_batchPdfForDate` の手順） | dbe05fe |
| 低-9 バックアップが同じ VPS | 直さない（運用の決めごと） | — |

あわせて、切り替えの控えと戻しの道具を足しました。`vps-app/tools/cutover_backup.py` は、この環境の DB と台帳だけを Drive で複製します。`vps-app/tools/restore_tabs_from_copy.py` は既定で確かめだけを行い、DeptApprovers・Settings は名前を書いたときだけ戻します。手順書 `docs/cutover-runbook.md` の B-8・B-15・B-16・C-2 をこれらに置き換えました（02d9a95・`tests/test_cutover_tools.py`）。
