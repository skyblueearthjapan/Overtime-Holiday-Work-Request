# 本番切り替え手順書（残業・休日出勤申請 → overtime.lineworks-local.net）

作成: 2026-11-26（担当 C）。もと: 計画書 `docs/vps-migration-plan.md` 6 章（6-1〜6-6）、共通スキル `gas-to-vps-migration` の `references/06-cutover-and-rollback.md`、工番部品の `docs/cutover-runbook.md`。
**この順番で行います。各手順の「確かめ」が通らなければ、先へ進みません。** 戻すと決めるのはオーナーです（C 章）。

〇〇 はまだ決まっていない値です。当日までに埋めます（最後の「確かめること」に一覧があります）。

## 値

| 項目 | 値 |
|---|---|
| VPS | `srv1508169`（`ssh lineworks-vps-user`）。`srv1300618` には触れない |
| VPS の置き場 | `/opt/overtime`（`vps-app/` の中身を直置き。git のフォルダではない。`.env` も直下、権限 600） |
| コンテナ | `overtime-db`・`overtime-api`・`overtime-worker`・`overtime-web`・`overtime-cloudflared`・`overtime-backup`（compose のサービス名は `db`・`api`・`worker`・`web`・`cloudflared`・`backup`） |
| DB | PostgreSQL 16、DB 名 `overtime`、ユーザー `overtime` |
| 新しい URL | `https://overtime.lineworks-local.net`（ステージングと同じホスト名。切り替えで `APP_ENV` だけ変える） |
| GAS scriptId（本番） | `1GPNBR9U9mw7TZPnl1pxtdVjFagUPxt9S6qSYlbNIEuHNvRLVspNIQNc7` |
| 本番の画面のデプロイ | `AKfycbwHnaECeni3FmmgtVzBZIM_xDMH3bw9Qx5sQMrW9J3XoRCOs_RmdyBensZX5Xta6sfIxA`。**戻す版は @131** |
| 移転モード入りの版 | @〇〇（A-3 で作る） |
| Sebastian が使っているデプロイ | 〇〇（A-3 で確かめる。画面のデプロイと同じなら同じ ID） |
| 本番の内部用のデプロイ（`GAS_INTERNAL_URL`） | `https://script.google.com/macros/s/〇〇/exec`（A-3 で作る） |
| 本番 DB スプレッドシート | `1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ`（PDF テンプレも同じ） |
| 蓄積台帳 | `1tKlT57SJxJijgfDVGafZ434sdVjscJJqN2uQ7P3sUyU` |
| 本番 PDF フォルダ | `1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B` |
| マスタ元（読むだけ） | `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ` |
| ステージング（安全装置の確かめに使う） | DB `1I0dPlH6w14UaiJVbO9i6wFCXlvGyPEJvDus8ekGptWY`、台帳 `16M5PxyozS_tTEpQmR6jizICgZov4SoK2M1IO4TG1oHI`、PDF フォルダ `1KsWanVEQr-gMsb_N9mxzplEjKFibsnrc`、GAS `1y027cB_59N1P2o3toawpCyPZ2k_W6mPETtgiDpPFtpY7TKgBWslnyLNM` |
| 業務の表（空にする） | `requests, work_logs, batch_logs, settings, dept_approvers, approver_map_rows, workers, orders, company_calendar, sheet_sync_state, audit_log` |
| 土台の表（空にする） | `jobs, idempotency_keys, sync_runs, sessions` |
| 残す表 | `google_credentials`（委任）、`alembic_version` |
| 社内ポータル | スプレッドシート `1DdhbJuz-zyS_2pKcM8y6Xf2pHofuzwNDgoW9eWcB0yg` の `APP_MASTER`。key 〇〇 の行、url 列 〇〇 |
| 手元の秘密の置き場 | `C:\Users\imaizumi.LINEWORKS-NET\.secrets\overtime\`（本番の内部用トークンは 〇〇 のファイル） |
| 切り替えの控えの置き場（VPS） | `/opt/overtime/cutover/`（`backups/` は自動で消えるので使わない） |

### コマンドの書き方

VPS のコマンドは、手元（Git Bash）から次の形で流します。毎回、先頭で `hostname` を確かめ、違えば何もしません。docker は `sg docker -c "…"` で動かします（lineworks ユーザーは sudo できません）。

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST: $(hostname)"; exit 9; }
cd /opt/overtime
# ここに手順のコマンド
EOF
```

DB へは次の関数で SQL を流します（上の形の中で定義してから使います）。

```sh
psql_q() { sg docker -c "docker compose exec -T db psql -U overtime -d overtime -v ON_ERROR_STOP=1 -X -P footer=off"; }
```

時間のかかるもの（build、移行）は `timeout 590 ssh …` で包みます（罠 2）。GAS の関数は、GAS エディタ `https://script.google.com/u/1/home/projects/1GPNBR9U9mw7TZPnl1pxtdVjFagUPxt9S6qSYlbNIEuHNvRLVspNIQNc7/edit` を開き、右上のアカウントが `imaizumi@lineworks-local.info` であることを確かめてから、関数を選んで「実行」します（運用関数は持ち主がエディタから実行したときだけ動きます。罠 48）。

---

## A. 前もって済ませること（前日まで）

- [ ] **A-1 受け入れ**: ステージングで `tools/qa/acceptance.py` を読むだけと `--write` の両方で流し、FAIL が 0（下のコマンド）。`docs/acceptance-checklist.md` の項目がすべて「済」。手元の自動テストが全件通る（`cd vps-app/api && .venv/Scripts/python -m pytest -q` の最後の行）。レビューの指摘が片付いている。

  ```sh
  ssh lineworks-vps-user 'bash -s' <<'EOF'
  set -eu
  test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
  cd /opt/overtime
  sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --mint-session --expect-env staging"
  sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --mint-session --expect-env staging --write"
  EOF
  ```

  確かめ: 両方とも最後に `PASS n / FAIL 0` と `session removed: True`。`--write` で作った申請 2 件は【受け入れテスト】の印つきで、最後に削除（canceled）されています。
- [ ] **A-2 リハーサル**: 本番からの移行のリハーサルを 2 回以上（B-9〜B-16 をステージングのまま流す）。`verifyProblems: []`、照合 SQL（B-12）の数を控え、移行直後の `sheet_reconcile --check` で差 0。戻しの練習（C 章）をステージングの GAS の写しで 1 回。
- [ ] **A-3 GAS に移転モードを入れる**（オーナーの許可を得てから。罠 21・9-1 #17）。
  1. `team-g-gas` の `src/Cutover.gs`・`src/Internal.gs` と 7-6 の行を `master` に入れて push します。`master` の `src/**` の push は GitHub Actions（`deploy-gas.yml`）が `clasp push --force` で **HEAD を書き換え、本番のトリガーはその瞬間から新しいコードで動きます**。`MIGRATED_TO_URL` が無い間は今と同じ動きなので、push した後に本番の `/exec` が今までどおり開くこと、1 分ごとの取り込みが成功していること（実行数の画面）を確かめます。
  2. 版を作って、画面のデプロイと Sebastian のデプロイに当てます。説明文は英数字にします（罠 27）。

     ```sh
     cd /d/projects/zangyo-kyujitsu-app/github-repo
     CLASP="npx -y @google/clasp@3.4.1"
     $CLASP show-authorized-user                     # imaizumi@lineworks-local.info であること（罠 28）
     $CLASP list-deployments --json > /tmp/overtime-deployments-before.json
     rm -rf /tmp/ot-pull && mkdir /tmp/ot-pull && (cd /tmp/ot-pull && echo '{"scriptId":"1GPNBR9U9mw7TZPnl1pxtdVjFagUPxt9S6qSYlbNIEuHNvRLVspNIQNc7"}' > .clasp.json && $CLASP pull)
     grep -A3 '"webapp"' /tmp/ot-pull/appsscript.json  # executeAs USER_DEPLOYING・access ANYONE_ANONYMOUS が HEAD にあること（罠 45）
     $CLASP create-version "migration mode and vps internal endpoint"   # 出てきた版の番号を 〇〇 として控える
     $CLASP update-deployment AKfycbwHnaECeni3FmmgtVzBZIM_xDMH3bw9Qx5sQMrW9J3XoRCOs_RmdyBensZX5Xta6sfIxA --versionNumber 〇〇 --description "migration mode"
     $CLASP list-deployments --json | grep -B2 -A6 AKfycbwHnaEC    # versionNumber が 〇〇 になったこと
     ```

     **clasp の `update-deployment` は「Redeployed … @N」と出ても版が変わらないことがあります**（`docs/skill-learnings-draft.md`。2026-11-26 に英数字の説明文でも起きました）。`list-deployments --json` の `versionNumber` が変わっていなければ、Apps Script API で当てます（clasp の認証ファイルのトークンを使います。トークンは画面に出しません）。

     ```sh
     DEPLOYMENT_ID=AKfycbwHnaECeni3FmmgtVzBZIM_xDMH3bw9Qx5sQMrW9J3XoRCOs_RmdyBensZX5Xta6sfIxA VERSION=〇〇 python - <<'PY'
     import json, os, urllib.parse, urllib.request
     SCRIPT_ID = "1GPNBR9U9mw7TZPnl1pxtdVjFagUPxt9S6qSYlbNIEuHNvRLVspNIQNc7"
     dep, ver = os.environ["DEPLOYMENT_ID"], int(os.environ["VERSION"])
     t = json.load(open(os.path.expanduser("~/.clasprc.json")))["tokens"]["default"]
     form = urllib.parse.urlencode({"client_id": t["client_id"], "client_secret": t["client_secret"],
                                    "refresh_token": t["refresh_token"], "grant_type": "refresh_token"}).encode()
     token = json.load(urllib.request.urlopen("https://oauth2.googleapis.com/token", form))["access_token"]
     body = {"deploymentConfig": {"scriptId": SCRIPT_ID, "versionNumber": ver, "manifestFileName": "appsscript",
                                  "description": "migration mode"}}
     req = urllib.request.Request(f"https://script.googleapis.com/v1/projects/{SCRIPT_ID}/deployments/{dep}",
                                  data=json.dumps(body).encode(), method="PUT",
                                  headers={"Authorization": "Bearer " + token, "Content-Type": "application/json"})
     print("versionNumber =", json.load(urllib.request.urlopen(req))["deploymentConfig"]["versionNumber"])
     PY
     ```

     Sebastian が別のデプロイ（〇〇）を使っていれば、同じ手順でそちらにも当てます。当てた直後に、本番の `/exec` が今までどおり開くことを画面で確かめます。
  3. 本番の内部用のデプロイを作ります: `$CLASP create-deployment --versionNumber 〇〇 --description "vps internal"`。出てきた deploymentId を `GAS_INTERNAL_URL`（`https://script.google.com/macros/s/<id>/exec`）として控え、`vps-app/api/app/config.py` の `PROD_GAS_DEPLOYMENT_IDS_FIXED` に足して（ステージングがこれを指せば起動しない歯止め）、新しいイメージを作っておきます（`docs/RUNBOOK.md` のデプロイ。`APP_IMAGE_TAG` を控える）。
  4. GAS エディタの「プロジェクトの設定 → スクリプト プロパティ」で 2 つ足します。`VPS_INTERNAL_TOKEN`＝本番の内部用トークン（32 字以上の乱数。`SEBASTIAN_API_TOKEN` とは別の値。手元の 〇〇 のファイルにも置く）、`VPS_ORIGIN`＝`https://overtime.lineworks-local.net`（当日 `enterMigrationModeForOvertime` が読む）。**`MIGRATED_TO_URL` は入れません**（入れた瞬間に移転モードになります）。
  5. GAS エディタで `showMigrationStatus` を 1 回実行し、権限の許可と今の状態を見ます（移転モード「なし」、内部用の token「あり」、トリガー 4 本、旧フォーム 21 本）。
- [ ] **A-4 内部用の入口の確かめ**（手元から。何も書きません）。

  ```sh
  URL="https://script.google.com/macros/s/〇〇/exec"
  TOKEN_FILE="$HOME/.secrets/overtime/〇〇"
  python - "$URL" "$TOKEN_FILE" <<'PY'
  import json, sys, urllib.request
  url, token = sys.argv[1], open(sys.argv[2]).read().strip()
  req = urllib.request.Request(url, data=json.dumps({"internal": True, "action": "internal.ping", "token": token}).encode(),
                               headers={"Content-Type": "application/json"})
  r = json.load(urllib.request.urlopen(req))
  print({k: r.get(k) for k in ("ok", "scriptId", "spreadsheetId", "spreadsheetTimeZone", "migrated", "forwardUrlSet", "accumulationKeyRows")})
  print(r.get("settings"))
  PY
  ```

  確かめ: `ok: True`、`scriptId` が `1GPNBR9U…`、`spreadsheetId` が `1Knx_kaQ…`、`migrated: False`、`settings` の `TEMPLATE_SSID`＝`1Knx_kaQ…`・`PDF_ROOT_FOLDER_ID`＝`1Bs1FvgR…`・`ACCUMULATION_SS_ID`＝`1tKlT57S…`、`spreadsheetTimeZone` が `.env` の `DB_SHEET_TIMEZONE`（〇〇）と同じ。
- [ ] **A-5 控えておくもの**（計画書 6-2）: DB のファイルの時間帯と地域設定、型の表（計画書 5-5）、蓄積台帳のタブと見出し、Settings の重なり（`accumulationKeyRows`）、DB の編集者の一覧、`操作` タブのボタンに割り当てたスクリプトの有無。
- [ ] **A-6 トリガーの持ち主**: GAS の「トリガー」画面で「すべての人のトリガー」を表示し、imaizumi@ 以外の持ち物が無いこと（gap-4 F-17）。あれば、その人に消してもらう段取りを決めておく。
- [ ] **A-7 部署の共有アカウント**: 13 部署の共有アカウントが新しい URL（ステージング）でログインできること。翌朝のピーク（6〜8 時）に初めてのログインが重ならないよう、前日の業務中に各部署の端末で 1 回開いてもらうかを、オーナーと決める（gap-2 S2・Q7）。
- [ ] **A-8 旧フォーム**: 21 本の直近の回答が 0 件であること（決定 #11 の条件）。`showMigrationStatus` の `POLL_LAST_TS` を控える。
- [ ] **A-9 `.env` の下書き**: 本番で変えるキーは `APP_ENV`・`GAS_INTERNAL_URL`・`GAS_INTERNAL_TOKEN`・`APP_IMAGE_TAG` だけ。`SEBASTIAN_API_TOKEN` が本番の GAS のプロパティ `SEBASTIAN_API_TOKEN` と同じ値であることを確かめておく（gap-4 F-16）。
- [ ] **A-10 オーナーへの声かけ**（下の「オーナーへの連絡」の 1・2）。当日の開始時刻を決め、「切り替えを始めます」の了承をもらう。

---

## B. 当日（業務の後の夜。朝バッチ 6:50 より前に終える）

| # | 手順 | 確かめ | 戻す範囲 |
|---|---|---|---|
| B-0 | `ssh lineworks-vps-user hostname` | `srv1508169` | — |
| B-1 | 進行中の件数を数える | submitted・承認済みで完了未記録・二次承認待ち・完了済みで PDF なし（2026-11-26 は 0・1・6・25）を控える | — |
| B-2 | GAS: 実行数の画面と `showMigrationStatus` | 動いている実行なし、未取り込みの回答 0、トリガー 4 本、移転モードなし | — |
| B-3 | GAS: **先にフォームを閉じる** `closeAllFormsForMigration` | 21 本が「回答を受け付けない」。ログに閉じた本数 | `reopenFormsAfterRollback` |
| B-4 | **流しきり 1**: 2 分待つ | 閉じた後に `pollNewResponses_` が完了。`showMigrationStatus` で `POLL_LAST_TS` 以降の回答 0 | — |
| B-5 | GAS: `enterMigrationModeForOvertime` | `/exec` が移転ページ。`/exec?page=approve2&date=<今日>` のボタンで新しいアプリの同じ画面が最上位の窓で開く | `exitMigrationMode` |
| B-6 | GAS: `removeAllTriggersForMigration` | ログの「削除したトリガー」に 4 本、「残っているトリガー: なし」。「すべての人のトリガー」も空 | `restoreTriggersAfterRollback` |
| B-7 | **流しきり 2**: 1〜2 分待ち、もう一度 `showMigrationStatus` | 実行数の画面に動いている実行なし。未取り込み 0、トリガーなし | — |
| B-8 | **控え**: DB と台帳を Drive API で複製（流しきり 2 の後） | 控えの ID を控える。控えの行数が元と同じ | — |
| B-9 | VPS の api・worker を止め、`pg_dump -Fc` | api・worker が exited、ダンプが 1KB 超 | `docker compose start api worker` |
| B-10 | 表を空にする（業務の表＋jobs・idempotency_keys・sync_runs・sessions） | 15 表が 0 件、`google_credentials` 1 件 | ダンプから `pg_restore` |
| B-11 | 移行 `--dry-run` → 本実行 | 終了コード 0、`verifyProblems: []`、`droppedColumns: {}`、件数がリハーサル以上 | B-10 から |
| B-12 | 照合 SQL、Settings の ID と時間帯 | 全部一致 | B-10 からか C へ |
| B-13 | `.env` を控えて本番の値に | 値と権限 600 | 控えを戻す |
| B-14 | **api だけ**起動し、安全装置と `internal.ping` | `NG` 0 行、healthz が production、ping が本番の DB と移転モード on | ログを見て直すか C へ |
| B-15 | マスタ同期と共有タブの取り込みを手で 1 回 | 作業員 50・工番 560・業務 NO 47・カレンダー 130 前後。DeptApprovers 13 部署、管理者 2 人 | 控えから 4 タブを写し戻す |
| B-16 | **worker を起動**し、行の反映と照合 `--check` | A 区分の差 0、反映待ち 0、`gas_ping` 成功、DeptApprovers が変わっていない | worker を止め、控えから写し戻して C へ |
| B-17 | GAS: `setSebastianForwardUrl` → Sebastian の `ping` | GAS の `/exec` 経由で `{ok:true, pong}`、VPS のログに届く | `clearSebastianForwardUrl` |
| B-18 | 画面と受け入れ（読むだけ）の確認。書き込みはしない | 期待どおり、acceptance の FAIL 0 | — |
| B-19 | 社内ポータルの url 列を新しい URL に（前に声をかける） | ポータルから新しい URL が開く | 元の `/exec` に戻す |
| B-20 | 最終確認 | 全コンテナ up、仕事の pending・failed 0、外から healthz 200、エラーログなし | — |
| B-21 | 引継ぎメモ・メモリ・スキルを更新し、オーナーへ報告 | — | — |

Sebastian の転送（B-17）は、計画書 6-4 の順のとおりポータル（B-19）より前に行います。移転モード（B-5）から転送の設定までの間、Sebastian には `{"error":"service is moving, retry later"}` が返るので、この間を短くするためです。

### B-0 ホスト名

```sh
ssh lineworks-vps-user hostname    # srv1508169 でなければ中止
```

### B-1 進行中の件数

GAS の画面（総務の滞留監視・二次承認・承認者画面）か、移行ツールの `--dry-run`（B-11 と同じコマンド。何も書きません）の報告で数えます。控える数: submitted、approved で `actualEndAt` が空、approved で `approvedBy2` が空、approved で正味 > 0 かつ PDF なし。

### B-2 GAS の今の状態

GAS エディタの左の「実行数」で、実行中の行が無いことを見ます。続けて `showMigrationStatus` を実行し、ログで次を確かめます。

- 移転モード: なし／内部用の token: あり／トリガー: `pollNewResponses_, nightlyUpdateAllForms_, morningBatchNoMail_, syncAllMasters`（4 本）
- 旧フォーム: 21 本（受付中 21 本）／`POLL_LAST_TS 以降の回答: 0 件`

### B-3 フォームを閉じる（移転モードより先）

GAS エディタで `closeAllFormsForMigration` を実行します。移転モードに入ると 1 分ごとの取り込みが止まるので、先に受付を閉じて、閉じた後の取り込みを 1 回流します（gap-4 F-8）。

確かめ: ログ「旧フォームの受付を閉じました: 21 本」、失敗 0 本。フォームと回答は消えません。閉じた記録はスクリプトのプロパティ `MIGRATION_CLOSED_FORMS` に残ります。

### B-4 流しきり 1

2 分待ちます。「実行数」で、閉じた時刻の後に `pollNewResponses_` が 1 回以上「完了」になったことを見ます。もう一度 `showMigrationStatus` を実行し、`POLL_LAST_TS 以降の回答: 0 件` を確かめます。0 でなければ、0 になるまで待ちます。

### B-5 移転モード

GAS エディタで `enterMigrationModeForOvertime` を実行します（`VPS_ORIGIN` を読み、`MIGRATED_TO_URL` を入れます）。

確かめ:
- 本番の `/exec`（`https://script.google.com/macros/s/AKfycbwHnaECeni3FmmgtVzBZIM_xDMH3bw9Qx5sQMrW9J3XoRCOs_RmdyBensZX5Xta6sfIxA/exec`）が「残業・休日出勤申請は新しいアドレスへ移りました」になる。
- `/exec?page=approve2&date=<今日 yyyy-MM-dd>` のボタンで、`https://overtime.lineworks-local.net/?page=approve2&date=<今日>` が**最上位の窓で**開く（GAS の枠の中で開かない）。

注意: ここから B-14 までは、移転ページのボタンの先がまだステージングか止まった VPS です。夜に行うのはこのためです。

戻す: `exitMigrationMode`。

### B-6 トリガーを消す

GAS エディタで `removeAllTriggersForMigration` を実行します。

確かめ: ログ「削除したトリガー: pollNewResponses_ (…), nightlyUpdateAllForms_ (…), morningBatchNoMail_ (…), syncAllMasters (…)」と「残っているトリガー: なし」。トリガーの画面で「すべての人のトリガー」も空。

戻す: `restoreTriggersAfterRollback`（朝バッチとマスタ同期の 2 本。フォームの 2 本も戻すと決めたときは `restoreAllTriggersAfterRollback`）。

### B-7 流しきり 2

1〜2 分待ちます（入口を通り過ぎた書き込みと PDF 作成が終わるのを待ちます。罠 30）。「実行数」に実行中の行が無いこと、`showMigrationStatus` でトリガー「なし」・未取り込み 0 を確かめます。**控え（B-8）はこの後に取ります**（GAS が最後に書いた後の状態を控えるため）。

### B-8 控え（DB と台帳の複製）

D 担当の `tools/cutover_backup.py` があればそれを使います。無ければ次を流します（委任トークンで Drive API の `files/{id}/copy`。複製はマイドライブの直下に置き、共有しません）。

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app api python -" <<'PY'
import json, urllib.parse, urllib.request
from datetime import datetime
from app.config import get_settings
from app.google_delegate import get_access_token
S = get_settings()
H = {"Authorization": "Bearer " + get_access_token(), "Content-Type": "application/json"}
def call(method, url, body=None):
    req = urllib.request.Request(url, data=None if body is None else json.dumps(body).encode(), method=method, headers=H)
    return json.load(urllib.request.urlopen(req))
def tab_rows(sid):
    meta = call("GET", f"https://sheets.googleapis.com/v4/spreadsheets/{sid}?fields=sheets.properties.title")
    titles = [s["properties"]["title"] for s in meta["sheets"]]
    q = "&".join("ranges=" + urllib.parse.quote(f"'{t}'!A:A") for t in titles)
    vr = call("GET", f"https://sheets.googleapis.com/v4/spreadsheets/{sid}/values:batchGet?{q}&majorDimension=COLUMNS")
    return {t: len((v.get("values") or [[]])[0]) for t, v in zip(titles, vr["valueRanges"])}
stamp = datetime.now().strftime("%Y-%m-%d %H%M")
for label, fid in (("DB", S.prod_spreadsheet_id), ("台帳", S.prod_ledger_spreadsheet_id)):
    c = call("POST", f"https://www.googleapis.com/drive/v3/files/{fid}/copy?supportsAllDrives=true&fields=id,name",
             {"name": f"【切替前の控え】残業・休日出勤申請 {label} {stamp}", "parents": ["root"]})
    a, b = tab_rows(fid), tab_rows(c["id"])
    print(label, "控えの ID:", c["id"], c["name"])
    for t in a:
        print("  ", "OK" if a[t] == b.get(t) else "NG", t, a[t], b.get(t))
PY
EOF
```

確かめ: 2 冊とも、すべてのタブが `OK`（A 列の行数が元と控えで同じ）。控えの ID を引継ぎメモに書きます。`NG` があれば、もう一度複製します。

### B-9 VPS を止めて DB を控える

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
mkdir -p cutover
sg docker -c "docker compose stop worker api"
sg docker -c "docker compose ps -a"
F=cutover/overtime_before_cutover_$(date +%Y%m%d_%H%M).dump
sg docker -c "docker compose exec -T db pg_dump -U overtime -Fc overtime" > "$F"
ls -l "$F"
test "$(stat -c %s "$F")" -gt 1024 && echo "DUMP OK: $F" || echo "DUMP TOO SMALL: $F"
EOF
```

確かめ: `overtime-api`・`overtime-worker` が `Exited`、`DUMP OK`。戻す: `sg docker -c "docker compose start api worker"`。

### B-10 表を空にする

移行ツールは自分が入れる表しか空か見ないので、仕事・冪等キー・同期の記録・セッション・マスタ・監査も手で空にします。`jobs` を残すとステージングで積んだ仕事が本番の GAS へ流れ、`sessions` を残すと受け入れテストのセッションが本番で生きます（罠 24b）。

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
psql_q() { sg docker -c "docker compose exec -T db psql -U overtime -d overtime -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
TRUNCATE requests, work_logs, batch_logs, settings, dept_approvers, approver_map_rows, workers, orders,
         company_calendar, sheet_sync_state, audit_log, jobs, idempotency_keys, sync_runs, sessions RESTART IDENTITY;
SELECT 'requests' t, count(*) FROM requests UNION ALL SELECT 'work_logs', count(*) FROM work_logs
UNION ALL SELECT 'batch_logs', count(*) FROM batch_logs UNION ALL SELECT 'settings', count(*) FROM settings
UNION ALL SELECT 'dept_approvers', count(*) FROM dept_approvers UNION ALL SELECT 'approver_map_rows', count(*) FROM approver_map_rows
UNION ALL SELECT 'workers', count(*) FROM workers UNION ALL SELECT 'orders', count(*) FROM orders
UNION ALL SELECT 'company_calendar', count(*) FROM company_calendar UNION ALL SELECT 'sheet_sync_state', count(*) FROM sheet_sync_state
UNION ALL SELECT 'audit_log', count(*) FROM audit_log UNION ALL SELECT 'jobs', count(*) FROM jobs
UNION ALL SELECT 'idempotency_keys', count(*) FROM idempotency_keys UNION ALL SELECT 'sync_runs', count(*) FROM sync_runs
UNION ALL SELECT 'sessions', count(*) FROM sessions UNION ALL SELECT 'google_credentials', count(*) FROM google_credentials;
SQL
EOF
```

確かめ: 15 表が 0、`google_credentials` が 1。戻す: `sg docker -c "docker compose exec -T db pg_restore -U overtime -d overtime --clean --if-exists" < cutover/overtime_before_cutover_<日時>.dump`。

### B-11 移行

先に `--dry-run`（読むだけ）で数字を見て、よければ本実行します。表は B-10 で空にしたので `--truncate` は付けません（空でなければツールが止まるので、それも確かめになります）。

```sh
timeout 590 ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
S=$(date +%H%M)
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api python tools/migrate_from_sheets.py --source sheets --dry-run" | tee cutover/migrate_dryrun_$S.json
echo "exit=${PIPESTATUS[0]}"
EOF
```

`verifyProblems: []`・`droppedColumns: {}`・件数（Requests・WorkLogs・BatchLogs・Settings・DeptApprovers）がリハーサル以上で、警告が既知のもの（WorkLogs の孤児・行の無い申請など）だけなら、`--dry-run` を外して同じコマンドで本実行します（出力は `cutover/migrate_run_$S.json`）。

確かめ: 終了コード 0（2 は止めた、1 は入れた後の照合が合わない）。戻す: B-10 からやり直し。

### B-12 照合（計画書 5-4）

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
psql_q() { sg docker -c "docker compose exec -T db psql -U overtime -d overtime -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
SELECT (SELECT count(*) FROM requests) requests, (SELECT count(*) FROM work_logs) work_logs, (SELECT count(*) FROM batch_logs) batch_logs;
SELECT status, request_type, count(*) FROM requests GROUP BY 1, 2 ORDER BY 1, 2;
SELECT to_char(target_date, 'YYYY-MM') ym, request_type, count(*) FROM requests GROUP BY 1, 2 ORDER BY 1, 2;
SELECT sum(approved_minutes) approved FROM requests;
SELECT sum(actual_minutes) actual, sum(break_minutes) brk, sum(net_minutes) net FROM work_logs;
WITH w AS (SELECT DISTINCT ON (request_id) * FROM work_logs ORDER BY request_id, id)
SELECT count(*) FILTER (WHERE r.status = 'submitted') submitted,
       count(*) FILTER (WHERE r.status = 'approved' AND coalesce(w.actual_end_at, '') = '') approved_not_done,
       count(*) FILTER (WHERE r.status = 'approved' AND r.approved_by2 = '') waiting_secondary,
       count(*) FILTER (WHERE r.status = 'approved' AND coalesce(w.net_minutes, 0) > 0 AND r.pdf_file_id = '') pdf_missing,
       count(*) FILTER (WHERE r.pdf_file_id <> '') with_pdf,
       count(*) FILTER (WHERE r.approved_by2 <> '') secondary_done,
       count(*) FILTER (WHERE r.approved_by LIKE '%(却下)') rejected,
       count(*) FILTER (WHERE r.worker_email <> '') with_worker_email,
       count(*) FILTER (WHERE r.target_date IS NULL) target_not_date,
       count(*) FILTER (WHERE w.request_id IS NULL) no_worklog
FROM requests r LEFT JOIN w ON w.request_id = r.request_id;
SELECT count(*) orphan_worklogs FROM work_logs w WHERE NOT EXISTS (SELECT 1 FROM requests r WHERE r.request_id = w.request_id);
SELECT key, value FROM settings WHERE key IN ('TEMPLATE_SSID', 'PDF_ROOT_FOLDER_ID', 'ACCUMULATION_SS_ID', 'SOURCE_SSID', 'APP_URL', 'PDF_MODE', 'ADMIN_EMAILS') ORDER BY key;
SELECT count(DISTINCT dept) dept_approver_depts FROM dept_approvers;
SQL
grep -E '^(PROD_SPREADSHEET_ID|PROD_PDF_FOLDER_ID|PROD_LEDGER_SPREADSHEET_ID|MASTER_SOURCE_SPREADSHEET_ID|DB_SHEET_TIMEZONE)=' .env
EOF
```

確かめ:
- 件数・状態×種別・月×種別・分の合計が、dry-run の報告とシートの数（B-1・リハーサルの控え）と一致。進行中の数が B-1 と一致。孤児（2026-11-26 は 32）・行の無い申請（同 17）が控えと一致。
- Settings の `TEMPLATE_SSID`＝`PROD_SPREADSHEET_ID`、`PDF_ROOT_FOLDER_ID`＝`PROD_PDF_FOLDER_ID`、`ACCUMULATION_SS_ID`＝`PROD_LEDGER_SPREADSHEET_ID`、`SOURCE_SSID`＝`MASTER_SOURCE_SPREADSHEET_ID`。`APP_URL` は GAS の `/exec` のまま（変えない）。
- DB のファイルの時間帯（A-4 の `spreadsheetTimeZone`）が `DB_SHEET_TIMEZONE` と同じ（gap-5 高-2）。
- DeptApprovers が 13 部署、`ADMIN_EMAILS` が 2 人。

戻す: 合わなければ B-10 からやり直すか、C へ。

### B-13 `.env` を本番の値に

手元から本番の内部用トークンを送ってから書き換えます（トークンは画面に出しません）。

```sh
ssh lineworks-vps-user 'test "$(hostname)" = srv1508169 || exit 9; umask 077; cat > /opt/overtime/cutover/gas_internal_token_prod.txt' < "$HOME/.secrets/overtime/〇〇"
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
cp -p .env cutover/.env.before-cutover
TOK=$(cat cutover/gas_internal_token_prod.txt)
test ${#TOK} -ge 32 || { echo "TOKEN TOO SHORT"; exit 8; }
sed -i "s|^APP_ENV=.*|APP_ENV=production|" .env
sed -i "s|^GAS_INTERNAL_URL=.*|GAS_INTERNAL_URL=https://script.google.com/macros/s/〇〇/exec|" .env
sed -i "s|^GAS_INTERNAL_TOKEN=.*|GAS_INTERNAL_TOKEN=$TOK|" .env
rm -f cutover/gas_internal_token_prod.txt
chmod 600 .env
grep -E '^(APP_ENV|APP_IMAGE_TAG|GAS_INTERNAL_URL|PUBLIC_ORIGIN|DB_SHEET_TIMEZONE)=' .env
stat -c '%a %n' .env
sg docker -c "docker image ls overtime-app"
EOF
```

確かめ: `APP_ENV=production`、`GAS_INTERNAL_URL` が本番の内部用のデプロイ、`APP_IMAGE_TAG` が A-3 で作った最新のイメージのタグ（古いタグだと `up -d` で古いコードに戻る。罠 24e）、`600 .env`。戻す: `cp -p cutover/.env.before-cutover .env`。

### B-14 api だけ起動して安全装置を確かめる

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
sg docker -c "docker compose up -d api"
sleep 20
sg docker -c "docker compose ps"
sg docker -c "docker compose exec -T web wget -qO- http://127.0.0.1/healthz"; echo
sg docker -c "docker compose exec -T api env PYTHONPATH=/app python -" <<'PY'
from app import safety
from app import config as C
from app.config import get_settings
S = get_settings()
print("env", S.app_env, "/ writable", safety.writable_spreadsheet_id(S), "(本番 DB であること)")
print("internal deployment", C.gas_deployment_id(S.gas_internal_url))
def refuse(label, sid, tab, purpose):
    try:
        safety.assert_writable(sid, tab, purpose, S); print("NG: allowed", label)
    except safety.UnsafeWriteError:
        print("OK: refused", label)
def allow(label, sid, tab, purpose):
    try:
        safety.assert_writable(sid, tab, purpose, S); print("OK: allowed", label)
    except Exception as e:
        print("NG: refused", label, e)
refuse("staging DB", C.STAGING_SPREADSHEET_ID_KNOWN, "Requests", "row_sync")
refuse("staging ledger", C.STAGING_LEDGER_SPREADSHEET_ID_KNOWN, "Requests", "row_sync")
refuse("staging pdf folder", C.STAGING_PDF_FOLDER_ID_KNOWN, "Requests", "row_sync")
refuse("master source", C.MASTER_SOURCE_SPREADSHEET_ID_FIXED, "作業員マスタ", "master_sync")
refuse("portal", "1DdhbJuz-zyS_2pKcM8y6Xf2pHofuzwNDgoW9eWcB0yg", "APP_MASTER", "row_sync")
refuse("prod ledger", C.PROD_LEDGER_SPREADSHEET_ID_FIXED, "Requests", "row_sync")
refuse("prod pdf folder", C.PROD_PDF_FOLDER_ID_FIXED, "Requests", "row_sync")
refuse("prod DB Settings", C.PROD_SPREADSHEET_ID_FIXED, "Settings", "row_sync")
refuse("prod DB DeptApprovers", C.PROD_SPREADSHEET_ID_FIXED, "DeptApprovers", "master_sync")
refuse("prod DB 申請書フォーム", C.PROD_SPREADSHEET_ID_FIXED, "申請書フォーム", "row_sync")
refuse("prod DB Requests by master_sync", C.PROD_SPREADSHEET_ID_FIXED, "Requests", "master_sync")
for tab in ("Requests", "WorkLogs", "BatchLogs"):
    allow("prod DB " + tab, C.PROD_SPREADSHEET_ID_FIXED, tab, "row_sync")
for tab in ("作業員マスタ", "工番マスタ", "社内カレンダーマスタ", "業務NOマスタ"):
    allow("prod DB " + tab, C.PROD_SPREADSHEET_ID_FIXED, tab, "master_sync")
PY
sg docker -c "docker compose exec -T api python -m app.worker run gas_ping"
EOF
```

確かめ: healthz が `{"ok":true,"env":"production"}`。`NG` が 0 行（1 行でもあれば先へ進まない）。`writable` が `1Knx_kaQ…`。`gas_ping: ok=True` で、summary の `problems: []`・`migrationMode: True`・`spreadsheetTimeZone` が `DB_SHEET_TIMEZONE` と同じ。戻す: api のログ（`sg docker -c "docker compose logs --tail 100 api"`）を見て直すか、C へ。

### B-15 マスタ同期と共有タブの取り込み（worker はまだ起動しない）

```sh
timeout 590 ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
sg docker -c "docker compose exec -T api python -m app.worker run sync_masters"
sg docker -c "docker compose exec -T api python -m app.worker run sync_sheet_config"
psql_q() { sg docker -c "docker compose exec -T db psql -U overtime -d overtime -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
SELECT (SELECT count(*) FROM workers) workers, (SELECT count(*) FROM orders) orders, (SELECT count(*) FROM company_calendar) calendar,
       (SELECT count(DISTINCT dept) FROM dept_approvers) approver_depts;
SELECT value FROM settings WHERE key = 'ADMIN_EMAILS';
SQL
EOF
```

確かめ: `sync_masters: ok=True`（4 タブとも `ok`、作業員 50・工番 560・業務 NO 47・カレンダー 130 前後）、`sync_sheet_config: ok=True`。DB の 4 タブの行数と先頭の行が B-8 の控えと同じ（シートを開いて見る）。DeptApprovers 13 部署、管理者 2 人。戻す: 控えから 4 タブの値を写し戻す（D 担当の `tools/restore_tabs_from_copy.py`）。

### B-16 worker を起動して、行の反映と照合を確かめる

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
sg docker -c "docker compose up -d"
sleep 60
sg docker -c "docker compose ps"
sg docker -c "docker compose exec -T worker python -m app.worker run sheet_reconcile --check"
psql_q() { sg docker -c "docker compose exec -T db psql -U overtime -d overtime -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
SELECT count(*) FILTER (WHERE desired_version > synced_version) row_sync_pending, count(*) FILTER (WHERE state = 'blocked') blocked FROM sheet_sync_state;
SELECT lane, state, count(*) FROM jobs GROUP BY 1, 2 ORDER BY 1, 2;
SELECT kind, ok, started_at, summary_json->'problems' problems FROM sync_runs ORDER BY id DESC LIMIT 10;
SQL
sg docker -c "docker compose logs --since 2m worker" | grep -iE 'error|exception|unsafe' || echo "worker: エラーなし"
EOF
```

確かめ: `sheet_reconcile: ok=True` で A 区分の差 0、`row_sync_pending 0`・`blocked 0`、jobs は 0 件（まだ誰も操作していない）、最新の `gas_ping` が `ok = t`。シートの `DeptApprovers`・`Settings` が B-8 の控えと同じ（VPS は書かない）。戻す: `sg docker -c "docker compose stop worker"`、控えから写し戻して C へ。

### B-17 Sebastian の転送

GAS エディタで `setSebastianForwardUrl` を実行します（転送先は `https://overtime.lineworks-local.net/api/sebastian`）。続けて手元から、Sebastian が使っている `/exec` に `ping` を送ります。

```sh
URL="https://script.google.com/macros/s/〇〇/exec"     # Sebastian が使っているデプロイ
python - "$URL" <<'PY'
import getpass, json, sys, urllib.request
token = getpass.getpass("SEBASTIAN_API_TOKEN: ")
req = urllib.request.Request(sys.argv[1], data=json.dumps({"action": "ping", "token": token}).encode(),
                             headers={"Content-Type": "application/json"})
print(urllib.request.urlopen(req).read().decode())
PY
ssh lineworks-vps-user 'test "$(hostname)" = srv1508169 || exit 9; cd /opt/overtime && sg docker -c "docker compose logs --since 3m web" | grep "/api/sebastian"'
```

確かめ: 返事が `{"ok":true,"pong":"<UTC の ISO、ミリ秒と Z>"}`。web のログに `POST /api/sebastian … 200` が出る（GAS を通って VPS に届いた）。戻す: `clearSebastianForwardUrl`（移転モードのままなら `service is moving` に戻る）。

### B-18 画面と受け入れ（書き込みはしない）

1. 受け入れの道具を読むだけで流します。production ではセッションを作れないので、ログインの要らない項目（healthz・入口・ヘッダー・Sebastian）だけが通り、ほかは SKIP になります。

   ```sh
   ssh lineworks-vps-user 'bash -s' <<'EOF'
   set -eu
   test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
   cd /opt/overtime
   sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --expect-env production"
   EOF
   curl -s https://overtime.lineworks-local.net/healthz; echo
   ```

   確かめ: `FAIL 0`、`H-2`（env が production）が PASS、外からの healthz が `{"ok":true,"env":"production"}`。
2. オーナーのアカウントと部署の共有アカウント 1 つで、画面を開いて見ます（JEV ハンズ）。TOP（今日の一覧）、承認者画面、二次承認（B-1 の待ち件数と同じ）、総務の月次と日次、スタッフ集計、完了未記録の 1 件が TOP の期限超過に出ること。古いリンク `…/exec?page=approve2&date=<日付>` → 移転ページ → ボタン → （セッションを空にしたので）ログイン → 新しいアプリの同じ画面。**ここでは書き込みの操作をしません。**

### B-19 社内ポータルの行き先

**前にオーナーへ声をかけます**（下の「オーナーへの連絡」の 3）。アプリの安全装置はポータルのシートへの書き込みを拒むので、api コンテナの中から委任トークンで直接書きます。既定は何も書かずに見るだけで、`APPLY=1` のときだけ、url 列が想定の古い値のときに書きます。

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
sg docker -c "docker compose exec -T -e APPLY=0 api env PYTHONPATH=/app python -" <<'PY'
import json, os, urllib.parse, urllib.request
from app.google_delegate import get_access_token
PORTAL = "1DdhbJuz-zyS_2pKcM8y6Xf2pHofuzwNDgoW9eWcB0yg"
KEY = "〇〇"                       # APP_MASTER の key 列の値
NEW = "https://overtime.lineworks-local.net/"
OLD_PREFIX = "https://script.google.com/macros/s/"
H = {"Authorization": "Bearer " + get_access_token(), "Content-Type": "application/json"}
def call(method, url, body=None):
    req = urllib.request.Request(url, data=None if body is None else json.dumps(body).encode(), method=method, headers=H)
    return json.load(urllib.request.urlopen(req))
rows = call("GET", f"https://sheets.googleapis.com/v4/spreadsheets/{PORTAL}/values/{urllib.parse.quote('APP_MASTER!A1:Z500')}")["values"]
head = [h.strip().lower() for h in rows[0]]
kc, uc = head.index("key"), head.index("url")
hit = [(i + 1, r) for i, r in enumerate(rows) if len(r) > kc and r[kc] == KEY]
assert len(hit) == 1, f"key {KEY} の行が {len(hit)} 行です"
n, row = hit[0]
cur = row[uc] if len(row) > uc else ""
col = chr(ord("A") + uc)
print(f"行 {n} 列 {col}: 今の url = {cur}")
if os.environ.get("APPLY") == "1":
    assert cur.startswith(OLD_PREFIX), "url が想定の古い値（GAS の /exec）ではないので書きません"
    rng = urllib.parse.quote(f"APP_MASTER!{col}{n}")
    call("PUT", f"https://sheets.googleapis.com/v4/spreadsheets/{PORTAL}/values/{rng}?valueInputOption=RAW", {"values": [[NEW]]})
    print("書き換えました →", NEW)
PY
EOF
```

1 回目（`APPLY=0`）で行と今の値（控える。戻すときに使う）を確かめ、同じコマンドを `APPLY=1` にしてもう一度流します。確かめ: ポータルを開き直して、このアプリのボタンの行き先が新しい URL。戻す: 同じ形で控えた `/exec` を書き戻す。

### B-20 最終確認

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
sg docker -c "docker compose ps"
psql_q() { sg docker -c "docker compose exec -T db psql -U overtime -d overtime -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
SELECT lane, state, count(*) FROM jobs GROUP BY 1, 2 ORDER BY 1, 2;
SELECT count(*) FILTER (WHERE desired_version > synced_version) row_sync_pending FROM sheet_sync_state;
SELECT kind, ok, started_at FROM sync_runs WHERE kind = 'gas_ping' ORDER BY id DESC LIMIT 1;
SQL
sg docker -c "docker compose logs --since 30m api worker" | grep -iE ' error|exception|traceback' || echo "エラーなし"
ls -l backups/ | tail -3
EOF
curl -s -o /dev/null -w '%{http_code}\n' https://overtime.lineworks-local.net/healthz
```

確かめ: 6 コンテナが `Up`、jobs の pending・failed が 0、反映待ち 0、`gas_ping` が `t`、エラーなし、外から 200。GAS の「実行数」に、切り替え後のトリガーの実行が無い。`backups/` が空や 1KB 未満なら、起動直後の空の DB のダンプが捨てられただけ（罠 52）なので、`sg docker -c "docker compose exec -T backup /usr/local/bin/run_backup.sh"` で 1 回取り、`[backup] ok … bytes` を確かめます。

### B-21 記録と報告

引継ぎメモ（`docs/handoff-<日付>.md`）・メモリ・スキルの下書き（`docs/skill-learnings-draft.md`）を更新し、オーナーへ報告します。報告に入れるもの: 移した件数と状態別の内訳とシートとの一致、進行中の件数、控えの場所（DB と台帳の複製の ID、`cutover/` のダンプ、`.env.before-cutover`）、利用者から見た変化（ポータル・古いリンク・ログイン画面）、判断して変えた点、オーナーにお願いすること。

---

## C. 元に戻す（計画書 6-6。戻すと決めるのはオーナー）

決定 A でシートは常に最新、決定 B で PDF と台帳は初めから GAS が作っています。そのため、書き出しを待つ・写し戻す作業は要らず、「VPS を止めて、GAS の移転モードを外す」が中心です。

1. **C-1 VPS の api を止める**（新しい書き込みを止める。worker は動かしたまま）。

   ```sh
   ssh lineworks-vps-user 'test "$(hostname)" = srv1508169 || exit 9; cd /opt/overtime && sg docker -c "docker compose stop api" && sg docker -c "docker compose ps"'
   ```

2. **C-2 worker の待ちを流しきる**。行の反映の待ち 0、GAS の仕事の待ち 0（GAS が動いていれば数分）になるまで、次を繰り返します。照合を `--check` で 1 回流し、差 0 を確かめます。`failed` や流しきれない GAS の仕事は、申請 ID と action を控えます（C-5 で使う）。**最後に worker を止めます**（止めないと、GAS が戻ってから書いた行を、毎晩の照合が DB の値で書き直します）。

   ```sh
   ssh lineworks-vps-user 'bash -s' <<'EOF'
   set -eu
   test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
   cd /opt/overtime
   psql_q() { sg docker -c "docker compose exec -T db psql -U overtime -d overtime -v ON_ERROR_STOP=1 -X -P footer=off"; }
   psql_q <<'SQL'
   SELECT count(*) FILTER (WHERE desired_version > synced_version) row_sync_pending, count(*) FILTER (WHERE state = 'blocked') blocked FROM sheet_sync_state;
   SELECT id, request_id, kind, state, attempts, left(last_error, 120) FROM jobs WHERE state NOT IN ('done', 'skipped') ORDER BY id;
   SQL
   EOF
   # 待ちが 0 になったら:
   ssh lineworks-vps-user 'test "$(hostname)" = srv1508169 || exit 9; cd /opt/overtime && sg docker -c "docker compose exec -T worker python -m app.worker run sheet_reconcile --check" && sg docker -c "docker compose stop worker"'
   ```

3. **C-3 GAS を戻す**: GAS エディタで `exitMigrationMode`（Sebastian の転送も外れます。内部用の入口は token が無ければ何もしないので残してよい）→ `restoreTriggersAfterRollback`（朝バッチとマスタ同期の 2 本。同じ関数のトリガーが既にあれば作りません）。旧フォームは決定 #11 のとおり閉じたままにします。**オーナーが旧フォームも戻すと決めたときだけ**、`restoreTriggersAfterRollback` の代わりに `reopenFormsAfterRollback` → `restoreAllTriggersAfterRollback`（4 本）を実行します。確かめ: `showMigrationStatus` で移転モード「なし」、トリガーが 2 本（フォームも戻したなら 4 本）。
4. **C-4 マスタ同期を手で 1 回**: GAS エディタで `syncAllMasters` を実行し、ログの `OK` 4 行を確かめます（VPS が書いていた DB のマスタ 4 タブに、GAS の同期が今の形で書けること。gap-4 F-6）。
5. **C-5 控えた仕事を GAS で処理する**: C-2 で控えた申請ごとに、メニューの PDF の作り直し、台帳は管理の関数で upsert（オーナーと相談）。
6. **C-6 確かめ**: 本番 `/exec` で一覧が出て、VPS で最後に作られた申請が見えること。Sebastian の `/exec` に `ping`（B-17 のコマンド）を送り、GAS の返事（`{ok:true, pong}`）が返ること。
7. **C-7 ポータル**: url 列を B-19 で控えた `/exec` に戻す（B-19 のコマンドの `NEW` を控えた値に、`OLD_PREFIX` を `https://overtime.lineworks-local.net` にして流す）。Settings の `APP_URL` は変えていないので何もしません。
8. **C-8 `.env`**: 必要なら `cp -p cutover/.env.before-cutover .env`。api と worker は、次に切り替えると決めるまで止めたままにします。戻した後は VPS の DB を使いません。もう一度切り替えるときは移行（B-8 から）からやり直します。
9. **C-9 利用者への案内**: オーナーと相談します（「オーナーへの連絡」の 5）。

- その日の朝バッチを VPS が既に流していれば、GAS の朝バッチは翌朝から動かします。同じ日に GAS の朝バッチを手で流すと、BatchLogs に 2 行目ができ、台帳の前日分をもう一度書き直します（結果は同じ）。
- 決定 A で型と表示形式を保つので、戻した後に日時の見え方が変わる行は出ない見込みです（リハーサルの戻しの練習で確かめる）。

### 当日の途中で戻すとき（B の「戻す範囲」）

- B-3〜B-7 の途中: 実行した GAS の関数を逆の順に戻します（`restoreTriggersAfterRollback`／`exitMigrationMode`／`reopenFormsAfterRollback`）。VPS は何も変えていません。
- B-9〜B-13 の途中: VPS は `sg docker -c "docker compose start api worker"`（ステージングのまま）、表はダンプから `pg_restore`、`.env` は控えを戻します。GAS は上と同じに戻します。シートはまだ VPS から書かれていません。
- B-14 以降: C 章の手順で戻します。

---

## D. 見守り（戻せる期間）

翌朝（6-5）に確かめること:

- 最初の申請が 1 分以内にシートの Requests の末尾に出たこと（型と表示形式も）。
- 最初の PDF が GAS の内部用の入口で作られ、DB とシートの 3 列が同じこと。本番で最初に作られた PDF 1 件をオーナーが見る。
- VPS の朝バッチ（6:50 ごろ）が前日分の PDF を作り、BatchLogs に `morningBatchNoMail` の 1 行、台帳に前日分を書いたこと（7:05 の確かめの結果も）。
- 統合メール通知システムの通知がいつもどおり出たこと。Sebastian の完了が届き、PDF ができたこと。

毎日見るもの:

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/overtime
psql_q() { sg docker -c "docker compose exec -T db psql -U overtime -d overtime -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
SELECT lane, state, count(*) FROM jobs GROUP BY 1, 2 ORDER BY 1, 2;
SELECT count(*) FILTER (WHERE desired_version > synced_version) pending, count(*) FILTER (WHERE state = 'blocked') blocked FROM sheet_sync_state;
SELECT kind, ok, started_at FROM sync_runs WHERE kind NOT IN ('gas_ping') ORDER BY id DESC LIMIT 12;
SELECT kind, ok, started_at FROM sync_runs WHERE kind = 'gas_ping' ORDER BY id DESC LIMIT 1;
SQL
ls -l backups/ | tail -3
sg docker -c "docker compose logs --since 24h api worker" | grep -ciE ' error|exception' || true
EOF
```

- `failed` の仕事が無い、反映待ち・blocked が 0、マスタ同期・共有タブ取り込み・毎晩の照合・朝バッチが `ok`、バックアップが 1KB 超。管理者の画面の「運用（管理者）」欄でも同じものが見えます。
- 完了済みで PDF なしの 25 件は切り替えでは触りません（Q13 で決める）。
- 戻せる期間が過ぎたら: ステージングの DB・台帳・PDF フォルダ・GAS の写しの片付けを、オーナーに確かめてから行います。

---

## オーナーへの連絡

1. **前日**: 当日の流れと開始時刻、利用者への案内を出すか（前例は「静かに」で出していない）、部署の共有アカウントで前日に 1 回ログインしてもらうか（A-7）、旧フォームを閉じること（決定 #11）を確かめる。
2. **当日の開始前**: 「切り替えを始めます」と声をかけ、了承をもらう（決定 #17）。本番 GAS の変更（A-3 の `master` への push と版の反映）も、この前に許可をもらう。
3. **ポータルの書き換えの前**（B-19）: 「ポータルの行き先を新しいアドレスに変えます」と声をかける。
4. **終わったら**: B-21 の報告。お願いすること（翌朝最初の PDF を見る、iPhone での表示の写真など）を最後に分けて書く。
5. **戻すとき**: 戻すかはオーナーが決める。戻した後の利用者への案内（「元のアドレスに戻しました」）もオーナーと相談する。

## 確かめること

- 移転モード入りの版の番号、Sebastian が使っているデプロイ、本番の内部用のデプロイ ID（A-3）。
- ポータルの `APP_MASTER` の key と url 列（B-19）。
- 本番の内部用トークンの手元のファイル名（A-3・B-13）。
- `DB_SHEET_TIMEZONE` の値（A-4・B-12）。
- リハーサルで控えた件数（B-11・B-12 の比べる相手）。
- `tools/cutover_backup.py`・`tools/restore_tabs_from_copy.py`（D 担当）が出来上がったら、B-8・B-15 をそれに置き換える。
