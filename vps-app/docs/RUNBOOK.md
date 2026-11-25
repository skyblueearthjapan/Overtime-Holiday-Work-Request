# 残業・休日出勤申請 VPS 版 運用手順（骨組み）

デプロイ・戻し・秘密の置き場の手順です。値の正は `docs/vps-migration-plan.md` の 4-0 章と `docs/vps-implementation-spec.md` の 1 章（リポジトリの直下の docs）です。作業の前に必ず読み合わせます。切り替えの手順（移転モード・流しきり・戻し）は段階 5 で足します。

## 最初に確かめること（毎回）

- `ssh lineworks-vps-user "hostname"` が `srv1508169` を返すこと。違えばすぐ中止します。
- `srv1300618` には絶対に触れません（別会社の本番が同居しています）。
- 置き場は `/opt/overtime`。リポジトリの `vps-app/` の**中身**をここに直置きします（git のフォルダではありません。`.env` も直下、権限 600）。ほかの `/opt/*` とコンテナには触れません。
- docker は `sg docker -c '...'` で動かします（lineworks ユーザーは sudo 不可）。
- コンテナ・ボリューム・ネットワークは `overtime-` で始まるものだけを扱います（`overtime-db`・`overtime-api`・`overtime-worker`・`overtime-web`・`overtime-cloudflared`・`overtime-backup`、`overtime-pgdata`、`overtime-net` 10.93.0.0/24）。

## 秘密の置き場

| もの | 置き場 | 備考 |
|---|---|---|
| `.env`（DB のパスワード、OAuth の secret、SESSION_SECRET、FERNET_KEY、GAS_INTERNAL_TOKEN、SEBASTIAN_API_TOKEN、TUNNEL_TOKEN） | VPS の `/opt/overtime/.env`（`chmod 600`） | git に入れない。見本は `.env.example`。確かめるときは `grep -E '^(APP_ENV\|APP_IMAGE_TAG)=' .env` のように秘密でない値だけを出す |
| 手元の控え | `C:\Users\imaizumi.LINEWORKS-NET\.secrets\overtime\` | git の外。Tunnel のトークンは `tunnel-token.env` |
| 委任トークン（`imaizumi@lineworks-local.info` のスプレッドシート・ドライブ） | DB の `google_credentials` 表（FERNET_KEY で暗号化） | FERNET_KEY を変えたら、オーナーが `/auth/google-delegate/start` から許可し直す |

## 初回の準備（ディレクターが行う。実行前にオーナーへ値を見せて一声かける）

1. `/opt/overtime` は作成済み（lineworks 所有、750）。Cloudflare Tunnel `overtime-vps`（`overtime.lineworks-local.net` → `http://web:80`）も作成済み。
2. 手元の `.env`（`.env.example` から作る）を送る: `ssh lineworks-vps-user 'umask 077; cat > /opt/overtime/.env && chmod 600 /opt/overtime/.env' < vps.env`
3. OAuth クライアント（GCP `sebastian01` の「Overtime Holiday Work Request (VPS)」）にリダイレクト URI を 2 つ登録する。
   - `https://overtime.lineworks-local.net/auth/callback`
   - `https://overtime.lineworks-local.net/auth/google-delegate/callback`
4. 下の「デプロイ」で起動する（staging。分からない ID は罠 50 のとおり仮の値で起動してよい）。
5. オーナーに、ログインしてから `https://overtime.lineworks-local.net/auth/google-delegate/start` を開いてスプレッドシート・ドライブを許可してもらう（未ログインで開くと 1 回目はログインだけで終わるので、もう一度同じアドレスを開く）。

## デプロイ（git archive 方式）

手元のリポジトリで、**コミット済みの** `vps-app/` を送ります（未コミットの変更は送られません。手元で消したファイルは VPS に残るので、要らなくなったものは VPS でも消します）。

```sh
cd /d/projects/zangyo-kyujitsu-app/github-repo
git archive --format=tar HEAD vps-app | ssh lineworks-vps-user 'test "$(hostname)" = srv1508169 || exit 9; rm -rf /tmp/overtime-sync && mkdir -p /tmp/overtime-sync && tar -x -C /tmp/overtime-sync && cp -a /tmp/overtime-sync/vps-app/. /opt/overtime/ && rm -rf /tmp/overtime-sync && echo synced'
```

イメージにタグを付けて出し、タグは `.env` の `APP_IMAGE_TAG` に書きます（コマンド行だけで渡すと、後の `docker compose up -d` で古いタグに戻ります＝罠 24e）。

```sh
TAG=$(git rev-parse --short HEAD)    # 手元で求める（VPS は git のフォルダではない）
ssh lineworks-vps-user 'grep -E "^APP_IMAGE_TAG=" /opt/overtime/.env; sg docker -c "docker image ls overtime-app"'   # 今のタグを控える
ssh lineworks-vps-user "test \"\$(hostname)\" = srv1508169 || exit 9; cd /opt/overtime && sed -i 's/^APP_IMAGE_TAG=.*/APP_IMAGE_TAG=$TAG/' .env && sg docker -c 'docker compose build api' && sg docker -c 'docker compose up -d' && sg docker -c 'docker compose ps'"
```

- build と起動は時間がかかります。`timeout 590 ssh ...` を裏で流します（罠 2）。
- api は起動のたびに `alembic upgrade head` を流してから立ち上がります。
- 設定の値が欠けている・危ない組み合わせ（staging なのに本番の ID、本番なのにステージングの GAS など）のときは、api と worker は起動しません（ログに名前が出ます）。
- 確かめ: `sg docker -c "docker compose exec -T web wget -qO- http://127.0.0.1/healthz"` が `{"ok":true,"env":"staging"}`。外から `curl -s https://overtime.lineworks-local.net/healthz` が 200。
- 入口の確かめ（curl）: `/` 302（ログインへ）、ログインなしの RPC 401、`/auth/login` 302（`hd`・PKCE・state・nonce 付き、`prompt` なし）、CSP などのヘッダ、JSON でない本文 415、別サイトの Origin 403、知らない関数 404。
- 初めて出すとき・`OVERTIME_NET_SUBNET` を変えたときは `docker compose down` → `up -d` でネットワークを作り直します。
- 切り替えでは、先にシートの控え → **api だけ**起動して安全装置を確かめる → worker を起動（罠 24）。

## 試験

- 手元: `cd vps-app/api && .venv/Scripts/python -m pytest -q`（SQLite。最後の行の passed/failed まで読む＝罠 44）。
- VPS の結合試験（PostgreSQL）: `docker build --target test -t overtime-test ./api` の後、DB `overtime_test` を作り `TEST_DATABASE_URL` を渡して `python -m pytest`（罠 51。DB 名に test が無ければ conftest が拒む）。

## 道具（tools/）

`tools/` はイメージに入っていません（罠 5）。`migrate_from_sheets.py`（移行。staging では本番の PDF の ID を空にし、元の値を `source_pdf_*` に残す）、`cutover_backup.py`（DB と台帳の控え）、`restore_tabs_from_copy.py`（控えから写し戻す。既定は確かめだけ）。使い方は各ファイルの先頭と `docs/cutover-runbook.md`。

```sh
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api python tools/<道具>.py ..."
```

## worker と定期処理

`python -m app.worker run <名前>` で手で動かせます（例 `run sync_masters`、`run sheet_reconcile --check`）。結果は `sync_runs`（kind=`schedule:<名前>`）に残ります。

| 名前 | いつ（日本時間） | 持ち主 |
|---|---|---|
| `sync_masters` | 毎日 0:30 | D |
| `sheet_reconcile` | 毎日 2:10 | D |
| `sync_sheet_config` | `SHEET_CONFIG_SYNC_MINUTES` ごと | D |
| `morning_batch` | 毎日 6:50 | M |
| `gas_ping` | 5 分ごと（起動時にも 1 回） | M |
| `purge` | 毎日 3:30（冪等キー 24 時間・期限切れのセッション・done の仕事 90 日。audit_log は消さない） | F |

GAS の仕事（`jobs` の `lane='gas'`）は全体で同時に 1 本。失敗は管理者の「運用（管理者）」欄（`ops_getStatus`）に出ます。

- シートへ書く前の関所（`services/write_gate`）: 本番は、合格した `gas_ping` が `SHEET_WRITE_PING_MAX_MINUTES`（既定 10）分以内で、GAS が移転モード（`migrated=true`）のときだけ、行の反映・照合・マスタ同期がシートへ書きます。ステージングは、ステージングの GAS への合格した `gas_ping` があれば書きます。止めているときは `ops_getStatus` の `sheetWriteGate.blocked` が `true` で、理由が出ます。
- `SKIP_DAILY_ON_STARTUP=1` で起動すると、起動した時に枠の中にある毎日の予定（0:30・2:10・6:50 など）を、その回だけ飛ばします（`sync_runs` に `skipped`）。切り替えの夜に使い、済んだら `.env` から消します。
- blocked の行（人がシートの行を消したなど）: `docker compose exec -T worker python -m app.services.row_sync status` で一覧、行を直したら `unblock <table> <pk>`、行が消えたなら `reappend <table> <pk>`（末尾に新しい行として書き直す）。
- 台帳の作り直し（`ops_rebuildLedger`）は、行の反映待ちが 0 のときだけ積めます。時間切れ・失敗しても送り直しません（台帳を見てから手で動かし直す）。

## ログインの取り消し

- 1 人（共有アカウント 1 つ）: 管理者の欄の `ops_revokeSessions`（そのメールの全端末）。退職・共有パスワードの変更のときは必ず使います。
- 全員: `sg docker -c "docker compose exec -T db psql -U overtime -d overtime -c 'DELETE FROM sessions'"`（全員がもう一度ログインします。セキュリティ確認 中-4 の当面の手当て）。
- ログインの成功・失敗は api のログ（`login ok`・`login rejected`）に残ります。

## バックアップ

- `overtime-backup` が毎日 3:00 に `backups/overtime_*.dump` を取り、`BACKUP_KEEP_DAYS` 日で消します。1KB 未満は `FAILED`（初回起動の直後は正常＝罠 52）。
- 切り替えの控えなど手で取ったダンプは `/opt/overtime/cutover/` に置きます（backups/ は消える＝罠 25）。

## 戻す（骨組み）

- アプリの版を戻す: `.env` の `APP_IMAGE_TAG` を控えたタグに書き戻して `docker compose up -d`。
- 本番の切り替えの戻し（GAS に戻す）は段階 5 で書きます（計画書 6-6）。
