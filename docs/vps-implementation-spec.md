# 残業・休日出勤申請 VPS 版 実装仕様（エージェント共通の取り決め）

作成: 2026-11-26（段階 2）。対象: 本番 GAS @131（scriptId `1GPNBR9U…`）を VPS（FastAPI＋PostgreSQL＋worker＋Cloudflare Tunnel）へ移す。
読む順: この文書 → `docs/vps-migration-plan.md`（4 章・5-5・6-3・8-4・9 章）→ `docs/vps-migration-gap-audit.md` → `docs/migration-inventory/` の自分の担当に関わる文書 → 自分の担当の GAS コード（`src/`）。

**方針: 今の GAS 版の機能と使い勝手をまったく残す。変えるのは 11 章「変えるもの」だけ。** 迷ったら GAS のコードの動きに合わせ、報告に書く。決定事項は計画書 9-1 の #1〜#25 と決定 A〜D（#18〜#21）。この文書と計画書が食い違えば、9-1 の決定が勝ち、次にこの文書が勝つ。

> 書き方: 「必須」は外せない約束。「既定」はオーナーが未回答の質問（計画書 9-2）にディレクターの案で決めた値で、設定 1 か所で変えられるように作る。〇〇 は段階 2b のスパイクで決まる値。秘密の値はこの文書にも git にも書かない。

参照実装（読むだけ。書き換えない）:

- 土台: `D:\projects\chozouhin-shinsei-app\vps-app\`（貯蔵品。auth・google_delegate・rpc・jobs・worker・safety・services/sheets・tools/migrate_from_sheets・compose・web）
- 業務 ID を `.env` 1 か所で読む作り・`jobs.register_kind`・`sync_runs` の予定枠・マスタ同期・書き出しの歯止め・受け入れの道具: `D:\projects\kouban-buhin-shinsei\vps-app\`（工番部品。config・jobs・worker・services/sync・services/sheets・tools/qa/acceptance.py）
- 共通スキル: `C:\Users\imaizumi.LINEWORKS-NET\.claude\skills\gas-to-vps-migration\`（`references/02-architecture.md` に参照実装の解説、`07-pitfalls.md` に罠 1〜59）

---

## 目次

1. 確定値
2. リポジトリの形と、写すもの・捨てるもの
3. 画面（5 ページ＋権限エラー）
4. RPC（画面の 29 関数・管理用・Sebastian の入口）
5. DB の表
6. シートへの行単位の反映（決定 A）・シートの取り込み（6-10、決定 #26）
7. GAS の内部用の入口と、きっかけごとの呼び順（決定 B）
8. 同じにしなければならない業務の決まり
9. 安全装置
10. 担当と境目（担当をまたぐ関数の表、作る順）
11. 変えるもの／変えないもの
12. テストと受け入れ表
13. 意図して残す固有の文字列

---

## 1. 確定値

| 値 | このアプリ | 出典 |
|---|---|---|
| アプリ名（フォルダ・DB・接頭辞） | `overtime` | 計画書 4-0 |
| 公開 URL | `https://overtime.lineworks-local.net`（staging も同じホスト名。切り替えで `APP_ENV` だけ変える） | 9-1 #2 |
| VPS | `srv1508169`（`ssh lineworks-vps-user`。作業前に `hostname`）。`srv1300618` には絶対に触れない | スキル |
| VPS の置き場 | `/opt/overtime`（`vps-app/` の中身を `git archive` で直置き。`.env` も直下、権限 600） | 4-0 |
| Tunnel 名 | `overtime-vps`。転送先 `http://web:80` | 4-0 |
| docker のサブネット | `10.93.0.0/24`。web の固定 IP `10.93.0.10` | 4-0 |
| compose のプロジェクト名 | `overtime` | — |
| コンテナ名 | `overtime-db`・`overtime-api`・`overtime-worker`・`overtime-web`・`overtime-cloudflared`・`overtime-backup` | 4-1 |
| ボリューム・ネットワーク | `overtime-pgdata`・`overtime-net` | — |
| DB | PostgreSQL 16。DB 名 `overtime`、ユーザー `overtime`。結合テスト用は `overtime_test`（名前に `test` が無ければ conftest が拒む） | — |
| クッキー | `__Host-overtime_session`・`__Host-overtime_login`・`__Host-overtime_delegate` | — |
| バックアップのダンプ名 | `overtime_*.dump`（毎日 3 時、1KB 未満は FAILED）。切り替えの控えは `/opt/overtime/cutover/` | スキル罠 25・26 |
| 手元の秘密の置き場 | `C:\Users\imaizumi.LINEWORKS-NET\.secrets\overtime\` | 4-0 |
| 本番 DB スプレッドシート | `1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ`（PDF テンプレも同じ） | 4-0 |
| マスタ元（`SOURCE_SSID`） | `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ`（読むだけ） | 4-0 |
| 蓄積台帳 | `1tKlT57SJxJijgfDVGafZ434sdVjscJJqN2uQ7P3sUyU`（VPS は書かない。GAS が書く） | 4-0 |
| 本番 PDF フォルダ | `1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B`（VPS は書かない。GAS が書く） | 4-0 |
| GAS scriptId（本番） | `1GPNBR9U9mw7TZPnl1pxtdVjFagUPxt9S6qSYlbNIEuHNvRLVspNIQNc7` | 4-0 |
| 委任の持ち主・オーナー | `imaizumi@lineworks-local.info`。Workspace のドメインは `lineworks-local.info` | 4-0 |
| 社内ポータルの URL（`PORTAL_URL`） | `src/adminApi.gs:554` の固定値をそのまま写す | [02 §1.2] |

### 1-1. 環境変数（`.env`、`vps-app/.env.example` に全キーを書く）

読むのは `api/app/config.py` の 1 か所だけ（工番部品の型）。値が欠ければ api も worker も起動しない（fail-closed）。「秘密」は `.env.example` に値を書かず、`CHANGE_ME` だけ置く。

| キー | 秘密 | 中身・検証 |
|---|---|---|
| `APP_ENV` | — | `staging`／`production` |
| `DATABASE_URL` | 秘密 | `postgresql+psycopg://overtime:<pw>@db:5432/overtime` |
| `POSTGRES_USER`／`POSTGRES_DB` | — | `overtime`／`overtime` |
| `POSTGRES_PASSWORD` | 秘密 | |
| `PUBLIC_ORIGIN` | — | `https://overtime.lineworks-local.net`（`https://` 必須、末尾 `/` なし） |
| `GOOGLE_OAUTH_CLIENT_ID` | — | GCP `sebastian01` の「Overtime Holiday Work Request (VPS)」 |
| `GOOGLE_OAUTH_CLIENT_SECRET` | 秘密 | |
| `SESSION_SECRET` | 秘密 | 32 文字以上 |
| `SESSION_MAX_AGE_DAYS` | — | **`90`**（9-1 #24） |
| `SESSION_IDLE_HOURS` | — | **`2160`**（90 日。無操作でも 90 日は切らない。9-1 #24） |
| `ALLOWED_DOMAIN` | — | `lineworks-local.info` |
| `BOOTSTRAP_ADMINS` | — | `imaizumi@lineworks-local.info`（空なら起動しない） |
| `FERNET_KEY` | 秘密 | 委任トークンの暗号化。変えたら委任の許可し直し |
| `DELEGATE_EMAIL` | — | `imaizumi@lineworks-local.info` |
| `PROD_SPREADSHEET_ID` | — | 本番 DB。コードの定数と違えば起動しない |
| `STAGING_SPREADSHEET_ID` | — | staging で必須。本番 DB の「コピーを作成」で作ったサンドボックス。本番と同じなら起動しない |
| `MASTER_SOURCE_SPREADSHEET_ID` | — | マスタ元。定数と違えば起動しない |
| `PROD_PDF_FOLDER_ID`／`STAGING_PDF_FOLDER_ID` | — | VPS は書かない。`internal.ping` の照合にだけ使う。staging では両方必須で、同じなら起動しない |
| `PROD_LEDGER_SPREADSHEET_ID`／`STAGING_LEDGER_SPREADSHEET_ID` | — | 同上（蓄積台帳。staging は台帳のコピー） |
| `GAS_INTERNAL_URL` | — | GAS の内部用デプロイの `/exec`。`^https://script\.google\.com/macros/s/[A-Za-z0-9_-]+/exec$` |
| `GAS_INTERNAL_TOKEN` | 秘密 | GAS のスクリプトのプロパティ `VPS_INTERNAL_TOKEN` と同じ値（32 バイト以上の乱数） |
| `GAS_CALL_TIMEOUT_SECONDS` | — | 既定 `330` |
| `SEBASTIAN_API_TOKEN` | 秘密 | GAS のスクリプトのプロパティ `SEBASTIAN_API_TOKEN` の値を写す（gap-4 F-16） |
| `SEBASTIAN_PDF_WAIT_SECONDS` | — | 既定 `40`。GAS の転送（`UrlFetchApp`）の待ち時間より短く。スパイクで 〇〇 に直す |
| `DB_SHEET_TIMEZONE` | — | DB スプレッドシートの時間帯（IANA 名）。〇〇（スパイクで「ファイル → 設定」と Sheets API の `properties.timeZone` を読んで決める）。シートの値と違えば行の反映を止める（gap-5 高-2） |
| `PORTAL_URL` | — | 社内ポータルの固定 URL |
| `SHEET_CONFIG_SYNC_MINUTES` | — | 既定 `5`（Q12 の既定） |
| `FORWARDED_ALLOW_IPS` | — | `10.93.0.10` |
| `OVERTIME_NET_SUBNET`／`OVERTIME_WEB_IP` | — | `10.93.0.0/24`／`10.93.0.10`（compose だけ） |
| `APP_IMAGE_TAG` | — | イメージのタグ（compose。切り替えの前に `.env` の値を確かめる＝罠 24e） |
| `TUNNEL_TOKEN` | 秘密 | compose だけ |
| `BACKUPS_HOST_PATH`／`BACKUP_KEEP_DAYS` | — | compose だけ |

使わないキー（貯蔵品・工番部品にあって、このアプリでは持たない）: `MAIL_REDIRECT_TO`・`MAIL_ALLOWED_DOMAINS`（メールを送らない。計画書 4-11）、`EXPORT_HOURLY_UNTIL`（決定 A で毎晩のタブ置き換えをしない）。

委任のスコープ（最初に決め、後から足さない＝罠 7）: `openid email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive`。`drive` は切り替えの控え（DB と台帳の複製）とステージングの準備の道具だけが使う。アプリのコード（api・worker）は Drive を呼ばない。`gmail.*` は要らない。

コードに定数で持つ ID は `config.py` だけ（`PROD_SPREADSHEET_ID_FIXED`・`MASTER_SOURCE_SPREADSHEET_ID_FIXED`・`PROD_GAS_SCRIPT_ID_FIXED`・`PROD_PDF_FOLDER_ID_FIXED`・`PROD_LEDGER_SPREADSHEET_ID_FIXED`）。ほかのモジュールに ID を直書きしない（`test_config_safety.py` が拒む）。

---

## 2. リポジトリの形と、写すもの・捨てるもの

### 2-1. 置き場所

- 新しいフォルダ `github-repo/vps-app/`（このリポジトリの直下）。土台は貯蔵品の `vps-app/` を写して作る。
- **GAS のファイル（`src/`）は G 担当だけが変える。** `master` へ `src/**` を push すると GitHub Actions（`.github/workflows/deploy-gas.yml`）が `clasp push --force` で GAS の HEAD を書き換え、本番のトリガー（HEAD のコードで動く）が新しいコードで動き出す。**`src/` の変更を `master` に入れるのは、オーナーの許可を得たディレクターだけ**（9-1 #17「本番 GAS の変更は事前に聞く」）。G は自分のブランチにコミットし、push しない。`vps-app/` の変更は Actions を動かさない。
- 担当エージェントは VPS・Cloudflare・Google・clasp・GitHub に触れない（ssh・デプロイ・外部 API・`git push` をしない）。手元のテストだけ。

```
vps-app/
  docker-compose.yml  .env.example
  api/
    Dockerfile  requirements.txt  requirements-dev.txt  pytest.ini  alembic.ini
    alembic/  env.py  versions/0001_initial_schema.py
    app/
      main.py  config.py  db.py  models.py  timeutil.py  jsutil.py  constants.py
      safety.py  auth.py  google_delegate.py  permissions.py  rpc.py  serialize.py
      sheet_queue.py  audit.py  jobs.py  worker.py
      pages.py                       # ?page= の入口（U）
      templates/pages/top.html approver.html approve2.html admin.html staffSummary.html no_auth.html
      domain/   register.py requests.py worklog.py approval.py secondary.py dashboards.py
                summaries.py masters.py breaks.py fiscal.py sebastian.py ops.py
      services/ sheets_api.py gas_client.py           # F
                sheet_codec.py row_sync.py reconcile.py config_sync.py master_sync.py   # D
                gas_jobs.py morning_batch.py ops_status.py                               # M
    tests/
  web/
    nginx.conf
    static/js/gas-shim.js  static/js/top.js approver.js approve2.js admin.js staffSummary.js no_auth.js
    static/css/mobile.css
    dev/mock-server.js  dev/shim.test.js
  backup/backup_entrypoint.sh
  cloudflare/setup_tunnel.py
  tools/migrate_from_sheets.py  tools/cutover_backup.py  tools/restore_tabs_from_copy.py
  tools/qa/acceptance.py
  docs/RUNBOOK.md
gas-tests/            # G。GAS の関数をスタブで動かす node の試験（src の外なので clasp は送らない）
src/Internal.gs src/Cutover.gs   # G（新）。ほかは 7-6 の決まった行だけ直す
```

### 2-2. 貯蔵品から写すもの

| 写すもの | 直すこと |
|---|---|
| `docker-compose.yml` | 名前・接頭辞・サブネット・固定 IP を 1 章に。`api` と `worker` は同じイメージ、`worker` は `healthcheck: disable: true`、`depends_on: api: service_healthy` |
| `api/Dockerfile` | `python:3.12-slim-bookworm`。**playwright・chromium・fonts-noto-cjk は入れない**（VPS は PDF を作らない）。`tzdata curl` だけ。マルチステージ（`base`→`test`→`runtime`、最後が `runtime`、`USER app`） |
| `requirements*.txt` | playwright・segno を外す。版はすべて固定 |
| `web/nginx.conf` | 3-7 の形（`/` を api のページへ、`/static/` だけ nginx） |
| `app/db.py`・`timeutil.py`・`main.py`・`alembic/env.py` | ほぼそのまま |
| `app/auth.py` | クッキー名、`prompt=select_account` を**外す**（gap-2 S2）、`/auth/me` を足す、期限は 1-1 |
| `app/google_delegate.py` | スコープを 1-1 に。`get_access_token` は残す（道具が呼ぶ＝罠 6） |
| `app/rpc.py` | 関数の一覧を 4 章に。冪等キー・Origin 照合・`commit` を rpc が行う形はそのまま |
| `app/safety.py` | 9 章の許可リスト（タブ単位・GAS の照合を足す） |
| `app/jobs.py`・`worker.py` | 工番部品の `register_kind` と `sync_runs` の予定枠を写し、レーン（10-3）を足す |
| `services/sheets.py` | API の呼び出し部分を `services/sheets_api.py` へ。書き出し（タブの置き換え）は捨てる |
| `tools/migrate_from_sheets.py` | 骨組み（`--source sheets`・`--dry-run`・`--snapshot-out`・空でなければ止める・1 トランザクション・`verify_in_db`・終了コード 2/1）。表の中身は書き直す |
| `backup/`・`cloudflare/setup_tunnel.py` | 接頭辞・ZONE・HOSTNAME・TUNNEL_NAME・DNS の comment。初回は「同じ名前の Tunnel があれば止める」行を足す（罠 24d・スキル 03 §5） |
| `tests/conftest.py`・`test_rpc.py`・`test_auth.py`・`test_config_safety.py`・`test_migrations.py`・`test_web_config.py`・`test_jobs.py` | このアプリの名前と関数に |

工番部品から写すもの: `config.py` の作り（業務 ID を `.env` の必須値にして定数と照合）、`jobs.register_kind(kind, handler, on_exhausted)`、`worker.py` の `sync_runs` で予定枠を判定する形と `python -m app.worker run <名前>`、`services/sync.py` の守り（読めない・空・有効な行 0 なら上書きしない）、`services/sheets.py` の `export_tables` の歯止め（DB の行が少なければ書かない）、`tools/qa/acceptance.py` の形。

### 2-3. 捨てるもの

HTML→PDF（`services/pdf.py`・`templates/pdf.html.j2`・`PdfRenderer`・`test_pdf.py`・chromium マーカー）、QR（segno）、メール（`services/mail.py`・`recipients.py`・`mailbody.py`・`mailpdf_common.py`・`mail_log` 表・`test_mail.py`）、Drive の保存（`services/drive.py`・`test_worker_drive.py`）、貯蔵品の業務（`domain/workflow.py`・`labels.py`・kiosk・`expectedStep`・段・分割・リコール）、`tools/resend_after_cutover.py`（送り直すメールが無い。計画書 6-5）、工番部品の xlsx（`partslist.py`・`jsnum.py` のうち xlsx 用）とまとめマスター（`archive.py`）、`web/static/index.html`・`app.js`（このアプリは 5 ページを移す）、`/vendor/`。

### 2-4. 段階 3 の完了条件（全担当）

次の grep が、意図して残したもの（13 章に理由を書く）を除いて 0 件。

```sh
grep -rnE "supplies|chozouhin|Chozouhin|CHOZOUHIN|1VKbB13|1HzP6wEF|1xCkpAIB|10\.91\.|10\.92\.|貯蔵品|\bparts\b|parts-|PARTS_|partslist|Job-Parts|ForParts|1QpPgRSH|1z4w_2izb|工番購入|部品表|kiosk|expectedStep|REQ-yyyy" vps-app/ gas-tests/ src/Internal.gs src/Cutover.gs --exclude-dir=.venv --exclude-dir=node_modules --exclude-dir=__pycache__
```

`1Knx_kaQ`・`1iu5Hoak`・`1tKlT57S`・`1Bs1FvgR`・`1GPNBR9U`・`imaizumi@` はこのアプリの正しい値なので、`config.py`・`.env.example`・`docs/RUNBOOK.md`・`src/Cutover.gs`（持ち主の確認）にだけ現れてよい。

---

## 3. 画面（5 ページ＋権限エラー）

### 3-1. 方針

- `src/top.html`・`approver.html`・`approve2.html`・`admin.html`・`staffSummary.html`・`no_auth.html` を、**HTML・CSS・文言・JS の処理をそのまま**にして api のテンプレート（Jinja2、自動エスケープ有効）で返す。PC での見た目は 1 px も変えない（9-1 #25）。
- 変えてよいのは次だけ: テンプレートの差し込みの書き方（3-3）、`google.script.run` を `gas-shim.js` に置き換えること（3-4。呼び出しの行は書き換えない）、CSP のための JS の置き場所とイベントの付け方（3-6）、XSS の直し（3-8）、スマホの CSS（3-9）、総務の日次一覧の直し（4-3 の 25 番、9-1 #22）、運用の欄の追加（3-10）。
- 「今日」をブラウザの時計で決める箇所（[02 §5]）はそのまま。

### 3-2. 入口 `GET /`（`app/pages.py`、U）

| 項目 | 決まり |
|---|---|
| ログイン | 全ページで必須（9-1 #3）。未ログインは `302 /auth/login?next=<今のパス＋クエリ>`（`?page=…&date=…` を保つ）。`next` は `auth.safe_next` を通す |
| `page` | 無指定は `top`。`top`・`approver`・`admin`・`approve2`・`staffSummary` 以外は `top`（`adminApi.gs:535-545`） |
| `page=redirect` | 既定: `302 /?page=top`（旧フォームの入口。Q9 の既定。フォームは閉じる＝9-1 #11） |
| `admin`・`approve2` | `permissions.is_admin(db, email)` が偽なら `no_auth`（題名「権限エラー」）。文は `総務部（管理者）権限がありません。\nApproverMap シートに role=admin で登録されているか、\nSettings シートの ADMIN_EMAILS にメールアドレスが含まれているか確認してください。\n（現在のアカウント: <email>）`。**デバッグ情報（`--- デバッグ情報 ---` 以下）は出さない**（Q8 の既定。H13） |
| `approver` | 管理者か、`permissions.is_any_dept_approver(db, email)` が真なら表示。偽なら `no_auth`、文は `承認者権限がありません。\nDeptApprovers または ApproverMap に登録されているか確認してください。\n（現在のアカウント: <email>）`。メールは必ずあるので「メールが空なら素通り」（H3）は起きない |
| `top`・`staffSummary` | 判定なし（誰でも） |
| 題名 | 5 ページとも `<title>残業・休日出勤申請</title>`、権限エラーは `<title>権限エラー</title>`（GAS の `setTitle`。staffSummary の `<title>スタッフビュー画面 …</title>` は「残業・休日出勤申請」に置き換える＝Q15 の既定。gap-2 S10） |
| 応答ヘッダー | `Cache-Control: no-store`。CSP などは nginx（3-7） |
| ほかのパス | `/?requestId=` は使わない（このアプリに詳細画面は無い） |

### 3-3. テンプレート変数

GAS の `<?= X ?>`（自動エスケープ）は `{{ X }}`、`<?!= X ?>` は 3-6 の JSON の塊へ移す。

| 変数 | 値 | 使い方 |
|---|---|---|
| `APP_URL` | `PUBLIC_ORIGIN + "/"`（例 `https://overtime.lineworks-local.net/`）。リンクは `{{ APP_URL }}?page=top` のまま | 全ページのヘッダー、no_auth の「トップ画面へ戻る」 |
| `PORTAL_URL` | `.env` の `PORTAL_URL` | 「社内ポータルへ移動」 |
| `DATE_PARAM` | `?date=` が `^\d{4}-\d{2}-\d{2}$` に合い、日付として正しいときだけその文字。それ以外は `''`（9-1 #13・[02 §6] の高） | approve2 の `urlDate`、admin の `URL_DATE_PARAM` |
| `REASONS_JSON`・`OT_HOURS_JSON`・`HD_HOURS_JSON` | `src/config.gs:42-52` の定数を写した `constants.py` の値 | top |
| `message` | 3-2 の文。改行 `\n` を含むが、`.msg` に `white-space` を足さない（今と同じく 1 行につながって見える。gap-2 S10） | no_auth |

### 3-4. `google.script.run` の代わり（`web/static/js/gas-shim.js`、U）

各ページの呼び出し行（40 か所）は 1 字も変えずに動くこと。

- 形: `google.script.run.withSuccessHandler(f).withFailureHandler(g).<fn>(...args)`。`withFailureHandler` が先、`withSuccessHandler` だけ、どちらも無し、の書き方も動く。呼ぶたびに新しい runner を返す（GAS と同じく、ハンドラの設定は 1 回の呼び出しにだけ効く）。
- 送り方: `fetch('/api/rpc/' + fn, {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(args), credentials:'same-origin'})`。`args` は引数をそのまま並べた配列。**省略された引数は送らない**（`api_markOvertimeDone(id)` は `["id"]`。gap-2 S6）。
- 返事:
  - 200 `{"ok":true,"result":R}` → `f(R)`。`R` が `null`・`{error}`・`{ok:false,…}`・`results[]` でも、そのまま成功ハンドラに渡す（成功の中で失敗を返す 5 つの形を変えない。gap-2 S6）。
  - 400 `{"ok":false,"error":msg}` → `g(err)`。`err` は `new Error(msg)`（`err.message` と `String(err)` の両方で文が出る。`String(err)` は `"Error: msg"` になるので、`toString` を `msg` を返すように上書きする）。
  - 409（同じ操作を処理中）→ 1 秒おきに最大 10 回、同じ冪等キーで送り直す。尽きたら `g(Error(サーバーの文))`。
  - 401 → 3-5。
  - 5xx・通信エラー → `g(Error('サーバーでエラーが起きました。時間をおいてもう一度お試しください。'))`（rpc の `ERR_INTERNAL` と同じ文）。冪等キーは残す。
- 冪等キー（9-1 #5・罠 17）: 書き込みの関数（4-1 の「書く」）にだけ `X-Idempotency-Key` を付ける。キーは「関数名＋引数の JSON」ごとに 1 つ作り、200・400 で捨てる。通信エラー・5xx・409 の後に同じ関数・同じ引数で呼ばれたら同じキーを使う。**同じ関数・同じ引数の呼び出しが返事待ちの間に来たら、送らずに前の呼び出しの返事を両方のハンドラに渡す**（二度押しで 2 件できない）。
- 呼び出しの記録: `console.log` はページの今の出力（`JSON.stringify(res)`）だけ。シムは何も出さない。

### 3-5. 操作中のログイン切れ（罠 20・gap-2 S3）

- 401 を受けたら、失敗ハンドラは呼ばない（今の画面は失敗で alert と読み直しをするため、二重に出る）。
- 書きかけを `sessionStorage` に退避する（キー `overtime-draft:<page>`、24 時間、同じメールのときだけ戻す）。対象: 申請モーダル（種別・部署・作業員・STEP・日付・過去日の選択・工番 3 件・理由・補足理由・予定時間・業務内容）、編集モーダル、総務の手動実績入力、承認者画面で管理者が選んだ承認者、staffSummary の月・部署・スタッフ。ログインから戻ったら、ページの初期化の後で戻す。
- `beforeunload` があれば外してから、`alert('ログインが切れました。ログイン後にもう一度押してください。')` を出し、`location.href = '/auth/login?next=' + encodeURIComponent(location.pathname + location.search)`。
- 楽観更新した行（top 6 か所・approver 3 か所・approve2 1 か所。gap-2 S3）は、移動してログインから戻るとページが読み直されるので、サーバーの今の状態で描き直される。書き込みは自動で送り直さない。冪等キーは捨てない。
- 先に確かめる: 申請モーダルを開くとき、`document.visibilitychange` で画面に戻ったとき、`GET /auth/me` を呼び、401 ならその場で上と同じ手順でログインへ送る（入力を始める前に切れていることに気づかせる）。

### 3-6. CSP と JS の置き場所（案 1 に決める）

計画書 4-3 の案 1 を採る（`script-src 'self'`。インラインの script と on 属性を禁止）。

- 各ページの `<script>` の中身を `web/static/js/<page>.js` に移す（中身は書き換えない。関数はグローバルのまま）。読み込みは `<script src="/static/js/gas-shim.js"></script><script src="/static/js/<page>.js"></script>` をページの今の script の位置に置く。
- `<?!= ?>` と JS の中の `<?= ?>` は、ページに `<script type="application/json" id="gas-template-data">{{ data_json }}</script>` を置いて渡す（実行されないので CSP に触れない）。`data_json` は `json.dumps` の後に `<`・`>`・`&` を `\u003c`・`\u003e`・`\u0026` に置き換えた文字。`<page>.js` の先頭で `var __T = JSON.parse(document.getElementById('gas-template-data').textContent);` とし、`var REASONS = __T.REASONS;`・`var URL_DATE_PARAM = __T.DATE_PARAM;`・`var urlDate = __T.DATE_PARAM;` のように今の変数名で受ける。
- on 属性（HTML に 81 個: onclick・oninput・onchange・onmouseenter・onmouseleave）と、JS の文字列で作る onclick（`top.html:1187-1243・1344・1361-1366`、`approver.html:734-737`、`approve2.html:321・349`、`admin.html:618-620`）は、`data-on-<event>="関数名"` と `data-args='<JSON 配列>'` に機械的に置き換え、`gas-shim.js` の 1 つの受け口（`document` への委譲）で呼ぶ。引数の中の `this` は `"$this"`、`event` は `"$event"`、`this.value` は `"$value"` と書き、受け口が置き換える。`mouseenter`・`mouseleave` は委譲できないので、受け口は `mouseover`／`mouseout` で `relatedTarget` を見て同じ動きにする。
- `approver.html:876` の `document.querySelector('[onclick*="approveAllInTable(\'…\')"]')` は、ボタンに `id="btnApproveAll-<種別>"` を付けて id で探す形に直す（見た目は同じ。gap-2 S5）。
- style 属性と `<style>` は書き換えない（`top.html:304` の `div[style*="grid-template-columns:1fr 1fr"]` が style の文字に頼っている）。

### 3-7. nginx（F）

- `location /static/` だけ nginx が `web/static` から返す（`expires -1`）。`/`・`/api/`・`/auth/`・`= /healthz` は api へ中継。`/api/sebastian` も同じ（レート制限は `/api/` と同じ）。
- セキュリティヘッダーは server レベルに 1 回だけ。CSP: `default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'self'; base-uri 'self'; form-action 'self'`。`frame-ancestors` は Q10 の既定（ポータルが iframe で埋め込んでいれば 1 行変える）。HSTS・nosniff・Referrer-Policy・Permissions-Policy は貯蔵品と同じ。
- 利用者の IP は `CF-Connecting-IP`。`limit_req` は貯蔵品と同じ値。

### 3-8. XSS の直し（必須。gap-2 S8・[02 §6]）

- 4 ページの `escHtml` と staffSummary の `esc` に `'` → `&#39;` を足す（表示は同じ）。
- `admin.html:618-620` の手動入力ボタンは 3-6 の `data-args`（JSON）にする。値は `escHtml` してから属性に入れる。
- `top.html:1150-1151` の PDF リンクの `pdfFileId` も `escHtml` を通す。
- `DATE_PARAM` は 3-3 の検証と 3-6 の JSON の塊でしか JS に入らない。
- 試験: 氏名が `x');alert(1)//` と `"><img src=x onerror=alert(1)>` の申請を作り、総務の日次表・承認者画面・TOP・二次承認・スタッフ集計で何も実行されず、文字のまま見える（12 章 U-6）。

### 3-9. スマホ（9-1 #25・罠 16）

- 今の HTML の `<meta name="viewport" content="width=device-width, initial-scale=1">` がそのまま効くので、各ページの今の `@media`（640px・768px・480px・820px・520px）が初めて動く。これをスマホ向けの配置の土台にする。
- 足す CSS は `web/static/css/mobile.css` の 1 ファイルに集め、各ページの最後の `<link>` で読む。**すべての規則を `@media (max-width: 820px)` 以下の中に書く**（PC の見た目を変えないため。`test_web_config.py` が確かめる）。
- 必ず入れる規則: `input[type=date], input[type=time], input[type=month] { -webkit-appearance:none; appearance:none; display:block; min-width:0; max-width:100%; text-align:left; box-sizing:border-box }` と `::-webkit-date-and-time-value { text-align:left }`（対象: `top.html:485・596・626`、`approve2.html:176`、`admin.html:421・425`、`approver.html:292`）。横にはみ出す表（admin・approve2・approver・staffSummary）は表を包む要素に `overflow-x:auto`。押せる大きさ（ボタン・リンク 40px 以上）。モーダルは画面幅に収める。
- Safari の日付の読み取り: 画面の JS で `new Date('yyyy-MM-dd HH:mm:ss')`（空白入り）のような Safari で `Invalid Date` になる書き方をしない。今のコードは `T` と `+09:00` を足す形と正規表現で読む形なので、そのまま保つ。U は 5 ページの `new Date(` を全部見て、文字を渡している箇所の一覧を報告に書く。

### 3-10. 運用の欄（管理者だけ。U が admin.html の末尾に足す）

総務画面（admin は管理者しか開けない）の一番下に、折りたたみの「運用（管理者）」欄を足す。今ある部分の見た目は変えない。中身は `ops_getStatus()` の結果（4-2）と次のボタン: ［共有タブを今すぐ取り込む］［マスタ同期］［失敗した仕事の再実行］（行ごと）［PDF の作り直し（申請 ID）］［指定日の PDF 一括］［台帳の作り直し］［休憩の再計算］［委任の許可をやり直す］（`/auth/google-delegate/start` へのリンク）［セッションの取り消し（メール）］。押す前に confirm を出す。

---

## 4. RPC

### 4-1. 取り決め（F が作り、W が関数を登録する）

- `POST /api/rpc/{fn}`、本文は引数の JSON 配列。Origin が `PUBLIC_ORIGIN`（無ければ `Sec-Fetch-Site: same-origin`）、`Content-Type: application/json`、`fn` が閉じたリストにある、本文が配列、を入口で確かめる。
- 応答: 200 `{"ok":true,"result":…}`／業務エラー 400 `{"ok":false,"error":"GAS と同じ文"}`／401 `{"ok":false,"error":"ログインが必要です。","login":"/auth/login"}`／403（管理者だけの関数を管理者以外が呼んだ）`{"ok":false,"error":"総務部（管理者）権限がありません。"}`／409（冪等キーが処理中）／501（未登録）／500（詳細はログだけ）。
- 業務関数の形: `def api_x(ctx: RpcContext, *args)`。引数は GAS と同じ並びの位置引数で、省略できる引数には既定値を持たせる。`ctx.user.email`（小文字）・`ctx.db`・`ctx.now`（1 回の呼び出しで同じ時刻、UTC）・`ctx.settings`。**業務関数は commit しない**（rpc が冪等キーと一緒に確定する）。業務エラーは `raise RpcError("GAS と同じ文")`。
- 書き込みの関数（下の表の「書く」）は `audit.record(db, user_email, fn, request_ids, ok)` を同じトランザクションで残す（F が rpc の中で呼ぶ）。
- 返す JSON の作り（GAS の直列化の癖を写す）:
  - GAS で `undefined` だったキーは出さない（例: `api_editRequest` の `newTargetDate` は日付を変えないとき出さない）。
  - 日時・日付は 4-3 の表の形の文字（JST、`T` も時差も付けない）。`null` を返すのは表に `null` と書いた所だけ。
  - 数は int。分の空はその関数が GAS で返していた値（多くは `Number('') = 0`）。
  - キーの並び順は気にしない（画面は名前で読む）。
- 閉じたリスト `RPC_FUNCTIONS` = 4-3 の 29 関数＋4-2 の 10 関数（計 39。`ops_runSheetIntake` は決定 #26 で足した）。`api_getFormUrl`・`api_getTodayRequestsForDept`・`api_getTodayRequestsForWorker` と、画面から呼ばれない保守の関数（`recalcWorkLogsBreakMinutes`・`rebuildAccumulationSS`・`syncAllMasters` ほか）は出さない（9-1 #11・#13）。保守は 4-2 の管理者だけの関数で行う。

### 4-2. 管理者だけの関数（W が `domain/ops.py` に作る。すべて `permissions.is_admin` が真のときだけ）

| 関数（引数） | 返す形 | 中身 |
|---|---|---|
| `ops_getStatus()` | `ops_status.get_ops_status(db)` の dict（M） | 行の反映の待ち件数・最も古い待ちの経過秒・止まっている行、GAS の仕事の待ち・実行中・失敗の件数と失敗の一覧（id・申請 ID・action・理由・回数）、最後の `gas_ping` の時刻と結果、朝バッチ・マスタ同期・共有タブ取り込み・毎晩の照合の最後の結果、照合で直した行の数 |
| `ops_retryJob(jobId)` | `{ok:true}` | `failed` の仕事を `pending` に戻す（書く） |
| `ops_syncSheetConfigNow()` | `config_sync.sync_sheet_config(db)` の結果 | 共有タブの取り込み（書く） |
| `ops_runMasterSync()` | `master_sync.sync_masters(db)` の結果 | マスタ同期（書く） |
| `ops_regeneratePdf(requestId)` | `{ok:true, jobId}` | `gas_jobs.enqueue_regenerate(db, requestId, reason="admin")`。今の `forceRegeneratePdf` の WorkLogs の書き換えと列の追加は写さない（gap-3 G3-L3） |
| `ops_batchPdfForDate(ymd)` | `{ok:true, jobId}` | `morning_batch.enqueue_manual_batch(db, ymd)`。今のメニュー「PDF一括生成」と同じく台帳は書かず、BatchLogs に `manualBatchPdfs` の 1 行 |
| `ops_rebuildLedger()` | `{ok:true, jobId}` | `gas_jobs.enqueue_ledger_rebuild(db)`（GAS の `rebuildAccumulationSS` を内部用の入口で呼ぶ） |
| `ops_recalcBreaks()` | `{ok:true, changed:N}` | `recalcWorkLogsBreakMinutes`（`worklog.gs:712-806`）を写す。`end <= start` の行は飛ばす。変わった行だけ `breakMinutes`・`netMinutes`・`updatedAt`（文字）・`updatedBy='break-recalc'` |
| `ops_revokeSessions(email)` | `{ok:true, revoked:N}` | `auth.revoke_sessions` |
| `ops_runSheetIntake()` | `sheet_intake.run_now(db)` の結果（`{ok, active, ingested, skipped, …}`） | シートの取り込み（6-10、決定 #26）を今すぐ 1 回。シートは読むだけ。DB に申請を入れる（書く） |

### 4-3. 画面の 29 関数（すべて W。シムと画面は U）

「書く」は 4-1 の冪等キーと監査の対象。「積む」は 7 章の GAS の仕事（M の `gas_jobs.enqueue_*` を呼ぶ）と 6 章の行の反映（`sheet_queue.mark_dirty`）。エラー文は [01 §1.4] の表の文字が正で、1 字も変えない（下では主なものだけ書く）。日時の形は「JST の文字」。シートの値が文字の行は、その文字をそのまま返す（`approval.gs:193-194`）。

| # | 関数（引数） | 返す形 | 権限 | 書くもの・積む仕事 | 二重押し・状態の照合 |
|---|---|---|---|---|---|
| 1 | `api_getDeptList()` | 部署名の配列（作業員マスタの部署の重複なし、JS の既定の並び＝UTF-16 の符号順） | なし | 読む | — |
| 2 | `api_getWorkersByDept(dept)` | `["<コード> <氏名>", …]`（コード順）。`dept` が空なら `[]` | なし | 読む | — |
| 3 | `api_getWorkerInfo()` | `{workerCode, workerName, dept, email}` か `null`。作業員マスタの「Googleアカウント」で始まる列と**大文字小文字を区別して完全一致**。今のマスタにその列が無いので常に `null`（gap-4 F-12） | なし | 読む | — |
| 4 | `api_getOrderList()` | `[{code, product, dest}]`（工番の重複は最初の行、code 昇順） | なし | 読む | — |
| 5 | `api_getHolidayCandidateDates()` | `[{date:'yyyy-MM-dd', label, type:'weekend'｜'holiday', holidayName?}]`。今週月曜〜来週金曜（+11 日）かつ今日（JST）以降。社内カレンダーの A〜D 列を位置で、曜日は C 列の文字そのまま、`出勤土曜` を除く。0 件・読めないときは土日＋計算の祝日（`formMap.gs:310-378` の `getKnownJapaneseHolidays_` を写す）。Calendar API は使わない（gap-5 中-6） | なし | 読む | — |
| 6 | `api_getDashboard()` | `{today:'yyyy-MM-dd', overtime:[item], holiday:[item]}`。item は `worklog.gs:460-480` のキー（`requestId, requestType, status, dept, workerCode, workerName, targetDate, targetDateLabel, approvedMinutes, actualStartAt, actualEndAt, …`、`actualStartAt/EndAt` は `yyyy-MM-dd HH:mm:ss` か `''`）。例外のとき `{today:'', overtime:[], holiday:[], error}` を 200 で | なし | 読む。表示条件は [01 §2.7] の表 | — |
| 7 | `api_submitRequest(data)` | 成功 `{ok:true, success:true, requestId}`。入力の誤りは `{ok:false, error}` を **200 で**。作業員コードと予定時間の誤りは 400（例外） | なし（誰の名義でも出せる。H4 は今のまま） | 書く: Requests 1 行（新しい行）＋WorkLogs の下書き行（`updatedAt` は日時のセル、`updatedBy='formSubmit'`）。工番マスタに無い工番は `exportError` に `工番マスタ未登録: X`（` / ` 区切り）。積む: 行の反映だけ | 冪等キー。同じ人・同じ日の申請は止めない（今と同じ） |
| 8 | `api_markOvertimeDone(id, manualEndTime?)` | `{ok:true, requestId, actualMinutes, breakMinutes, netMinutes, pdf:null}`（`pdf` は GAS の仕事で後から作るので常に `null`。画面は使っていない） | `assertSelf_`（`workerEmail` が空でなく、ログインのメールと**大文字小文字を区別して**違えば `本人のみ操作可能です。`） | 書く: WorkLogs（行が無ければ作る）。積む: `internal.afterComplete`（`source:"screen"`、`wlOverride`、`statusAtEnqueue`） | 再実行は上書き（Q9）。`requests` の行を先にロック（gap-1 G3） |
| 9 | `api_markHolidayStart(id, manualStartTime?)` | `{ok:true, requestId, actualStartAt:'yyyy-MM-dd HH:mm:ss'}` | `assertSelf_` | 書く: WorkLogs の `actualStartAt`・`updatedAt`・`updatedBy`。仕事は積まない（台帳も PDF も触らない） | 何度でも上書き。完了後も可（gap-1 G14） |
| 10 | `api_markHolidayDone(id, manualEndTime?)` | `{ok:true, requestId, actualMinutes, breakMinutes, netMinutes, pdf:null}` | `assertSelf_` | 書く: WorkLogs（開始も JST の文字で書き直す）。積む: `internal.afterComplete`（`source:"screen"`） | 開始が無いと `休日開始が記録されていません（開始ボタンを押してください）。` |
| 11 | `api_editRequest(id, patch)` | `{ok:true, newStatus, newMinutes}`、休日で日付を変えたときだけ `newTargetDate:'yyyy-MM-dd'` | なし | 書く: `approvedMinutes`、休日の `targetDate`、`modifiedAt`（`yyyy-MM-dd HH:mm:ss` の文字として扱う）・`modifiedBy`。PDF があれば DB の PDF の 3 列を空にし、積む: `internal.regenerate`（`reason:"edit"`、`oldPdfFileId`）。PDF が無ければ何も積まない（Q8） | 取消済み・完了済みは拒む。残業は `OT_HOURS` と完全一致 |
| 12 | `api_deleteRequest(id)` | `{ok:true}` | なし | 書く: status=`canceled`、DB の PDF の 3 列を空、WorkLogs の行があれば 5 列を空・`updatedAt`（文字）・`updatedBy`=メール。積む: `internal.afterDelete`（`oldPdfFileId`・`requestType`・`targetDate`） | 2 回目は `既にキャンセル済みです。` |
| 13 | `api_getApproverDepts()` | 管理者 `{isAdmin:true, approvers:[{email, name, depts:[…]}], currentEmail}`、承認者 `{isAdmin:false, approvers:[自分], currentEmail}`（`deptAuth.gs:62-82` の形） | 承認者でなければ 400 `承認者権限がありません。DeptApprovers または ApproverMap に登録してください。` | 読む | — |
| 14 | `api_getApproverDashboard(deptsOrDept)` | `{today, overtime, holiday, isAdmin, selectedDepts}`。item の `submittedAt` は `yyyy-MM-dd HH:mm:ss`、`targetDateLabel` は `fmtDateLabelJa_`。部署なし `{today:'', overtime:[], holiday:[], isAdmin, error:'部署が指定されていません。'}`、権限なし `error:'部署「X」の承認権限がありません。'`、例外のとき `{…, isAdmin:true, error}`（H2 の癖は今のまま） | 部署ごとに `can_approve_dept` | 読む。canceled の当日分も出す（Q4） | — |
| 15 | `api_approveRequest(id)` | `{ok:true, requestId, approvedBy, approvedAt:'yyyy-MM-dd HH:mm:ss'}`。承認済みなら `{ok:true, message:'既に承認済みです。'}` | `can_approve_dept(email, 申請の部署)` | 書く: status=`approved`、`approvedAt`（日時）、`approvedBy`。WorkLogs に `actualEndAt` があれば積む: `internal.afterApprove` | 2 つのタブで押しても 2 回目は `既に承認済みです。`、仕事は 1 回だけ |
| 16 | `api_rejectRequest(id)` | `{ok:true, requestId}` | 同上 | 書く: status=`canceled`、`approvedBy`=`<email>(却下)`、`approvedAt`。台帳・PDF は触らない | submitted 以外は `未承認状態でのみ却下可能です（現在: X）。` |
| 17 | `api_approveRequestsBatch(ids)` | `{ok:true, results:[{requestId, ok, message?｜error?, approvedBy?, approvedAt?}]}`（`approval.gs:388-475` の形と順） | 1 件ずつ `can_approve_dept` | 15 と同じを 1 件ずつ。行ロックは ID を並べ替えて取る。結果の並びは Requests の行の順、その後に見つからなかった ID（JS のオブジェクトのキーの順＝数字だけのキーが先）。配列でない引数は包まない（文字列は 1 字ずつの ID、数・オブジェクトは対象なし＝GAS と同じ） | 50 件を超えると 400 `一括承認は50件までです。` |
| 18 | `api_approverMonthlySummary(ym, deptsOrDept)` | `buildMonthlySummary_` の形（[01 §5.1]）。`charts.dept` は作り直さない（H14） | 全部署で `can_approve_dept` | 読む | — |
| 19 | `api_getSecondaryApprovalList(ymd)` | `{date, items:[…]}`（`approval.gs:651-668` のキー。`submittedAt`・`approvedAt`・`approvedAt2` は `yyyy-MM-dd HH:mm`、`isRetroactive`・`daysAgo` 付き。並びは同じ） | 管理者 | 読む（GAS の見出しの自動追加は写さない＝Q17） | — |
| 20 | `api_secondaryApprove(id)` | `{ok:true, requestId, approvedBy2, approvedAt2:'yyyy-MM-dd HH:mm:ss', pdfRegenerated:false}`（PDF は後で作るので常に `false`。画面は使っていない）。済みなら `{ok:true, message:'既に二次承認済みです。'}` | 管理者 | 書く: `approvedBy2`・`approvedAt2`（日時）。PDF があれば 3 列を空にして積む: `internal.regenerate`（`reason:"secondary"`）。無ければ積む: `internal.ledgerUpsert`（単票はいつも台帳を書く） | 一次承認前は `一次承認が完了していません（現在: X）。` |
| 21 | `api_secondaryApproveBatch(ids)` | `{ok:true, results:[…]}` | 管理者（件数の確かめが先、権限が後＝`approval.gs:776-779`） | 新しく二次承認したものだけ 20 と同じに積む。6 件以上なら GAS の仕事の間を 500ms 空ける | 50 件超は 400 |
| 22 | `api_adminDeptOptions()` | `['ALL', …部署]` | 管理者 | 読む | — |
| 23 | `api_adminMonthlySummary(ym, dept)` | [01 §5.1] | 管理者 | 読む | — |
| 24 | `api_adminFiscalTrend(ymd, dept)` | `{fiscalYearStart, fiscalYearEnd, deptFilter, labels, values}`（4/1 始まり） | 管理者 | 読む | — |
| 25 | `api_adminDailyDetail(ymd, dept)` | `{date, deptFilter, total, pdfMissing, items}`。**直す（9-1 #22）**: item の `targetDate`（`yyyy-MM-dd`）・`pdfGeneratedAt`・`actualStartAt`・`actualEndAt`（`yyyy-MM-dd HH:mm:ss`、空は `''`）を文字で返し、画面の過去申請のバッジのために `isRetroactive`・`daysAgo` も足す。並びは `(dept+workerName)` の `localeCompare(…,'ja')`（ICU の照合で写す。gap-1 G11） | 管理者 | 読む | — |
| 26 | `api_adminManualComplete(id, end, start?)` | `{ok:true, actualMinutes, breakMinutes, netMinutes}` | 管理者 | 書く: WorkLogs（`updatedBy`=メール）。積む: `internal.afterComplete`（`source:"admin"`、`wlOverride`） | `end <= start` は `終了時刻は開始時刻より後にしてください。`（エラー文は [01 §1.4]） |
| 27 | `api_adminSpecialClauseCount(ymd, dept)` | `{fiscalYear:'YYYY-04-01 ~ YYYY+1-03-31', list:[…]}` | 管理者 | 読む | — |
| 28 | `api_adminPendingWatch()` | 配列（[01 §5.4] のキー） | 管理者 | 読む | — |
| 29 | `api_getClosingSummary(ym, dept, workerToken)` | [01 §5.5] の `personal`／`list` の形（データが無くても `null` は返さない。`{mode:'list', people:[]}` など） | なし | 読む | — |

「読む」の関数も 1 回の呼び出しの中で同じ時刻・同じスナップショットを使う。`api_getDashboard` などのサーバーの「今日」は JST（8-1）。

### 4-4. Sebastian の入口 `POST /api/sebastian`（W が中身、F が経路）

- 経路は F が `main.py` に置き、`domain.sebastian.handle(db, raw_body: bytes) -> dict` を呼んで、返った dict を **いつも HTTP 200** の `application/json` で返す。ログイン・Origin・冪等キー・Content-Type の確かめは通さない。本文は Content-Type に関係なく生の文字を `json.loads`（Pydantic の本文モデルを使わない。gap-1 G6）。
- 返事の文字と形は [01 §1.2] の表と同じ（`api.gs` の文字どおり。末尾に `: <requestId>` が付く文、`{error, action}` で `ok` が無い形、JSON の誤り・トークン未設定・不一致の 3 つは `action` も無い）。`invalid JSON: ` の後ろは Python の例外文になる（受け入れる差）。トークンは `hmac.compare_digest`。
- action:
  - `ping` → `{ok:true, pong:"<UTC の ISO 8601、ミリ秒と Z 付き>"}`
  - `completeOvertime`（`manualEndTime` は無視）: 開始＝対象日 17:20 JST、終了＝今、**休憩は時刻窓方式**（9-1 #23。画面と同じ `breaks.calc_break_by_clock_time`）、`updatedBy='sebastian-agent'`。積む: `internal.afterComplete`（`source:"sebastian"`、優先度高）。
  - `markHolidayStart`（`manualStartTime` 任意）・`markHolidayDone`（`manualEndTime` 任意）: 画面の 9・10 と同じ計算、`updatedBy='sebastian-agent'`。
  - 本人確認（`assertSelf_`）はしない（今と同じ）。行ロックは画面と同じ。
- 返す `status` は書く前に読んだ値。`workerName` は正規化しない生の値。`targetDate` は**シートの型どおり**: その行の `targetDate` が日付のセルなら「その日の 0 時（`DB_SHEET_TIMEZONE`）」を UTC にした ISO 文字（`2026-11-24T15:00:00.000Z` の形）、文字ならその文字。
- `pdf`（`completeOvertime`・`markHolidayDone`）: status が approved でなければ `null`。approved なら、積んだ仕事の結果を `gas_jobs.wait_for_result(db_factory, job_id, SEBASTIAN_PDF_WAIT_SECONDS)` で待ち、`{ok:true, pdfFileId, pdfName, folderId}`・`{already:true, pdfFileId}`・`{error:"<GAS の文>"}` のどれかで返す。待ち切れたら `{error:"PDF作成の結果待ちが時間切れになりました（作成は続いています）"}` を入れて返し、仕事は続ける。
- GAS の転送役の取り決め（G が作る）: 移転モードの GAS の `doPost` は、本文が `"internal": true` を持たなければ、受けた本文（生の文字）をそのまま `SEBASTIAN_FORWARD_URL`（スクリプトのプロパティ。値は `https://overtime.lineworks-local.net/api/sebastian`）へ `UrlFetchApp.fetch(url, {method:'post', contentType:'application/json', payload: 本文, muteHttpExceptions:true, followRedirects:false, headers:{'X-Overtime-Forwarded':'gas'}})` で送り、返事が 200 で JSON として読めればその文字をそのまま返す。200 以外は `{"error":"upstream error: <code>"}`、例外は `{"error":"upstream unavailable: <message>"}`。プロパティが無く移転モードなら `{"error":"service is moving, retry later"}`。どれも GAS の `ContentService` の JSON（HTTP 200）。

---

## 5. DB の表（F が `models.py` と `0001_initial_schema` に全部作る）

共通: 文字の列は `NOT NULL DEFAULT ''`（空とシートの空を同じに扱う）。日時は `UTCDateTime`（timestamptz、時差のない値は拒む）。enum は `CheckConstraint`。業務の表はすべて `row_version int NOT NULL DEFAULT 1`（書くたびに +1）、`sheet_row int NULL`（シートの行番号。まだ書いていない新しい行は NULL）、`cell_kinds jsonb NOT NULL DEFAULT '{}'`（型が混ざる列の、行ごとの型。6-3）。**シートの値は正規化せずに生のまま入れる**（GAS が `normalize_` するところは業務の側で `jsutil.normalize_gas()` を通す）。

### 5-1. 業務の表（A 区分。行単位でシートへ反映する）

`requests`（シート `Requests`、見出し 1 行目。列は見出しの文字で対応させる。見出し → 列）:

| シートの見出し | 列 | 型 |
|---|---|---|
| `requestId` | `request_id` | text 主キー（UUID 36 字） |
| `requestType(overtime/holiday)` | `request_type` | text（`overtime`／`holiday`） |
| `status(submitted/approved/canceled)` | `status` | text（`submitted`／`approved`／`canceled` だけ。移行で `modified` があれば止める） |
| `dept`・`workerCode`・`workerName`・`workerEmail` | `dept`・`worker_code`・`worker_name`・`worker_email` | text |
| `targetDate` | `target_date` date NULL ＋ `target_date_text` text（日付にならない元の文字を残す。gap-3 G3-L2） | |
| `submittedAt`・`approvedAt`・`approvedAt2`・`pdfGeneratedAt`・`modifiedAt` | `submitted_at`・`approved_at`・`approved_at2`・`pdf_generated_at`・`modified_at` | timestamptz NULL |
| `approvedBy`・`approvedBy2`・`modifiedBy` | `approved_by`・`approved_by2`・`modified_by` | text |
| `approvedMinutes` | `approved_minutes` | int NULL（申請時の予定。承認で変わらない） |
| `reason`・`reasonDetail`・`workContent` | `reason`・`reason_detail`・`work_content` | text |
| `jobId1`〜`jobId3`・`hrMailSentAt`・`exportError` | `job_id1`〜`job_id3`・`hr_mail_sent_at`・`export_error` | text（いつも空の列も持つ） |
| `workNo1`〜`3`・`orderNo1`〜`3`・`customer1`〜`3`・`product1`〜`3` | `work_no1`… | text（`orderNo` は工番の写し） |
| `pdfFileId`・`pdfFolderId` | `pdf_file_id`・`pdf_folder_id` | text |
| （追加）| `old_pdf_file_id_pending` | text。PDF の 3 列を空にした操作の、まだゴミ箱へ移していない前の ID（7 章の仕事が使う） |

- シートの見出しに無い列（自動で足された列が今の本番に無い場合）は、移行の dry-run で `missingColumns` に出して止める。シートにあって決まりに無い見出しは `droppedColumns` に出す（触らない）。
- `isRetroactive` は列を作らない。`submitted_at` の JST の日付 > `target_date` で毎回求める（gap-5 中-7）。

`work_logs`（シート `WorkLogs`、1 行目は説明、見出しは 2 行目、データは 3 行目から）:

| 見出し | 列 | 型 |
|---|---|---|
| — | `id` | bigint 主キー（WorkLogs は Requests と 1 対 1 ではない。孤児 32 行・行の無い申請 17 件。外部キーを張らない） |
| `requestId` | `request_id` | text、索引（一意ではない。同じ ID が複数あれば `sheet_row` の小さい行を使う＝GAS の `findWorkLogRowNo_` と同じ） |
| `actualStartAt`・`actualEndAt`・`updatedAt` | `actual_start_at`・`actual_end_at`・`updated_at` | text（GAS が読んで返す JST の `yyyy-MM-dd HH:mm:ss`。日時のセルだった行も GAS と同じ読み方で文字にして持ち、`cell_kinds` に元の型を残す） |
| `actualMinutes`・`breakMinutes`・`netMinutes` | `actual_minutes`・`break_minutes`・`net_minutes` | int NULL（空は NULL。0 と区別する） |
| `updatedBy` | `updated_by` | text（決まった値 `formSubmit`・`sebastian-agent`・`unknown`・`delete`・`admin-manual`・`break-recalc` は同じ文字） |

`batch_logs`（シート `BatchLogs`、見出し 1 行目）: `id` bigint 主キー、`executed_at`（`実行日時`、timestamptz）、`batch_name`（`バッチ名`）、`target_date`（`対象日`、date）＋`target_date_text`、`success`・`skip`・`fail`（`成功`・`スキップ`・`失敗`、int）、`error_detail`（`エラー詳細`、text。複数は `\n` でつなぐ）。

### 5-2. シートを正とする表（B 区分。VPS は書かない。取り込むだけ）

| 表 | 元 | 持ち方 |
|---|---|---|
| `settings` | `Settings`（1 行目説明・2 行目見出し・3 行目から A=key、B=value） | `key`（前後の空白を取る）主キー、`value`、`sheet_row`。同じキーが 2 行以上あれば**下の行が勝つ**。重なったキーは `sync_runs` の詳細に出す（gap-5 中-2・gap-4 F-14） |
| `dept_approvers` | `DeptApprovers`（見出し 1 行目） | `sheet_row` 主キー、`dept`（A）、`emails`（B の生の文字）、`name`（C）、`col_d`・`col_e`（D `役職`・E `通知メール`。持つだけ。貯蔵品との約束＝gap-4 F-5） |
| `approver_map_rows` | `ApproverMap`（1 行目説明・2 行目見出し、今は空） | `sheet_row` 主キー、`cells` jsonb（1 行目から生のまま）。`permissions` は GAS と同じく **1 行目を見出しとして**読む（今は効かない＝今と同じ） |

`StampMap`・`休憩マスタ` は取り込まない（VPS で使う所が無い。StampMap は GAS の内部用の入口が直接読む。休憩マスタは 9-1 #23 で使わなくなった）。

### 5-3. マスタの表（C 区分。マスタ元から作り、DB のタブにも書く）

| 表 | 元のタブ（マスタ元 → DB） | 列 |
|---|---|---|
| `workers` | `作業員マスタ` → `作業員マスタ` | `sort_order` 主キー、`worker_code`、`worker_name`、`dept`、`active_flag`、`google_account`（「Googleアカウント」で始まる列。無ければ空）、`raw` jsonb（見出し付きの 1 行） |
| `orders` | `工番マスタ` → `工番マスタ` | `sort_order` 主キー、`code`・`customer`（受注先）・`dest`（納入先）・`dest_address`・`product`・`quantity`・`imported_at`（text） |
| `company_calendar` | `社内カレンダーマスタ` → 同名 | `sort_order` 主キー、`date_text`（A）、`kind`（B）、`weekday`（C、文字のまま）、`note`（D） |
| （表なし） | `業務NO.マスタ` → `業務NOマスタ` | タブへ書くだけ |

### 5-4. 土台の表

| 表 | 列（主なもの） |
|---|---|
| `sessions` | `token_hash`（SHA-256）主キー、`email`、`name`、`created_at`、`expires_at`（`min(now+IDLE, created_at+MAX_AGE)`、90 日） |
| `google_credentials` | 貯蔵品と同じ（暗号化したリフレッシュトークン、許可されたスコープ） |
| `idempotency_keys` | 主キー（`user_email`, `key`）、`fn`、`state`（`running`／`done`）、`response_json`、`created_at`。24 時間で消す |
| `jobs` | `id`、`kind`（7 章の action 名か `morning_batch` など）、`lane`（`gas`／`default`）、`priority`（0 通常・10 Sebastian）、`request_id`、`args_json`、`requires_json`（`[{table, pk, version}]`。この版まで行の反映が済むまで動かさない）、`guard_json`（積んだ時点の status など）、`state`（`pending`／`running`／`done`／`failed`／`skipped`）、`attempts`、`run_after`、`leased_until`、`last_error`、`result_json`、`batch_id`（一括の操作で積んだ仕事のまとまり）、`created_at`・`updated_at` |
| `sheet_sync_state` | 主キー（`table_name`, `pk`）、`desired_version`、`synced_version`、`synced_values` jsonb（最後にシートへ書いた、またはシートにあると分かっている値。列ごと）、`state`（`ok`／`retry`／`blocked`）、`attempts`、`next_try_at`、`last_error`、`first_dirty_at`、`last_synced_at`、`seq`（bigserial。新しい行を作った順） |
| `sync_runs` | `id`、`kind`（`sync_masters`・`sync_sheet_config`・`sheet_reconcile`・`morning_batch`・`gas_ping`・`purge`・予定枠 `schedule:<名前>`）、`started_at`、`finished_at`、`ok`、`summary_json`（件数・直した行・差の一覧） |
| `audit_log` | `id`、`at`、`user_email`、`fn`、`request_ids`（text 配列を JSON で）、`ok`、`error` |

alembic は `0001_initial_schema` 1 本。`test_migrations.py`（SQLite で upgrade → `compare_metadata == []` → downgrade）を通す。

### 5-5. 行ロック

- 申請 1 件ごとに `requests` の行を `SELECT … FOR UPDATE`（GAS の全体ロックの代わり。別の申請は待たない）。
- **`work_logs` だけを書く処理（完了・開始・総務の手動入力・Sebastian の 3 action・削除）も、先に `requests` の同じ行をロックする**（gap-1 G3。承認と完了が同時でも、どちらかが必ず PDF の仕事を積む）。申請が無い ID は GAS と同じエラー。
- 一括の処理は ID を並べ替えてからロックする。
- SQLite のテストでは `FOR UPDATE` を無視してよい（`store.lock_request` の 1 か所で切り替える）。

---

## 6. シートへの行単位の反映（決定 A。D が作る）

### 6-1. どのタブを書くか

| 区分 | タブ | VPS の扱い |
|---|---|---|
| A（VPS が行単位で書く） | `Requests`・`WorkLogs`・`BatchLogs` | 書くたびに、その行の変わった列だけを同じ位置へ。新しい行は末尾。行もタブも消さない |
| C（マスタ同期だけが書く） | `作業員マスタ`・`工番マスタ`・`社内カレンダーマスタ`・`業務NOマスタ` | 6-8 |
| B（読むだけ） | `Settings`・`DeptApprovers`（5 列）・`ApproverMap`・`StampMap`・`休憩マスタ` | 書かない。`config_sync` が数分ごとに読む |
| D（GAS が使う） | `申請書フォーム`・`操作`・`休憩計算_フォーム` | 触らない（読みもしない） |
| E（移さない） | `FormMap`・`FormTemplates`・`ExportHistory`・`部署マスタ` ほか | 触らない |

### 6-2. 流れ

1. W の業務関数は、行を書いたら同じトランザクションで `sheet_queue.mark_dirty(db, table_name, pk)` を呼ぶ（`row_version` を +1 し、`sheet_sync_state.desired_version` をそれに合わせる。新しい行は `sheet_sync_state` を作る）。
2. worker の行の反映のレーン（全体で 1 本。2 秒ごと）が、`desired_version > synced_version` かつ `next_try_at <= now` の行を、新しい行は `seq` 順、ほかは `first_dirty_at` 順に最大 50 件取る。
3. DB の**今の値**を読み（積んだときの値ではない）、`synced_values` と比べて違う列だけを書く。
4. 書けたら `synced_version = 読んだ時点の row_version`、`synced_values` を更新。GAS の仕事は `requires_json` の版を見て待つ（7 章）。

### 6-3. 書き方（型と表示形式）

- Sheets API の `spreadsheets.batchUpdate` の `updateCells` で、値と表示形式を 1 回で書く（`fields: "userEnteredValue,userEnteredFormat.numberFormat"`）。値は `numberValue`（シリアル値・分）か `stringValue`（文字。入力として解釈されない＝RAW と同じ。`'` を付けない）。空は `userEnteredValue` を持たないセル（値を消す）。
- 列ごとの型の決まり（`services/sheet_codec.py` の `COLUMN_RULES`。初めの値は下の表、スパイクで本番の型の表を取って直す）:

| 種類 | 列 | 書き方 |
|---|---|---|
| 日時のセル | Requests `submittedAt`・`approvedAt`・`approvedAt2`・`pdfGeneratedAt`、BatchLogs `実行日時`、WorkLogs の下書き行の `updatedAt` | `DB_SHEET_TIMEZONE` の壁時計でシリアル値を作り、`numberValue`。表示形式は 6-4 |
| 型が混ざる列 | Requests `targetDate`・`modifiedAt`、BatchLogs `対象日`、WorkLogs `updatedAt` | 行の `cell_kinds[列]` が `serial` なら日付・日時のセル、`string` なら文字。新しい行では `auto`＝その列の最後のデータ行と同じ型と表示形式 |
| JST の文字 | WorkLogs `actualStartAt`・`actualEndAt`・完了などの `updatedAt` | `stringValue`、表示形式 `@`（TEXT。GAS の `updateWorkLog_` と同じ） |
| 分・回数 | `approvedMinutes`・WorkLogs の 3 列・BatchLogs の 3 列 | 整数の `numberValue`。NULL は空 |
| 文字 | そのほか | `stringValue`。自由入力の `1/2` も文字のまま（GAS の `setValues` は日付に変えていた。受け入れる差＝11 章） |

- W が型を決める所: WorkLogs の下書き行（申請のとき）は `cell_kinds.updatedAt='serial'`、完了・開始・削除・総務の手動入力・Sebastian・休憩の再計算で書くときは `'string'`（GAS の書き方と同じ）。ほかは書かない（D が `auto` を解く）。
- シリアル値の作り方: Sheets のシリアル値は 1899-12-30 からの日数で、時刻は小数。timestamptz を `DB_SHEET_TIMEZONE` の壁時計に直してから数える（ミリ秒まで）。日付は 0 時。

### 6-4. 行の位置・並び・表示形式

- 今ある行（移行した行と、一度書いた行）は `sheet_row` に書く。書く前にその行の ID の列（Requests・WorkLogs は `requestId`、BatchLogs は `実行日時` と `バッチ名`）を `batchGet` で読み、DB と同じか確かめる。違えば ID の列を全部読んで位置を探し直し、`sheet_row` を直してから書く。見つからなければ書かずに `state='blocked'`、管理者の欄に出す。
- 新しい行は、ID の列を読んで「最後のデータ行の次」に書く（GAS の `appendRow` と同じ）。シートの行数が足りなければ `appendDimension`（ROWS、末尾に足すだけ）。書いた位置を `sheet_row` に入れる。
- ただし Requests の新しい行は、先に requestId だけを `values.append`（範囲 `'Requests'!A:A`、RAW、`insertDataOption=OVERWRITE`）で書き、返事の `updatedRange` の行を `sheet_row` にして、残りの列をその行へ `updateCells` で書く（`row_sync.reserve_new_rows`。決定 #26）。Sebastian が同じタブへ `INSERT_ROWS` で追記するのと同時でも、行を決めるのはシートの側なので、互いの行を上書きしない。間で落ちた行（requestId だけの行）は、次の回に ID の列から見つけて使う。OVERWRITE は行を挿入しない（シートの末尾なら行が足される）。
- 新しい行の日付・日時のセルの表示形式は、同じ列の最後のデータ行の `numberFormat` を写す（そのセルが文字なら、日付は `yyyy/MM/dd`、日時は `yyyy/MM/dd H:mm:ss`）。表示形式の無いシリアル値は、戻した GAS が数値として読み、未承認滞留 48h などが狂う（gap-4 F-1）。
- 行・列・タブの挿入・削除・並べ替え・作り直しは使わない。申請の削除は値の書き換え（status=`canceled`、WorkLogs の 5 列を空）で表す。
- 見出しの確かめ: 書く前に見出しを読み（WorkLogs は 2 行目。1 行目の説明に触れない）、`COLUMN_RULES` の見出しが 1 つでも無ければレーン全体を止めて知らせる（列を足さない。gap-1 G15）。決まりに無い右側の列には触れない。
- 時間帯の確かめ: レーンは起動時と 1 時間ごとに `spreadsheets.get` の `properties.timeZone` を読み、`DB_SHEET_TIMEZONE` と違えばレーンを止めて知らせる（gap-5 高-2）。

### 6-5. GAS が書く列との取り合い

- GAS の内部用の入口は、Requests の `pdfGeneratedAt`・`pdfFileId`・`pdfFolderId` をシートに直接書く（今の関数のまま）。
- M は GAS の返事を受けたら、`gas_jobs.record_pdf_result` で DB の 3 列に入れ、**同じトランザクションで `sheet_sync_state.synced_values` の 3 列も同じ値に直し、`desired_version`・`synced_version` を同じだけ進める**（シートに既にある値なので、レーンが書き直さない）。
- 二次承認・編集・削除で 3 列を空にしたときは、レーンが空を書く（GAS の `clearExistingPdfForRegeneration_` と同じ結果）。GAS の仕事はその版の反映を待ってから動く。
- 行の反映は列ごとの差分なので、GAS が仕事の途中に別の列が変わっても、3 列を古い値で上書きしない。

### 6-6. 目標と失敗のとき

- 遅れ（`first_dirty_at` から書くまで）は 1 分以内、ふだんは 10 秒以内。5 分を超えた行があれば `ops_getStatus` に「遅れ」と出し、`/healthz` の詳細（管理者だけ）にも出す。
- 1 行の失敗は 10 秒・30 秒・1 分・2 分・5 分、以後 5 分ごとに再試行（ほかの行は先へ進む）。429・5xx は `SheetsApi` が指数的に 4 回まで再試行する。
- 安全装置の拒否・見出しの違い・時間帯の違いは、再試行せずにレーンを止める（`sync_runs` に記録、管理者の欄に赤で出す）。

### 6-7. 毎晩の照合（`sheet_reconcile`、2 時台）

- A 区分の 3 タブを `UNFORMATTED_VALUE`（`SERIAL_NUMBER`）と `FORMATTED_VALUE` で全行読み、DB と 1 セルずつ比べる。日時は**秒の単位**で比べ（GAS が書いた時刻はミリ秒を持つため）、表示形式の種類（日付か文字か）も見る。
- `desired_version > synced_version` の行（反映待ち）は比べない。
- 差があれば DB の値でその行を書き直し、直した行・列・前後の値を `sync_runs.summary_json` と管理者の欄に残す。シートにあって DB に無い行（GAS や人が書いた行、Sebastian が足してまだ取り込んでいない行）は消さず・書かずに報告する（`extraRows`・`notInDb`）。
- 歯止め（罠 49）: DB の `requests` か `work_logs` が 0 行、またはシートより DB の行が少ないときは、1 行も書かずに止めて知らせる。Requests の「DB に無い requestId の行」（取り込み待ち。6-10）は `FOREIGN_TOLERANCE`（20）行までは数から除く（超えたら止める）。
- `python -m app.worker run sheet_reconcile --check` は書かずに数えるだけ（切り替えの B-16 と戻しの C-2 で使う）。

### 6-8. マスタ同期（決定 D。`master_sync.sync_masters`、0:30 ごろと手動）

- 読む: `MASTER_SOURCE_SPREADSHEET_ID` の `作業員マスタ`・`業務NO.マスタ`・`工番マスタ`・`社内カレンダーマスタ`。読み取りの再試行は GAS（`syncMasters.gs`）と同じ回数。
- 規則（`syncMasters.gs` を写す）: 1 列目が空の行を捨てた後の「見出し＋先頭 80 行」で作業員を切る（元の並び）。Date は「JST の 0 時ちょうど → `yyyy-MM-dd`、それ以外 → `yyyy-MM-dd HH:mm:ss`」。UNFORMATTED で読み、マスタ元の時間帯で日時に直してから同じ規則で文字にする。書式だけの列は無視する。
- 書く: DB の 4 タブへ 1 行目から、GAS の `setValues` と同じ USER_ENTERED で書き、余った行は値だけ消す（`values:batchClear`、行は消さない）。タブを消して作り直さない。
- 守り: 元が読めない・見出しだけ・有効な行 0 なら、そのタブは上書きしない（今の GAS は見出しだけに置き換える＝11 章の変えるもの）。タブごとに失敗を握り、ほかのタブは続ける。
- 同じ読み取りから `workers`・`orders`・`company_calendar` を全置換する。

### 6-9. 共有タブの取り込み（`config_sync.sync_sheet_config`、5 分ごとと手動）

- `Settings`・`DeptApprovers`・`ApproverMap` を読んで 5-2 の表を全置換する。読めない・空・有効な行 0 なら上書きしない。
- キーと見出しは前後の空白を取り、`normalize_gas`（前後の空白を取り、全角英数を半角に）を値にも当てて比べる（Settings はキーの前後の空白だけ。gap-4 F-11・F-14）。
- Settings の `SOURCE_SSID`・`TEMPLATE_SSID`・`PDF_ROOT_FOLDER_ID`・`ACCUMULATION_SS_ID` が `.env` の値と違えば、`sync_runs` と管理者の欄に「食い違い」を出す（書き込み先は `.env` が決める）。

### 6-10. シートの取り込み（`sheet_intake`、60 秒ごとと手動。決定 #26。担当 D2）

Sebastian（同じ VPS の `/root/.hermes` の AI。こちらからは変えない）は、本番 DB スプレッドシートの Requests へ Sheets API で新しい申請を直接足す: `values().append(range='Requests!A:A', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS')`。行は見出しの名前から作り、`requestId`（uuid4）・`requestType`・`status='submitted'`・`dept`・`workerCode`・`workerName`・`workerEmail`（部署の共有アカウント）・`targetDate`（`yyyy-MM-dd`）・`submittedAt`（JST の `M/D/YYYY H:MM:SS`。USER_ENTERED なので日時のセルになる）・`approvedMinutes`・`reason`・`workContent`・`reasonDetail`・`workNo1..3`・`orderNo1..3`・`customer1..3`・`product1..3` を入れ、ほかは空。Requests を全列読んで requestId を探し、完了は GAS の `doPost`（切り替えの後は VPS の `/api/sebastian` へ転送）で送る。GAS の時代は何もしなくてもその行が全画面・集計に出た。VPS では DB が正なので、取り込まないと見えない。

- **いつ動くか**: worker の予定 `sheet_intake`（`every_minutes=1`）、管理者の `ops_runSheetIntake()`、`/api/sebastian` の完了・開始が DB に無い requestId を指したとき（その場で 1 回取り込んでから処理する。GAS では追記の直後でも見つかったため）。
- **関所**: production は `write_gate.check` が通る（新しい `gas_ping` が `migrated=true`）ときだけ。切り替えの前（GAS が正）と戻した後は取り込まない。staging はステージングの写しを読むだけなので、いつでも動く。
- **鍵**: 同時に 1 つだけ（プロセスの中は `threading.Lock`、PostgreSQL は `pg_try_advisory_lock(739310002)`）。取れなければ `{busy:true}` で何もしない（手での実行と Sebastian の入口は少し待つ）。入れる前に DB にあるかを確かめ、主キーの重なりは「もうある」として飛ばす（冪等）。
- **読む**: 時間帯を確かめ（`row_sync.check_timezone`）、`row_sync.load_tab` で見出しと requestId の列と行の数を読む（安い読み取り）。DB に無い requestId の行と、どの申請の行でもない requestId の空の行だけを、移行の道具と同じく UNFORMATTED と FORMATTED で読む（罠 11）。
- **読み方**: `sheet_codec.decode_row`（移行の道具 `tools/migrate_from_sheets.py` と同じ関数）。入れる行は `sheet_row`＝その行、`row_version=1`、`cell_kinds`＝行ごとの型。`sheet_sync_state` は反映済み（`desired_version=synced_version=1`、`synced_values`＝`sheet_codec.synced_values_for`＝今シートにある値）なので、行の反映は書き直さない。後で承認などをすると、行の反映が requestId で行を確かめ、同じ行の変わった列だけを書く（ずれていれば探し直す）。
- **入れない行**（入れずに記録し、その行へは書かない）: requestId が空・同じ requestId が 2 行以上・読めない値（`CodecError`）・`requestType` と `status` が決まりにない・`targetDate` が日付でない・`submittedAt`・`dept`・`workerCode` が空・requestId が 64 字を超える。staging では `pdfFileId`・`pdfFolderId` のある行も入れない。毎回読み直すので、人がシートを直せば次の回に入る。
- **WorkLogs**: 作らない。GAS は Sebastian が足した行に何もしなかった（`ensureWorkLogRow_` は `api_submitRequest` と旧フォームの取り込みだけ）。実績を記録するとき `store.update_worklog` が行を作る（`updateWorkLog_` と同じ）。
- **記録**: 予定の回は `sync_runs`（`schedule:sheet_intake`。3 日より古いものは取り込みが消す）、手での回は `sheet_intake`、取り込んだ申請は `sheet_intake_ingest`（消さない）。`ops_getStatus` の `sheetIntake`（最後の回の `ingested`・`skipped`・`stopped`）と `lastRuns.sheet_intake` に出る。
- **シートは書かない**。DB に無い行（取り込み待ち・入れない行）は、行の反映・照合（6-7）・戻しの道具のどれも上書き・移動・削除しない。戻しの道具は Requests に控えに無い requestId の行があれば `--apply` を止める（`rowsNotInCopy`。`--drop-rows-not-in-copy` を付けたときだけ戻す）。控えの道具（`cutover_backup`）は複製を作るだけ。
- **残る差**: Sebastian の追記から画面に出るまで最大 1 分ほど（予定の間隔）。Sebastian の `INSERT_ROWS` は、A 列の表の途中に空の行があるとそこへ行を挿入して下の行をずらす（GAS の時代も同じ）。ずれた行は行の反映が requestId で探し直す。照合は位置の違う行を `idMismatch` に出して書かない。

---

## 7. GAS の内部用の入口と、きっかけごとの呼び順（決定 B）

### 7-1. 取り決め（G が `src/Internal.gs` に作り、F の `gas_client` が呼ぶ）

| 項目 | 決め |
|---|---|
| 入口 | 今の `doPost`（`api.gs:10`）の先頭で本文を読み、`body.internal === true` なら `handleInternal_(body)` へ。Sebastian の 4 action とは別の道 |
| URL | 同じプロジェクトの内部用のデプロイ（画面用・Sebastian 用と分ける）の `/exec`。デプロイはディレクターがオーナーの許可を得て作る。G はコードだけ |
| 認証 | 本文の `token` をスクリプトのプロパティ `VPS_INTERNAL_TOKEN` と長さ一定の比較で比べる。プロパティが無い・違うときは何もせず `{"ok":false,"error":"unauthorized","retryable":false}` |
| 本文 | `{"internal":true,"token":"…","action":"internal.<名前>","jobId":"<uuid>","requestId":"…","expect":{…},"args":{…}}` |
| 返事 | いつも HTTP 200 の JSON。成功 `{"ok":true,"action","jobId","requestId","pdf":{…}?,"ledger":{…}?,"trash":{…}?,"durationMs"}`。失敗 `{"ok":false,"error":"<コード>","message":"…","retryable":true｜false,"pdf":{…}?}`（失敗でも PDF ができていれば `pdf` を入れる。VPS は先に記録する） |
| `pdf` | `{"status":"created｜already｜skipped｜error","pdfFileId","pdfName","folderId","pdfGeneratedAt":"yyyy-MM-dd HH:mm:ss","pdfGeneratedAtEpochMs":<数>,"error"}`。`created` と `already` のときは、書いた後のシートの 3 列を読み直して返す |
| `ledger` | `{"status":"upserted｜removed｜rewritten｜rebuilt｜skipped｜error","sheetName","error"}` |
| `trash` | `{"status":"trashed｜skipped｜error","fileId","error"}` |
| エラーのコード | `unauthorized`・`bad_request`（false）・`row_not_synced`（`expect` と合わない。true）・`busy`（ロック待ち切れ。true）・`ledger_unavailable`（台帳が開けない。true）・`ledger_header_mismatch`（false）・`internal`（そのほかの例外。true） |
| 冪等 | `jobId` を鍵に、返事を `CacheService.getScriptCache()` に 6 時間（`ij:<jobId>`）。同じ `jobId` が来たら処理せずに同じ返事に `"replayed":true` を足して返す。キャッシュの例外は握って処理を続ける（罠 36）。6 時間を過ぎた再送でも、PDF は `already`、台帳は requestId の上書きなので二重にならない |
| `expect` | `{"status","approvedBy2","actualEndAt","targetDate","approvedMinutes","pdfFileId"}` のうち VPS が送ったもの。GAS は `getRequestFullData_` と `buildWorkLogsMapByRequestId_` でシートを読み、`normalize_` した文字で比べる（日時は JST の `yyyy-MM-dd HH:mm:ss`、`targetDate` は `yyyy-MM-dd`、数は文字にして）。違えば何もせず `row_not_synced` |
| 実行の印 | token を確かめた実行の中だけ、グローバルの `__OVERTIME_INTERNAL_OK__ = true` を立てる。移転モードの `getDb_()` の止めはこの印があれば通す |
| ロック | 外側にロックを足さない（今の `generatePdf*_` の中の `waitLock(20000)` をそのまま使う）。ロックの待ち切れは `busy` |
| 時間 | 1 回の呼び出しで 1 件だけ。GAS の上限 6 分。VPS の待ちは `GAS_CALL_TIMEOUT_SECONDS`（330 秒）。時間切れは「結果が分からない」として同じ `jobId` で送り直す |
| 台帳の前の確かめ | 台帳を書く action は、関数を呼ぶ前に Settings の `ACCUMULATION_SS_ID`（同じキーは下の行）を `SpreadsheetApp.openById` で開けるか確かめ、開けなければ書かずに `ledger_unavailable`（今の `getAccumulationSS_` の「新しく作る」を動かさない）。書く年度タブ（`^(残業|休日出勤)_\d{4}年度$`、`computeAccFiscalYear_` と種別で決める）が既にあれば 1 行目が `ACC_HEADERS_`（27 列）と 1 字ずつ同じか確かめ、違えば `ledger_header_mismatch`。タブが無ければ通す（年度の始めに今の関数が作る）。`_backup_` 付きのタブには触れない |

### 7-2. action（中身は今の関数を、今と同じ順序・条件で呼ぶ）

| action | `args` | GAS がすること（この順） |
|---|---|---|
| `internal.ping` | — | 何も書かない。`{ok:true, scriptId: ScriptApp.getScriptId(), spreadsheetId: getDb_().getId(), spreadsheetTimeZone, scriptTimeZone, migrated: !!getMigratedToUrl_(), forwardUrlSet, settings:{TEMPLATE_SSID, PDF_ROOT_FOLDER_ID, ACCUMULATION_SS_ID, PDF_MODE, SOURCE_SSID}, accumulationKeyRows:<件数>, apiVersion:1}` |
| `internal.afterApprove` | — | `approval.gs:328-341` と同じ: WorkLogs に `actualEndAt` があるときだけ、`try { PDF_MODE が direct を含めば generatePdfDirect_(id) でなければ generatePdfForRequest_(id); upsertAccumulationRow_(id) } catch { pdf=error、台帳は書かない }`。`actualEndAt` が無ければ両方 `skipped` |
| `internal.afterComplete` | `source`（`screen`／`admin`／`sebastian`）、`statusAtEnqueue`、`wlOverride`（screen・admin のときだけ。`{actualStartAt, actualEndAt, actualMinutes, breakMinutes, netMinutes}`） | `statusAtEnqueue === 'approved'` なら `try { sebastian: generatePdfForRequest_(id)（いつも XLOOKUP 方式、override なし）／それ以外: PDF_MODE で generatePdfDirect_(id, wlOverride) か generatePdfForRequest_(id, wlOverride) } catch { pdf=error }`。**PDF の成否に関係なく** `upsertAccumulationRow_(id)`（`worklog.gs:193-207`・`api.gs:102-113`・`adminApi.gs:498-512`） |
| `internal.regenerate` | `reason`（`secondary`／`edit`／`admin`）、`oldPdfFileId` | まず `oldPdfFileId` があれば `DriveApp.getFileById(…).setTrashed(true)`（失敗は `trash.status=error` に記録して続ける＝gap-5 中-3）。secondary: `_headerCache = {}; try { generatePdfDirect_(id) }`（いつも直接方式、二次承認印入り）、その後いつも `upsertAccumulationRow_(id)`。edit: `try { PDF_MODE で generatePdf* }`、その後 `try { upsertAccumulationRow_(id) }`。admin: PDF_MODE で作るだけ、台帳は書かない |
| `internal.ledgerUpsert` | — | `upsertAccumulationRow_(id)`（PDF の無い二次承認の単票） |
| `internal.afterDelete` | `oldPdfFileId`、`requestType`、`targetDate`（`yyyy-MM-dd`） | 前の PDF をゴミ箱へ（失敗は記録して続ける）→ `try { removeAccumulationRow_(id, requestType, targetDate) }`。`removeAccumulationRow_` が文字の日付で年度を正しく決めるかを G が試験で確かめる |
| `internal.batchPdf` | `targetDate` | シートの `pdfFileId` があれば `skipped`。無ければ PDF_MODE で `generatePdfDirect_(id)` か `generatePdfForRequest_(id)`。例外は `pdf.status=error, error=e.message`（`ok:true` で返す。朝バッチが数える）。台帳は書かない |
| `internal.ledgerRewriteDay` | `targetDate`（`yyyy-MM-dd`） | `appendToAccumulationSS_(new Date(y, m-1, d))`（前日分を消して書き直す今の癖のまま） |
| `internal.ledgerRebuild` | — | `rebuildAccumulationSS()`。5 分を超えそうなら今の関数のまま試し、時間切れは `internal` で返す（管理者が分かるように） |

- どの action も、呼ぶ前に `getRequestFullData_(id)` で status を読み、VPS が送った `expect` と照合する。PDF を作る関数自身も status を確かめる（今のまま）。
- 印鑑・テンプレ・フォルダ・ファイル名・`操作!B3`・`netMinutes` の補正・一時コピーは、今の関数がするとおり（変えない）。一時コピーを失敗時も片付ける `finally` と、テンプレの書かないセル（Q19）は、スパイクの後でオーナーが決めるまで入れない。

### 7-3. VPS 側の仕事（M。`services/gas_jobs.py`）

- 仕事は `jobs` の `lane='gas'`。worker の GAS のレーンは**全体で同時に 1 本**（gap-1 G2・gap-3 G3-M1）。`priority` の高い順、同じなら `id` 順。**同じ申請の仕事は積んだ順で、前の仕事が `done`／`skipped` になるまで次は動かさない**。前の仕事が `failed` のままなら、同じ申請の次の仕事も止めて管理者の欄に出す（台帳を古い値で上書きしないため）。別の申請は先へ進める。
- 動かす前に:
  1. `requires_json` の各行で `sheet_sync_state.synced_version >= version` を確かめる。まだなら `run_after = now + 5 秒`（試行回数は増やさない）。
  2. DB の申請を読み直す。`afterApprove`・`afterComplete`・`regenerate`・`ledgerUpsert`・`batchPdf` で status が `canceled` なら `skipped`（削除の後に古い仕事で PDF や台帳を作らない。削除の仕事が台帳を消す）。
  3. `expect` を作る（DB の今の値。シートは反映済みのはず）。
  4. 最後の `gas_ping` の照合（9-3）が 10 分以内に成功していなければ動かさない（`run_after` を 1 分後に）。
- 結果の記録（同じトランザクション）: `pdf.status` が `created`／`already` なら `record_pdf_result`（6-5）。`trash` が済めば `old_pdf_file_id_pending` を空に。`jobs.result_json` に返事を入れる。
- 再試行: `retryable:true`・通信エラー・時間切れは 1・2・4・8・16・30 分、その後は 30 分ごとに 24 時間まで（Q20 の既定「再試行する」）。`retryable:false` と 24 時間を過ぎたものは `failed`。利用者の操作は止めない。
- 一括の操作で 6 件以上積んだまとまり（`batch_id`）は、同じまとまりの呼び出しの間を 500ms 空ける（GAS の `Utilities.sleep(500)` の写し）。

### 7-4. きっかけごとの呼び順（今と同じ順序・条件。計画書 4-6 の表）

| きっかけ | VPS の操作のトランザクション（W） | 積む GAS の仕事（M の関数） |
|---|---|---|
| 一次承認（個別・一括は 1 件ずつ） | status・approvedAt・approvedBy。WorkLogs の `actualEndAt` を同じトランザクションで読む | `actualEndAt` があれば `enqueue_after_approve`。requires: Requests の行 |
| 残業・休日の完了（画面） | WorkLogs | `enqueue_after_complete(source="screen", wl_override, status_at_enqueue)`。requires: Requests・WorkLogs |
| 総務の手動入力 | WorkLogs | 同上 `source="admin"` |
| Sebastian の完了（残業・休日） | WorkLogs（`updatedBy='sebastian-agent'`） | 同上 `source="sebastian"`、`priority=10` |
| 休日の開始（画面・Sebastian） | WorkLogs | なし |
| 二次承認（単票） | approvedBy2・approvedAt2。PDF があれば 3 列を空にして `old_pdf_file_id_pending` | PDF があれば `enqueue_regenerate(reason="secondary", old_pdf_file_id)`、無ければ `enqueue_ledger_upsert` |
| 二次承認（一括） | 新しく二次承認したものだけ上と同じ | 同上（新しく二次承認したものだけ） |
| 編集 | approvedMinutes・targetDate・modifiedAt・modifiedBy。PDF があれば 3 列を空に | PDF があったときだけ `enqueue_regenerate(reason="edit")` |
| 削除 | canceled・3 列を空・WorkLogs の 5 列を空 | `enqueue_after_delete`（PDF の有無に関係なく。台帳の行を消す） |
| 却下・申請 | — | なし |
| 朝バッチ | 7-5 | 7-5 |
| 管理者の作り直し | 3 列を空にして `old_pdf_file_id_pending` | `enqueue_regenerate(reason="admin")` |

### 7-5. 朝バッチ（`morning_batch.run_morning_batch`、毎朝 6:50 JST。gap-3 G3-H1）

GAS の `morningBatchNoMail_`（`mail.gs:1454-1468`）→ `batchGeneratePdfs_`（`pdfExport.gs:570-600`）→ `logBatchResult_` → `appendToAccumulationSS_` の順を写す。仕事 `morning_batch`（`lane='gas'`）1 本の中で、次を順に行う（outbox に積んで先へ進まない）。

1. 対象日＝JST の昨日。DB から「対象日が昨日で approved」の申請を、GAS の `listApprovedRequestsByDate_` と同じ並びで拾う（`items`）。
2. 1 件ずつ: `pdfFileId` があれば `skip+1`。無ければ `internal.batchPdf` を呼び、`created`／`already` なら `ok+1` と記録、`error` なら `fail+1` と `errors` に `<requestId> (<dept>/<workerName>): <message>`。通信の失敗が 3 回続いたら（30 秒おき）その件は `fail`、`message` は `GAS呼び出し失敗: <内容>`、あとで送り直す仕事（`internal.batchPdf`、Q20）を積む。`items` が 6 件以上なら 1 件ごとに 500ms 待つ。**1 件ごとに `priority` の高い仕事（Sebastian）があれば先に流す。**
3. `batch_logs` に 1 行（`実行日時`=今、`バッチ名`=`morningBatchNoMail`、`対象日`=`yyyy-MM-dd`、成功・スキップ・失敗、エラー詳細は `\n` でつなぐ）を書き、`mark_dirty`。
4. `internal.ledgerRewriteDay(昨日)` を呼ぶ（台帳を開けなければ `failed` にして管理者へ）。
5. 7:05 までに Requests と BatchLogs の行の反映が済んだことを確かめ、済んでいなければ管理者の欄に出す（統合メール通知システムが 7:10 ごろ読む見込み。gap-5 中-8）。
6. 未完了でも前日の approved なら PDF を作る癖は今のまま（Q18 の既定）。2 で送り直した件が後で成功したら、その日の `internal.ledgerRewriteDay` をもう一度積む。

`ops_batchPdfForDate(ymd)` は 1〜3 を同じに行い（対象日は指定日、`バッチ名`=`manualBatchPdfs`）、4 は行わない。

### 7-6. GAS 側の変更（G。決まった所だけ）

| ファイル | 変更 |
|---|---|
| `src/Internal.gs`（新） | 7-1・7-2 のすべて。`INTERNAL_API_VERSION = 1` |
| `src/Cutover.gs`（新） | 工番部品の `src/Cutover.gs` を写して直す。`MIGRATION_TRIGGER_FUNCTIONS = ['morningBatchNoMail_','syncAllMasters','nightlyUpdateAllForms_','pollNewResponses_']`。`renderMigratedPage_(url, page, date)` は `page`（5 ページの名前だけ）と `date`（`^\d{4}-\d{2}-\d{2}$` だけ）を引き継いだ `https://overtime.lineworks-local.net/?page=…&date=…` へのボタン。`<base target="_top">` とボタンの `target="_top"` は必須、1 秒後の `window.top.location` は「動けばよい」（罠 29・gap-2 S4）。運用関数: `enterMigrationModeForOvertime()`（URL は引数の要る `enterMigrationMode(url)` を包む。既定の URL を持たない＝罠 24d）、`setSebastianForwardUrl()`、`clearSebastianForwardUrl()`、`setInternalToken()`（エディタで実行すると乱数を作って保存し、値は画面に出さず長さだけ出す。VPS へは手で写す）、`closeAllFormsForMigration()`（FormMap の 21 本の受付を閉じて ID を記録。フォームと回答は消さない）、`removeAllTriggersForMigration()`、`exitMigrationMode()`（転送のプロパティも消す）、`restoreTriggersAfterRollback()`（`morningBatchNoMail_`・`syncAllMasters` だけ。同じ関数のトリガーがあれば作らない）、`showMigrationStatus()`（読むだけ。版・移転モード・転送先・トリガー・`POLL_LAST_TS` 以降の未取り込みの回答数。`getDb_()` の止めを通らない）。運用関数は `Session.getActiveUser().getEmail() === Session.getEffectiveUser().getEmail()`（持ち主）のときだけ通す（罠 48） |
| `src/config.gs` の `getDb_()`（69 行） | 先頭に `assertNotMigrated_()`。ただし `__OVERTIME_INTERNAL_OK__` が立っていれば通す（罠 47） |
| `src/adminApi.gs` の `doGet`（534 行） | 先頭で移転モードなら `renderMigratedPage_(url, e.parameter.page, e.parameter.date)` を返す |
| `src/api.gs` の `doPost`（10 行） | 先頭で本文を読み、`internal === true` なら `handleInternal_`。それ以外で転送先があれば 4-4 の転送。移転モードで転送先が無ければ `{"error":"service is moving, retry later"}`。どれでもなければ今の処理 |
| トリガーの 4 関数 | 先頭で `if (skipBecauseMigrated_('<名前>')) return;` |
| `getDb_()` を通らない入口（gap-4 F-7） | `setupAllTriggers_`・`setup*Trigger_`・`cleanupFormSubmitTriggers_`・`reprocessAllResponses_`・`debugProcessLatest_`・`debugPollStatus_` の先頭に `assertNotMigrated_()`。`api_getHolidayCandidateDates` の `CalendarApp` への逃げ道の先頭も同じ |
| `src/menu.gs` の `onOpen` | 移転モードなら「移転済み（新しいアドレス）」の 1 項目だけ |

- ほかの関数（PDF・台帳・業務）は 1 行も変えない。`appsscript.json` も変えない（`webapp` の設定は既にある。罠 45）。
- 移転モードでないとき（`MIGRATED_TO_URL` が無い）は、内部用の入口が増えた以外、今とまったく同じ動きであること（G の試験で確かめる）。

---

## 8. 同じにしなければならない業務の決まり

### 8-1. 時刻（決定 C。必須）

- サーバーの「今」は `timeutil.now_utc()` だけ。「今日」は `timeutil.today_jst()`。`date.today()`・`datetime.now()`・`datetime.utcnow()` を使わない（F の試験が `app/` を grep して拒む）。各モジュールは `from app import timeutil` として `timeutil.now_utc()` を呼ぶ（試験で `monkeypatch.setattr(timeutil, "now_utc", …)` が効くように）。
- 年度・締め月・月の経過日数・30 日の窓・日時の解釈は、すべて Asia/Tokyo の日付で行う。DB に入れる日時は UTC。JSON は JST の文字。
- 年度は 2 通り（Q16。今のまま）:
  - **3/16 始まり**（`fiscal.fy_start_0316(d)`: `m > 3 or (m == 3 and day >= 16)` ならその年の 3/16、でなければ前年の 3/16）: 申請の日付の検査（残業の年度内・休日の年度開始）、スタッフ集計の年度、台帳の年度タブ（GAS が決める）。
  - **4/1 始まり**（`fiscal.fy_start_0401(d)`）: 総務の年度推移・特別条項。
- 締め月（`fiscal.closing_ym(d)`）: 日が 15 以下なら前月分、16 以上なら当月分。期間は M/16〜(M+1)/16 未満。
- 残業の通常の申請は、送られた日付を捨てて JST の今日にする（`formSubmit.gs:710-712`）。
- ブラウザ側の「今日」は変えない（[02 §5]）。

### 8-2. 実績と休憩（今のまま。#23 だけ変える）

- 残業の開始は対象日 17:20 JST 固定。終了は手入力 `HH:mm`（`^\d{2}:\d{2}$` に合えば対象日のその時刻。合わなければ黙って今の時刻＝gap-1 G13）か今。時・分が範囲外（`25:00` など）は、GAS では無効な日付で例外になる見込み（推測）なので、400 `時刻が不正です: <値>` を返す（受け入れる差）。
- 休日の開始は手入力か今。完了は開始（WorkLogs）から。終了が開始より前でも 0 分（エラーにしない）。
- `actualMinutes = max(0, js_round((end - start) をミリ秒にして / 60000))`、`netMinutes = max(0, actual - break)`。
- 休憩は**時刻窓方式**（`breaks.calc_break_by_clock_time(start, end)`。`worklog.gs:111-129` を写す）: 開始・終了を JST の `HH:mm` にし（秒は捨てる）、分で比べる。窓 10:00-10:10、12:15-13:00、15:00-15:10、17:10-17:20、19:20-19:30、21:30-21:40。`開始 <= 窓開始 かつ 終了 >= 窓終了` の窓だけ足す。日をまたぐと終了の分が小さくなり 0。
- **Sebastian の残業完了も同じ時刻窓方式**（9-1 #23。切り替えの日から。過去の行は直さない）。休憩マスタは使わない。
- 総務の手動入力: 開始は引数、無ければ残業 `17:20`、休日は WorkLogs の `actualStartAt` の最初の `HH:mm`。開始・終了とも対象日の時刻。`end <= start` はエラー。
- 対象日（Requests の `targetDate`）が空・読めない申請で、残業完了・休日の開始と完了・総務の手動入力を押したとき、VPS は 400 `作業実施日が無効です: <値>` を返す（GAS は `Utilities.formatDate` の英語の例外文になる見込み。受け入れる差。正しさレビュー 中-1）。
- 手入力の時刻・日付の形の確かめは文字列の末尾まで（`\Z`）。末尾に改行がある値は形に合わないので、GAS と同じく黙って今の時刻にする。

### 8-2a. そのほかの受け入れる差（2026-11-26 のレビューで決めた。正しさ・セキュリティ・データ安全）

| 場面 | GAS | VPS |
|---|---|---|
| Sebastian の本文が `null` | TypeError で HTML のエラー画面 | `{"error":"unauthorized"}` |
| Sebastian で想定外の例外 | 例外の文を `{error, action}` で返す | `{"error":"internal error", action}`（詳細はログだけ。セキュリティ 中-1） |
| Sebastian で完了を確定した後、PDF の結果待ちに例外 | — | `ok:true` のまま `pdf:{error:"PDF作成の結果を確かめられませんでした（作成は続いている可能性があります）"}`（再送による二重の完了を防ぐ） |
| ダッシュボードで想定外の例外 | 例外の文を `error` に入れる | `データを取得できませんでした。`（詳細はログだけ。RpcError の文はそのまま） |
| 残業の予定時間が整数の列に入らない大きさ（`1e20`・`1e308`） | `Infinity` などを書く | 400 `残業の予定時間が不正です: X` |
| 集計の年が 1〜9998 の外（`9999-12`・`0.5-01`） | 空の集計 | 400（同じ形の文の `yearMonth…で指定してください`） |
| 申請の文字の項目が 50,000 字を超える | `appendRow` が例外（シートのセルの上限） | 400 `<項目>が長すぎます（50000字まで）。`（GAS も失敗する所だけ。中身は変えない） |
| `api_submitRequest(null)`・`api_editRequest(id, null)` | TypeError | `{}` として扱う（`申請種別が不正です: ` など） |
| 引数が多すぎる RPC | 余りを捨てる | 同じく捨てる（足りない引数は、既定の値の無い関数だけ 400 `引数の数が正しくありません。`） |
| マスタに見出しが無い | 例外 | 空の一覧（表の作りが固定なので起きない見込み） |
| `ops_retryJob` の番号 | （GAS に無い） | bool でない整数か、数字だけの文字だけを受ける |
| 0000 年の日付 | 有効な日付として年度の検査で拒む | 同じ（Python の日付は 1 年からなので 0 年は月日だけ確かめる） |
| 一括承認の ID が `null`・`1.0` | キーは `"null"`・`"1"` | 同じ（`js_property_key`） |
| 申請 ID の前後の空白（シートの行の ID） | `normalize_(行の ID) === 引数` で探す | 同じく行の ID の前後の空白を無視して探す（引数は正規化しない） |

### 8-2b. シートへ書く値は式にしない（セキュリティ 高-1）

- 行の反映・照合は `updateCells` の `stringValue`（文字のセル）で書くので、`=1+1`・`+81-90-…`・`-部品`・`@x` は文字のまま入る（`formulaValue` はどこでも使わない。試験 `tests/test_formula_safety.py`）。
- マスタ同期は GAS の `setValues` と同じ USER_ENTERED で書くが、`=` `+` `-` `@` で始まる文字（ただの数の `-5`・`+3` を除く）は先頭に `'` を付けて文字のセルにする。
- 戻しの道具（`tools/restore_tabs_from_copy.py`）は RAW で書く。移行の道具はシートを読むだけ。
- 台帳と PDF のひな形へ書く GAS の側は GAS の担当が直す。

### 8-3. 丸め（gap-1 G5）

`jsutil.js_round(x) = math.floor(x + 0.5)`、`jsutil.js_floor = math.floor`。Python の `round()` を使わない（F の試験が `app/domain` を grep して拒む）。使う所: 実働分、月次の予測 `round(pace × 月の日数)`、残業の予定 `round(h × 60)`、画面の `fmtMin` は画面の JS のまま。滞留時間は `floor((now - submittedAt) / 1 時間)`。境界の試験: 19:29:59・19:30:00・19:30:40 の終了、予測が .5 になる例（1 分を 2 日で割り 31 日を掛ける＝15.5 → 16）。

### 8-4. 集計（[01 §5] を写す）

- 土台 `buildJoinedRecords_`: status が空・canceled を除き、対象日が `from <= d < to`（JST）。**submitted も含む**。WorkLogs を requestId で結ぶ（同じ ID が複数なら GAS の Map と同じく**後の行**が勝つ。`buildWorkLogsMapByRequestId_` は上から読んで `set` するため）。
- 月次: 40h=2400 分・60h=3600 分・注意 1800 分、判定はすべて `>=`。未来月は月末まで経過とみなす。集計の単位は `dept + '__' + workerName`。
- 特別条項: 4/1〜3/31、月合計 `>= 3600` の月数、`isDanger >= 6`、`isWarn >= 5`。
- 滞留: submitted 全件、`isOver48h: hours >= 48`。`submittedAt` が読めない行は GAS では関数全体が失敗する見込み（推測）。VPS はその行を飛ばす（受け入れる差）。
- スタッフ集計: approved の net だけ、締め月、3/16 年度、年間超過は締め月ごとの合計 `>= 2400` の月数。画面の判定 `>` は画面のまま。
- 並び: 文字の比べは JS の `<`（UTF-16 の符号順）。`localeCompare(…,'ja')` の 2 か所（総務の日次一覧、部署の円グラフ）は ICU（`PyICU` が無ければ `locale` に頼らず、JS の照合と同じ結果になる表を試験で確かめる。合わない並びは受け入れる差として報告）。

### 8-5. 入力の検査（`api_submitRequest`。文は 1 字も変えない。[01 §2.2]）

各値は `normalize_gas`。順に:
1. `申請種別が不正です: X`
2. `部署が指定されていません。`
3. `作業員が指定されていません。`
4. `理由を選択してください。`
5. `理由が「その他」の場合、補足理由が必須です。`（reason に `その他` を**含む**とき）
6. `予定時間を選択してください。`
7. 残業で遡及でなければ日付を JST の今日に置き換える。それ以外は `作業実施日を選択してください。`・`作業実施日の形式が不正です: X`・`作業実施日が無効です: X`
8. 残業: `作業実施日の形式が不正です（yyyy-MM-dd）: X`・`作業実施日が無効です: X`・`未来日は申請できません: X`・`年度開始日（yyyy-MM-dd）より前の日付は申請できません。`
9. 休日: 年度開始より前を拒む。遡及なら未来日を拒む（`遡及申請では未来日を選択できません: X`）。通常の休日は未来の上限なし
10. 400（例外）: `作業員コードが取得できません。`・`残業の予定時間が不正です: X`・`休日の予定時間が不正です: X`

11. 400（例外）: 文字の項目（部署・作業員・業務 NO.・理由・補足理由・業務内容）が 50,000 字を超えると `<項目>が長すぎます（50000字まで）。`（シートのセルの上限。GAS も書き込みで失敗する。8-2a）

1〜9 は `{ok:false, error}` を 200 で返す。`isRetroactive` は `true` か文字の `'true'`（大文字小文字を無視）で真。予定時間は一覧と照合しない（Q13。編集だけ照合）。理由も一覧と照合しない。

### 8-6. 権限（F の `permissions.py`）

- 管理者 = Settings の `ADMIN_EMAILS` と `GENERAL_AFFAIRS_EMAILS`（キーの前後の空白を取る）の和（カンマ区切り、小文字で比べる）＋`BOOTSTRAP_ADMINS`。一覧が空でも全員を管理者にしない。ApproverMap の `role=admin`（1 行目を見出しとして読む。今は効かない）も GAS と同じに見る。
- 部署の承認者 = DeptApprovers が 2 行以上あればそれだけ（A 部署・B メールのカンマ区切りを小文字に）。無いか 1 行以下なら ApproverMap。管理者は全部署。
- `assertSelf_`: `workerEmail` が空でなく、ログインのメールと大文字小文字を区別して違えば拒む（Q24 の癖のまま）。

### 8-7. PDF の方式（GAS が決める）

`PDF_MODE` は「`direct` を**含むか**」で判定する（本番の値は `direct で直接書込方式（未設定=XLOOKUP方式）`。gap-1 G4）。判定は GAS の今の関数がするので、VPS は `PDF_MODE` を解釈しない（`ops_getStatus` に表示するだけ）。二次承認はいつも直接方式、Sebastian はいつも XLOOKUP 方式（7-2）。

---

## 9. 安全装置（F が `safety.py`、最重要）

### 9-1. 環境ごとの許可リスト

| 対象 | production | staging |
|---|---|---|
| 書いてよいスプレッドシート | `PROD_SPREADSHEET_ID` だけ | `STAGING_SPREADSHEET_ID` だけ（本番 DB の「コピーを作成」で作ったサンドボックス） |
| 書いてよいタブ | `Requests`・`WorkLogs`・`BatchLogs`（`row_sync`・`reconcile` だけ）、`作業員マスタ`・`工番マスタ`・`社内カレンダーマスタ`・`業務NOマスタ`（`master_sync` だけ） | 同じ |
| 書いてはいけない（どの環境も） | マスタ元・蓄積台帳（両方の環境の ID）・PDF フォルダ・ほかのスプレッドシートすべて。B・D・E 区分のタブ（特に `DeptApprovers`・`Settings`・`申請書フォーム`・`操作`） | |
| 使ってよい書き方 | `updateCells`、`values.batchUpdate`、`values.batchClear`（C 区分の余りの行だけ）、`appendDimension`（ROWS、A 区分の末尾だけ）、`values.append`（Requests の新しい行の requestId だけ。RAW・OVERWRITE。6-4・決定 #26） | 同じ |
| 使わない書き方 | `deleteDimension`・`insertDimension`・`deleteSheet`・`addSheet`・`duplicateSheet`・`sortRange`・`moveDimension`・A 区分の `clear` | |
| Drive | api・worker は呼ばない。`tools/cutover_backup.py` だけが複製（元は `PROD_SPREADSHEET_ID` と `PROD_LEDGER_SPREADSHEET_ID` だけ） | 同じ |
| GAS の内部用の入口 | `internal.ping` の `spreadsheetId == PROD_SPREADSHEET_ID`、`scriptId == PROD_GAS_SCRIPT_ID_FIXED`、Settings の `TEMPLATE_SSID == PROD_SPREADSHEET_ID`、`PDF_ROOT_FOLDER_ID == PROD_PDF_FOLDER_ID`、`ACCUMULATION_SS_ID == PROD_LEDGER_SPREADSHEET_ID`、`spreadsheetTimeZone == DB_SHEET_TIMEZONE` | `spreadsheetId == STAGING_SPREADSHEET_ID`、`scriptId != PROD_GAS_SCRIPT_ID_FIXED`（サンドボックスに紐づく GAS の写し）、Settings の 3 つの ID が `STAGING_*` と同じで、どれも本番の ID と違う、時間帯が同じ |

- 確かめる時: 起動時（設定の検証。合わなければ起動しない）と、**書き込みの直前**（`SheetsApi` の書く関数の中で `safety.assert_writable(spreadsheet_id, tab, purpose)`）。
- GAS の照合が合わなければ、`sync_runs(kind='gas_ping', ok=false)` に理由を残し、GAS のレーンは全部止まる（7-3 の 4）。管理者の欄と `/healthz` の詳細に出す。
- staging の準備（ディレクターがする）: 本番 DB の「コピーを作成」（紐づく GAS の写しも一緒にできる）、台帳のコピー、staging 用の PDF フォルダ。写しの Settings の `TEMPLATE_SSID`（＝写し自身）・`PDF_ROOT_FOLDER_ID`・`ACCUMULATION_SS_ID` を staging の ID に直す。写しの GAS に G のコードを入れ、内部用のデプロイを作り、`VPS_INTERNAL_TOKEN` を置く。写しのトリガーは作らない。
- staging の `GAS_INTERNAL_URL`・`GAS_INTERNAL_TOKEN` は本番と別の値。本番の GAS は切り替えの準備（計画書 6-2）まで呼ばない。
- `/healthz` は外には `{"ok":true,"env":"staging|production"}` だけ。詳細（遅れ・GAS の照合）はログインした管理者が `/healthz?detail=1` で見る。

### 9-2. ほかの守り

- 本番 GAS（`src/`）を `master` に入れない（2-1）。
- 委任のスコープは 1-1 のまま。足すときは先に VPS に置いてからオーナーに許可し直してもらう（罠 7）。
- `tools/qa/acceptance.py` は `APP_ENV=staging` でしか動かない。作ったセッションは最後に必ず消す。作る申請の業務内容に【受け入れテスト】の印。
- 移行の道具は `--source sheets` のとき `PROD_SPREADSHEET_ID` 以外を拒む（staging への移行は本番を読む＝読むだけ）。

---

## 10. 担当と境目

### 10-1. 担当（それぞれ git worktree の自分のブランチ。ほかの担当のファイルを書き換えない。push しない）

| 担当 | 範囲（書いてよいファイル） | 始める条件 |
|---|---|---|
| **F 土台** | `docker-compose.yml`・`.env.example`・`api/Dockerfile`・`requirements*.txt`・`pytest.ini`・`alembic*`（全表）・`app/main.py`・`config.py`・`db.py`・`models.py`（全表）・`timeutil.py`・`jsutil.py`・`constants.py`・`safety.py`・`auth.py`（`/auth/me` も）・`google_delegate.py`・`permissions.py`・`rpc.py`（38 関数の閉じたリスト、未登録は 501）・`serialize.py`・`sheet_queue.py`・`audit.py`・`jobs.py`（レーン・優先度・`requires`・`register_kind`）・`worker.py`（3 本のスレッド・予定・`run` の CLI）・`services/sheets_api.py`・`services/gas_client.py`・`/api/sebastian` の経路・`web/nginx.conf`・`backup/`・`cloudflare/`・`docs/RUNBOOK.md`・`tests/conftest.py` と土台の試験 | すぐ |
| **U 画面** | `app/pages.py`・`app/templates/pages/*`・`web/static/**`・`web/dev/**`・`tests/test_pages.py`・`tests/test_static_assets.py` | すぐ（10-2 の形を信じ、mock-server で単独で進める） |
| **G GAS** | `src/Internal.gs`・`src/Cutover.gs`・7-6 の決まった行・`gas-tests/**` | すぐ |
| **W 業務** | `app/domain/**`（`ops.py` を含む）・`tools/qa/acceptance.py`・`tests/test_domain_*.py`・`tests/test_sebastian.py` | F を main に取り込んだ後 |
| **D データ** | `services/sheet_codec.py`・`row_sync.py`・`reconcile.py`・`config_sync.py`・`master_sync.py`・`tools/migrate_from_sheets.py`・`tools/cutover_backup.py`・`tools/restore_tabs_from_copy.py`（控えから A・C 区分のタブを値と表示形式ごと写し戻す。gap-4 F-10）・対応する試験 | F を取り込んだ後 |
| **M GAS の仕事** | `services/gas_jobs.py`・`morning_batch.py`・`ops_status.py`・対応する試験 | F を取り込んだ後 |

作る順: **F・U・G を並行 → F を main へ（全試験が通ってから）→ W・D・M を並行 → U・G・W・D・M を順に取り込み、全試験 → つなぎ目の確かめ → レビュー 2 本（正しさ・安全）→ 修正。** U と G は F を待たないが、取り込みは F の後。

### 10-2. 担当をまたぐ関数（名前と引数をここで固定する。呼ぶ側は相手の完成を待たず、この形を信じて書く。試験では偽物を入れてよい）

| 関数 | 持ち主 | 呼ぶ側 |
|---|---|---|
| `config.get_settings() -> Settings` | F | 全員 |
| `db.get_db()`・`db.get_sessionmaker()` | F | 全員 |
| `timeutil.now_utc() -> datetime`・`now_jst()`・`today_jst() -> date`・`to_jst_text(dt, pattern) -> str`（`'yyyy-MM-dd'`・`'yyyy-MM-dd HH:mm'`・`'yyyy-MM-dd HH:mm:ss'`・`'M/d'` などの GAS の書き方を受ける） | F | 全員 |
| `jsutil.js_round(x: float) -> int`・`js_floor(x) -> int`・`normalize_gas(s) -> str`・`js_sort_key(s) -> tuple` | F | W・D |
| `permissions.is_admin(db, email) -> bool`・`dept_approver_map(db) -> dict[str, list[str]]`・`can_approve_dept(db, email, dept) -> bool`・`is_any_dept_approver(db, email) -> bool`・`approver_profiles(db) -> list[dict]` | F | U（pages）・W |
| `rpc.register(name)`（デコレーター）・`RpcContext`・`RpcError(message)` | F | W |
| `auth.require_user`・`auth.get_optional_user`・`auth.revoke_sessions(db, email) -> int` | F | U・W |
| `sheet_queue.mark_dirty(db, table_name: str, pk: str) -> int`（新しい `row_version` を返す） | F | W・M・D |
| `sheet_queue.current_version(db, table_name, pk) -> int` | F | M |
| `jobs.enqueue(db, kind, args: dict, *, lane="gas", request_id="", priority=0, requires: list[dict]=(), guard: dict|None=None, batch_id: str="") -> int` | F | M |
| `jobs.register_kind(kind, handler, *, lane, on_exhausted=None)`・`JobContext`・`SkipJob`・`RetryLater(seconds)` | F | M・D |
| `worker.register_schedule(name, *, at=(h, m) | every_minutes=n, task)` | F | M・D |
| `services.sheets_api.SheetsApi.for_db(db)`（`get_spreadsheet`・`batch_get_values(ids, ranges, render)`・`batch_update(spreadsheet_id, requests, *, tab, purpose)`・`values_batch_update`・`values_batch_clear`・`append_rows(spreadsheet_id, tab, n)`。書く関数はすべて `safety.assert_writable` を通す） | F | D |
| `services.gas_client.GasClient.from_settings(settings).call(action: str, *, job_id: str, request_id: str = "", expect: dict | None = None, args: dict | None = None, timeout: float | None = None) -> dict`（返事の JSON。通信の失敗・時間切れは `GasTransportError`） | F | M |
| `safety.assert_writable(spreadsheet_id, tab, purpose)`・`safety.verify_gas_ping(ping: dict, settings) -> list[str]`（合わない点の一覧。空なら合格） | F | D・M |
| `domain.sebastian.handle(db, raw_body: bytes) -> dict` | W | F（経路） |
| `domain.breaks.calc_break_by_clock_time(start: datetime, end: datetime) -> int` | W | W（Sebastian・再計算） |
| `domain.fiscal.fy_start_0316(d) -> date`・`fy_start_0401(d) -> date`・`closing_ym(d) -> str` | W | M（朝バッチの対象は使わない。参考） |
| `domain.ops.*`（4-2 の 10 関数） | W | U（画面から RPC で） |
| `services.gas_jobs.enqueue_after_approve(db, request_id) -> int` | M | W |
| `services.gas_jobs.enqueue_after_complete(db, request_id, *, source: str, status_at_enqueue: str, wl_override: dict | None = None, priority: int = 0) -> int` | M | W |
| `services.gas_jobs.enqueue_regenerate(db, request_id, *, reason: str, old_pdf_file_id: str = "", batch_id: str = "") -> int` | M | W |
| `services.gas_jobs.enqueue_ledger_upsert(db, request_id, *, batch_id: str = "") -> int` | M | W |
| `services.gas_jobs.enqueue_after_delete(db, request_id, *, request_type: str, target_date: str, old_pdf_file_id: str) -> int` | M | W |
| `services.gas_jobs.enqueue_ledger_rebuild(db) -> int` | M | W |
| `services.gas_jobs.wait_for_result(session_factory, job_id: int, timeout_seconds: float) -> dict | None`（`None`＝待ち切れ） | M | W（Sebastian） |
| `services.gas_jobs.record_pdf_result(db, request_id, pdf: dict) -> None`（6-5） | M | M（D は試験で使う） |
| `services.morning_batch.enqueue_manual_batch(db, ymd: str) -> int` | M | W |
| `services.ops_status.get_ops_status(db) -> dict` | M | W |
| `services.config_sync.sync_sheet_config(db) -> dict`（`{ok, counts, overlaps, mismatches}`） | D | W・worker |
| `services.master_sync.sync_masters(db) -> dict`（`{ok, tabs:{名前:{ok, rows, skipped_reason}}}`） | D | W・worker |
| `services.row_sync.run_once(session_factory, limit=50) -> dict` | D | worker（F が登録の口を作り、D が `register_schedule`／レーンに入れる） |
| `services.reconcile.run(session_factory, *, check_only: bool) -> dict` | D | worker・tools |
| `services.sheet_codec.COLUMN_RULES`・`encode_cell(rule, value, kind, tz) -> dict`・`decode_cell(rule, unformatted, formatted, tz) -> (value, kind)` | D | D・tools |
| `services.sheet_codec.decode_row(spec, col, urow, frow, tz, sheet_row) -> (record: dict, errors: list[str])`・`synced_values_for(spec, record) -> dict`（1 行の読み方と反映済みの値。決定 #26） | D2 | tools（移行）・`sheet_intake` |
| `services.sheet_intake.run(session_factory, *, api=None, app_settings=None, wait_seconds=0.0) -> dict`（`{ok, active, ingested, skipped, busy?, stopped?}`）・`run_now(db) -> dict`・`ensure_request(db, request_id) -> bool`・`status(db) -> dict`（6-10） | D2 | worker・W（`ops_runSheetIntake`・Sebastian の入口）・M（`ops_status`） |
| `services.sheets_api.SheetsApi.append_values(spreadsheet_id, tab, values, *, purpose="row_sync") -> int`（書いた最初の行。RAW・OVERWRITE、Requests の requestId だけ） | D2（F の土台に足した） | D（`row_sync.reserve_new_rows`） |
| `google_delegate.get_access_token(scopes=None) -> str`・`get_delegate_credentials(db)` | F | D・tools |

`wl_override` の形は `{"actualStartAt": "yyyy-MM-dd HH:mm:ss", "actualEndAt": "…", "actualMinutes": int, "breakMinutes": int, "netMinutes": int}`。`requires` の形は `[{"table": "requests", "pk": "<requestId>", "version": 7}, {"table": "work_logs", "pk": "<id>", "version": 3}]`（M が `sheet_queue.current_version` で作る）。

### 10-3. worker の形（F が骨組み、D・M が中身を登録）

- `python -m app.worker` は 3 本のスレッドを動かす: ①行の反映のレーン（`row_sync.run_once` を 2 秒ごと）、②GAS のレーン（`lane='gas'` の仕事を 1 本ずつ。`pg_try_advisory_lock` で全体 1 本を守る）、③予定と `lane='default'` の仕事。スレッドごとに DB のセッションを持つ。SIGTERM で止まる。
- 予定（`sync_runs` の `schedule:<名前>` で枠を判定。起動の直後に走らせ直さない）:

| 名前 | 時刻 | 持ち主 |
|---|---|---|
| `sync_masters` | 毎日 0:30 | D |
| `sheet_reconcile` | 毎日 2:10 | D |
| `sync_sheet_config` | `SHEET_CONFIG_SYNC_MINUTES` ごと | D |
| `morning_batch` | 毎日 6:50 | M |
| `gas_ping` | 5 分ごと（起動時にも 1 回） | M |
| `sheet_intake` | 1 分ごと（6-10。production は GAS が移転モードのときだけ取り込む） | D2 |
| `purge` | 毎日 3:30（冪等キー 24 時間、期限切れのセッション、`done` の仕事 90 日、`audit_log` は消さない） | F |

- `python -m app.worker run <名前>`（例 `run sync_masters`、`run sheet_reconcile --check`）で手で動かせる。

---

## 11. 変えるもの／変えないもの

### 11-1. 変えるもの（これ以外は変えない）

| 変えること | 根拠 |
|---|---|
| 全画面で Google ログイン（共有アカウントのまま）。メールが空のときの「管理者扱い」「判定を飛ばす」「`unknown`」が起きなくなる | #3・#4 |
| ログインの持ち時間 90 日（無操作でも）。ログインで `prompt=select_account`・`consent` を付けない | #24・gap-2 S2 |
| 保守の処理は管理者だけ（4-2）。画面から呼ばれない公開関数を出さない。`?date=` の形を確かめる | #13 |
| 旧フォームの経路を移さない（`api_getFormUrl`・取り込み・選択肢の更新）。`?page=redirect` は top へ | #11・Q9 の既定 |
| 書き込みに冪等キーと状態の照合（二重押しで 2 件できない） | #5 |
| PDF・台帳は GAS が裏で作る。押した直後の返事の `pdf` は `null`、`pdfRegenerated` は `false`。PDF のリンクは数十秒〜で出る | #19 |
| PDF の仕事の失敗を再試行する（今なら PDF なしで終わる申請に後から PDF と台帳の行ができる） | Q20 の既定・gap-1 G7 |
| 台帳が開けないとき新しい台帳を作らず、失敗として知らせる。書く前に見出しを確かめる | 計画書 4-7・gap-5 中-1 |
| マスタ同期で、元が読めない・空なら DB のタブを上書きしない | #21 |
| シートの業務データのタブは毎晩の置き換えではなく、書くたびに行単位で反映（行の位置・型・表示形式を保つ） | #18 |
| 共有タブ（Settings・DeptApprovers）を直してから VPS の判定に効くまで最大 5 分（［今すぐ取り込む］あり）。GAS の PDF の印鑑はすぐ効く | Q12 の既定 |
| 総務の日次一覧が表示される（日時を文字で返す。過去申請のバッジも出る）。手動入力ボタンの XSS を直す | #22・gap-2 S1・S8 |
| Sebastian の残業完了の休憩を画面と同じ時刻窓方式に（過去の行は直さない） | #23 |
| スマホ向けの配置（PC は今と同じ） | #25 |
| 権限エラー画面にデバッグ情報を出さない | Q8 の既定 |
| エスケープの抜けを直す（`'`、PDF リンク）。iframe への埋め込みを自分のサイトだけに | 3-8・Q10 の既定 |
| 二次承認一覧が見出しを書き足さない | Q17 |
| 一時コピーの片付け・テンプレの書かないセルは、スパイクの後でオーナーが決めるまで変えない | 計画書 4-6・Q19 |
| 受け入れる小さな差: 自由入力の `1/2` などが日付にならず文字のまま、`invalid JSON:` の後ろの文、時刻の範囲外の文、`submittedAt` が読めない滞留の行を飛ばす、ロック待ちの GAS のエラー文が出ない | 6-3・4-4・8-2・8-4 |

### 11-2. 変えないもの（計画書 3-2 の癖を含む。主なもの）

権限の無い作成・編集・削除・TOP・スタッフ集計、自己承認、17:20 固定、窓をまるごと含むときだけの休憩、日をまたぐと休憩 0、境界の `>=`（スタッフ集計の画面だけ `>`）、年度 2 通り、月次は submitted も数える、未来月は月末まで経過、承認者画面の canceled の当日表示、過去の submitted 残業は TOP に出ない、休日の日付の編集で年度・未来を確かめない、PDF が無い申請の編集は台帳を書かない、完了の再実行は上書きで PDF は作り直さない、申請時は予定時間を照合しない、通常の休日は未来の上限なし、朝バッチは未完了でも PDF を作る（Q18）、台帳の癖（1 次承認日が空・朝バッチの薄い行・却下で消さない）、作業員の自動選択が効かない、承認者画面の KPI の数え方、STEP 2 のサマリーに過去日が出ない、admin のフィルタは［月次を更新］まで反映しない、CSV は BOM なし、approve2 の注記「翌朝バッチで…」、単票の二次承認に confirm なし、メールの大文字小文字の扱い（Q24）、作業員マスタ 80 人の上限、`charts.dept` に全部署（H14）、画面の文言・alert・confirm のすべて（[02 §2.8]）、エラー文のすべて（[01 §1.2・§1.4]）、メールを送らない。

---

## 12. テストと受け入れ表

### 12-1. 動かし方

- Python: F が `vps-app/api/.venv` を作る（`python -m venv .venv` → `.venv/Scripts/python -m pip install -r requirements.txt -r requirements-dev.txt`）。各担当は `cd vps-app/api && .venv/Scripts/python -m pytest -q`。**最後の行（passed/failed の数）まで読み、1 件でも落ちていたら報告に書いて先へ進まない**（罠 44）。SQLite（メモリー）で動くこと。外の API（Google・GAS）は呼ばない（偽物を入れる）。
- PostgreSQL の結合試験（ディレクターが VPS で）: `docker build --target test`、DB `overtime_test`、`TEST_DATABASE_URL` を渡して `python -m pytest`（罠 51）。行ロック・`skip_locked`・advisory lock・jsonb はここで確かめる（`@pytest.mark.postgres`）。
- JS（U）: `node --test vps-app/web/dev/`（シムの試験。外部の npm パッケージを使わない）。画面の確かめは `node vps-app/web/dev/mock-server.js`（nginx と同じ CSP、`/__calls`、`/__fail?status=401|500&count=1`）。
- GAS（G）: `node --test gas-tests/`。`src/*.gs` を `vm` の文脈に読み込み、`SpreadsheetApp`・`PropertiesService`・`CacheService`・`LockService`・`DriveApp`・`ContentService`・`UrlFetchApp`・`Utilities`・`Session`・`ScriptApp`・`Logger` の偽物で動かす。clasp を使わない。
- 受け入れ（ディレクターが staging で）: `docker compose exec -T api env PYTHONPATH=/app python - < tools/qa/acceptance.py`。

### 12-2. 担当ごとの必須の試験

| 担当 | 試験 |
|---|---|
| F | `test_config_safety.py`（必須値・定数との照合・staging と本番の ID の取り違え・タブの許可リスト・使わない書き方の拒否・`verify_gas_ping` の全条件・ほかのモジュールへの ID の直書きを拒む grep）、`test_rpc.py`（Origin・415・404・400・501・401・403・冪等キーの再生と 409・業務エラーで鍵が残らない・38 関数の閉じたリスト）、`test_auth.py`（PKCE・hd・90 日・`prompt` を付けない・`/auth/me`・`safe_next`）、`test_google_delegate.py`、`test_jobs.py`（レーン・優先度・同じ申請の順番・前が failed なら止まる・`requires` で待つと試行回数が増えない・リース切れの拾い直し）、`test_worker_schedule.py`（`sync_runs` の枠・起動直後に走らない）、`test_timeutil.py`（`date.today`・`datetime.now`・Python の `round(` を `app/` から grep で拒む）、`test_jsutil.py`（`js_round(14.5)==15`・`normalize_gas` の全角）、`test_permissions.py`（2 キーの和・空でも全員にしない・DeptApprovers 2 行以上・ApproverMap の 1 行目見出し）、`test_migrations.py`、`test_web_config.py`（CSP・インラインの script と on 属性が 0・`mobile.css` の規則がすべて 820px 以下の `@media` の中・本番イメージに tests が無い・nginx のヘッダーが server レベル） |
| U | `test_pages.py`（ログインへの 302 と `next`・page の振り分け・不明は top・`redirect` は top・admin/approve2/approver の `no_auth` と文・デバッグ情報が無い・題名・`DATE_PARAM` の検証と `</script>` を含む値・テンプレート変数）、`test_static_assets.py`（5 ページに on 属性 0・`escHtml` が `'` を逃がす・`[onclick*=` が無い・`new Date(` に空白入りの文字を渡す箇所が無い）、`web/dev/shim.test.js`（ハンドラの付け方 4 通り・省略引数・200 の `{error}` が成功に届く・400 で `err.message`・401 で失敗ハンドラを呼ばずログインへ・下書きの退避と復元・同じ呼び出しの二重送信が 1 回・409 の送り直し・冪等キーの保持と破棄） |
| G | `gas-tests/internal.test.js`（token 違い・無し・`jobId` の再生・`expect` の不一致・各 action の呼び順と条件（7-2 の表の行ごと）・台帳が開けないとき作らない・見出しの不一致・PDF が例外のとき afterApprove は台帳を書かず afterComplete は書く・前の PDF をゴミ箱へ移せなくても続ける）、`cutover.test.js`（移転モードでないとき今と同じ・`doGet` の移転ページに `page`・`date` と `target="_top"`・不正な `page`／`date` を捨てる・`getDb_()` の止めと内部用の印・トリガー 4 関数が何もしない・F-7 の入口が止まる・`onOpen` が 1 項目・Sebastian の転送の 4 通り・運用関数が持ち主以外で拒む） |
| W | `test_domain_requests.py`・`worklog.py`・`approval.py`・`secondary.py`・`dashboards.py`・`summaries.py`・`masters.py`・`ops.py`（29 関数それぞれ: 返す形のキーと日時の形・エラー文の 1 字ずつ・積む仕事と `mark_dirty`・権限）、`test_breaks.py`（窓の境界・19:29:59／19:30:00／19:30:40・日またぎ 0・17:10 の窓は 17:20 開始で数えない）、`test_fiscal_jst.py`（**06:30 JST に通常残業を出すと targetDate が当日**、**3/16 08:00 JST に 3/15 の遡及残業が拒まれる**、**15 日 23:30 と 16 日 00:30 の締め月**、4/1 と 3/16 の年度、30 日の窓）、`test_rounding.py`（予測 15.5→16）、`test_concurrency.py`（承認と完了が同時でも PDF の仕事が必ず 1 つ積まれる＝`postgres` マーカーと SQLite の順序の両方、一括の ID の並べ替え、**2 つのタブで同じ申請を承認しても仕事は 1 つ**、**同じ冪等キーの二重送信で申請は 1 件**）、`test_sebastian.py`（4 action × 成功・各エラーの返事が [01 §1.2] の表と 1 字ずつ同じ、`text/plain` の本文、`targetDate` の 2 通りの形、`pdf` の 4 形と待ち切れ、休憩が時刻窓方式） |
| D | `test_sheet_codec.py`（シリアル値と時間帯の往復・`cell_kinds`・`auto` の解決・表示形式の写し・WorkLogs の `@`）、`test_row_sync.py`（差分だけ書く・新しい行は末尾で作った順・ID の確かめと探し直し・見つからなければ止める・見出しの違いで止まる・時間帯の違いで止まる・GAS が書いた 3 列を上書きしない・削除は値の書き換え・行やタブを消す要求を出さない）、`test_reconcile.py`（秒の単位・反映待ちの行は比べない・DB が少なければ書かない・`--check`）、`test_config_sync.py`（Settings の下の行が勝つ・キーの空白・空なら上書きしない・DeptApprovers の 5 列を持つ）、`test_master_sync.py`（80 人・Date の規則・USER_ENTERED・余りの行は値だけ消す・空なら上書きしない・タブごとの失敗）、`test_migrate_from_sheets.py`（本番 ID 以外を拒む・空でなければ止める・`modified` で止める・孤児と行の無い申請・`sheet_row` と `cell_kinds` と `synced_values` が入る・移した直後の照合が差 0・連番を進める＝罠 9） |
| M | `test_gas_jobs.py`（7-4 の表の行ごとに積む action と args・`requires` を待つ・canceled なら skipped・`record_pdf_result` で 3 列と `synced_values` が同じになる・再試行の間隔と 24 時間・`retryable:false` で failed・同じ申請の順番・`batch_id` の 500ms・`gas_ping` が合わないと止まる）、`test_morning_batch.py`（成功・スキップ・失敗の数、エラー詳細の形、BatchLogs の 1 行の値と型、ledgerRewriteDay が最後、6 件以上の待ち、Sebastian を先に流す、送り直しで ledgerRewriteDay をもう一度）、`test_ops_status.py` |
| D2 | `test_sheet_intake.py`（決定 #26: Sebastian と同じ `values.append`（USER_ENTERED・INSERT_ROWS）の行を取り込み、シートへは書かない・反映済みで入り行の反映と照合が書き直さない・WorkLogs を作らない・**Sebastian の追記と VPS の新しい行が同時**（VPS が行を取る前／取った後の 2 通り）でも互いの行が残る・**取り込んだ行を承認すると同じ行の 3 列だけが変わる**・行がずれても requestId で探し直す・requestId の空・重なり・決まりに無い status・日付でない targetDate・読めない分は入れずに記録し、行の反映と照合を動かしてもその行は変わらない・**冪等と鍵**（入れ子の実行は busy、2 度目は何もしない）・production は移転モードのときだけ・時間帯が違えば止まる・1 分ごとの登録・`ops_runSheetIntake` と `ops_getStatus.sheetIntake`・Sebastian の完了がその場で取り込む・`values.append` は Requests の文字だけ）、`test_reconcile.py`（取り込み待ちの行を数えず書かない、上限を超えたら止める）、`test_cutover_tools.py`（戻しの道具が控えに無い行を消す前に止まる） |

### 12-3. 受け入れ表（計画書 3 章と抜け調査の指摘 → 確かめ方）

「自動」は 12-2 の試験か `tools/qa/acceptance.py`。「JEV」は JEV ハンズでオーナーが確かめる実機の操作（スキル 05 §4-2。報告は「実機で押した／自動だけ／まだ」に分ける）。「SB」はサンドボックス（DB のコピーと紐づく GAS の写し）での照合。

| # | 項目（出典） | 確かめ方 |
|---|---|---|
| A-1 | 画面 5 ページ＋権限エラー、`?page=`、不明は top、`redirect`（3-1） | 自動 U `test_pages.py`。JEV: 各ページを開き、ヘッダーのリンクを全部押す |
| A-2 | 入口判定（admin・approve2・approver）（3-1） | 自動 U。JEV: 承認者でない共有アカウントで approver・admin を開き権限エラー |
| A-3 | テンプレート変数・`DATE_PARAM`（3-1、[02 §6] 高） | 自動 U。JEV: `?page=approve2&date=2026-11-25`、`?page=admin&date=2026-11-25` で自動表示 |
| A-4 | 29 関数の形・エラー文・失敗の 6 通り（3-1、gap-2 S6・S9） | 自動 W と acceptance。SB: 同じ入力で GAS と VPS の返事を突き合わせる（日次一覧は GAS が null なので比べない） |
| A-5 | Sebastian の 4 action（3-1、gap-1 G6、gap-5 高-1） | 自動 W `test_sebastian.py`。SB: 写しの GAS の転送を通して `ping`・完了。切り替え後に本物の Sebastian の完了 1 件（6-5） |
| A-6 | 状態の遷移・削除の副作用（3-1、gap-1 G12） | 自動 W。JEV: 申請 → 承認 → 完了 → 二次承認 → 削除を 1 件通す |
| A-7 | 実績と休憩、Sebastian の休憩（3-1、#23、gap-1 G13・G14） | 自動 W `test_breaks.py`・`test_sebastian.py` |
| A-8 | 集計（月次・年度推移・特別条項・滞留・スタッフ集計）（3-1、gap-1 G5・G11・G18） | 自動 W。SB: 総当たり比較（全月×全部署、全締め月×全員。計画書 5-4） |
| A-9 | PDF の見た目（3-1、gap-3 G3-H4） | SB: スパイクの 10 件を本番の PDF と画素ごとに比べる。オーナーが実物を見る |
| A-10 | PDF を作る場面×方式×既存の PDF（3-1、gap-1 G8・G9、gap-3 G3-H2） | 自動 G `internal.test.js`・M `test_gas_jobs.py`。SB: 各きっかけを 1 件ずつ通し、PDF と台帳の行を見る |
| A-11 | 台帳（3-1、gap-3 G3-L5、gap-5 中-1） | SB: 台帳の行を GAS が書いた行と 1 セルずつ比べる。見出しを壊したコピーで `ledger_header_mismatch` |
| A-12 | 朝バッチと BatchLogs（3-1、gap-3 G3-H1、gap-4 F-15、gap-5 中-4・中-8） | 自動 M `test_morning_batch.py`。SB: 前日分 3 件（1 件は PDF あり）で流し、BatchLogs の 1 行・台帳・7:05 の確かめ |
| A-13 | マスタ同期と DB の 4 タブ（3-1、#21、gap-4 F-6・F-12、gap-5 高-3） | 自動 D。SB: GAS の `syncAllMasters` の後と VPS の `sync_masters` の後のタブを 1 セルずつ比べる（値と表示） |
| A-14 | 承認者の表と管理者（3-1、gap-1 G10、gap-4 F-5・F-11） | 自動 F `test_permissions.py`・D `test_config_sync.py`。B-16 で DeptApprovers が変わっていない |
| A-15 | トリガーと旧フォーム（3-1、#11、gap-4 F-7・F-8・F-17・F-19） | 自動 G `cutover.test.js`。手順書のリハーサル |
| A-16 | ロックと同時実行（3-1、gap-1 G2・G3・G17、gap-3 G3-M1） | 自動 W `test_concurrency.py`（postgres）。SB: 完了済み 20 件の一括承認で台帳に 20 行が重ならない |
| A-17 | 今のまま移す癖（3-2 の全項目） | 自動 W（項目ごとに 1 試験。名前に Q 番号を入れる）。JEV: 作業員の自動選択が効かない、STEP 2 のサマリー、CSV |
| A-18 | 直すと決めたもの（3-3 の全行、11-1） | 自動（各担当）。JEV: 日次一覧が出る・手動入力ができる（#22） |
| A-19 | **06:30 JST の試験**（gap-1 G1、決定 C） | 自動 W `test_fiscal_jst.py`。SB: 朝 6:30 台に staging で通常残業を 1 件出し、targetDate とシートの行を見る |
| A-20 | **二重押し**（#5、罠 17） | 自動 F・U・W。JEV: 申請送信を素早く 2 回押して 1 件 |
| A-21 | **2 つのタブでの承認**（罠 18） | 自動 W。JEV: 2 つのタブで同じ申請を承認し、2 回目が「既に承認済みです。」相当で PDF は 1 つ |
| A-22 | **Safari の日付**（[02 §5]、罠 16） | 自動 U `test_static_assets.py`。JEV 不可のため iPhone の実機（オーナーの写真）で TOP の申請・完了・期限超過の表示、承認者画面 |
| A-23 | 401 の注入 9 件（gap-2 S3） | 自動 U `shim.test.js`。mock-server の `status=401` で楽観更新の 9 操作を手で押す |
| A-24 | 移転ページと iframe（gap-2 S4） | 自動 G。リハーサル: GAS の `/exec?page=approve2&date=…` のボタンで新しいアプリが最上位の窓で開く |
| A-25 | CSP で全ボタンが動く（gap-2 S5） | 自動 F・U。JEV: 5 ページの全ボタン・入力欄・マウスを載せる表示 |
| A-26 | スマホ（#25、gap-2 S7） | 自動 F（`mobile.css`）。iPhone と現場の Android の実機で TOP の申請・完了、承認者画面の承認、二次承認。PC（1366px）で今の画面と見比べて差が無い |
| A-27 | XSS（gap-2 S8、[02 §6]） | 自動 U。JEV: 3-8 の氏名の申請を作り 5 画面を開く |
| A-28 | 題名・改行・CSV（gap-2 S10） | 自動 U。JEV: ［表示中をCSV］が保存できる |
| A-29 | 行の反映の型（gap-4 F-1、gap-5 高-2・中-9、決定 A） | SB: VPS が足した行・直した行を写しの GAS の `getValues()` で読み、日時の列が `Date`。`getRequestById_`・`computeIsRetroactive_`・承認者画面の提出列・未承認滞留が元の値と同じ |
| A-30 | 行の位置と並び・消さない（gap-4 F-2・F-3・F-4、gap-3 G3-M5） | 自動 D。SB: 人が行を挿入したコピーで探し直しと止まり方 |
| A-43 | Sebastian の直接の追記（決定 #26、6-10） | 自動 D2 `test_sheet_intake.py`。SB: ステージングの写しへ Sebastian と同じ呼び方で 1 行足し、2 分以内に画面に出る → 承認 → 同じ行が書き換わる |
| A-31 | 毎晩の照合（決定 A） | 自動 D。SB: 移行の直後に `--check` で差 0（計画書 5-6） |
| A-32 | GAS が止まったとき（決定 B、4-6a） | SB: 写しの GAS の token を外して操作し、画面は動き、仕事が再試行され、管理者の欄に出て、戻すと流れる |
| A-33 | ステージングから本番に書かない（9 章） | 自動 F。staging で本番の ID・台帳・マスタ元への書き込みが拒まれる。切り替えの B-14 |
| A-34 | Settings の重なり・時間帯（gap-5 中-2・高-2） | 自動 D・F。段階 2 で数えた結果を `.env` と照合 |
| A-35 | 印鑑の画像の例外・前の PDF の持ち主が違う（gap-3 G3-L1、gap-5 中-3） | SB: 持ち主が部署のアカウントの PDF を持つ申請の編集・二次承認 |
| A-36 | 再試行で PDF が 2 つできない（gap-3 G3-M3） | 自動 G（`jobId` の再生）・M |
| A-37 | `PDF_MODE` を含むかで判定（gap-1 G4、gap-5 低-3） | 自動 G（偽物の Settings に本番の文） |
| A-38 | 候補日の逃げ道（gap-5 中-6、gap-4 F-9） | 自動 W `test_domain_masters.py` |
| A-39 | `updatedBy` の決まった値・過去申請の求め方（gap-5 低-2・中-7） | 自動 W・D。SB: 月×種別の過去申請の件数を新旧で比べる |
| A-40 | 移行の照合（計画書 5-4） | 自動 D と B-12 の SQL |
| A-41 | 速さ | acceptance で中央値を測る（目安: 一覧 100ms 以下、承認 100ms 以下） |
| A-42 | 手順書のリハーサルと戻しの練習（計画書 5-6・6-6） | ディレクターが SB と staging で 2 回 |

---

## 13. 意図して残す固有の文字列

（各担当が、残した理由と一緒にここへ足す。）

- `imaizumi@lineworks-local.info`: `.env.example` の `BOOTSTRAP_ADMINS`・`DELEGATE_EMAIL` の例と `docs/RUNBOOK.md`。コードと試験には書かない（試験は `owner@lineworks-local.info`）。
- 1 章の ID: `app/config.py` の定数と `.env.example` の見本だけ。
- `C:\Users\imaizumi.LINEWORKS-NET\.secrets\overtime\`: `docs/RUNBOOK.md`（計画書 4-0 の確定値）。
- `cloudflare/setup_tunnel.py` の `TOKEN_ENV`（社内共通の Cloudflare API トークンの置き場）: 参照実装と同じなので残す。

## 確かめること（段階 2b のスパイクと w2403c で）

- `DB_SHEET_TIMEZONE` の値と、Requests・WorkLogs・BatchLogs の型の表（6-3 の `COLUMN_RULES` を直す）。
- PDF 1 件の所要時間と、GAS の `UrlFetchApp` の待ち時間の上限（`SEBASTIAN_PDF_WAIT_SECONDS` を決める）。
- Sebastian が使っている deploymentId と、Cloudflare が GAS からの転送を通すか。
- 台帳のタブの見出し・`ACCUMULATION_SS_ID` の行数・Settings の重なり。
- ポータルが iframe で埋め込んでいるか（Q10。`frame-ancestors` の 1 行）。
