# スキルに書き足す候補（残業・休日出勤申請の移行で分かったこと）

- clasp 3.4.1 の `update-deployment <id> -V N` は「Redeployed … @N」と出ても、版が変わらないことがある（説明文が英数字でも起きた。2026-11-26）。`list-deployments --json` で versionNumber を必ず確かめ、変わっていなければ Apps Script API `PUT /v1/projects/{scriptId}/deployments/{deploymentId}`（`deploymentConfig.versionNumber`）で当てる。clasp の認証ファイルのトークンで呼べる。
- 本番 GAS の時間主導トリガーは公開版ではなく HEAD のコードで動く。移行用の変更（移転モードなど）を master に入れて clasp push すると、その瞬間に本番の裏の処理が変わる。切り替えの合図まで別ブランチに置く。
- GCP コンソールは、時間がたつと本人確認のパスワードを求める（OAuth クライアント作成の画面）。オーナーが席を外す前に、GCP の作業を先に済ませる。
- 「スプレッドシートのコピーを作成」でできたサンドボックスは、Settings タブが本番の PDF フォルダ・台帳・フォーム・宛先を指したまま。コピー直後に Settings を試験用へ書き換え、フォームの対応表のタブは名前を変えて無効にする。Sheets の自動補完で値が変わることがあるので、書いた後に画面で確かめる。
- JEV ハンズの画面読みは ID を伏せ字にする。ID はボタンで写して PowerShell の Get-Clipboard で読むか、Drive API の検索で取る。
- 本番の DB スプレッドシートを、ほかのアプリ（貯蔵品の VPS 版）が DeptApprovers の読み先にしていた。最初に、ほかのアプリのリポジトリを DB の ID で grep する。
- DB スプレッドシートの時間帯が GMT（Sheets API の properties.timeZone は "Etc/GMT"、地域は en_US）だった。Apps Script の `getSpreadsheetTimeZone()` はこのとき空の文字を返す。GAS の返事だけで時間帯を照合すると必ず不合格になるので、GMT 系の設定のときは空を同じ値として扱い、本当の値は Sheets API で確かめる。
- 外の仕組み（Sebastian）が、Sheets API で Requests に新しい申請の行を直接足していた（`values.append`、USER_ENTERED、INSERT_ROWS）。GAS の doPost だけを見ていると気づかない。移す前に、シートに直接書いている仕組みが無いかを、その仕組みのコードで確かめる（DB の ID で grep）。VPS 側に「シートからの取り込み」を足して、外の仕組みは変えずに済ませた。
- 移行の道具が本番だけを移す元として受け付ける作りだと、ステージングにも本番の行が入る。本番の PDF の ID などは、ステージングでは空にして入れる（ステージングの GAS もオーナーの権限で動くため、本番のファイルを捨てられる）。
- worker を `python -m app.worker` で起動すると、同じファイルが `__main__` と `app.worker` の 2 つとして読まれる。各モジュールの `register_worker()` は `app.worker` 側に予定を登録するので、`__main__` 側で回すと予定が空になる（ステージングで、1 分ごとの取り込み・gas_ping・行の反映の予定が一度も動いていなかった）。試験は部品として読むので気づけない。`if __name__ == "__main__": from app import worker; worker.main()` とし、起動ログの schedules= の一覧を必ず目で見る。貯蔵品・工番部品の worker も同じ作りなら確かめる。
