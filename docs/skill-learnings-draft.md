# スキルに書き足す候補（残業・休日出勤申請の移行で分かったこと）

- clasp 3.4.1 の `update-deployment <id> -V N` は「Redeployed … @N」と出ても、版が変わらないことがある（説明文が英数字でも起きた。2026-11-26）。`list-deployments --json` で versionNumber を必ず確かめ、変わっていなければ Apps Script API `PUT /v1/projects/{scriptId}/deployments/{deploymentId}`（`deploymentConfig.versionNumber`）で当てる。clasp の認証ファイルのトークンで呼べる。
- 本番 GAS の時間主導トリガーは公開版ではなく HEAD のコードで動く。移行用の変更（移転モードなど）を master に入れて clasp push すると、その瞬間に本番の裏の処理が変わる。切り替えの合図まで別ブランチに置く。
- GCP コンソールは、時間がたつと本人確認のパスワードを求める（OAuth クライアント作成の画面）。オーナーが席を外す前に、GCP の作業を先に済ませる。
- 「スプレッドシートのコピーを作成」でできたサンドボックスは、Settings タブが本番の PDF フォルダ・台帳・フォーム・宛先を指したまま。コピー直後に Settings を試験用へ書き換え、フォームの対応表のタブは名前を変えて無効にする。Sheets の自動補完で値が変わることがあるので、書いた後に画面で確かめる。
- JEV ハンズの画面読みは ID を伏せ字にする。ID はボタンで写して PowerShell の Get-Clipboard で読むか、Drive API の検索で取る。
- 本番の DB スプレッドシートを、ほかのアプリ（貯蔵品の VPS 版）が DeptApprovers の読み先にしていた。最初に、ほかのアプリのリポジトリを DB の ID で grep する。
