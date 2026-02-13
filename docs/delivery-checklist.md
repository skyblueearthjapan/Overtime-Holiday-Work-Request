# 納品チェックリスト（〇×方式）

> 実装完了後に実施する受け入れ確認用。〇×で確認するだけで完了判定できる。

---

## 0. 事前情報・成果物確認

- [ ] DB スプレッドシート ID が確定している
- [ ] PDF テンプレスプレッドシート ID が確定している
- [ ] PDF 保存先フォルダ ID（`PDF_ROOT_FOLDER_ID`）が確定している
- [ ] 総務宛先（`HR_MAIL_TO`）が Settings に設定済み（カンマ区切り複数可）
- [ ] Web アプリ URL（`APP_URL`）が Settings に設定済み
- [ ] スコープ（Drive/Form/Gmail/UrlFetch）が有効で認可済み
- [ ] シート名・ヘッダ名がテンプレ通りで変更されていない

---

## 1. DB 構造（スプレッドシート）チェック

### 1-1. 必須シート存在（コードが参照する正式名）

- [ ] Settings（設定値 Key/Value）
- [ ] Requests（申請データ）
- [ ] WorkLogs（実績データ、ヘッダ行=2行目）
- [ ] 部署マスタ（`SHEET.DEPTS`）
- [ ] 作業員マスタ（`SHEET.WORKERS`）
- [ ] 業務NOマスタ（`SHEET.JOBS`）
- [ ] 工番マスタ（`SHEET.ORDERS`）
- [ ] 休憩マスタ（適用区分/実働分_下限/実働分_上限/休憩分）
- [ ] ApproverMap（部署 × 承認者メール × role × 有効フラグ）
- [ ] FormTemplates（テンプレフォーム ID 管理）
- [ ] FormMap（部署×種別フォーム URL キャッシュ）

### 1-2. Requests 主要列（存在チェック）

- [ ] requestId
- [ ] status(submitted/approved/canceled)
- [ ] requestType(overtime/holiday)
- [ ] dept
- [ ] workerCode
- [ ] workerName
- [ ] workerEmail
- [ ] targetDate
- [ ] approvedMinutes
- [ ] submittedAt
- [ ] approvedAt
- [ ] approvedBy
- [ ] pdfFileId
- [ ] pdfGeneratedAt
- [ ] pdfFolderId

### 1-3. WorkLogs 主要列（存在チェック、ヘッダ行=2行目）

- [ ] requestId
- [ ] actualStartAt（休日用）
- [ ] actualEndAt（残業・休日）
- [ ] actualMinutes
- [ ] breakMinutes
- [ ] netMinutes
- [ ] updatedAt
- [ ] updatedBy

### 1-4. 作業員マスタ 主要列

- [ ] 作業員コード
- [ ] 氏名
- [ ] 部署
- [ ] Googleアカウント（メール）
- [ ] 在籍フラグ

### 1-5. ApproverMap 主要列

- [ ] 部署
- [ ] 承認者メール
- [ ] role(approver/admin)
- [ ] 有効フラグ

---

## 2. 申請（Google フォーム連携）チェック

### 2-1. フォーム仕様

- [ ] 残業フォームが存在し、工番はプルダウン
- [ ] 休日フォームが存在し、工番はプルダウン
- [ ] 予定時間の値が仕様通り（残業 0.5 刻み、休日 = 半日/1 日）
- [ ] 理由の選択肢が仕様通り（その他は補足理由必須）
- [ ] テンプレフォーム ID が FormTemplates に登録済み

### 2-2. onFormSubmit の動作

- [ ] フォーム送信で Requests に 1 行追加される
- [ ] requestId（UUID）が必ず入る
- [ ] status が submitted で登録される
- [ ] approvedMinutes が「分」で正規化されている
- [ ] submittedAt がタイムスタンプで入る
- [ ] WorkLogs にプレースホルダ行が作成される
- [ ] TOP に表示できる最小情報（部署・氏名・日付・種別・予定）が揃う

---

## 3. Web アプリ TOP 画面（作業者側）チェック

### 3-1. UI 要素

- [ ] 作業者情報がGoogleアカウントで自動表示される（部署/氏名/コード）
- [ ] 「残業申請」ボタンがある（押すと部署対応フォームへ遷移）
- [ ] 「休日出勤申請」ボタンがある（同上）
- [ ] 「承認者画面」ナビリンクがある
- [ ] 「総務部画面」ナビリンクがある
- [ ] 本日の申請一覧が表示される（種別/承認/予定/実績/休憩/実残業/PDF/操作）
- [ ] 各行に「承認待ち/承認済」ラベルが表示される

### 3-2. 実績ボタン

- [ ] 残業：承認済みで完了ボタンが表示される
- [ ] 休日：承認済みで開始ボタン → 開始済みで完了ボタンが表示される
- [ ] 完了済み → 「完了済」ラベルに切り替わる
- [ ] ボタン押下で WorkLogs にタイムスタンプが記録される
- [ ] PDF 生成済みなら Drive 直リンクが表示される

---

## 4. 承認者画面チェック（部署別アクセス制御）

### 4-1. アクセス制御

- [ ] ApproverMap にある承認者のみ部署が表示される
- [ ] admin ロールは全部署が表示される
- [ ] 承認者は「自部署」の一覧だけ見える

### 4-2. 承認処理

- [ ] 承認ボタン押下で Requests.status が approved になる
- [ ] approvedAt / approvedBy が記録される
- [ ] 一括承認ボタンで未承認全件を一括承認できる
- [ ] TOP 画面の該当行ラベルが承認済みに切り替わる

### 4-3. KPI 表示

- [ ] 本日の申請件数が表示される
- [ ] 未承認件数が表示される
- [ ] 承認済件数が表示される

---

## 5. 実績（WorkLogs）計算チェック

### 5-1. 残業

- [ ] 完了時刻が記録される
- [ ] 開始は 17:20 固定で actualStartAt に記録される
- [ ] 実績時間（actualMinutes）= end - start
- [ ] 休憩マスタから breakMinutes が差し引かれる
- [ ] netMinutes = actualMinutes - breakMinutes（0 未満にならない）

### 5-2. 休日

- [ ] 開始時刻（actualStartAt）が記録される
- [ ] 完了時刻（actualEndAt）が記録される
- [ ] breakMinutes が休憩マスタから差し引かれている
- [ ] netMinutes が正しく算出される（負にならない）

### 5-3. 本人チェック

- [ ] assertSelf_ により本人以外は操作不可

---

## 6. PDF 自動生成チェック（申請書）

### 6-1. トリガー条件

- [ ] 実績確定（完了ボタン）＋ status=approved を起点に PDF 生成が走る
- [ ] 未承認時は WorkLogs のみ更新（PDF は生成しない）

### 6-2. 転記・出力

- [ ] テンプレ SS を複製 → 操作!B3 に requestId セット → 再計算待ち
- [ ] 「申請書フォーム」シートを A4 PDF 化
- [ ] Drive に保存される
- [ ] 保存先：`PDF_ROOT_FOLDER_ID/YYYY.MM.DD` フォルダ（なければ自動作成）
- [ ] Requests.pdfFileId / pdfGeneratedAt / pdfFolderId が入る
- [ ] 一時コピー SS はゴミ箱に移動される
- [ ] 既に PDF 生成済みの場合は再生成しない（already フラグ）

---

## 7. 自動メールチェック（総務通知）

### 7-1. 夕方メール（2 回）

- [ ] 17 時台に送信される（承認時間＝予定時間で OK）
- [ ] 18 時台に送信される（同様）
- [ ] 本文に「部署」「氏名」「残業/休日」「承認時間」が部署別グループで列挙される
- [ ] APP_URL（アプリ URL）が本文に含まれる

### 7-2. 朝メール（翌朝）

- [ ] 7 時台に送信される
- [ ] 実績一覧が CSV + Excel（xlsx）の両方で添付される
- [ ] 本文に「残業 PDF 何件」「休日 PDF 何件」「合計何件」報告が含まれる
- [ ] 添付の実績は WorkLogs.netMinutes（実績）である

### 7-3. トリガー

- [ ] `setupAllTriggers_()` を実行して全トリガーが設定される
- [ ] フォーム更新（6:30）+ 夕方メール（17:10, 18:10）+ 朝メール（7:10）
- [ ] トリガーが重複しない（同名ハンドラ存在チェックあり）

---

## 8. 総務部管理画面（見える化 DX）チェック

### 8-1. アクセス制御

- [ ] admin ロールのみアクセス可能（`?page=admin`）

### 8-2. 月次 40h 監視（実績）

- [ ] 月次個人別合算（残業＋休日）が正しく集計される（netMinutes）
- [ ] 40h(2400分)/60h(3600分) ラインを基準に表示される
- [ ] 予測（ペース換算：累計/経過日×月日数）が算出される
- [ ] 注意対象（30h 超 or 予測 40h 超）が抽出される
- [ ] PDF 未作成件数が分かる（承認済み＋実績あり＋pdf なし）

### 8-3. グラフ（Canvas 自前描画、外部ライブラリ不使用）

- [ ] 個人別棒グラフが表示される（実績=濃色＋予測=薄色、40h/60h ライン）
- [ ] 部署別配分（円グラフ）が表示される
- [ ] 年度推移（折れ線、4/1〜3/31）が表示される

### 8-4. 特別条項（年 6 回）管理

- [ ] 年度内の「月 60h 超」回数が個人別に出る
- [ ] 5 回以上は警告、6 回は危険表示

### 8-5. 未承認滞留監視

- [ ] status=submitted の一覧が出る
- [ ] 48h 超が警告表示される

### 8-6. PDF リンク開封

- [ ] 日次テーブルから PDF をワンクリックで Drive 開封できる

### 8-7. フィルタ・検索・ソート・CSV

- [ ] 部署フィルタ / 氏名検索 / 40h超のみ / 60h超のみ / 予測40h超のみ / PDF未作成のみ
- [ ] 列ヘッダクリックでソート
- [ ] 表示中データを CSV エクスポート

---

## 9. パフォーマンス／安全性／運用チェック

- [ ] 主要処理は requestId で更新し、同一申請を重複生成しない
- [ ] 二重クリック対策（LockService でロック取得）がある
- [ ] Logger / try-catch エラーハンドリングが入っている
- [ ] 権限エラー時の表示が分かりやすい（日本語メッセージ）
- [ ] タイムゾーンが `Asia/Tokyo` で統一されている
- [ ] UI ボタンは処理中に disabled + 「処理中...」表示になる

---

## 最終受け入れテスト（必ずこの順で実施）

1. [ ] 残業フォーム送信 → Requests 登録 → TOP に表示
2. [ ] 承認者画面で承認 → TOP ラベルが「承認済」に変更
3. [ ] 残業完了ボタン → WorkLogs 更新 → netMinutes 算出
4. [ ] PDF 生成 → 日付フォルダに保存 → Requests に pdfFileId 記録
5. [ ] 休日フォーム送信 → 承認 → 開始 → 完了 → PDF 生成（休日フロー）
6. [ ] `sendEveningMail_()` 手動実行 → 文面 OK
7. [ ] `sendMorningMail_()` 手動実行 → CSV/Excel 添付 OK ＋ PDF 件数 OK
8. [ ] 総務画面で月次集計 ＝ WorkLogs と一致
9. [ ] 年度 60h 回数と未承認滞留が表示される

---

## 納品物チェック（提出物）

- [ ] GAS ソース一式（.gs 13 ファイル ＋ .html 4 ファイル）
- [ ] DB スプレッドシート（テンプレに従う全 16 シート）
- [ ] PDF テンプレートシート（操作!B3 → 申請書フォームの構成）
- [ ] 設定値一覧（Settings キー表：PDF_ROOT_FOLDER_ID, TEMPLATE_SSID, HR_MAIL_TO, APP_URL, ADMIN_EMAILS）
- [ ] トリガー設定手順（`setupAllTriggers_()` 実行手順）
- [ ] 運用手順（総務宛先変更方法、承認者追加方法、部署追加方法）

---

## GAS ソースファイル一覧（全 17 ファイル）

| ファイル | 役割 |
|----------|------|
| config.gs | 定数・設定・ユーティリティ |
| masters.gs | マスタ読込 + api_getWorkerInfo |
| formMap.gs | FormMap CRUD + api_getFormUrl |
| formUtils.gs | フォーム質問検索・選択肢設定 |
| formGenerate.gs | getOrCreateDeptForm_ + updateDeptFormChoices_ |
| formBatch.gs | nightlyUpdateAllForms_（毎朝6:30） |
| formSubmit.gs | handleFormSubmit_ + Requests/WorkLogs書込 |
| approval.gs | 承認権限 + api_approveRequest + api_getApproverDepts |
| worklog.gs | 開始/完了 + 休憩控除 + api_getTodayRequestsForWorker |
| pdfExport.gs | PDF生成（テンプレSS複製→Drive保存） |
| mail.gs | 夕方メール2回 + 朝メール（CSV/Excel添付） |
| adminApi.gs | 総務部API（日次/月次/年度/特別条項/滞留）+ doGet |
| menu.gs | onOpenメニュー + setupAllTriggers_ |
| admin.html | 総務部DX画面（KPI/Canvas/テーブル/フィルタ） |
| top.html | 作業者TOP画面（申請/一覧/完了ボタン） |
| approver.html | 承認者画面（部署選択/承認/一括承認） |
| no_auth.html | 権限エラー画面（フレンドリー表示） |
