# Google フォーム連携仕様書（確定版）

部署 prefill ＋ 作業員候補絞り込みを実現するための、部署別フォーム自動生成方式の確定仕様。

---

## 目次

1. [方式概要](#1-方式概要)
2. [Web アプリ導線](#2-web-アプリ導線)
3. [フォーム生成方式（方式A：テンプレ複製＋キャッシュ）](#3-フォーム生成方式方式aテンプレ複製キャッシュ)
4. [フォーム項目仕様](#4-フォーム項目仕様)
5. [工番の扱い](#5-工番の扱い)
6. [DB シート追加（FormTemplates / FormMap）](#6-db-シート追加formtemplatesformmap)
7. [フォーム URL 決定ロジック（API）](#7-フォーム-url-決定ロジックapi)
8. [マスタ更新時のフォーム選択肢更新](#8-マスタ更新時のフォーム選択肢更新)
9. [onFormSubmit → Requests 格納マッピング](#9-onformsubmit--requests-格納マッピング)
10. [後続処理への影響](#10-後続処理への影響)
11. [GAS 関数一覧](#11-gas-関数一覧)

---

## 1. 方式概要

### 背景・なぜこの方式が必要か

- 部署を prefill で固定したい
- 作業員プルダウンを部署で絞りたい
- **Google フォームは「開くたびに選択肢を動的に変える」ことができない**
- → 部署ごとに作業員候補が違うため、フォームは部署単位で分ける必要がある

### 確定方式

**方式A：テンプレ複製 ＋ FormMap キャッシュ**

- テンプレフォームを 2 つ用意（残業 / 休日）
- (種別, 部署) の組み合わせで初回アクセス時にテンプレを複製
- 2 回目以降は FormMap のキャッシュ URL に即リダイレクト

---

## 2. Web アプリ導線

### トップ画面（/）

```
「残業申請」ボタン押下
  └→ 部署選択モーダル表示（ページ遷移なし・即表示）
      └→ 部署ボタンを 1 クリック → 対応フォームへ即遷移

「休日出勤申請」ボタン押下
  └→ 部署選択モーダル表示（同上）
      └→ 部署ボタンを 1 クリック → 対応フォームへ即遷移
```

### ポイント

- 部署選択は **別ページではなくモーダル**（ページ切替の遅さを解消）
- 「すぐ開く」体感を維持
- 部署ボタンの一覧は部署マスタから動的取得

---

## 3. フォーム生成方式（方式A：テンプレ複製＋キャッシュ）

### テンプレフォーム

| テンプレ名 | type | 用途 |
|------------|------|------|
| OvertimeTemplate | overtime | 残業申請フォームの雛形 |
| HolidayTemplate | holiday | 休日出勤申請フォームの雛形 |

### 生成フロー

```
1. FormMap から (type, dept) を検索
2. レコードが存在する → formUrl を返す（生成しない）
3. レコードが存在しない → 新規生成:
   a. FormTemplates から該当テンプレを取得
   b. テンプレフォームを DriveApp.getFileById().makeCopy() で複製
   c. タイトル設定:
      - 残業: 【残業申請】{部署名}
      - 休日: 【休日申請】{部署名}
   d. 説明文: 運用注意を差し込み（固定テンプレ）
   e. 質問の選択肢を更新:
      - 作業員: 作業員マスタを dept でフィルタ → choice 設定
      - 業務ID: 業務NOマスタを dept でフィルタ → choice 設定
      - 理由: 固定 choice（理由リスト）
      - 予定時間: 種別に応じた choice
   f. FormMap に formId / formUrl / updatedAt / isActive=true を保存
4. formUrl を返す
```

### prefill の代替手法（確定）

prefill URL を毎回生成するのではなく、**選択肢を 1 つだけにして実質固定** にする:

- **部署**: 選択肢をその部署名 1 つだけにする → 入力者が変更不可
- **申請種別**: 選択肢を「残業」or「休日」1 つだけにする → 同上

この方式は prefill より堅牢（URL 改ざんで変更されるリスクがない）。

---

## 4. フォーム項目仕様

### 4.1 残業フォーム（質問順）

| # | 質問 | 入力形式 | 選択肢 / 制約 | 備考 |
|---|------|----------|---------------|------|
| 1 | 申請種別 | プルダウン（1択固定） | `残業` | 実質 prefill |
| 2 | 部署 | プルダウン（1択固定） | `{部署名}` | 実質 prefill |
| 3 | 作業員 | プルダウン | 部署で絞り込み。表示例: `A001 今泉雄二` | 作業員マスタ(在籍フラグ=true)から生成 |
| 4 | 作業実施日 | 日付 | ― | Date 型 |
| 5 | 工番 | プルダウン | 工番マスタ全件。表示: `工番｜品名／納入先` | 件数が多い場合は絞り込み列追加を検討 |
| 6 | 業務ID | プルダウン | 部署で絞り込み | 業務NOマスタから生成 |
| 7 | 理由 | プルダウン | 固定リスト（下記参照） | ― |
| 8 | 補足理由 | 段落（長文） | ― | 理由=「その他」の時のみ必須 |
| 9 | 予定時間 | プルダウン | `0.5` / `1.0` / `1.5` / `2.0` / `2.5` / `3.0` / `3.5` / `4.0` | 単位: 時間 |

### 4.2 休日フォーム（質問順）

| # | 質問 | 入力形式 | 選択肢 / 制約 | 備考 |
|---|------|----------|---------------|------|
| 1〜8 | *(残業フォームと同一)* | ― | ― | ― |
| 9 | 予定時間 | プルダウン | `半日` / `1日` | ― |

### 4.3 理由リスト（固定・GAS コード準拠）

- 急ぎ作業を要するため
- 納期遅延解消のため
- マシントラブル対応のため
- 業者対応のため
- 顧客対応のため
- その他（→補足理由必須）

> ※ 「その他」選択時は補足理由（質問 8）が必須。

### 4.4 正規化ルール（DB 保持は分単位）

| 種別 | フォーム表示値 | approvedMinutes（分） |
|------|----------------|----------------------|
| 残業 | 0.5 | 30 |
| 残業 | 1.0 | 60 |
| 残業 | 1.5 | 90 |
| 残業 | 2.0 | 120 |
| 残業 | 2.5 | 150 |
| 残業 | 3.0 | 180 |
| 残業 | 3.5 | 210 |
| 残業 | 4.0 | 240 |
| 休日 | 半日 | 240 |
| 休日 | 1日 | 480 |

---

## 5. 工番の扱い

### V1（確定：プルダウン）

- フォーム入力は **プルダウン**（工番マスタ全件から生成）
- 表示形式: `工番｜品名／納入先`（品名・納入先がある場合）
- 部署による絞り込みは V1 では行わない（工番マスタに部署列がないため）

### スケーラビリティ注意

工番が数千〜万件になるとフォームのプルダウンが重くなる。対策として工番マスタに以下の列追加を推奨:

| 追加列候補 | 目的 |
|------------|------|
| isActive | 稼働中の工番のみ TRUE → 候補を絞り込み |
| validTo | この日付を過ぎたら候補から除外 |
| dept | 部署で絞れるようにする（将来） |

まずは現行件数で動作確認し、重ければ絞り込み列を追加する。

### 将来拡張

| 段階 | 方式 |
|------|------|
| V1（確定） | プルダウン（全件） |
| 拡張 | isActive / validTo で絞り込み |
| 上位拡張 | Web アプリ側で工番検索 → prefill してフォームへ渡す |

---

## 6. DB シート追加（FormTemplates / FormMap）

> 詳細カラム定義は [spreadsheet-structure.md](./spreadsheet-structure.md) を参照。

### 6.1 FormTemplates

テンプレートフォーム（残業/休日の 2 レコード）を管理。

| type | templateFormId | templateUrl | note |
|------|----------------|-------------|------|
| overtime | *(ID)* | *(URL)* | 残業テンプレ |
| holiday | *(ID)* | *(URL)* | 休日テンプレ |

### 6.2 FormMap

(type, dept) 単位で自動生成されたフォームを管理。

| type | dept | formId | formUrl | updatedAt | isActive |
|------|------|--------|---------|-----------|----------|
| overtime | 製造部 | *(ID)* | *(URL)* | 2026-01-15T06:30:00 | TRUE |
| holiday | 製造部 | *(ID)* | *(URL)* | 2026-01-15T06:30:00 | TRUE |
| ... | ... | ... | ... | ... | ... |

---

## 7. フォーム URL 決定ロジック（API）

### 関数: `getOrCreateDeptForm(type, dept)`

| 項目 | 内容 |
|------|------|
| **入力** | `type`: `overtime` or `holiday`、`dept`: 部署名（部署マスタの値） |
| **出力** | `formUrl`（部署専用フォーム URL） |

### 処理手順（確定）

```
function getOrCreateDeptForm(type, dept) {
  // 1. FormMap から (type, dept) を検索
  //    → あれば formUrl を返す

  // 2. なければ生成
  //    a. FormTemplates から該当テンプレを取得
  //    b. テンプレフォームを複製
  //    c. タイトル設定: 【残業申請】{dept} or 【休日申請】{dept}
  //    d. 説明文設定: 固定テンプレ
  //    e. 選択肢更新:
  //       - 申請種別: 1択固定（残業 or 休日）
  //       - 部署: 1択固定（dept）
  //       - 作業員: 作業員マスタを dept でフィルタ（在籍フラグ=true）
  //       - 業務ID: 業務NOマスタを dept でフィルタ
  //       - 理由: 固定リスト
  //       - 予定時間: 種別に応じた choice
  //    f. FormMap に保存 (formId, formUrl, updatedAt, isActive=true)

  // 3. formUrl を返す
}
```

---

## 8. マスタ更新時のフォーム選択肢更新

部署異動・人員追加があるため、既存フォームの選択肢を定期更新する必要がある。

### 方針（確定：毎朝バッチ）

| 項目 | 内容 |
|------|------|
| **タイミング** | 毎朝 6:30（時間トリガー） |
| **対象** | FormMap 上の全 (type, dept) で isActive=true のフォーム |
| **更新内容** | 作業員 choice / 業務 ID choice を最新マスタで上書き |
| **関数** | `nightlyUpdateAllForms()` |

### 処理フロー

```
function nightlyUpdateAllForms() {
  // 1. FormMap の全行（isActive=true）を取得
  // 2. 各行について updateDeptFormChoices(type, dept) を呼び出し
  // 3. updatedAt を更新
}

function updateDeptFormChoices(type, dept) {
  // 1. FormMap から formId を取得
  // 2. FormApp.openById(formId)
  // 3. 作業員マスタを dept でフィルタ → 作業員質問の choice を更新
  // 4. 業務NOマスタを dept でフィルタ → 業務ID質問の choice を更新
  // 5. FormMap.updatedAt を更新
}
```

### 補足

- 管理画面に「フォーム候補更新」ボタンを設けて手動実行も可能にする（任意）
- 部署マスタに新部署が追加された場合、その部署のフォームは初回アクセス時に自動生成されるため事前対応不要

---

## 9. onFormSubmit → Requests 格納マッピング

### 9.1 フォーム項目 → Requests 列のマッピング（確定）

| フォーム項目 | Requests 列 | 変換・補完 |
|-------------|-------------|-----------|
| 申請種別（残業/休日） | requestType | `残業` → `overtime` / `休日` → `holiday` |
| 部署 | dept | フォーム値を基本採用。作業員マスタの部署で上書き可（改ざん対策） |
| 作業員（`A001 今泉雄二`） | workerCode / workerName | 先頭コード抽出 → マスタから氏名補完（表示文字列から氏名抽出でも可） |
| 作業実施日 | targetDate | `yyyy-MM-dd` に正規化 |
| 工番 | orderNo1 | V1 は 1 件のみ。将来 3 件対応（orderNo1〜3） |
| 業務ID（業務NO） | jobId1 / workNo1 | 業務 ID を保存し、表示用に業務 NO/名を補完 |
| 理由 | reason | 文字列そのまま保存 |
| 補足理由 | reasonDetail | reason=「その他」のとき必須。※ Requests に `reasonDetail` 列追加を推奨 |
| 予定時間（残業: 0.5〜4.0） | approvedMinutes | 時間 → 分に換算（×60） |
| 予定時間（休日: 半日/1日） | approvedMinutes | 半日=240 / 1日=480 |

### 9.2 GAS が自動で埋める列（確定）

| Requests 列 | 値 |
|-------------|------|
| requestId | UUID（`Utilities.getUuid()`） |
| status | `submitted` |
| submittedAt | `new Date()`（現在日時） |
| workerEmail | 作業員マスタから workerCode で引き当て |

### 9.3 作業員コード抽出ロジック

```javascript
// フォーム回答例: "A001 今泉雄二"
// → workerCode = "A001", workerName = "今泉雄二"
const answer = "A001 今泉雄二";
const workerCode = answer.split(" ")[0];  // "A001"
// マスタから氏名・メール・部署を補完
```

---

## 10. 後続処理への影響

### 10.1 トップ表示

- Requests に 1 行追加された瞬間（status=submitted）から表示
- **変更なし**

### 10.2 承認

- 承認ボタンで status=approved に変更
- **変更なし**

### 10.3 実績（完了ボタン）

- 残業: 完了のみ入力（17:20 固定起点）
- 休日: 開始/完了
- **変更なし**

### 10.4 PDF 生成

- **完了ボタン押下時に PDF 生成**
- 条件: `status = approved` を必須（承認前の PDF 乱発防止）
- 未承認で完了が押された場合: WorkLogs のみ更新し、承認後に朝バッチで PDF 生成（保険）

### 10.5 夕方メール（2 回）

- 対象: `status = approved` のみ
- 本文に含める情報:
  - 部署名・作業員名・予定時間
  - APP_URL

### 10.6 朝メール

- 実績（netMinutes）一覧を CSV/Excel 添付
- PDF 作成件数も本文に記載

---

## 11. GAS 関数一覧

### 11.1 フォーム生成・送信

| 関数名 | トリガー | 概要 |
|--------|----------|------|
| `getOrCreateDeptForm_(type, dept)` | Web アプリから呼び出し | FormMap 検索 → 未生成ならテンプレ複製＋onFormSubmitトリガー付与 → formUrl 返却 |
| `updateDeptFormChoices_(type, dept)` | バッチ / 手動 | 指定 (type, dept) のフォームの選択肢を最新マスタで更新 |
| `nightlyUpdateAllForms_()` | 時間トリガー（毎朝 6:30） | FormMap 全件の選択肢を一括更新 |
| `addFormSubmitTrigger_(formId)` | フォーム生成時に自動呼出 | 新規フォームに handleFormSubmit_ トリガーを付与（重複防止付き） |
| `handleFormSubmit_(e)` | フォーム送信トリガー（installable） | フォーム回答 → Requests 書き込み＋WorkLogs プレースホルダ作成 |
| `appendRequestRow_(obj)` | handleFormSubmit_ から呼出 | ヘッダ名ベースで Requests に 1 行追加（列順変更耐性あり） |
| `ensureWorkLogRow_(requestId)` | handleFormSubmit_ から呼出 | WorkLogs にプレースホルダ行を作成 |
| `lookupOrderInfo_(orderNo)` | handleFormSubmit_ から呼出 | 工番マスタから受注先/納入先/品名を補完 |

### 11.2 承認

| 関数名 | トリガー | 概要 |
|--------|----------|------|
| `api_getTodayRequestsForDept(dept)` | Web アプリ（承認者画面） | 部署別の本日申請一覧を取得（権限チェック付き） |
| `api_approveRequest(requestId)` | Web アプリ（承認ボタン） | status→approved、approvedAt/approvedBy を記録 |
| `isAdmin_(email)` | 内部 | ApproverMap で admin ロールかチェック |
| `canApproveDept_(email, dept)` | 内部 | 指定部署の承認権限があるかチェック |

### 11.3 実績（開始/完了）

| 関数名 | トリガー | 概要 |
|--------|----------|------|
| `api_markOvertimeDone(requestId)` | Web アプリ（残業完了ボタン） | 17:20 固定起点で実績算出→休憩控除→net→WorkLogs更新→PDF生成 |
| `api_markHolidayStart(requestId)` | Web アプリ（休日開始ボタン） | actualStartAt を記録 |
| `api_markHolidayDone(requestId)` | Web アプリ（休日完了ボタン） | start〜end で実績算出→休憩控除→net→WorkLogs更新→PDF生成 |
| `calcBreakMinutesByMaster_(type, actualMinutes)` | 内部 | 休憩マスタから該当区間の休憩分を返す |
| `getRequestById_(requestId)` | 内部 | Requests から 1 件取得 |
| `updateWorkLog_(requestId, patch)` | 内部 | WorkLogs の指定行をパッチ更新 |

### 11.4 PDF 生成

| 関数名 | トリガー | 概要 |
|--------|----------|------|
| `generatePdfForRequest_(requestId)` | 完了ボタンから呼出 | テンプレSS複製→操作!B3にrequestIdセット→申請書フォームをPDF化→Drive保存 |
| `getOrCreateDateFolder_(rootFolderId, dateObj)` | 内部 | yyyy.MM.dd フォルダを取得 or 作成 |
| `exportSheetToPdfBlob_(ssId, sheetId, filename)` | 内部 | 指定シートを A4 PDF Blob にエクスポート |

### 11.5 Web アプリ / メニュー

| 関数名 | トリガー | 概要 |
|--------|----------|------|
| `doGet(e)` / `doPost(e)` | Web アプリ | 部署選択モーダル表示、フォーム URL リダイレクト |
| `onOpen()` | スプレッドシート起動 | 管理者メニュー（フォーム全更新 / テスト生成） |
| `setupTriggers_()` | 手動 1 回実行 | nightlyUpdateAllForms_ の毎朝トリガーを作成 |

---

## 12. ソースファイル構成

```
src/
├── config.gs          # 定数・設定・ユーティリティ（SHEET, Q, REASONS, getSheetHeaderIndex_）
├── masters.gs         # マスタ読込（loadDeptList_, loadWorkersByDept_, loadJobsByDept_, loadOrderChoices_）
├── formMap.gs         # FormMap CRUD + FormTemplates 取得
├── formUtils.gs       # フォーム質問検索・選択肢設定ヘルパー
├── formGenerate.gs    # 核心: getOrCreateDeptForm_ + updateDeptFormChoices_
├── formBatch.gs       # 毎朝バッチ: nightlyUpdateAllForms_
├── formSubmit.gs      # onFormSubmit ハンドラ: handleFormSubmit_ + Requests/WorkLogs 書き込み
├── approval.gs        # 承認権限チェック + 承認実行 API
├── worklog.gs         # 開始/完了ボタン + 休憩控除 + net 算出
├── pdfExport.gs       # PDF 生成（テンプレSS複製→操作!B3→PDF→Drive保存）
└── menu.gs            # onOpen メニュー + setupTriggers_
```

### Settings 必須キー（PDF 生成用）

| Key | 説明 |
|-----|------|
| `PDF_ROOT_FOLDER_ID` | PDF 保存先ルートフォルダの Google Drive ID |
| `TEMPLATE_SSID` | 申請書テンプレートのスプレッドシート ID（操作/申請書フォーム シートを含む） |
