# スプレッドシート構造定義書

残業・休日出勤申請システムで使用する Google スプレッドシートの全シート構成を定義する。

---

## 目次

1. [操作](#1-操作)
2. [README](#2-readme)
3. [Settings](#3-settings)
4. [部署マスタ](#4-部署マスタ)
5. [作業員マスタ](#5-作業員マスタ)
6. [業務NOマスタ](#6-業務noマスタ)
7. [工番マスタ](#7-工番マスタ)
8. [休憩マスタ](#8-休憩マスタ)
9. [ApproverMap](#9-approvermap)
10. [Requests](#10-requests)
11. [WorkLogs](#11-worklogs)
12. [ExportHistory](#12-exporthistory)
13. [FormTemplates](#13-formtemplates)
14. [FormMap](#14-formmap)
15. [申請書フォーム](#15-申請書フォーム)
16. [休憩計算_フォーム](#16-休憩計算_フォーム)

---

## 1. 操作

UI 操作用シート（ボタン配置等）。システム利用者が直接触るダッシュボード的な役割。

---

## 2. README

スプレッドシートの使い方・注意事項を記載した説明シート。

---

## 3. Settings

システム全体の設定値を **Key / Value** 形式で管理するシート。

| 行 | Key | 説明 |
|----|-----|------|
| ― | 各種設定キー | Web アプリ URL、PDF 出力先フォルダ ID、メール送信先など |

> ※ GAS から `getSettings()` 等で読み取り、各機能の動作パラメータとして使用する想定。

---

## 4. 部署マスタ

別マスタから自動転記する想定。ここにある部署名をフォーム選択肢や権限判定に利用。

| 列 | ヘッダー (Row 2) | 型 | 説明 |
|----|-------------------|-----|------|
| A | 部署 | String | 部署名 |

- **データ開始行**: Row 3

---

## 5. 作業員マスタ

別マスタから自動転記想定。Google アカウントは Web アプリの本人判定に使用。

| 列 | ヘッダー (Row 2) | 型 | 説明 |
|----|-------------------|-----|------|
| A | 作業員コード | String | 作業員の一意コード |
| B | 氏名 | String | 作業員名 |
| C | 部署 | String | 所属部署（部署マスタと一致） |
| D | 担当業務ID（カンマ区切り） | String | 担当する業務IDをカンマ区切りで列挙 |
| E | Googleアカウント（メール） | String | ログイン判定に使用する Gmail アドレス |
| F | 在籍フラグ | Boolean/String | 在籍中かどうか |

- **データ開始行**: Row 3

---

## 6. 業務NOマスタ

別マスタから自動転記想定。フォームの選択肢・申請書明細に使用。

| 列 | ヘッダー (Row 2) | 型 | 説明 |
|----|-------------------|-----|------|
| A | 業務ID | String | 業務の一意 ID |
| B | 業務NO | String | 業務番号（表示用） |
| C | 業務名 | String | 業務名称 |
| D | 部署 | String | 対象部署 |
| E | 説明 | String | 業務の説明・備考 |

- **データ開始行**: Row 3

---

## 7. 工番マスタ

別マスタから自動転記想定。フォームの選択肢・申請書明細に使用。

| 列 | ヘッダー (Row 2) | 型 | 説明 |
|----|-------------------|-----|------|
| A | 工番 | String | 工番コード |
| B | 受注先 | String | 受注先名 |
| C | 納入先 | String | 納入先名 |
| D | 納入先住所 | String | 納入先の住所 |
| E | 品名 | String | 製品・品目名 |
| F | 数量 | Number | 数量 |
| G | 取込日時 | DateTime | マスタ取込日時 |

- **データ開始行**: Row 3

---

## 8. 休憩マスタ

実績計算用。実働分（開始〜完了）に応じて控除する休憩分を定義。上限が空欄の場合は無限大として扱う。

| 列 | ヘッダー (Row 2) | 型 | 説明 |
|----|-------------------|-----|------|
| A | 適用区分 | String | `overtime` または `holiday` |
| B | 実働分_下限 | Number | この区間の下限（分） |
| C | 実働分_上限 | Number | この区間の上限（分）。空欄 = 上限なし |
| D | 休憩分 | Number | 控除する休憩時間（分） |
| E | 備考 | String | 補足説明 |

- **データ開始行**: Row 3（Row 3 はヘッダー行の場合あり、実データは Row 4〜）

### サンプルデータ

| 適用区分 | 実働分_下限 | 実働分_上限 | 休憩分 | 備考 |
|----------|-------------|-------------|--------|------|
| holiday | 0 | 240 | 0 | 4h未満は休憩なし（例） |
| holiday | 240 | 360 | 30 | 4h以上〜6h未満は30分（例） |
| holiday | 360 | *(空欄)* | 45 | 6h以上は45分（例） |
| overtime | 0 | *(空欄)* | 0 | 残業は起点固定(17:20)のため休憩控除なし（必要なら定義） |

---

## 9. ApproverMap

部署 × Google アカウントで承認画面の入室可否と承認権限を管理。

| 列 | ヘッダー (Row 2) | 型 | 説明 |
|----|-------------------|-----|------|
| A | 部署 | String | 対象部署 |
| B | 承認者メール | String | 承認者の Google アカウント |
| C | role | String | `approver` または `admin` |
| D | 有効フラグ | Boolean/String | この承認者が有効かどうか |

- **データ開始行**: Row 3

---

## 10. Requests

申請データの本体。1 行 = 1 申請。

| 列 | ヘッダー (Row 1) | 型 | 説明 |
|----|-------------------|-----|------|
| A | requestId | String | 申請の一意 ID |
| B | requestType | String | `overtime`（残業）/ `holiday`（休日出勤） |
| C | status | String | `submitted` / `approved` / `canceled` |
| D | dept | String | 申請者の部署 |
| E | workerCode | String | 作業員コード |
| F | workerName | String | 作業員氏名 |
| G | workerEmail | String | 作業員メールアドレス |
| H | targetDate | Date | 対象日（残業日 or 休日出勤日） |
| I | submittedAt | DateTime | 申請日時 |
| J | approvedAt | DateTime | 承認日時 |
| K | approvedBy | String | 承認者メールアドレス |
| L | approvedMinutes | Number | 承認された時間（分） |
| M | reason | String | 申請理由 |
| N | workContent | String | 作業内容 |
| O | jobId1 | String | 業務ID 1 |
| P | jobId2 | String | 業務ID 2 |
| Q | jobId3 | String | 業務ID 3 |
| R | workNo1 | String | 工番 1 |
| S | workNo2 | String | 工番 2 |
| T | workNo3 | String | 工番 3 |
| U | orderNo1 | String | 受注先 1 |
| V | orderNo2 | String | 受注先 2 |
| W | orderNo3 | String | 受注先 3 |
| X | customer1 | String | 納入先 1 |
| Y | customer2 | String | 納入先 2 |
| Z | customer3 | String | 納入先 3 |
| AA | product1 | String | 品名 1 |
| AB | product2 | String | 品名 2 |
| AC | product3 | String | 品名 3 |
| AD | hrMailSentAt | DateTime | 人事部門へのメール送信日時 |
| AE | pdfGeneratedAt | DateTime | PDF 生成日時 |
| AF | pdfFileId | String | 生成した PDF の Google Drive ファイル ID |
| AG | pdfFolderId | String | PDF 保存先フォルダ ID |
| AH | exportError | String | エクスポート時エラーメッセージ |

- **データ開始行**: Row 2
- **備考**: 業務・工番・受注先・納入先・品名は最大 3 セットまで登録可能

---

## 11. WorkLogs

実績データ。残業は `actualEndAt` のみ入力（開始は 17:20 固定）。休日出勤は開始・終了とも入力。

| 列 | ヘッダー (Row 2) | 型 | 説明 |
|----|-------------------|-----|------|
| A | requestId | String | 紐づく申請 ID（Requests.requestId） |
| B | actualStartAt | DateTime | 実績開始日時 |
| C | actualEndAt | DateTime | 実績終了日時 |
| D | actualMinutes | Number | 実績時間（分）= 終了 − 開始 |
| E | breakMinutes | Number | 休憩時間（分）（休憩マスタから自動算出） |
| F | netMinutes | Number | 正味時間（分）= actualMinutes − breakMinutes |
| G | updatedAt | DateTime | 更新日時 |
| H | updatedBy | String | 更新者メールアドレス |

- **データ開始行**: Row 3
- **注記 (Row 1)**: ※実績。残業は actualEndAt のみ入力。（開始は 17:20 固定）

---

## 12. ExportHistory

PDF 出力・メール送信等のエクスポート履歴を記録。

| 列 | ヘッダー (Row 1) | 型 | 説明 |
|----|-------------------|-----|------|
| A | exportId | String | エクスポートの一意 ID |
| B | requestId | String | 対象の申請 ID |
| C | exportType | String | エクスポート種別（`pdf` / `mail` 等） |
| D | exportedAt | DateTime | エクスポート実行日時 |
| E | exportedBy | String | 実行者メールアドレス |
| F | fileId | String | 生成ファイルの Google Drive ID |
| G | status | String | 成功/失敗のステータス |
| H | errorMessage | String | エラー時のメッセージ |

- **データ開始行**: Row 2

---

## 13. FormTemplates

Google フォームのテンプレート管理シート。残業・休日それぞれのテンプレートフォーム ID を保持する。部署別フォーム自動生成時にこのテンプレートを複製元として使用する。

| 列 | ヘッダー (Row 1) | 型 | 説明 |
|----|-------------------|-----|------|
| A | type | String | `overtime` または `holiday` |
| B | templateFormId | String | テンプレートとなる Google フォームの ID |
| C | templateUrl | String | テンプレートフォームの URL |
| D | note | String | 備考（例：残業テンプレ / 休日テンプレ） |

- **データ開始行**: Row 2

### 初期データ

| type | templateFormId | templateUrl | note |
|------|----------------|-------------|------|
| overtime | *(フォームID)* | *(URL)* | 残業テンプレ |
| holiday | *(フォームID)* | *(URL)* | 休日テンプレ |

---

## 14. FormMap

部署 × 種別ごとに自動生成されたフォームの管理シート。`getOrCreateDeptForm()` で生成・参照される。

| 列 | ヘッダー (Row 1) | 型 | 説明 |
|----|-------------------|-----|------|
| A | type | String | `overtime` または `holiday` |
| B | dept | String | 部署名（部署マスタと一致） |
| C | formId | String | 生成されたフォームの ID |
| D | formUrl | String | 生成されたフォームの URL |
| E | updatedAt | DateTime | 最終更新日時（選択肢更新含む） |
| F | isActive | Boolean | 有効フラグ |

- **データ開始行**: Row 2
- **ユニークキー**: (type, dept) の組み合わせ
- **自動生成**: フォーム未生成の (type, dept) でアクセスがあった場合、FormTemplates のテンプレを複製して自動生成し、本シートに記録する

---

## 15. 申請書フォーム

PDF 出力用の申請書テンプレート。セル結合を多用したレイアウト定義。GAS から値を埋め込み、PDF として書き出す。

### レイアウト概要

- **上部**: 会社名・タイトル（「残業申請書」/「休日出勤届」）
- **申請者情報ブロック**: 申請日、部署名、作業員名
- **作業内容ブロック**: 対象日、作業時間（開始〜終了）、作業内容、理由
- **明細テーブル**: 業務NO・工番・受注先・納入先・品名（最大 3 行）
- **承認欄**: 承認者名、承認日
- **実績欄**: 実績開始〜終了、実働時間、休憩時間、正味時間

> ※ セル座標は実装時に GAS コードで直接参照するため、レイアウト変更時はコード側も修正が必要。

---

## 16. 休憩計算_フォーム

休憩時間の自動計算ロジックを確認・テストするための補助シート。

### 構成

- **入力エリア**: 適用区分（overtime/holiday）、実働時間（分）を入力
- **計算結果エリア**: 休憩マスタを参照し、該当する休憩分を自動表示
- **参照テーブル**: 休憩マスタの内容を転記・参照

> ※ 数式ベースの計算シート。GAS から直接使用するのではなく、休憩マスタの定義が正しいか手動検証する目的。

---

## シート間リレーション

```
作業員マスタ.部署  ──→  部署マスタ.部署
作業員マスタ.担当業務ID  ──→  業務NOマスタ.業務ID
ApproverMap.部署  ──→  部署マスタ.部署
Requests.workerCode  ──→  作業員マスタ.作業員コード
Requests.dept  ──→  部署マスタ.部署
Requests.jobId1〜3  ──→  業務NOマスタ.業務ID
Requests.workNo1〜3  ──→  工番マスタ.工番
WorkLogs.requestId  ──→  Requests.requestId
ExportHistory.requestId  ──→  Requests.requestId
休憩マスタ  ──→  WorkLogs（休憩分の自動算出に使用）
FormMap.type  ──→  FormTemplates.type（テンプレ複製元）
FormMap.dept  ──→  部署マスタ.部署
FormMap（生成時）  ──→  作業員マスタ（部署絞り込みで選択肢セット）
FormMap（生成時）  ──→  業務NOマスタ（部署絞り込みで選択肢セット）
```

---

## 補足

- **Row 1 が注記、Row 2 がヘッダーのシート**: 部署マスタ、作業員マスタ、業務NOマスタ、工番マスタ、休憩マスタ、ApproverMap、WorkLogs
- **Row 1 がヘッダーのシート**: Requests、ExportHistory、FormTemplates、FormMap
- マスタ系シート（部署・作業員・業務NO・工番）は別マスタからの自動転記を想定しており、手動編集は原則不要
- `requestType` は `overtime`（残業）と `holiday`（休日出勤）の 2 種類
- 残業の開始時刻は **17:20 固定**（定時終了時刻）
