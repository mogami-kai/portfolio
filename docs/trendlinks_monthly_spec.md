# Trendlinks実績報告（MNP/光）再設計 仕様書

## 1. 目的
LINEグループ投稿から、月次で以下を集計する。

1. 店舗別の `MNP件数` / `光件数`
2. 人別の `MNP件数` / `光件数`
3. 人別の `累計実績`（月内累計・全期間累計）
4. 人別の `日当たりMNP` / `日当たり光`

> 光件数は `1G光 + 10G光` の合算値とする。

---

## 2. 前提・設計方針
- レイヤーは **Ops + Admin の2層**で運用する（Middle承認は使わない）。
- 雑談メッセージは集計対象外にする。
- 既存LINE投稿のフォーマット揺れを許容しつつ、集計に必要な最小項目のみ抽出する。

---

## 3. 対象データ（入力）
LINE投稿本文（テキスト）から、下記項目を抽出する。

### 3.1 必須抽出項目
- 日付
- 店舗
- 名前
- 稼働日数（例: `稼働数(12日)`）
- 当日mnp件数
- 計mnp件数
- 当日1G光件数
- 計1G光件数
- 当日10G光件数
- 計10G光件数

### 3.2 派生項目
- `年月`（yyyy-MM）
- `光_当日 = 当日1G光 + 当日10G光`
- `光_計 = 計1G光 + 計10G光`

---

## 4. データモデル

## 4.1 Ops正規化テーブル: `LINE_SALES_DAILY`
1投稿 = 1レコードで保持する。

| 列名 | 型 | 説明 |
|---|---|---|
| message_id | string | LINEメッセージID（重複排除キー） |
| posted_at | datetime | Webhook受信日時 |
| report_date | date | 投稿内の日付 |
| ym | string | `yyyy-MM` |
| store_raw | string | 投稿内の店舗名（生） |
| person_raw | string | 投稿内の名前（生） |
| work_days | number | 稼働日数 |
| mnp_day | number | 当日mnp |
| mnp_cum | number | 計mnp |
| h1g_day | number | 当日1G光 |
| h1g_cum | number | 計1G光 |
| h10g_day | number | 当日10G光 |
| h10g_cum | number | 計10G光 |
| hikari_day | number | `h1g_day + h10g_day` |
| hikari_cum | number | `h1g_cum + h10g_cum` |
| parse_status | enum | `OK / PARTIAL / NG` |
| is_business | bool | 業務投稿判定 |
| skip_reason | string | 除外理由（雑談、項目不足など） |
| normalized_store | string | 店舗名正規化後 |
| normalized_person | string | 人名正規化後 |

### 4.2 Admin集計テーブル

#### A. `MONTHLY_PERSON_STORE`
`年月 × 人 × 店舗` 粒度。

| 列名 | 説明 |
|---|---|
| ym | 年月 |
| person | 正規化済み人名 |
| store | 正規化済み店舗 |
| work_days_month | 月内の最新稼働日数（最大値） |
| mnp_month | 月内当日mnp合計 |
| hikari_month | 月内当日光合計 |
| mnp_avg_per_day | `mnp_month / work_days_month` |
| hikari_avg_per_day | `hikari_month / work_days_month` |
| mnp_month_end_cum | 月内最終投稿日の計mnp |
| hikari_month_end_cum | 月内最終投稿日の計光 |

#### B. `MONTHLY_PERSON_TOTAL`
`年月 × 人` 粒度（店舗横断）。

| 列名 | 説明 |
|---|---|
| ym | 年月 |
| person | 正規化済み人名 |
| work_days_month | 月内の最新稼働日数（最大値） |
| mnp_month | 月内当日mnp合計 |
| hikari_month | 月内当日光合計 |
| mnp_avg_per_day | 日当たりmnp |
| hikari_avg_per_day | 日当たり光 |
| mnp_month_end_cum | 月末時点計mnp |
| hikari_month_end_cum | 月末時点計光 |
| mnp_lifetime_cum | 全期間累計mnp |
| hikari_lifetime_cum | 全期間累計光 |

#### C. `MONTHLY_STORE_TOTAL`
`年月 × 店舗` 粒度。

| 列名 | 説明 |
|---|---|
| ym | 年月 |
| store | 正規化済み店舗 |
| mnp_month | 月内当日mnp合計 |
| hikari_month | 月内当日光合計 |

---

## 5. 業務投稿判定（雑談除外）

## 5.1 判定ルール
以下をすべて満たした場合のみ `is_business = true`。

1. `日付` 行を検出
2. `店舗` 行を検出
3. `名前` 行を検出
4. 指標行を1つ以上検出（`当日mnp` / `当日1G光` / `当日10G光`）

未達の場合は `is_business = false` とし、`skip_reason` に理由を格納。

## 5.2 追加除外条件
- スタンプ・短文・返信のみ（例: 「了解」「お疲れ」）
- 指標ゼロのみ、かつ必須ヘッダー欠落

---

## 6. 正規化ルール

## 6.1 店舗名
- 全角/半角空白を除去
- 末尾スペースを削除
- 表記ゆれ辞書で統一（例: `ビックカメラ新宿西口` / `ビック新宿西口`）

## 6.2 人名
- カッコ除去（例: `（` `）`）
- 余分な空白・敬称削除
- 表記ゆれ辞書で統一（例: カナ/漢字/略称）

## 6.3 数値
- `0件`、`０件` の両方対応
- 取得失敗時は `null` で保持し `parse_status = PARTIAL`

---

## 7. 集計ロジック

## 7.1 月次件数
- `mnp_month = Σ mnp_day`
- `hikari_month = Σ hikari_day`

## 7.2 月末累計
- 同一 `ym × person` で `report_date` が最大の行を採用
- `mnp_month_end_cum = その行の mnp_cum`
- `hikari_month_end_cum = その行の hikari_cum`

## 7.3 日当たり実績
- `mnp_avg_per_day = mnp_month / work_days_month`
- `hikari_avg_per_day = hikari_month / work_days_month`
- `work_days_month <= 0` または `null` の場合は `0` で返す

## 7.4 全期間累計（人別）
- `mnp_lifetime_cum = Σ mnp_month (過去月含む)`
- `hikari_lifetime_cum = Σ hikari_month (過去月含む)`

---

## 8. 出力仕様（Trendlinks実績報告版）

## 8.1 必須ヘッダー（人別）
- 年月
- 人名
- 店舗
- 稼働日数
- MNP
- 光
- 日当たりMNP
- 日当たり光
- 月末MNP累計
- 月末光累計
- 全期間MNP累計
- 全期間光累計

## 8.2 必須ヘッダー（店舗別）
- 年月
- 店舗
- MNP
- 光

---

## 9. 更新タイミング
- Ops: LINE受信ごとに `LINE_SALES_DAILY` へ追記
- Admin: 日次バッチ（1日1回）または手動実行で月次集計更新

---

## 10. 品質管理（運用ルール）
- `parse_status = NG/PARTIAL` を週次で確認
- `normalized_person/store` の未解決値を辞書へ追加
- 月末締め時に
  - `Σ当日値` と `月末累計` の乖離チェック
  - 稼働日数未入力者のチェック

---

## 11. KPI定義（最終）
- **MNP実績**: 月内 `当日mnp` 合計
- **光実績**: 月内 `(当日1G光 + 当日10G光)` 合計
- **日当たりMNP**: `MNP実績 / 稼働日数`
- **日当たり光**: `光実績 / 稼働日数`
- **人別累計実績**: 全期間の月次実績合計

以上。
