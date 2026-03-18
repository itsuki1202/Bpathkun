# CX分配くん 実装計画書

## 概要

ブライトパスくんの既存成績ランキングを参照してインセンティブを計算する、**管理者専用ページ「CX分配くん」**を追加する。  
**既存の `app.py` のロジック（`admin_page` / 各種集計関数 / `calculation.py` など）には一切手を加えない。**

---

## 変更ファイル一覧

### 新規作成

#### [NEW] `utils/cx_distributor.py`
CX分配くん専用の計算ロジックモジュール。

- `load_cx_config(month)` — 月別CX設定JSONを読み込む
- `save_cx_config(config, month)` — 月別CX設定JSONに保存
- `get_default_cx_members(user_master)` — `user_master.json` からスタッフ一覧・役職を初期生成
- `calculate_cx_distribution(cx_config, shop_rank_df, team_rank_df, individual_scored_df)` — STEP1〜3の計算を実行し結果DataFrameを返す

#### [NEW] `data/cx_data/` ディレクトリ
月別CX設定データの保存先。ファイル名例：`data/cx_data/202603/cx_config.json`

#### [NEW] `docs/cx_distributor/task.md`
タスク管理ファイル。

---

### 変更ファイル

#### [MODIFY] `app.py`

> [!IMPORTANT]
> 既存の全関数・全ロジックには一切触れない。追記のみで対応する。

追記箇所：
1. **インポート追記**（先頭付近）：`from utils.cx_distributor import ...`
2. **新関数 `cx_distributor_page()`** を末尾に追加（既存関数とは完全独立）
3. **サイドバーボタン追記**（管理者専用ボタン群の末尾）：「CX分配くん」ボタンを追記
4. **ルーティング追記**（`if mode == "設定 (Admin)":` ブロックの末尾）：`elif mode == "CX分配くん":` を追記

---

## 実装詳細

### `utils/cx_distributor.py` の構造

```python
CX_DATA_DIR = "data/cx_data"

# cx_config.json のスキーマ（月別保存）
{
  "month": "202603",
  "cx_total_points": 1000000,   # 当月CXポイント原資
  "base_pt": 50,                 # ベースポイント（変更可能）
  "calculation_mode": "A",       # "A" or "B"
  "skill_allowances": {          # スキル手当単価
    "フロントスペシャリスト": 0,
    "グランマイスター": 0,
    "マイスター": 0,
    "プレマイスター": 0
  },
  "members": [
    {
      "name": "スタッフ名",
      "role": "role_member",     # role_manager / role_leader / role_member
      "skill_type": "マイスター", # スキル手当区分
      "manual_pt": null          # 手動ポイント（手動追加スタッフ用。nullなら自動計算）
    }
  ]
}
```

### `cx_distributor_page()` の画面構成

**2タブ構成：**
- **「⚙️ 設定」タブ**：スタッフ一覧・役職・スキル手当・CXポイント・ベースポイント・計算モードを設定・保存
- **「📊 計算結果」タブ**：STEP1〜3の計算を実行し、全スタッフの結果を一覧表示

---

## STEP 計算ロジック実装方針

### STEP1（モードA/B で分岐）
- **モードA**：各スタッフの `スキル手当金額 × 0.85`（端数切り捨て）を合算して固定費総額を算出。余剰原資 = CXポイント - 固定費総額
- **モードB**：固定費総額 = 0。余剰原資 = CXポイント原資100%

### STEP2（ポイント付与）
- 各スタッフの `Total_Pt = Base_Pt + 成果Pt`
- 成果Ptは役職に応じて既存ランキングの `Rank` 列を参照
  - `role_manager`：店舗Rankを `aggregate_shop_scores()` から取得
  - `role_leader`：チームRankを `aggregate_team_scores()` から取得
  - `role_member`：個人Rankを `calculate_scores()` から取得（店長・リーダーを除外した上での再Rank付け）
- `manual_pt` が設定されているスタッフはそのptをそのまま `Total_Pt` として使用

### STEP3（最終支給額）
- 1Pt単価 = 余剰原資 ÷ 全スタッフのTotal_Pt合算
- 歩合額 = 個人Total_Pt × 1Pt単価（四捨五入）
- **モードA**：最終支給額 = 旧手当 + 歩合額
- **モードB**：最終支給額 = 歩合額のみ

---

## 検証計画

### 自動テスト
既存テスト `tests/test_calculation.py` でカバーされているロジックには触れないため、新規テストとして追加する：

```
# tests/test_cx_distributor.py を新規作成
# 実行コマンド：
python -m pytest tests/test_cx_distributor.py -v
```

テストケース：
- `calculate_cx_distribution()` のモードA/B両方の計算結果が正しいか
- メンバー1人・複数人でのポイント計算
- 線形補間（スタッフの成果Pt計算）のエッジケース（1人の場合、2人の場合）

### 手動確認（管理者でログイン後）
1. サイドバーに「CX分配くん」ボタンが表示されることを確認
2. 「設定」タブでスタッフ一覧が自動転記されることを確認
3. 各項目（CXポイント・ベースポイント・スキル手当・計算モード）を入力・保存できることを確認
4. 「計算結果」タブで計算結果が表示されることを確認（モードA / モードB の切り替えも確認）
5. 手動追加スタッフが追加でき、手動ポイントが計算に反映されることを確認
6. 既存のブライトパスくんの各ページ（総合・順位・件数など）に影響が出ていないことを確認
