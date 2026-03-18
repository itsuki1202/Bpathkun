# CX分配くん 実装完了レポート

## 変更ファイル一覧

| ファイル | 変更種別 | 内容 |
|---|---|---|
| `utils/cx_distributor.py` | **新規作成** | CX計算ロジック専用モジュール（STEP1〜3・設定読み書き） |
| `app.py` | **追記のみ** | インポート追加・UIページ関数追加・サイドバーボタン・ルーティング |
| `docs/cx_distributor/task.md` | **新規作成** | タスク管理ファイル |
| `docs/cx_distributor/implementation_plan.md` | **新規作成** | 実装計画書 |

既存の calculation.py / data_manager.py / auth_manager.py などのロジックは一切変更していません。

---

## CX計算ロジックの実装内容

- **STEP1**：モードA → 旧手当（支援費×0.85）を天引き / モードB → 固定費0円
- **STEP2**：Base_Pt（設定値）＋ 役職別成果Pt（店長/リーダー/スタッフ線形補間）
- **STEP3**：1Pt単価 = 余剰原資 ÷ 総ポイント → 最終支給額算出
- **月別保存先**：`data/cx_data/YYYYMM/cx_config.json`
