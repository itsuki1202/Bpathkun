# CX分配くん 実装タスク

## フェーズ1：計算モジュールの作成
- [x] 実装計画書の作成
- [/] `utils/cx_distributor.py` の実装
  - [ ] `load_cx_config()` / `save_cx_config()`
  - [ ] `get_default_cx_members()` （user_masterからの初期生成）
  - [ ] `calculate_cx_distribution()` （STEP1〜3）
- [ ] `data/cx_data/` ディレクトリ作成（コード内で自動生成）

## フェーズ2：UI（app.py への追記）
- [ ] インポート追記
- [ ] `cx_distributor_page()` 関数の実装
  - [ ] 設定タブ（スタッフ一覧・役職・スキル手当・CXポイント・計算モード）
  - [ ] 計算結果タブ（STEP1〜3の結果一覧表示）
- [ ] サイドバーボタン追記
- [ ] ルーティング追記

## フェーズ3：検証
- [ ] 手動動作確認（設定保存・計算結果表示）
- [ ] 既存ページへの影響がないことの確認
