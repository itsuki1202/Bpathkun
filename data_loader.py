import pandas as pd
from master_data import REQUIRED_COLUMNS, COLUMN_MAPPING

def load_data(file):
    """
    CSV または Excel ファイルからデータを読み込む。
    Args:
        file: Streamlit の UploadedFile オブジェクト、またはファイルパス文字列
    Returns:
        pd.DataFrame: 読み込んだデータ。エラー時は None
        str: エラーメッセージ（エラーなければ None）
    """
    try:
        # ファイル種別の判定
        if isinstance(file, str):
            filename = file
            file_obj = file
        else:
            filename = file.name
            file_obj = file

        if filename.endswith('.csv'):
            df = pd.read_csv(file_obj)
        elif filename.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file_obj)
        else:
            return None, "未対応のファイル形式です。CSV または Excel ファイルをアップロードしてください。"

        # カラム名のクリーンアップ
        df.columns = [str(c).strip().replace('\n', '').replace('\r', '') for c in df.columns]

        # 必須列の候補
        staff_aliases = ["スタッフ名", "氏名", "スタッフ", "名前", "Name"]
        shop_aliases = ["店舗名", "店舗", "Shop", "拠点"]
        team_aliases = ["チーム名", "チーム", "Team"]

        # もし現在のヘッダーで必須項目が見つからない場合、データ行からヘッダーを探す
        def find_headers(target_df):
            has_staff = any(a in target_df.columns for a in staff_aliases)
            has_shop = any(a in target_df.columns for a in shop_aliases)
            return has_staff and has_shop

        if not find_headers(df):
            # ファイルの最初の10行を調べて、ヘッダーらしき行を探す
            for i in range(min(10, len(df))):
                row_values = [str(v).strip() for v in df.iloc[i]]
                if any(any(a in v for a in staff_aliases) for v in row_values):
                    # この行をヘッダーとして再設定
                    df.columns = row_values
                    df = df.iloc[i+1:].reset_index(drop=True)
                    # 再度クリーンアップ
                    df.columns = [str(c).strip().replace('\n', '').replace('\r', '') for c in df.columns]
                    break

        # マッピングの適用（最終的な標準名に変換）
        new_cols = []
        for c in df.columns:
            mapped = c
            if any(a == c for a in staff_aliases): mapped = "スタッフ名"
            elif any(a == c for a in shop_aliases): mapped = "店舗名"
            elif any(a == c for a in team_aliases): mapped = "チーム名"
            elif c in ["代理店", "代理店名", "Agency"]: mapped = "代理店名"
            new_cols.append(mapped)
        df.columns = new_cols

        # 必須列の最終確認
        required_core = ["スタッフ名", "店舗名", "チーム名"]
        missing_cols = [col for col in required_core if col not in df.columns]

        if missing_cols:
            # 最後の手段: 部分一致
            for target in missing_cols:
                for actual in df.columns:
                    # targetが「スタッフ名」のとき「スタッフ」が含まれていれば
                    clean_target = target.replace("名", "")
                    if clean_target in str(actual):
                        df.rename(columns={actual: target}, inplace=True)
                        break
            
            # 再チェック
            missing_cols = [col for col in required_core if col not in df.columns]
            if missing_cols:
                actual_cols = ", ".join(list(df.columns)[:10])
                return None, f"不足している列があります: {', '.join(missing_cols)} (読み取られた列: {actual_cols}...)"

        # 代理店名は任意
        if "代理店名" not in df.columns:
            df["代理店名"] = "不明"

        return df, None

    except Exception as e:
        return None, f"ファイルの読み込みエラー: {e}"
