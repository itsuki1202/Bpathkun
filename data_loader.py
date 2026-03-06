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

        # 英語列名を日本語に変換（レガシーデータ対応）
        df.rename(columns=COLUMN_MAPPING, inplace=True)

        # 必須列の確認（最低限: 氏名・店舗・チーム）
        required_jp_cols = ["氏名", "店舗", "チーム"]
        missing_cols = [col for col in required_jp_cols if col not in df.columns]

        if missing_cols:
            return None, f"不足している列があります: {', '.join(missing_cols)}"

        return df, None

    except Exception as e:
        return None, f"ファイルの読み込みエラー: {e}"
