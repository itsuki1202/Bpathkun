import json
import os
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data')

def get_file_path(filename, month=None):
    if month:
        return os.path.join(DATA_DIR, month, filename)
    return os.path.join(DATA_DIR, filename)

def load_json(filepath, default=None):
    if not os.path.exists(filepath):
        return default if default is not None else {}
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading {filepath}: {e}")
        return default if default is not None else {}

def save_json(filepath, data):
    try:
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving {filepath}: {e}")
        return False

def load_user_master(month=None):
    filepath = get_file_path('user_master.json', month)
    return load_json(filepath, default={"users": []})

def save_user_master(data, month=None):
    filepath = get_file_path('user_master.json', month)
    return save_json(filepath, data)

def load_scoring_config(month=None):
    filepath = get_file_path('scoring_config.json', month)
    return load_json(filepath, default={"items": []})

def save_scoring_config(data, month=None):
    filepath = get_file_path('scoring_config.json', month)
    return save_json(filepath, data)

def get_users_df(month=None):
    data = load_user_master(month)
    if not data['users']:
        return pd.DataFrame(columns=["name", "shop", "team", "standard_hours"])
    return pd.DataFrame(data['users'])

def apply_user_master_merge(df, month=None):
    """
    実績データ(df)に対し、ユーザーマスターの情報を結合して店舗・チーム名を最新化する。
    名前の一致判定において前後の空白や改行を削除することで、結合漏れを防ぐ。
    """
    if df is None or df.empty:
        return df
        
    master_df = get_users_df(month)
    if master_df.empty:
        return df
        
    # カラム名の正規化（スタッフ名列の特定）
    name_col = next((c for c in ['スタッフ名', '氏名', 'Name'] if c in df.columns), None)
    if not name_col:
        return df
        
    # 結合用に一時的な正規化カラムを作成（トリミング）
    df['_name_tmp'] = df[name_col].astype(str).str.strip()
    master_df['_name_tmp'] = master_df['name'].astype(str).str.strip()
    
    # 結合（店舗・チーム情報を取得。master側を優先するため、df側の既存カラムはリネームして後で補完に使う）
    # 重複を避けるため、master側から必要な3列のみ抽出
    # name_master としないのは、元の名前を表示に使いたいため
    merged = pd.merge(
        df, 
        master_df[['_name_tmp', 'shop', 'team']], 
        on='_name_tmp', 
        how='left'
    )
    
    # 元の店舗・チーム列を特定
    orig_shop_col = next((c for c in ['店舗名', '店舗', 'Shop'] if c in df.columns), '店舗名')
    orig_team_col = next((c for c in ['チーム名', 'チーム', 'Team'] if c in df.columns), 'チーム名')
    
    # マスターの情報があればそれを採用、なければ元の情報を維持
    if 'shop' in merged.columns:
        merged['店舗名'] = merged['shop'].fillna(merged[orig_shop_col] if orig_shop_col in merged.columns else pd.NA)
    if 'team' in merged.columns:
        merged['チーム名'] = merged['team'].fillna(merged[orig_team_col] if orig_team_col in merged.columns else pd.NA)
        
    # 不要な一時列の削除
    cols_to_drop = ['_name_tmp', 'shop', 'team']
    merged = merged.drop(columns=[c for c in cols_to_drop if c in merged.columns])
    
    # 元の不要な店舗・チーム列があれば削除（'店舗名','チーム名'に統一されたため）
    final_drops = []
    if orig_shop_col != '店舗名' and orig_shop_col in merged.columns: final_drops.append(orig_shop_col)
    if orig_team_col != 'チーム名' and orig_team_col in merged.columns: final_drops.append(orig_team_col)
    if final_drops:
        merged = merged.drop(columns=final_drops)
        
    return merged

def archive_month_data(month, df=None):
    """
    指定した年月(YYYYMM)のフォルダに、現在のユーザーマスターと評価項目設定、
    および実績データを保存する。
    """
    # 現在の最新設定をロード
    master = load_user_master()
    config = load_scoring_config()
    
    # アーカイブ先に保存
    save_user_master(master, month=month)
    save_scoring_config(config, month=month)
    
    # 実績データがあればCSVとして保存
    if df is not None:
        filepath = get_file_path('performance_data.csv', month=month)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        df.to_csv(filepath, index=False, encoding='utf-8-sig')
    
    return True

def load_performance_data(month=None):
    """
    指定した月、または最新の実績データをロードする。
    """
    filepath = get_file_path('performance_data.csv', month=month)
    if os.path.exists(filepath):
        try:
            return pd.read_csv(filepath)
        except Exception as e:
            print(f"Error loading performance data: {e}")
            return None
    return None
