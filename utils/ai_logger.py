import os
import pandas as pd
from datetime import datetime

def get_log_dir(target_month: str) -> str:
    """Gets the data directory for the given target_month."""
    safe_month = target_month.replace("/", "_")
    return os.path.join('data', safe_month)

def get_log_path(target_month: str) -> str:
    """Returns the path to the current_month_log.csv for the given month."""
    return os.path.join(get_log_dir(target_month), 'current_month_log.csv')

def append_current_month_log(df: pd.DataFrame, target_month: str) -> bool:
    """
    Appends the provided DataFrame as a snapshot to current_month_log.csv.
    Adds an 'UpdateDate' column to track temporal progression.
    Returns True if successful, False otherwise.
    """
    try:
        log_dir = get_log_dir(target_month)
        os.makedirs(log_dir, exist_ok=True)
        log_path = get_log_path(target_month)
        
        # Prepare snapshot data
        snapshot_df = df.copy()
        snapshot_df['UpdateDate'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if os.path.exists(log_path):
            existing_df = pd.read_csv(log_path)
            combined_df = pd.concat([existing_df, snapshot_df], ignore_index=True)
        else:
            combined_df = snapshot_df
            
        combined_df.to_csv(log_path, index=False, encoding='utf-8-sig')
        return True
    except Exception as e:
        print(f"Failed to append log for {target_month}: {e}")
        return False

def get_current_month_log(target_month: str) -> pd.DataFrame:
    """
    Retrieves the historical snapshots for the current month.
    """
    log_path = get_log_path(target_month)
    if os.path.exists(log_path):
        try:
            return pd.read_csv(log_path)
        except Exception:
            pass
    return pd.DataFrame()
