import os
import pandas as pd
import numpy as np
from utils.data_manager import get_file_path as get_data_path, load_scoring_config
from utils.calculation import calculate_scores, calculate_organization_scores


def get_available_fiscal_years():
    """Returns a sorted list of available fiscal years based on existing data folders."""
    data_dir = 'data'
    if not os.path.exists(data_dir):
        return []
    
    fys = set()
    for entry in os.listdir(data_dir):
        if os.path.isdir(os.path.join(data_dir, entry)) and '_' in entry:
            try:
                year, month = map(int, entry.split('_'))
                # Fiscal year starts in April
                fy = year if month >= 4 else year - 1
                fys.add(fy)
            except ValueError:
                continue
    return sorted(list(fys), reverse=True)

def get_months_for_fiscal_year(fy):
    """Returns a list of YYYY_MM strings that belong to the given fiscal year."""
    months = []
    # April to December of the FY
    for m in range(4, 13):
        months.append(f"{fy}_{m:02d}")
    # January to March of the next year
    for m in range(1, 4):
        months.append(f"{fy + 1}_{m:02d}")
        
    # Filter only those that exist
    return [m for m in months if os.path.exists(os.path.join('data', m))]

def safe_zscore(series):
    """Calculates Z-score safely, handling cases with 0 std dev or empty series."""
    if series.empty:
        return series
    mean = series.mean()
    std = series.std(ddof=0) # Population standard deviation
    if std == 0 or pd.isna(std):
        return pd.Series(50.0, index=series.index)
    return ((series - mean) / std) * 10 + 50

def load_and_score_month(target_month, df_denom_func):
    """Loads data, config, and calculates scores for a specific month.
    df_denom_func is a function to calculate df_denom for that month if needed.
    """
    perf_path = get_data_path('current_performance.csv', target_month)
    if not os.path.exists(perf_path):
        return None, None, None, None
        
    df_perf = pd.read_csv(perf_path)
    config = load_scoring_config(target_month)
    
    # Calculate df_denom using the provided app-level function
    df_denom = df_denom_func(df_perf, target_month)
    
    scored_indiv = calculate_scores(df_perf, config, df_denom)
    scored_team = calculate_organization_scores(df_perf, config, 'team', scored_indiv)
    scored_shop = calculate_organization_scores(df_perf, config, 'shop', scored_indiv)
    
    return scored_indiv, scored_team, scored_shop, config

def calculate_annual_zscores(fy, df_denom_func):
    """
    Calculates aggregated Z-scores for a fiscal year.
    Returns: df_indiv_agg, df_team_agg, df_shop_agg
    """
    months = get_months_for_fiscal_year(fy)
    
    all_indiv_z = []
    all_team_z = []
    all_shop_z = []
    
    name_col = 'スタッフ名' # Primary key for individual
    team_col = 'チーム名'
    shop_col = '店舗名'
    
    for month in months:
        scored_indiv, scored_team, scored_shop, config = load_and_score_month(month, df_denom_func)
        if scored_indiv is None or scored_indiv.empty:
            continue
            
        # 1. Individual Z-Scores
        # Which columns to z-score? 'Total_Score' and all '[item]_score'
        score_cols = [c for c in scored_indiv.columns if c.endswith('_score')]
        if 'Total_Score' not in score_cols and 'Total_Score' in scored_indiv.columns:
            score_cols.append('Total_Score')
            
        temp_indiv = scored_indiv[[name_col, shop_col, team_col]].copy()
        
        # We need fallback logic if standard keys aren't present
        actual_name = next((c for c in ['スタッフ名', '氏名', 'Name'] if c in scored_indiv.columns), 'Name')
        actual_shop = next((c for c in ['店舗名', '店舗', 'Shop'] if c in scored_indiv.columns), 'Shop')
        actual_team = next((c for c in ['チーム名', 'チーム', 'Team'] if c in scored_indiv.columns), 'Team')
        
        temp_indiv['Standard_Name'] = scored_indiv[actual_name]
        temp_indiv['Standard_Shop'] = scored_indiv[actual_shop] if actual_shop in scored_indiv.columns else 'Unknown'
        temp_indiv['Standard_Team'] = scored_indiv[actual_team] if actual_team in scored_indiv.columns else 'Unknown'
        
        # Calculate Z-Scores
        temp_indiv['Z_Total_Score'] = safe_zscore(scored_indiv['Total_Score'])
        
        for item in config.get('items', []):
            i_name = item['name']
            i_col = f"{i_name}_score"
            if i_col in scored_indiv.columns:
                temp_indiv[f"Z_{i_col}"] = safe_zscore(scored_indiv[i_col])
                
        all_indiv_z.append(temp_indiv)
        
        # 2. Team Z-Scores
        if scored_team is not None and not scored_team.empty and actual_team in scored_team.columns:
            temp_team = pd.DataFrame({'Standard_Team': scored_team[actual_team]})
            temp_team['Z_Total_Score'] = safe_zscore(scored_team['Total_Score'])
            for item in config.get('items', []):
                i_col = f"{item['name']}_score"
                if i_col in scored_team.columns:
                    temp_team[f"Z_{i_col}"] = safe_zscore(scored_team[i_col])
            all_team_z.append(temp_team)
            
        # 3. Shop Z-Scores
        if scored_shop is not None and not scored_shop.empty and actual_shop in scored_shop.columns:
            temp_shop = pd.DataFrame({'Standard_Shop': scored_shop[actual_shop]})
            temp_shop['Z_Total_Score'] = safe_zscore(scored_shop['Total_Score'])
            for item in config.get('items', []):
                i_col = f"{item['name']}_score"
                if i_col in scored_shop.columns:
                    temp_shop[f"Z_{i_col}"] = safe_zscore(scored_shop[i_col])
            all_shop_z.append(temp_shop)

    # Aggregate
    df_indiv_agg = pd.DataFrame()
    df_team_agg = pd.DataFrame()
    df_shop_agg = pd.DataFrame()
    
    if all_indiv_z:
        df_indiv_all = pd.concat(all_indiv_z, ignore_index=True)
        # Assuming the latest Shop/Team assignment for a Name is their current one
        latest_assign = df_indiv_all.drop_duplicates(subset=['Standard_Name'], keep='last')[['Standard_Name', 'Standard_Shop', 'Standard_Team']]
        
        numeric_z_cols = [c for c in df_indiv_all.columns if c.startswith('Z_')]
        df_indiv_agg = df_indiv_all.groupby('Standard_Name')[numeric_z_cols].mean().reset_index()
        df_indiv_agg = df_indiv_agg.merge(latest_assign, on='Standard_Name', how='left')
        df_indiv_agg.rename(columns={'Standard_Name': 'スタッフ名', 'Standard_Shop': '所属店舗', 'Standard_Team': '所属チーム'}, inplace=True)
        df_indiv_agg['総合_Z_Rank'] = df_indiv_agg['Z_Total_Score'].rank(ascending=False, method='min')
        
    if all_team_z:
        df_team_all = pd.concat(all_team_z, ignore_index=True)
        numeric_z_cols = [c for c in df_team_all.columns if c.startswith('Z_')]
        df_team_agg = df_team_all.groupby('Standard_Team')[numeric_z_cols].mean().reset_index()
        df_team_agg.rename(columns={'Standard_Team': 'チーム名'}, inplace=True)
        df_team_agg['総合_Z_Rank'] = df_team_agg['Z_Total_Score'].rank(ascending=False, method='min')

    if all_shop_z:
        df_shop_all = pd.concat(all_shop_z, ignore_index=True)
        numeric_z_cols = [c for c in df_shop_all.columns if c.startswith('Z_')]
        df_shop_agg = df_shop_all.groupby('Standard_Shop')[numeric_z_cols].mean().reset_index()
        df_shop_agg.rename(columns={'Standard_Shop': '店舗名'}, inplace=True)
        df_shop_agg['総合_Z_Rank'] = df_shop_agg['Z_Total_Score'].rank(ascending=False, method='min')

    return df_indiv_agg, df_team_agg, df_shop_agg
