import pandas as pd
import numpy as np

def calculate_relative_score(val, all_values, weight):
    """【相対評価（共通）】実績値のランキング。1位満点、以降等分割減点。実績0は0点。"""
    if pd.isna(val) or val <= 0: return 0
    valid_values = [v for v in all_values if pd.notna(v)]
    N = len(valid_values)
    if N <= 0: return 0
    if N == 1: return weight
    rank = sum(1 for v in valid_values if v > val) + 1 
    # 等分割減点
    step = weight / N
    score = weight - (rank - 1) * step
    return max(0, score)

def calculate_absolute_score(val, target, weight):
    """【個人・絶対評価】達成割合×配点（MAX満点キャップ）"""
    if pd.isna(val) or pd.isna(target) or target <= 0: return 0
    achievement_rate = val / target
    return min(achievement_rate * weight, weight)

def calculate_org_absolute_score(val, target, weight):
    """【組織・絶対評価】ALL or NOTHING (100%達成で満点、それ以外は0点)"""
    if pd.isna(val) or pd.isna(target) or target <= 0: return 0
    if val >= target: return weight
    return 0

def calculate_relative_absolute_score(val, target, all_values, all_targets, weight):
    """【相対絶対（共通）】達成率のランキング。等分割減点。
    ※動的目標値（全体平均の50%）をクリアしている場合は、最低保証として配点の50%を与える。
    """
    if pd.isna(val) or pd.isna(target) or target <= 0 or val <= 0: return 0
    my_rate = val / target
    
    # 参加者全員の分子合計 ÷ 分母合計で全体平均率を算出する
    valid_pairs = [(v, t) for v, t in zip(all_values, all_targets) if pd.notna(v) and pd.notna(t) and t > 0]
    total_val = sum(v for v, t in valid_pairs)
    total_denom = sum(t for v, t in valid_pairs)
    
    avg_rate = total_val / total_denom if total_denom > 0 else 0
    
    # 全体平均率の50%を「最低目標値」とする
    min_target_rate = avg_rate * 0.5
    
    # ランキング評価 (Rank Evaluation)
    valid_rates = [v/t for v, t in valid_pairs]
    N = len(valid_rates)
    
    if N <= 0: return 0
    if N == 1: 
        score = weight
    else:
        rank = sum(1 for r in valid_rates if r > my_rate) + 1
        step = weight / N
        score = max(0, weight - (rank - 1) * step)
        
    # 最低保証ロジック (Minimum Guarantee Logic)
    # 個人の獲得率が「最低目標値」を超えている場合、順位評価とは別に『配点の50%』を最低保証として付与
    if min_target_rate > 0 and my_rate > min_target_rate:
        guarantee = weight * 0.5
        if score < guarantee:
            score = guarantee
            
    return score

def calculate_denominators(df_perf, df_orders, df_terminals, df_users):
    """ (Keep logic same as before) """
    results = []
    name_col = next((c for c in ['スタッフ名', '氏名', 'Name'] if c in df_perf.columns), None)
    if not name_col: return pd.DataFrame()
        
    for idx, row in df_perf.iterrows():
        staff_name = row[name_col]
        std_hours = 7.5
        if not df_users.empty and 'name' in df_users.columns:
            user_row = df_users[df_users['name'] == staff_name]
            if not user_row.empty:
                v = user_row.iloc[0].get('standard_hours', 7.5)
                std_hours = float(v) if pd.notna(v) else 7.5
                
        total_hours = row.get('総勤務時間', 0)
        total_hours = float(total_hours) if pd.notna(total_hours) and total_hours != "" else 0
        actual_days = total_hours / std_hours if std_hours > 0 else 0
            
        staff_orders = df_orders[df_orders['name'] == staff_name] if df_orders is not None else pd.DataFrame()
        staff_terms = df_terminals[df_terminals['name'] == staff_name] if df_terminals is not None else pd.DataFrame()
        
        order_sum = float(staff_orders['order_count'].sum()) if not staff_orders.empty and 'order_count' in staff_orders.columns else 0.0
        term_sum = float(staff_terms['terminal_count'].sum()) if not staff_terms.empty and 'terminal_count' in staff_terms.columns else 0.0
        term_days = float(staff_terms[staff_terms['terminal_count'] >= 1]['date'].nunique()) if not staff_terms.empty and 'terminal_count' in staff_terms.columns else 0.0
            
        avg_term = term_sum / actual_days if actual_days > 0 else 0.0
        avg_order = order_sum / actual_days if actual_days > 0 else 0.0
        
        val_avg_attendance = actual_days * 0.25
        val_order = order_sum * 0.25
        val_term = term_sum * 0.20
        val_term_days = term_days * 0.10
        val_avg_term = avg_term * 0.10
        val_avg_order = avg_order * 0.10
        denominator = sum([val_avg_attendance, val_order, val_term, val_term_days, val_avg_term, val_avg_order])
        
        results.append({
            name_col: staff_name, '成績分母': denominator, '実質出勤日数(平均出勤)': actual_days,
            'オーダー数(合計)': order_sum, '端末販売数(合計)': term_sum, '端末販売数出勤日': term_days,
            '平均端末販売数': avg_term, '平均オーダー': avg_order, '総勤務時間': total_hours, '基準勤務時間': std_hours,
            '分母内訳_平均出勤': val_avg_attendance, '分母内訳_オーダー数': val_order, '分母内訳_端末販売数': val_term,
            '分母内訳_端末出勤日': val_term_days, '分母内訳_平均端末': val_avg_term, '分母内訳_平均オーダー': val_avg_order
        })
    df_res = pd.DataFrame(results)
    if not df_res.empty and '成績分母' in df_res.columns:
        max_d = df_res['成績分母'].max()
        df_res['正規化成績分母'] = df_res['成績分母'] / max_d if max_d > 0 else 1.0
    return df_res

def _get_individual_base_eval_value(val, staff_name, df_denom, name_col, eval_type):
    """ 個人の実績値を評価用に補正する（相対評価・相対絶対でランキングに使う値）"""
    if pd.isna(val) or val <= 0: return val
    if eval_type == 'shop_achievement': return val # no denominator
    if df_denom is not None and not df_denom.empty and name_col:
        d_row = df_denom[df_denom[name_col] == staff_name]
        if not d_row.empty:
            denom_val = d_row.iloc[0].get('成績分母', 1.0)
            if denom_val > 0:
                return val / denom_val
    return val

def normalize_col_name(col):
    import unicodedata
    import re
    if not isinstance(col, str): return str(col)
    # Remove all spaces, newlines, and carriage returns
    col = re.sub(r'[\s　\n\r]+', '', col)
    return unicodedata.normalize('NFKC', col)

def calculate_scores(df, config, df_denom=None):
    """ 個人のスコア計算を行う """
    scored_df = df.copy()
    items = config.get('items', [])
    # Define internal column names (Standardized JP)
    name_col = 'スタッフ名'
    shop_col = '店舗名'
    team_col = 'チーム名'
    total_score_col = 'Total_Score'
    scored_df[total_score_col] = 0.0

    for item in items:
        target_col = normalize_col_name(item['name'])
        if target_col not in df.columns: continue
        
        weight = item.get('weight', 0)
        eval_type = item.get('evaluation_type', 'relative')
        new_col = f"{target_col}_score"
        
        all_targets = []
        for idx in range(len(df)):
            tar = 1.0
            staff_name = df.iloc[idx][name_col] if name_col in df.columns else ""
            
            if shop_col and eval_type == 'shop_achievement':
                 shop_name = df.iloc[idx][shop_col]
                 shop_targets = item.get('shop_targets', {})
                 s_tar = float(shop_targets.get(shop_name, 0))
                 if s_tar > 0:
                     count_mem = len(df[df[shop_col] == shop_name])
                     tar = s_tar / count_mem if count_mem else 1.0
            else:
                 # Normal target retrieval
                 targets_dict = item.get('targets', {})
                 if staff_name in targets_dict:
                     tar = float(targets_dict[staff_name])
            
            all_targets.append(tar)

        # Pre-calculate shop totals if this is a shop_achievement item
        shop_totals = {}
        if eval_type in ['shop_achievement', 'independent_achievement'] and shop_col:
            # Drop NaN and ensure numeric representation for sum
            valid_df = df.dropna(subset=[target_col])
            numeric_vals = pd.to_numeric(valid_df[target_col], errors='coerce').fillna(0)
            shop_totals = valid_df.assign(_temp_val=numeric_vals).groupby(shop_col)['_temp_val'].sum().to_dict()

        # Pre-calculate team totals for team_absolute and team_relative
        team_totals = {}
        team_col = next((c for c in ['チーム名', 'チーム', 'Team'] if c in df.columns), None)
        if eval_type in ['team_absolute', 'team_relative', 'shop_achievement', 'independent_achievement'] and team_col:
            valid_df = df.dropna(subset=[target_col])
            numeric_vals = pd.to_numeric(valid_df[target_col], errors='coerce').fillna(0)
            team_totals = valid_df.assign(_temp_val=numeric_vals).groupby(team_col)['_temp_val'].sum().to_dict()

        raw_values = df[target_col].tolist()
        eval_base_values = []
        denom_values = []
        for idx_val, val in enumerate(raw_values):
            staff_name = df.iloc[idx_val][name_col] if name_col in df.columns else ""
            
            # Extract denominator value
            denom = 1.0
            
            # 使用する分母の種別をコンフィグから取得
            denom_type = item.get('denominator_type', 'SAITO式')
            
            # 1. コンフィグで「SAITO式」以外が指定されている場合
            if denom_type != 'SAITO式':
                # 柔軟な列名検索: 
                # 1. 正確なマッチング (例: "スマホ総販分母")
                # 2. 項目名と「分母」の両方を含む列を探す (例: "スマホ総販(分母)")
                
                spec_col_1 = f"{target_col}分母"
                spec_col_2 = f"{target_col}の分母"
                spec_col_3 = f"{target_col}_分母"
                
                d_val = None
                if spec_col_1 in df.columns:
                    d_val = df.iloc[idx_val][spec_col_1]
                elif spec_col_2 in df.columns:
                    d_val = df.iloc[idx_val][spec_col_2]
                elif spec_col_3 in df.columns:
                    d_val = df.iloc[idx_val][spec_col_3]
                elif denom_type in df.columns:
                    d_val = df.iloc[idx_val][denom_type]
                else:
                    # 部分一致検索: 項目名が含まれていて、かつ「分母」というキーワードがある列
                    matched_denom_col = next((c for c in df.columns if target_col in str(c) and "分母" in str(c)), None)
                    if matched_denom_col:
                        d_val = df.iloc[idx_val][matched_denom_col]
                    
                if d_val is not None and pd.notna(d_val) and float(d_val) > 0:
                    denom = float(d_val)
            # 2. 個別分母の指定がない（または見つからない）場合は、共通のSAITO式分母（df_denom）を利用する
            elif eval_type != 'shop_achievement' and df_denom is not None and not df_denom.empty and name_col:
                d_row = df_denom[df_denom[name_col] == staff_name]
                if not d_row.empty:
                    d_val = d_row.iloc[0].get('成績分母', 1.0)
                    if d_val > 0: denom = d_val
            
            # Calculate eval_val directly utilizing the resolved denominator
            # For shop_achievement, the eval_val is the shop's total achievement rate against the shop_target
            if eval_type == 'shop_achievement':
                if not shop_col:
                    eval_val = 0
                else:
                    shop_name = df.iloc[idx_val][shop_col]
                    shop_total_perf = shop_totals.get(shop_name, 0)
                    shop_targets_dict = item.get('shop_targets', {})
                    shop_target = float(shop_targets_dict.get(shop_name, 0))
                    if shop_target > 0:
                        eval_val = shop_total_perf / shop_target
                    else:
                        eval_val = 0 # No target set -> 0
            elif eval_type == 'team_absolute':
                if not team_col:
                    eval_val = 0
                else:
                    team_name = df.iloc[idx_val][team_col]
                    team_total_perf = team_totals.get(team_name, 0)
                    
                    # For team_absolute, the target is the sum of the personal targets of the team members
                    # (since targets were divided by member count earlier, we sum them up here)
                    team_members_indices = df[df[team_col] == team_name].index.tolist()
                    team_target = sum([all_targets[i] for i in team_members_indices])
                    
                    if team_target > 0:
                        eval_val = team_total_perf / team_target
                    else:
                        eval_val = 0
            elif eval_type == 'team_relative':
                 if not team_col:
                     eval_val = 0
                 else:
                     team_name = df.iloc[idx_val][team_col]
                     team_total_perf = team_totals.get(team_name, 0)
                     
                     # For team_relative, we divide by the sum of denominators for the team
                     team_members_indices = df[df[team_col] == team_name].index.tolist()
                     
                     # We need to compute the individual denominators first if we want to sum them
                     # But at this point we are inside the loop.
                     # Let's extract the individual denom logic to a function or calculate the team denom here.
                     
                     team_denom_sum = 0
                     for ti in team_members_indices:
                         t_d_val = 1.0
                         if denom_type != 'SAITO式':
                             spec_col_1 = f"{target_col}分母"
                             spec_col_2 = f"{target_col}_分母"
                             if spec_col_1 in df.columns:
                                 t_d_val = df.iloc[ti][spec_col_1]
                             elif spec_col_2 in df.columns:
                                 t_d_val = df.iloc[ti][spec_col_2]
                             elif denom_type in df.columns:
                                 t_d_val = df.iloc[ti][denom_type]
                         elif df_denom is not None and not df_denom.empty and name_col:
                             t_staff_name = df.iloc[ti][name_col]
                             d_row = df_denom[df_denom[name_col] == t_staff_name]
                             if not d_row.empty:
                                 t_d_val = d_row.iloc[0].get('成績分母', 1.0)
                         
                         if t_d_val is not None and pd.notna(t_d_val) and float(t_d_val) > 0:
                             team_denom_sum += float(t_d_val)
                     
                     if team_denom_sum > 0:
                         eval_val = team_total_perf / team_denom_sum
                     else:
                         eval_val = 0
            elif pd.isna(val) or val <= 0:
                eval_val = val
            else:
                eval_val = val / denom
                
            denom_values.append(denom)
            eval_base_values.append(eval_val)

        scores = []
        for idx_val, val in enumerate(raw_values):
            score = 0
            if pd.isna(val): val = 0
            
            if eval_type in ['absolute', 'shop_achievement', 'team_absolute']:
                # For shop_achievement and team_absolute, eval_val is the organization's achievement rate
                if eval_type == 'absolute':
                    target = all_targets[idx_val]
                
                # Check for staged achievement targets
                st_targets = item.get('targets', {})
                st_weights = item.get('weights', {})
                w1 = st_weights.get('w1', weight)
                w2 = st_weights.get('w2', 0)
                w3 = st_weights.get('w3', 0)
                
                ar1 = st_targets.get('acquisition_rate_1', 0)
                if pd.isna(ar1) or str(ar1).strip() == '': ar1 = st_targets.get('achievement_rate_1', 0)
                ar1 = float(ar1) if pd.notna(ar1) and str(ar1).strip() != '' else 0.0
                
                ar2 = st_targets.get('acquisition_rate_2', 0)
                if pd.isna(ar2) or str(ar2).strip() == '': ar2 = st_targets.get('achievement_rate_2', 0)
                ar2 = float(ar2) if pd.notna(ar2) and str(ar2).strip() != '' else 0.0
                
                ar3 = st_targets.get('acquisition_rate_3', 0)
                if pd.isna(ar3) or str(ar3).strip() == '': ar3 = st_targets.get('achievement_rate_3', 0)
                ar3 = float(ar3) if pd.notna(ar3) and str(ar3).strip() != '' else 0.0
                
                if ar1 > 0:
                    # Staged scoring based on eval_val (team achievement rate or individual rate based on eval_type)
                    # Convert eval_val to percentage (0.0-1.0 to 0-100) to match UI entered targets
                    eval_val_pct = eval_base_values[idx_val] * 100.0
                    if pd.isna(eval_val_pct):
                        score = 0
                    elif eval_val_pct >= ar1:
                        score = w1
                    elif ar2 > 0 and eval_val_pct >= ar2:
                        score = w2
                    elif ar3 > 0 and eval_val_pct >= ar3:
                        score = w3
                    else:
                        score = 0
                        
                    # Individual Salvage Logic (only applies if the team missed the target)
                    if eval_type == 'team_absolute' and score == 0:
                        individual_salvage = item.get('flags', {}).get('individual_salvage', False)
                        if individual_salvage:
                            # Check if the individual hit their own target
                            indiv_target = all_targets[idx_val]
                            indiv_eval = (val / indiv_target) if indiv_target > 0 else 0
                            if indiv_eval >= ar1:
                                score = w1
                            elif ar2 > 0 and indiv_eval >= ar2:
                                score = w2
                            elif ar3 > 0 and indiv_eval >= ar3:
                                score = w3
                                
                    # Team Salvage Logic (only applies if the shop missed the target)
                    if eval_type == 'shop_achievement' and score == 0:
                        team_salvage = item.get('flags', {}).get('team_salvage', False)
                        if team_salvage:
                            # Check if the team hit their team target
                            team_name = df.iloc[idx_val][team_col] if team_col else None
                            if team_name:
                                team_total_perf = team_totals.get(team_name, 0)
                                team_members_indices = df[df[team_col] == team_name].index.tolist()
                                team_target = sum([all_targets[i] for i in team_members_indices])
                                team_eval = (team_total_perf / team_target) if team_target > 0 else 0
                                if team_eval >= ar1:
                                    score = w1
                                elif ar2 > 0 and team_eval >= ar2:
                                    score = w2
                                elif ar3 > 0 and team_eval >= ar3:
                                    score = w3
                else:
                    # Linear scoring based on personal target (only applies to normal absolute evaluation)
                    if eval_type == 'absolute':
                        score = min(calculate_absolute_score(val, target, weight), weight)
                    elif eval_type == 'shop_achievement' or eval_type == 'team_absolute':
                        # Shop/Team achievement with no stages set: award weight if achievement >= 100%
                        eval_val = eval_base_values[idx_val]
                        if eval_val >= 1.0:
                            score = weight
                        else:
                            score = 0
                            # Salvage check
                            if eval_type == 'team_absolute':
                                individual_salvage = item.get('flags', {}).get('individual_salvage', False)
                                if individual_salvage:
                                    indiv_target = all_targets[idx_val]
                                    indiv_eval = (val / indiv_target) if indiv_target > 0 else 0
                                    if indiv_eval >= 1.0:
                                        score = weight
                                        
                            if eval_type == 'shop_achievement':
                                team_salvage = item.get('flags', {}).get('team_salvage', False)
                                if team_salvage:
                                    team_name = df.iloc[idx_val][team_col] if team_col else None
                                    if team_name:
                                        team_total_perf = team_totals.get(team_name, 0)
                                        team_members_indices = df[df[team_col] == team_name].index.tolist()
                                        team_target = sum([all_targets[i] for i in team_members_indices])
                                        team_eval = (team_total_perf / team_target) if team_target > 0 else 0
                                        if team_eval >= 1.0:
                                            score = weight
            elif eval_type == 'independent_achievement':
                # 【達成評価 (独立加算)】
                # 個人・チーム・店舗の達成を独立して判定し、それぞれ配点1, 2, 3を加算する
                st_targets = item.get('targets', {})
                st_weights = item.get('weights', {})
                
                # 目標値の取得
                p_target = float(st_targets.get('personal_target', 0))
                t_target = float(st_targets.get('team_target', 0))
                
                # 配点の取得
                w1 = float(st_weights.get('w1', weight))
                w2 = float(st_weights.get('w2', 0))
                w3 = float(st_weights.get('w3', 0))
                
                person_score = 0
                team_score = 0
                shop_score = 0
                
                # 1. 個人判定
                if p_target > 0 and val >= p_target:
                    person_score = w1
                    
                # 2. チーム判定
                if team_col and t_target > 0:
                    team_name = normalize_col_name(df.iloc[idx_val][team_col])
                    team_total = team_totals.get(team_name, 0)
                    if team_total >= t_target:
                        team_score = w2
                        
                # 3. 店舗判定
                if shop_col:
                    shop_name = normalize_col_name(df.iloc[idx_val][shop_col])
                    shop_total = shop_totals.get(shop_name, 0)
                    shop_targets_dict = item.get('shop_targets', {})
                    s_target = float(shop_targets_dict.get(shop_name, 0))
                    if s_target > 0 and shop_total >= s_target:
                        shop_score = w3
                        
                score = person_score + team_score + shop_score

            elif eval_type == 'relative_absolute':
                eval_val = eval_base_values[idx_val]
                target = all_targets[idx_val]
                score = calculate_relative_absolute_score(eval_val, target, eval_base_values, all_targets, weight)
                if val <= 0: score = 0
            else: # 【相対評価】
                eval_val = eval_base_values[idx_val]
                score = calculate_relative_score(eval_val, eval_base_values, weight)
                if val <= 0: score = 0
            
            scores.append(score)
        
        scored_df[new_col] = scores
        scored_df[f"{target_col}_denom"] = denom_values
        scored_df[total_score_col] += scored_df[new_col]
        
    scored_df['Rank'] = scored_df[total_score_col].rank(ascending=False, method='min')
    
    # SAITO式分母を「Total_Denominator」として保持（組織集計用）
    if df_denom is not None and not df_denom.empty:
        denom_subset = df_denom[[name_col, '成績分母']].rename(columns={'成績分母': 'Total_Denominator'})
        # すでに列がある場合は削除（再計算時などの重複防止）
        if 'Total_Denominator' in scored_df.columns:
            scored_df = scored_df.drop(columns=['Total_Denominator'])
        scored_df = pd.merge(scored_df, denom_subset, on=name_col, how='left')
    
    return scored_df

def calculate_organization_scores(df_original, config, org_level='team', scored_df=None):
    if df_original is None or df_original.empty:
        return pd.DataFrame()
        
    df = df_original.copy()
    """
    組織（チーム・店舗）用のスコア計算。
    成績分母での補正は行わず、組織の実績合計を元に新たに計算する。
    """
    items = config.get('items', [])
    org_col = next((c for c in ['チーム名', 'チーム', 'Team'] if c in df.columns), None) if org_level == 'team' else next((c for c in ['店舗名', '店舗', 'Shop'] if c in df.columns), None)
    if not org_col or org_col not in df.columns: return pd.DataFrame()
        
    def normalize_col_name(col):
        import unicodedata
        import re
        if not isinstance(col, str): return col
        col = re.sub(r'[\s　]+', '', col)
        return unicodedata.normalize('NFKC', col)

    item_names = [normalize_col_name(it['name']) for it in items if normalize_col_name(it['name']) in df.columns]
    
    # 組織集計時に設定された独自分母列も集計する
    spec_denoms = []
    for it in items:
        dtype = it.get('denominator_type')
        if dtype and dtype != 'SAITO式':
            cname = normalize_col_name(it.get('name'))
            if f"{cname}分母" in df.columns:
                spec_denoms.append(f"{cname}分母")
            elif f"{cname}_分母" in df.columns:
                spec_denoms.append(f"{cname}_分母")
            elif dtype in df.columns:
                spec_denoms.append(dtype)
                
    # Remove duplicates
    spec_denoms = list(set(spec_denoms))
                  
    agg_dict = {col: 'sum' for col in item_names + spec_denoms}
    
    name_col = next((c for c in ['スタッフ名', '氏名', 'Name'] if c in df.columns), None)
    if name_col: agg_dict[name_col] = 'count'
    shop_col = next((c for c in ['店舗名', '店舗', 'Shop'] if c in df.columns), None)
    if org_level == 'team' and shop_col and shop_col != org_col:
        agg_dict[shop_col] = 'first'
        
    org_agg = df.groupby(org_col).agg(agg_dict).reset_index()
    total_score_col = 'Total_Score'
    org_agg[total_score_col] = 0.0
    
    # SAITO式分母の組織集計
    team_saito_denoms = {}
    shop_saito_denoms = {}
    if scored_df is not None and 'Total_Denominator' in scored_df.columns:
        team_saito_denoms = scored_df.groupby(org_col)['Total_Denominator'].sum().to_dict()
        if shop_col and shop_col in scored_df.columns:
            shop_saito_denoms = scored_df.groupby(shop_col)['Total_Denominator'].sum().to_dict()
    
    for item in items:
        target_col = normalize_col_name(item['name'])
        if target_col not in org_agg.columns: continue
        weight = float(item.get('weight', 0))
        eval_type = item.get('evaluation_type', 'relative')
        dtype = item.get('denominator_type', 'SAITO式')
        new_col = f"{target_col}_score"
        
        raw_values = org_agg[target_col].tolist()
        
        all_targets = []
        shop_perf_totals = {}
        if eval_type == 'shop_achievement' and shop_col:
            shop_perf_totals = df.groupby(shop_col)[target_col].sum().to_dict()

        for i, row in org_agg.iterrows():
            my_org = row[org_col]
            my_shop = row[shop_col] if org_level == 'team' and shop_col in org_agg.columns else my_org
            
            # チーム/組織の実績評価に使用する分母（ターゲット）の算出
            t_val = 1.0
            
            if eval_type == 'shop_achievement':
                # 店舗達成評価の場合：店舗目標をそのまま使用（按分しない）
                shop_targets = item.get('shop_targets', {})
                t_val = float(shop_targets.get(my_shop, 0))
            else:
                # 相対評価・絶対評価の場合
                if dtype == 'SAITO式':
                    # SAITO式分母の合計を使用
                    if my_org in team_saito_denoms:
                        t_val = team_saito_denoms[my_org]
                    else:
                        # フォールバック：人数
                        t_val = float(row[name_col]) if name_col in row else 1.0
                else:
                    # 独自分母を使用（集計済みカラムから取得）
                    spec_col = None
                    if f"{target_col}分母" in org_agg.columns: spec_col = f"{target_col}分母"
                    elif f"{target_col}_分母" in org_agg.columns: spec_col = f"{target_col}_分母"
                    elif dtype in org_agg.columns: spec_col = dtype
                    
                    if spec_col:
                        t_val = float(row[spec_col])
                    else:
                        # 目標設定から合算（旧ロジック互換）
                        team_members_indices = df[df[org_col] == my_org].index.tolist()
                        targets_dict = item.get('targets', {})
                        p_target_def = float(targets_dict.get('personal_target', 0))
                        team_target_sum = 0
                        for ti in team_members_indices:
                            t_staff_name = df.iloc[ti][name_col] if name_col in df.columns else ""
                            team_target_sum += float(targets_dict.get(t_staff_name, p_target_def))
                        t_val = team_target_sum if team_target_sum > 0 else float(len(team_members_indices))
            
            if t_val <= 0: t_val = 1.0
            all_targets.append(t_val)
            
        scores = []
        for idx, val in enumerate(raw_values):
            score = 0
            target = all_targets[idx]
            
            # 使用する実績値を評価タイプに応じて決定
            eval_perf_val = val
            if eval_type == 'shop_achievement' and shop_col:
                my_shop = org_agg.iloc[idx][shop_col] if shop_col in org_agg.columns else org_agg.iloc[idx][org_col]
                eval_perf_val = shop_perf_totals.get(my_shop, 0)

            if eval_type in ['absolute', 'shop_achievement', 'team_absolute']:
                # ステップ評価の取得
                st_targets = item.get('targets', {})
                st_weights = item.get('weights', {})
                w1 = float(st_weights.get('w1', weight)) if pd.notna(st_weights.get('w1')) else float(weight)
                w2 = float(st_weights.get('w2', 0)) if pd.notna(st_weights.get('w2')) else 0.0
                w3 = float(st_weights.get('w3', 0)) if pd.notna(st_weights.get('w3')) else 0.0
                
                ar1 = st_targets.get('acquisition_rate_1', 0)
                if pd.isna(ar1) or str(ar1).strip() == '': ar1 = st_targets.get('achievement_rate_1', 0)
                ar1 = float(ar1) if pd.notna(ar1) and str(ar1).strip() != '' else 0.0
                
                ar2 = st_targets.get('acquisition_rate_2', 0)
                if pd.isna(ar2) or str(ar2).strip() == '': ar2 = st_targets.get('achievement_rate_2', 0)
                ar2 = float(ar2) if pd.notna(ar2) and str(ar2).strip() != '' else 0.0
                
                ar3 = st_targets.get('acquisition_rate_3', 0)
                if pd.isna(ar3) or str(ar3).strip() == '': ar3 = st_targets.get('achievement_rate_3', 0)
                ar3 = float(ar3) if pd.notna(ar3) and str(ar3).strip() != '' else 0.0
                
                # 達成率の計算
                org_rate = (eval_perf_val / target if target > 0 else 0)
                
                if ar1 > 0:
                    # ステップ判定 (UI入力は % 単位なので 100 倍して比較)
                    rate_pct = org_rate * 100.0
                    if rate_pct >= ar1: score = w1
                    elif ar2 and rate_pct >= float(ar2): score = w2
                    elif ar3 and rate_pct >= float(ar3): score = w3
                    else: score = 0
                else:
                    # 目標1が設定されていない場合は100%達成で満点
                    if org_rate >= 1.0:
                        score = max(w1, w2, w3)
                    else:
                        score = 0
            
            elif eval_type == 'relative_absolute':
                score = calculate_relative_absolute_score(eval_perf_val, target, raw_values, all_targets, weight)
                if val <= 0: score = 0
            elif eval_type in ['team_relative', 'relative']:
                # 効率(実績/ターゲット)のランキング
                org_efficiency_values = [raw_values[j] / all_targets[j] for j in range(len(raw_values))]
                score = calculate_relative_score(org_efficiency_values[idx], org_efficiency_values, weight)
                if val <= 0: score = 0
            else:
                score = calculate_relative_score(eval_perf_val, raw_values, weight)
                if val <= 0: score = 0
                
            scores.append(score)
            
        org_agg[new_col] = scores
        org_agg[total_score_col] += org_agg[new_col]
        
    org_agg['Rank'] = org_agg[total_score_col].rank(ascending=False, method='min')
    
    # Identify Leader/Manager from individual scored_df if provided
    if scored_df is not None and name_col in scored_df.columns:
        best_col = 'Leader' if org_level == 'team' else 'Manager'
        leaders = scored_df.sort_values('Total_Score', ascending=False).groupby(org_col)[name_col].first().reset_index()
        leaders.rename(columns={name_col: best_col}, inplace=True)
        org_agg = pd.merge(org_agg, leaders, on=org_col, how='left')
        
        # 店舗成績の場合、マネージャー名はすべて空欄にする（ユーザー要望）
        if org_level == 'shop' and 'Manager' in org_agg.columns:
            org_agg['Manager'] = ""
    else:
        org_agg['Leader' if org_level == 'team' else 'Manager'] = ""
        
    return org_agg.sort_values('Rank')

def aggregate_team_scores(scored_df, df_original=None, config=None):
    if df_original is not None and config is not None:
        return calculate_organization_scores(df_original, config, 'team', scored_df)
    return pd.DataFrame()

def aggregate_shop_scores(scored_df, df_original=None, config=None):
    if df_original is not None and config is not None:
        return calculate_organization_scores(df_original, config, 'shop', scored_df)
    return pd.DataFrame()

def get_calculation_audit_df(df, config, df_denom):
    """
    計算プロセスの詳細（分母、達成率、しきい値、配点）を一覧化し、管理者が検証できるようにします。
    """
    audit_rows = []
    items = config.get('items', [])
    
    # 名称列の特定
    name_col = None
    for c in ['スタッフ名', '氏名', 'Name']:
        if c in df.columns: name_col = c; break
    if not name_col: name_col = df.columns[0]
    
    # 店舗・チーム列の特定
    shop_col = None; team_col = None
    for c in ['店舗名', '店舗', 'Shop']:
        if c in df.columns: shop_col = c; break
    for c in ['チーム名', 'チーム', 'Team']:
        if c in df.columns: team_col = c; break
        
    for item in items:
        item_name = item.get('name', '')
        weight = item.get('weight', 0)
        eval_type = item.get('evaluation_type', 'relative')
        denom_type = item.get('denominator_type', 'SAITO式')
        
        # 配点としきい値の取得
        st_weights = item.get('weights', {})
        w1 = st_weights.get('w1', weight)
        w2 = st_weights.get('w2', 0)
        w3 = st_weights.get('w3', 0)
        
        st_targets = item.get('targets', {})
        ar1 = st_targets.get('acquisition_rate_1', 0)
        if pd.isna(ar1) or str(ar1).strip() == '': ar1 = st_targets.get('achievement_rate_1', 0)
        ar1 = float(ar1) if pd.notna(ar1) and str(ar1).strip() != '' else 0.0
        
        ar2 = st_targets.get('acquisition_rate_2', 0)
        if pd.isna(ar2) or str(ar2).strip() == '': ar2 = st_targets.get('achievement_rate_2', 0)
        ar2 = float(ar2) if pd.notna(ar2) and str(ar2).strip() != '' else 0.0
        
        ar3 = st_targets.get('acquisition_rate_3', 0)
        if pd.isna(ar3) or str(ar3).strip() == '': ar3 = st_targets.get('achievement_rate_3', 0)
        ar3 = float(ar3) if pd.notna(ar3) and str(ar3).strip() != '' else 0.0

        # 実績列の特定
        norm_name = normalize_col_name(item_name)
        perf_col = None
        for c in df.columns:
            if normalize_col_name(c) == norm_name:
                perf_col = c
                break
        
        if not perf_col: continue
            
        # 分母の取得
        denom_col = f"{item_name}分母"
        
        for idx, row in df.iterrows():
            person_name = row[name_col]
            perf_val = row[perf_col]
            if pd.isna(perf_val): perf_val = 0
            
            # 分母の特定
            denom_val = 0
            if denom_type == 'SAITO式':
                if df_denom is not None and not df_denom.empty:
                    # '正規化成績分母' があれば優先、なければ '成績分母'
                    d_col = '正規化成績分母' if '正規化成績分母' in df_denom.columns else '成績分母'
                    # name_col (スタッフ名) でマッチング
                    match_denom = df_denom[df_denom[name_col] == person_name]
                    if not match_denom.empty:
                        denom_val = match_denom[d_col].iloc[0]
            else:
                denom_val = row.get(denom_col, 0)
            
            if pd.isna(denom_val) or denom_val == 0: denom_val = 1.0 # 0割回避
            
            # ターゲットの特定
            target_val = 1.0
            if eval_type == 'shop_achievement' and shop_col:
                shop_name = row[shop_col]
                s_targets = item.get('shop_targets', {})
                s_tar = float(s_targets.get(shop_name, 0))
                if s_tar > 0:
                    count_mem = len(df[df[shop_col] == shop_name])
                    target_val = s_tar / count_mem if count_mem else 1.0
            else:
                t_dict = item.get('targets', {})
                target_val = float(t_dict.get(person_name, 1.0))

            if target_val <= 0: target_val = 1.0
            
            # 達成率 (実績 / (ターゲット * 分母))
            # 多くのロジックでこの形式（または相当）が使われる
            rate = (perf_val / (target_val * denom_val))
            rate_pct = rate * 100.0
            
            # 最終スコアの計算を簡易再現
            score = 0
            if eval_type in ['absolute', 'team_absolute', 'shop_achievement']:
                if ar1 > 0:
                    if rate_pct >= ar1: score = w1
                    elif ar2 > 0 and rate_pct >= ar2: score = w2
                    elif ar3 > 0 and rate_pct >= ar3: score = w3
                elif rate >= 1.0:
                    score = max(w1, w2, w3)
            
            audit_rows.append({
                "対象者": person_name,
                "店舗": row.get(shop_col, "-") if shop_col else "-",
                "チーム": row.get(team_col, "-") if team_col else "-",
                "項目名": item_name,
                "評価型": eval_type,
                "実績値": perf_val,
                "分母種別": denom_type,
                "計算ターゲット": round(float(target_val), 2),
                "計算分母": round(float(denom_val), 3),
                "達成率(%)": round(float(rate_pct), 1),
                "段階(1/2/3)": f"{ar1}/{ar2}/{ar3}" if ar1 > 0 else "100%",
                "配点(1/2/3)": f"{w1}/{w2}/{w3}" if ar1 > 0 else f"{max(w1,w2,w3)}",
                "最終スコア": score
            })
            
    return pd.DataFrame(audit_rows)
