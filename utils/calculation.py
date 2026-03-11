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
    
    # 内部で使用する標準カラム名
    name_col = 'スタッフ名'
    shop_col = '店舗名'
    team_col = 'チーム名'
    total_score_col = 'Total_Score'
    scored_df[total_score_col] = 0.0

    # データフレームのカラム名を一括で正規化し、検索しやすくする
    raw_cols = df.columns.tolist()
    norm_to_raw = {normalize_col_name(c): c for c in raw_cols}
    
    for item in items:
        item_name = item.get('name', '')
        # 項目名を正規化してカラムを特定
        norm_item_name = normalize_col_name(item_name)
        if norm_item_name not in norm_to_raw:
            continue
        
        target_col = norm_to_raw[norm_item_name]
        weight = float(item.get('weight', 0))
        eval_type = item.get('evaluation_type', 'relative')
        new_col = f"{norm_item_name}_score"
        
        # 分母種別の特定
        denom_type = item.get('denominator_type', 'SAITO式')
        
        scores = []
        denom_values = []
        
        for idx in range(len(df)):
            staff_name = str(df.iloc[idx][name_col]).strip() if name_col in df.columns else ""
            val = float(pd.to_numeric(df.iloc[idx][target_col], errors='coerce')) if pd.notna(df.iloc[idx][target_col]) else 0.0
            
            # --- 1. 分母 (Denominator) の決定 ---
            denom = 1.0
            if denom_type != 'SAITO式':
                # 独自分母の検索
                spec_col = next((c for c in df.columns if norm_item_name in normalize_col_name(c) and "分母" in c), None)
                if not spec_col and denom_type in df.columns: spec_col = denom_type
                
                if spec_col:
                    d_val = pd.to_numeric(df.iloc[idx][spec_col], errors='coerce')
                    if pd.notna(d_val) and d_val > 0: denom = float(d_val)
            elif df_denom is not None and not df_denom.empty:
                # SAITO式分母を使用
                d_row = df_denom[df_denom[name_col].str.strip() == staff_name] if name_col in df_denom.columns else pd.DataFrame()
                if not d_row.empty:
                    d_val = d_row.iloc[0].get('成績分母', 1.0)
                    if d_val > 0: denom = float(d_val)
            
            # --- 2. 目標値 (Target) の決定 ---
            target = 1.0
            if eval_type == 'individual_achievement' and shop_col in df.columns:
                # 個人達成: 店舗別個人目標を取得
                shop_name = df.iloc[idx][shop_col]
                shop_targets = item.get('shop_targets', {})
                ind_target = float(shop_targets.get(shop_name, 0))
                # 目標が未設定（0件）の場合は評価なし（0点）
                if ind_target <= 0:
                    scores.append(0)
                    denom_values.append(0.0)
                    continue
                # 個人判定: 個人実績 >= 個人目標 → 満点 / それ以下 → 0点
                score = weight if val >= ind_target else 0
                scores.append(score)
                denom_values.append(1.0)
                continue
            elif eval_type in ['shop_achievement'] and shop_col in df.columns:
                shop_name = df.iloc[idx][shop_col]
                shop_targets = item.get('shop_targets', {})
                target = float(shop_targets.get(shop_name, 0))
                shop_total_perf = df[df[shop_col] == shop_name][target_col].astype(float).sum()
                val_for_calc = shop_total_perf
                denom = 1.0
            else:
                targets_dict = item.get('targets', {})
                target = float(targets_dict.get(staff_name, 1.0))
                val_for_calc = val
            
            if target <= 0: target = 1.0
            
            # --- 3. 効率 / 達成率の計算 ---
            # 分母補正した実績 = 実績 / 分母
            # 達成率 = (分母補正実績) / 目標
            efficiency = val_for_calc / denom
            achievement_rate = efficiency / target
            
            # --- 4. スコア判定 (Scoring) ---
            score = 0
            if eval_type in ['absolute', 'shop_achievement', 'individual_achievement', 'team_absolute']:
                # ステップ判定 (しきい値が設定されている場合)
                st_targets = item.get('targets', {})
                st_weights = item.get('weights', {})
                ar1 = float(st_targets.get('achievement_rate_1', 0) or 0)
                
                if ar1 > 0:
                    # パーセンテージ判定
                    rate_pct = achievement_rate * 100.0
                    w1 = float(st_weights.get('w1', weight))
                    w2 = float(st_weights.get('w2', 0))
                    w3 = float(st_weights.get('w3', 0))
                    ar2 = float(st_targets.get('achievement_rate_2', 0) or 0)
                    ar3 = float(st_targets.get('achievement_rate_3', 0) or 0)
                    
                    if rate_pct >= ar1: score = w1
                    elif ar2 > 0 and rate_pct >= ar2: score = w2
                    elif ar3 > 0 and rate_pct >= ar3: score = w3
                else:
                    # 線形評価
                    score = min(achievement_rate * weight, weight)
                
                # 救済措置 (Salvage)
                if eval_type == 'team_absolute' and score == 0:
                    if item.get('flags', {}).get('individual_salvage', False):
                        if (val / target) >= 1.0: score = weight # 個人が100%達成なら救済
            
            elif eval_type == 'relative_absolute':
                # 相対絶対評価は別途関数にて計算に回すため、ここでは効率を記録
                score = achievement_rate # 後で calculate_relative_absolute_score に渡すための仮値
            else:
                # 相対評価も「効率」を元に行う
                score = efficiency

            scores.append(score)
            denom_values.append(denom)

        # 全員分の計算が終わった後、相対評価を行う
        if eval_type == 'relative':
            final_scores = []
            for s_val in scores:
                final_scores.append(calculate_relative_score(s_val, scores, weight))
            scores = final_scores
        elif eval_type == 'relative_absolute':
            # 相対絶対評価 (50%最低保証あり)
            final_scores = []
            # 本来のターゲットに対する達成率リスト
            all_rates = scores 
            for r_val in all_rates:
                # すでに achievement_rate が入っているので target=1 として呼ぶ
                final_scores.append(calculate_relative_absolute_score(r_val, 1.0, all_rates, [1.0]*len(all_rates), weight))
            scores = final_scores
        
        scored_df[new_col] = scores
        scored_df[f"{norm_item_name}_denom"] = denom_values
        scored_df[total_score_col] += scored_df[new_col]
        
    scored_df['Rank'] = scored_df[total_score_col].rank(ascending=False, method='min')
    
    # SAITO式分母のメタデータ紐付け
    if df_denom is not None and not df_denom.empty:
        denom_subset = df_denom[[name_col, '成績分母']].rename(columns={'成績分母': 'Total_Denominator'})
        if 'Total_Denominator' in scored_df.columns:
            scored_df = scored_df.drop(columns=['Total_Denominator'])
        scored_df = pd.merge(scored_df, denom_subset, on=name_col, how='left')
    
    return scored_df

def calculate_organization_scores(df_original, config, org_level='team', scored_df=None):
    if df_original is None or df_original.empty:
        return pd.DataFrame()
        
    df = df_original.copy()
    items = config.get('items', [])
    org_col = next((c for c in ['チーム名', 'チーム', 'Team'] if c in df.columns), None) if org_level == 'team' else next((c for c in ['店舗名', '店舗', 'Shop'] if c in df.columns), None)
    if not org_col or org_col not in df.columns: return pd.DataFrame()
    
    # 名称列の特定
    name_col = 'スタッフ名'
    shop_col = '店舗名'

    # カラム名正規化マップ
    raw_cols = df.columns.tolist()
    norm_to_raw = {normalize_col_name(c): c for c in raw_cols}
    item_names = [norm_to_raw[normalize_col_name(it['name'])] for it in items if normalize_col_name(it['name']) in norm_to_raw]
    
    # 組織集計時に使用する分母列もリストアップ
    spec_denoms = []
    for it in items:
        dtype = it.get('denominator_type')
        norm_n = normalize_col_name(it.get('name'))
        if dtype and dtype != 'SAITO式':
            spec_col = next((c for c in df.columns if norm_n in normalize_col_name(c) and "分母" in c), None)
            if not spec_col and dtype in df.columns: spec_col = dtype
            if spec_col: spec_denoms.append(spec_col)
                
    spec_denoms = list(set(spec_denoms))
    agg_dict = {col: 'sum' for col in item_names + spec_denoms}
    
    if name_col in df.columns: agg_dict[name_col] = 'count'
    if org_level == 'team' and shop_col in df.columns and shop_col != org_col:
        agg_dict[shop_col] = 'first'
        
    org_agg = df.groupby(org_col).agg(agg_dict).reset_index()
    total_score_col = 'Total_Score'
    org_agg[total_score_col] = 0.0
    
    # SAITO式分母の組織集計
    team_saito_denoms = {}
    if scored_df is not None and 'Total_Denominator' in scored_df.columns:
        team_saito_denoms = scored_df.groupby(org_col)['Total_Denominator'].sum().to_dict()
    
    for item in items:
        item_name = item.get('name', '')
        norm_item_name = normalize_col_name(item_name)
        if norm_item_name not in norm_to_raw: continue
        
        target_col = norm_to_raw[norm_item_name]
        weight = float(item.get('weight', 0))
        eval_type = item.get('evaluation_type', 'relative')
        dtype = item.get('denominator_type', 'SAITO式')
        new_col = f"{norm_item_name}_score"
        
        scores = []
        for i, row in org_agg.iterrows():
            my_org = row[org_col]
            my_shop = row[shop_col] if org_level == 'team' and shop_col in org_agg.columns else my_org
            
            val = float(row[target_col])
            
            # --- 1. 組織分母 (Target or Denom Sum) ---
            t_val = 1.0
            if eval_type == 'individual_achievement':
                # 個人達成: チーム/店舗合計 >= 個人目標 × 人数
                shop_targets = item.get('shop_targets', {})
                base_target = float(shop_targets.get(my_shop, 0))
                # 目標が未設定（0件）の場合は評価なし（0点）
                if base_target <= 0:
                    scores.append(0)
                    continue
                count_mem = float(row[name_col]) if name_col in row else 1.0
                org_target = base_target * count_mem
                # 組織判定: 合計実績 >= 組織目標 → 満点 / 未達 → 0点
                score = weight if val >= org_target else 0
                scores.append(score)
                continue
            elif eval_type == 'shop_achievement':
                # shop_achievement は「店舗全体の達成率」で全員同点が付く評価軸。
                # チーム集計のとき、チーム分の実績合計だけでは店舗目標に届かず0点になるバグを防ぐため、
                # scored_df の個人スコア（全員同点のはず）をそのまま引き継ぐ。
                if scored_df is not None and new_col in scored_df.columns and org_col in scored_df.columns:
                    team_score_val = scored_df[scored_df[org_col] == my_org][new_col].dropna()
                    score = float(team_score_val.iloc[0]) if not team_score_val.empty else 0.0
                else:
                    shop_targets = item.get('shop_targets', {})
                    t_val = float(shop_targets.get(my_shop, 0))
                    if t_val <= 0:
                        score = 0.0
                    else:
                        if df_original is not None and shop_col in df_original.columns and target_col in df_original.columns:
                            shop_total = float(df_original[df_original[shop_col] == my_shop][target_col].sum())
                        else:
                            shop_total = val
                        achievement_rate_fb = shop_total / t_val
                        st_targets = item.get('targets', {})
                        st_weights = item.get('weights', {})
                        ar1 = float(st_targets.get('achievement_rate_1', 0) or 0)
                        ar2 = float(st_targets.get('achievement_rate_2', 0) or 0)
                        ar3 = float(st_targets.get('achievement_rate_3', 0) or 0)
                        w1 = float(st_weights.get('w1', weight))
                        w2 = float(st_weights.get('w2', 0))
                        w3 = float(st_weights.get('w3', 0))
                        rate_pct = achievement_rate_fb * 100.0
                        if ar1 > 0 and rate_pct >= ar1: score = w1
                        elif ar2 > 0 and rate_pct >= ar2: score = w2
                        elif ar3 > 0 and rate_pct >= ar3: score = w3
                        else: score = 0.0
                scores.append(score)
                continue
            else:
                if dtype == 'SAITO式':
                    t_val = team_saito_denoms.get(my_org, float(row[name_col]) if name_col in row else 1.0)
                else:
                    spec_col = next((c for c in org_agg.columns if norm_item_name in normalize_col_name(c) and "分母" in c), None)
                    if not spec_col and dtype in org_agg.columns: spec_col = dtype
                    t_val = float(row[spec_col]) if spec_col else 1.0
            
            if t_val <= 0: t_val = 1.0
            
            # --- 2. 達成率の計算 ---
            achievement_rate = val / t_val
            
            # --- 3. スコア判定 ---
            score = 0
            if eval_type in ['absolute', 'shop_achievement', 'individual_achievement', 'team_absolute']:
                st_targets = item.get('targets', {})
                st_weights = item.get('weights', {})
                ar1 = float(st_targets.get('achievement_rate_1', 0) or 0)
                
                if ar1 > 0:
                    rate_pct = achievement_rate * 100.0
                    w1 = float(st_weights.get('w1', weight))
                    w2 = float(st_weights.get('w2', 0))
                    w3 = float(st_weights.get('w3', 0))
                    ar2 = float(st_targets.get('achievement_rate_2', 0) or 0)
                    ar3 = float(st_targets.get('achievement_rate_3', 0) or 0)
                    
                    if rate_pct >= ar1: score = w1
                    elif ar2 > 0 and rate_pct >= ar2: score = w2
                    elif ar3 > 0 and rate_pct >= ar3: score = w3
                elif achievement_rate >= 1.0:
                    score = weight
            
            elif eval_type == 'relative_absolute':
                # 本来は全組織のリストが必要だが、ここでは簡易化（必要なら全体計算を別途回す）
                score = achievement_rate # 仮
            elif eval_type in ['team_relative', 'relative']:
                score = achievement_rate # 効率として記録し、後でランキング
            else:
                score = val # 順位用
            
            scores.append(score)
            
        # ランキング系の後処理
        if eval_type in ['team_relative', 'relative']:
            scores = [calculate_relative_score(s, scores, weight) for s in scores]
        elif eval_type == 'relative_absolute':
            # 全体平均などを考慮した相対絶対評価
            all_rates = scores
            scores = [calculate_relative_absolute_score(r, 1.0, all_rates, [1.0]*len(all_rates), weight) for r in all_rates]

        org_agg[new_col] = scores
        org_agg[total_score_col] += org_agg[new_col]
        
    org_agg['Rank'] = org_agg[total_score_col].rank(ascending=False, method='min')
    
    # リーダー/マネージャーの特定
    if scored_df is not None and name_col in scored_df.columns:
        if org_level == 'team':
            # リーダー判定: user_masterの並び順を直接参照し、チームごとの先頭スタッフをリーダーに確定する
            # groupby.first()への依存を廃止し、CSVの並び順に影響されない確実な方法
            user_master = config.get('_user_master', {})
            if user_master.get('users'):
                import unicodedata, re
                def _norm(t):
                    t = str(t) if t else ''
                    t = unicodedata.normalize('NFKC', t)
                    t = re.sub(r'[\s　]+', '', t)
                    return t
                # チームごとの先頭スタッフを名簿から直接抽出（名簿順が絶対的な優先）
                team_to_leader = {}
                for u in user_master['users']:
                    team_key = _norm(u.get('team', ''))
                    if team_key not in team_to_leader:
                        team_to_leader[team_key] = u.get('name', '')
                # org_aggのチーム名を正規化してマッピング
                org_agg['Leader'] = org_agg[org_col].apply(
                    lambda t: team_to_leader.get(_norm(t), '')
                )
            else:
                # user_masterがない場合はdf_originalから取得（フォールバック）
                leaders = df_original.groupby(org_col, sort=False)[name_col].first().reset_index()
                leaders.rename(columns={name_col: 'Leader'}, inplace=True)
                org_agg = pd.merge(org_agg, leaders, on=org_col, how='left')
        else:
            # 店長判定: ユーザーマスターの 'shop_managers' フィールドから直接取得
            user_master = config.get('_user_master', {})
            if not user_master:
                from utils.data_manager import load_user_master
                user_master = load_user_master()
            
            shop_managers = user_master.get('shop_managers', {})
            org_agg['Manager'] = org_agg[org_col].map(shop_managers).fillna("")
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
    """ 計算プロセスの詳細（監査用） """
    audit_rows = []
    items = config.get('items', [])
    name_col = 'スタッフ名'
    shop_col = '店舗名'
    team_col = 'チーム名'
    
    # カラム名正規化マップ
    raw_cols = df.columns.tolist()
    norm_to_raw = {normalize_col_name(c): c for c in raw_cols}
    
    for item in items:
        item_name = item.get('name', '')
        norm_item_name = normalize_col_name(item_name)
        if norm_item_name not in norm_to_raw: continue
        
        perf_col = norm_to_raw[norm_item_name]
        weight = float(item.get('weight', 0))
        eval_type = item.get('evaluation_type', 'relative')
        denom_type = item.get('denominator_type', 'SAITO式')
        
        # 配点しきい値
        st_targets = item.get('targets', {})
        st_weights = item.get('weights', {})
        ar1 = float(st_targets.get('achievement_rate_1', 0) or 0)
        w1 = float(st_weights.get('w1', weight))
        w2 = float(st_weights.get('w2', 0))
        w3 = float(st_weights.get('w3', 0))
        ar2 = float(st_targets.get('achievement_rate_2', 0) or 0)
        ar3 = float(st_targets.get('achievement_rate_3', 0) or 0)
        
        for idx, row in df.iterrows():
            staff_name = str(row[name_col]).strip() if name_col in row else ""
            val = float(pd.to_numeric(row[perf_col], errors='coerce')) if pd.notna(row[perf_col]) else 0.0
            
            # 分母の特定
            denom = 1.0
            if denom_type != 'SAITO式':
                spec_col = next((c for c in df.columns if norm_item_name in normalize_col_name(c) and "分母" in c), None)
                if not spec_col and denom_type in df.columns: spec_col = denom_type
                if spec_col:
                    d_val = pd.to_numeric(row[spec_col], errors='coerce')
                    if pd.notna(d_val) and d_val > 0: denom = float(d_val)
            elif df_denom is not None:
                d_row = df_denom[df_denom[name_col].str.strip() == staff_name] if name_col in df_denom.columns else pd.DataFrame()
                if not d_row.empty:
                    d_val = d_row.iloc[0].get('成績分母', 1.0)
                    if d_val > 0: denom = float(d_val)
            
            # 目標の特定
            target = 1.0
            val_for_calc = val
            if eval_type in ['shop_achievement', 'individual_achievement'] and shop_col in df.columns:
                shop_name = row[shop_col]
                shop_targets = item.get('shop_targets', {})
                # 店舗目標を直接引用
                target = float(shop_targets.get(shop_name, 0))
                
                if eval_type == 'shop_achievement':
                    # 店舗全員の実績合計
                    shop_total_perf = df[df[shop_col] == shop_name][perf_col].astype(float).sum()
                    val_for_calc = shop_total_perf
                else:
                    val_for_calc = val
                denom = 1.0
            else:
                target = float(st_targets.get(staff_name, 1.0))
            if target <= 0: target = 1.0
            
            # 達成率
            rate = (val_for_calc / denom) / target
            rate_pct = rate * 100.0
            
            # 最終スコアの計算を簡易再現
            score = 0
            if eval_type in ['absolute', 'team_absolute', 'shop_achievement', 'individual_achievement']:
                if ar1 > 0:
                    if rate_pct >= ar1: score = w1
                    elif ar2 > 0 and rate_pct >= ar2: score = w2
                    elif ar3 > 0 and rate_pct >= ar3: score = w3
                elif rate >= 1.0:
                    score = w1 if w1 > 0 else weight
            
            audit_rows.append({
                "対象者": staff_name,
                "店舗": row.get(shop_col, "-"),
                "チーム": row.get(team_col, "-"),
                "項目名": item_name,
                "評価型": eval_type,
                "実績値": val,
                "分母種別": denom_type,
                "計算ターゲット": round(target, 2),
                "計算分母": round(denom, 3),
                "達成率(%)": round(rate_pct, 1),
                "段階(1/2/3)": f"{ar1}/{ar2}/{ar3}" if ar1 > 0 else "100%",
                "配点(1/2/3)": f"{w1}/{w2}/{w3}" if ar1 > 0 else f"{w1}",
                "最終スコア": score
            })
            
    return pd.DataFrame(audit_rows)
