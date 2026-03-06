
import streamlit as st
import pandas as pd
import numpy as np
import io
from utils.data_manager import load_user_master, save_user_master, load_scoring_config, save_scoring_config
from utils.calculation import calculate_scores, calculate_denominators, aggregate_team_scores, aggregate_shop_scores
from data_loader import load_data
from master_data import PERFORMANCE_ITEMS, NEW_FORMAT_COLUMNS

# Admin Password (hardcoded for simplicity as requested)
ADMIN_PASSWORD = "admin123"

def admin_page():
    st.header("管理者設定")
    
    tab1, tab2, tab3 = st.tabs(["メンバー管理", "評価項目設定", "成績データ管理"])
    
    # --- Tab 1: Member Management ---
    with tab1:
        st.subheader("メンバー管理 (User Master)")
        st.markdown("ここで設定した「店舗」「チーム」情報が、成績集計時の正となります。")
        
        # --- File Upload/Download Area ---
        col_dl, col_ul = st.columns(2)
        
        # Template Download
        with col_dl:
            # Create template dataframe
            template_df = pd.DataFrame(columns=["店舗", "チーム", "氏名", "基準時間"])
            # Add a sample row
            template_df.loc[0] = ["イオン天草店", "チームA", "山田太郎", 7.5]
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                template_df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            st.download_button(
                label="登録用フォーマット(Excel)をダウンロード",
                data=buffer.getvalue(),
                file_name="member_upload_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # File Upload
        with col_ul:
            uploaded_master = st.file_uploader("メンバーリストを一括登録 (Excel)", type=["xlsx"])

        # Load current data
        user_data = load_user_master()
        users_list = user_data.get('users', [])
        df_users = pd.DataFrame(users_list)
        
        if df_users.empty:
            df_users = pd.DataFrame(columns=["shop", "team", "name", "standard_hours"])
        else:
            # Ensure correct column order if data exists
            cols = ["shop", "team", "name", "standard_hours"]
            # Filter to user columns present in df
            existing_cols = [c for c in cols if c in df_users.columns]
            df_users = df_users[existing_cols]
        
        # Overwrite with uploaded data if present
        if uploaded_master is not None:
            try:
                uploaded_df = pd.read_excel(uploaded_master)
                # Map Japanese headers to English keys
                column_map = {
                    "氏名": "name",
                    "店舗": "shop",
                    "チーム": "team",
                    "基準時間": "standard_hours"
                }
                # Rename columns
                uploaded_df = uploaded_df.rename(columns=column_map)
                
                # Check if required columns exist (at least name)
                if "name" in uploaded_df.columns:
                    # Keep valid columns
                    valid_cols = [c for c in ["shop", "team", "name", "standard_hours"] if c in uploaded_df.columns]
                    df_users = uploaded_df[valid_cols]
                    st.success("ファイルを読み込みました。下の表で確認・編集し、「保存」ボタンを押してください。")
                else:
                    st.error("アップロードされたファイルに「氏名」列が見つかりません。")
            except Exception as e:
                st.error(f"ファイル読み込みエラー: {e}")

        # Column Config for display (Localization)
        column_config = {
            "name": "氏名",
            "shop": "店舗",
            "team": "チーム",
            "standard_hours": st.column_config.NumberColumn(
                "基準時間",
                min_value=0.0,
                step=0.5,
                format="%.1f"
            )
        }

        edited_users = st.data_editor(
            df_users, 
            column_config=column_config,
            num_rows="dynamic", 
            use_container_width=True
        )
        
        if st.button("メンバー情報を保存"):
            # Convert back to list of dicts and save
            # Drop empty rows if any
            clean_data = edited_users.dropna(how='all').to_dict(orient='records')
            # Ensure standard_hours is numeric
            for u in clean_data:
                try:
                    u['standard_hours'] = float(u.get('standard_hours', 7.5))
                except:
                    u['standard_hours'] = 7.5
                    
            save_user_master({"users": clean_data})
            st.success("メンバー情報を保存しました！")

    # --- Tab 2: Scoring Config ---
    with tab2:
        st.subheader("評価項目・配点設定")
        
        # Get distinct shops for headers
        user_data_master = load_user_master()
        users_list_master = user_data_master.get('users', [])
        distinct_shops = sorted(list(set([u.get('shop') for u in users_list_master if u.get('shop')])))
        
        # --- File Upload/Download Area ---
        col_dl_score, col_ul_score = st.columns(2)
        
        # Template Data for Scoring (Base columns)
        base_template_data = {
            "項目名": ["HS新規", "スマホ総販"],
            "分母種別": ["SAITO式", "SAITO式"],
            "配点1": [60, 100], "配点2": [0, 0], "配点3": [0, 0],
            "評価軸": ["相対評価", "相対評価"],
            "獲得率目標1": [0.5, 0], "獲得率目標2": [0, 0], "獲得率目標3": [0, 0],
            "達成率目標1": [1.0, 0], "達成率目標2": [0, 0], "達成率目標3": [0, 0],
            "個人救済": [False, False], "チーム救済": [False, False]
        }
        # Add dynamic shop columns to template
        for shop in distinct_shops:
            base_template_data[f"目標_{shop}"] = [0, 0]
            
        with col_dl_score:
            template_df_score = pd.DataFrame(base_template_data)
            buffer_score = io.BytesIO()
            with pd.ExcelWriter(buffer_score, engine='openpyxl') as writer:
                template_df_score.to_excel(writer, index=False, sheet_name='ScoringConfig')
            
            st.download_button(
                label="設定用フォーマット(Excel)をダウンロード",
                data=buffer_score.getvalue(),
                file_name="scoring_config_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col_ul_score:
            uploaded_scoring = st.file_uploader("評価設定を一括登録 (Excel)", type=["xlsx"], key="scoring_uploader")

        # Load and Flatten Config
        config = load_scoring_config()
        items = config.get('items', [])
        
        # Mapping for Evaluation Types (Display <-> Internal)
        eval_type_map = {
            "relative": "相対評価", "absolute": "絶対評価", 
            "relative_absolute": "相対絶対", "shop_achievement": "店舗達成率",
            "team_absolute": "チーム絶対", "team_relative": "チーム相対", "team_step": "チーム段階",
            "shop_absolute": "店舗絶対", "shop_step": "店舗段階"
        }
        eval_type_map_rev = {v: k for k, v in eval_type_map.items()}

        flat_items = []
        for item in items:
            etype_display = eval_type_map.get(item.get('evaluation_type'), item.get('evaluation_type'))
            
            flat_item = {
                "項目名": item.get('name', ''),
                "分母種別": item.get('denominator_type', 'SAITO式'),
                "配点1": item.get('weight', 0),
                "配点2": item.get('weights', {}).get('w2', 0),
                "配点3": item.get('weights', {}).get('w3', 0),
                "評価軸": etype_display,
                "獲得率目標1": item.get('targets', {}).get('acquisition_rate_1', 0.0),
                "獲得率目標2": item.get('targets', {}).get('acquisition_rate_2', 0.0),
                "獲得率目標3": item.get('targets', {}).get('acquisition_rate_3', 0.0),
                "達成率目標1": item.get('targets', {}).get('achievement_rate_1', 0.0),
                "達成率目標2": item.get('targets', {}).get('achievement_rate_2', 0.0),
                "達成率目標3": item.get('targets', {}).get('achievement_rate_3', 0.0),
                "個人救済": item.get('flags', {}).get('individual_salvage', False),
                "チーム救済": item.get('flags', {}).get('team_salvage', False),
            }
            
            # Add shop targets
            shop_targets = item.get('shop_targets', {})
            for shop in distinct_shops:
                flat_item[f"目標_{shop}"] = shop_targets.get(shop, 0)
                
            flat_items.append(flat_item)
            
        df_items = pd.DataFrame(flat_items)
        if df_items.empty:
             df_items = pd.DataFrame(columns=base_template_data.keys())
        
        # Ensure all dynamic shop columns exist in dataframe for display
        for shop in distinct_shops:
            col_name = f"目標_{shop}"
            if col_name not in df_items.columns:
                df_items[col_name] = 0

        # Process Upload
        if uploaded_scoring is not None:
            try:
                uploaded_df_sc = pd.read_excel(uploaded_scoring)
                # Basic validation: check if '項目名' exists
                if "項目名" in uploaded_df_sc.columns:
                    # Fill missing columns
                    # Base columns
                    for col in base_template_data.keys():
                        if col not in uploaded_df_sc.columns:
                             # Default values
                            if "配点" in col or "目標" in col:
                                uploaded_df_sc[col] = 0
                            elif "救済" in col:
                                uploaded_df_sc[col] = False
                            else:
                                uploaded_df_sc[col] = ""
                    
                    # Ensure shop columns exist
                    for shop in distinct_shops:
                        col_name = f"目標_{shop}"
                        if col_name not in uploaded_df_sc.columns:
                            uploaded_df_sc[col_name] = 0
                            
                    df_items = uploaded_df_sc
                    st.success("設定ファイルを読み込みました。確認して保存してください。")
                else:
                    st.error("「項目名」列が見つかりません。")
            except Exception as e:
                st.error(f"読み込みエラー: {e}")

        # Column Config
        column_config_score = {
            "項目名": "項目名",
            "分母種別": st.column_config.SelectboxColumn(
                "分母種別", options=["SAITO式", "ドコモデータ", "DM", "独自"], required=True
            ),
            "評価軸": st.column_config.SelectboxColumn(
                "評価軸",
                options=list(eval_type_map.values()),
                required=True
            ),
            "配点1": st.column_config.NumberColumn("配点1", min_value=0, max_value=1000),
            "配点2": st.column_config.NumberColumn("配点2", min_value=0, max_value=1000),
            "配点3": st.column_config.NumberColumn("配点3", min_value=0, max_value=1000),
            "獲得率目標1": st.column_config.NumberColumn("獲得率目標1", format="%.2f"),
            "獲得率目標2": st.column_config.NumberColumn("獲得率目標2", format="%.2f"),
            "獲得率目標3": st.column_config.NumberColumn("獲得率目標3", format="%.2f"),
            "達成率目標1": st.column_config.NumberColumn("達成率目標1", format="%.2f"),
            "達成率目標2": st.column_config.NumberColumn("達成率目標2", format="%.2f"),
            "達成率目標3": st.column_config.NumberColumn("達成率目標3", format="%.2f"),
            "個人救済": "個人救済(ON/OFF)",
            "チーム救済": "チーム救済(ON/OFF)"
        }
        
        # Add column config for shops (Optional formatting)
        for shop in distinct_shops:
            column_config_score[f"目標_{shop}"] = st.column_config.NumberColumn(f"目標_{shop}")
        
        edited_items = st.data_editor(
            df_items, 
            column_config=column_config_score,
            num_rows="dynamic", 
            use_container_width=True
        )
        
        if st.button("評価設定を保存"):
            # Reconstruct nested structure
            new_items = []
            for _, row in edited_items.iterrows():
                # Map Japanese Display back to Internal Key
                internal_etype = eval_type_map_rev.get(row['評価軸'], 'relative')
                
                # Collect shop targets
                shop_targets_data = {}
                for shop in distinct_shops:
                    col_name = f"目標_{shop}"
                    if col_name in row:
                        shop_targets_data[shop] = row[col_name]

                new_item = {
                    "id": f"item_{row['項目名']}", 
                    "name": row['項目名'],
                    "weight": row['配点1'], # Main weight is weight1
                    "weights": { # Multi-weights
                        "w1": row['配点1'],
                        "w2": row['配点2'],
                        "w3": row['配点3']
                    },
                    "evaluation_type": internal_etype,
                    "denominator_type": row['分母種別'],
                    "targets": {
                        "acquisition_rate_1": row.get('獲得率目標1', 0),
                        "acquisition_rate_2": row.get('獲得率目標2', 0),
                        "acquisition_rate_3": row.get('獲得率目標3', 0),
                        "achievement_rate_1": row.get('達成率目標1', 0),
                        "achievement_rate_2": row.get('達成率目標2', 0),
                        "achievement_rate_3": row.get('達成率目標3', 0)
                    },
                    "flags": {
                        "individual_salvage": row['個人救済'],
                        "team_salvage": row['チーム救済']
                    },
                    "shop_targets": shop_targets_data
                }
                new_items.append(new_item)
            
            save_scoring_config({"items": new_items})
            st.success("評価設定を保存しました！")



    # --- Tab 3: Performance Data Management ---
    with tab3:
        st.subheader("成績データ一括登録")
        st.markdown("ここでアップロードしたエクセルデータが、「総合」「件数」「順位」すべてに反映されます。")
        
        # Format Generation
        col_dl_perf, col_ul_perf = st.columns(2)
        
        with col_dl_perf:
            st.markdown("**① フォーマットの取得**")
            # Placeholder to align height
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            perf_template_df = pd.DataFrame(columns=NEW_FORMAT_COLUMNS)
            # Add header styling with openpyxl explicitly in real app, here we just output standard Excel
            buffer_perf = io.BytesIO()
            with pd.ExcelWriter(buffer_perf, engine='openpyxl') as writer:
                perf_template_df.to_excel(writer, index=False, sheet_name='PerformanceData')
                
            st.download_button(
                label="成績データ登録用フォーマット(Excel)をダウンロード",
                data=buffer_perf.getvalue(),
                file_name="performance_data_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col_ul_perf:
            st.markdown("**② データのアップロード**")
            import os
            uploaded_perf = st.file_uploader("成績データをアップロード", label_visibility="collapsed", type=["csv", "xlsx"], key="perf_uploader")
            
            if uploaded_perf is not None:
                # Use load_data logic which applies mapping
                loaded_df, err = load_data(uploaded_perf)
                if err:
                    st.error(err)
                else:
                    # Save persistence to current_performance.csv
                    os.makedirs('data', exist_ok=True)
                    loaded_df.to_csv('data/current_performance.csv', index=False, encoding='utf-8-sig')
                    st.success("成績データをシステムに反映しました！")
                    # Clear sample mode if any
                    st.session_state['use_sample'] = False
                    
        # Debug/Preview area
        import os
        if os.path.exists("data/current_performance.csv"):
             st.markdown("---")
             st.write("▼ 現在登録されているデータ (Preview)")
             st.dataframe(pd.read_csv("data/current_performance.csv").head(10))
             
             st.markdown("---")
             st.subheader("💡 成績分母の計算根拠 (デバッグ用)")
             if st.checkbox("計算プロセスと内訳の詳細を表示する"):
                 try:
                     df_cur = pd.read_csv("data/current_performance.csv")
                     df_denom_debug = get_denominator_df(df_cur)
                     if not df_denom_debug.empty:
                         st.write("各メンバーの「成績分母」および計算に使用されたパラメータの一覧です。")
                         st.dataframe(df_denom_debug)
                     else:
                         st.info("計算可能なデータがありません。")
                 except Exception as e:
                     st.error(f"デバッグ情報の生成エラー: {e}")

        st.markdown("---")
        st.subheader("日別実績データ登録 (分母計算用)")
        st.markdown("後ほど分母計算で使用する、日別のオーダー数・端末販売数のエクセルをアップロードします。")
        
        # We need download buttons for daily formats too
        col_dl_daily, col_ul_daily = st.columns(2)
        
        with col_dl_daily:
            st.markdown("**① 日別フォーマットの取得**")
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            
            # Generate a complex empty format as requested
            daily_cols = ["販売店名", "受付担当者"] + [str(i) for i in range(1, 32)]
            daily_df = pd.DataFrame(columns=daily_cols)
            
            buf_daily = io.BytesIO()
            with pd.ExcelWriter(buf_daily, engine='openpyxl') as w:
                # To simulate the 3 header rows:
                # Row 1: empty prefixes
                # Row 2: empty months
                # Row 3: days 1..31
                # Row 4: headers
                header1 = pd.DataFrame([[""]*len(daily_cols)])
                header2 = pd.DataFrame([[""]*len(daily_cols)])
                header3 = pd.DataFrame([["", ""] + [str(i) for i in range(1, 32)]])
                header4 = pd.DataFrame([daily_cols])
                
                final_df = pd.concat([header1, header2, header3, header4, daily_df], ignore_index=True)
                final_df.to_excel(w, index=False, header=False, sheet_name='DailyDataFormat')
                
            st.download_button(
                label="日別実績登録用フォーマット(Excel)をDL",
                data=buf_daily.getvalue(),
                file_name="daily_performance_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_daily_tmpl"
            )
            
        with col_ul_daily:
            st.markdown("**② 日別データのアップロード**")
            # --- Orders ---
            st.caption("■ 日別オーダー数")
            uploaded_orders = st.file_uploader("日別オーダー数", label_visibility="collapsed", type=["xlsx"], key="order_uploader")
            if uploaded_orders is not None:
                from utils.daily_loader import load_daily_data
                df_orders, err = load_daily_data(uploaded_orders, val_name="order_count")
                if err:
                    st.error(err)
                elif df_orders is not None and not df_orders.empty:
                    os.makedirs('data', exist_ok=True)
                    df_orders.to_csv('data/daily_orders.csv', index=False, encoding='utf-8-sig')
                    st.success(f"{len(df_orders)}件のオーダーデータを保存しました！")
                else:
                    st.warning("有効なデータが見つかりませんでした。")
                    
            # --- Terminals ---
            st.caption("■ 日別端末販売数")
            uploaded_terms = st.file_uploader("日別端末販売数", label_visibility="collapsed", type=["xlsx"], key="term_uploader")
            if uploaded_terms is not None:
                from utils.daily_loader import load_daily_data
                df_terms, err = load_daily_data(uploaded_terms, val_name="terminal_count")
                if err:
                    st.error(err)
                elif df_terms is not None and not df_terms.empty:
                    os.makedirs('data', exist_ok=True)
                    df_terms.to_csv('data/daily_terminals.csv', index=False, encoding='utf-8-sig')
                    st.success(f"{len(df_terms)}件の端末データを保存しました！")
                else:
                    st.warning("有効なデータが見つかりませんでした。")
                    
        # Preview
        if os.path.exists('data/daily_orders.csv'):
            st.write("▼ 日別オーダー数 (Preview)")
            st.dataframe(pd.read_csv('data/daily_orders.csv').head(5))
        if os.path.exists('data/daily_terminals.csv'):
            st.write("▼ 日別端末販売数 (Preview)")
            st.dataframe(pd.read_csv('data/daily_terminals.csv').head(5))



def get_denominator_df(df):
    """ Helper to calculate denominators on the fly processing local CSVs """
    import os
    if os.path.exists('data/daily_orders.csv'):
        df_orders = pd.read_csv('data/daily_orders.csv')
    else:
        df_orders = pd.DataFrame()
        
    if os.path.exists('data/daily_terminals.csv'):
        df_terms = pd.read_csv('data/daily_terminals.csv')
    else:
        df_terms = pd.DataFrame()
        
    user_master = load_user_master()
    df_users = pd.DataFrame(user_master.get('users', []))
    
    return calculate_denominators(df, df_orders, df_terms, df_users)


# --- Page Functions ---

def ranking_page(df):
    st.header("順位表 (Ranking)")
    
    if df is None:
        st.info("データをアップロードしてください。")
        return

    # Load Config
    config = load_scoring_config()
    if not config or not config.get('items'):
        st.warning("評価設定がありません。管理者メニューから設定してください。")
        st.dataframe(df)
        return

    # Calculate Denominators
    df_denom = get_denominator_df(df)
    
    # Calculate Scores with weights
    scored_df = calculate_scores(df, config, df_denom)
    
    # --- Robust Merge with User Master ---
    user_master = load_user_master()
    if user_master.get('users'):
        df_master = pd.DataFrame(user_master['users'])
        
        # Identify columns
        name_col_score = 'スタッフ名' if 'スタッフ名' in scored_df.columns else 'Name'
        shop_col_score = '店舗名' if '店舗名' in scored_df.columns else 'Shop'
        team_col_score = 'チーム名' if 'チーム名' in scored_df.columns else 'Team'
        
        # Merge if possible
        if name_col_score in scored_df.columns and 'name' in df_master.columns:
            # Perform merge (keep all score rows)
            merged_df = pd.merge(scored_df, df_master[['name', 'shop', 'team']], 
                               left_on=name_col_score, right_on='name', 
                               how='left', suffixes=('', '_master'))
            
            # Prefer Master Data (shop_master) > Uploaded Data (shop_col_score)
            # Create standardized '店舗' and 'チーム' columns
            
            # Shop Logic
            if 'shop' in merged_df.columns:
                # If master has shop, use it. If nan, fallback to upload.
                merged_df['店舗名'] = merged_df['shop'].fillna(merged_df[shop_col_score] if shop_col_score in merged_df.columns else pd.NA)
            elif shop_col_score in merged_df.columns:
                merged_df['店舗名'] = merged_df[shop_col_score]
            else:
                merged_df['店舗名'] = "未所属"
                
            # Team Logic
            if 'team' in merged_df.columns:
                merged_df['チーム名'] = merged_df['team'].fillna(merged_df[team_col_score] if team_col_score in merged_df.columns else pd.NA)
            elif team_col_score in merged_df.columns:
                merged_df['チーム名'] = merged_df[team_col_score]
            else:
                merged_df['チーム名'] = "未所属"
                
            # Clean up redundant columns if desired, but keeping them for debug is safer.
            # Update scored_df
            scored_df = merged_df
        else:
            # If no merge possible, just ensure standardized columns exist
            if shop_col_score != '店舗' and shop_col_score in scored_df.columns:
                scored_df['店舗名'] = scored_df[shop_col_score]
            if team_col_score != 'チーム' and team_col_score in scored_df.columns:
                scored_df['チーム名'] = scored_df[team_col_score]

    # Ensure items exist
    items = config.get('items', [])
    item_names = [item.get('name') for item in items if item.get('name')]
    
    name_col = 'スタッフ名' if 'スタッフ名' in scored_df.columns else 'Name'
    shop_col = '店舗名' if '店舗名' in scored_df.columns else 'Shop'
    team_col = 'チーム名' if 'チーム名' in scored_df.columns else 'Team'

    # --- Main Views ---
    tab_overall, tab_category = st.tabs(["総合ランキング", "部門別ランキング"])

    # 1. Overall Ranking Tab
    with tab_overall:
        col_left, col_right = st.columns(2, gap="large")
        
        with col_left:
            st.subheader("個人総合ランキング")
            if 'Rank' in scored_df.columns:
                disp_cols = ['Rank', name_col, 'Total_Score', shop_col, team_col]
                valid_item_cols = [col for col in item_names if col in scored_df.columns]
                for col in valid_item_cols:
                    if col not in disp_cols:
                        disp_cols.append(col)
                
                valid_cols = [c for c in disp_cols if c in scored_df.columns]
                
                show_df = scored_df.sort_values("Rank")[valid_cols].copy()
                show_df['Total_Score'] = show_df['Total_Score'].map('{:.0f}点'.format)
                
                # Rename columns back to Japanese standard if needed
                rename_dict = {}
                if name_col != 'スタッフ名': rename_dict[name_col] = 'スタッフ名'
                if shop_col != '店舗名': rename_dict[shop_col] = '店舗名'
                if team_col != 'チーム名': rename_dict[team_col] = 'チーム名'
                show_df.rename(columns=rename_dict, inplace=True)
                
                st.dataframe(show_df, hide_index=True, use_container_width=True, height=625)
            else:
                st.error("ランキング計算エラー")

        with col_right:
            st.subheader("チーム総合ランキング")
            team_df = aggregate_team_scores(scored_df, df_original=df, config=config)
            if not team_df.empty:
                t_col = 'チーム' if 'チーム' in team_df.columns else 'Team'
                s_col = '店舗' if '店舗' in team_df.columns else 'Shop'
                
                disp_cols = ['Rank', t_col, 'Leader', 'Total_Score', s_col]
                valid_cols = [c for c in disp_cols if c in team_df.columns]
                
                show_team = team_df[valid_cols].copy()
                show_team['Total_Score'] = show_team['Total_Score'].map('{:.0f}点'.format)
                show_team.rename(columns={'Leader': '最高得点者'}, inplace=True)
                
                st.dataframe(show_team, hide_index=True, use_container_width=True, height=270)
            else:
                st.info("チーム情報がありません")

            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            
            st.subheader("店舗総合ランキング")
            shop_df = aggregate_shop_scores(scored_df, df_original=df, config=config)
            if not shop_df.empty:
                s_col = '店舗' if '店舗' in shop_df.columns else 'Shop'
                m_col = 'Manager'
                
                disp_cols = ['Rank', s_col, m_col, 'Total_Score']
                valid_cols = [c for c in disp_cols if c in shop_df.columns]
                
                show_shop = shop_df[valid_cols].copy()
                show_shop['Total_Score'] = show_shop['Total_Score'].map('{:.0f}点'.format)
                show_shop.rename(columns={m_col: '最高得点者'}, inplace=True)
                
                st.dataframe(show_shop, hide_index=True, use_container_width=True, height=270)
            else:
                st.info("店舗情報がありません")

    # 2. Category Ranking Tab
    with tab_category:
        st.subheader("部門別ランキング (Top 5)")
        if not item_names:
            st.info("評価項目が設定されていません")
        else:
            cat_tabs = st.tabs(item_names)
            for i, item_name in enumerate(item_names):
                with cat_tabs[i]:
                    if item_name not in scored_df.columns:
                        st.warning(f"「{item_name}」データがありません")
                        continue
                    
                    sorted_df = scored_df.sort_values(item_name, ascending=False).head(5)
                    
                    summary_rows = []
                    rank = 1
                    for _, row in sorted_df.iterrows():
                        val = row[item_name]
                        summary_rows.append({
                            "順位": f"{rank}位",
                            "スタッフ名": row.get(name_col, ''),
                            "実績": val,
                            "店舗名": row.get(shop_col, ''),
                            "チーム名": row.get(team_col, '')
                        })
                        rank += 1
                    st.dataframe(pd.DataFrame(summary_rows), hide_index=True, use_container_width=True)


def count_page(df):
    st.header("件数")
    
    if df is None:
        st.info("データをアップロードしてください。")
        return

    # Load scoring config just to get the list of items to display
    config = load_scoring_config()
    if not config or not config.get('items'):
         st.warning("評価設定がありません。設定を行ってください。")
         st.dataframe(df)
         return
         
    items = config.get('items', [])
    item_names = [item.get('name') for item in items if item.get('name')]
    
    # Optional: also merge with user master to ensure shop/team names are standardized
    user_master = load_user_master()
    if user_master.get('users'):
        df_master = pd.DataFrame(user_master['users'])
        name_col = 'スタッフ名' if 'スタッフ名' in df.columns else 'Name'
        shop_col = '店舗名' if '店舗名' in df.columns else 'Shop'
        team_col = 'チーム名' if 'チーム名' in df.columns else 'Team'
        
        if name_col in df.columns and 'name' in df_master.columns:
            merged_df = pd.merge(df, df_master[['name', 'shop', 'team']], 
                               left_on=name_col, right_on='name', how='left')
            if 'shop' in merged_df.columns:
                merged_df['店舗名'] = merged_df['shop'].fillna(merged_df[shop_col] if shop_col in merged_df.columns else pd.NA)
            elif shop_col in merged_df.columns:
                merged_df['店舗名'] = merged_df[shop_col]
            else:
                merged_df['店舗名'] = "未所属"
                
            if 'team' in merged_df.columns:
                merged_df['チーム名'] = merged_df['team'].fillna(merged_df[team_col] if team_col in merged_df.columns else pd.NA)
            elif team_col in merged_df.columns:
                merged_df['チーム名'] = merged_df[team_col]
            else:
                merged_df['チーム名'] = "未所属"
            
            # Use merged dataframe for calculations
            base_df = merged_df
        else:
            base_df = df.copy()
            if shop_col != '店舗名' and shop_col in base_df.columns: base_df['店舗名'] = base_df[shop_col]
            if team_col != 'チーム名' and team_col in base_df.columns: base_df['チーム名'] = base_df[team_col]
    else:
        base_df = df.copy()
        shop_col = '店舗' if '店舗' in base_df.columns else 'Shop'
        team_col = 'チーム' if 'チーム' in base_df.columns else 'Team'
        if shop_col != '店舗名' and shop_col in base_df.columns: base_df['店舗名'] = base_df[shop_col]
        if team_col != 'チーム名' and team_col in base_df.columns: base_df['チーム名'] = base_df[team_col]

    name_col = 'スタッフ名' if 'スタッフ名' in base_df.columns else 'Name'

    # 1. 個人件数表 (Top Section)
    st.subheader("個人件数")
    
    # Select columns to display: Name, Shop, Team + Item counts
    disp_cols = [name_col, '店舗名', 'チーム名']
    valid_items = [i for i in item_names if i in base_df.columns]
    disp_cols.extend(valid_items)
    
    # Filter available columns
    actual_disp_cols = [c for c in disp_cols if c in base_df.columns]
    
    show_personal = base_df[actual_disp_cols].copy()
    
    # Rename for display if needed
    rename_dict = {}
    if name_col != 'スタッフ名': rename_dict[name_col] = 'スタッフ名'
    show_personal.rename(columns=rename_dict, inplace=True)
    
    # Fill NA with 0 for count columns
    for col in valid_items:
        if col in show_personal.columns:
            show_personal[col] = pd.to_numeric(show_personal[col], errors='coerce').fillna(0)
    
    st.dataframe(show_personal, hide_index=True, use_container_width=True, height=400)

    # 2. チーム件数表 (Middle Section, Red Tint)
    st.subheader("チーム件数")
    if 'チーム名' in base_df.columns:
        # Aggregate logic
        agg_cols = {col: 'sum' for col in valid_items}
        # Include Shop for display if available
        if '店舗名' in base_df.columns:
            try:
                team_counts = base_df.groupby(['チーム名', '店舗名'], as_index=False).agg(agg_cols)
            except:
                team_counts = base_df.groupby('チーム名', as_index=False).agg(agg_cols)
        else:
            team_counts = base_df.groupby('チーム名', as_index=False).agg(agg_cols)
        
        # Sort by first item for some rough ordering, or keep standard
        if valid_items:
             team_counts = team_counts.sort_values(by=valid_items[0], ascending=False)
             
        def highlight_team(val): return 'background-color: #FFF0F2' # Light Pink
        styled_team = team_counts.style.map(highlight_team)
        st.dataframe(styled_team, hide_index=True, use_container_width=True, height=250)
    else:
        st.info("チーム情報がありません")

    # 3. 店舗件数表 (Bottom Section, Green Tint)
    st.subheader("店舗件数")
    if '店舗名' in base_df.columns:
        agg_cols = {col: 'sum' for col in valid_items}
        shop_counts = base_df.groupby('店舗名', as_index=False).agg(agg_cols)
        
        if valid_items:
             shop_counts = shop_counts.sort_values(by=valid_items[0], ascending=False)
             
        def highlight_shop(val): return 'background-color: #F0FFF0' # Light Green
        styled_shop = shop_counts.style.map(highlight_shop)
        st.dataframe(styled_shop, hide_index=True, use_container_width=True, height=250)
    else:
        st.info("店舗情報がありません")

def comprehensive_page(df):
    st.header("総合")
    
    if df is None:
        st.info("データをアップロードしてください。")
        return

    # Check for scoring config to display results
    config = load_scoring_config()
    if not config or not config.get('items'):
         st.warning("評価設定がありません。設定を行ってください。")
         st.dataframe(df)
         return
         
    df_denom = get_denominator_df(df)
    scored_df = calculate_scores(df, config, df_denom)
    
    # Setup base columns
    name_col = 'スタッフ名' if 'スタッフ名' in scored_df.columns else 'Name'
    shop_col = '店舗名' if '店舗名' in scored_df.columns else 'Shop'
    team_col = 'チーム名' if 'チーム名' in scored_df.columns else 'Team'
    
    # 1. 個人成績表 (Top Section, Full width)
    st.subheader("個人成績")
    if 'Rank' in scored_df.columns:
        # Collect display columns (Rank, Name, Team, Shop, Total_Score + Items)
        items = config.get('items', [])
        item_names = [item.get('name') for item in items if item.get('name')]
        
        disp_cols = ['Rank', name_col, 'Total_Score', shop_col, team_col]
        valid_item_cols = [col for col in item_names if col in scored_df.columns]
        for col in valid_item_cols:
            if col not in disp_cols:
                disp_cols.append(col)
                
        valid_cols = [c for c in disp_cols if c in scored_df.columns]
        show_df = scored_df.sort_values("Rank")[valid_cols].copy()
        
        # Formatting
        show_df['Total_Score'] = show_df['Total_Score'].map('{:.1f}'.format)
        
        # Display as un-styled dataframe first (Streamlit native card style applies)
        st.dataframe(show_df, hide_index=True, use_container_width=True, height=400)
    else:
        st.error("ランキング計算エラー")

    # 2. チーム成績表 (Middle Section, Full width + Red Tint)
    st.subheader("チーム成績")
    team_df = aggregate_team_scores(scored_df, df_original=df, config=config)
    if not team_df.empty:
        t_col = 'チーム' if 'チーム' in team_df.columns else 'Team'
        s_col = '店舗' if '店舗' in team_df.columns else 'Shop'
        
        disp_cols_team = ['Rank', t_col, 'Leader', 'Total_Score', s_col]
        valid_cols_team = [c for c in disp_cols_team if c in team_df.columns]
        
        show_team = team_df[valid_cols_team].copy()
        show_team['Total_Score'] = show_team['Total_Score'].map('{:.1f}'.format)
        show_team.rename(columns={'Leader': '最高得点者'}, inplace=True)
        
        # Apply Pandas Styler for background color
        def highlight_team(val): return 'background-color: #FFF0F2' # Very light red/pink
        styled_team = show_team.style.map(highlight_team)
        
        st.dataframe(styled_team, hide_index=True, use_container_width=True, height=250)
    else:
        st.info("チーム情報がありません")

    # 3. 店舗成績表 (Bottom Section, Full width + Green Tint)
    st.subheader("店舗成績")
    shop_df = aggregate_shop_scores(scored_df, df_original=df, config=config)
    if not shop_df.empty:
        s_col = '店舗' if '店舗' in shop_df.columns else 'Shop'
        m_col = 'Manager'
        
        disp_cols_shop = ['Rank', s_col, m_col, 'Total_Score']
        valid_cols_shop = [c for c in disp_cols_shop if c in shop_df.columns]
        
        show_shop = shop_df[valid_cols_shop].copy()
        show_shop['Total_Score'] = show_shop['Total_Score'].map('{:.1f}'.format)
        show_shop.rename(columns={m_col: '最高得点者'}, inplace=True)
        
        # Apply Pandas Styler for background color
        def highlight_shop(val): return 'background-color: #F0FFF0' # Very light green
        styled_shop = show_shop.style.map(highlight_shop)
        
        st.dataframe(styled_shop, hide_index=True, use_container_width=True, height=250)
    else:
        st.info("店舗情報がありません")

# =====================================================
# 成績詳細ページ
# =====================================================
def individual_detail_page(df):
    st.header("📋 成績詳細")

    if df is None:
        st.info("データをアップロードしてください。")
        return

    config = load_scoring_config()
    if not config or not config.get('items'):
        st.warning("評価設定がありません。管理者メニューから設定してください。")
        return

    df_denom = get_denominator_df(df)
    scored_df = calculate_scores(df, config, df_denom)

    # ユーザーマスターとマージ
    user_master = load_user_master()
    if user_master.get('users'):
        df_master = pd.DataFrame(user_master['users'])
        name_col = 'スタッフ名' if 'スタッフ名' in scored_df.columns else 'Name'
        if name_col in scored_df.columns and 'name' in df_master.columns:
            merged = pd.merge(scored_df, df_master[['name', 'shop', 'team']],
                              left_on=name_col, right_on='name', how='left')
            if 'shop' in merged.columns:
                merged['店舗名'] = merged['shop'].fillna(merged.get('店舗名', ''))
            if 'team' in merged.columns:
                merged['チーム名'] = merged['team'].fillna(merged.get('チーム名', ''))
            scored_df = merged

    name_col = 'スタッフ名' if 'スタッフ名' in scored_df.columns else 'Name'
    shop_col = '店舗名' if '店舗名' in scored_df.columns else 'Shop'
    team_col = 'チーム名' if 'チーム名' in scored_df.columns else 'Team'
    items = config.get('items', [])
    item_names = [item.get('name') for item in items if item.get('name')]

    # スタッフ選択
    staff_list = sorted(scored_df[name_col].dropna().unique().tolist())
    selected_staff = st.selectbox("スタッフを選択", staff_list)

    if not selected_staff:
        return

    row = scored_df[scored_df[name_col] == selected_staff].iloc[0]

    # プロフィールカード
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("総合順位", f"{int(row.get('Rank', 0))}位")
    col2.metric("総合スコア", f"{row.get('Total_Score', 0):.1f}点")
    col3.metric("店舗", str(row.get(shop_col, '-')))
    col4.metric("チーム", str(row.get(team_col, '-')))

    st.markdown("---")
    st.subheader("📊 項目別スコア詳細")

    # 項目別スコアテーブル
    detail_rows = []
    for item in items:
        iname = item.get('name', '')
        weight = item.get('weight', 0)
        score_col = f"{iname}_score"
        perf_val = row.get(iname, 0)
        score_val = row.get(score_col, 0)
        achievement = (score_val / weight * 100) if weight > 0 else 0
        detail_rows.append({
            "項目名": iname,
            "実績": perf_val if not pd.isna(perf_val) else 0,
            "獲得点数": round(score_val, 1) if not pd.isna(score_val) else 0,
            "配点(満点)": weight,
            "達成率": f"{achievement:.0f}%",
            "評価軸": item.get('evaluation_type', '-')
        })

    detail_df = pd.DataFrame(detail_rows)
    st.dataframe(detail_df, hide_index=True, use_container_width=True)

    # 分母デバッグ情報
    if df_denom is not None and not df_denom.empty:
        denom_row = df_denom[df_denom[name_col] == selected_staff]
        if not denom_row.empty:
            st.markdown("---")
            st.subheader("🔢 成績分母の内訳")
            denom_disp = denom_row.T.reset_index()
            denom_disp.columns = ["項目", "値"]
            st.dataframe(denom_disp, hide_index=True, use_container_width=True)


# =====================================================
# 分析ページ
# =====================================================
def analysis_page(df):
    st.header("📈 分析")

    if df is None:
        st.info("データをアップロードしてください。")
        return

    config = load_scoring_config()
    if not config or not config.get('items'):
        st.warning("評価設定がありません。管理者メニューから設定してください。")
        return

    df_denom = get_denominator_df(df)
    scored_df = calculate_scores(df, config, df_denom)

    # ユーザーマスターとマージ
    user_master = load_user_master()
    if user_master.get('users'):
        df_master = pd.DataFrame(user_master['users'])
        name_col_s = 'スタッフ名' if 'スタッフ名' in scored_df.columns else 'Name'
        if name_col_s in scored_df.columns and 'name' in df_master.columns:
            merged = pd.merge(scored_df, df_master[['name', 'shop', 'team']],
                              left_on=name_col_s, right_on='name', how='left')
            if 'shop' in merged.columns:
                merged['店舗名'] = merged['shop'].fillna(merged.get('店舗名', ''))
            if 'team' in merged.columns:
                merged['チーム名'] = merged['team'].fillna(merged.get('チーム名', ''))
            scored_df = merged

    name_col = 'スタッフ名' if 'スタッフ名' in scored_df.columns else 'Name'
    shop_col = '店舗名' if '店舗名' in scored_df.columns else 'Shop'
    team_col = 'チーム名' if 'チーム名' in scored_df.columns else 'Team'
    items = config.get('items', [])
    item_names = [item.get('name') for item in items if item.get('name')]

    tab_shop, tab_team, tab_item = st.tabs(["店舗別分析", "チーム別分析", "項目別分析"])

    with tab_shop:
        st.subheader("店舗別 平均スコア")
        if shop_col in scored_df.columns:
            shop_summary = scored_df.groupby(shop_col)['Total_Score'].agg(['mean', 'max', 'min', 'count']).reset_index()
            shop_summary.columns = ['店舗', '平均スコア', '最高スコア', '最低スコア', '人数']
            shop_summary['平均スコア'] = shop_summary['平均スコア'].round(1)
            shop_summary = shop_summary.sort_values('平均スコア', ascending=False)
            st.dataframe(shop_summary, hide_index=True, use_container_width=True)

            st.subheader("店舗別 項目別合計")
            valid_items = [i for i in item_names if i in scored_df.columns]
            if valid_items and shop_col in scored_df.columns:
                shop_item_df = scored_df.groupby(shop_col)[valid_items].sum().reset_index()
                st.dataframe(shop_item_df, hide_index=True, use_container_width=True)
        else:
            st.info("店舗情報がありません")

    with tab_team:
        st.subheader("チーム別 平均スコア")
        if team_col in scored_df.columns:
            team_summary = scored_df.groupby(team_col)['Total_Score'].agg(['mean', 'max', 'min', 'count']).reset_index()
            team_summary.columns = ['チーム', '平均スコア', '最高スコア', '最低スコア', '人数']
            team_summary['平均スコア'] = team_summary['平均スコア'].round(1)
            team_summary = team_summary.sort_values('平均スコア', ascending=False)
            st.dataframe(team_summary, hide_index=True, use_container_width=True)

            st.subheader("チーム別 項目別合計")
            valid_items = [i for i in item_names if i in scored_df.columns]
            if valid_items and team_col in scored_df.columns:
                team_item_df = scored_df.groupby(team_col)[valid_items].sum().reset_index()
                st.dataframe(team_item_df, hide_index=True, use_container_width=True)
        else:
            st.info("チーム情報がありません")

    with tab_item:
        st.subheader("項目別 実績分布")
        valid_items = [i for i in item_names if i in scored_df.columns]
        for iname in valid_items:
            col_l, col_r = st.columns([2, 1])
            with col_l:
                item_data = pd.to_numeric(scored_df[iname], errors='coerce').dropna()
                if not item_data.empty:
                    stats = {
                        "合計": item_data.sum(),
                        "平均": item_data.mean(),
                        "最大": item_data.max(),
                        "最小": item_data.min(),
                        "標準偏差": item_data.std()
                    }
                    st.markdown(f"**{iname}**")
                    st.dataframe(
                        pd.DataFrame([{k: round(v, 2) for k, v in stats.items()}]),
                        hide_index=True, use_container_width=True
                    )
            with col_r:
                score_col = f"{iname}_score"
                if score_col in scored_df.columns:
                    score_data = pd.to_numeric(scored_df[score_col], errors='coerce').dropna()
                    if not score_data.empty:
                        cfg_item = next((it for it in items if it.get('name') == iname), {})
                        w = cfg_item.get('weight', 1)
                        avg_achievement = (score_data.mean() / w * 100) if w > 0 else 0
                        st.metric(f"{iname} 平均達成率", f"{avg_achievement:.0f}%")
            st.markdown("---")


# =====================================================
# 配点シミュレーションページ
# =====================================================
def simulation_page(df):
    st.header("🎯 配点シミュレーション")
    st.markdown("評価設定を仮で変更して、スコアへの影響をシミュレーションします。実際の設定は変更されません。")

    if df is None:
        st.info("データをアップロードしてください。")
        return

    config = load_scoring_config()
    if not config or not config.get('items'):
        st.warning("評価設定がありません。管理者メニューから設定してください。")
        return

    items = config.get('items', [])
    item_names = [item.get('name') for item in items]

    st.subheader("⚙️ 配点の仮変更")
    st.markdown("各項目の配点を調整して「シミュレーション実行」を押してください。")

    import copy
    sim_config = copy.deepcopy(config)
    sim_items = sim_config.get('items', [])

    cols_per_row = 3
    sim_weights = {}

    rows = [item_names[i:i+cols_per_row] for i in range(0, len(item_names), cols_per_row)]
    for row_items in rows:
        cols = st.columns(cols_per_row)
        for j, iname in enumerate(row_items):
            item = next((it for it in sim_items if it.get('name') == iname), {})
            current_w = item.get('weight', 0)
            new_w = cols[j].number_input(
                f"{iname} 配点",
                min_value=0, max_value=500,
                value=int(current_w),
                step=5,
                key=f"sim_{iname}"
            )
            sim_weights[iname] = new_w

    # 合計配点表示
    total_sim = sum(sim_weights.values())
    st.info(f"シミュレーション合計配点: **{total_sim}点**")

    if st.button("▶ シミュレーション実行", type="primary"):
        # sim_config に新配点を適用
        for item in sim_items:
            iname = item.get('name')
            if iname in sim_weights:
                item['weight'] = sim_weights[iname]
                if 'weights' in item:
                    item['weights']['w1'] = sim_weights[iname]

        df_denom = get_denominator_df(df)
        sim_scored = calculate_scores(df, sim_config, df_denom)

        name_col = 'スタッフ名' if 'スタッフ名' in sim_scored.columns else 'Name'

        st.subheader("📊 シミュレーション結果")
        sim_disp_cols = ['Rank', name_col, 'Total_Score'] + [
            f"{n}_score" for n in item_names if f"{n}_score" in sim_scored.columns
        ]
        valid_cols = [c for c in sim_disp_cols if c in sim_scored.columns]
        show = sim_scored.sort_values('Rank')[valid_cols].copy()
        show['Total_Score'] = show['Total_Score'].round(1)
        st.dataframe(show, hide_index=True, use_container_width=True)

        # 現行vs シミュレーション比較
        st.subheader("📉 現行 vs シミュレーション 比較")
        orig_scored = calculate_scores(df, config, df_denom)
        compare = pd.merge(
            orig_scored[[name_col, 'Total_Score', 'Rank']].rename(
                columns={'Total_Score': '現行スコア', 'Rank': '現行順位'}
            ),
            sim_scored[[name_col, 'Total_Score', 'Rank']].rename(
                columns={'Total_Score': 'SIMスコア', 'Rank': 'SIM順位'}
            ),
            on=name_col, how='outer'
        )
        compare['スコア差'] = (compare['SIMスコア'] - compare['現行スコア']).round(1)
        compare['順位差'] = (compare['現行順位'] - compare['SIM順位']).astype(int)
        compare = compare.sort_values('SIM順位')
        st.dataframe(compare, hide_index=True, use_container_width=True)


# =====================================================
# 年間表彰ページ
# =====================================================
def annual_awards_page(df):
    st.header("🏆 年間表彰")
    st.markdown("月次成績データを累積して、年間の優秀者を表彰します。")

    import os, glob

    # 月次データのアップロード
    st.subheader("📁 月次データ管理")
    col_ul, col_dl = st.columns(2)

    with col_ul:
        month_label = st.text_input("月ラベル (例: 2026_01)", placeholder="YYYY_MM")
        uploaded_month = st.file_uploader("月次成績データをアップロード", type=["csv", "xlsx"], key="annual_uploader")
        if uploaded_month is not None and month_label:
            loaded_df, err = load_data(uploaded_month)
            if err:
                st.error(err)
            else:
                os.makedirs('data/annual', exist_ok=True)
                save_path = f'data/annual/{month_label}.csv'
                loaded_df.to_csv(save_path, index=False, encoding='utf-8-sig')
                st.success(f"{month_label} のデータを保存しました。")

    with col_dl:
        annual_files = sorted(glob.glob('data/annual/*.csv'))
        if annual_files:
            st.markdown("**登録済み月次データ:**")
            for f in annual_files:
                fname = os.path.basename(f).replace('.csv', '')
                st.markdown(f"- {fname}")
        else:
            st.info("月次データがまだ登録されていません。")

    if not annual_files:
        return

    st.markdown("---")
    st.subheader("📊 年間累積集計")

    # 配点設定の読み込み
    config = load_scoring_config()
    if not config or not config.get('items'):
        st.warning("評価設定がありません。管理者メニューから設定してください。")
        return

    items = config.get('items', [])
    name_col = None

    all_monthly = []
    for fpath in annual_files:
        try:
            month_df = pd.read_csv(fpath)
            month_name = os.path.basename(fpath).replace('.csv', '')
            df_denom = get_denominator_df(month_df)
            scored = calculate_scores(month_df, config, df_denom)
            # name_col の特定
            if not name_col:
                name_col = 'スタッフ名' if 'スタッフ名' in scored.columns else 'Name'
            scored['月'] = month_name
            all_monthly.append(scored[[name_col, 'Total_Score', '月']])
        except Exception as e:
            st.warning(f"{fpath} の読み込みエラー: {e}")

    if not all_monthly:
        st.error("有効なデータがありません。")
        return

    if not name_col:
        name_col = 'スタッフ名'

    combined = pd.concat(all_monthly, ignore_index=True)

    # 月別スコア一覧
    pivot = combined.pivot_table(index=name_col, columns='月', values='Total_Score', aggfunc='sum').fillna(0)
    pivot['年間合計'] = pivot.sum(axis=1)
    pivot['参加月数'] = (combined.groupby(name_col)['月'].nunique())
    pivot['月平均'] = (pivot['年間合計'] / pivot['参加月数']).round(1)
    pivot = pivot.sort_values('年間合計', ascending=False).reset_index()
    pivot.insert(0, '順位', range(1, len(pivot)+1))

    st.dataframe(pivot, hide_index=True, use_container_width=True)

    # TOP3 表彰
    st.markdown("---")
    st.subheader("🥇 年間表彰")
    medals = ["🥇 1位", "🥈 2位", "🥉 3位"]
    top3 = pivot.head(3)
    cols = st.columns(3)
    for i, (_, row) in enumerate(top3.iterrows()):
        if i >= 3:
            break
        with cols[i]:
            st.markdown(f"### {medals[i]}")
            st.markdown(f"**{row[name_col]}**")
            st.metric("年間合計", f"{row['年間合計']:.1f}点")
            st.metric("月平均", f"{row['月平均']:.1f}点")

    # Excel エクスポート
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        pivot.to_excel(writer, index=False, sheet_name='年間集計')
        combined.to_excel(writer, index=False, sheet_name='月別明細')
    st.download_button(
        label="📥 年間集計をExcelでダウンロード",
        data=buf.getvalue(),
        file_name="annual_awards.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def main():
    st.set_page_config(page_title="ブライトパスくん", layout="wide")
    st.title("ブライトパスくん")
    
    # CSS injection
    # CSS injection for Material Design & Strict Layout
    st.markdown('''
    <style>
        /* Custom Google-like styling */
        .main {
            background-color: #F8F9FA;
            font-family: 'Roboto', 'Noto Sans JP', sans-serif;
        }
        h1, h2, h3, h4 {
            color: #202124;
            font-weight: 500;
        }
        /* Buttons */
        .stButton>button {
            background-color: #1A73E8;
            color: white;
            border-radius: 4px;
            border: none;
            padding: 8px 24px;
            font-weight: 500;
            transition: box-shadow 0.2s;
        }
        .stButton>button:hover {
            box-shadow: 0 1px 2px 0 rgba(60,64,67,0.3), 0 1px 3px 1px rgba(60,64,67,0.15);
            background-color: #1b66c9;
            color: white;
        }
        /* Tabs */
        .stTabs [data-baseweb="tab-list"] {
            gap: 24px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 50px;
            white-space: pre-wrap;
            background-color: transparent;
            border-radius: 4px 4px 0 0;
            gap: 1px;
            padding-top: 10px;
            padding-bottom: 10px;
            font-weight: 500;
        }
        /* Upload UI Japanese localization (Aggressive CSS Override) */
        [data-testid="stFileUploadDropzone"] > div > small,
        [data-testid="stFileUploaderDropzone"] > div > small,
        [data-testid="stFileUploadDropzone"] > div > div > span,
        [data-testid="stFileUploaderDropzone"] > div > div > span {
            display: none !important;
            visibility: hidden !important;
        }
        [data-testid="stFileUploadDropzone"] > div::before,
        [data-testid="stFileUploaderDropzone"] > div::before {
            content: "エクセルをここにドラッグ＆ドロップ \A (または右のボタンから選択)";
            white-space: pre;
            display: block;
            text-align: center;
            font-size: 14px;
            color: #4A4A4A;
            margin-bottom: 8px;
        }
        [data-testid="stFileUploadDropzone"] > div::after,
        [data-testid="stFileUploaderDropzone"] > div::after {
            content: "上限 200MB / CSV, XLSX";
            display: block;
            text-align: center;
            font-size: 12px;
            color: #8C8C8C;
        }
        /* Import Google Fonts */
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&family=Noto+Sans+JP:wght@300;400;500;700&display=swap');

        /* Global Font Settings */
        html, body, [class*="css"] {
            padding: 10px 24px !important;
            font-weight: 600 !important;
            border: none !important;
            background-color: #1A73E8 !important; /* Google Blue */
            color: #FFFFFF !important;
            box-shadow: 0 1px 2px 0 rgba(60,64,67,0.3) !important;
            transition: all 0.2s ease-in-out !important;
            width: 100% !important; /* Sidebarに合わせる */
        }
        .stButton > button:hover {
            box-shadow: 0 2px 4px 1px rgba(60,64,67,0.2) !important;
            background-color: #1765CC !important;
            transform: translateY(-1px);
        }
        
        /* Outline/Secondary button for Admin Login/Logout etc to not be fully blue if needed, 
           but per request we try to make them look like the image. We'll leave them unified for now. */

        /* Download Button specific */
        .stDownloadButton > button {
            border-radius: 6px !important;
            font-weight: 600 !important;
            background-color: #FFFFFF !important;
            color: #1A73E8 !important;
            border: 1px solid #DADCE0 !important;
            width: 100% !important;
        }
        .stDownloadButton > button:hover {
            background-color: #F8FAFD !important;
            border: 1px solid #D2E3FC !important;
        }

        /* 3. タブのボタンスタイル化 */
        [data-testid="stTabs"] {
            background-color: transparent !important;
        }
        [data-testid="stTabs"] button {
            font-size: 1rem;
            font-weight: 600;
            color: #5F6368 !important;
            /* Inactive tab style = gray outline button */
            border: 1px solid #DADCE0 !important;
            border-radius: 6px !important;
            background-color: #FFFFFF !important;
            padding: 8px 24px !important;
            margin-right: 12px !important;
            margin-bottom: 16px !important;
            box-shadow: 0 1px 2px rgba(60,64,67,0.1);
            transition: all 0.2s ease-in-out !important;
            height: auto !important;
            min-height: 40px !important;
        }
        [data-testid="stTabs"] button:hover {
            background-color: #F8F9FA !important;
            border-color: #1A73E8 !important;
            color: #1A73E8 !important;
        }
        [data-testid="stTabs"] button[aria-selected="true"] {
            /* Active tab style = '投稿の作成' reference blue button */
            background-color: #1A73E8 !important;
            color: #FFFFFF !important;
            border: 1px solid #1A73E8 !important;
            box-shadow: 0 1px 3px rgba(60,64,67,0.3) !important;
        }
        [data-testid="stTabs"] button[aria-selected="true"] p {
            color: #FFFFFF !important;
        }
        
        /* 邪魔なデフォルトのタブのアンダーラインを消す */
        div[data-baseweb="tab-highlight"] { display: none !important; }
        
        
        /* 4. DataFrames (Tables) - Card Style & Strict Layout */
        [data-testid="stDataFrame"] {
            background-color: #FFFFFF;
            border-radius: 8px;
            box-shadow: 0 1px 3px 0 rgba(60,64,67,0.15), 0 1px 2px 0 rgba(60,64,67,0.3);
            border: 1px solid #E8EAED;
            margin-bottom: 0 !important; /* Reset margin for strict grid */
        }
        
        /* Hide element index in dataframe to tighten layout */
        [data-testid="stDataFrame"] > div > div {
             border-radius: 8px; /* Force internal rounding */
        }
        
        /* Table Headers */
        thead tr th {
            color: #5F6368 !important;
            font-size: 0.85rem !important;
            font-weight: 600 !important;
            border-bottom: 2px solid #E8EAED !important;
            background-color: #F8F9FA !important; /* わずかに色をつけてヘッダー帯を明確に */
        }

        /* File Uploader */
        [data-testid="stFileUploaderDropzone"] {
            background-color: #FFFFFF;
            border: 1px dashed #1A73E8;
            border-radius: 8px;
        }
        [data-testid="stFileUploaderDropzone"] button {
            border-radius: 6px !important;
            background-color: #FFFFFF !important;
            color: #1A73E8 !important;
            box-shadow: none !important;
            border: 1px solid #DADCE0 !important;
        }
    </style>
    ''', unsafe_allow_html=True)
    # --- Sidebar ---
    st.sidebar.header("メニュー")
    
    # State for Mode
    if 'current_mode' not in st.session_state:
        st.session_state['current_mode'] = "総合"

    # Admin Login state
    if 'is_admin' not in st.session_state:
        st.session_state['is_admin'] = False
        
    with st.sidebar.expander("管理者メニュー"):
        if not st.session_state['is_admin']:
            password = st.text_input("パスワード", type="password")
            if st.button("ログイン"):
                if password == ADMIN_PASSWORD:
                    st.session_state['is_admin'] = True
                    st.rerun()
                else:
                    st.error("パスワードが違います")
        else:
            if st.button("ログアウト"):
                st.session_state['is_admin'] = False
                if st.session_state['current_mode'] == "設定 (Admin)":
                    st.session_state['current_mode'] = "総合"
                st.rerun()
                
    st.sidebar.markdown("<p style='color: #5F6368; font-weight: bold; font-size: 0.9rem; margin-bottom: 0px;'>モード選択</p>", unsafe_allow_html=True)
    
    # Custom Button Navigation
    if st.sidebar.button("総合", use_container_width=True, type="primary" if st.session_state['current_mode'] == "総合" else "secondary"):
        st.session_state['current_mode'] = "総合"
        st.rerun()
        
    if st.sidebar.button("件数", use_container_width=True, type="primary" if st.session_state['current_mode'] == "件数" else "secondary"):
        st.session_state['current_mode'] = "件数"
        st.rerun()
        
    if st.sidebar.button("順位", use_container_width=True, type="primary" if st.session_state['current_mode'] == "順位" else "secondary"):
        st.session_state['current_mode'] = "順位"
        st.rerun()

    if st.sidebar.button("成績詳細", use_container_width=True, type="primary" if st.session_state['current_mode'] == "成績詳細" else "secondary"):
        st.session_state['current_mode'] = "成績詳細"
        st.rerun()

    if st.sidebar.button("分析", use_container_width=True, type="primary" if st.session_state['current_mode'] == "分析" else "secondary"):
        st.session_state['current_mode'] = "分析"
        st.rerun()

    if st.sidebar.button("配点シミュレーション", use_container_width=True, type="primary" if st.session_state['current_mode'] == "配点シミュレーション" else "secondary"):
        st.session_state['current_mode'] = "配点シミュレーション"
        st.rerun()

    if st.sidebar.button("年間表彰", use_container_width=True, type="primary" if st.session_state['current_mode'] == "年間表彰" else "secondary"):
        st.session_state['current_mode'] = "年間表彰"
        st.rerun()
        
    if st.session_state['is_admin']:
        if st.sidebar.button("設定 (Admin)", use_container_width=True, type="primary" if st.session_state['current_mode'] == "設定 (Admin)" else "secondary"):
            st.session_state['current_mode'] = "設定 (Admin)"
            st.rerun()
            
    mode = st.session_state['current_mode']

    # --- Global Data Loading (from saved file) ---
    import os
    df = None
    data_path = "data/current_performance.csv"
    if os.path.exists(data_path):
        # We save & load as csv for internal persistence
        try:
            df = pd.read_csv(data_path)
            # Ensure proper typing especially for text columns to avoid float conversion
            if 'スタッフ名' in df.columns:
                df['スタッフ名'] = df['スタッフ名'].astype(str)
        except Exception as e:
             st.sidebar.error(f"データ読込エラー: {e}")
    else:
        if mode != "設定 (Admin)":
            st.sidebar.warning("成績データが未登録です。管理者メニューからアップロードしてください。")

    # --- Routing ---
    if mode == "設定 (Admin)":
        admin_page()
    elif mode == "順位":
        ranking_page(df)
    elif mode == "件数":
        count_page(df)
    elif mode == "成績詳細":
        individual_detail_page(df)
    elif mode == "分析":
        analysis_page(df)
    elif mode == "配点シミュレーション":
        simulation_page(df)
    elif mode == "年間表彰":
        annual_awards_page(df)
    else:
        comprehensive_page(df)

if __name__ == "__main__":
    main()

