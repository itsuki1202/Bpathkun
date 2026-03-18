"""
CX分配くん 計算ロジックモジュール
- 月別CX設定データの読み書き
- STEP1〜3のインセンティブ計算
- user_masterからの初期メンバー生成
"""

import os
import json
import math
import pandas as pd

# CX専用データディレクトリ（プロジェクトルートからの相対パス）
_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CX_DATA_DIR = os.path.join(_BASE_DIR, "data", "cx_data")

# スキル手当区分の一覧
SKILL_TYPES = ["フロントスペシャリスト", "グランマイスター", "マイスター", "プレマイスター", "なし"]

# 役職フラグの一覧
ROLE_TYPES = ["role_manager", "role_leader", "role_member"]
ROLE_LABELS = {
    "role_manager": "店長",
    "role_leader": "リーダー",
    "role_member": "スタッフ",
}
ROLE_LABELS_REV = {v: k for k, v in ROLE_LABELS.items()}

# 店長の成果ポイントデフォルトテーブル（順位→ポイント）
MANAGER_DEFAULT_PTS = {1: 130, 2: 90, 3: 50}

# リーダーの成果ポイントデフォルトテーブル（順位→ポイント）
LEADER_DEFAULT_PTS = {1: 120, 2: 90, 3: 70, 4: 50, 5: 30, 6: 10, 7: 0}

# 店長デフォルト比率（1位を 1.0 とした比率）
_TOP_MGR = MANAGER_DEFAULT_PTS[1]
MANAGER_DEFAULT_RATIOS = {
    k: round(v / _TOP_MGR, 6) for k, v in MANAGER_DEFAULT_PTS.items()
}

# リーダーデフォルト比率（1位を 1.0 とした比率、7位 = 0固定）
_TOP_LDR = LEADER_DEFAULT_PTS[1]
LEADER_DEFAULT_RATIOS = {
    k: round(v / _TOP_LDR, 6) if _TOP_LDR > 0 else 0
    for k, v in LEADER_DEFAULT_PTS.items()
}


def _get_cx_config_path(month: str) -> str:
    """月別CX設定ファイルのパスを返す"""
    return os.path.join(CX_DATA_DIR, month, "cx_config.json")


def load_cx_config(month: str) -> dict:
    """月別CX設定JSONを読み込んで返す。存在しない場合はデフォルト値を返す。"""
    path = _get_cx_config_path(month)
    if not os.path.exists(path):
        return _default_cx_config(month)
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"CX設定の読み込みエラー ({path}): {e}")
        return _default_cx_config(month)


def save_cx_config(config: dict, month: str) -> bool:
    """月別CX設定JSONに保存する。"""
    path = _get_cx_config_path(month)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"CX設定の保存エラー ({path}): {e}")
        return False


def _default_cx_config(month: str) -> dict:
    """デフォルトのCX設定を返す"""
    return {
        "month": month,
        "cx_total_points": 0,
        "base_pt": 50,
        "calculation_mode": "A",
        "skill_allowances": {
            "フロントスペシャリスト": 0,
            "グランマイスター": 0,
            "マイスター": 0,
            "プレマイスター": 0,
        },
        # 店長成果ポイントテーブル（順位: pt）
        "manager_pts": {str(k): v for k, v in MANAGER_DEFAULT_PTS.items()},
        # リーダー成果ポイントテーブル（順位: pt, 7位以下は"7"キーに0）
        "leader_pts": {str(k): v for k, v in LEADER_DEFAULT_PTS.items()},
        # スタッフ線形補間の上限（1位のpt）・下限（最下位のpt）
        "member_pt_top": 100,
        "member_pt_bottom": 0,
        "members": [],
    }


def get_default_cx_members(user_master: dict) -> list:
    """
    user_master.json の情報からCX分配くん用の初期メンバーリストを生成する。
    - shop_managers に登録されている名前 → role_manager（店長）
    - 各チームの先頭スタッフ → role_leader（リーダー）
    - それ以外 → role_member（スタッフ）
    """
    users = user_master.get("users", [])
    shop_managers_dict = user_master.get("shop_managers", {})

    # 店長名のセット（正規化して比較）
    manager_names = set()
    for name in shop_managers_dict.values():
        if name:
            manager_names.add(str(name).strip())

    # チームごとの先頭スタッフをリーダーとして判定
    seen_teams = {}
    leader_names = set()
    for u in users:
        team = str(u.get("team", "")).strip()
        name = str(u.get("name", "")).strip()
        if team and team not in seen_teams:
            seen_teams[team] = name
            leader_names.add(name)

    members = []

    # user_masterの全スタッフを追加
    for u in users:
        name = str(u.get("name", "")).strip()
        if not name:
            continue

        if name in manager_names:
            role = "role_manager"
        elif name in leader_names:
            role = "role_leader"
        else:
            role = "role_member"

        members.append({
            "name": name,
            "role": role,
            "skill_type": "なし",
            "manual_pt": None,
        })

    # shop_managersに登録されているが、usersリストにいない店長も追加
    for shop, mgr_name in shop_managers_dict.items():
        if not mgr_name:
            continue
        mgr_name = str(mgr_name).strip()
        existing_names = {m["name"] for m in members}
        if mgr_name not in existing_names:
            members.append({
                "name": mgr_name,
                "role": "role_manager",
                "skill_type": "なし",
                "manual_pt": None,
            })

    return members


def _calc_member_score_pt(rank: int, total_members: int, pt_top: int = 100, pt_bottom: int = 0) -> int:
    """
    スタッフ（role_member）の成果ポイントを線形補間で算出する。
    1位=pt_top、最下位=pt_bottom、人数に応じて等間隔で減少。
    計算式: round(pt_top - ((rank-1) / (total_members-1)) * (pt_top - pt_bottom))
    total_membersが1の場合はpt_top固定。
    """
    if total_members <= 1:
        return pt_top
    return round(pt_top - ((rank - 1) / (total_members - 1)) * (pt_top - pt_bottom))


def calculate_cx_distribution(
    cx_config: dict,
    shop_rank_df: pd.DataFrame,
    team_rank_df: pd.DataFrame,
    individual_scored_df: pd.DataFrame,
    user_master: dict,
) -> dict:
    """
    CX分配のSTEP1〜3を計算し、各スタッフの結果を辞書として返す。

    Parameters:
        cx_config: load_cx_config()で取得したCX設定
        shop_rank_df: aggregate_shop_scores()の結果（店舗ランキング）
        team_rank_df: aggregate_team_scores()の結果（チームランキング）
        individual_scored_df: calculate_scores()の結果（個人スコア・Rank列付き）
        user_master: load_user_master()の結果

    Returns:
        {
          "step1": {
              "old_allowances": {name: int},     # 個人旧手当
              "fixed_cost_total": int,            # 固定費総額
              "surplus_fund": float,              # 変動分配用余剰原資
          },
          "step2": {
              "points": {name: {"base_pt":int, "performance_pt":int, "total_pt":int}},
          },
          "step3": {
              "system_total_pt": int,
              "pt_unit_price": float,
              "results": {name: {"歩合額":int, "旧手当":int, "最終支給額":int}},
          },
          "mode": "A" or "B",
          "base_pt": int,
          "pt_unit_price": float,
        }
    """
    mode = cx_config.get("calculation_mode", "A")
    cx_total = cx_config.get("cx_total_points", 0)
    base_pt = cx_config.get("base_pt", 50)
    skill_allowances = cx_config.get("skill_allowances", {})
    members = cx_config.get("members", [])

    # ランキングデータの準備（Rank列）
    # 店舗ランキング: 店舗名 → Rank
    shop_rank_map = {}
    shop_manager_rank_map = {}  # 店長名 → Rank
    if shop_rank_df is not None and not shop_rank_df.empty:
        # 店舗名列の特定
        shop_col = next((c for c in ["店舗名", "店舗", "Shop"] if c in shop_rank_df.columns), None)
        if shop_col:
            for _, row in shop_rank_df.iterrows():
                shop_rank_map[str(row[shop_col]).strip()] = int(row.get("Rank", 99))
            # 店長名 → Rank のマッピングも作成
            mgr_col = "Manager"
            if mgr_col in shop_rank_df.columns:
                for _, row in shop_rank_df.iterrows():
                    mgr = str(row.get(mgr_col, "")).strip()
                    if mgr:
                        shop_manager_rank_map[mgr] = int(row.get("Rank", 99))

    # チームランキング: チーム名 → Rank
    team_rank_map = {}
    leader_rank_map = {}  # リーダー名 → Rank
    if team_rank_df is not None and not team_rank_df.empty:
        team_col = next((c for c in ["チーム名", "チーム", "Team"] if c in team_rank_df.columns), None)
        if team_col:
            for _, row in team_rank_df.iterrows():
                team_rank_map[str(row[team_col]).strip()] = int(row.get("Rank", 99))
            # リーダー名 → Rank のマッピングも作成
            ldr_col = "Leader"
            if ldr_col in team_rank_df.columns:
                for _, row in team_rank_df.iterrows():
                    ldr = str(row.get(ldr_col, "")).strip()
                    if ldr:
                        leader_rank_map[ldr] = int(row.get("Rank", 99))

    # 個人ランキング（role_member 用）: スタッフ名 → Rank（役職者を除外してから再付け）
    member_rank_map = {}
    if individual_scored_df is not None and not individual_scored_df.empty:
        name_col = next(
            (c for c in ["スタッフ名", "氏名", "Name"] if c in individual_scored_df.columns), None
        )
        if name_col and "Total_Score" in individual_scored_df.columns:
            # 役職者（店長・リーダー）の名前を取得
            manager_names_set = set()
            leader_names_set = set()
            for m in members:
                if m["role"] == "role_manager":
                    manager_names_set.add(m["name"])
                elif m["role"] == "role_leader":
                    leader_names_set.add(m["name"])

            # 役職者を除外したデータフレーム
            member_df = individual_scored_df[
                ~individual_scored_df[name_col].astype(str).str.strip().isin(
                    manager_names_set | leader_names_set
                )
            ].copy()

            # 役職者除外後に再Rankを付ける（Total_Scoreの降順）
            if not member_df.empty:
                member_df["_cx_rank"] = member_df["Total_Score"].rank(
                    ascending=False, method="min"
                ).astype(int)
                for _, row in member_df.iterrows():
                    nm = str(row[name_col]).strip()
                    member_rank_map[nm] = int(row["_cx_rank"])

    total_members_for_linear = len(member_rank_map)

    # 設定からポイントテーブルを取得（設定なければデフォルト）
    raw_mgr = cx_config.get("manager_pts", {str(k): v for k, v in MANAGER_DEFAULT_PTS.items()})
    manager_pts = {int(k): int(v) for k, v in raw_mgr.items()}
    raw_ldr = cx_config.get("leader_pts", {str(k): v for k, v in LEADER_DEFAULT_PTS.items()})
    leader_pts = {int(k): int(v) for k, v in raw_ldr.items()}
    member_pt_top = int(cx_config.get("member_pt_top", 100))
    member_pt_bottom = int(cx_config.get("member_pt_bottom", 0))

    # ---- STEP1: 旧手当算出と余剰原資計算 ----
    old_allowances = {}  # name → 旧手当（切り捨て）
    for m in members:
        name = m["name"]
        skill_type = m.get("skill_type", "なし")
        base_allowance = skill_allowances.get(skill_type, 0) if skill_type != "なし" else 0
        # 旧手当 = 支援費 × 0.85（端数切り捨て）
        old_allowances[name] = math.floor(base_allowance * 0.85)

    if mode == "A":
        fixed_cost_total = sum(old_allowances.values())
        surplus_fund = cx_total - fixed_cost_total
    else:  # モードB
        fixed_cost_total = 0
        surplus_fund = cx_total

    # ---- STEP2: ポイント付与 ----
    pt_details = {}  # name → {base_pt, performance_pt, total_pt}

    for m in members:
        name = m["name"]
        role = m.get("role", "role_member")
        manual_pt = m.get("manual_pt", None)

        if manual_pt is not None:
            # 手動ポイントが設定されている場合はそのままTotal_Ptとして使用
            try:
                total_pt = int(manual_pt)
            except (ValueError, TypeError):
                total_pt = base_pt
            pt_details[name] = {
                "base_pt": "-",
                "performance_pt": "-",
                "total_pt": total_pt,
                "is_manual": True,
            }
            continue

        # 成果ポイントの計算
        performance_pt = 0
        if role == "role_manager":
            # 店長：店長名→Rankのマップから取得
            rank = shop_manager_rank_map.get(name, 99)
            performance_pt = manager_pts.get(rank, 0)
        elif role == "role_leader":
            # リーダー：リーダー名→Rankのマップから取得
            rank = leader_rank_map.get(name, 99)
            # 7位以下はキー7の値を使用
            performance_pt = leader_pts.get(rank, leader_pts.get(7, 0))
        else:
            # スタッフ（役職者除外後Rank）：線形補間
            rank = member_rank_map.get(name, None)
            if rank is not None and total_members_for_linear > 0:
                performance_pt = _calc_member_score_pt(rank, total_members_for_linear, member_pt_top, member_pt_bottom)
            else:
                performance_pt = 0

        total_pt = base_pt + performance_pt
        pt_details[name] = {
            "base_pt": base_pt,
            "performance_pt": performance_pt,
            "total_pt": total_pt,
            "is_manual": False,
        }

    # ---- STEP3: 1Pt単価と最終支給額 ----
    system_total_pt = sum(d["total_pt"] for d in pt_details.values())

    if system_total_pt > 0 and surplus_fund > 0:
        pt_unit_price = surplus_fund / system_total_pt
    else:
        pt_unit_price = 0.0

    results = {}
    for name, pts in pt_details.items():
        total_pt = pts["total_pt"]
        歩合額 = round(total_pt * pt_unit_price)
        旧手当 = old_allowances.get(name, 0)

        if mode == "A":
            最終支給額 = 旧手当 + 歩合額
        else:  # モードB
            最終支給額 = 歩合額

        results[name] = {
            "base_pt": pts["base_pt"],
            "performance_pt": pts["performance_pt"],
            "total_pt": total_pt,
            "is_manual": pts.get("is_manual", False),
            "歩合額": 歩合額,
            "旧手当": 旧手当,
            "最終支給額": 最終支給額,
        }

    return {
        "step1": {
            "old_allowances": old_allowances,
            "fixed_cost_total": fixed_cost_total,
            "surplus_fund": surplus_fund,
        },
        "step3": {
            "system_total_pt": system_total_pt,
            "pt_unit_price": pt_unit_price,
            "results": results,
        },
        "mode": mode,
        "base_pt": base_pt,
        "pt_unit_price": pt_unit_price,
    }
