"""
認証管理モジュール
- ユーザー情報（config.yaml）の読み書き
- ログイン履歴（login_log.csv）の記録・読み込み
- パスワードのハッシュ化ユーティリティ
"""

import os
import yaml
import bcrypt
import csv
import pandas as pd
from datetime import datetime

# ファイルパス定義
AUTH_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "auth")
CONFIG_PATH = os.path.join(AUTH_DIR, "config.yaml")
LOG_PATH = os.path.join(AUTH_DIR, "login_log.csv")


def load_auth_config() -> dict:
    """config.yaml を読み込んで返す"""
    if not os.path.exists(CONFIG_PATH):
        # 初期設定が存在しない場合は空の構造を返す
        return {
            "credentials": {"usernames": {}},
            "cookie": {
                "expiry_days": 1,
                "key": "brightpath_secret_key_2026",
                "name": "brightpath_auth"
            },
            "preauthorized": {"emails": []}
        }
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def save_auth_config(config: dict):
    """config.yaml に書き込む"""
    os.makedirs(AUTH_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False)


def hash_password(plain_password: str) -> str:
    """パスワードをbcryptでハッシュ化して返す"""
    return bcrypt.hashpw(plain_password.encode(), bcrypt.gensalt()).decode()


def verify_password(plain_password: str, hashed: str) -> bool:
    """パスワードを検証する"""
    return bcrypt.checkpw(plain_password.encode(), hashed.encode())


def get_all_users(config: dict) -> list:
    """ユーザー一覧を取得する [{username, name, email, role}, ...]"""
    users = []
    for uname, udata in config.get("credentials", {}).get("usernames", {}).items():
        users.append({
            "username": uname,
            "name": udata.get("name", ""),
            "email": udata.get("email", ""),
            "role": udata.get("role", "viewer"),
        })
    return users


def add_user(config: dict, username: str, display_name: str, email: str,
             plain_password: str, role: str = "viewer") -> tuple:
    """
    ユーザーを追加する
    Returns: (成功フラグ, メッセージ)
    """
    usernames = config.get("credentials", {}).get("usernames", {})
    if username in usernames:
        return False, f"ユーザーID「{username}」は既に存在します"
    if not username or not plain_password:
        return False, "ユーザーIDとパスワードは必須です"
    usernames[username] = {
        "name": display_name,
        "email": email,
        "password": hash_password(plain_password),
        "role": role,
    }
    config["credentials"]["usernames"] = usernames
    save_auth_config(config)
    return True, f"ユーザー「{username}」を追加しました"


def delete_user(config: dict, username: str) -> tuple:
    """
    ユーザーを削除する
    Returns: (成功フラグ, メッセージ)
    """
    usernames = config.get("credentials", {}).get("usernames", {})
    if username not in usernames:
        return False, f"ユーザーID「{username}」は存在しません"
    if username == "admin":
        return False, "管理者アカウントは削除できません"
    del usernames[username]
    config["credentials"]["usernames"] = usernames
    save_auth_config(config)
    return True, f"ユーザー「{username}」を削除しました"


def reset_password(config: dict, username: str, new_password: str) -> tuple:
    """
    パスワードをリセットする
    Returns: (成功フラグ, メッセージ)
    """
    usernames = config.get("credentials", {}).get("usernames", {})
    if username not in usernames:
        return False, f"ユーザーID「{username}」は存在しません"
    if not new_password:
        return False, "新しいパスワードを入力してください"
    usernames[username]["password"] = hash_password(new_password)
    config["credentials"]["usernames"] = usernames
    save_auth_config(config)
    return True, f"「{username}」のパスワードを変更しました"


def log_login(username: str, display_name: str, status: str, message: str = ""):
    """
    ログイン履歴を login_log.csv に追記する
    status: "success" or "failed"
    """
    os.makedirs(AUTH_DIR, exist_ok=True)
    # ファイルが存在しない場合はヘッダーも書く
    file_exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["timestamp", "username", "display_name", "status", "message"])
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            username,
            display_name,
            status,
            message,
        ])


def load_login_log() -> pd.DataFrame:
    """login_log.csv を DataFrame で返す"""
    if not os.path.exists(LOG_PATH):
        return pd.DataFrame(columns=["timestamp", "username", "display_name", "status", "message"])
    try:
        df = pd.read_csv(LOG_PATH)
        if df.empty:
            return pd.DataFrame(columns=["timestamp", "username", "display_name", "status", "message"])
        df["timestamp"] = pd.to_datetime(df["timestamp"])
        return df.sort_values("timestamp", ascending=False).reset_index(drop=True)
    except Exception:
        return pd.DataFrame(columns=["timestamp", "username", "display_name", "status", "message"])
