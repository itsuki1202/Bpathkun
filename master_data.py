# Performance Items based on User Config (Japanese Keys)
PERFORMANCE_ITEMS = {
    "スマホ総販": {"label": "スマホ総販", "target": 0, "weight": 80},
    "GOLD": {"label": "GOLD", "target": 0, "weight": 20},
    "GOLD　U": {"label": "GOLD　U", "target": 0, "weight": 10},
    "PLATINUM": {"label": "PLATINUM", "target": 0, "weight": 30},
    "カード合計": {"label": "カード合計", "target": 0, "weight": 40},
    "MAX": {"label": "MAX", "target": 0, "weight": 30},
    "ポイ活": {"label": "ポイ活", "target": 0, "weight": 50},
    "home5G": {"label": "home5G", "target": 0, "weight": 40},
    "光": {"label": "光", "target": 0, "weight": 50},
    "イエナカ合計": {"label": "イエナカ合計", "target": 0, "weight": 10},
    "でんき": {"label": "でんき", "target": 0, "weight": 30},
    "あんしんS": {"label": "あんしんS", "target": 0, "weight": 50},
    "アフィリエイト": {"label": "アフィリエイト", "target": 0, "weight": 50},
    "あんしんS+アフィ": {"label": "あんしんS+アフィ", "target": 0, "weight": 40},
    "ディズニー": {"label": "ディズニー", "target": 0, "weight": 50},
    "Lemino・GoogleOne\nNetflix・YouTube Premium": {"label": "Lemino・GoogleOne\nNetflix・YouTube Premium", "target": 0, "weight": 20},
    "コーティング": {"label": "コーティング", "target": 0, "weight": 10},
    "Amazon": {"label": "Amazon", "target": 0, "weight": 20},
}

# Mapping from Legacy System Column Name (English) to New Display Name (Japanese)
# This allows old sample data to be partially mapped if uploaded
COLUMN_MAPPING = {
    "Name": "スタッフ名",
    "Shop": "店舗名",
    "Team": "チーム名",
    "Agency": "代理店名",
    "氏名": "スタッフ名",
    "店舗": "店舗名",
    "チーム": "チーム名",
    "代理店": "代理店名",
    # Legacy mappings (best effort)
    "Smartphone_Total": "スマホ総販",
    "dCard_Individual": "GOLD", # Approximate
    "AnshinPack": "あんしんS",
    "Ienaka": "イエナカ合計",
    "Eximo": "MAX", # Guess
    "Coating": "コーティング",
    "Affiliate": "アフィリエイト"
}

# Reverse mapping: Not strictly needed if we use Japanese internally
REVERSE_COLUMN_MAPPING = {v: k for k, v in COLUMN_MAPPING.items()}

REQUIRED_COLUMNS_JP = [
    "スタッフ名", "店舗名", "チーム名", "代理店名"
]

REQUIRED_COLUMNS = REQUIRED_COLUMNS_JP + list(PERFORMANCE_ITEMS.keys())

# 新フォーマットの列定義（成績データ登録用テンプレート）
NEW_FORMAT_COLUMNS = ["スタッフ名", "店舗名", "チーム名", "代理店名", "基準時間"] + list(PERFORMANCE_ITEMS.keys())
