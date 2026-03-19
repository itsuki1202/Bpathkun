import streamlit as st
import pandas as pd
import numpy as np

# 一時的なUIモックアップ用スクリプト（本番には影響しません）
st.set_page_config(layout="wide", page_title="Bpathkun UI Proposal")

st.markdown("""
<style>
    /* Google Fonts (Noto Sans JP) のインポート */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700&display=swap');

    /* アプリ全体のフォントをGoogleフォントに変更 */
    html, body, [class*="css"], .stApp, p, div, span, label, h1, h2, h3, h4, h5, h6 {
        font-family: 'Noto Sans JP', sans-serif !important;
    }

    /* 全体の背景色をわずかにグレーにしてカードを引き立たせる */
    .stApp {
        background-color: #F8FAFC;
    }
    
    /* サイドバーの背景色を白にして境界を作る */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }
    
    /* サイドバーのメニューボタンをBOXらしく（丸みを帯びたブロック） */
    [data-testid="stSidebar"] div.stButton > button {
        width: 100%;
        justify-content: flex-start;
        border: 1px solid transparent !important;
        border-radius: 8px !important;
        background-color: #F8FAFC !important; /* 薄いグレーのBOX */
        color: #475569 !important;
        font-size: 14px !important;
        padding: 0.75rem 1rem;
        margin-bottom: 4px;
        transition: all 0.2s ease-in-out;
    }
    [data-testid="stSidebar"] div.stButton > button:hover {
        background-color: #EFF6FF !important; /* ホバー時に青っぽいBOXに */
        color: #1D4ED8 !important;
        border-color: #BFDBFE !important;
        font-weight: 600;
        transform: translateY(-1px);
    }

    /* メインコンテンツの表などを「カード」のように白く浮かせる */
    [data-testid="stVerticalBlock"] > div > div.stDataFrame {
        background-color: #FFFFFF;
        padding: 1rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        border: 1px solid #E2E8F0;
    }
    
    /* メトリクス（数値）もカード化 */
    [data-testid="metric-container"] {
        background-color: #FFFFFF;
        padding: 1.5rem 1rem;
        border-radius: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
        border: 1px solid #E2E8F0;
    }
    
    /* タイトルの装飾（Googleライクなスッキリとした洗練フォント） */
    h1 {
        font-size: 2rem !important;
        color: #0F172A;
        font-weight: 700;
        letter-spacing: -0.025em;
    }
    
    /* サブヘッダー（総合＞個人成績など）を立体的なBOXに */
    h3 {
        font-size: 1.1rem !important;
        color: #1E293B;
        background-color: #FFFFFF; /* ボックスの背景 */
        padding: 10px 16px !important;
        border-radius: 8px; /* 角丸 */
        border-left: 6px solid #3B82F6; /* Googleらしいブルーのアクセント線 */
        display: inline-block; /* 幅を文字に合わせる */
        margin-top: 1.5rem !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08); /* 立体感のための影 */
        border: 1px solid #E2E8F0; /* 境界線 */
        border-left-width: 6px; /* 左側だけ太く */
    }
    
    /* ログアウトボタンの特別装飾（控えめな文字） */
    [data-testid="stSidebar"] div.stButton:nth-child(1) > button {
        background-color: transparent !important;
        color: #64748B !important;
    }
    [data-testid="stSidebar"] div.stButton:nth-child(1) > button:hover {
        color: #EF4444 !important;
        background-color: #FEF2F2 !important;
        border-color: #FCA5A5 !important;
    }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.button("ログアウト", key="logout")
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 11px; color: #94A3B8; font-weight: 700; letter-spacing: 0.05em; margin-bottom: 12px;'>MAIN MENU</div>", unsafe_allow_html=True)
    
    # 完全に絵文字を排除したクリーンなメニュー（CSSでBOX化）
    menus = ["総合", "件数", "順位", "成績詳細", "分析", "ルール説明", "CX分配くん", "最新データをDL"]
    for i, menu in enumerate(menus):
        st.button(menu, key=f"menu_{i}")
        
    st.markdown("<div style='margin-top: 3rem; font-size: 12px; color: #94A3B8;'>Logged in as:<br><span style='color: #475569; font-weight: 600;'>Administrator</span></div>", unsafe_allow_html=True)

st.title("Brightpathkun")
st.caption("Last updated: 2026/03/18 17:50")

# メトリクス（数値サマリー）
cols = st.columns(4)
with cols[0]:
    st.metric(label="全店 スマホ販売", value="342件", delta="24件 先月比")
with cols[1]:
    st.metric(label="平均達成率", value="98.5%", delta="-1.5%", delta_color="normal")
with cols[2]:
    st.metric(label="トップ店舗", value="倉敷店", delta="維持")
with cols[3]:
    st.metric(label="CX分配額", value="¥ 124,500", delta="¥ 12,000 先月比")

# BOX化した見出し（h3要素として出力）
st.markdown("### 総合 > 個人成績")

# ダミーデータ (本番データに近い構造)
df = pd.DataFrame({
    "店舗": ["倉敷店", "倉敷店", "倉敷店", "岡山店", "岡山店"],
    "チーム": ["MGN4", "エース", "小橋女学院", "Alpha", "Beta"],
    "スタッフ名": ["四倉祐貴", "高橋美和", "小橋実央", "山田太郎", "鈴木花子"],
    "店頭PI": [40.0, 40.0, 40.0, 38.5, 45.0],
    "スマホ総販": [51.9, 67.7, 47.4, 55.0, 60.1],
    "モトローラ": [0.0, 10.0, 0.0, 5.0, 10.0],
    "合計点": [91.9, 117.7, 87.4, 98.5, 115.1]
})

st.dataframe(df, use_container_width=True, hide_index=True)

st.markdown("### 店舗別 トレンド")
chart_data = pd.DataFrame(
    np.cumsum(np.random.randn(20, 3) * 2 + 5, axis=0),
    columns=['倉敷店', '岡山店', '玉野店']
)
# プロットの色を指定して、少し淡い青・緑・オレンジにする
st.line_chart(chart_data, color=["#3B82F6", "#10B981", "#F59E0B"])
