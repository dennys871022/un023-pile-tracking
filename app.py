import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import plotly.express as px
import re
import datetime

st.set_page_config(page_title="UN023 排樁進度-雲端同步版", layout="wide")
st.title("🚧 UN023 排樁進度管理 (全自動雲端存檔)")

# --- 1. 連接 Google 試算表 ---
# 在 Streamlit Cloud Secrets 中設定好試算表網址
# 格式: [connections.gsheets] -> spreadsheet = "https://docs.google.com/spreadsheets/d/1y3Qnlx9qFwV6S6pyFTsT4rlXP_Tb8qd9tNhRBTjBHao/edit?usp=sharing"
conn = st.connection("gsheets", type=GSheetsConnection)

# 讀取底圖 (座標)
@st.cache_data
def load_base_data():
    try:
        try:
            df = pd.read_csv('排樁座標.csv', encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv('排樁座標.csv', encoding='big5')
    except: return None
    
    x_col = next((col for col in df.columns if 'X' in col.upper()), None)
    y_col = next((col for col in df.columns if 'Y' in col.upper()), None)
    text_col = next((col for col in df.columns if '內容' in col or '值' in col or 'VALUE' in col.upper()), None)
    
    df['樁號'] = df[text_col].apply(lambda x: re.sub(r'\\[^;]+;|[{}]', '', str(x)).strip())
    df['樁號大寫'] = df['樁號'].str.upper()
    df_piles = df[df['樁號大寫'].str.match(r'^P\d+$')].copy()
    df_piles['X'] = pd.to_numeric(df_piles[x_col], errors='coerce')
    df_piles['Y'] = pd.to_numeric(df_piles[y_col], errors='coerce')
    return df_piles.dropna(subset=['X', 'Y'])

df_base = load_base_data()

# 讀取 Google 試算表中的歷史紀錄
def get_history():
    try:
        return conn.read(ttl="1s") # ttl=1s 確保每次重新整理都抓最新
    except:
        return pd.DataFrame(columns=['樁號', '施工日期', '施作順序'])

df_history = get_history()

# --- 2. 施工登錄邏輯 ---
st.markdown("### 📝 今日施工登錄")
col_date, col_mode = st.columns([1, 3])
today_date = col_date.date_input("施工日期", datetime.date.today())
mode = col_mode.radio("模式：", ["連續", "4支循環", "3支循環"], horizontal=True)

step = 1
if "4支" in mode: step = 4
elif "3支" in mode: step = 3

# 寫入試算表的函數
def save_to_gsheets(new_piles):
    global df_history
    current_max_seq = 0 if df_history.empty else df_history['施作順序'].max()
    
    new_rows = []
    for p_id in new_piles:
        # 樁號上限與重複檢查
        p_num = int(re.search(r'\d+', p_id).group())
        if p_num <= 613 and p_id not in df_history['樁號'].values:
            current_max_seq += 1
            new_rows.append({
                '樁號': p_id,
                '施工日期': str(today_date),
                '施作順序': current_max_seq
            })
    
    if new_rows:
        updated_df = pd.concat([df_history, pd.DataFrame(new_rows)], ignore_index=True)
        conn.update(data=updated_df)
        st.success(f"✅ 自動同步成功！新增 {len(new_rows)} 支樁至雲端試算表。")
        st.cache_data.clear() # 強制清除緩存以更新圖面

# 輸入分頁
t1, t2 = st.tabs(["🎯 起點推算", "✏️ 手動輸入"])
with t1:
    with st.form("auto"):
        c1, c2, c3 = st.columns(3)
        s_n = c1.number_input("起始數字", 1, 613, 1)
        dir = c2.radio("方向", ["遞增", "遞減"])
        num = c3.number_input("支數", 1, 50, 10)
        if st.form_submit_button("確認並同步雲端"):
            p_list = []
            curr = s_n
            for _ in range(num):
                if curr < 1 or curr > 613: break
                p_list.append(f"P{curr}")
                curr = curr + step if dir == "遞增" else curr - step
            save_to_gsheets(p_list)

with t2:
    with st.form("manual"):
        raw = st.text_input("輸入區間 (如 1-50, 60)")
        if st.form_submit_button("確認並同步雲端"):
            p_list = []
            if raw:
                items = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw).strip())
                for item in items:
                    if '-' in item:
                        s, e = map(int, item.split('-'))
                        rng = range(s, e+1, step) if s<=e else range(s, e-1, -step)
                        for i in rng: p_list.append(f"P{i}")
                    elif item.isdigit(): p_list.append(f"P{item}")
            save_to_gsheets(p_list)

# --- 3. 繪圖顯示 ---
df_plot = df_base.merge(df_history, left_on='樁號大寫', right_on='樁號', how='left', suffixes=('', '_h'))
df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
df_plot['標籤'] = df_plot.apply(lambda r: f"{r['樁號']} ({int(r['施作順序'])})" if pd.notna(r['施作順序']) else r['樁號'], axis=1)

st.sidebar.markdown(f"### 📊 雲端統計\n**已完成總數：** {len(df_history)}")
h = st.sidebar.slider("畫布高度", 500, 2500, 1000)

fig = px.scatter(
    df_plot, x='X', y='Y', text='標籤', color='狀態',
    color_discrete_map={'未完成': 'lightgrey'},
    hover_data={'X': False, 'Y': False, '標籤': False}
)
fig.update_traces(textposition='top center', marker=dict(size=10, line=dict(width=1, color='DarkSlateGrey')))
fig.update_layout(
    xaxis=dict(visible=False), yaxis=dict(scaleanchor="x", scaleratio=1, visible=False),
    plot_bgcolor='white', height=h, margin=dict(l=0,r=0,t=0,b=0)
)
st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

# 匯出今日報告
if st.button("📄 產生今日施工報告 (CSV匯出)"):
    report = df_history[df_history['施工日期'] == str(today_date)]
    st.download_button("點此下載報告", report.to_csv(index=False).encode('utf-8-sig'), f"施工報告_{today_date}.csv")
