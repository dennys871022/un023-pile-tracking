import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import plotly.express as px
import re
import datetime

st.set_page_config(page_title="UN023 排樁進度-雲端自動同步", layout="wide")
st.title("📊 UN023 排樁進度管理系統 (雲端自動存檔)")

# --- 1. 建立雲端串接 ---
conn = st.connection("gsheets", type=GSheetsConnection)

# 讀取底圖座標 (請確保 GitHub 專案內仍有 排樁座標.csv)
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
    # 限制 P1-P613
    df_piles = df[df['樁號大寫'].str.match(r'^P\d+$')].copy()
    df_piles['數字'] = df_piles['樁號大寫'].str.extract(r'(\d+)').astype(int)
    df_piles = df_piles[df_piles['數字'] <= 613]
    
    df_piles['X'] = pd.to_numeric(df_piles[x_col], errors='coerce')
    df_piles['Y'] = pd.to_numeric(df_piles[y_col], errors='coerce')
    return df_piles.dropna(subset=['X', 'Y'])

df_base = load_base_data()

# 讀取試算表歷史紀錄
def get_cloud_data():
    try:
        # 清除快取以確保抓到最新資料
        return conn.read(ttl="1s")
    except:
        return pd.DataFrame(columns=['樁號', '施工日期', '施作順序'])

df_history = get_cloud_data()

# --- 2. 施工進度登錄邏輯 ---
st.markdown("### 📝 今日進度錄入")
c_date, c_mode = st.columns([1, 3])
today = str(c_date.date_input("施工日期", datetime.date.today()))
mode = c_mode.radio("模式：", ["連續施工", "4支一循環 (1, 5, 9...)", "3支一循環 (1, 4, 7...)"], horizontal=True)

step = 1
if "4支" in mode: step = 4
elif "3支" in mode: step = 3

def sync_to_cloud(pile_ids):
    global df_history
    max_seq = 0 if df_history.empty else df_history['施作順序'].max()
    new_data = []
    
    for pid in pile_ids:
        # 防呆：檢查是否在 613 內且試算表尚未存在
        num_part = re.search(r'\d+', pid)
        if num_part and int(num_part.group()) <= 613:
            if pid not in df_history['樁號'].values:
                max_seq += 1
                new_data.append({'樁號': pid, '施工日期': today, '施作順序': max_seq})
    
    if new_data:
        updated = pd.concat([df_history, pd.DataFrame(new_data)], ignore_index=True)
        conn.update(data=updated)
        st.success(f"✅ 已成功自動同步至雲端！(新增 {len(new_data)} 支)")
        st.cache_data.clear()
        st.rerun()

t1, t2 = st.tabs(["🎯 起點自動推算", "✏️ 手動輸入區間"])

with t1:
    with st.form("auto"):
        col1, col2, col3 = st.columns(3)
        s_num = col1.number_input("起始樁號數字", 1, 613, 1)
        direct = col2.radio("方向", ["遞增 (+)", "遞減 (-)"])
        total = col3.number_input("施作支數", 1, 50, 10)
        if st.form_submit_button("確認並同步雲端"):
            target_list = []
            cur = s_num
            for _ in range(total):
                if cur < 1 or cur > 613: break
                target_list.append(f"P{cur}")
                cur = cur + step if direct == "遞增 (+)" else cur - step
            sync_to_cloud(target_list)

with t2:
    with st.form("manual"):
        raw = st.text_input("輸入區間 (例如: 1-50, 60)")
        if st.form_submit_button("確認並同步雲端"):
            target_list = []
            if raw:
                items = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw).strip())
                for item in items:
                    if '-' in item:
                        s, e = map(int, item.split('-'))
                        rng = range(s, e+1, step) if s<=e else range(s, e-1, -step)
                        for i in rng: target_list.append(f"P{i}")
                    elif item.isdigit(): target_list.append(f"P{item}")
            sync_to_cloud(target_list)

# --- 3. 視覺化圖面 ---
df_plot = df_base.merge(df_history, left_on='樁號大寫', right_on='樁號', how='left', suffixes=('', '_h'))
df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')

def make_label(row):
    if pd.notna(row['施作順序']):
        return f"{row['樁號']} ({int(row['施作順序'])})"
    return row['樁號']

df_plot['顯示標籤'] = df_plot.apply(make_label, axis=1)

# 側邊欄設定
st.sidebar.markdown(f"### 📊 雲端即時統計\n**累計完成總數：** {len(df_history)}")
plot_h = st.sidebar.slider("畫布高度", 500, 2500, 1000)

fig = px.scatter(
    df_plot, x='X', y='Y', text='顯示標籤', color='狀態',
    color_discrete_map={'未完成': 'lightgrey'}, # 未完成固定為灰色，其餘日期自動上色
    hover_data={'X': False, 'Y': False, '顯示標籤': False}
)

fig.update_traces(textposition='top center', marker=dict(size=10, line=dict(width=1, color='DarkSlateGrey')))
fig.update_layout(
    xaxis=dict(visible=False, showgrid=False),
    yaxis=dict(scaleanchor="x", scaleratio=1, visible=False, showgrid=False),
    plot_bgcolor='white', margin=dict(l=0, r=0, t=0, b=0), height=plot_h,
    legend=dict(title="施工日期", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
)

st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

# 產生今日報告
if st.button("📄 匯出今日施工報告"):
    today_report = df_history[df_history['施工日期'] == today]
    if not today_report.empty:
        csv = today_report.to_csv(index=False).encode('utf-8-sig')
        st.download_button("點此下載報告 (CSV)", csv, f"施工日報_{today}.csv", "text/csv")
    else:
        st.warning("今日尚無施工紀錄。")
