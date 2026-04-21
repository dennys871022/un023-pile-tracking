import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import plotly.express as px
import re
import datetime

st.set_page_config(page_title="UN023 排樁進度管理系統", layout="wide")
st.title("📊 UN023 排樁進度管理系統")

# 1. 建立雲端與底圖連接
conn = st.connection("gsheets", type=GSheetsConnection)

@st.cache_data
def load_base_data():
    try:
        try:
            df = pd.read_csv('排樁座標.csv', encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv('排樁座標.csv', encoding='big5')
    except:
        return None
    
    x_col = next((col for col in df.columns if 'X' in col.upper()), None)
    y_col = next((col for col in df.columns if 'Y' in col.upper()), None)
    text_col = next((col for col in df.columns if '內容' in col or '值' in col), None)
    
    # 清洗資料並限制 P1 伺 P613
    df['樁號'] = df[text_col].apply(lambda x: re.sub(r'\\[^;]+;|[{}]', '', str(x)).strip())
    df['樁號大寫'] = df['樁號'].str.upper()
    df_piles = df[df['樁號大寫'].str.match(r'^P\d+$')].copy()
    df_piles['數字'] = df_piles['樁號大寫'].str.extract(r'(\d+)').astype(int)
    df_piles = df_piles[df_piles['數字'] <= 613]
    
    df_piles['X'] = pd.to_numeric(df_piles[x_col], errors='coerce')
    df_piles['Y'] = pd.to_numeric(df_piles[y_col], errors='coerce')
    
    # 強制去重，防止 P1 重複出現導致元件報錯
    return df_piles.drop_duplicates(subset=['樁號大寫']).dropna(subset=['X', 'Y'])

df_base = load_base_data()

def get_cloud_data():
    try:
        data = conn.read(ttl=0)
        return data if not data.empty else pd.DataFrame(columns=['樁號', '施工日期', '施作順序'])
    except:
        return pd.DataFrame(columns=['樁號', '施工日期', '施作順序'])

df_history = get_cloud_data()

# 2. 施工進度登錄
st.markdown("### 📝 施工進度登錄")
c_date, c_mode = st.columns([1, 3])
today_str = str(c_date.date_input("施工日期", datetime.date.today()))
mode_opt = c_mode.radio("模式選擇：", ["連續施作", "4支一循環 (1, 5...)", "3支一循環 (1, 4...)"], horizontal=True)

step = 1
if "4支" in mode_opt: step = 4
elif "3支" in mode_opt: step = 3

def sync_to_cloud(input_list):
    current_history = get_cloud_data()
    max_seq = 0 if current_history.empty else current_history['施作順序'].max()
    new_rows = []
    
    for pid in input_list:
        pid = pid.upper().strip()
        if not current_history.empty and pid in current_history['樁號'].values:
            continue
        num_match = re.search(r'\d+', pid)
        if num_match and int(num_match.group()) <= 613:
            max_seq += 1
            new_rows.append({'樁號': pid, '施工日期': today_str, '施作順序': max_seq})
    
    if new_rows:
        updated = pd.concat([current_history, pd.DataFrame(new_rows)], ignore_index=True)
        conn.update(data=updated)
        st.success(f"同步完成：新增 {len(new_rows)} 支")
        st.cache_data.clear()
        st.rerun()

t1, t2 = st.tabs(["🎯 起點自動推算", "✏️ 區間手動輸入"])

with t1:
    with st.form("auto_form"):
        col1, col2, col3 = st.columns(3)
        s_num = col1.number_input("起始編號", 1, 613, 1)
        dir_opt = col2.radio("方向", ["遞增 (+)", "遞減 (-)"])
        total_piles = col3.number_input("施作數量", 1, 100, 10)
        if st.form_submit_button("🚀 確認並同步雲端"):
            targets = []
            curr = s_num
            for _ in range(total_piles):
                if curr < 1 or curr > 613: break
                targets.append(f"P{curr}")
                curr = curr + step if "遞增" in dir_opt else curr - step
            sync_to_cloud(targets)

with t2:
    with st.form("manual_form"):
        raw_val = st.text_input("輸入區間 (如: 1:50, 60)")
        if st.form_submit_button("🚀 確認並同步雲端"):
            targets = []
            if raw_val:
                items = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw_val).strip())
                for item in items:
                    if ':' in item or '-' in item:
                        try:
                            s, e = map(int, re.split(r'[:-]', item))
                            rng = range(s, e+1, step) if s <= e else range(s, e-1, -step)
                            for i in rng: targets.append(f"P{i}")
                        except: pass
                    elif item.isdigit():
                        targets.append(f"P{item}")
            sync_to_cloud(targets)

# 3. 視覺化顯示
if df_base is not None:
    df_plot = df_base.merge(df_history, left_on='樁號大寫', right_on='樁號', how='left', suffixes=('', '_c'))
    df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
    df_plot['顯示標籤'] = df_plot.apply(lambda r: f"{r['樁號']} ({int(r['施作順序'])})" if pd.notna(r['施作順序']) else r['樁號'], axis=1)

    st.sidebar.header("⚙️ 系統設定")
    h_val = st.sidebar.slider("圖面高度", 500, 2500, 1000)
    
    if st.sidebar.button("⚠️ 清空雲端數據"):
        conn.update(data=pd.DataFrame(columns=['樁號', '施工日期', '施作順序']))
        st.rerun()

    fig = px.scatter(
        df_plot, x='X', y='Y', text='顯示標籤', color='狀態',
        color_discrete_map={'未完成': 'lightgrey'},
        hover_data={'X': False, 'Y': False, '顯示標籤': False}
    )
    fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='DarkSlateGrey')))
    fig.update_layout(
        xaxis=dict(visible=False, showgrid=False),
        yaxis=dict(scaleanchor="x", scaleratio=1, visible=False, showgrid=False),
        plot_bgcolor='white', margin=dict(l=0, r=0, t=0, b=0), height=h_val,
        legend=dict(title="施工日期 (自動換色)", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

    c1, c2 = st.columns(2)
    today_done = len(df_history[df_history['施工日期'] == today_str])
    c1.metric(f"📅 {today_str} 進度", f"{today_done} 支")
    c2.metric("🏆 累計總進度", f"{len(df_history)} / 613")
