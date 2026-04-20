import streamlit as st
import pandas as pd
import plotly.express as px
import re
import datetime

st.set_page_config(page_title="排樁進度管理系統", layout="wide")
st.title("UN023 排樁工程進度自動化儀表板")

# --- 1. 自動讀取底圖 (從 GitHub 讀取) ---
@st.cache_data
def load_base_data():
    try:
        try:
            df = pd.read_csv('排樁座標.csv', encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv('排樁座標.csv', encoding='big5')
    except FileNotFoundError:
        return None
    
    x_col = next((col for col in df.columns if 'X' in col.upper()), None)
    y_col = next((col for col in df.columns if 'Y' in col.upper()), None)
    text_col = next((col for col in df.columns if '內容' in col or '值' in col or 'VALUE' in col.upper()), None)

    if not (x_col and y_col and text_col): return None

    def clean_autocad_text(text):
        if not isinstance(text, str): return str(text)
        text = re.sub(r'\\[^;]+;', '', text)
        text = re.sub(r'[{}]', '', text)
        return text.strip()

    df['樁號'] = df[text_col].apply(clean_autocad_text)
    pile_pattern = re.compile(r'^[A-Za-z]\-?\d+$')
    df_piles = df[df['樁號'].apply(lambda x: bool(pile_pattern.match(x)))].copy()
    
    df_piles['X'] = pd.to_numeric(df_piles[x_col], errors='coerce')
    df_piles['Y'] = pd.to_numeric(df_piles[y_col], errors='coerce')
    df_piles['樁號大寫'] = df_piles['樁號'].str.upper().str.strip()
    return df_piles.dropna(subset=['X', 'Y'])

df_base = load_base_data()

if df_base is None:
    st.error("找不到『排樁座標.csv』。請確認已將檔案上傳至 GitHub 專案庫中。")
    st.stop()

# --- 2. 歷史進度管理 (側邊欄) ---
st.sidebar.header("📂 進度檔案管理")
history_file = st.sidebar.file_uploader("匯入歷史進度存檔 (CSV)", type="csv")

if 'history' not in st.session_state:
    st.session_state['history'] = []

if history_file is not None:
    df_hist = pd.read_csv(history_file)
    st.session_state['history'] = df_hist.to_dict('records')
    st.sidebar.success("✅ 歷史進度已成功匯入！")

# 匯出進度按鈕
if st.session_state['history']:
    df_download = pd.DataFrame(st.session_state['history'])
    csv_data = df_download.to_csv(index=False).encode('utf-8-sig')
    st.sidebar.download_button(
        label="💾 下載最新進度存檔 (收工前務必下載)",
        data=csv_data,
        file_name=f"排樁進度紀錄_{datetime.date.today()}.csv",
        mime="text/csv"
    )

st.sidebar.markdown("***")

# --- 3. 今日進度登錄 ---
st.markdown("### 📝 施工進度登錄")

col_date, col_mode = st.columns([1, 3])
with col_date:
    today_date = st.date_input("施工日期")
with col_mode:
    mode = st.radio("施工模式：", ["連續施工 (1, 2...)", "4支一循環 (1, 5...)", "3支一循環 (1, 4...)"], horizontal=True)

step = 1
if "4支" in mode: step = 4
elif "3支" in mode: step = 3

tab1, tab2 = st.tabs(["🎯 起點自動推算模式", "✏️ 自由輸入模式 (區間)"])

# 處理紀錄寫入的函數
def process_piles(pile_list):
    current_max_seq = 0
    if st.session_state['history']:
        current_max_seq = max([item['施作順序'] for item in st.session_state['history']])
    
    added_count = 0
    for p_id in pile_list:
        # 檢查是否已存在紀錄中
        if not any(x['樁號'] == p_id for x in st.session_state['history']):
            current_max_seq += 1
            st.session_state['history'].append({
                '樁號': p_id,
                '施工日期': str(today_date),
                '施作順序': current_max_seq
            })
            added_count += 1
    return added_count

with tab1:
    with st.form("calc_form"):
        c1, c2, c3 = st.columns(3)
        with c1: start_num = st.number_input("起始樁號數字", min_value=1, value=1, step=1)
        with c2: direction = st.radio("方向", ["編號遞增 (+)", "編號遞減 (-)"])
        with c3: count = st.number_input("預計施作數量", min_value=1, value=10, step=1)

        if st.form_submit_button("✅ 寫入進度"):
            temp_list = []
            curr = start_num
            for _ in range(count):
                if curr < 1: break
                p_id = f"P{curr}"
                if p_id not in temp_list: temp_list.append(p_id)
                curr = curr + step if "遞增" in direction else curr - step
            
            added = process_piles(temp_list)
            st.success(f"已記錄！今日新增 {added} 支樁。")

with tab2:
    with st.form("free_form"):
        completed_input = st.text_input("輸入完成樁號區間 (如: 1-100, 105)：")
        if st.form_submit_button("✅ 寫入進度"):
            temp_list = []
            if completed_input:
                clean_input = re.sub(r'[pP]', '', completed_input)
                raw_items = re.split(r'[,\s]+', clean_input.strip())
                for item in raw_items:
                    if not item: continue
                    if '-' in item:
                        parts = item.split('-')
                        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                            s_num, e_num = int(parts[0]), int(parts[1])
                            rng = range(s_num, e_num + 1, step) if s_num <= e_num else range(s_num, e_num - 1, -step)
                            for i in rng:
                                p_id = f"P{i}"
                                if p_id not in temp_list: temp_list.append(p_id)
                    elif item.isdigit():
                        p_id = f"P{item}"
                        if p_id not in temp_list: temp_list.append(p_id)
            
            added = process_piles(temp_list)
            st.success(f"已記錄！今日新增 {added} 支樁。")

# --- 4. 資料合併與動態標籤 ---
df_plot = df_base.copy()
df_plot['狀態'] = '未完成'
df_plot['施作順序'] = None
df_plot['顯示標籤'] = df_plot['樁號']

if st.session_state['history']:
    df_hist = pd.DataFrame(st.session_state['history'])
    # 將歷史紀錄合併至底圖
    df_plot = df_plot.merge(df_hist, left_on='樁號大寫', right_on='樁號', how='left', suffixes=('', '_hist'))
    
    # 狀態改為施工日期
    df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
    df_plot['施作順序'] = df_plot['施作順序']
    
    def make_label(row):
        if pd.notna(row['施作順序']):
            return f"{row['樁號']} ({int(row['施作順序'])})"
        return row['樁號']
    
    df_plot['顯示標籤'] = df_plot.apply(make_label, axis=1)

# --- 5. 繪製視覺化圖面與裁切設定 ---
st.sidebar.header("🛠️ 圖面裁切設定")
x_min_val, x_max_val = float(df_base['X'].min()), float(df_base['X'].max())
y_min_val, y_max_val = float(df_base['Y'].min()), float(df_base['Y'].max())

x_range = st.sidebar.slider("X 軸範圍", x_min_val, x_max_val, (x_min_val, x_max_val))
y_range = st.sidebar.slider("Y 軸範圍", y_min_val, y_max_val, (y_min_val, y_max_val))

df_final = df_plot[(df_plot['X'] >= x_range[0]) & (df_plot['X'] <= x_range[1]) & (df_plot['Y'] >= y_range[0]) & (df_plot['Y'] <= y_range[1])]

# 確保 '未完成' 永遠是灰色，其他日期交由 Plotly 自動分配顏色
color_map = {'未完成': 'lightgrey'}

fig = px.scatter(
    df_final, x='X', y='Y', text='顯示標籤', color='狀態',
    color_discrete_map=color_map,
    hover_data={'X': False, 'Y': False, '顯示標籤': False}
)

fig.update_traces(
    textposition='top center', 
    marker=dict(size=10, line=dict(width=1, color='DarkSlateGrey')), 
    textfont=dict(size=12, color='black') 
)

if not df_final.empty:
    x_min, x_max = df_final['X'].min(), df_final['X'].max()
    y_min, y_max = df_final['Y'].min(), df_final['Y'].max()
    x_margin, y_margin = (x_max - x_min) * 0.05, (y_max - y_min) * 0.05

    fig.update_layout(
        xaxis=dict(range=[x_min - x_margin, x_max + x_margin], visible=False, showgrid=False), 
        yaxis=dict(range=[y_min - y_margin, y_max + y_margin], scaleanchor="x", scaleratio=1, visible=False, showgrid=False),
        plot_bgcolor='white', margin=dict(l=0, r=0, t=0, b=0), height=850,
        legend=dict(title="施工日期", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

# --- 6. 進度統計 ---
today_str = str(today_date)
today_completed = len(df_final[df_final['狀態'] == today_str])
total_completed = len(df_final[df_final['狀態'] != '未完成'])

c1, c2 = st.columns(2)
c1.metric(label=f"📅 {today_str} 完成數量", value=today_completed)
c2.metric(label="✅ 累計總完成數量", value=total_completed)
