import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import plotly.express as px
import re
import datetime

st.set_page_config(page_title="UN023 排樁進度管理系統", layout="wide")
st.title("🚧 UN023 排樁進度管理系統 (雲端同步版)")

# --- 1. 建立雲端串接 ---
conn = st.connection("gsheets", type=GSheetsConnection)

# 讀取底圖座標 (自動去重)
@st.cache_data
def load_base_data():
    try:
        try:
            df = pd.read_csv('排樁座標.csv', encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv('排樁座標.csv', encoding='big5')
    except Exception as e:
        st.error(f"讀取底圖失敗: {e}")
        return None
    
    # 自動抓取座標與內容欄位
    x_col = next((col for col in df.columns if 'X' in col.upper()), None)
    y_col = next((col for col in df.columns if 'Y' in col.upper()), None)
    text_col = next((col for col in df.columns if '內容' in col or '值' in col), None)
    
    # 清洗與過濾 (限制 P1-P613)
    df['樁號'] = df[text_col].apply(lambda x: re.sub(r'\\[^;]+;|[{}]', '', str(x)).strip())
    df['樁號大寫'] = df['樁號'].str.upper()
    df_piles = df[df['樁號大寫'].str.match(r'^P\d+$')].copy()
    df_piles['數字'] = df_piles['樁號大寫'].str.extract(r'(\d+)').astype(int)
    df_piles = df_piles[df_piles['數字'] <= 613]
    
    df_piles['X'] = pd.to_numeric(df_piles[x_col], errors='coerce')
    df_piles['Y'] = pd.to_numeric(df_piles[y_col], errors='coerce')
    
    # 關鍵：去除重複的樁號座標，防止圖面點位混亂
    return df_piles.drop_duplicates(subset=['樁號大寫']).dropna(subset=['X', 'Y'])

df_base = load_base_data()

# 讀取試算表紀錄 (ttl=0 確保每次操作都重新獲取最新雲端數據)
def get_cloud_data():
    try:
        data = conn.read(ttl=0)
        # 確保欄位存在
        if data.empty:
            return pd.DataFrame(columns=['樁號', '施工日期', '施作順序'])
        return data
    except:
        return pd.DataFrame(columns=['樁號', '施工日期', '施作順序'])

df_history = get_cloud_data()

# --- 2. 施工進度登錄邏輯 ---
st.markdown("### 📝 今日施工登錄")
c_date, c_mode = st.columns([1, 3])
today_str = str(c_date.date_input("施工日期", datetime.date.today()))
mode_opt = c_mode.radio("模式選擇：", ["連續施作", "4支一循環 (1, 5...)", "3支一循環 (1, 4...)"], horizontal=True)

step = 1
if "4支" in mode_opt: step = 4
elif "3支" in mode_opt: step = 3

# 同步函數：確保不重複寫入
def sync_to_cloud(input_list):
    # 1. 移除輸入清單中的重複項
    unique_input = list(dict.fromkeys(input_list))
    
    # 2. 獲取當前最新歷史
    current_history = get_cloud_data()
    max_seq = 0 if current_history.empty else current_history['施作順序'].max()
    
    new_rows = []
    for pid in unique_input:
        pid = pid.upper().strip()
        # 檢查是否已在雲端存在
        if not current_history.empty and pid in current_history['樁號'].values:
            continue
        
        # 限制 P1-P613
        match = re.search(r'\d+', pid)
        if match and int(match.group()) <= 613:
            max_seq += 1
            new_rows.append({
                '樁號': pid,
                '施工日期': today_str,
                '施作順序': max_seq
            })
    
    if new_rows:
        updated_df = pd.concat([current_history, pd.DataFrame(new_rows)], ignore_index=True)
        conn.update(data=updated_df)
        st.success(f"✅ 雲端同步完成！新增 {len(new_rows)} 支。")
        st.cache_data.clear() # 清除繪圖緩存
        st.rerun() # 強制刷新畫面
    else:
        st.warning("⚠️ 所選樁號皆已存在於雲端或超出範圍，未更新。")

# 輸入模式分頁
tab1, tab2 = st.tabs(["🎯 起點自動推算模式", "✏️ 區間輸入模式"])

with tab1:
    with st.form("auto_form"):
        col1, col2, col3 = st.columns(3)
        s_num = col1.number_input("起始樁號 (數字)", 1, 613, 1)
        dir_opt = col2.radio("方向", ["遞增 (+)", "遞減 (-)"])
        total_piles = col3.number_input("施作數量", 1, 100, 10)
        
        if st.form_submit_button("🚀 確認並同步至雲端"):
            targets = []
            curr = s_num
            for _ in range(total_piles):
                if curr < 1 or curr > 613: break
                targets.append(f"P{curr}")
                curr = curr + step if "遞增" in dir_opt else curr - step
            sync_to_cloud(targets)

with tab2:
    with st.form("manual_form"):
        raw_val = st.text_input("輸入區間 (例如: 1-50, 60)")
        if st.form_submit_button("🚀 確認並同步至雲端"):
            targets = []
            if raw_val:
                items = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw_val).strip())
                for item in items:
                    if '-' in item:
                        try:
                            s, e = map(int, item.split('-'))
                            rng = range(s, e+1, step) if s <= e else range(s, e-1, -step)
                            for i in rng: targets.append(f"P{i}")
                        except: pass
                    elif item.isdigit():
                        targets.append(f"P{item}")
            sync_to_cloud(targets)

# --- 3. 圖面顯示 ---
if df_base is not None:
    # 合併雲端歷史資料
    df_plot = df_base.merge(df_history, left_on='樁號大寫', right_on='樁號', how='left', suffixes=('', '_cloud'))
    df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
    
    # 產生含順序的標籤
    df_plot['顯示標籤'] = df_plot.apply(
        lambda r: f"{r['樁號']} ({int(r['施作順序'])})" if pd.notna(r['施作順序']) else r['樁號'], axis=1
    )

    # 側邊欄設定
    st.sidebar.header("⚙️ 圖面設定")
    st.sidebar.write(f"**累計完成：** {len(df_history)} 支")
    h_slider = st.sidebar.slider("調整圖面高度", 500, 2500, 1000)
    
    if st.sidebar.button("🧨 危險：清空雲端試算表"):
        conn.update(data=pd.DataFrame(columns=['樁號', '施工日期', '施作順序']))
        st.sidebar.success("雲端已清空")
        st.rerun()

    # 繪圖
    fig = px.scatter(
        df_plot, x='X', y='Y', text='顯示標籤', color='狀態',
        color_discrete_map={'未完成': 'lightgrey'}, # 只有未完成固定灰色，其餘日期自動上色
        hover_data={'X': False, 'Y': False, '顯示標籤': False}
    )
    
    fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='DarkSlateGrey')))
    fig.update_layout(
        xaxis=dict(visible=False, showgrid=False),
        yaxis=dict(scaleanchor="x", scaleratio=1, visible=False, showgrid=False),
        plot_bgcolor='white', margin=dict(l=0, r=0, t=0, b=0), height=h_slider,
        legend=dict(title="施工日期 (換色)", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

    # 進度統計卡片
    c1, c2 = st.columns(2)
    today_done = len(df_history[df_history['施工日期'] == today_str])
    c1.metric(f"📅 {today_str} 進度", f"{today_done} 支")
    c2.metric("🏆 累計總進度", f"{len(df_history)} / 613")
