import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import plotly.express as px
import re
import datetime

st.set_page_config(page_title="UN023 排樁進度管理系統", layout="wide")
st.title("📊 UN023 排樁進度管理系統 (全自動雲端同步)")

# --- 1. 初始化連線 ---
conn = st.connection("gsheets", type=GSheetsConnection)

# 讀取座標底圖 (加入強力去重邏輯)
@st.cache_data
def load_base_data():
    try:
        # 嘗試讀取 CSV
        try:
            df = pd.read_csv('排樁座標.csv', encoding='utf-8')
        except:
            df = pd.read_csv('排樁座標.csv', encoding='big5')
        
        # 抓取必要欄位
        x_col = next((c for c in df.columns if 'X' in c.upper()), None)
        y_col = next((c for c in df.columns if 'Y' in c.upper()), None)
        text_col = next((c for c in df.columns if '內容' in c or '值' in c), None)
        
        # 資料清洗
        df['樁號'] = df[text_col].apply(lambda x: re.sub(r'\\[^;]+;|[{}]', '', str(x)).strip().upper())
        # 只保留 P1-P613
        df = df[df['樁號'].str.match(r'^P\d+$')]
        df['數字'] = df['樁號'].str.extract(r'(\d+)').astype(int)
        df = df[df['數字'] <= 613]
        
        # 座標轉數值
        df['X'] = pd.to_numeric(df[x_col], errors='coerce')
        df['Y'] = pd.to_numeric(df[y_col], errors='coerce')
        
        # *** 核心修正：強制去重 ***
        # 每個樁號只保留第一筆座標，徹底解決 Streamlit 選項重複錯誤
        df = df.drop_duplicates(subset=['樁號'])
        
        return df.dropna(subset=['X', 'Y']).sort_values('數字')
    except Exception as e:
        st.error(f"底圖載入出錯: {e}")
        return None

df_base = load_base_data()

# 讀取雲端試算表 (TTL=0 確保即時同步)
def get_cloud_data():
    try:
        df = conn.read(ttl=0)
        if df is None or df.empty:
            return pd.DataFrame(columns=['樁號', '施工日期', '施作順序'])
        # 確保資料格式正確
        df['樁號'] = df['樁號'].astype(str).str.upper().str.strip()
        df['施作順序'] = pd.to_numeric(df['施作順序'], errors='coerce')
        return df.drop_duplicates(subset=['樁號']) # 雲端也要去重
    except:
        return pd.DataFrame(columns=['樁號', '施工日期', '施作順序'])

df_history = get_cloud_data()

# --- 2. 施工登錄介面 ---
st.markdown("### 📝 施工進度登錄")

# 建立模式選擇 (使用靜態選項，避免動態生成導致錯誤)
c_date, c_mode = st.columns([1, 2])
today = str(c_date.date_input("施工日期", datetime.date.today()))
mode = c_mode.radio("跳支模式：", ["連續 (1, 2, 3...)", "4支一循環 (1, 5, 9...)", "3支一循環 (1, 4, 7...)"], horizontal=True)

step = 1
if "4支" in mode: step = 4
elif "3支" in mode: step = 3

def sync_to_cloud(new_piles):
    # 重新抓取最新歷史，避免覆蓋
    latest_hist = get_cloud_data()
    max_seq = 0 if latest_hist.empty else latest_hist['施作順序'].max()
    
    final_to_add = []
    for pid in new_piles:
        pid = pid.upper().strip()
        # 檢查是否已存在
        if pid not in latest_hist['樁號'].values:
            max_seq += 1
            final_to_add.append({'樁號': pid, '施工日期': today, '施作順序': max_seq})
    
    if final_to_add:
        updated_df = pd.concat([latest_hist, pd.DataFrame(final_to_add)], ignore_index=True)
        conn.update(data=updated_df)
        st.success(f"✅ 已成功更新 {len(final_to_add)} 支樁至雲端！")
        st.cache_data.clear()
        st.rerun()
    else:
        st.info("ℹ️ 輸入的樁號均已在紀錄中，無須更新。")

# 分頁輸入
tab1, tab2 = st.tabs(["🎯 起點自動推算", "✏️ 區間手動輸入"])

with tab1:
    with st.form("auto_form"):
        col1, col2, col3 = st.columns(3)
        s_num = col1.number_input("開始數字 (P)", 1, 613, 1)
        direct = col2.radio("方向", ["遞增 (+)", "遞減 (-)"])
        amount = col3.number_input("施作數量", 1, 100, 10)
        if st.form_submit_button("確認執行並同步雲端"):
            plist = []
            curr = s_num
            for _ in range(amount):
                if 1 <= curr <= 613:
                    plist.append(f"P{curr}")
                curr = curr + step if direct == "遞增 (+)" else curr - step
            sync_to_cloud(plist)

with tab2:
    with st.form("manual_form"):
        raw_in = st.text_input("輸入區間 (例如: 1-50, 60)")
        if st.form_submit_button("確認執行並同步雲端"):
            plist = []
            if raw_in:
                # 濾掉非數字字符但保留連字號
                clean = re.sub(r'[pP]', '', raw_in)
                parts = re.split(r'[,\s]+', clean.strip())
                for p in parts:
                    if '-' in p:
                        try:
                            s, e = map(int, p.split('-'))
                            rng = range(s, e+1, step) if s <= e else range(s, e-1, -step)
                            for i in rng: plist.append(f"P{i}")
                        except: pass
                    elif p.isdigit():
                        plist.append(f"P{p}")
            sync_to_cloud(plist)

# --- 3. 圖面渲染 ---
if df_base is not None:
    # 合併資料
    df_plot = df_base.merge(df_history, on='樁號', how='left')
    df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
    
    # 產生顯示標籤
    df_plot['標籤'] = df_plot.apply(
        lambda r: f"{r['樁號']} ({int(r['施作順序'])})" if pd.notna(r['施作順序']) else r['樁號'], axis=1
    )

    # 側邊欄設定
    st.sidebar.header("⚙️ 圖面設定")
    st.sidebar.metric("已完成總數", len(df_history))
    h = st.sidebar.slider("畫布高度", 600, 2500, 1000)
    
    if st.sidebar.button("🧨 危險：重設雲端資料"):
        conn.update(data=pd.DataFrame(columns=['樁號', '施工日期', '施作順序']))
        st.rerun()

    # 繪圖
    fig = px.scatter(
        df_plot, x='X', y='Y', text='標籤', color='狀態',
        color_discrete_map={'未完成': 'lightgrey'}, # 只有未完成固定灰色
        hover_data={'X': False, 'Y': False, '標籤': False}
    )
    
    fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='DarkSlateGrey')))
    fig.update_layout(
        xaxis=dict(visible=False, showgrid=False),
        yaxis=dict(scaleanchor="x", scaleratio=1, visible=False, showgrid=False),
        plot_bgcolor='white', margin=dict(l=0, r=0, t=0, b=0), height=h,
        legend=dict(title="施工日期 (自動換色)", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

    # 底部統計報告
    st.markdown("---")
    st.subheader(f"📊 今日施工摘要 ({today})")
    today_data = df_history[df_history['施工日期'] == today]
    st.write(f"今日完成支數: **{len(today_data)}** 支")
    if not today_data.empty:
        st.download_button("📥 下載今日報表 (CSV)", today_data.to_csv(index=False).encode('utf-8-sig'), f"Daily_Report_{today}.csv")
