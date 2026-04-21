import streamlit as st
import pandas as pd
import plotly.express as px
import re
import datetime

st.set_page_config(page_title="UN023 排樁進度管理系統", layout="wide")
st.title("📊 UN023 排樁進度管理系統 (單機高穩定版)")

# --- 1. 讀取座標底圖 ---
@st.cache_data
def load_base_data():
    try:
        try:
            df = pd.read_csv('排樁座標.csv', encoding='utf-8')
        except:
            df = pd.read_csv('排樁座標.csv', encoding='big5')
        
        x_col = next((c for c in df.columns if 'X' in c.upper()), None)
        y_col = next((c for c in df.columns if 'Y' in c.upper()), None)
        text_col = next((c for c in df.columns if '內容' in c or '值' in c), None)
        
        df['樁號'] = df[text_col].apply(lambda x: re.sub(r'\\[^;]+;|[{}]', '', str(x)).strip().upper())
        df = df[df['樁號'].str.match(r'^P\d+$')]
        df['數字'] = df['樁號'].str.extract(r'(\d+)').astype(int)
        df = df[df['數字'] <= 613]
        
        df['X'] = pd.to_numeric(df[x_col], errors='coerce')
        df['Y'] = pd.to_numeric(df[y_col], errors='coerce')
        
        # 強制去重，防止圖面點位重複報錯
        df = df.drop_duplicates(subset=['樁號'])
        return df.dropna(subset=['X', 'Y']).sort_values('數字')
    except Exception as e:
        st.error(f"底圖載入出錯: {e}")
        return None

df_base = load_base_data()

# --- 2. 歷史進度管理 (側邊欄上傳/下載) ---
st.sidebar.header("📂 進度檔案管理")
history_file = st.sidebar.file_uploader("1️⃣ 每日開工：匯入昨日進度檔 (CSV)", type="csv")

if 'history' not in st.session_state:
    st.session_state['history'] = []

# 讀取上傳的歷史紀錄
if history_file is not None:
    try:
        df_hist = pd.read_csv(history_file)
        # 確保不會重複載入
        st.session_state['history'] = df_hist.drop_duplicates(subset=['樁號']).to_dict('records')
        st.sidebar.success("✅ 歷史進度已成功匯入！")
    except Exception as e:
        st.sidebar.error("檔案讀取失敗，請確認格式。")

# 匯出進度按鈕 (每日收工)
if st.session_state['history']:
    df_download = pd.DataFrame(st.session_state['history'])
    csv_data = df_download.to_csv(index=False).encode('utf-8-sig')
    st.sidebar.download_button(
        label="2️⃣ 每日收工：下載最新進度存檔",
        data=csv_data,
        file_name=f"排樁進度紀錄_{datetime.date.today()}.csv",
        mime="text/csv",
        type="primary"
    )
    
if st.sidebar.button("🗑️ 清空目前網頁暫存紀錄"):
    st.session_state['history'] = []
    st.rerun()

st.sidebar.markdown("---")

# --- 3. 施工登錄介面 ---
st.markdown("### 📝 施工進度登錄")

c_date, c_mode = st.columns([1, 2])
today = str(c_date.date_input("施工日期", datetime.date.today()))
mode = c_mode.radio("跳支模式：", ["連續 (1, 2...)", "4支一循環 (1, 5...)", "3支一循環 (1, 4...)"], horizontal=True)

step = 1
if "4支" in mode: step = 4
elif "3支" in mode: step = 3

def process_piles(new_piles):
    current_max_seq = 0
    if st.session_state['history']:
        current_max_seq = max([int(item['施作順序']) for item in st.session_state['history'] if pd.notna(item['施作順序'])])
    
    added_count = 0
    for pid in new_piles:
        pid = pid.upper().strip()
        # 檢查是否已在歷史紀錄中
        if not any(x['樁號'] == pid for x in st.session_state['history']):
            current_max_seq += 1
            st.session_state['history'].append({
                '樁號': pid,
                '施工日期': today,
                '施作順序': current_max_seq
            })
            added_count += 1
            
    if added_count > 0:
        st.success(f"✅ 已登錄！新增 {added_count} 支樁。")
    else:
        st.info("ℹ️ 輸入的樁號均已登錄過，無新增。")

tab1, tab2 = st.tabs(["🎯 起點自動推算", "✏️ 區間手動輸入"])

with tab1:
    with st.form("auto_form"):
        col1, col2, col3 = st.columns(3)
        s_num = col1.number_input("開始數字 (P)", 1, 613, 1)
        direct = col2.radio("方向", ["遞增 (+)", "遞減 (-)"])
        amount = col3.number_input("施作數量", 1, 100, 10)
        if st.form_submit_button("確認登錄"):
            plist = []
            curr = s_num
            for _ in range(amount):
                if 1 <= curr <= 613:
                    plist.append(f"P{curr}")
                curr = curr + step if "遞增" in direct else curr - step
            process_piles(plist)

with tab2:
    with st.form("manual_form"):
        raw_in = st.text_input("輸入區間 (例如: 1-50, 60)")
        if st.form_submit_button("確認登錄"):
            plist = []
            if raw_in:
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
            process_piles(plist)

# --- 4. 圖面合併與渲染 ---
if df_base is not None:
    df_plot = df_base.copy()
    
    if st.session_state['history']:
        df_hist = pd.DataFrame(st.session_state['history'])
        df_plot = df_plot.merge(df_hist, on='樁號', how='left')
        df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
        df_plot['標籤'] = df_plot.apply(
            lambda r: f"{r['樁號']} ({int(r['施作順序'])})" if pd.notna(r['施作順序']) else r['樁號'], axis=1
        )
    else:
        df_plot['狀態'] = '未完成'
        df_plot['標籤'] = df_plot['樁號']

    st.sidebar.metric("目前網頁暫存完成數", len(st.session_state['history']))
    h = st.sidebar.slider("畫布高度", 600, 2500, 1000)

    fig = px.scatter(
        df_plot, x='X', y='Y', text='標籤', color='狀態',
        color_discrete_map={'未完成': 'lightgrey'}, 
        hover_data={'X': False, 'Y': False, '標籤': False}
    )
    
    fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='DarkSlateGrey')))
    fig.update_layout(
        xaxis=dict(visible=False, showgrid=False),
        yaxis=dict(scaleanchor="x", scaleratio=1, visible=False, showgrid=False),
        plot_bgcolor='white', margin=dict(l=0, r=0, t=0, b=0), height=h,
        legend=dict(title="施工日期", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

    st.markdown("---")
    st.subheader(f"📊 今日新增摘要 ({today})")
    if st.session_state['history']:
        today_data = [x for x in st.session_state['history'] if x['施工日期'] == today]
        st.write(f"今日新增支數: **{len(today_data)}** 支")
    else:
        st.write("今日新增支數: **0** 支")
