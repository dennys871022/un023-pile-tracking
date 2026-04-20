import streamlit as st
import pandas as pd
import plotly.express as px
import re
import datetime

st.set_page_config(page_title="排樁進度管理系統", layout="wide")
st.title("UN023 排樁工程進度管理系統")

# 1. 自動讀取底圖（從 GitHub 專案目錄讀取）
@st.cache_data
def load_base_data():
    base_filename = '排樁座標.csv'
    try:
        try:
            df = pd.read_csv(base_filename, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(base_filename, encoding='big5')
    except FileNotFoundError:
        return None
    
    # 尋找座標與內容欄位
    x_col = next((col for col in df.columns if 'X' in col.upper()), None)
    y_col = next((col for col in df.columns if 'Y' in col.upper()), None)
    text_col = next((col for col in df.columns if '內容' in col or '值' in col or 'VALUE' in col.upper()), None)

    if not (x_col and y_col and text_col): return None

    # 清理 AutoCAD 格式文字
    def clean_text(text):
        if not isinstance(text, str): return str(text)
        text = re.sub(r'\\[^;]+;', '', text)
        text = re.sub(r'[{}]', '', text)
        return text.strip()

    df['樁號'] = df[text_col].apply(clean_text)
    # 統一樁號格式並篩選
    df['樁號大寫'] = df['樁號'].str.upper().str.strip()
    df_piles = df[df['樁號大寫'].str.match(r'^P\d+$')].copy()
    
    df_piles['X'] = pd.to_numeric(df_piles[x_col], errors='coerce')
    df_piles['Y'] = pd.to_numeric(df_piles[y_col], errors='coerce')
    return df_piles.dropna(subset=['X', 'Y'])

df_base = load_base_data()

if df_base is None:
    st.error("找不到底圖檔案。請確認『排樁座標.csv』已上傳至 GitHub。")
    st.stop()

# 2. 進度檔案管理
st.sidebar.header("📂 進度存檔管理")
history_file = st.sidebar.file_uploader("匯入昨日進度存檔 (CSV)", type="csv")

if 'history' not in st.session_state:
    st.session_state['history'] = []

if history_file is not None:
    df_hist = pd.read_csv(history_file)
    st.session_state['history'] = df_hist.to_dict('records')
    st.sidebar.success("✅ 歷史紀錄已匯入")

# 下載按鈕
if st.session_state['history']:
    df_dl = pd.DataFrame(st.session_state['history'])
    csv_data = df_dl.to_csv(index=False).encode('utf-8-sig')
    st.sidebar.download_button(
        label="💾 下載最新進度存檔",
        data=csv_data,
        file_name=f"排樁進度_{datetime.date.today()}.csv",
        mime="text/csv"
    )

if st.sidebar.button("🗑️ 清空本次所有紀錄"):
    st.session_state['history'] = []
    st.rerun()

# 3. 施工登錄介面
st.markdown("### 📝 施工進度登錄")
col_date, col_mode = st.columns([1, 3])
with col_date:
    today_date = st.date_input("施工日期", datetime.date.today())
with col_mode:
    mode = st.radio("施工模式：", ["連續 (1, 2...)", "4支循環 (1, 5...)", "3支循環 (1, 4...)"], horizontal=True)

step = 1
if "4支" in mode: step = 4
elif "3支" in mode: step = 3

def process_piles(pile_list):
    current_max_seq = 0
    if st.session_state['history']:
        current_max_seq = max([item['施作順序'] for item in st.session_state['history']])
    
    added = 0
    for p_id in pile_list:
        # 樁號上限 613
        p_num_match = re.search(r'\d+', p_id)
        if p_num_match and int(p_num_match.group()) > 613:
            continue
        
        # 檢查是否重複
        if not any(x['樁號'] == p_id for x in st.session_state['history']):
            current_max_seq += 1
            st.session_state['history'].append({
                '樁號': p_id,
                '施工日期': str(today_date),
                '施作順序': current_max_seq
            })
            added += 1
    return added

t1, t2 = st.tabs(["🎯 起點自動推算", "✏️ 手動區間輸入"])

with t1:
    with st.form("auto_form"):
        c1, c2, c3 = st.columns(3)
        start_num = c1.number_input("起始數字", min_value=1, max_value=613, value=1)
        direction = c2.radio("方向", ["遞增 (+)", "遞減 (-)"])
        count = c3.number_input("施作支數", min_value=1, value=10)
        if st.form_submit_button("✅ 確認執行"):
            piles = []
            curr = start_num
            for _ in range(count):
                if curr < 1 or curr > 613: break
                piles.append(f"P{curr}")
                curr = curr + step if "遞增" in direction else curr - step
            added = process_piles(piles)
            st.success(f"成功記錄 {added} 支樁")

with t2:
    with st.form("manual_form"):
        raw_input = st.text_input("輸入區間 (如 1-50, 60)：")
        if st.form_submit_button("✅ 確認執行"):
            piles = []
            if raw_input:
                items = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw_input).strip())
                for item in items:
                    if '-' in item:
                        p = item.split('-')
                        if len(p) == 2:
                            s, e = int(p[0]), int(p[1])
                            rng = range(s, e + 1, step) if s <= e else range(s, e - 1, -step)
                            for i in rng: piles.append(f"P{i}")
                    elif item.isdigit():
                        piles.append(f"P{item}")
            added = process_piles(piles)
            st.success(f"成功記錄 {added} 支樁")

# 4. 資料合併與圖面顯示
df_plot = df_base.copy()
if st.session_state['history']:
    df_hist = pd.DataFrame(st.session_state['history'])
    df_plot = df_plot.merge(df_hist, left_on='樁號大寫', right_on='樁號', how='left', suffixes=('', '_h'))
    df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
    df_plot['標籤'] = df_plot.apply(lambda r: f"{r['樁號']} ({int(r['施作順序'])})" if pd.notna(r['施作順序']) else r['樁號'], axis=1)
else:
    df_plot['狀態'] = '未完成'
    df_plot['標籤'] = df_plot['樁號']

# 5. 繪製
st.sidebar.markdown("***")
h = st.sidebar.slider("畫布高度", 500, 2500, 800)

fig = px.scatter(
    df_plot, x='X', y='Y', text='標籤', color='狀態',
    color_discrete_map={'未完成': 'lightgrey'},
    hover_data={'X': False, 'Y': False, '標籤': False}
)

fig.update_traces(textposition='top center', marker=dict(size=10, line=dict(width=1, color='DarkSlateGrey')))
fig.update_layout(
    xaxis=dict(visible=False, showgrid=False),
    yaxis=dict(scaleanchor="x", scaleratio=1, visible=False, showgrid=False),
    plot_bgcolor='white', margin=dict(l=0, r=0, t=0, b=0), height=h
)

st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

# 6. 統計
st.markdown(f"### 📊 統計摘要 (截至 {today_date})")
c1, c2 = st.columns(2)
c1.metric("今日進度", len(df_plot[df_plot['狀態'] == str(today_date)]))
c2.metric("累計總完成", len(df_plot[df_plot['狀態'] != '未完成']))
