import streamlit as st
import pandas as pd
import plotly.express as px
import re
import datetime
import io

st.set_page_config(page_title="UN023 排樁進度管理系統", layout="wide")
st.title("📊 UN023 排樁進度管理系統 (機台分流與圖文報表版)")

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
        
        df = df.drop_duplicates(subset=['樁號'])
        return df.dropna(subset=['X', 'Y']).sort_values('數字')
    except Exception as e:
        st.error(f"底圖載入出錯: {e}")
        return None

df_base = load_base_data()

# --- 2. 歷史進度管理 ---
if 'history' not in st.session_state:
    st.session_state['history'] = []

st.sidebar.header("📂 系統存檔區")
history_file = st.sidebar.file_uploader("1️⃣ 匯入昨日進度 CSV (接續紀錄)", type="csv")

if history_file is not None:
    try:
        df_hist = pd.read_csv(history_file)
        if '機台' not in df_hist.columns: df_hist['機台'] = 'A車'
        # 確保順序與日期欄位正確
        st.session_state['history'] = df_hist.drop_duplicates(subset=['樁號']).to_dict('records')
        st.sidebar.success("✅ 歷史紀錄已同步")
    except:
        st.sidebar.error("讀取失敗，請確認檔案格式。")

if st.session_state['history']:
    csv_data = pd.DataFrame(st.session_state['history']).to_csv(index=False).encode('utf-8-sig')
    st.sidebar.download_button("2️⃣ 下載今日備份 CSV", data=csv_data, file_name=f"Backup_{datetime.date.today()}.csv", mime="text/csv")

if st.sidebar.button("🗑️ 清空所有紀錄"):
    st.session_state['history'] = []
    st.rerun()

# --- 3. 施工登錄 ---
st.markdown("### 📝 施工進度登錄")
c1, c2, c3 = st.columns([1, 1, 2])
today = str(c1.date_input("日期"))
machine = c2.radio("施工機台：", ["A車", "B車"], horizontal=True)
mode = c3.radio("模式：", ["連續", "4支一循環", "3支一循環"], horizontal=True)

step = 1
if "4支" in mode: step = 4
elif "3支" in mode: step = 3

def process_piles(pile_list):
    hist_df = pd.DataFrame(st.session_state['history']) if st.session_state['history'] else pd.DataFrame(columns=['樁號','機台','施作順序'])
    
    # 分開計算機台順序
    m_hist = hist_df[hist_df['機台'] == machine]
    current_max = 0 if m_hist.empty else m_hist['施作順序'].max()
    
    added = 0
    for pid in pile_list:
        pid = pid.upper().strip()
        if not any(x['樁號'] == pid for x in st.session_state['history']):
            current_max += 1
            st.session_state['history'].append({
                '樁號': pid,
                '施工日期': today,
                '機台': machine,
                '施作順序': current_max
            })
            added += 1
    st.success(f"已登錄！{machine} 新增 {added} 支。")

t1, t2 = st.tabs(["起點推算", "區間輸入"])
with t1:
    with st.form("f1"):
        cc1, cc2, cc3 = st.columns(3)
        s_n = cc1.number_input("起始數字", 1, 613, 1)
        dir = cc2.radio("方向", ["遞增", "遞減"])
        num = cc3.number_input("支數", 1, 100, 10)
        if st.form_submit_button("確認登錄"):
            plist = []
            curr = s_n
            for _ in range(num):
                if 1 <= curr <= 613: plist.append(f"P{curr}")
                curr = curr + step if dir == "遞增" else curr - step
            process_piles(plist)
with t2:
    with st.form("f2"):
        raw = st.text_input("輸入區間 (如 1-50)")
        if st.form_submit_button("確認登錄"):
            plist = []
            if raw:
                nums = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw))
                for n in nums:
                    if '-' in n:
                        s, e = map(int, n.split('-'))
                        rng = range(s, e+1, step) if s <= e else range(s, e-1, -step)
                        for i in rng: plist.append(f"P{i}")
                    elif n.isdigit(): plist.append(f"P{n}")
            process_piles(plist)

# --- 4. 圖面渲染 ---
df_plot = df_base.copy()
if st.session_state['history']:
    df_h = pd.DataFrame(st.session_state['history'])
    df_plot = df_plot.merge(df_h, on='樁號', how='left')
    df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
    df_plot['標籤'] = df_plot.apply(lambda r: f"{r['樁號']}({r['機台'][0]}{int(r['施作順序'])})" if pd.notna(r['施作順序']) else r['樁號'], axis=1)
else:
    df_plot['狀態'] = '未完成'
    df_plot['標籤'] = df_plot['樁號']

fig = px.scatter(df_plot, x='X', y='Y', text='標籤', color='狀態', color_discrete_map={'未完成':'lightgrey'})
fig.update_traces(textposition='top center', marker=dict(size=12))
fig.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=800, margin=dict(l=0,r=0,t=0,b=0))
st.plotly_chart(fig, use_container_width=True)

# --- 5. 強大 Excel 報表匯出 ---
st.markdown("---")
if st.session_state['history']:
    def to_excel(df_exp, base_info):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 1. 明細表
            df_full = base_info.merge(df_exp, on='樁號', how='right')
            df_full[['樁號', '施工日期', '機台', '施作順序', 'X', 'Y']].to_excel(writer, sheet_name='施工明細', index=False)
            
            # 2. 圖表統計頁
            summary = df_exp.pivot_table(index='施工日期', columns='機台', values='樁號', aggfunc='count', fill_value=0)
            summary.to_excel(writer, sheet_name='進度圖表')
            
            workbook = writer.book
            worksheet = writer.sheets['進度圖表']
            
            # 建立長條圖 (A/B 車進度)
            chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
            for i, col in enumerate(summary.columns):
                chart.add_series({
                    'name':       ['進度圖表', 0, i + 1],
                    'categories': ['進度圖表', 1, 0, len(summary), 0],
                    'values':     ['進度圖表', 1, i + 1, len(summary), i + 1],
                })
            worksheet.insert_chart('E2', chart)
            
            # 建立樁位分布圖 (模擬圖色)
            scatter = workbook.add_chart({'type': 'scatter'})
            # 未完成
            unfinished = df_plot[df_plot['狀態'] == '未完成']
            # 已完成
            finished = df_plot[df_plot['狀態'] != '未完成']
            
            # 在 Excel 畫出點位圖
            scatter.add_series({
                'name': '已完成樁位',
                'categories': ['施工明細', 1, 4, len(df_full), 4], # X 座標
                'values':     ['施工明細', 1, 5, len(df_full), 5], # Y 座標
                'marker':     {'type': 'circle', 'size': 8, 'border': {'color': 'red'}, 'fill': {'color': 'red'}},
            })
            scatter.set_title({'name': '全區樁位分布進度圖'})
            scatter.set_x_axis({'visible': False})
            scatter.set_y_axis({'visible': False})
            worksheet.insert_chart('E25', scatter)
            
        return output.getvalue()

    excel_data = to_excel(pd.DataFrame(st.session_state['history']), df_base)
    st.download_button("📊 下載 Excel 完整報表 (含座標分佈圖)", data=excel_data, file_name=f"Report_{datetime.date.today()}.xlsx")
