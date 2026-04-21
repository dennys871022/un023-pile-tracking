import streamlit as st
import pandas as pd
import plotly.express as px
import re
import datetime
import io

st.set_page_config(page_title="UN023 排樁進度管理系統", layout="wide")
st.title("📊 UN023 排樁進度管理系統 (圖文報表版)")

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
        
        return df.drop_duplicates(subset=['樁號']).dropna(subset=['X', 'Y']).sort_values('數字')
    except Exception as e:
        st.error(f"底圖讀取錯誤: {e}")
        return None

df_base = load_base_data()

# --- 2. 歷史紀錄管理 ---
if 'history' not in st.session_state:
    st.session_state['history'] = []

st.sidebar.header("📂 系統存檔與報表匯出")
upload_file = st.sidebar.file_uploader("1️⃣ 每日開工：匯入歷史 Excel 或 CSV", type=['csv', 'xlsx'])

if upload_file is not None:
    try:
        if upload_file.name.endswith('.csv'):
            df_hist = pd.read_csv(upload_file)
        else:
            # 如果上傳的是 Excel，只讀取「施工明細」這頁
            df_hist = pd.read_excel(upload_file, sheet_name='施工明細')
        
        # 確保順序為整數，清除重複
        if '機台' not in df_hist.columns:
            df_hist['機台'] = 'A車'
        st.session_state['history'] = df_hist.drop_duplicates(subset=['樁號']).to_dict('records')
        st.sidebar.success("✅ 歷史紀錄已成功同步")
    except Exception as e:
        st.sidebar.error("檔案讀取失敗，請確認檔案。")

# 刪除重置按鈕
if st.sidebar.button("🗑️ 清空所有紀錄"):
    st.session_state['history'] = []
    st.rerun()

# --- 3. 施工登錄 ---
st.markdown("### 📝 進度登錄")
c1, c2, c3 = st.columns([1, 1, 2])
work_date = str(c1.date_input("施工日期"))
machine = c2.radio("施工機台：", ["A車", "B車"], horizontal=True)
mode = c3.radio("模式：", ["連續", "4支一循環", "3支一循環"], horizontal=True)

step = 1
if "4支" in mode: step = 4
elif "3支" in mode: step = 3

def save_data(piles):
    hist_df = pd.DataFrame(st.session_state['history']) if st.session_state['history'] else pd.DataFrame(columns=['樁號','機台','施作順序'])
    
    # 獨立計算該機台的最後編號
    m_data = hist_df[hist_df['機台'] == machine]
    last_seq = 0 if m_data.empty else m_data['施作順序'].max()
    
    new_entries = []
    for p in piles:
        p = p.upper().strip()
        # 排除已施作的樁號，防止重複登錄
        if not any(d['樁號'] == p for d in st.session_state['history']):
            last_seq += 1
            new_entries.append({
                '樁號': p,
                '施工日期': work_date,
                '機台': machine,
                '施作順序': int(last_seq)
            })
    
    if new_entries:
        st.session_state['history'].extend(new_entries)
        st.success(f"✅ {machine} 成功登錄 {len(new_entries)} 支。")
    else:
        st.info("ℹ️ 這些樁號已經登錄過了。")

t1, t2 = st.tabs(["🎯 起點推算", "✏️ 區間輸入"])
with t1:
    with st.form("auto"):
        cc1, cc2, cc3 = st.columns(3)
        start = cc1.number_input("起點 P", 1, 613, 1)
        direct = cc2.radio("方向", ["遞增", "遞減"])
        num = cc3.number_input("數量", 1, 100, 10)
        if st.form_submit_button("執行登錄"):
            plist = []
            curr = start
            for _ in range(num):
                if 1 <= curr <= 613: plist.append(f"P{curr}")
                curr = curr + step if direct == "遞增" else curr - step
            save_data(plist)
with t2:
    with st.form("manual"):
        raw = st.text_input("輸入區間 (例如: 1-50 或 100-92)")
        if st.form_submit_button("執行登錄"):
            plist = []
            if raw:
                items = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw))
                for i in items:
                    if '-' in i:
                        s, e = map(int, i.split('-'))
                        # 修正邊界問題：確保 e 也會被納入計算
                        rng = range(s, e + 1, step) if s <= e else range(s, e - step, -step)
                        for n in rng: plist.append(f"P{n}")
                    elif i.isdigit(): plist.append(f"P{i}")
            save_data(plist)

# --- 4. 網頁圖面預覽與合併資料 ---
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
fig.update_traces(textposition='top center', marker=dict(size=10, line=dict(width=1, color='DarkSlateGrey')))
fig.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=750, margin=dict(l=0, r=0, t=0, b=0))
st.plotly_chart(fig, use_container_width=True)

# --- 5. 唯一 Excel 報表匯出 (放置於側邊欄) ---
st.sidebar.markdown("---")
if st.session_state['history']:
    st.sidebar.success(f"目前累積 {len(st.session_state['history'])} 支完成進度")
    
    def export_excel(df_exp, df_full_plot):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 建立第一頁：施工明細 (同時包含座標，供散佈圖使用)
            df_export_full = df_full_plot.copy()
            df_export_full['是否完成'] = df_export_full['狀態'].apply(lambda x: '未完成' if x == '未完成' else '已完成')
            
            # 只抽出已完成的明細供查看
            df_completed = df_export_full[df_export_full['是否完成'] == '已完成']
            df_completed[['樁號', '施工日期', '機台', '施作順序', 'X', 'Y']].to_excel(writer, sheet_name='施工明細', index=False)
            
            # 建立第二頁：全區進度圖 (XY 散佈圖)
            workbook = writer.book
            worksheet = workbook.add_worksheet('全區進度圖')
            writer.sheets['全區進度圖'] = worksheet
            
            # 將全部的點位(含完成與未完成)寫入隱藏的欄位中，讓 Excel 讀取作圖
            undone = df_export_full[df_export_full['是否完成'] == '未完成']
            done = df_export_full[df_export_full['是否完成'] == '已完成']
            
            undone[['X', 'Y']].to_excel(writer, sheet_name='全區進度圖', startcol=15, index=False)
            done[['X', 'Y']].to_excel(writer, sheet_name='全區進度圖', startcol=18, index=False)
            
            chart = workbook.add_chart({'type': 'scatter'})
            
            # 畫未完成 (灰色)
            if not undone.empty:
                chart.add_series({
                    'name': '未完成樁位',
                    'categories': ['全區進度圖', 1, 15, len(undone), 15],
                    'values':     ['全區進度圖', 1, 16, len(undone), 16],
                    'marker':     {'type': 'circle', 'size': 5, 'fill': {'color': '#D3D3D3'}, 'border': {'color': '#D3D3D3'}},
                })
            
            # 畫已完成 (紅色)
            if not done.empty:
                chart.add_series({
                    'name': '今日/已完成樁位',
                    'categories': ['全區進度圖', 1, 18, len(done), 18],
                    'values':     ['全區進度圖', 1, 19, len(done), 19],
                    'marker':     {'type': 'circle', 'size': 8, 'fill': {'color': 'red'}, 'border': {'color': 'black'}},
                })
            
            chart.set_title({'name': '排樁全區施工進度分佈圖'})
            chart.set_x_axis({'visible': False})
            chart.set_y_axis({'visible': False})
            chart.set_size({'width': 800, 'height': 600})
            
            worksheet.insert_chart('B2', chart)
            
        return output.getvalue()

    excel_file = export_excel(pd.DataFrame(st.session_state['history']), df_plot)
    
    # 這個就是唯一、且最明顯的匯出按鈕！
    st.sidebar.download_button(
        label="📥 2️⃣ 收工：匯出 Excel 總報表", 
        data=excel_file, 
        file_name=f"排樁進度總報表_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
else:
    st.sidebar.info("尚無資料可供匯出")
