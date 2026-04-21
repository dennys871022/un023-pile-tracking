import streamlit as st
import pandas as pd
import plotly.express as px
import re
import datetime
import io

st.set_page_config(page_title="UN023 排樁進度系統 V4", layout="wide")
st.title("🚧 UN023 排樁進度管理系統 (高穩定報表版)")

# --- 1. 座標底圖讀取 (強化容錯) ---
@st.cache_data
def load_base_data():
    try:
        # 嘗試多種編碼讀取
        try:
            df = pd.read_csv('排樁座標.csv', encoding='utf-8')
        except:
            df = pd.read_csv('排樁座標.csv', encoding='big5')
        
        # 欄位模糊匹配
        x_col = next((c for c in df.columns if 'X' in c.upper() or '座標' in c), None)
        y_col = next((c for c in df.columns if 'Y' in c.upper() or '座標' in c), None)
        text_col = next((c for c in df.columns if '內容' in c or '值' in c or '樁號' in c), None)
        
        if not all([x_col, y_col, text_col]):
            st.error("CSV 欄位格式不符，請檢查是否有 '位置 X', '位置 Y', '內容' 欄位")
            return None

        # 清洗內容並過濾 P1-P613
        df['樁號'] = df[text_col].apply(lambda x: re.sub(r'\\[^;]+;|[{}]', '', str(x)).strip().upper())
        df = df[df['樁號'].str.match(r'^P\d+$')]
        df['數字'] = df['樁號'].str.extract(r'(\d+)').astype(int)
        df = df[(df['數字'] >= 1) & (df['數字'] <= 613)]
        
        df['X'] = pd.to_numeric(df[x_col], errors='coerce')
        df['Y'] = pd.to_numeric(df[y_col], errors='coerce')
        
        # 去重，防止重疊報錯
        return df.drop_duplicates(subset=['樁號']).dropna(subset=['X', 'Y']).sort_values('數字')
    except Exception as e:
        st.error(f"底圖載入失敗: {e}")
        return None

df_base = load_base_data()

# --- 2. 數據持久化 ---
if 'history' not in st.session_state:
    st.session_state['history'] = []

st.sidebar.header("📂 數據導入與存檔")
up_file = st.sidebar.file_uploader("1️⃣ 導入歷史資料 (Excel/CSV)", type=['csv', 'xlsx'])

if up_file:
    try:
        if up_file.name.endswith('.csv'):
            df_up = pd.read_csv(up_file)
        else:
            # 優先讀取「施工明細」分頁
            df_up = pd.read_excel(up_file, sheet_name=None)
            if '施工明細' in df_up:
                df_up = df_up['施工明細']
            else:
                df_up = list(df_up.values())[0] # 找不到就讀第一頁
        
        # 標準化欄位
        df_up.columns = [c.strip() for c in df_up.columns]
        if '機台' not in df_up.columns: df_up['機台'] = 'A車'
        
        st.session_state['history'] = df_up.drop_duplicates(subset=['樁號']).to_dict('records')
        st.sidebar.success("✅ 資料同步成功")
    except Exception as e:
        st.sidebar.error(f"讀取失敗: {e}")

if st.sidebar.button("🗑️ 重設所有進度"):
    st.session_state['history'] = []
    st.rerun()

# --- 3. 施工作業登錄 ---
st.markdown("### 📝 進度錄入")
c1, c2, c3 = st.columns([1, 1, 2])
work_date = str(c1.date_input("施工日期"))
machine = c2.radio("施工機台：", ["A車", "B車"], horizontal=True)
mode = c3.radio("模式：", ["連續", "4支循環", "3支循環"], horizontal=True)

step = 1
if "4支" in mode: step = 4
elif "3支" in mode: step = 3

def save_piles(piles):
    if not piles: return
    hist_df = pd.DataFrame(st.session_state['history']) if st.session_state['history'] else pd.DataFrame(columns=['樁號','機台','施作順序'])
    
    # *** 關鍵：AB車個別計數 ***
    m_data = hist_df[hist_df['機台'] == machine]
    seq = 0 if m_data.empty else m_data['施作順序'].max()
    
    added = 0
    for p in piles:
        p = p.upper().strip()
        # 檢查是否已做過
        if not any(d['樁號'] == p for d in st.session_state['history']):
            seq += 1
            st.session_state['history'].append({
                '樁號': p,
                '施工日期': work_date,
                '機台': machine,
                '施作順序': int(seq)
            })
            added += 1
    
    if added > 0:
        st.success(f"✅ {machine} 已成功新增 {added} 支！")
    else:
        st.warning("⚠️ 這些樁號皆已存在，未重複登錄。")

tab_auto, tab_man = st.tabs(["🎯 起點推算", "✏️ 手動區間"])
with tab_auto:
    with st.form("auto"):
        cc1, cc2, cc3 = st.columns(3)
        start_p = cc1.number_input("起始 P", 1, 613, 1)
        direct = cc2.radio("方向", ["遞增", "遞減"])
        count_p = cc3.number_input("支數", 1, 100, 10)
        if st.form_submit_button("執行登錄"):
            plist = []
            curr = start_p
            for _ in range(count_p):
                if 1 <= curr <= 613: plist.append(f"P{curr}")
                curr = curr + step if direct == "遞增" else curr - step
            save_piles(plist)

with tab_man:
    with st.form("manual"):
        raw_val = st.text_input("輸入區間 (如: 1-50 或 100-92)")
        if st.form_submit_button("執行登錄"):
            plist = []
            if raw_val:
                items = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw_val))
                for it in items:
                    if '-' in it:
                        s, e = map(int, it.split('-'))
                        # 修正：確保包含結束點
                        r_step = step if s <= e else -step
                        r_end = e + 1 if s <= e else e - 1
                        for n in range(s, r_end, r_step):
                            plist.append(f"P{n}")
                    elif it.isdigit():
                        plist.append(f"P{it}")
            save_piles(plist)

# --- 4. 平面圖預覽 ---
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
fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='black')))
fig.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=800)
st.plotly_chart(fig, use_container_width=True)

# --- 5. Excel 報表匯出 (修正圖表與資料遺失問題) ---
st.sidebar.markdown("---")
if st.session_state['history']:
    def get_excel_report(history_list, base_df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 1. 施工明細
            df_exp = pd.DataFrame(history_list)
            # 合併座標
            df_full = df_exp.merge(base_df[['樁號', 'X', 'Y']], on='樁號', how='left')
            df_full.to_excel(writer, sheet_name='施工明細', index=False)
            
            # 2. 圖表統計與資料點位
            workbook = writer.book
            worksheet = workbook.add_worksheet('全區進度圖')
            
            # 準備已完成與未完成的座標數據 (寫入 Excel 遠端欄位供圖表讀取)
            done = df_plot[df_plot['狀態'] != '未完成']
            undone = df_plot[df_plot['狀態'] == '未完成']
            
            # 將 XY 數據寫入 worksheet 方便圖表抓取 (隱藏在第 20 欄以後)
            undone[['X', 'Y']].to_excel(writer, sheet_name='全區進度圖', startcol=20, index=False)
            done[['X', 'Y']].to_excel(writer, sheet_name='全區進度圖', startcol=23, index=False)
            
            # 建立 XY 散佈圖
            chart = workbook.add_chart({'type': 'scatter'})
            
            # 未完成點
            if not undone.empty:
                chart.add_series({
                    'name': '未完成',
                    'categories': ['全區進度圖', 1, 20, len(undone), 20],
                    'values':     ['全區進度圖', 1, 21, len(undone), 21],
                    'marker':     {'type': 'circle', 'size': 4, 'fill': {'color': '#E0E0E0'}, 'border': {'color': '#E0E0E0'}},
                })
            
            # 已完成點
            if not done.empty:
                chart.add_series({
                    'name': '已完成進度',
                    'categories': ['全區進度圖', 1, 23, len(done), 23],
                    'values':     ['全區進度圖', 1, 24, len(done), 24],
                    'marker':     {'type': 'circle', 'size': 7, 'fill': {'color': 'red'}, 'border': {'color': 'black'}},
                    'data_labels': {'value': False}
                })
            
            chart.set_title({'name': '排樁工程全區施工進度圖'})
            chart.set_x_axis({'visible': False})
            chart.set_y_axis({'visible': False, 'reverse': False})
            chart.set_size({'width': 900, 'height': 600})
            
            worksheet.insert_chart('B2', chart)
            
        return output.getvalue()

    excel_out = get_excel_report(st.session_state['history'], df_base)
    st.sidebar.download_button(
        label="📥 2️⃣ 收工：匯出 Excel 完整報表",
        data=excel_out,
        file_name=f"UN023_進度報表_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
else:
    st.sidebar.info("請先登錄進度")
