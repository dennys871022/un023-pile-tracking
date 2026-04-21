import streamlit as st
import pandas as pd
import plotly.express as px
import re
import datetime
import io
import xlsxwriter.utility

st.set_page_config(page_title="UN023 排樁進度系統", layout="wide")
st.title("🚧 UN023 排樁進度管理系統 (日期換色+標籤全開版)")

# --- 1. 座標底圖讀取 ---
@st.cache_data
def load_base_data():
    try:
        try:
            df = pd.read_csv('排樁座標.csv', encoding='utf-8')
        except:
            df = pd.read_csv('排樁座標.csv', encoding='big5')
        
        x_col = next((c for c in df.columns if 'X' in c.upper() or '座標' in c), None)
        y_col = next((c for c in df.columns if 'Y' in c.upper() or '座標' in c), None)
        text_col = next((c for c in df.columns if '內容' in c or '值' in c or '樁號' in c), None)
        
        df['樁號'] = df[text_col].apply(lambda x: re.sub(r'\\[^;]+;|[{}]', '', str(x)).strip().upper())
        df = df[df['樁號'].str.match(r'^P\d+$')]
        df['數字'] = df['樁號'].str.extract(r'(\d+)').astype(int)
        df = df[(df['數字'] >= 1) & (df['數字'] <= 613)]
        
        df['X'] = pd.to_numeric(df[x_col], errors='coerce')
        df['Y'] = pd.to_numeric(df[y_col], errors='coerce')
        
        return df.drop_duplicates(subset=['樁號']).dropna(subset=['X', 'Y']).sort_values('數字')
    except Exception as e:
        st.error(f"底圖載入失敗: {e}")
        return None

df_base = load_base_data()

# --- 2. 數據導入與管理 ---
if 'history' not in st.session_state:
    st.session_state['history'] = []

st.sidebar.header("📂 數據導入")
up_file = st.sidebar.file_uploader("1️⃣ 每日開工：導入歷史 Excel 報表", type=['csv', 'xlsx'])

if up_file:
    try:
        if up_file.name.endswith('.csv'):
            df_up = pd.read_csv(up_file)
        else:
            # 指定 openpyxl 引擎讀取 Excel
            df_up = pd.read_excel(up_file, sheet_name='施工明細', engine='openpyxl')
        
        if '機台' not in df_up.columns: df_up['機台'] = 'A車'
        
        st.session_state['history'] = df_up.drop_duplicates(subset=['樁號']).to_dict('records')
        st.sidebar.success("✅ 歷史資料已成功同步！")
    except Exception as e:
        st.sidebar.error(f"讀取失敗，錯誤碼: {e}")

if st.sidebar.button("🗑️ 清空網頁暫存"):
    st.session_state['history'] = []
    st.rerun()

# --- 3. 施工作業登錄 ---
st.markdown("### 📝 進度登錄")
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
    
    # 獨立 A/B 車的順序計數器
    m_data = hist_df[hist_df['機台'] == machine]
    seq = 0 if m_data.empty else m_data['施作順序'].max()
    
    added = 0
    for p in piles:
        p = p.upper().strip()
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
        st.success(f"✅ {machine} 已新增 {added} 支 (施作順序累計至 {seq})！")
    else:
        st.warning("⚠️ 樁號重複，未登錄。")

tab_auto, tab_man = st.tabs(["🎯 起點推算", "✏️ 區間輸入"])
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
                        r_step = step if s <= e else -step
                        r_end = e + 1 if s <= e else e - 1
                        for n in range(s, r_end, r_step): plist.append(f"P{n}")
                    elif it.isdigit(): plist.append(f"P{it}")
            save_piles(plist)

# --- 4. 網頁平面圖預覽 ---
df_plot = df_base.copy()
if st.session_state['history']:
    df_h = pd.DataFrame(st.session_state['history'])
    # 核心修復：強制刪除上傳檔案中帶有的 X, Y 座標，避免欄位重疊衝突
    df_h = df_h.drop(columns=['X', 'Y', '標籤', '狀態'], errors='ignore')
    
    df_plot = df_plot.merge(df_h, on='樁號', how='left')
    df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
    # 產生 P100(A1) 格式標籤
    df_plot['標籤'] = df_plot.apply(lambda r: f"{r['樁號']}({r['機台'][0]}{int(r['施作順序'])})" if pd.notna(r['施作順序']) else r['樁號'], axis=1)
else:
    df_plot['狀態'] = '未完成'
    df_plot['標籤'] = df_plot['樁號']

fig = px.scatter(df_plot, x='X', y='Y', text='標籤', color='狀態', color_discrete_map={'未完成':'lightgrey'})
fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='black')))
fig.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=800, legend_title_text='施工日期 (自動換色)')
st.plotly_chart(fig, use_container_width=True)

# --- 5. 強大 Excel 圖表報表匯出 ---
st.sidebar.markdown("---")
if st.session_state['history']:
    def get_excel_report(history_list, base_df, full_plot_df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 1. 施工明細分頁
            df_exp = pd.DataFrame(history_list)
            # 先清掉可能存在的舊座標，再與底圖座標合併
            df_exp = df_exp.drop(columns=['X', 'Y'], errors='ignore')
            df_full = df_exp.merge(base_df[['樁號', 'X', 'Y']], on='樁號', how='left')
            df_full.to_excel(writer, sheet_name='施工明細', index=False)
            
            # 2. 全區進度圖分頁
            workbook = writer.book
            worksheet = workbook.add_worksheet('全區進度圖')
            writer.sheets['全區進度圖'] = worksheet
            chart = workbook.add_chart({'type': 'scatter'})
            
            # 動態寫入數據：將資料依日期分開，以利 Excel 自動賦予不同顏色
            col_idx = 10 # 將輔助數據隱藏在 K 欄之後
            
            # 處理未完成資料
            undone = full_plot_df[full_plot_df['狀態'] == '未完成']
            if not undone.empty:
                undone[['X', 'Y']].to_excel(writer, sheet_name='全區進度圖', startcol=col_idx, index=False)
                chart.add_series({
                    'name': '未完成',
                    'categories': ['全區進度圖', 1, col_idx, len(undone), col_idx],
                    'values':     ['全區進度圖', 1, col_idx+1, len(undone), col_idx+1],
                    'marker':     {'type': 'circle', 'size': 4, 'fill': {'color': '#E0E0E0'}, 'border': {'color': '#E0E0E0'}}
                })
                col_idx += 3
            
            # 處理已完成資料 (依日期迴圈)
            dates = df_full['施工日期'].dropna().unique()
            for d in sorted(dates):
                date_data = full_plot_df[full_plot_df['施工日期'] == d].reset_index(drop=True)
                if not date_data.empty:
                    # 將 X, Y, 標籤 寫入 Excel
                    date_data[['X', 'Y', '標籤']].to_excel(writer, sheet_name='全區進度圖', startcol=col_idx, index=False)
                    
                    # 建立 Excel 客製化資料標籤 (顯示 P100(A1))
                    custom_labels = []
                    label_col_letter = xlsxwriter.utility.xl_col_to_name(col_idx + 2)
                    for row_idx in range(len(date_data)):
                        custom_labels.append({'value': f'=全區進度圖!${label_col_letter}${row_idx + 2}'})
                    
                    # 加入該日期的 Series
                    chart.add_series({
                        'name': str(d),
                        'categories': ['全區進度圖', 1, col_idx, len(date_data), col_idx],
                        'values':     ['全區進度圖', 1, col_idx+1, len(date_data), col_idx+1],
                        'marker':     {'type': 'circle', 'size': 7},
                        'data_labels': {'custom': custom_labels, 'position': 'above'}
                    })
                    col_idx += 4
            
            # 設定圖表外觀
            chart.set_title({'name': '排樁工程全區施工進度圖'})
            chart.set_x_axis({'visible': False})
            chart.set_y_axis({'visible': False, 'reverse': False})
            chart.set_size({'width': 1200, 'height': 800}) # 加大畫布避免文字擠壓
            worksheet.insert_chart('B2', chart)
            
        return output.getvalue()

    excel_out = get_excel_report(st.session_state['history'], df_base, df_plot)
    st.sidebar.download_button(
        label="📥 2️⃣ 收工：匯出 Excel 圖加表",
        data=excel_out,
        file_name=f"UN023_報表_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
