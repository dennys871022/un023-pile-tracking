import streamlit as st
import pandas as pd
import plotly.express as px
import re
import datetime
import io
import xlsxwriter.utility
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="UN023 排樁進度系統 V8", layout="wide")
st.title("🏗️ UN023 排樁工程全自動進度管理系統")

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

# --- 2. 雲端服務初始化 ---
@st.cache_resource
def init_gspread():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = json.loads(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(st.secrets["sheet_url"]).sheet1
        
        headers = sheet.row_values(1)
        if not headers:
            sheet.append_row(['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
        return sheet
    except Exception as e:
        st.error(f"雲端連線失敗，請檢查 Secrets 設定。錯誤: {e}")
        return None

sheet = init_gspread()

def get_cloud_data():
    if sheet is None: return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
    try:
        records = sheet.get_all_records()
        if not records:
            return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
        df = pd.DataFrame(records)
        df['樁號'] = df['樁號'].astype(str).str.upper().str.strip()
        return df
    except:
        return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])

df_history = get_cloud_data()

# --- 3. 手動與自動備份邏輯 ---
st.sidebar.header("📂 備份與還原")
up_file = st.sidebar.file_uploader("匯入歷史進度 (CSV/Excel)", type=['csv', 'xlsx'])
if up_file:
    file_id = f"{up_file.name}_{up_file.size}"
    if st.session_state.get('loaded_file') != file_id:
        try:
            if up_file.name.endswith('.csv'):
                df_up = pd.read_csv(up_file)
            else:
                df_up = pd.read_excel(up_file, sheet_name='施工明細', engine='openpyxl')
            
            new_rows = []
            current_piles = df_history['樁號'].tolist()
            for _, row in df_up.iterrows():
                p = str(row['樁號']).upper().strip()
                if p not in current_piles:
                    base_info = df_base[df_base['樁號'] == p]
                    x_v = base_info['X'].iloc[0] if not base_info.empty else 0
                    y_v = base_info['Y'].iloc[0] if not base_info.empty else 0
                    new_rows.append([p, str(row['施工日期']), str(row.get('機台', 'A車')), int(row.get('施作順序', 1)), x_v, y_v])
            
            if new_rows:
                sheet.append_rows(new_rows)
                st.sidebar.success(f"已從檔案同步 {len(new_rows)} 筆至雲端")
                st.rerun()
            st.session_state['loaded_file'] = file_id
        except Exception as e:
            st.sidebar.error(f"檔案匯入失敗: {e}")

# --- 4. 進度登錄介面 ---
st.markdown("### 📝 施工進度登錄")
c1, c2, c3 = st.columns([1, 1, 2])
work_date = str(c1.date_input("日期"))
machine = c2.radio("機台", ["A車", "B車"], horizontal=True)
mode = c3.radio("模式", ["4支一循環 (1, 5...)", "3支一循環 (1, 4...)"], horizontal=True)

step = 4 if "4支" in mode else 3

def save_to_cloud(piles):
    if not piles or sheet is None: return
    m_data = df_history[df_history['機台'] == machine]
    seq = 0 if m_data.empty else pd.to_numeric(m_data['施作順序']).max()
    
    new_data = []
    for p in piles:
        p = p.upper().strip()
        if p not in df_history['樁號'].values:
            seq += 1
            base = df_base[df_base['樁號'] == p]
            x, y = (base['X'].iloc[0], base['Y'].iloc[0]) if not base.empty else (0, 0)
            new_data.append([p, work_date, machine, int(seq), float(x), float(y)])
    
    if new_data:
        sheet.append_rows(new_data)
        st.success(f"✅ {machine} 成功登錄 {len(new_data)} 支。")
        st.rerun()

tab_auto, tab_man = st.tabs(["🎯 起點推算", "✏️ 手動輸入"])
with tab_auto:
    with st.form("auto"):
        sc1, sc2, sc3 = st.columns(3)
        s_p = sc1.number_input("起始 P", 1, 613, 1)
        direct = sc2.radio("方向", ["遞增", "遞減"])
        count = sc3.number_input("數量", 1, 100, 10)
        if st.form_submit_button("執行登錄"):
            plist = []
            curr = s_p
            for _ in range(count):
                if 1 <= curr <= 613: plist.append(f"P{curr}")
                curr = curr + step if direct == "遞增" else curr - step
            save_to_cloud(plist)

with tab_man:
    with st.form("manual"):
        raw_in = st.text_input("輸入區間 (如: 1-50)")
        if st.form_submit_button("執行登錄"):
            plist = []
            if raw_in:
                parts = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw_in))
                for pt in parts:
                    if '-' in pt:
                        s, e = map(int, pt.split('-'))
                        r_step = step if s <= e else -step
                        for n in range(s, e + (1 if s <= e else -1), r_step): plist.append(f"P{n}")
                    elif pt.isdigit(): plist.append(f"P{pt}")
            save_to_cloud(plist)

# --- 5. 全區進度圖 (網頁嵌入版) ---
st.markdown("---")
st.subheader("🗺️ 全區施工進度圖")

# 圖表設定
with st.expander("⚙️ 圖表顯示設定"):
    sc1, sc2, sc3 = st.columns(3)
    show_labels = sc1.checkbox("顯示樁號標籤", value=True)
    f_size = sc2.slider("字體大小", 8, 24, 12)
    map_h = sc3.slider("地圖高度", 600, 2000, 1000)

# 合併資料 (防衝突處理)
df_plot = df_base.copy()
if not df_history.empty:
    hist_clean = df_history[['樁號', '施工日期', '機台', '施作順序']].copy()
    df_plot = df_plot.merge(hist_clean, on='樁號', how='left')
    df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
    df_plot['標籤'] = df_plot.apply(
        lambda r: f"{r['樁號']}({r['機台'][0]}{int(r['施作順序'])})" if pd.notna(r['施作順序']) else r['樁號'], axis=1
    )
else:
    df_plot['狀態'] = '未完成'
    df_plot['標籤'] = df_plot['樁號']

# 繪圖
fig = px.scatter(
    df_plot, x='X', y='Y', text='標籤' if show_labels else None,
    color='狀態', color_discrete_map={'未完成': '#4F4F4F'} # 深灰色
)
fig.update_traces(textposition='top center', textfont_size=f_size, marker=dict(size=14, line=dict(width=1, color='white')))
fig.update_layout(
    xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False),
    height=map_h, margin=dict(l=10, r=10, t=10, b=10), plot_bgcolor='white'
)
st.plotly_chart(fig, use_container_width=True)

# --- 6. Excel 導出功能 (保持原樣) ---
st.sidebar.markdown("---")
if not df_history.empty:
    def get_excel_report(history_df, full_plot_df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            history_df.to_excel(writer, sheet_name='施工明細', index=False)
            workbook = writer.book
            worksheet = workbook.add_worksheet('全區進度圖')
            chart = workbook.add_chart({'type': 'scatter'})
            col_idx = 10
            undone = full_plot_df[full_plot_df['狀態'] == '未完成']
            if not undone.empty:
                undone[['X', 'Y']].to_excel(writer, sheet_name='全區進度圖', startcol=col_idx, index=False)
                chart.add_series({
                    'name': '未完成',
                    'categories': ['全區進度圖', 1, col_idx, len(undone), col_idx],
                    'values':     ['全區進度圖', 1, col_idx+1, len(undone), col_idx+1],
                    'marker':     {'type': 'circle', 'size': 4, 'fill': {'color': '#696969'}, 'border': {'color': '#696969'}}
                })
                col_idx += 3
            dates = history_df['施工日期'].dropna().unique()
            for d in sorted(dates):
                date_data = full_plot_df[full_plot_df['施工日期'] == d].reset_index(drop=True)
                if not date_data.empty:
                    date_data[['X', 'Y', '標籤']].to_excel(writer, sheet_name='全區進度圖', startcol=col_idx, index=False)
                    custom_labels = [{'value': f'=全區進度圖!${xlsxwriter.utility.xl_col_to_name(col_idx + 2)}${ri + 2}'} for ri in range(len(date_data))]
                    chart.add_series({
                        'name': str(d),
                        'categories': ['全區進度圖', 1, col_idx, len(date_data), col_idx],
                        'values':     ['全區進度圖', 1, col_idx+1, len(date_data), col_idx+1],
                        'marker':     {'type': 'circle', 'size': 7},
                        'data_labels': {'custom': custom_labels, 'position': 'above'}
                    })
                    col_idx += 4
            chart.set_title({'name': '排樁全區進度圖'}); chart.set_size({'width': 2400, 'height': 1400}); worksheet.insert_chart('B2', chart)
        return output.getvalue()

    excel_out = get_excel_report(df_history, df_plot)
    st.sidebar.download_button("📥 匯出 Excel 總報表", excel_out, f"UN023_Report_{datetime.date.today()}.xlsx", type="primary")
