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

st.set_page_config(page_title="UN023 排樁進度管理系統", layout="wide")
st.title("🏗️ UN023 排樁進度管理 (雲端全自動圖表版)")

# --- 1. 底圖讀取 ---
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

# --- 2. 雲端服務初始化 (建立明細與繪圖分頁) ---
@st.cache_resource
def init_gspread():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = json.loads(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(st.secrets["sheet_url"])
        
        try:
            sheet_main = spreadsheet.worksheet("施工明細")
        except:
            sheet_main = spreadsheet.add_worksheet("施工明細", 1000, 15)
            sheet_main.append_row(['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
            
        try:
            sheet_chart = spreadsheet.worksheet("系統繪圖區(勿動)")
        except:
            sheet_chart = spreadsheet.add_worksheet("系統繪圖區(勿動)", 700, 30)
            
        return spreadsheet, sheet_main, sheet_chart
    except Exception as e:
        st.error(f"雲端連線失敗: {e}")
        return None, None, None

spreadsheet, sheet_main, sheet_chart = init_gspread()

def get_cloud_data():
    if sheet_main is None: return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
    try:
        records = sheet_main.get_all_records()
        df = pd.DataFrame(records)
        df['樁號'] = df['樁號'].astype(str).str.upper().str.strip()
        df['施作順序'] = pd.to_numeric(df['施作順序'], errors='coerce')
        return df
    except:
        return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])

df_history = get_cloud_data()

# --- 3. ★ 核心功能：強制生成 Google Sheets 雲端圖表 ★ ---
def sync_gs_visuals():
    if not spreadsheet or sheet_main is None or sheet_chart is None: return
    
    try:
        # 準備餵給圖表的矩陣資料
        plot_df = df_base[['樁號', 'X', 'Y']].copy()
        if not df_history.empty:
            hist = df_history.copy()
            hist['標籤'] = hist.apply(lambda r: f"{r['樁號']}({r['機台'][0]}{int(r['施作順序'])})", axis=1)
            plot_df = plot_df.merge(hist[['樁號', '施工日期', '標籤']], on='樁號', how='left')
        else:
            plot_df['施工日期'] = float('nan')
            plot_df['標籤'] = plot_df['樁號']

        gs_data = pd.DataFrame()
        gs_data['標籤'] = plot_df['標籤']
        gs_data['X'] = plot_df['X']
        gs_data['未完成'] = plot_df['Y'].where(plot_df['施工日期'].isna(), '')

        dates = sorted(plot_df['施工日期'].dropna().unique())
        for d in dates:
            gs_data[str(d)] = plot_df['Y'].where(plot_df['施工日期'] == d, '')

        # 更新隱藏分頁的數據
        sheet_chart.clear()
        sheet_chart.update([gs_data.columns.values.tolist()] + gs_data.fillna('').values.tolist())

        # 透過 API 建立圖表
        main_id = sheet_main.id
        chart_id = sheet_chart.id
        num_rows = len(gs_data) + 1
        num_cols = len(gs_data.columns)

        # 尋找並刪除舊圖表
        ss_meta = spreadsheet.fetch_sheet_metadata({'includeGridData': False})
        main_sheet_meta = next((s for s in ss_meta.get('sheets', []) if s['properties']['sheetId'] == main_id), None)

        requests = []
        if main_sheet_meta and 'charts' in main_sheet_meta:
            for chart in main_sheet_meta['charts']:
                requests.append({"deleteChart": {"chartId": chart['chartId']}})

        # 建立多序列散佈圖 (按日期分色)
        series_list = []
        for i in range(2, num_cols):
            series_list.append({
                "series": {"sourceRange": {"sources": [{"sheetId": chart_id, "startRowIndex": 0, "endRowIndex": num_rows, "startColumnIndex": i, "endColumnIndex": i+1}]}},
                "targetAxis": "LEFT_AXIS"
            })

        requests.append({
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "雲端自動同步進度圖 (依日期分色)",
                        "basicChart": {
                            "chartType": "SCATTER",
                            "legendPosition": "RIGHT_LEGEND",
                            "axis": [{"position": "BOTTOM_AXIS", "title": "X 座標"}, {"position": "LEFT_AXIS", "title": "Y 座標"}],
                            "domains": [{"domain": {"sourceRange": {"sources": [{"sheetId": chart_id, "startRowIndex": 0, "endRowIndex": num_rows, "startColumnIndex": 1, "endColumnIndex": 2}]}}}],
                            "series": series_list
                        }
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {"sheetId": main_id, "rowIndex": 1, "columnIndex": 7},
                            "widthPixels": 900,
                            "heightPixels": 700
                        }
                    }
                }
            }
        })

        spreadsheet.batch_update({"requests": requests})
    except Exception as e:
        print(f"雲端圖表更新失敗: {e}") # 背景報錯不影響主程式

# --- 4. 數據備份與手動上傳 ---
st.sidebar.header("📂 備份與還原機制")
up_file = st.sidebar.file_uploader("匯入歷史進度 (Excel/CSV)", type=['csv', 'xlsx'])
if up_file:
    try:
        df_up = pd.read_excel(up_file, sheet_name='施工明細') if up_file.name.endswith('.xlsx') else pd.read_csv(up_file)
        new_rows = []
        current_piles = df_history['樁號'].tolist()
        for _, row in df_up.iterrows():
            p = str(row['樁號']).upper().strip()
            if p not in current_piles:
                base = df_base[df_base['樁號'] == p]
                x, y = (base['X'].iloc[0], base['Y'].iloc[0]) if not base.empty else (0,0)
                new_rows.append([p, str(row['施工日期']), str(row.get('機台','A車')), int(row.get('施作順序',1)), x, y])
        if new_rows:
            sheet_main.append_rows(new_rows)
            sync_gs_visuals() # 更新雲端圖表
            st.sidebar.success(f"已同步 {len(new_rows)} 筆至雲端")
            st.rerun()
    except Exception as e: st.sidebar.error(f"還原失敗: {e}")

# --- 5. 施工登錄 ---
st.markdown("### 📝 進度登錄")
c1, c2, c3 = st.columns([1, 1, 2])
work_date = str(c1.date_input("日期"))
machine = c2.radio("施工機台", ["A車", "B車"], horizontal=True)
mode = c3.radio("模式", ["4支一循環", "3支一循環"], horizontal=True)
step = 4 if "4支" in mode else 3

def save_to_cloud(piles):
    if not piles or sheet_main is None: return
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
        sheet_main.append_rows(new_data)
        sync_gs_visuals() # 同步更新 Google Sheets 圖表
        st.success("✅ 雲端同步與圖表更新完成！")
        st.rerun()

tab1, tab2 = st.tabs(["🎯 起點推算", "✏️ 手動輸入"])
with tab1:
    with st.form("auto"):
        sc1, sc2, sc3 = st.columns(3)
        s_p = sc1.number_input("起始 P", 1, 613, 1)
        direct = sc2.radio("方向", ["遞增", "遞減"])
        count = sc3.number_input("數量", 1, 100, 10)
        if st.form_submit_button("確認登錄"):
            plist = []
            curr = s_p
            for _ in range(count):
                if 1 <= curr <= 613: plist.append(f"P{curr}")
                curr = curr + step if direct == "遞增" else curr - step
            save_to_cloud(plist)
with tab2:
    with st.form("manual"):
        raw_in = st.text_input("輸入區間 (如: 1-50)")
        if st.form_submit_button("確認登錄"):
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

# --- 6. 網頁平面圖預覽 ---
st.markdown("---")
st.subheader("🗺️ 現場施工進度全區圖")
df_plot = df_base.copy()
if not df_history.empty:
    hist_clean = df_history.drop(columns=['X', 'Y', '數字'], errors='ignore')
    df_plot = df_plot.merge(hist_clean, on='樁號', how='left')
    df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
    df_plot['標籤'] = df_plot.apply(lambda r: f"{r['樁號']}({r['機台'][0]}{int(r['施作順序'])})" if pd.notna(r['施作順序']) else r['樁號'], axis=1)
else:
    df_plot['狀態'] = '未完成'
    df_plot['標籤'] = df_plot['樁號']

fig = px.scatter(df_plot, x='X', y='Y', text='標籤', color='狀態', color_discrete_map={'未完成': '#4F4F4F'})
fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='white')))
fig.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=800, plot_bgcolor='white')
st.plotly_chart(fig, use_container_width=True)

# --- 7. 下載最強 Excel 報表 ---
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
                chart.add_series({'name': '未完成', 'categories': ['全區進度圖', 1, col_idx, len(undone), col_idx], 'values': ['全區進度圖', 1, col_idx+1, len(undone), col_idx+1], 'marker': {'type': 'circle', 'size': 4, 'fill': {'color': '#696969'}, 'border': {'color': '#696969'}}})
                col_idx += 3
            dates = history_df['施工日期'].dropna().unique()
            for d in sorted(dates):
                date_data = full_plot_df[full_plot_df['施工日期'] == d].reset_index(drop=True)
                if not date_data.empty:
                    date_data[['X', 'Y', '標籤']].to_excel(writer, sheet_name='全區進度圖', startcol=col_idx, index=False)
                    custom_labels = [{'value': f'=全區進度圖!${xlsxwriter.utility.xl_col_to_name(col_idx + 2)}${ri + 2}'} for ri in range(len(date_data))]
                    chart.add_series({'name': str(d), 'categories': ['全區進度圖', 1, col_idx, len(date_data), col_idx], 'values': ['全區進度圖', 1, col_idx+1, len(date_data), col_idx+1], 'marker': {'type': 'circle', 'size': 7}, 'data_labels': {'custom': custom_labels, 'position': 'above'}})
                    col_idx += 4
            chart.set_title({'name': '全區進度圖'}); chart.set_size({'width': 2400, 'height': 1400}); worksheet.insert_chart('B2', chart)
        return output.getvalue()
    st.sidebar.download_button("📥 匯出 Excel 總報表", get_excel_report(df_history, df_plot), f"UN023_Report_{datetime.date.today()}.xlsx", type="primary")
