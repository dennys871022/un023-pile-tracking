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

st.set_page_config(page_title="UN023 排樁進度系統 V11", layout="wide")
st.title("🏗️ UN023 排樁進度管理 (自動偵測重建版)")

# 1. 底圖載入
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

# 2. 雲端連線函數 (動態獲取分頁)
def get_gs_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = json.loads(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        ss = client.open_by_url(st.secrets["sheet_url"])
        
        # 檢查並重建施工明細
        try:
            sh_main = ss.worksheet("施工明細")
        except:
            sh_main = ss.add_worksheet("施工明細", 1000, 15)
            sh_main.append_row(['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
            
        # 檢查並重建繪圖區
        try:
            sh_chart = ss.worksheet("系統繪圖區(勿動)")
        except:
            sh_chart = ss.add_worksheet("系統繪圖區(勿動)", 700, 50)
            
        return ss, sh_main, sh_chart
    except Exception as e:
        st.error(f"雲端連線失敗: {e}")
        return None, None, None

spreadsheet, sheet_main, sheet_chart = get_gs_sheets()

def get_cloud_data():
    if sheet_main is None: return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
    try:
        records = sheet_main.get_all_records()
        df = pd.DataFrame(records)
        df['樁號'] = df['樁號'].astype(str).str.upper().str.strip()
        df['施作順序'] = pd.to_numeric(df.get('施作順序', 0), errors='coerce').fillna(0)
        if '機台' not in df.columns: df['機台'] = 'A車'
        return df
    except:
        return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])

df_history = get_cloud_data()

# 3. 雲端圖表引擎 (包含重建邏輯)
def sync_gs_visuals():
    ss, sh_m, sh_c = get_gs_sheets()
    if not ss or df_history.empty: return
    
    try:
        plot_df = df_base[['樁號', 'X', 'Y']].copy()
        hist = df_history.copy()
        def label_func(r):
            m = str(r.get('機台', 'A'))[0]
            s = r.get('施作順序', 0)
            return f"{r['樁號']}({m}{int(s)})"
        hist['標籤'] = hist.apply(label_func, axis=1)
        plot_df = plot_df.merge(hist[['樁號', '施工日期', '標籤']], on='樁號', how='left')

        gs_data = pd.DataFrame()
        gs_data['X'] = plot_df['X']
        gs_data['標籤'] = plot_df['標籤']
        gs_data['未完成'] = plot_df['Y'].where(plot_df['施工日期'].isna(), '')

        dates = sorted(plot_df['施工日期'].dropna().unique())
        for d in dates:
            gs_data[str(d)] = plot_df['Y'].where(plot_df['施工日期'] == d, '')

        sh_c.clear()
        sh_c.update([gs_data.columns.values.tolist()] + gs_data.fillna('').values.tolist())

        m_id = sh_m.id; c_id = sh_c.id
        num_rows = len(gs_data) + 1; num_cols = len(gs_data.columns)

        meta = ss.fetch_sheet_metadata()
        target_meta = next((s for s in meta['sheets'] if s['properties']['sheetId'] == m_id), None)
        reqs = []
        if target_meta and 'charts' in target_meta:
            for c in target_meta['charts']: reqs.append({"deleteChart": {"chartId": c['chartId']}})

        series = []
        for i in range(2, num_cols):
            series.append({
                "series": {"sourceRange": {"sources": [{"sheetId": c_id, "startRowIndex": 0, "endRowIndex": num_rows, "startColumnIndex": i, "endColumnIndex": i+1}]}},
                "targetAxis": "LEFT_AXIS"
            })

        reqs.append({
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "雲端自動進度圖",
                        "basicChart": {
                            "chartType": "SCATTER", "legendPosition": "RIGHT_LEGEND",
                            "axis": [{"position": "BOTTOM_AXIS", "title": "X"}, {"position": "LEFT_AXIS", "title": "Y"}],
                            "domains": [{"domain": {"sourceRange": {"sources": [{"sheetId": c_id, "startRowIndex": 0, "endRowIndex": num_rows, "startColumnIndex": 0, "endColumnIndex": 1}]}}}],
                            "series": series
                        }
                    },
                    "position": {"overlayPosition": {"anchorCell": {"sheetId": m_id, "rowIndex": 1, "columnIndex": 8}, "widthPixels": 1000, "heightPixels": 800}}
                }
            }
        })
        ss.batch_update({"requests": reqs})
    except Exception as e:
        st.warning(f"圖表更新失敗: {e}")

# 4. 數據操作
st.sidebar.header("📂 數據備份與還原")
up_file = st.sidebar.file_uploader("匯入 Excel/CSV", type=['csv', 'xlsx'])
if up_file:
    try:
        df_up = pd.read_excel(up_file, sheet_name='施工明細') if up_file.name.endswith('.xlsx') else pd.read_csv(up_file)
        new_rows = []
        curr_p = df_history['樁號'].tolist()
        for _, row in df_up.iterrows():
            p = str(row['樁號']).upper().strip()
            if p not in curr_p:
                b = df_base[df_base['樁號'] == p]
                x, y = (b['X'].iloc[0], b['Y'].iloc[0]) if not b.empty else (0,0)
                new_rows.append([p, str(row['施工日期']), str(row.get('機台','A車')), int(row.get('施作順序',1)), x, y])
        if new_rows:
            sheet_main.append_rows(new_rows); sync_gs_visuals(); st.sidebar.success("同步完成"); st.rerun()
    except Exception as e: st.sidebar.error(f"還原失敗: {e}")

st.markdown("### 📝 進度登錄")
c1, c2, c3 = st.columns([1, 1, 2])
work_date = str(c1.date_input("日期"))
machine = c2.radio("機台", ["A車", "B車"], horizontal=True)
mode = c3.radio("模式", ["4支一循環", "3支一循環"], horizontal=True)
step = 4 if "4支" in mode else 3

def save_proc(piles):
    if not piles or sheet_main is None: return
    m_data = df_history[df_history['機台'] == machine]
    seq = 0 if m_data.empty else pd.to_numeric(m_data['施作順序']).max()
    new_d = []
    for p in piles:
        p = p.upper().strip()
        if p not in df_history['樁號'].values:
            seq += 1
            b = df_base[df_base['樁號'] == p]
            x, y = (b['X'].iloc[0], b['Y'].iloc[0]) if not b.empty else (0, 0)
            new_d.append([p, work_date, machine, int(seq), float(x), float(y)])
    if new_d:
        sheet_main.append_rows(new_d); sync_gs_visuals(); st.success("雲端已更新"); st.rerun()

t1, t2 = st.tabs(["🎯 推算", "✏️ 手動"])
with t1:
    with st.form("a"):
        sc1, sc2, sc3 = st.columns(3); sp = sc1.number_input("起始 P", 1, 613, 1)
        dr = sc2.radio("方向", ["遞增", "遞減"]); ct = sc3.number_input("數量", 1, 100, 10)
        if st.form_submit_button("登錄"):
            plist = []; cur = sp
            for _ in range(ct):
                if 1 <= cur <= 613: plist.append(f"P{cur}")
                cur = cur + step if dr == "遞增" else cur - step
            save_proc(plist)
with t2:
    with st.form("m"):
        raw = st.text_input("區間 (1-50)")
        if st.form_submit_button("登錄"):
            plist = []
            if raw:
                pts = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw))
                for pt in pts:
                    if '-' in pt:
                        s, e = map(int, pt.split('-')); rs = step if s <= e else -step
                        for n in range(s, e + (1 if s <= e else -1), rs): plist.append(f"P{n}")
                    elif pt.isdigit(): plist.append(f"P{pt}")
            save_proc(plist)

# 5. 網頁平面圖 (含縮放平移)
st.markdown("---")
st.subheader("🗺️ 現場施工全區圖 (左鍵平移 / 滾輪縮放)")
df_plot = df_base.copy()
if not df_history.empty:
    hc = df_history.drop(columns=['X', 'Y', '數字'], errors='ignore')
    df_plot = df_plot.merge(hc, on='樁號', how='left')
    df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
    def lbl_f(r):
        if pd.isna(r.get('施作順序')): return r['樁號']
        m = str(r.get('機台', 'A'))[0]
        return f"{r['樁號']}({m}{int(r['施作順序'])})"
    df_plot['標籤'] = df_plot.apply(lbl_f, axis=1)
else:
    df_plot['狀態'] = '未完成'; df_plot['標籤'] = df_plot['樁號']

fig = px.scatter(df_plot, x='X', y='Y', text='標籤', color='狀態', color_discrete_map={'未完成': '#4F4F4F'})
fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='white')))
fig.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=900, plot_bgcolor='white', dragmode='pan')
st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

# 6. 下載報表
if not df_history.empty:
    def exp_xl(h_df, p_df):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
            h_df.to_excel(wr, sheet_name='施工明細', index=False)
            wb = wr.book; ws = wb.add_worksheet('全區進度圖'); ch = wb.add_chart({'type': 'scatter'})
            col = 10; und = p_df[p_df['狀態'] == '未完成']
            if not und.empty:
                und[['X', 'Y']].to_excel(wr, sheet_name='全區進度圖', startcol=col, index=False)
                ch.add_series({'name': '未完成', 'categories': ['全區進度圖', 1, col, len(und), col], 'values': ['全區進度圖', 1, col+1, len(und), col+1], 'marker': {'type': 'circle', 'size': 4, 'fill': {'color': '#696969'}, 'border': {'color': '#696969'}}})
                col += 3
            dts = h_df['施工日期'].dropna().unique()
            for d in sorted(dts):
                dd = p_df[p_df['施工日期'] == d].reset_index(drop=True)
                if not dd.empty:
                    dd[['X', 'Y', '標籤']].to_excel(wr, sheet_name='全區進度圖', startcol=col, index=False)
                    clbls = [{'value': f'=全區進度圖!${xlsxwriter.utility.xl_col_to_name(col+2)}${ri+2}'} for ri in range(len(dd))]
                    ch.add_series({'name': str(d), 'categories': ['全區進度圖', 1, col, len(dd), col], 'values': ['全區進度圖', 1, col+1, len(dd), col+1], 'marker': {'type': 'circle', 'size': 7}, 'data_labels': {'custom': clbls, 'position': 'above'}})
                    col += 4
            ch.set_title({'name': '全區進度圖'}); ch.set_size({'width': 2400, 'height': 1400}); ws.insert_chart('B2', ch)
        return out.getvalue()
    st.sidebar.download_button("📥 匯出 Excel 總報表", exp_xl(df_history, df_plot), f"Report_{datetime.date.today()}.xlsx", type="primary")
