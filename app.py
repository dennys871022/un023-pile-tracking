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

st.set_page_config(page_title="UN023 排樁進度系統 V12", layout="wide")
st.title("🏗️ UN023 排樁進度管理 (雲端同步圖表版)")

# 1. 座標底圖讀取
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

# 2. 雲端即時連線 (不使用快取，確保能偵測到手動刪除分頁)
def get_gs_connection():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = json.loads(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        ss = client.open_by_url(st.secrets["sheet_url"])
        
        # 施工明細檢查/重建
        try:
            sh_main = ss.worksheet("施工明細")
        except:
            sh_main = ss.add_worksheet("施工明細", 1000, 20)
            sh_main.append_row(['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
            
        # 繪圖區檢查/重建
        try:
            sh_chart = ss.worksheet("系統繪圖區(勿動)")
        except:
            sh_chart = ss.add_worksheet("系統繪圖區(勿動)", 700, 60)
            
        return ss, sh_main, sh_chart
    except Exception as e:
        st.error(f"雲端連線異常: {e}")
        return None, None, None

# 3. 獲取資料
def fetch_current_data(sh_main):
    if sh_main is None: return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
    try:
        records = sh_main.get_all_records()
        if not records: return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
        df = pd.DataFrame(records)
        df['樁號'] = df['樁號'].astype(str).str.upper().str.strip()
        if '機台' not in df.columns: df['機台'] = 'A車'
        df['施作順序'] = pd.to_numeric(df.get('施作順序', 0), errors='coerce').fillna(0)
        return df
    except:
        return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])

ss, sh_main, sh_chart = get_gs_connection()
df_history = fetch_current_data(sh_main)

# 4. ★ 強制更新雲端彩色圖表與標籤 ★
def update_cloud_chart():
    # 重新獲取最新分頁控制權
    ss_now, m_now, c_now = get_gs_connection()
    if not ss_now or df_history.empty: return
    
    try:
        # 整理標籤與日期
        plot_df = df_base[['樁號', 'X', 'Y']].copy()
        hist = df_history.copy()
        def label_maker(r):
            m = str(r.get('機台', 'A'))[0]
            s = r.get('施作順序', 0)
            return f"{r['樁號']}({m}{int(s)})"
        hist['標籤'] = hist.apply(label_maker, axis=1)
        plot_df = plot_df.merge(hist[['樁號', '施工日期', '標籤']], on='樁號', how='left')

        # 建立多序列矩陣：A(X), B(標籤), C(未完成), D...(日期Y)
        gs_matrix = pd.DataFrame()
        gs_matrix['X'] = plot_df['X']
        gs_matrix['標籤'] = plot_df['標籤']
        gs_matrix['未完成'] = plot_df['Y'].where(plot_df['施工日期'].isna(), '')
        
        dates = sorted(plot_df['施工日期'].dropna().unique())
        for d in dates:
            gs_matrix[str(d)] = plot_df['Y'].where(plot_df['施工日期'] == d, '')

        # 清空並寫入繪圖分頁
        c_now.clear()
        c_now.update([gs_matrix.columns.values.tolist()] + gs_matrix.fillna('').values.tolist())

        # API 指令：刪除舊圖並建立包含圖例與日期的散佈圖
        m_id = m_now.id; c_id = c_now.id
        num_rows = len(gs_matrix) + 1
        num_cols = len(gs_matrix.columns)

        meta = ss_now.fetch_sheet_metadata()
        sheet_meta = next((s for s in meta['sheets'] if s['properties']['sheetId'] == m_id), None)
        
        reqs = []
        if sheet_meta and 'charts' in sheet_meta:
            for ch in sheet_meta['charts']:
                reqs.append({"deleteChart": {"chartId": ch['chartId']}})

        # 建立序列 (從「未完成」到所有日期)
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
                        "title": "排樁全區施工進度圖",
                        "basicChart": {
                            "chartType": "SCATTER",
                            "legendPosition": "RIGHT_LEGEND",
                            "axis": [
                                {"position": "BOTTOM_AXIS", "title": "X 座標"},
                                {"position": "LEFT_AXIS", "title": "Y 座標"}
                            ],
                            "domains": [{
                                "domain": {"sourceRange": {"sources": [{"sheetId": c_id, "startRowIndex": 0, "endRowIndex": num_rows, "startColumnIndex": 0, "endColumnIndex": 1}]}}
                            }],
                            "series": series
                        }
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {"sheetId": m_id, "rowIndex": 1, "columnIndex": 8},
                            "widthPixels": 1200, "heightPixels": 850
                        }
                    }
                }
            }
        })
        ss_now.batch_update({"requests": reqs})
        st.success("✅ 雲端圖表已強制重繪 (包含日期圖例)")
    except Exception as e:
        st.error(f"雲端繪圖引擎出錯: {e}")

# 5. UI 操作介面
st.sidebar.header("📂 數據備份")
if st.sidebar.button("🔄 強制同步雲端圖表"):
    update_cloud_chart()

up_file = st.sidebar.file_uploader("匯入歷史 Excel/CSV", type=['csv', 'xlsx'])
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
            sh_main.append_rows(new_rows)
            st.sidebar.success(f"已同步 {len(new_rows)} 筆")
            update_cloud_chart()
            st.rerun()
    except Exception as e: st.sidebar.error(f"還原失敗: {e}")

st.markdown("### 📝 進度登錄")
c1, c2, c3 = st.columns([1, 1, 2])
work_date = str(c1.date_input("施工日期"))
machine = c2.radio("機台", ["A車", "B車"], horizontal=True)
mode = c3.radio("模式", ["4支一循環", "3支一循環"], horizontal=True)
step = 4 if "4支" in mode else 3

def save_data(piles):
    if not piles or sh_main is None: return
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
        sh_main.append_rows(new_d)
        update_cloud_chart()
        st.rerun()

t1, t2 = st.tabs(["🎯 自動推算", "✏️ 手動輸入"])
with t1:
    with st.form("a"):
        cc1, cc2, cc3 = st.columns(3); sp = cc1.number_input("起始 P", 1, 613, 1)
        dr = cc2.radio("方向", ["遞增", "遞減"]); ct = cc3.number_input("支數", 1, 100, 10)
        if st.form_submit_button("執行登錄"):
            plist = []; cur = sp
            for _ in range(ct):
                if 1 <= cur <= 613: plist.append(f"P{cur}")
                cur = cur + step if dr == "遞增" else cur - step
            save_data(plist)
with t2:
    with st.form("m"):
        raw = st.text_input("輸入區間 (如 1-50)")
        if st.form_submit_button("執行登錄"):
            plist = []
            if raw:
                pts = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw))
                for pt in pts:
                    if '-' in pt:
                        s, e = map(int, pt.split('-')); rs = step if s <= e else -step
                        for n in range(s, e + (1 if s <= e else -1), rs): plist.append(f"P{n}")
                    elif pt.isdigit(): plist.append(f"P{pt}")
            save_data(plist)

# 6. 網頁平面圖 (含平移/縮放)
st.markdown("---")
st.subheader("🗺️ 現場施工圖 (左鍵平移 / 滾輪縮放)")
df_p = df_base.copy()
if not df_history.empty:
    hc = df_history.drop(columns=['X', 'Y', '數字'], errors='ignore')
    df_p = df_p.merge(hc, on='樁號', how='left')
    df_p['狀態'] = df_p['施工日期'].fillna('未完成')
    def lbl_gen(r):
        if pd.isna(r.get('施作順序')): return r['樁號']
        m = str(r.get('機台', 'A'))[0]
        return f"{r['樁號']}({m}{int(r['施作順序'])})"
    df_p['標籤'] = df_p.apply(lbl_gen, axis=1)
else:
    df_p['狀態'] = '未完成'; df_p['標籤'] = df_p['樁號']

fig = px.scatter(df_p, x='X', y='Y', text='標籤', color='狀態', color_discrete_map={'未完成': '#4F4F4F'})
fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='white')))
fig.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=950, plot_bgcolor='white', dragmode='pan')
st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

# 7. Excel 下載
if not df_history.empty:
    def xl_gen(h_df, p_df):
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
            ch.set_title({'name': '全區進度圖'}); ch.set_size({'width': 2400, 'height': 1500}); ws.insert_chart('B2', ch)
        return out.getvalue()
    st.sidebar.download_button("📥 匯出 Excel 總報表", xl_gen(df_history, df_p), f"Report_{datetime.date.today()}.xlsx", type="primary")
