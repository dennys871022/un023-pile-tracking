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

st.set_page_config(page_title="UN023 排樁進度系統 V17", layout="wide")
st.title("🏗️ UN023 排樁進度管理 (自動歸檔版)")

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

def get_gs_connection():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = json.loads(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        ss = client.open_by_url(st.secrets["sheet_url"])
        
        try:
            sh_main = ss.worksheet("施工明細")
        except:
            sh_main = ss.add_worksheet("施工明細", 1000, 20)
            sh_main.append_row(['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
            
        try:
            sh_chart = ss.worksheet("系統繪圖區")
        except:
            sh_chart = ss.add_worksheet("系統繪圖區", 700, 60)
            
        return ss, sh_main, sh_chart
    except Exception as e:
        st.error(f"雲端連線異常: {e}")
        return None, None, None

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

def process_status_logic(df_hist, df_b):
    plot_df = df_b[['樁號', 'X', 'Y']].copy()
    if df_hist.empty:
        plot_df['狀態'] = '未完成'
        plot_df['標籤'] = plot_df['樁號']
        return plot_df
    
    hist = df_hist.copy()
    
    def label_maker(r):
        m = str(r.get('機台', 'A'))[0]
        s = r.get('施作順序', 0)
        return f"{r['樁號']}({m}{int(s)})"
    
    hist['標籤'] = hist.apply(label_maker, axis=1)
    hist['施工日期_DT'] = pd.to_datetime(hist['施工日期'], errors='coerce')
    
    max_date = hist['施工日期_DT'].max()
    if pd.notna(max_date):
        monday = max_date - pd.Timedelta(days=max_date.weekday())
        def set_status(dt):
            if pd.isna(dt): return '未完成'
            if dt < monday: return '[已完成]'
            return dt.strftime('%Y-%m-%d')
        hist['狀態'] = hist['施工日期_DT'].apply(set_status)
    else:
        hist['狀態'] = '未完成'
        
    plot_df = plot_df.merge(hist[['樁號', '狀態', '標籤']], on='樁號', how='left')
    plot_df['狀態'] = plot_df['狀態'].fillna('未完成')
    plot_df['標籤'] = plot_df['標籤'].fillna(plot_df['樁號'])
    return plot_df

def sync_to_chart_sheet():
    ss_now, m_now, c_now = get_gs_connection()
    if not ss_now or df_history.empty: return
    
    try:
        plot_df = process_status_logic(df_history, df_base)
        
        gs_matrix = pd.DataFrame()
        gs_matrix['X'] = plot_df['X']
        gs_matrix['標籤'] = plot_df['標籤']
        
        gs_matrix['未完成'] = plot_df['Y'].where(plot_df['狀態'] == '未完成', None)
        gs_matrix['[已完成]'] = plot_df['Y'].where(plot_df['狀態'] == '[已完成]', None)
        
        valid_dates = sorted([s for s in plot_df['狀態'].unique() if s not in ['未完成', '[已完成]']])
        for d in valid_dates:
            gs_matrix[d] = plot_df['Y'].where(plot_df['狀態'] == d, None)

        gs_matrix = gs_matrix.astype(object).where(pd.notnull(gs_matrix), None)
        out_data = [gs_matrix.columns.values.tolist()] + gs_matrix.values.tolist()

        c_now.clear()
        c_now.update("A1", out_data)
        st.success("✅ 雲端繪圖數據已同步")
    except Exception as e:
        st.error(f"同步失敗: {e}")

st.sidebar.header("📂 備份與同步")
if st.sidebar.button("🔄 手動同步雲端數據"):
    sync_to_chart_sheet()

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
            sh_main.append_rows(new_rows)
            st.sidebar.success(f"已同步 {len(new_rows)} 筆")
            sync_to_chart_sheet()
            st.rerun()
    except Exception as e: st.sidebar.error(f"還原失敗: {e}")

st.markdown("### 📝 進度登錄")
c1, c2, c3 = st.columns([1, 1, 2])
work_date = str(c1.date_input("日期"))
machine = c2.radio("機台", ["A車", "B車"], horizontal=True)
mode = c3.radio("模式", ["4支一循環", "2支一循環"], horizontal=True)
step = 4 if "4支" in mode else 2

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
        sync_to_chart_sheet()
        st.rerun()

t1, t2 = st.tabs(["🎯 推算", "✏️ 手動"])
with t1:
    with st.form("a"):
        cc1, cc2, cc3 = st.columns(3); sp = cc1.number_input("起始 P", 1, 613, 1)
        dr = cc2.radio("方向", ["遞增", "遞減"]); ct = cc3.number_input("數量", 1, 100, 10)
        if st.form_submit_button("登錄"):
            plist = []; cur = sp
            for _ in range(ct):
                if 1 <= cur <= 613: plist.append(f"P{cur}")
                cur = cur + step if dr == "遞增" else cur - step
            save_data(plist)
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
            save_data(plist)

st.markdown("---")
st.subheader("🗺️ 現場施工全區圖 (左鍵平移 / 滾輪縮放)")
df_p = process_status_logic(df_history, df_base)

color_map = {'未完成': '#696969', '[已完成]': '#FFB6C1'}
colors_seq = px.colors.qualitative.Plotly

fig = px.scatter(
    df_p, x='X', y='Y', text='標籤', color='狀態',
    color_discrete_map=color_map,
    color_discrete_sequence=colors_seq
)
fig.update_traces(
    textposition='top center', 
    textfont=dict(size=8),
    marker=dict(size=10, line=dict(width=1, color='white'))
)
fig.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=950, plot_bgcolor='white', dragmode='pan')
st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

if not df_history.empty:
    def xl_gen(h_df, p_df):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
            h_df.to_excel(wr, sheet_name='施工明細', index=False)
            wb = wr.book; ws = wb.add_worksheet('全區進度圖'); ch = wb.add_chart({'type': 'scatter'})
            col = 10
            
            states = ['未完成', '[已完成]'] + sorted([s for s in p_df['狀態'].unique() if s not in ['未完成', '[已完成]']])
            colors = {'未完成': '#696969', '[已完成]': '#FFB6C1'}
            fallback_colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2']
            color_idx = 0
            
            for state in states:
                sub_df = p_df[p_df['狀態'] == state].reset_index(drop=True)
                if sub_df.empty: continue
                
                sub_df[['X', 'Y', '標籤']].to_excel(wr, sheet_name='全區進度圖', startcol=col, index=False)
                
                marker_color = colors.get(state)
                if not marker_color:
                    marker_color = fallback_colors[color_idx % len(fallback_colors)]
                    color_idx += 1
                
                series_data = {
                    'name': state,
                    'categories': ['全區進度圖', 1, col, len(sub_df), col],
                    'values': ['全區進度圖', 1, col+1, len(sub_df), col+1],
                    'marker': {'type': 'circle', 'size': 6, 'fill': {'color': marker_color}, 'border': {'color': marker_color}}
                }
                
                if state != '未完成':
                    clbls = [{'value': f'=全區進度圖!${xlsxwriter.utility.xl_col_to_name(col+2)}${ri+2}'} for ri in range(len(sub_df))]
                    series_data['data_labels'] = {'custom': clbls, 'position': 'above', 'font': {'size': 8}}
                    
                ch.add_series(series_data)
                col += 4
                
            ch.set_title({'name': '全區進度圖'}); ch.set_size({'width': 2400, 'height': 1500}); ws.insert_chart('B2', ch)
        return out.getvalue()
    st.sidebar.download_button("📥 匯出 Excel 總報表", xl_gen(df_history, df_p), f"Report_{datetime.date.today()}.xlsx", type="primary")
