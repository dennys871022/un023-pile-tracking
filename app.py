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

try:
    import matplotlib.pyplot as plt
    from adjustText import adjust_text
    MATPLOTLIB_READY = True
except ImportError:
    MATPLOTLIB_READY = False

st.set_page_config(page_title="UN023 排樁進度系統 V24", layout="wide")
st.title("🏗️ UN023 排樁進度管理 (完美版面客製版)")

@st.cache_resource
def setup_chinese_font():
    import os
    import urllib.request
    import matplotlib.font_manager as fm
    font_path = 'NotoSansCJKtc-Regular.otf'
    if not os.path.exists(font_path):
        try:
            url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
            urllib.request.urlretrieve(url, font_path)
        except Exception as e:
            pass
    if os.path.exists(font_path):
        fm.fontManager.addfont(font_path)
        return fm.FontProperties(fname=font_path).get_name()
    return None

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
        df['施作順序'] = pd.to_numeric(df.get('施作順序', 0), errors='coerce').fillna(0)
        return df
    except:
        return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])

ss, sh_main, sh_chart = get_gs_connection()
df_history = fetch_current_data(sh_main)

# --- 核心邏輯：動態統計 ---
total_done_auto = len(df_history)
today_done_auto = 0
week_start_str = ""
if not df_history.empty:
    df_history['施工日期_DT'] = pd.to_datetime(df_history['施工日期'], errors='coerce')
    latest_dt = df_history['施工日期_DT'].max()
    today_done_auto = len(df_history[df_history['施工日期_DT'] == latest_dt])
    
    monday = latest_dt - pd.Timedelta(days=latest_dt.weekday())
    this_week_data = df_history[df_history['施工日期_DT'] >= monday]
    if not this_week_data.empty:
        earliest_this_week = this_week_data['施工日期_DT'].min()
        roc_y = earliest_this_week.year - 1911
        week_start_str = f"{roc_y}/{earliest_this_week.month:02d}/{earliest_this_week.day:02d}"

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

# --- UI 介面 ---
st.sidebar.header("📂 備份與同步")
if st.sidebar.button("🔄 手動同步雲端數據"):
    sync_to_chart_sheet()

st.markdown("### 📝 進度登錄")
c1, c2, c3 = st.columns([1, 1, 2])
work_date = c1.date_input("日期")
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
            new_d.append([p, str(work_date), machine, int(seq), float(x), float(y)])
    if new_d:
        sh_main.append_rows(new_d)
        sync_to_chart_sheet()
        st.rerun()

t1, t2 = st.tabs(["🎯 推算", "✏️ 手動"])
with t1:
    with st.form("a"):
        cc1, cc2, cc3 = st.columns(3); sp = cc1.number_input("起始 P", 1, 613, 1)
        dr = cc2.radio("方向", ["遞增", "遞減"]); ct = cc3.number_input("數量", 1, 100, 10)
        if st.form_submit_button("執行登錄"):
            plist = []; cur = sp
            for _ in range(ct):
                if 1 <= cur <= 613: plist.append(f"P{cur}")
                cur = cur + step if dr == "遞增" else cur - step
            save_data(plist)
with t2:
    with st.form("m"):
        raw = st.text_input("區間 (1-50)")
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

st.markdown("---")
df_p = process_status_logic(df_history, df_base)

# --- 網頁圖表樣式：加大圓圈防重疊 ---
color_map = {'未完成': '#696969', '[已完成]': '#FFB6C1'}
fig = px.scatter(
    df_p, x='X', y='Y', text='標籤', color='狀態',
    color_discrete_map=color_map, color_discrete_sequence=px.colors.qualitative.Plotly,
    custom_data=['樁號']
)
# 設定未完成為大號空心圓
fig.update_traces(
    selector=dict(name='未完成'),
    marker=dict(symbol='circle-open', size=16, line=dict(width=2, color='#A9A9A9')),
    textposition='top right' # 強制避開圓圈中心
)
# 設定已完成與日期為實心圓
fig.update_traces(
    selector=lambda t: t.name != '未完成',
    marker=dict(symbol='circle', size=16, line=dict(width=1, color='white')),
    textposition='top right'
)
fig.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=900, plot_bgcolor='white', dragmode='pan')

try:
    selection_event = st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True}, on_select="rerun", selection_mode=('box', 'lasso'))
    selected_piles = [pt["customdata"][0] for pt in selection_event["selection"]["points"]] if selection_event and "selection" in selection_event and selection_event["selection"]["points"] else []
except:
    st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})
    selected_piles = []

# --- PDF 報表自訂內容 ---
if not df_history.empty:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📄 PDF 報表自訂內容")
    pdf_loc_note = st.sidebar.text_input("圖表右上角大字標題", "滯洪池BC")
    pdf_week_est = st.sidebar.number_input("本週預計完成 (支)", value=36)
    pdf_today_done = st.sidebar.number_input("本日完成 (支) [自動統計]", value=today_done_auto)
    pdf_cum_done = st.sidebar.number_input("累積完成 (支) [自動統計]", value=total_done_auto)
    
    st.sidebar.markdown("### 📥 下載區")
    
    # 完美復原 Excel 生成器 (不破壞任何格式、包含圓圈、隱藏座標軸)
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
                
                # Excel 繪圖核心設定 (完全復原版)
                series_data = {
                    'name': state,
                    'categories': ['全區進度圖', 1, col, len(sub_df), col],
                    'values': ['全區進度圖', 1, col+1, len(sub_df), col+1],
                    'marker': {'type': 'circle', 'size': 6, 'fill': {'color': marker_color}, 'border': {'color': marker_color}}
                }
                
                # Excel 未完成空心圓設定
                if state == '未完成':
                    series_data['marker'] = {'type': 'circle', 'size': 6, 'fill': {'color': 'none'}, 'border': {'color': marker_color}}
                
                if state != '未完成':
                    clbls = [{'value': f'=全區進度圖!${xlsxwriter.utility.xl_col_to_name(col+2)}${ri+2}'} for ri in range(len(sub_df))]
                    series_data['data_labels'] = {'custom': clbls, 'position': 'above', 'font': {'size': 8}}
                    
                ch.add_series(series_data)
                col += 4
            
            today_str = datetime.date.today().strftime('%Y-%m-%d')
            ch.set_title({'name': f'{today_str} 施作進度回報'})
            ch.set_size({'width': 2400, 'height': 1500})
            ch.set_x_axis({'visible': False, 'major_gridlines': {'visible': False}})
            ch.set_y_axis({'visible': False, 'major_gridlines': {'visible': False}})
            ws.insert_chart('B2', ch)
        return out.getvalue()
    
    st.sidebar.download_button("🟢 匯出 Excel (全區報表)", xl_gen(df_history, df_p), f"Report_{datetime.date.today()}.xlsx", type="secondary")

    if not MATPLOTLIB_READY:
        st.sidebar.error("⚠️ 請確保已在 GitHub 加入 matplotlib 與 adjustText")
    else:
        if selected_piles:
            pdf_df = df_p[df_p['樁號'].isin(selected_piles)].copy()
            pdf_btn_text = "🔴 匯出 PDF (您框選的局部範圍)"
        else:
            pdf_df = df_p.copy()
            pdf_btn_text = "🔴 匯出 PDF (全區圖)"

        # PDF 智慧繪圖生成器
        def pdf_gen(p_df, loc_text, w_est, t_done, c_done, w_start):
            font_name = setup_chinese_font()
            if font_name: plt.rcParams['font.family'] = font_name
            plt.rcParams['axes.unicode_minus'] = False
            
            fig = plt.figure(figsize=(24, 16))
            
            # 讓出左半邊給文字，圖表放在右側
            ax = fig.add_axes([0.45, 0.1, 0.5, 0.75]) 
            
            states = ['未完成', '[已完成]'] + sorted([s for s in p_df['狀態'].unique() if s not in ['未完成', '[已完成]']])
            colors = {'未完成': '#808080', '[已完成]': '#FFB6C1'}
            fallback_colors = px.colors.qualitative.Plotly
            color_idx = 0
            
            texts = []
            for state in states:
                sub_df = p_df[p_df['狀態'] == state]
                if sub_df.empty: continue
                c = colors.get(state, fallback_colors[color_idx % len(fallback_colors)])
                if state not in colors: color_idx += 1
                
                # 增大點位(s=180)，呈現相切效果
                if state == '未完成':
                    ax.scatter(sub_df['X'], sub_df['Y'], color=c, s=180, facecolors='none', edgecolors=c, lw=1.5, zorder=2, label=state)
                else:
                    ax.scatter(sub_df['X'], sub_df['Y'], color=c, s=180, zorder=3, label=state)
                    # 放入文字標籤，準備交給防撞演算法處理
                    for _, row in sub_df.iterrows():
                        texts.append(ax.text(row['X'], row['Y'], row['標籤'], fontsize=9, ha='center', va='center'))

            ax.margins(0.1)
            
            # 暴力防撞演算法：強迫文字離開點的中心，並帶上指示線
            adjust_text(texts, ax=ax, 
                        force_points=0.8, expand_points=(1.5, 1.5), 
                        force_text=0.5, expand_text=(1.2, 1.2),
                        arrowprops=dict(arrowstyle='-', color='black', lw=0.8, alpha=0.6))
            
            ax.set_aspect('equal', adjustable='datalim')
            ax.axis('off')
            
            # --- 圖例放回右上角 ---
            ax.legend(loc='upper right', bbox_to_anchor=(1.1, 1.05), fontsize=14, markerscale=0.8)
            
            # --- 版面文字繪製 ---
            roc_year = datetime.date.today().year - 1911
            today_str = f"{roc_year}/{datetime.date.today().month:02d}/{datetime.date.today().day:02d}"
            
            latest_dt = pd.to_datetime(df_history['施工日期'], errors='coerce').max()
            sunday = latest_dt + datetime.timedelta(days=(6 - latest_dt.weekday()))
            roc_sun_y = sunday.year - 1911
            week_range = f"{w_start}~{roc_sun_y}/{sunday.month:02d}/{sunday.day:02d}"
            
            # 左上角：大標題
            fig.text(0.05, 0.90, f"{today_str}施作進度回報", fontsize=50, fontweight='bold')
            
            # 左側縮排：資訊區塊
            info_lines = [
                f"本週預計完成-{w_est}支",
                f"{week_range}",
                f"本日完成-{t_done}支",
                f"{today_str}",
                f"累積完成-{c_done}支"
            ]
            fig.text(0.12, 0.82, "\n".join(info_lines), fontsize=32, linespacing=1.6, va='top')
            
            # 右上角：區域備註 (置中於右側圖表上方)
            fig.text(0.70, 0.90, loc_text, fontsize=55, fontweight='bold', ha='center')
            
            buf = io.BytesIO()
            plt.savefig(buf, format='pdf', bbox_inches='tight')
            plt.close(fig)
            return buf.getvalue()

        st.sidebar.download_button(pdf_btn_text, pdf_gen(pdf_df, pdf_loc_note, pdf_week_est, pdf_today_done, pdf_cum_done, week_start_str), f"Plan_{datetime.date.today()}.pdf", type="primary")
