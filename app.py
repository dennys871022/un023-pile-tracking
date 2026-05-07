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

st.set_page_config(page_title="UN023 排樁進度系統 V31", layout="wide")
st.title("🏗️ UN023 排樁進度管理 (PDF預覽與排版微調版)")

# === 字體設定 ===
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

# === 資料庫邏輯 ===
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

# === 動態統計邏輯 ===
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

# === UI 介面 ===
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

# === 網頁預覽圖 (維持穩定空心/實心大圓) ===
color_map = {'未完成': '#696969', '[已完成]': '#FFB6C1'}
fig_web = px.scatter(
    df_p, x='X', y='Y', text='標籤', color='狀態',
    color_discrete_map=color_map, color_discrete_sequence=px.colors.qualitative.Plotly,
    custom_data=['樁號']
)
fig_web.update_traces(
    selector=dict(name='未完成'),
    marker=dict(symbol='circle-open', size=16, line=dict(width=2, color='#A9A9A9')),
    textposition='top right'
)
fig_web.update_traces(
    selector=lambda t: t.name != '未完成',
    marker=dict(symbol='circle', size=16, line=dict(width=1, color='white')),
    textposition='top right'
)
fig_web.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=900, plot_bgcolor='white', dragmode='pan')

st.subheader("🗺️ 網頁選取區 (框選範圍即可針對該區生成PDF)")
try:
    selection_event = st.plotly_chart(fig_web, use_container_width=True, config={'scrollZoom': True}, on_select="rerun", selection_mode=('box', 'lasso'))
    selected_piles = [pt["customdata"][0] for pt in selection_event["selection"]["points"]] if selection_event and "selection" in selection_event and selection_event["selection"]["points"] else []
except:
    st.plotly_chart(fig_web, use_container_width=True, config={'scrollZoom': True})
    selected_piles = []

# === PDF 報表自訂與版面位置調整 ===
if not df_history.empty:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📄 PDF 報表文字內容")
    pdf_loc_note = st.sidebar.text_input("右上角位置標題", "滯洪池BC")
    pdf_week_est = st.sidebar.number_input("本週預計完成 (支)", value=36)
    pdf_today_done = st.sidebar.number_input("本日完成 (支) [自動統計]", value=today_done_auto)
    pdf_cum_done = st.sidebar.number_input("累積完成 (支) [自動統計]", value=total_done_auto)
    
    # 全新功能：版面微調滑桿 (放在 Form 內避免滑動時不斷卡頓)
    st.sidebar.markdown("### 📐 PDF 版面位置微調")
    with st.sidebar.form("layout_controls"):
        st.caption("調整 X, Y 數值 (0.00 ~ 1.00) 可移動文字在 PDF 上的位置。")
        pos_title_y = st.slider("大標題 高度 (Y)", 0.0, 1.0, 0.90, 0.01)
        pos_info_x = st.slider("統計資訊 左右 (X)", 0.0, 1.0, 0.05, 0.01)
        pos_info_y = st.slider("統計資訊 高度 (Y)", 0.0, 1.0, 0.85, 0.01)
        pos_loc_x = st.slider("右上角標題 左右 (X)", 0.0, 1.0, 0.70, 0.01)
        pos_loc_y = st.slider("右上角標題 高度 (Y)", 0.0, 1.0, 0.95, 0.01)
        pos_leg_x = st.slider("圖例 左右 (X)", 0.0, 1.5, 1.15, 0.01)
        pos_leg_y = st.slider("圖例 高度 (Y)", 0.0, 1.5, 1.05, 0.01)
        st.form_submit_button("🔄 套用排版並更新下方預覽圖")
    
    st.sidebar.markdown("### 📥 下載區")
    
    # 穩定版 Excel
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
                if state == '未完成':
                    series_data['marker'] = {'type': 'circle', 'size': 6, 'fill': {'none': True}, 'border': {'color': marker_color}}
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

    if MATPLOTLIB_READY:
        is_local_mode = bool(selected_piles)
        pdf_target_df = df_p[df_p['樁號'].isin(selected_piles)].copy() if is_local_mode else df_p.copy()
        
        # 建立 PDF 圖形物件 (供網頁預覽及下載使用)
        def create_pdf_figure():
            font_name = setup_chinese_font()
            if font_name: plt.rcParams['font.family'] = font_name
            plt.rcParams['axes.unicode_minus'] = False
            
            fig = plt.figure(figsize=(24, 16))
            # 回退至 V28 穩定配置：圖表佔據右側，左下角留白給您貼 Excel
            ax = fig.add_axes([0.45, 0.1, 0.5, 0.75]) 
            
            states = ['未完成', '[已完成]'] + sorted([s for s in pdf_target_df['狀態'].unique() if s not in ['未完成', '[已完成]']])
            colors = {'未完成': '#808080', '[已完成]': '#FFB6C1'}
            fallback_colors = px.colors.qualitative.Plotly
            color_idx = 0
            
            lbl_fontsize = 18 if is_local_mode else 9
            
            texts = []
            for state in states:
                sub_df = pdf_target_df[pdf_target_df['狀態'] == state]
                if sub_df.empty: continue
                c = colors.get(state, fallback_colors[color_idx % len(fallback_colors)])
                if state not in colors: color_idx += 1
                
                if state == '未完成':
                    ax.scatter(sub_df['X'], sub_df['Y'], facecolors='none', edgecolors=c, s=180, lw=1.5, zorder=2, label=state)
                else:
                    # 依據範例圖例格式：日期 樁號 ○ 施作順序 (此處先標日期供圖例使用，文字標籤會在下方加入)
                    legend_label = f"{state} 樁號 ○ 施作順序"
                    ax.scatter(sub_df['X'], sub_df['Y'], color=c, s=180, zorder=3, label=legend_label)
                    for _, row in sub_df.iterrows():
                        texts.append(ax.text(row['X'], row['Y'], row['標籤'], fontsize=lbl_fontsize, ha='center', va='center'))

            ax.margins(0.1)
            
            adjust_text(texts, ax=ax, 
                        expand_points=(2.5, 2.5), 
                        expand_text=(1.5, 1.5),
                        arrowprops=dict(arrowstyle='-', color='black', lw=1.0, alpha=0.8),
                        max_iterations=800)
            
            ax.set_aspect('equal', adjustable='datalim')
            ax.axis('off')
            
            # 使用側邊欄設定的 X, Y 座標
            ax.legend(loc='upper right', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28, markerscale=1.5)
            
            roc_year = datetime.date.today().year - 1911
            today_str = f"{roc_year}/{datetime.date.today().month:02d}/{datetime.date.today().day:02d}"
            
            latest_dt = pd.to_datetime(df_history['施工日期'], errors='coerce').max()
            if pd.isna(latest_dt):
                latest_dt = datetime.date.today()
            sunday = latest_dt + datetime.timedelta(days=(6 - latest_dt.weekday()))
            roc_sun_y = sunday.year - 1911
            week_range = f"{week_start_str}~{roc_sun_y}/{sunday.month:02d}/{sunday.day:02d}"
            
            # 使用側邊欄設定的 X, Y 座標放置文字
            fig.text(0.05, pos_title_y, f"{today_str} 施作進度回報", fontsize=50, fontweight='bold')
            
            info_lines = [
                f"本週預計完成 {pdf_week_est} 支",
                f"{week_range}",
                f"本日完成 {pdf_today_done} 支",
                f"{today_str}",
                f"累積完成 {pdf_cum_done} 支"
            ]
            fig.text(pos_info_x, pos_info_y, "\n".join(info_lines), fontsize=35, linespacing=1.6, va='top')
            
            fig.text(pos_loc_x, pos_loc_y, pdf_loc_note, fontsize=55, fontweight='bold', ha='center')
            
            return fig

        # 產生供展示及下載的 Figure
        pdf_fig = create_pdf_figure()
        
        # --- 全新功能：網頁內即時預覽 PDF 實際輸出長相 ---
        st.markdown("---")
        st.subheader("👁️ PDF 最終版面預覽區 (可從左側調整文字位置)")
        st.info("💡 下方顯示的就是您即將下載的 PDF 報表排版。您可以透過左側側邊欄的滑桿自由移動標題、圖例與文字區塊的位置。")
        st.pyplot(pdf_fig)
        
        # 將 Figure 轉存為 Bytes 供下載按鈕使用
        buf = io.BytesIO()
        pdf_fig.savefig(buf, format='pdf', bbox_inches='tight')
        plt.close(pdf_fig) # 防止記憶體溢出
        pdf_bytes = buf.getvalue()

        pdf_btn_text = "🔴 匯出 PDF (您框選的局部範圍)" if is_local_mode else "🔴 匯出 PDF (全區圖)"
        st.sidebar.download_button(pdf_btn_text, pdf_bytes, f"Plan_{datetime.date.today()}.pdf", type="primary")
