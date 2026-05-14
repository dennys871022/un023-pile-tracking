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
    MATPLOTLIB_READY = True
except ImportError:
    MATPLOTLIB_READY = False

st.set_page_config(page_title="UN023 排樁進度系統 V50", layout="wide")
st.title("🏗️ UN023 排樁進度管理 (文字細節精準版)")

# 初始化 Session State
if 'sel_a' not in st.session_state:
    st.session_state.sel_a = []
if 'sel_b' not in st.session_state:
    st.session_state.sel_b = []

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
        sh_main = ss.worksheet("施工明細")
        return ss, sh_main
    except Exception as e:
        st.error(f"雲端連線異常: {e}")
        return None, None

def load_settings(ss):
    default_settings = {
        "pdf_loc_note_right": "滯洪池B.C",
        "pdf_loc_note_left": "滯洪池A",
        "fig_scale": 1.5, "marker_size": 180, "lbl_fontsize": 18, "text_offset": 20,
        "pos_title_y": 0.90, "pos_info_x": 0.05, "pos_info_y": 0.85,
        "pos_loc_x": 0.70, "pos_loc_y": 0.95, "pos_loc_x_left": 0.22, "pos_loc_y_left": 0.55,
        "pos_leg_x": 0.00, "pos_leg_y": 0.00
    }
    if ss is None: return default_settings
    try:
        sh = ss.worksheet("系統設定")
        records = sh.get_all_records()
        loaded = {}
        for r in records:
            k = r.get('Key')
            v = r.get('Value')
            if k in default_settings:
                if isinstance(default_settings[k], int):
                    loaded[k] = int(float(v))
                elif isinstance(default_settings[k], float):
                    loaded[k] = float(v)
                else:
                    loaded[k] = str(v)
        return {**default_settings, **loaded}
    except:
        return default_settings

def save_settings(ss, settings_dict):
    if ss is None: return
    try:
        sh = ss.worksheet("系統設定")
        sh.clear()
        out = [['Key', 'Value']]
        for k, v in settings_dict.items(): out.append([k, str(v)])
        sh.append_rows(out)
    except: pass

def fetch_current_data(sh_main):
    if sh_main is None: return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
    try:
        records = sh_main.get_all_records()
        if not records: return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])
        df = pd.DataFrame(records)
        df['樁號'] = df['樁號'].astype(str).str.upper().str.strip()
        return df
    except: return pd.DataFrame(columns=['樁號', '施工日期', '機台', '施作順序', 'X', 'Y'])

ss, sh_main = get_gs_connection()
if 'ui_settings' not in st.session_state:
    st.session_state.ui_settings = load_settings(ss)

s = st.session_state.ui_settings
df_history = fetch_current_data(sh_main)

total_done_auto = len(df_history)
total_perc = (total_done_auto / 613) * 100 if 613 > 0 else 0
today_done_auto_a = 0
today_done_auto_b = 0
cum_done_a = 0
cum_done_b = 0
this_week_done_a = 0
this_week_done_b = 0
week_start_str = ""
today_state_key = ""

if not df_history.empty:
    df_history['施工日期_DT'] = pd.to_datetime(df_history['施工日期'], errors='coerce')
    latest_dt = df_history['施工日期_DT'].max()
    today_data = df_history[df_history['施工日期_DT'] == latest_dt]
    today_done_auto_a = len(today_data[today_data['機台'].astype(str).str.upper().str.contains('A')])
    today_done_auto_b = len(today_data[today_data['機台'].astype(str).str.upper().str.contains('B')])
    
    cum_done_a = len(df_history[df_history['機台'].astype(str).str.upper().str.contains('A')])
    cum_done_b = len(df_history[df_history['機台'].astype(str).str.upper().str.contains('B')])
    
    today_state_key = latest_dt.strftime('%m/%d')
    monday = latest_dt - pd.Timedelta(days=latest_dt.weekday())
    this_week_data = df_history[df_history['施工日期_DT'] >= monday]
    
    if not this_week_data.empty:
        earliest_this_week = this_week_data['施工日期_DT'].min()
        week_start_str = f"{earliest_this_week.year-1911}/{earliest_this_week.month:02d}/{earliest_this_week.day:02d}"
        this_week_done_a = len(this_week_data[this_week_data['機台'].astype(str).str.upper().str.contains('A')])
        this_week_done_b = len(this_week_data[this_week_data['機台'].astype(str).str.upper().str.contains('B')])
    else:
        week_start_str = f"{monday.year-1911}/{monday.month:02d}/{monday.day:02d}"

def process_status_logic(df_hist, df_b):
    plot_df = df_b[['樁號', 'X', 'Y', '數字']].copy().sort_values('數字').reset_index(drop=True)
    dx = plot_df['X'].diff().bfill(); dy = plot_df['Y'].diff().bfill()
    dx_fwd = (plot_df['X'].shift(-1) - plot_df['X']).ffill(); dy_fwd = (plot_df['Y'].shift(-1) - plot_df['Y']).ffill()
    plot_df['is_horizontal'] = (dx + dx_fwd).abs() >= (dy + dy_fwd).abs()
    
    if df_hist.empty:
        plot_df['狀態'] = '未完成'; plot_df['標籤'] = plot_df['樁號']; plot_df['純順序'] = ""
        return plot_df
    
    hist = df_hist.copy()
    hist['標籤'] = hist.apply(lambda r: f"{r['樁號']}({str(r.get('機台','A'))[0]}{int(r.get('施作順序',0))})", axis=1)
    hist['純順序'] = hist.apply(lambda r: f"({str(r.get('機台','A'))[0]}{int(r.get('施作順序',0))})", axis=1)
    hist['施工日期_DT'] = pd.to_datetime(hist['施工日期'], errors='coerce')
    max_date = hist['施工日期_DT'].max(); monday_dt = max_date - pd.Timedelta(days=max_date.weekday())
    hist['狀態'] = hist['施工日期_DT'].apply(lambda dt: '未完成' if pd.isna(dt) else ('[已完成]' if dt < monday_dt else dt.strftime('%m/%d')))
    
    plot_df = plot_df.merge(hist[['樁號', '狀態', '標籤', '純順序']], on='樁號', how='left')
    plot_df['狀態'] = plot_df['狀態'].fillna('未完成'); plot_df['標籤'] = plot_df['標籤'].fillna(plot_df['樁號']); plot_df['純順序'] = plot_df['純順序'].fillna("")
    return plot_df

df_p = process_status_logic(df_history, df_base)

def get_local_stats(sel_list, p_df):
    if not sel_list: return 0, 0
    sub = p_df[p_df['樁號'].isin(sel_list)]
    total = len(sub)
    done = len(sub[sub['狀態'] != '未完成'])
    return done, total

local_a_done, local_a_total = get_local_stats(st.session_state.sel_a, df_p)
local_b_done, local_b_total = get_local_stats(st.session_state.sel_b, df_p)

st.markdown("### 📝 進度登錄")
c1, c2, c3 = st.columns([1, 1, 2])
work_date = c1.date_input("日期"); machine = c2.radio("機台", ["A車", "B車"], horizontal=True)
mode = c3.radio("模式", ["4支一循環", "2支一循環"], horizontal=True); step = 4 if "4支" in mode else 2

def save_data(piles):
    if not piles or sh_main is None: return
    m_data = df_history[df_history['機台'] == machine]
    seq = 0 if m_data.empty else pd.to_numeric(m_data['施作順序'], errors='coerce').max()
    new_d = []
    for p in piles:
        p = p.upper().strip()
        if p not in df_history['樁號'].values:
            seq += 1; b = df_base[df_base['樁號'] == p]
            x, y = (b['X'].iloc[0], b['Y'].iloc[0]) if not b.empty else (0, 0)
            new_d.append([p, str(work_date), machine, int(seq), float(x), float(y)])
    if new_d: sh_main.append_rows(new_d); st.rerun()

t1, t2 = st.tabs(["🎯 推算", "✏️ 手動"])
with t1:
    with st.form("a"):
        cc1, cc2, cc3 = st.columns(3); sp = cc1.number_input("起始 P", 1, 613, 1)
        dr = cc2.radio("方向", ["遞增", "遞減"]); ct = cc3.number_input("數量", 1, 100, 10)
        if st.form_submit_button("執行登錄"):
            plist = []; cur = sp
            for _ in range(int(ct)):
                if 1 <= cur <= 613: plist.append(f"P{cur}")
                cur = cur + step if dr == "遞增" else cur - step
            save_data(plist)
with t2:
    with st.form("m"):
        raw = st.text_input("區間 (1-50)"); 
        if st.form_submit_button("執行登錄"):
            plist = []
            if raw:
                pts = re.split(r'[,\s]+', re.sub(r'[pP]', '', raw))
                for pt in pts:
                    if '-' in pt:
                        s_idx, e_idx = map(int, pt.split('-')); rs = step if s_idx <= e_idx else -step
                        for n in range(s_idx, e_idx + (1 if s_idx <= e_idx else -1), rs): plist.append(f"P{n}")
                    elif pt.isdigit(): plist.append(f"P{pt}")
            save_data(plist)

st.markdown("---")
fig_web = px.scatter(df_p, x='X', y='Y', text='標籤', color='狀態', color_discrete_map={'未完成': '#696969', '[已完成]': '#FFB6C1'}, custom_data=['樁號'])
fig_web.update_traces(selector=dict(name='未完成'), marker=dict(symbol='circle-open', size=16, line=dict(width=2, color='#A9A9A9')), textposition='top right')
fig_web.update_traces(selector=lambda t: t.name != '未完成', marker=dict(symbol='circle', size=16, line=dict(width=1, color='white')), textposition='top right')
fig_web.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=900, plot_bgcolor='white', dragmode='pan')

st.subheader("🗺️ 網頁選取區 (框選後可指定局部範圍)")
try:
    selection_event = st.plotly_chart(fig_web, use_container_width=True, config={'scrollZoom': True}, on_select="rerun", selection_mode=('box', 'lasso'))
    selected_piles = [pt["customdata"][0] for pt in selection_event["selection"]["points"]] if selection_event and "selection" in selection_event and selection_event["selection"]["points"] else []
except: selected_piles = []

if selected_piles:
    st.success(f"🎯 畫面上目前已選取： **{len(selected_piles)}** 支樁位 (確認數量後請點擊下方按鈕分配)")
else:
    st.caption("💡 提示：請在地圖上方框選樁位，此處會即時顯示選取數量。")

col_btn1, col_btn2, col_btn3 = st.columns(3)
with col_btn1:
    if st.button("📌 設定為 A機範圍"): st.session_state.sel_a = selected_piles; st.rerun()
with col_btn2:
    if st.button("📌 設定為 B機範圍"): st.session_state.sel_b = selected_piles; st.rerun()
with col_btn3:
    if st.button("🗑️ 清除選取"): st.session_state.sel_a = []; st.session_state.sel_b = []; st.rerun()

st.info(f"當前暫存狀態：A機 {len(st.session_state.sel_a)} 支樁 | B機 {len(st.session_state.sel_b)} 支樁")

if not df_history.empty:
    st.sidebar.markdown("### 📄 PDF 報表文字內容")
    pdf_loc_note_right = st.sidebar.text_input("右側主標題", s['pdf_loc_note_right'])
    pdf_loc_note_left = st.sidebar.text_input("左側副標題", s['pdf_loc_note_left'])
    pdf_week_est = st.sidebar.number_input("本週預計完成 (支)", value=36)
    
    st.sidebar.markdown("### 🎛️ PDF 圖表幾何微調")
    with st.sidebar.form("geom"):
        fig_scale = st.slider("排樁間距拉開倍率", 1.0, 5.0, s['fig_scale'], 0.1)
        marker_size = st.slider("圓圈大小", 50, 400, s['marker_size'], 10)
        lbl_fontsize = st.slider("樁號文字大小", 8, 40, s['lbl_fontsize'], 1)
        text_offset = st.slider("文字離圓圈距離", 5, 60, s['text_offset'], 1)
        st.form_submit_button("🔄 套用幾何設定")

    st.sidebar.markdown("### 📐 PDF 文字位置微調")
    with st.sidebar.form("layout"):
        pos_title_y = st.slider("大標題高度 (Y)", 0.0, 1.0, s['pos_title_y'], 0.01)
        pos_info_x = st.slider("資訊區左右 (X)", 0.0, 1.0, s['pos_info_x'], 0.01)
        pos_info_y = st.slider("資訊區高度 (Y)", 0.0, 1.0, s['pos_info_y'], 0.01)
        pos_loc_x = st.slider("右側標題 (X)", 0.0, 1.0, s['pos_loc_x'], 0.01)
        pos_loc_y = st.slider("右側標題 (Y)", 0.0, 1.0, s['pos_loc_y'], 0.01)
        pos_loc_x_left = st.slider("左側標題 (X)", 0.0, 1.0, s['pos_loc_x_left'], 0.01)
        pos_loc_y_left = st.slider("左側標題 (Y)", 0.0, 1.0, s['pos_loc_y_left'], 0.01)
        pos_leg_x = st.slider("圖例左右 (X)", -1.0, 1.5, s['pos_leg_x'], 0.01)
        pos_leg_y = st.slider("圖例高度 (Y)", -1.0, 1.5, s['pos_leg_y'], 0.01)
        st.form_submit_button("🔄 套用位置設定")

    if st.sidebar.button("💾 記憶當前排版與標題 (永久儲存)"):
        new_s = {"pdf_loc_note_right": pdf_loc_note_right, "pdf_loc_note_left": pdf_loc_note_left, "fig_scale": fig_scale, "marker_size": marker_size, "lbl_fontsize": lbl_fontsize, "text_offset": text_offset, "pos_title_y": pos_title_y, "pos_info_x": pos_info_x, "pos_info_y": pos_info_y, "pos_loc_x": pos_loc_x, "pos_loc_y": pos_loc_y, "pos_loc_x_left": pos_loc_x_left, "pos_loc_y_left": pos_loc_y_left, "pos_leg_x": pos_leg_x, "pos_leg_y": pos_leg_y}
        save_settings(ss, new_s); st.session_state.ui_settings = new_s; st.sidebar.success("✅ 設定已記憶")

    if MATPLOTLIB_READY:
        def draw_pdf_axis(ax, target_df, scale_factor=1.0, is_main=False):
            if target_df.empty: ax.axis('off'); return
            states = ['未完成', '[已完成]'] + sorted([s for s in target_df['狀態'].unique() if s not in ['未完成', '[已完成]']])
            colors = {'未完成': '#808080', '[已完成]': '#FFB6C1'}; color_idx = 0; pal = px.colors.qualitative.Plotly
            for state in states:
                sub = target_df[target_df['狀態'] == state]
                if sub.empty: continue
                c = colors.get(state, pal[color_idx % len(pal)]); 
                if state not in colors: color_idx += 1
                if state == '未完成': ax.scatter(sub['X'], sub['Y'], facecolors='none', edgecolors=c, s=marker_size*scale_factor, lw=1.5, zorder=2)
                else:
                    ax.scatter(sub['X'], sub['Y'], color=c, s=marker_size*scale_factor, zorder=3, label=f"{state} 樁號 ○ 施作順序" if is_main else None)
                    if state == today_state_key:
                        for _, row in sub.iterrows():
                            is_h = row['is_horizontal']; p = row['樁號']; s_txt = row['純順序']; off = text_offset*scale_factor; fs = lbl_fontsize*scale_factor
                            if is_h: 
                                ax.annotate(p, (row['X'], row['Y']), xytext=(0, off), textcoords='offset points', fontsize=fs, fontweight='bold', ha='center', va='bottom', zorder=4)
                                if s_txt: ax.annotate(s_txt, (row['X'], row['Y']), xytext=(0, -off), textcoords='offset points', fontsize=fs, color=c, ha='center', va='top', zorder=4)
                            else:
                                ax.annotate(p, (row['X'], row['Y']), xytext=(-off, 0), textcoords='offset points', fontsize=fs, fontweight='bold', ha='right', va='center', zorder=4)
                                if s_txt: ax.annotate(s_txt, (row['X'], row['Y']), xytext=(off, 0), textcoords='offset points', fontsize=fs, color=c, ha='left', va='center', zorder=4)
            ax.margins(0.1); ax.set_aspect('equal', adjustable='datalim'); ax.axis('off')

        def create_pdf_figure():
            font_name = setup_chinese_font(); 
            if font_name: plt.rcParams['font.family'] = font_name
            fig = plt.figure(figsize=(24 * fig_scale, 16 * fig_scale))
            has_a, has_b = bool(st.session_state.sel_a), bool(st.session_state.sel_b)
            
            if not (has_a or has_b):
                ax = fig.add_axes([0.45, 0.1, 0.5, 0.75]); draw_pdf_axis(ax, df_p, 1.0, True)
                ax.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28 * fig_scale, markerscale=1.5)
            else:
                if has_a and has_b:
                    ax_a = fig.add_axes([0.35, 0.1, 0.30, 0.75]); draw_pdf_axis(ax_a, df_p[df_p['樁號'].isin(st.session_state.sel_a)], 1.0, True); ax_a.set_title("A機作業區", fontsize=40*fig_scale, fontweight='bold', y=-0.05); ax_a.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28*fig_scale, markerscale=1.5)
                    ax_b = fig.add_axes([0.68, 0.1, 0.30, 0.75]); draw_pdf_axis(ax_b, df_p[df_p['樁號'].isin(st.session_state.sel_b)], 1.0, False); ax_b.set_title("B機作業區", fontsize=40*fig_scale, fontweight='bold', y=-0.05)
                elif has_a:
                    ax_a = fig.add_axes([0.45, 0.1, 0.5, 0.75]); draw_pdf_axis(ax_a, df_p[df_p['樁號'].isin(st.session_state.sel_a)], 1.0, True); ax_a.set_title("A機作業區", fontsize=40*fig_scale, fontweight='bold', y=-0.05); ax_a.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28*fig_scale, markerscale=1.5)
                elif has_b:
                    ax_b = fig.add_axes([0.45, 0.1, 0.5, 0.75]); draw_pdf_axis(ax_b, df_p[df_p['樁號'].isin(st.session_state.sel_b)], 1.0, True); ax_b.set_title("B機作業區", fontsize=40*fig_scale, fontweight='bold', y=-0.05); ax_b.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28*fig_scale, markerscale=1.5)

            roc_y = datetime.date.today().year - 1911; today_roc = f"{roc_y}/{datetime.date.today().month:02d}/{datetime.date.today().day:02d}"
            
            latest_dt = pd.to_datetime(df_history['施工日期'], errors='coerce').max()
            if pd.isna(latest_dt): latest_dt = datetime.date.today()
            sunday = latest_dt + datetime.timedelta(days=(6 - latest_dt.weekday()))
            week_end_str = f"{sunday.year-1911}/{sunday.month:02d}/{sunday.day:02d}"
            
            a_pct_str = f" ({(local_a_done/local_a_total)*100:.2f}%)" if local_a_total > 0 else ""
            b_pct_str = f" ({(local_b_done/local_b_total)*100:.2f}%)" if local_b_total > 0 else ""
            
            info_lines = [
                f"本週預計完成 {pdf_week_est} 支",
                f"{week_start_str}~{week_end_str}",
                f"本週累積 A機:{this_week_done_a}支 B機:{this_week_done_b}支",
                f"本日完成 A機:{today_done_auto_a}支 B機:{today_done_auto_b}支",
                f"選取區 A機:{local_a_done}/{local_a_total}{a_pct_str}",
                f"　　　 B機:{local_b_done}/{local_b_total}{b_pct_str}",
                f"總累積完成 {total_done_auto} 支 ({total_done_auto}/613, {total_perc:.2f}%)",
                f"各別累積 A機:{cum_done_a}支 B機:{cum_done_b}支"
            ]
            fig.text(0.05, pos_title_y, f"{today_roc} 施作進度回報", fontsize=50 * fig_scale, fontweight='bold')
            fig.text(pos_info_x, pos_info_y, "\n".join(info_lines), fontsize=35 * fig_scale, linespacing=1.6, va='top')
            fig.text(pos_loc_x, pos_loc_y, pdf_loc_note_right, fontsize=55 * fig_scale, fontweight='bold', ha='center')
            fig.text(pos_loc_x_left, pos_loc_y_left, pdf_loc_note_left, fontsize=55 * fig_scale, fontweight='bold', ha='center')
            return fig

        pdf_fig = create_pdf_figure(); st.markdown("---"); st.pyplot(pdf_fig)
        buf = io.BytesIO(); pdf_fig.savefig(buf, format='pdf', bbox_inches='tight'); plt.close(pdf_fig)
        st.sidebar.download_button("🔴 匯出 PDF 報表", buf.getvalue(), f"Plan_{datetime.date.today()}.pdf", type="primary")

    def xl_gen(h_df, p_df):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
            h_df.to_excel(wr, sheet_name='施工明細', index=False); wb = wr.book; ws = wb.add_worksheet('全區進度圖'); ch = wb.add_chart({'type': 'scatter'})
            col = 10; states = ['未完成', '[已完成]'] + sorted([s for s in p_df['狀態'].unique() if s not in ['未完成', '[已完成]']])
            colors = {'未完成': '#696969', '[已完成]': '#FFB6C1'}; pal = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2']; ci = 0
            for s in states:
                sub = p_df[p_df['狀態'] == s].reset_index(drop=True); 
                if sub.empty: continue
                sub[['X', 'Y', '標籤']].to_excel(wr, sheet_name='全區進度圖', startcol=col, index=False)
                mc = colors.get(s, pal[ci % len(pal)]); 
                if s not in colors: ci += 1
                sd = {'name': s, 'categories': ['全區進度圖', 1, col, len(sub), col], 'values': ['全區進度圖', 1, col+1, len(sub), col+1], 'marker': {'type': 'circle', 'size': 6, 'fill': {'color': mc}, 'border': {'color': mc}}}
                if s == '未完成': sd['marker']['fill'] = {'none': True}
                if s != '未完成': sd['data_labels'] = {'custom': [{'value': f'=全區進度圖!${xlsxwriter.utility.xl_col_to_name(col+2)}${ri+2}'} for ri in range(len(sub))], 'position': 'above', 'font': {'size': 8}}
                ch.add_series(sd); col += 4
            ch.set_x_axis({'visible': False}); ch.set_y_axis({'visible': False}); ws.insert_chart('B2', ch)
        return out.getvalue()
    st.sidebar.download_button("🟢 匯出 Excel (全區報表)", xl_gen(df_history, df_p), f"Report_{datetime.date.today()}.xlsx")
