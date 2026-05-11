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

st.set_page_config(page_title="UN023 排樁進度系統 V44", layout="wide")
st.title("🏗️ UN023 排樁進度管理 (雙機獨立統計版)")

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
    except:
        try:
            sh = ss.add_worksheet("系統設定", 50, 2)
            out = [['Key', 'Value']]
            for k, v in default_settings.items():
                out.append([k, str(v)])
            sh.append_rows(out)
            return default_settings
        except:
            return default_settings
    try:
        records = sh.get_all_records()
        loaded = {}
        for r in records:
            k = r.get('Key')
            v = r.get('Value')
            if k in default_settings:
                try:
                    if isinstance(default_settings[k], float):
                        loaded[k] = float(v)
                    elif isinstance(default_settings[k], int):
                        loaded[k] = int(float(v))
                    else:
                        loaded[k] = str(v)
                except:
                    loaded[k] = default_settings[k]
        return {**default_settings, **loaded}
    except:
        return default_settings

def save_settings(ss, settings_dict):
    if ss is None: return
    try:
        sh = ss.worksheet("系統設定")
        sh.clear()
        out = [['Key', 'Value']]
        for k, v in settings_dict.items():
            out.append([k, str(v)])
        sh.append_rows(out)
    except Exception as e:
        st.error(f"設定儲存失敗: {e}")

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

ss, sh_main = get_gs_connection()

if 'ui_settings' not in st.session_state:
    st.session_state.ui_settings = load_settings(ss)

s = st.session_state.ui_settings

df_history = fetch_current_data(sh_main)

total_done_auto = len(df_history)
today_done_auto_a = 0
today_done_auto_b = 0
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
    
    today_state_key = latest_dt.strftime('%m/%d')
    
    monday = latest_dt - pd.Timedelta(days=latest_dt.weekday())
    this_week_data = df_history[df_history['施工日期_DT'] >= monday]
    if not this_week_data.empty:
        earliest_this_week = this_week_data['施工日期_DT'].min()
        roc_y = earliest_this_week.year - 1911
        week_start_str = f"{roc_y}/{earliest_this_week.month:02d}/{earliest_this_week.day:02d}"
        
        this_week_done_a = len(this_week_data[this_week_data['機台'].astype(str).str.upper().str.contains('A')])
        this_week_done_b = len(this_week_data[this_week_data['機台'].astype(str).str.upper().str.contains('B')])

def process_status_logic(df_hist, df_b):
    plot_df = df_b[['樁號', 'X', 'Y', '數字']].copy()
    plot_df = plot_df.sort_values('數字').reset_index(drop=True)
    
    dx = plot_df['X'].diff().bfill()
    dy = plot_df['Y'].diff().bfill()
    dx_fwd = (plot_df['X'].shift(-1) - plot_df['X']).ffill()
    dy_fwd = (plot_df['Y'].shift(-1) - plot_df['Y']).ffill()
    dx_avg = dx + dx_fwd
    dy_avg = dy + dy_fwd
    plot_df['is_horizontal'] = dx_avg.abs() >= dy_avg.abs()
    
    if df_hist.empty:
        plot_df['狀態'] = '未完成'; plot_df['標籤'] = plot_df['樁號']; plot_df['純順序'] = ""
        return plot_df
        
    hist = df_hist.copy()
    hist['標籤'] = hist.apply(lambda r: f"{r['樁號']}({str(r.get('機台','A'))[0]}{int(r.get('施作順序',0))})", axis=1)
    hist['純順序'] = hist.apply(lambda r: f"({str(r.get('機台','A'))[0]}{int(r.get('施作順序',0))})", axis=1)
    hist['施工日期_DT'] = pd.to_datetime(hist['施工日期'], errors='coerce')
    max_date = hist['施工日期_DT'].max()
    monday = max_date - pd.Timedelta(days=max_date.weekday())
    
    hist['狀態'] = hist['施工日期_DT'].apply(lambda dt: '未完成' if pd.isna(dt) else ('[已完成]' if dt < monday else dt.strftime('%m/%d')))
    
    plot_df = plot_df.merge(hist[['樁號', '狀態', '標籤', '純順序']], on='樁號', how='left')
    plot_df['狀態'] = plot_df['狀態'].fillna('未完成')
    plot_df['標籤'] = plot_df['標籤'].fillna(plot_df['樁號'])
    plot_df['純順序'] = plot_df['純順序'].fillna("")
    return plot_df

df_p = process_status_logic(df_history, df_base)

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
color_map_web = {'未完成': '#696969', '[已完成]': '#FFB6C1'}
fig_web = px.scatter(df_p, x='X', y='Y', text='標籤', color='狀態', color_discrete_map=color_map_web, custom_data=['樁號'])
fig_web.update_traces(selector=dict(name='未完成'), marker=dict(symbol='circle-open', size=16, line=dict(width=2, color='#A9A9A9')), textposition='top right')
fig_web.update_traces(selector=lambda t: t.name != '未完成', marker=dict(symbol='circle', size=16, line=dict(width=1, color='white')), textposition='top right')
fig_web.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=900, plot_bgcolor='white', dragmode='pan')

st.subheader("🗺️ 網頁選取區 (框選後可指定為局部圖)")
try:
    selection_event = st.plotly_chart(fig_web, use_container_width=True, config={'scrollZoom': True}, on_select="rerun", selection_mode=('box', 'lasso'))
    selected_piles = [pt["customdata"][0] for pt in selection_event["selection"]["points"]] if selection_event and "selection" in selection_event and selection_event["selection"]["points"] else []
except:
    st.plotly_chart(fig_web, use_container_width=True, config={'scrollZoom': True})
    selected_piles = []

col_btn1, col_btn2, col_btn3 = st.columns(3)
with col_btn1:
    if st.button("📌 設定為 A機範圍"):
        st.session_state.sel_a = selected_piles
        st.rerun()
with col_btn2:
    if st.button("📌 設定為 B機範圍"):
        st.session_state.sel_b = selected_piles
        st.rerun()
with col_btn3:
    if st.button("🗑️ 清除所有局部圖"):
        st.session_state.sel_a = []
        st.session_state.sel_b = []
        st.rerun()

st.info(f"當前暫存狀態：A機 {len(st.session_state.sel_a)} 支樁 | B機 {len(st.session_state.sel_b)} 支樁")

if not df_history.empty:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📄 PDF 報表文字內容")
    
    pdf_loc_note_right = st.sidebar.text_input("右側主標題", s['pdf_loc_note_right'])
    pdf_loc_note_left = st.sidebar.text_input("左側副標題", s['pdf_loc_note_left'])
    
    pdf_week_est = st.sidebar.number_input("本週預計完成 (支)", value=36)
    pdf_this_week_done_a = st.sidebar.number_input("本週累積 A機 (支) [自動統計]", value=this_week_done_a)
    pdf_this_week_done_b = st.sidebar.number_input("本週累積 B機 (支) [自動統計]", value=this_week_done_b)
    pdf_today_done_a = st.sidebar.number_input("本日完成 A機 (支) [自動統計]", value=today_done_auto_a)
    pdf_today_done_b = st.sidebar.number_input("本日完成 B機 (支) [自動統計]", value=today_done_auto_b)
    pdf_cum_done = st.sidebar.number_input("總累積完成 (支) [自動統計]", value=total_done_auto)
    
    st.sidebar.markdown("### 🎛️ PDF 圖表幾何微調")
    with st.sidebar.form("geom_controls"):
        fig_scale = st.slider("畫布排樁間距拉開倍率", 1.0, 5.0, s['fig_scale'], 0.1)
        marker_size = st.slider("圓圈大小", 50, 400, s['marker_size'], 10)
        lbl_fontsize = st.slider("樁號文字大小", 8, 40, s['lbl_fontsize'], 1)
        text_offset = st.slider("文字離圓圈距離", 5, 60, s['text_offset'], 1)
        st.form_submit_button("🔄 套用幾何設定")

    st.sidebar.markdown("### 📐 PDF 文字位置微調")
    with st.sidebar.form("layout_controls"):
        pos_title_y = st.slider("左上大標題 高度 (Y)", 0.0, 1.0, s['pos_title_y'], 0.01)
        pos_info_x = st.slider("統計資訊 左右 (X)", 0.0, 1.0, s['pos_info_x'], 0.01)
        pos_info_y = st.slider("統計資訊 高度 (Y)", 0.0, 1.0, s['pos_info_y'], 0.01)
        pos_loc_x = st.slider("右側位置標題 (X)", 0.0, 1.0, s['pos_loc_x'], 0.01)
        pos_loc_y = st.slider("右側位置標題 (Y)", 0.0, 1.0, s['pos_loc_y'], 0.01)
        pos_loc_x_left = st.slider("左側位置標題 (X)", 0.0, 1.0, s['pos_loc_x_left'], 0.01)
        pos_loc_y_left = st.slider("左側位置標題 (Y)", 0.0, 1.0, s['pos_loc_y_left'], 0.01)
        pos_leg_x = st.slider("圖例 左右 (X)", -1.0, 1.5, s['pos_leg_x'], 0.01)
        pos_leg_y = st.slider("圖例 高度 (Y)", -1.0, 1.5, s['pos_leg_y'], 0.01)
        st.form_submit_button("🔄 套用文字位置")

    if st.sidebar.button("💾 記憶當前排版與標題 (永久儲存)"):
        new_settings = {
            "pdf_loc_note_right": pdf_loc_note_right,
            "pdf_loc_note_left": pdf_loc_note_left,
            "fig_scale": fig_scale,
            "marker_size": marker_size,
            "lbl_fontsize": lbl_fontsize,
            "text_offset": text_offset,
            "pos_title_y": pos_title_y,
            "pos_info_x": pos_info_x,
            "pos_info_y": pos_info_y,
            "pos_loc_x": pos_loc_x,
            "pos_loc_y": pos_loc_y,
            "pos_loc_x_left": pos_loc_x_left,
            "pos_loc_y_left": pos_loc_y_left,
            "pos_leg_x": pos_leg_x,
            "pos_leg_y": pos_leg_y
        }
        save_settings(ss, new_settings)
        st.session_state.ui_settings = new_settings
        st.sidebar.success("✅ 設定已成功寫入雲端！明天打開也會保持這個版面。")

    if MATPLOTLIB_READY:
        def draw_pdf_axis(ax, target_df, scale_factor=1.0, is_main=False):
            if target_df.empty:
                ax.axis('off')
                return
                
            states = ['未完成', '[已完成]'] + sorted([s for s in target_df['狀態'].unique() if s not in ['未完成', '[已完成]']])
            colors = {'未完成': '#808080', '[已完成]': '#FFB6C1'}
            fallback_colors = px.colors.qualitative.Plotly
            color_idx = 0
            
            msize = marker_size * scale_factor
            fsize = lbl_fontsize * scale_factor
            offset = text_offset * scale_factor
            
            for state in states:
                sub_df = target_df[target_df['狀態'] == state]
                if sub_df.empty: continue
                c = colors.get(state, fallback_colors[color_idx % len(fallback_colors)])
                if state not in colors: color_idx += 1
                
                if state == '未完成':
                    ax.scatter(sub_df['X'], sub_df['Y'], facecolors='none', edgecolors=c, s=msize, lw=1.5, zorder=2)
                else:
                    legend_label = f"{state} 樁號 ○ 施作順序" if is_main else ""
                    ax.scatter(sub_df['X'], sub_df['Y'], color=c, s=msize, zorder=3, label=legend_label if legend_label else None)
                    
                    if state == today_state_key:
                        for _, row in sub_df.iterrows():
                            is_horiz = row['is_horizontal']
                            p_text = row['樁號']
                            s_text = row['純順序']
                            
                            if is_horiz:
                                ax.annotate(p_text, (row['X'], row['Y']), xytext=(0, offset), textcoords='offset points', fontsize=fsize, fontweight='bold', color='black', ha='center', va='bottom', zorder=4)
                                if s_text:
                                    ax.annotate(s_text, (row['X'], row['Y']), xytext=(0, -offset), textcoords='offset points', fontsize=fsize, color=c, ha='center', va='top', zorder=4)
                            else:
                                ax.annotate(p_text, (row['X'], row['Y']), xytext=(-offset, 0), textcoords='offset points', fontsize=fsize, fontweight='bold', color='black', ha='right', va='center', zorder=4)
                                if s_text:
                                    ax.annotate(s_text, (row['X'], row['Y']), xytext=(offset, 0), textcoords='offset points', fontsize=fsize, color=c, ha='left', va='center', zorder=4)

            ax.margins(0.1)
            ax.set_aspect('equal', adjustable='datalim')
            ax.axis('off')

        def create_pdf_figure():
            font_name = setup_chinese_font()
            if font_name: plt.rcParams['font.family'] = font_name
            plt.rcParams['axes.unicode_minus'] = False
            
            fig = plt.figure(figsize=(24 * fig_scale, 16 * fig_scale))
            
            has_local_a = bool(st.session_state.sel_a)
            has_local_b = bool(st.session_state.sel_b)
            has_any_local = has_local_a or has_local_b
            
            if not has_any_local:
                ax_main = fig.add_axes([0.45, 0.1, 0.5, 0.75]) 
                draw_pdf_axis(ax_main, df_p, scale_factor=1.0, is_main=True)
                ax_main.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28 * fig_scale, markerscale=1.5)
            else:
                if has_local_a and has_local_b:
                    ax_a = fig.add_axes([0.35, 0.1, 0.30, 0.75])
                    draw_pdf_axis(ax_a, df_p[df_p['樁號'].isin(st.session_state.sel_a)], scale_factor=1.0, is_main=True)
                    ax_a.set_title("A機作業區", fontsize=40 * fig_scale, fontweight='bold', y=-0.05)
                    ax_a.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28 * fig_scale, markerscale=1.5)
                    
                    ax_b = fig.add_axes([0.68, 0.1, 0.30, 0.75])
                    draw_pdf_axis(ax_b, df_p[df_p['樁號'].isin(st.session_state.sel_b)], scale_factor=1.0, is_main=False)
                    ax_b.set_title("B機作業區", fontsize=40 * fig_scale, fontweight='bold', y=-0.05)
                elif has_local_a:
                    ax_a = fig.add_axes([0.45, 0.1, 0.5, 0.75])
                    draw_pdf_axis(ax_a, df_p[df_p['樁號'].isin(st.session_state.sel_a)], scale_factor=1.0, is_main=True)
                    ax_a.set_title("A機作業區", fontsize=40 * fig_scale, fontweight='bold', y=-0.05)
                    ax_a.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28 * fig_scale, markerscale=1.5)
                elif has_local_b:
                    ax_b = fig.add_axes([0.45, 0.1, 0.5, 0.75])
                    draw_pdf_axis(ax_b, df_p[df_p['樁號'].isin(st.session_state.sel_b)], scale_factor=1.0, is_main=True)
                    ax_b.set_title("B機作業區", fontsize=40 * fig_scale, fontweight='bold', y=-0.05)
                    ax_b.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28 * fig_scale, markerscale=1.5)

            roc_year = datetime.date.today().year - 1911
            today_str = f"{roc_year}/{datetime.date.today().month:02d}/{datetime.date.today().day:02d}"
            latest_dt = pd.to_datetime(df_history['施工日期'], errors='coerce').max()
            if pd.isna(latest_dt): latest_dt = datetime.date.today()
            sunday = latest_dt + datetime.timedelta(days=(6 - latest_dt.weekday()))
            week_range = f"{week_start_str}~{sunday.year-1911}/{sunday.month:02d}/{sunday.day:02d}"
            
            fig.text(0.05, pos_title_y, f"{today_str} 施作進度回報", fontsize=50 * fig_scale, fontweight='bold')
            info_lines = [
                f"本週預計完成 {pdf_week_est} 支",
                f"{week_range}",
                f"本週累積 A機:{pdf_this_week_done_a}支 B機:{pdf_this_week_done_b}支",
                f"本日完成 A機:{pdf_today_done_a}支 B機:{pdf_today_done_b}支",
                f"{today_str}",
                f"累積完成 {pdf_cum_done} 支"
            ]
            fig.text(pos_info_x, pos_info_y, "\n".join(info_lines), fontsize=35 * fig_scale, linespacing=1.6, va='top')
            
            fig.text(pos_loc_x, pos_loc_y, pdf_loc_note_right, fontsize=55 * fig_scale, fontweight='bold', ha='center')
            fig.text(pos_loc_x_left, pos_loc_y_left, pdf_loc_note_left, fontsize=55 * fig_scale, fontweight='bold', ha='center')
            
            return fig

        pdf_fig = create_pdf_figure()
        st.markdown("---")
        st.subheader("👁️ PDF 最終版面預覽區")
        st.pyplot(pdf_fig)
        
        buf = io.BytesIO()
        pdf_fig.savefig(buf, format='pdf', bbox_inches='tight')
        plt.close(pdf_fig)
        pdf_bytes = buf.getvalue()
        
        st.sidebar.markdown("### 📥 下載區")
        has_local_download = bool(st.session_state.sel_a) or bool(st.session_state.sel_b)
        pdf_btn_text = "🔴 匯出 PDF 報表 (局部圖)" if has_local_download else "🔴 匯出 PDF 報表 (全區圖)"
        st.sidebar.download_button(pdf_btn_text, pdf_bytes, f"Plan_{datetime.date.today()}.pdf", type="primary")

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
                marker_color = colors.get(state, fallback_colors[color_idx % len(fallback_colors)])
                if state not in colors: color_idx += 1
                series_data = {'name': state, 'categories': ['全區進度圖', 1, col, len(sub_df), col], 'values': ['全區進度圖', 1, col+1, len(sub_df), col+1], 'marker': {'type': 'circle', 'size': 6, 'fill': {'color': marker_color}, 'border': {'color': marker_color}}}
                if state == '未完成': series_data['marker'] = {'type': 'circle', 'size': 6, 'fill': {'none': True}, 'border': {'color': marker_color}}
                if state != '未完成':
                    clbls = [{'value': f'=全區進度圖!${xlsxwriter.utility.xl_col_to_name(col+2)}${ri+2}'} for ri in range(len(sub_df))]
                    series_data['data_labels'] = {'custom': clbls, 'position': 'above', 'font': {'size': 8}}
                ch.add_series(series_data)
                col += 4
            ch.set_x_axis({'visible': False}); ch.set_y_axis({'visible': False}); ws.insert_chart('B2', ch)
        return out.getvalue()
    st.sidebar.download_button("🟢 匯出 Excel (全區報表)", xl_gen(df_history, df_p), f"Report_{datetime.date.today()}.xlsx")
