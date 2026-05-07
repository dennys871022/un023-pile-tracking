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
    import matplotlib.gridspec as gridspec
    MATPLOTLIB_READY = True
except ImportError:
    MATPLOTLIB_READY = False

st.set_page_config(page_title="UN023 排樁進度系統 V29", layout="wide")
st.title("🏗️ UN023 排樁進度管理 (手繪報表擬真版)")

# --- 字體設定 ---
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

# --- 資料庫邏輯 (不變，確保穩定性) ---
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
    except Exception:
        return None, None

ss, sh_main = get_gs_connection()
df_history = fetch_current_data(sh_main) if sh_main else pd.DataFrame()

# --- 動態統計邏輯 ---
total_done_auto = len(df_history)
today_done_auto = 0
week_start_str = ""
today_roc_str = ""
if not df_history.empty:
    df_history['施工日期_DT'] = pd.to_datetime(df_history['施工日期'], errors='coerce')
    latest_dt = df_history['施工日期_DT'].max()
    today_done_auto = len(df_history[df_history['施工日期_DT'] == latest_dt])
    today_roc_str = f"{latest_dt.year-1911}/{latest_dt.month:02d}/{latest_dt.day:02d}"
    monday = latest_dt - pd.Timedelta(days=latest_dt.weekday())
    this_week_data = df_history[df_history['施工日期_DT'] >= monday]
    if not this_week_data.empty:
        earliest_this_week = this_week_data['施工日期_DT'].min()
        week_start_str = f"{earliest_this_week.year-1911}/{earliest_this_week.month:02d}/{earliest_this_week.day:02d}"

# --- 側邊欄自訂 (新增表格數據輸入) ---
st.sidebar.markdown("### 📄 PDF 報表自訂內容")
pdf_loc_note = st.sidebar.text_input("右上角位置", "滯洪池BC")
pdf_week_est = st.sidebar.number_input("本週預計完成 (支)", value=36)
pdf_today_done = st.sidebar.number_input("本日完成 (支)", value=today_done_auto)
pdf_cum_done = st.sidebar.number_input("累積完成 (支)", value=total_done_auto)
pdf_comment = st.sidebar.text_area("下方備註", "下午1:00~3:30洽談灌漿，導航下半午完3支樁")

st.sidebar.markdown("#### 📊 表格數據設定")
tab_total_days = st.sidebar.number_input("預計施工總天數", value=100)
tab_daily_target = st.sidebar.number_input("預計平均支數/日", value=6.5, step=0.1)
tab_concrete_used = st.sidebar.number_input("今日水泥用量 (m3)", value=21.0)

# --- 繪圖輔助功能 (穩定版) ---
def process_status_logic(df_hist, df_b):
    plot_df = df_b[['樁號', 'X', 'Y']].copy()
    if df_hist.empty:
        plot_df['狀態'] = '未完成'; plot_df['標籤'] = plot_df['樁號']
        return plot_df
    hist = df_hist.copy()
    hist['標籤'] = hist.apply(lambda r: f"{r['樁號']}({str(r.get('機台','A'))[0]}{int(r.get('施作順序',0))})", axis=1)
    hist['施工日期_DT'] = pd.to_datetime(hist['施工日期'], errors='coerce')
    max_date = hist['施工日期_DT'].max()
    monday = max_date - pd.Timedelta(days=max_date.weekday())
    hist['狀態'] = hist['施工日期_DT'].apply(lambda dt: '[已完成]' if dt < monday else dt.strftime('%m/%d'))
    plot_df = plot_df.merge(hist[['樁號', '狀態', '標籤']], on='樁號', how='left')
    plot_df['狀態'] = plot_df['狀態'].fillna('未完成')
    plot_df['標籤'] = plot_df['標籤'].fillna(plot_df['樁號'])
    return plot_df

df_p = process_status_logic(df_history, df_base)

# --- PDF 智慧報表生成引擎 V29 ---
def pdf_gen_v29(p_df, loc, w_est, t_done, c_done, w_start, comment, sel_piles):
    font_name = setup_chinese_font()
    if font_name: plt.rcParams['font.family'] = font_name
    
    # 建立巨大畫布
    fig = plt.figure(figsize=(24, 18))
    gs = gridspec.GridSpec(2, 2, width_ratios=[1, 1.2], height_ratios=[1, 0.15])
    
    # 1. 右側：現場樁位圖
    ax_map = fig.add_subplot(gs[0, 1])
    
    # 過濾顯示範圍
    target_df = p_df[p_df['樁號'].isin(sel_piles)].copy() if sel_piles else p_df.copy()
    
    # 顏色定義
    states = ['未完成', '[已完成]'] + sorted([s for s in target_df['狀態'].unique() if s not in ['未完成', '[已完成]']])
    color_palette = px.colors.qualitative.Plotly
    colors = {'未完成': '#D3D3D3', '[已完成]': '#FFB6C1'}
    
    texts = []
    for i, state in enumerate(states):
        sub = target_df[target_df['狀態'] == state]
        if sub.empty: continue
        c = colors.get(state, color_palette[i % len(color_palette)])
        
        if state == '未完成':
            ax_map.scatter(sub['X'], sub['Y'], facecolors='none', edgecolors=c, s=180, lw=1.5, alpha=0.5)
        else:
            # 實心圓點
            ax_map.scatter(sub['X'], sub['Y'], color=c, s=200, zorder=3, label=f"{state} 樁號 ○ 施作順序")
            for _, row in sub.iterrows():
                # 文字放大且與圓點同色
                texts.append(ax_map.text(row['X'], row['Y'], row['標籤'], fontsize=12, fontweight='bold', color=c))

    adjust_text(texts, ax=ax_map, expand_points=(1.8, 1.8), arrowprops=dict(arrowstyle='-', color='gray', lw=0.5))
    ax_map.set_aspect('equal'); ax_map.axis('off')
    ax_map.set_title(loc, loc='right', fontsize=55, fontweight='bold', pad=20)
    ax_map.legend(loc='upper right', bbox_to_anchor=(1, 1.05), fontsize=18, frameon=False)

    # 2. 左側：標題與統計
    ax_text = fig.add_subplot(gs[0, 0])
    ax_text.axis('off')
    
    # 報表主標題
    roc_today = f"{datetime.date.today().year-1911}/{datetime.date.today().month:02d}/{datetime.date.today().day:02d}"
    ax_text.text(0, 0.95, f"{roc_today}施作進度回報", fontsize=55, fontweight='bold')
    
    # 統計區
    latest_dt = pd.to_datetime(df_history['施工日期'], errors='coerce').max()
    sunday = latest_dt + datetime.timedelta(days=(6 - latest_dt.weekday()))
    week_range = f"{w_start}~{sunday.year-1911}/{sunday.month:02d}/{sunday.day:02d}"
    
    summary_txt = (
        f"本週預計完成-{w_est}支\n"
        f"{week_range}\n"
        f"本日完成-{t_done}支\n"
        f"{roc_today}\n"
        f"累積完成-{c_done}支"
    )
    ax_text.text(0.05, 0.88, summary_txt, fontsize=30, linespacing=1.6, va='top')

    # 3. 左下側：結構表格
    table_data = [
        ["結構預估每日施工進度", ""],
        ["預計開始時間", "2026/5/6 前期作業工期", "2026/5/5"],
        ["預計施作完工日", f"{tab_total_days} 個施作天數", "2026/8/13"],
        ["隔日預計完工日", f"{613-c_done} 剩餘施作支數", "尚未計算"],
        ["預估每日進度(支/日)", f"{tab_daily_target} 本日施作目標(支)", f"{t_done}"],
        ["材料使用情形", "", ""],
        ["今日用量(m3)", f"{tab_concrete_used} 今日水泥用量(m3)", f"{tab_concrete_used}"],
    ]
    
    table = ax_text.table(cellText=table_data, loc='bottom', cellLoc='left', bbox=[0, 0.05, 0.95, 0.5])
    table.auto_set_font_size(False)
    table.set_fontsize(18)
    for i in range(len(table_data)):
        for j in range(len(table_data[0])):
            table[(i,j)].set_height(0.06)
            if i == 0 or i == 5: table[(i,j)].set_facecolor('#E0E0E0')

    # 4. 底部備註
    ax_note = fig.add_subplot(gs[1, :])
    ax_note.axis('off')
    ax_note.text(0, 0.5, f"備註：{comment}", fontsize=25, fontweight='bold', bbox=dict(facecolor='none', edgecolor='black', pad=10))

    buf = io.BytesIO()
    plt.savefig(buf, format='pdf', bbox_inches='tight')
    plt.close(fig)
    return buf.getvalue()

# --- 下載按鈕區 ---
st.sidebar.markdown("---")
if st.sidebar.download_button("🔴 下載 115年擬真工程報表 (PDF)", pdf_gen_v29(df_p, pdf_loc_note, pdf_week_est, pdf_today_done, pdf_cum_done, week_start_str, pdf_comment, selected_piles), f"Progress_Report_{datetime.date.today()}.pdf", type="primary"):
    st.balloons()
