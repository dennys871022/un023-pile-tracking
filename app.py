import streamlit as st
import pandas as pd
import plotly.express as px
import re
import datetime
import io

st.set_page_config(page_title="UN023 排樁進度管理系統", layout="wide")
st.title("📊 UN023 排樁進度管理系統 (雙機報表版)")

# --- 1. 讀取座標底圖 ---
@st.cache_data
def load_base_data():
    try:
        try:
            df = pd.read_csv('排樁座標.csv', encoding='utf-8')
        except:
            df = pd.read_csv('排樁座標.csv', encoding='big5')
        
        x_col = next((c for c in df.columns if 'X' in c.upper()), None)
        y_col = next((c for c in df.columns if 'Y' in c.upper()), None)
        text_col = next((c for c in df.columns if '內容' in c or '值' in c), None)
        
        df['樁號'] = df[text_col].apply(lambda x: re.sub(r'\\[^;]+;|[{}]', '', str(x)).strip().upper())
        df = df[df['樁號'].str.match(r'^P\d+$')]
        df['數字'] = df['樁號'].str.extract(r'(\d+)').astype(int)
        df = df[df['數字'] <= 613]
        
        df['X'] = pd.to_numeric(df[x_col], errors='coerce')
        df['Y'] = pd.to_numeric(df[y_col], errors='coerce')
        
        df = df.drop_duplicates(subset=['樁號'])
        return df.dropna(subset=['X', 'Y']).sort_values('數字')
    except Exception as e:
        st.error(f"底圖載入出錯: {e}")
        return None

df_base = load_base_data()

# --- 2. 歷史進度管理 (系統存檔) ---
st.sidebar.header("📂 系統存檔區 (防遺失)")
history_file = st.sidebar.file_uploader("1️⃣ 每日開工：匯入昨日進度 CSV", type="csv")

if 'history' not in st.session_state:
    st.session_state['history'] = []

if history_file is not None:
    try:
        df_hist = pd.read_csv(history_file)
        # 相容舊版 CSV：如果以前的存檔沒有機台，自動補上
        if '機台' not in df_hist.columns:
            df_hist['機台'] = '未指定'
            
        st.session_state['history'] = df_hist.drop_duplicates(subset=['樁號']).to_dict('records')
        st.sidebar.success("✅ 歷史進度已成功匯入！")
    except Exception as e:
        st.sidebar.error("檔案讀取失敗，請確認格式。")

if st.session_state['history']:
    df_download = pd.DataFrame(st.session_state['history'])
    csv_data = df_download.to_csv(index=False).encode('utf-8-sig')
    st.sidebar.download_button(
        label="2️⃣ 每日收工：下載系統存檔 (CSV)",
        data=csv_data,
        file_name=f"排樁系統備份_{datetime.date.today()}.csv",
        mime="text/csv",
        type="primary"
    )
    
if st.sidebar.button("🗑️ 清空目前網頁暫存紀錄"):
    st.session_state['history'] = []
    st.rerun()

st.sidebar.markdown("---")

# --- 3. 施工登錄介面 ---
st.markdown("### 📝 施工進度登錄")

c_date, c_machine, c_mode = st.columns([1, 1, 2])
today = str(c_date.date_input("施工日期", datetime.date.today()))
machine_val = c_machine.radio("施工機台：", ["A車", "B車"], horizontal=True)
mode = c_mode.radio("跳支模式：", ["連續 (1, 2...)", "4支一循環 (1, 5...)", "3支一循環 (1, 4...)"], horizontal=True)

step = 1
if "4支" in mode: step = 4
elif "3支" in mode: step = 3

def process_piles(new_piles):
    current_max_seq = 0
    if st.session_state['history']:
        current_max_seq = max([int(item['施作順序']) for item in st.session_state['history'] if pd.notna(item['施作順序'])])
    
    added_count = 0
    for pid in new_piles:
        pid = pid.upper().strip()
        if not any(x['樁號'] == pid for x in st.session_state['history']):
            current_max_seq += 1
            st.session_state['history'].append({
                '樁號': pid,
                '施工日期': today,
                '機台': machine_val,
                '施作順序': current_max_seq
            })
            added_count += 1
            
    if added_count > 0:
        st.success(f"✅ 已登錄！{machine_val} 新增 {added_count} 支樁。")
    else:
        st.info("ℹ️ 輸入的樁號均已登錄過，無新增。")

tab1, tab2 = st.tabs(["🎯 起點自動推算", "✏️ 區間手動輸入"])

with tab1:
    with st.form("auto_form"):
        col1, col2, col3 = st.columns(3)
        s_num = col1.number_input("開始數字 (P)", 1, 613, 1)
        direct = col2.radio("方向", ["遞增 (+)", "遞減 (-)"])
        amount = col3.number_input("施作數量", 1, 100, 10)
        if st.form_submit_button("確認登錄"):
            plist = []
            curr = s_num
            for _ in range(amount):
                if 1 <= curr <= 613:
                    plist.append(f"P{curr}")
                curr = curr + step if "遞增" in direct else curr - step
            process_piles(plist)

with tab2:
    with st.form("manual_form"):
        raw_in = st.text_input("輸入區間 (例如: 1-50, 60)")
        if st.form_submit_button("確認登錄"):
            plist = []
            if raw_in:
                clean = re.sub(r'[pP]', '', raw_in)
                parts = re.split(r'[,\s]+', clean.strip())
                for p in parts:
                    if '-' in p:
                        try:
                            s, e = map(int, p.split('-'))
                            rng = range(s, e+1, step) if s <= e else range(s, e-1, -step)
                            for i in rng: plist.append(f"P{i}")
                        except: pass
                    elif p.isdigit():
                        plist.append(f"P{p}")
            process_piles(plist)

# --- 4. 圖面合併與渲染 ---
if df_base is not None:
    df_plot = df_base.copy()
    
    if st.session_state['history']:
        df_hist = pd.DataFrame(st.session_state['history'])
        df_plot = df_plot.merge(df_hist, on='樁號', how='left')
        df_plot['狀態'] = df_plot['施工日期'].fillna('未完成')
        
        # 標籤顯示：包含機台與順序
        def make_label(r):
            if pd.notna(r['施作順序']):
                machine_str = r['機台'] if '機台' in r and pd.notna(r['機台']) else ''
                return f"{r['樁號']} ({int(r['施作順序'])}) {machine_str}"
            return r['樁號']
            
        df_plot['標籤'] = df_plot.apply(make_label, axis=1)
    else:
        df_plot['狀態'] = '未完成'
        df_plot['標籤'] = df_plot['樁號']

    st.sidebar.metric("目前網頁暫存完成數", len(st.session_state['history']))
    h = st.sidebar.slider("畫布高度", 600, 2500, 1000)

    fig = px.scatter(
        df_plot, x='X', y='Y', text='標籤', color='狀態',
        color_discrete_map={'未完成': 'lightgrey'}, 
        hover_data={'X': False, 'Y': False, '標籤': False}
    )
    
    fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='DarkSlateGrey')))
    fig.update_layout(
        xaxis=dict(visible=False, showgrid=False),
        yaxis=dict(scaleanchor="x", scaleratio=1, visible=False, showgrid=False),
        plot_bgcolor='white', margin=dict(l=0, r=0, t=0, b=0), height=h,
        legend=dict(title="施工日期", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

    # --- 5. 產生圖文並茂的 Excel 報表 ---
    st.markdown("---")
    st.subheader("📑 匯出呈報總表 (Excel 圖加表)")
    
    if st.session_state['history']:
        def generate_excel_report(df_export):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 第一頁：原始清單
                df_export.to_excel(writer, index=False, sheet_name='進度明細表')
                
                # 第二頁：統計與圖表
                # 建立透視表：列為日期，欄為機台
                if '機台' in df_export.columns:
                    summary = df_export.pivot_table(index='施工日期', columns='機台', values='樁號', aggfunc='count', fill_value=0)
                else:
                    summary = df_export.groupby('施工日期').size().reset_index(name='完成支數').set_index('施工日期')
                
                summary.to_excel(writer, sheet_name='每日圖表統計')
                
                # 畫 Excel 長條圖
                workbook = writer.book
                worksheet = writer.sheets['每日圖表統計']
                chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'}) # 堆疊長條圖
                
                for i, col in enumerate(summary.columns):
                    chart.add_series({
                        'name':       ['每日圖表統計', 0, i + 1],
                        'categories': ['每日圖表統計', 1, 0, len(summary), 0],
                        'values':     ['每日圖表統計', 1, i + 1, len(summary), i + 1],
                        'data_labels': {'value': True} # 顯示柱子上的數字
                    })
                
                chart.set_title({'name': '排樁每日施工進度 (A車/B車)'})
                chart.set_x_axis({'name': '日期'})
                chart.set_y_axis({'name': '完成支數'})
                chart.set_size({'width': 720, 'height': 480})
                
                worksheet.insert_chart('E2', chart)
                
            return output.getvalue()
        
        # 執行匯出
        df_to_export = pd.DataFrame(st.session_state['history'])
        excel_data = generate_excel_report(df_to_export)
        
        st.download_button(
            label="📊 點此下載 Excel 總表 (含統計圖與數據)",
            data=excel_data,
            file_name=f"排樁進度總表_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("尚無歷史紀錄，無法產生報表。")
