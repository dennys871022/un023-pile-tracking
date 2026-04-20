import streamlit as st
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="排樁進度管理系統", layout="wide")
st.title("UN023 排樁工程進度自動化儀表板")

# 1. 檔案上傳區
uploaded_file = st.file_uploader("請上傳由 AutoCAD 匯出的 CSV 座標檔", type="csv")

if uploaded_file is not None:
    st.info("系統正在讀取 CAD 座標資料...")

    @st.cache_data
    def load_and_clean_data(file):
        try:
            df = pd.read_csv(file, encoding='utf-8')
        except UnicodeDecodeError:
            file.seek(0)
            df = pd.read_csv(file, encoding='big5') 
        
        x_col = next((col for col in df.columns if 'X' in col.upper()), None)
        y_col = next((col for col in df.columns if 'Y' in col.upper()), None)
        text_col = next((col for col in df.columns if '內容' in col or '值' in col or 'VALUE' in col.upper()), None)

        if not (x_col and y_col and text_col):
            return None

        def clean_autocad_text(text):
            if not isinstance(text, str):
                return str(text)
            text = re.sub(r'\\[^;]+;', '', text)
            text = re.sub(r'[{}]', '', text)
            return text.strip()

        df['樁號'] = df[text_col].apply(clean_autocad_text)
        
        pile_pattern = re.compile(r'^[A-Za-z]\-?\d+$')
        df_piles = df[df['樁號'].apply(lambda x: bool(pile_pattern.match(x)))].copy()
        
        df_piles['X'] = pd.to_numeric(df_piles[x_col], errors='coerce')
        df_piles['Y'] = pd.to_numeric(df_piles[y_col], errors='coerce')
        
        return df_piles.dropna(subset=['X', 'Y'])

    df = load_and_clean_data(uploaded_file)

    if df is not None and not df.empty:
        st.success(f"解析成功！共匯入 {len(df)} 支樁位座標。")

        # 初始化 Session State，確保按下確認鍵後狀態不會因拉動側邊欄而消失
        if 'active_piles' not in st.session_state:
            st.session_state['active_piles'] = []

        # --- 3. 進度輸入區 ---
        st.markdown("### 📝 今日進度與施作順序登錄")
        
        mode = st.radio(
            "選擇施工模式：", 
            ["連續施工 (1, 2, 3...)", "4支一循環 (1, 5, 9...)", "3支一循環 (1, 4, 7...)"], 
            horizontal=True
        )

        step = 1
        if "4支" in mode: step = 4
        elif "3支" in mode: step = 3

        # 建立兩個分頁供選擇輸入模式
        tab1, tab2 = st.tabs(["🎯 起點自動推算模式", "✏️ 自由輸入模式 (區間/混合)"])

        with tab1:
            with st.form("calc_form"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    start_num = st.number_input("起始樁號 (數字)", min_value=1, value=1, step=1)
                with col2:
                    direction = st.radio("施作方向", ["編號遞增 (+)", "編號遞減 (-)"])
                with col3:
                    count = st.number_input("預計施作數量 (支)", min_value=1, value=10, step=1)

                submit_calc = st.form_submit_button("✅ 確認執行")

                if submit_calc:
                    temp_list = []
                    current_num = start_num
                    for _ in range(count):
                        if current_num < 1:
                            break
                        p_id = f"P{current_num}"
                        if p_id not in temp_list:
                            temp_list.append(p_id)
                        
                        if "遞增" in direction:
                            current_num += step
                        else:
                            current_num -= step
                    st.session_state['active_piles'] = temp_list

        with tab2:
            with st.form("free_form"):
                completed_input = st.text_input("輸入完成樁號區間 (例如: 1-100)：")
                submit_free = st.form_submit_button("✅ 確認執行")

                if submit_free:
                    temp_list = []
                    if completed_input:
                        clean_input = re.sub(r'[pP]', '', completed_input)
                        raw_items = re.split(r'[,\s]+', clean_input.strip())
                        
                        for item in raw_items:
                            if not item: continue
                            
                            if '-' in item:
                                parts = item.split('-')
                                if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                                    s_num, e_num = int(parts[0]), int(parts[1])
                                    rng = range(s_num, e_num + 1, step) if s_num <= e_num else range(s_num, e_num - 1, -step)
                                    for i in rng:
                                        p_id = f"P{i}"
                                        if p_id not in temp_list: temp_list.append(p_id)
                            elif item.isdigit():
                                p_id = f"P{item}"
                                if p_id not in temp_list: temp_list.append(p_id)
                    st.session_state['active_piles'] = temp_list

        # 從 Session State 讀取目前欲著色的樁號清單
        completed_ordered_list = st.session_state['active_piles']
        seq_map = {p_id: i + 1 for i, p_id in enumerate(completed_ordered_list)}

        # --- 4. 資料處理與動態標籤 ---
        df['樁號大寫'] = df['樁號'].str.upper().str.strip()
        df['施作順序'] = df['樁號大寫'].map(seq_map)
        df['狀態'] = df['施作順序'].apply(lambda x: '今日完成' if pd.notnull(x) else '未完成')
        
        def make_label(row):
            if row['狀態'] == '今日完成':
                return f"{row['樁號']} ({int(row['施作順序'])})"
            return row['樁號']
        
        df['顯示標籤'] = df.apply(make_label, axis=1)

        # --- 5. 繪製視覺化圖面與裁切設定 ---
        st.sidebar.header("🛠️ 圖面設定與裁切")
        plot_height = st.sidebar.slider("圖面放大倍率 (高度)", 500, 2500, 800, 100)

        st.sidebar.markdown("***")
        st.sidebar.subheader("✂️ 濾除 CAD 雜訊")
        
        x_min_val, x_max_val = float(df['X'].min()), float(df['X'].max())
        y_min_val, y_max_val = float(df['Y'].min()), float(df['Y'].max())

        x_range = st.sidebar.slider("X 軸範圍限制", x_min_val, x_max_val, (x_min_val, x_max_val))
        y_range = st.sidebar.slider("Y 軸範圍限制", y_min_val, y_max_val, (y_min_val, y_max_val))

        df_plot = df[(df['X'] >= x_range[0]) & (df['X'] <= x_range[1]) & (df['Y'] >= y_range[0]) & (df['Y'] <= y_range[1])].copy()

        fig = px.scatter(
            df_plot, x='X', y='Y', 
            text='顯示標籤', 
            color='狀態',
            color_discrete_map={'未完成': 'lightgrey', '今日完成': 'red'},
            hover_data={'X': False, 'Y': False, '顯示標籤': False}
        )
        
        fig.update_traces(
            textposition='top center', 
            marker=dict(size=10, line=dict(width=1, color='DarkSlateGrey')), 
            textfont=dict(size=12, color='black') 
        )
        
        if not df_plot.empty:
            x_min, x_max = df_plot['X'].min(), df_plot['X'].max()
            y_min, y_max = df_plot['Y'].min(), df_plot['Y'].max()
            x_margin, y_margin = (x_max - x_min) * 0.05, (y_max - y_min) * 0.05

            fig.update_layout(
                xaxis=dict(range=[x_min - x_margin, x_max + x_margin], visible=False, showgrid=False), 
                yaxis=dict(range=[y_min - y_margin, y_max + y_margin], scaleanchor="x", scaleratio=1, visible=False, showgrid=False),
                plot_bgcolor='white', 
                margin=dict(l=0, r=0, t=0, b=0), 
                height=plot_height,
                legend=dict(title="施工狀態", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )

        st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True})

        # --- 6. 進度統計 ---
        valid_completed = len(df_plot[df_plot['狀態'] == '今日完成'])
        st.metric(label="✅ 今日預計完成數量", value=valid_completed)
        
    else:
        st.error("無法正確讀取 CSV 檔。請確認檔案是由 AutoCAD 匯出，且包含 X、Y 座標欄位。")
