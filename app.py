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

        # --- 3. 進度輸入區 (跳支與排序智慧版) ---
        st.markdown("### 📝 今日進度與施作順序登錄")
        
        mode = st.radio(
            "選擇施工模式：", 
            ["連續施工 (1, 2, 3...)", "4支一循環 (1, 5, 9...)", "3支一循環 (1, 4, 7...)"], 
            horizontal=True
        )

        step = 1
        if "4支" in mode: step = 4
        elif "3支" in mode: step = 3

        st.caption(f"💡 提示：請直接輸入數字或區間 (如 1-100)，系統會自動補 P 並依每 {step} 支跳號點亮。")
        completed_input = st.text_input("輸入完成樁號區間：")

        completed_ordered_list = []
        if completed_input:
            clean_input = re.sub(r'[pP]', '', completed_input)
            raw_items = re.split(r'[,\s]+', clean_input.strip())
            
            for item in raw_items:
                if not item: continue
                
                if '-' in item:
                    parts = item.split('-')
                    if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                        start, end = int(parts[0]), int(parts[1])
                        # 依照輸入順序與步長生成
                        rng = range(start, end + 1, step) if start <= end else range(start, end - 1, -step)
                        for i in rng:
                            p_id = f"P{i}"
                            if p_id not in completed_ordered_list:
                                completed_ordered_list.append(p_id)
                elif item.isdigit():
                    p_id = f"P{item}"
                    if p_id not in completed_ordered_list:
                        completed_ordered_list.append(p_id)
        
        # 建立順序對照表
        seq_map = {p_id: i + 1 for i, p_id in enumerate(completed_ordered_list)}

        # --- 4. 資料處理與動態標籤 ---
        df['樁號大寫'] = df['樁號'].str.upper().str.strip()
        df['施作順序'] = df['樁號大寫'].map(seq_map)
        df['狀態'] = df['施作順序'].apply(lambda x: '今日完成' if pd.notnull(x) else '未完成')
        
        # 產生顯示標籤：若完成則顯示 "P1 (1)"，未完成則顯示 "P1"
        def make_label(row):
            if row['狀態'] == '今日完成':
                return f"{row['樁號']} ({int(row['施作順序'])})"
            return row['樁號']
        
        df['顯示標籤'] = df.apply(make_label, axis=1)

        # --- 5. 繪製視覺化圖面 ---
        st.sidebar.header("圖面設定")
        plot_height = st.sidebar.slider("圖面放大倍率 (高度)", 500, 2500, 800, 100)

        fig = px.scatter(
            df, x='X', y='Y', 
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
        
        # 緊縮座標軸消除白邊
        x_min, x_max = df['X'].min(), df['X'].max()
        y_min, y_max = df['Y'].min(), df['Y'].max()
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

        # 6. 進度統計
        valid_completed = len(df[df['狀態'] == '今日完成'])
        st.metric(label="✅ 今日預計完成數量", value=valid_completed)
        
    else:
        st.error("無法正確讀取 CSV 檔。請確認檔案是由 AutoCAD 匯出，且包含 X、Y 座標欄位。")
