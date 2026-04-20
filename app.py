import streamlit as st
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="排樁進度管理系統", layout="wide")
st.title("UN023 排樁工程進度自動化儀表板")

# 1. 檔案上傳區 (改為上傳 CSV)
uploaded_file = st.file_uploader("請上傳由 AutoCAD 匯出的 CSV 座標檔", type="csv")

if uploaded_file is not None:
    st.info("系統正在讀取 CAD 座標資料...")

    @st.cache_data
    def load_and_clean_data(file):
        # 嘗試讀取 CSV (處理中文編碼問題)
        try:
            df = pd.read_csv(file, encoding='utf-8')
        except UnicodeDecodeError:
            file.seek(0)
            df = pd.read_csv(file, encoding='big5') # 台灣 CAD 常見編碼
        
        # 尋找對應的欄位名稱 (X座標、Y座標、內容)
        x_col = next((col for col in df.columns if 'X' in col.upper()), None)
        y_col = next((col for col in df.columns if 'Y' in col.upper()), None)
        text_col = next((col for col in df.columns if '內容' in col or '值' in col or 'VALUE' in col.upper()), None)

        if not (x_col and y_col and text_col):
            return None

        # 清理 AutoCAD 格式文字 (消除字型亂碼)
        def clean_autocad_text(text):
            if not isinstance(text, str):
                return str(text)
            # 移除 \f...; 等 AutoCAD 格式碼
            text = re.sub(r'\\[^;]+;', '', text)
            text = re.sub(r'[{}]', '', text)
            return text.strip()

        df['樁號'] = df[text_col].apply(clean_autocad_text)
        
        # 篩選出真正的樁號 (以英文字母開頭接數字，如 P500, A1)
        pile_pattern = re.compile(r'^[A-Za-z]\-?\d+$')
        df_piles = df[df['樁號'].apply(lambda x: bool(pile_pattern.match(x)))].copy()
        
        # 確保座標為數值格式
        df_piles['X'] = pd.to_numeric(df_piles[x_col], errors='coerce')
        df_piles['Y'] = pd.to_numeric(df_piles[y_col], errors='coerce')
        
        return df_piles.dropna(subset=['X', 'Y'])

    df = load_and_clean_data(uploaded_file)

    if df is not None and not df.empty:
        st.success(f"解析成功！共匯入 {len(df)} 支排樁/預壘樁座標。")

        # 3. 進度輸入區
        completed_input = st.text_input("輸入今日完成樁號 (請以半形逗號分隔，例如 P500, P199, B-4)")
        # 將輸入的字串轉為大寫並清除多餘空白
        completed_list = [p.strip().upper() for p in completed_input.split(',')] if completed_input else []

        # 4. 更新圖面狀態
        df['樁號大寫'] = df['樁號'].str.upper()
        df['狀態'] = df['樁號大寫'].apply(lambda x: '今日完成' if x in completed_list else '未完成')
# --- 5. 繪製視覺化圖面 ---
        
        # 在側邊欄加入一個手動縮放控制
        st.sidebar.header("圖面設定")
        plot_height = st.sidebar.slider("圖面放大倍率 (高度)", min_value=500, max_value=2500, value=1000, step=100)

        # 計算資料邊界，強行緊縮
        x_min, x_max = df['X'].min(), df['X'].max()
        y_min, y_max = df['Y'].min(), df['Y'].max()
        # 增加 2% 的微小邊距避免點被切到
        x_margin = (x_max - x_min) * 0.02
        y_margin = (y_max - y_min) * 0.02

        fig = px.scatter(
            df, x='X', y='Y', text='樁號', color='狀態',
            color_discrete_map={'未完成': 'lightgrey', '今日完成': 'red'},
            hover_data={'X': False, 'Y': False, '樁號大寫': False}
        )
        
        fig.update_traces(
            textposition='top center', 
            marker=dict(size=10, line=dict(width=1, color='DarkSlateGrey')), 
            textfont=dict(size=12, color='black') 
        )
        
        fig.update_layout(
            # 強制 X 軸範圍
            xaxis=dict(
                range=[x_min - x_margin, x_max + x_margin],
                visible=False,
                showgrid=False
            ), 
            # 強制 Y 軸範圍並鎖定 1:1
            yaxis=dict(
                range=[y_min - y_margin, y_max + y_margin],
                scaleanchor="x", 
                scaleratio=1,
                visible=False,
                showgrid=False
            ),
            plot_bgcolor='white', 
            margin=dict(l=0, r=0, t=0, b=0), # 全域邊距歸零
            height=plot_height,              # 使用滑桿控制的高度
            legend=dict(
                title="施工狀態", 
                orientation="h", 
                yanchor="bottom", 
                y=1.02, 
                xanchor="right", 
                x=1
            )
        )

        # 顯示圖表
        st.plotly_chart(
            fig, 
            use_container_width=True, 
            config={
                'scrollZoom': True,      # 開啟滑輪縮放
                'displayModeBar': True,  # 顯示工具列
                'modeBarButtonsToAdd': ['drawline', 'drawopenpath', 'eraselayer']
            }
        )
        # 6. 進度統計
        valid_completed = len(df[df['狀態'] == '今日完成'])
        st.metric(label="✅ 今日已點亮樁數", value=valid_completed)
        
    else:
        st.error("無法正確讀取 CSV 檔。請確認檔案是由 AutoCAD 匯出，且包含 X、Y 座標欄位。")
