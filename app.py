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

# 5. 繪製視覺化圖面
        # 計算圖形的長寬分佈，自動決定最佳畫布高度，消除過多白邊
        x_spread = df['X'].max() - df['X'].min()
        y_spread = df['Y'].max() - df['Y'].min()
        aspect_ratio = y_spread / x_spread if x_spread > 0 else 1
        
        # 以網頁可用寬度約 900px 為基準，推算合理的動態高度
        dynamic_height = max(500, min(1200, int(900 * aspect_ratio)))

        fig = px.scatter(
            df, x='X', y='Y', text='樁號', color='狀態',
            color_discrete_map={'未完成': 'lightgrey', '今日完成': 'red'},
            hover_data={'X': False, 'Y': False, '樁號大寫': False}
        )
        
        # 微調標記大小與字體，避免全區圖的文字過度重疊
        fig.update_traces(
            textposition='top center', 
            marker=dict(size=8, line=dict(width=1, color='DarkSlateGrey')), 
            textfont=dict(size=10, color='black') 
        )
        
        fig.update_layout(
            xaxis=dict(visible=False), 
            yaxis=dict(visible=False, scaleanchor="x", scaleratio=1),
            plot_bgcolor='white', 
            margin=dict(l=0, r=0, t=40, b=0), # 將四周系統預設的白邊歸零
            height=dynamic_height,            # 套用動態計算的最佳高度
            legend=dict(title="施工狀態", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )

        # 啟動雙擊重置畫面大小的功能
        st.plotly_chart(fig, use_container_width=True, config={'doubleClick': 'reset'})

        # 6. 進度統計
        valid_completed = len(df[df['狀態'] == '今日完成'])
        st.metric(label="✅ 今日已點亮樁數", value=valid_completed)
        
    else:
        st.error("無法正確讀取 CSV 檔。請確認檔案是由 AutoCAD 匯出，且包含 X、Y 座標欄位。")
