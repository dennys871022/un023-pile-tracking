import streamlit as st
import pdfplumber
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="排樁進度管理系統", layout="wide")
st.title("UN023 排樁工程進度自動化儀表板")

# 1. 檔案上傳區
uploaded_file = st.file_uploader("請上傳排樁施工圖 (PDF格式)", type="pdf")

if uploaded_file is not None:
    st.info("系統正在解析圖面座標，請稍候...")

    # 2. 擷取座標函數 (使用快取避免重複讀取)
    @st.cache_data
    def extract_pile_data(file):
        piles = []
        with pdfplumber.open(file) as pdf:
            for i, page in enumerate(pdf.pages):
                words = page.extract_words()
                for word in words:
                    text = word['text']
                    # 匹配樁號規則：大寫字母開頭接數字 (例如 A1, P237, B2)
                    if re.match(r'^[A-Z]\d{1,3}$', text):
                        piles.append({
                            '樁號': text,
                            'X': word['x0'],
                            'Y': -word['top'], # Y軸反轉以符合平面圖視覺
                            '頁碼': i + 1
                        })
        return pd.DataFrame(piles)

    df = extract_pile_data(uploaded_file)

    if not df.empty:
        st.success(f"解析完成！共抓取 {len(df)} 支樁位座標。")

        # 3. 進度輸入區
        completed_input = st.text_input("輸入今日完成樁號 (請以半形逗號分隔，如 A1,A2,P237)")
        completed_list = [p.strip() for p in completed_input.split(',')] if completed_input else []

        # 更新狀態
        df['狀態'] = df['樁號'].apply(lambda x: '今日完成' if x in completed_list else '未完成')

        # 4. 繪製視覺化圖面
        fig = px.scatter(
            df, x='X', y='Y', text='樁號', color='狀態',
            color_discrete_map={'未完成': 'lightgrey', '今日完成': 'red'},
            title="排樁平面動態著色圖"
        )
        fig.update_traces(textposition='top center', marker=dict(size=12, line=dict(width=1, color='DarkSlateGrey')))
        fig.update_layout(xaxis_visible=False, yaxis_visible=False, plot_bgcolor='white', height=700)

        st.plotly_chart(fig, use_container_width=True)

        # 5. 統計資訊
        valid_completed = len([p for p in completed_list if p in df['樁號'].values])
        st.metric(label="今日完成數量", value=valid_completed)
    else:
        st.warning("無法從此 PDF 中辨識到符合規則的樁號。請確認檔案是否為向量格式。")
