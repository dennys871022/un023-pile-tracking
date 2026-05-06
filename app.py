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

st.set_page_config(page_title="UN023 排樁進度系統 V21", layout="wide")
st.title("🏗️ UN023 排樁進度管理 (局部框選雙輸出版)")

# ── 常數定義（避免 Magic Number 散落各處）──────────────────────────────────
CYCLE_STEPS = {"4支一循環": 4, "2支一循環": 2}
STATE_COLORS = {"未完成": "#696969", "[已完成]": "#FFB6C1"}
FALLBACK_COLORS = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b", "#e377c2"]
PILE_MIN, PILE_MAX = 1, 613


def get_state_color(state: str, fallback_idx: int) -> str:
    """統一的顏色查詢，xl_gen 與 pdf_gen 共用，避免重複定義。"""
    return STATE_COLORS.get(state, FALLBACK_COLORS[fallback_idx % len(FALLBACK_COLORS)])


# ── 字體設定（雲端中文字體）──────────────────────────────────────────────────
@st.cache_resource
def setup_chinese_font():
    import os
    import urllib.request
    import matplotlib.font_manager as fm
    font_path = "NotoSansCJKtc-Regular.otf"
    if not os.path.exists(font_path):
        try:
            url = (
                "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/"
                "TraditionalChinese/NotoSansCJKtc-Regular.otf"
            )
            urllib.request.urlretrieve(url, font_path)
        except Exception as e:
            st.warning(f"字體下載失敗，將使用系統備用字體: {e}")
    if os.path.exists(font_path):
        fm.fontManager.addfont(font_path)
        return fm.FontProperties(fname=font_path).get_name()
    return None


# ── 底圖載入 ──────────────────────────────────────────────────────────────
@st.cache_data
def load_base_data() -> pd.DataFrame | None:
    """從 CSV 載入樁位座標，自動偵測編碼。"""
    try:
        try:
            import chardet
            with open("排樁座標.csv", "rb") as f:
                enc = chardet.detect(f.read())["encoding"] or "utf-8"
        except ImportError:
            # chardet 不存在時退回手動嘗試
            enc = "utf-8"
            try:
                pd.read_csv("排樁座標.csv", encoding=enc, nrows=1)
            except UnicodeDecodeError:
                enc = "big5"

        df = pd.read_csv("排樁座標.csv", encoding=enc)

        x_col   = next((c for c in df.columns if "X" in c.upper() or "座標" in c), None)
        y_col   = next((c for c in df.columns if "Y" in c.upper() or "座標" in c), None)
        text_col = next((c for c in df.columns if "內容" in c or "值" in c or "樁號" in c), None)

        if not all([x_col, y_col, text_col]):
            st.error("CSV 欄位辨識失敗，請確認含有 X/Y 座標與樁號欄位。")
            return None

        df["樁號"] = df[text_col].apply(
            lambda x: re.sub(r"\[^;]+;|[{}]", "", str(x)).strip().upper()
        )
        df = df[df["樁號"].str.match(r"^P\d+$")]
        df["數字"] = df["樁號"].str.extract(r"(\d+)").astype(int)
        df = df[(df["數字"] >= PILE_MIN) & (df["數字"] <= PILE_MAX)]
        df["X"] = pd.to_numeric(df[x_col], errors="coerce")
        df["Y"] = pd.to_numeric(df[y_col], errors="coerce")
        return (
            df.drop_duplicates(subset=["樁號"])
            .dropna(subset=["X", "Y"])
            .sort_values("數字")
        )
    except FileNotFoundError:
        st.error("找不到 排樁座標.csv，請確認檔案已放置在正確路徑。")
        return None
    except Exception as e:
        st.error(f"底圖載入失敗: {e}")
        return None


df_base = load_base_data()

# 建立樁號 → 座標的快速查詢字典（O(1) 取代迴圈內逐行過濾）
if df_base is not None:
    base_lookup: dict[str, dict] = df_base.set_index("樁號")[["X", "Y"]].to_dict("index")
else:
    base_lookup = {}


# ── Google Sheets 連線（快取，避免每次 rerun 重建 OAuth）──────────────────
@st.cache_resource
def get_gs_connection():
    """建立並快取 Google Sheets 連線，失敗時回傳 None tuple。"""
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        try:
            creds_dict = json.loads(st.secrets["gcp_service_account"])
            sheet_url  = st.secrets["sheet_url"]
        except KeyError as e:
            st.error(f"缺少必要的 secrets 設定：{e}。請在 .streamlit/secrets.toml 中加入。")
            return None, None, None

        creds  = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        ss     = client.open_by_url(sheet_url)

        try:
            sh_main = ss.worksheet("施工明細")
        except gspread.exceptions.WorksheetNotFound:
            sh_main = ss.add_worksheet("施工明細", 1000, 20)
            sh_main.append_row(["樁號", "施工日期", "機台", "施作順序", "X", "Y"])

        try:
            sh_chart = ss.worksheet("系統繪圖區")
        except gspread.exceptions.WorksheetNotFound:
            sh_chart = ss.add_worksheet("系統繪圖區", 700, 60)

        return ss, sh_main, sh_chart

    except gspread.exceptions.APIError as e:
        st.error(f"Google Sheets API 錯誤: {e}")
        return None, None, None
    except Exception as e:
        st.error(f"雲端連線異常: {e}")
        return None, None, None


def fetch_current_data(sh_main) -> pd.DataFrame:
    """從雲端工作表讀取所有施工記錄。"""
    empty = pd.DataFrame(columns=["樁號", "施工日期", "機台", "施作順序", "X", "Y"])
    if sh_main is None:
        return empty
    try:
        records = sh_main.get_all_records()
        if not records:
            return empty
        df = pd.DataFrame(records)
        df["樁號"] = df["樁號"].astype(str).str.upper().str.strip()
        if "機台" not in df.columns:
            df["機台"] = "A車"
        df["施作順序"] = pd.to_numeric(df.get("施作順序", 0), errors="coerce").fillna(0)
        return df
    except gspread.exceptions.APIError as e:
        st.warning(f"資料讀取失敗（API 錯誤）: {e}")
        return empty
    except Exception as e:
        st.warning(f"資料讀取失敗: {e}")
        return empty


# ── Session State：確保 df_history 在 rerun 間保持一致 ─────────────────────
ss_conn, sh_main, sh_chart = get_gs_connection()

if "df_history" not in st.session_state:
    st.session_state.df_history = fetch_current_data(sh_main)

df_history: pd.DataFrame = st.session_state.df_history


def refresh_history():
    """從雲端重新抓取並更新 session_state。"""
    st.session_state.df_history = fetch_current_data(sh_main)


# ── 狀態計算邏輯 ──────────────────────────────────────────────────────────
@st.cache_data
def process_status_logic(
    df_hist_json: str,   # 用 JSON 字串作為 cache key（DataFrame 不可 hash）
    df_base_json: str,
) -> pd.DataFrame:
    df_hist = pd.read_json(io.StringIO(df_hist_json))
    df_b    = pd.read_json(io.StringIO(df_base_json))

    plot_df = df_b[["樁號", "X", "Y"]].copy()
    if df_hist.empty:
        plot_df["狀態"] = "未完成"
        plot_df["標籤"] = plot_df["樁號"]
        return plot_df

    hist = df_hist.copy()

    def label_maker(r):
        m = str(r.get("機台", "A"))[0]
        s = r.get("施作順序", 0)
        return f"{r['樁號']}({m}{int(s)})"

    hist["標籤"]       = hist.apply(label_maker, axis=1)
    hist["施工日期_DT"] = pd.to_datetime(hist["施工日期"], errors="coerce")

    max_date = hist["施工日期_DT"].max()
    if pd.notna(max_date):
        monday = max_date - pd.Timedelta(days=max_date.weekday())

        def set_status(dt):
            if pd.isna(dt):
                return "未完成"
            return "[已完成]" if dt < monday else dt.strftime("%Y-%m-%d")

        hist["狀態"] = hist["施工日期_DT"].apply(set_status)
    else:
        hist["狀態"] = "未完成"

    plot_df = plot_df.merge(hist[["樁號", "狀態", "標籤"]], on="樁號", how="left")
    plot_df["狀態"] = plot_df["狀態"].fillna("未完成")
    plot_df["標籤"] = plot_df["標籤"].fillna(plot_df["樁號"])
    return plot_df


def get_plot_df() -> pd.DataFrame:
    """取得目前最新的狀態 DataFrame（帶快取）。"""
    return process_status_logic(
        df_history.to_json(),
        df_base.to_json() if df_base is not None else "{}",
    )


# ── 雲端圖表同步 ──────────────────────────────────────────────────────────
def sync_to_chart_sheet():
    _, m_now, c_now = get_gs_connection()
    if not m_now or df_history.empty:
        return
    try:
        plot_df   = get_plot_df()
        gs_matrix = pd.DataFrame()
        gs_matrix["X"]      = plot_df["X"]
        gs_matrix["標籤"]   = plot_df["標籤"]
        gs_matrix["未完成"] = plot_df["Y"].where(plot_df["狀態"] == "未完成", None)
        gs_matrix["[已完成]"] = plot_df["Y"].where(plot_df["狀態"] == "[已完成]", None)

        valid_dates = sorted(
            [s for s in plot_df["狀態"].unique() if s not in ["未完成", "[已完成]"]]
        )
        for d in valid_dates:
            gs_matrix[d] = plot_df["Y"].where(plot_df["狀態"] == d, None)

        gs_matrix = gs_matrix.astype(object).where(pd.notnull(gs_matrix), None)
        out_data  = [gs_matrix.columns.values.tolist()] + gs_matrix.values.tolist()
        c_now.clear()
        c_now.update("A1", out_data)
        st.success("✅ 雲端繪圖數據已同步")
    except gspread.exceptions.APIError as e:
        st.error(f"同步失敗（API 錯誤）: {e}")
    except Exception as e:
        st.error(f"同步失敗: {e}")


# ── 資料儲存 ──────────────────────────────────────────────────────────────
def save_data(piles: list[str], work_date: str, machine: str):
    """將新樁號寫入 Google Sheets，並更新 session_state。"""
    if not piles or sh_main is None:
        return

    m_data = df_history[df_history["機台"] == machine]
    seq    = 0 if m_data.empty else pd.to_numeric(m_data["施作順序"]).max()
    new_d  = []

    for p in piles:
        p = p.upper().strip()
        if p not in df_history["樁號"].values:
            seq += 1
            coords = base_lookup.get(p, {"X": 0, "Y": 0})
            new_d.append([p, work_date, machine, int(seq), float(coords["X"]), float(coords["Y"])])

    if new_d:
        try:
            sh_main.append_rows(new_d)
        except gspread.exceptions.APIError as e:
            st.error(f"寫入失敗（API 錯誤）: {e}")
            return

        # 更新 session_state（不重抓，直接 append 加快速度）
        new_df = pd.DataFrame(new_d, columns=["樁號", "施工日期", "機台", "施作順序", "X", "Y"])
        st.session_state.df_history = pd.concat(
            [st.session_state.df_history, new_df], ignore_index=True
        )
        sync_to_chart_sheet()
        st.rerun()


# ── 側邊欄 ────────────────────────────────────────────────────────────────
st.sidebar.header("📂 備份與同步")

if st.sidebar.button("🔄 手動同步雲端數據"):
    refresh_history()
    sync_to_chart_sheet()

up_file = st.sidebar.file_uploader("匯入 Excel/CSV", type=["csv", "xlsx"])
if up_file:
    try:
        df_up = (
            pd.read_excel(up_file, sheet_name="施工明細")
            if up_file.name.endswith(".xlsx")
            else pd.read_csv(up_file)
        )
        new_rows   = []
        curr_piles = df_history["樁號"].tolist()

        for _, row in df_up.iterrows():
            p      = str(row["樁號"]).upper().strip()
            if p not in curr_piles:
                coords = base_lookup.get(p, {"X": 0, "Y": 0})
                new_rows.append([
                    p,
                    str(row["施工日期"]),
                    str(row.get("機台", "A車")),
                    int(row.get("施作順序", 1)),
                    float(coords["X"]),
                    float(coords["Y"]),
                ])

        if new_rows:
            sh_main.append_rows(new_rows)
            refresh_history()
            sync_to_chart_sheet()
            st.sidebar.success(f"已同步 {len(new_rows)} 筆")
            st.rerun()
        else:
            st.sidebar.info("無新增資料（所有樁號已存在）。")

    except KeyError as e:
        st.sidebar.error(f"欄位缺失：{e}，請確認 Excel 含有正確欄位名稱。")
    except Exception as e:
        st.sidebar.error(f"還原失敗: {e}")


# ── 進度登錄 UI ───────────────────────────────────────────────────────────
st.markdown("### 📝 進度登錄")
col_date, col_machine, col_mode = st.columns([1, 1, 2])
work_date = str(col_date.date_input("日期"))
machine   = col_machine.radio("機台", ["A車", "B車"], horizontal=True)
mode      = col_mode.radio("模式", list(CYCLE_STEPS.keys()), horizontal=True)
step      = CYCLE_STEPS[mode]

tab_auto, tab_manual = st.tabs(["🎯 推算", "✏️ 手動"])

with tab_auto:
    with st.form("form_auto"):
        cc1, cc2, cc3 = st.columns(3)
        start_pile = cc1.number_input("起始 P", PILE_MIN, PILE_MAX, 1)
        direction  = cc2.radio("方向", ["遞增", "遞減"])
        count      = cc3.number_input("數量", 1, 100, 10)

        if st.form_submit_button("登錄"):
            plist, cur = [], start_pile
            for _ in range(count):
                if PILE_MIN <= cur <= PILE_MAX:
                    plist.append(f"P{cur}")
                cur = cur + step if direction == "遞增" else cur - step
            save_data(plist, work_date, machine)

with tab_manual:
    with st.form("form_manual"):
        raw = st.text_input("區間 (例：1-50, 55, 60)")
        if st.form_submit_button("登錄"):
            plist = []
            if raw:
                pts = re.split(r"[,\s]+", re.sub(r"[pP]", "", raw))
                for pt in pts:
                    if "-" in pt:
                        try:
                            s, e = map(int, pt.split("-"))
                            rs   = step if s <= e else -step
                            for n in range(s, e + (1 if s <= e else -1), rs):
                                plist.append(f"P{n}")
                        except ValueError:
                            st.warning(f"無法解析區間：{pt}")
                    elif pt.isdigit():
                        plist.append(f"P{pt}")
            save_data(plist, work_date, machine)


# ── 主圖表 ────────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("🗺️ 現場施工全區圖 (框選範圍即可局部導出 PDF)")
st.info(
    "💡 **局部導出教學**：將游標移至圖表右上角，點選 `⬚ (Box Select)` 或 `⸠ (Lasso Select)`，"
    "在圖面上畫出您想要的範圍。側邊欄的 PDF 按鈕將自動鎖定該範圍！若想取消，請在空白處點擊兩下。"
)

df_p          = get_plot_df()
colors_seq    = px.colors.qualitative.Plotly

fig = px.scatter(
    df_p, x="X", y="Y", text="標籤", color="狀態",
    color_discrete_map=STATE_COLORS,
    color_discrete_sequence=colors_seq,
    custom_data=["樁號"],
)
fig.update_traces(
    textposition="top center",
    textfont=dict(size=8),
    marker=dict(size=10, line=dict(width=1, color="white")),
)
fig.update_layout(
    xaxis_visible=False,
    yaxis=dict(scaleanchor="x", scaleratio=1, visible=False),
    height=950,
    plot_bgcolor="white",
    dragmode="pan",
)

selected_piles: list[str] = []
try:
    selection_event = st.plotly_chart(
        fig,
        use_container_width=True,
        config={"scrollZoom": True},
        on_select="rerun",
        selection_mode=("box", "lasso"),
    )
    if (
        selection_event
        and "selection" in selection_event
        and selection_event["selection"]["points"]
    ):
        selected_piles = [pt["customdata"][0] for pt in selection_event["selection"]["points"]]

except TypeError:
    # 兼容較舊的 Streamlit 版本
    st.plotly_chart(fig, use_container_width=True, config={"scrollZoom": True})


# ── 報表下載區 ────────────────────────────────────────────────────────────
if not df_history.empty:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📥 報表與圖面下載")

    def xl_gen(h_df: pd.DataFrame, p_df: pd.DataFrame) -> bytes:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as wr:
            h_df.to_excel(wr, sheet_name="施工明細", index=False)
            wb = wr.book
            ws = wb.add_worksheet("全區進度圖")
            ch = wb.add_chart({"type": "scatter"})
            col = 10

            states = ["未完成", "[已完成]"] + sorted(
                [s for s in p_df["狀態"].unique() if s not in ["未完成", "[已完成]"]]
            )
            fallback_idx = 0

            for state in states:
                sub_df = p_df[p_df["狀態"] == state].reset_index(drop=True)
                if sub_df.empty:
                    continue

                sub_df[["X", "Y", "標籤"]].to_excel(
                    wr, sheet_name="全區進度圖", startcol=col, index=False
                )

                marker_color = get_state_color(state, fallback_idx)
                if state not in STATE_COLORS:
                    fallback_idx += 1

                series_data = {
                    "name": state,
                    "categories": ["全區進度圖", 1, col, len(sub_df), col],
                    "values":     ["全區進度圖", 1, col + 1, len(sub_df), col + 1],
                    "marker": {
                        "type": "circle", "size": 6,
                        "fill":   {"color": marker_color},
                        "border": {"color": marker_color},
                    },
                }
                if state != "未完成":
                    clbls = [
                        {"value": f"=全區進度圖!${xlsxwriter.utility.xl_col_to_name(col + 2)}${ri + 2}"}
                        for ri in range(len(sub_df))
                    ]
                    series_data["data_labels"] = {
                        "custom": clbls, "position": "above", "font": {"size": 8}
                    }
                ch.add_series(series_data)
                col += 4

            today_str = datetime.date.today().strftime("%Y-%m-%d")
            ch.set_title({"name": f"{today_str} 施作進度回報"})
            ch.set_size({"width": 2400, "height": 1500})
            ch.set_x_axis({"visible": False, "major_gridlines": {"visible": False}})
            ch.set_y_axis({"visible": False, "major_gridlines": {"visible": False}})
            ws.insert_chart("B2", ch)
        return out.getvalue()

    st.sidebar.download_button(
        "🟢 匯出 Excel (全區報表)",
        xl_gen(df_history, df_p),
        f"Report_{datetime.date.today()}.xlsx",
        type="secondary",
    )

    if not MATPLOTLIB_READY:
        st.sidebar.error("⚠️ 請確保已在 requirements.txt 加入 matplotlib 與 adjustText")
    else:
        if selected_piles:
            st.sidebar.success(f"🎯 已鎖定框選範圍 ({len(selected_piles)} 個點位)")
            pdf_df       = df_p[df_p["樁號"].isin(selected_piles)].copy()
            pdf_btn_text = "🔴 匯出 PDF (您框選的局部範圍)"
            is_partial   = True
        else:
            st.sidebar.info("🗺️ 目前為全區模式")
            pdf_df       = df_p.copy()
            pdf_btn_text = "🔴 匯出 PDF (全區圖)"
            is_partial   = False

        def pdf_gen(p_df: pd.DataFrame, is_partial: bool = False) -> bytes:
            """產生 PDF 進度圖，is_partial=True 時標題加上「局部圖」。"""
            font_name = setup_chinese_font()
            if font_name:
                plt.rcParams["font.family"] = font_name
            else:
                plt.rcParams["font.sans-serif"] = [
                    "Microsoft JhengHei", "PingFang TC", "SimHei", "Arial Unicode MS"
                ]
            plt.rcParams["axes.unicode_minus"] = False

            fig_pdf, ax = plt.subplots(figsize=(24, 16))

            states = ["未完成", "[已完成]"] + sorted(
                [s for s in p_df["狀態"].unique() if s not in ["未完成", "[已完成]"]]
            )
            fallback_idx = 0
            texts = []

            for state in states:
                sub_df = p_df[p_df["狀態"] == state]
                if sub_df.empty:
                    continue

                c = get_state_color(state, fallback_idx)
                if state not in STATE_COLORS:
                    fallback_idx += 1

                ax.scatter(sub_df["X"], sub_df["Y"], label=state, color=c, s=15, zorder=2)

                if state != "未完成":
                    for _, row in sub_df.iterrows():
                        texts.append(
                            ax.text(row["X"], row["Y"], row["標籤"], fontsize=8,
                                    ha="center", va="center")
                        )

            ax.margins(0.15)
            adjust_text(
                texts, ax=ax,
                expand_points=(2.5, 2.5),
                expand_text=(1.5, 1.5),
                arrowprops=dict(arrowstyle="-", color="gray", lw=0.5, alpha=0.8),
                max_iterations=300,
            )

            ax.set_aspect("equal", adjustable="datalim")
            ax.axis("off")
            today_str  = datetime.date.today().strftime("%Y-%m-%d")
            title_text = (
                f"{today_str} 施作進度回報 (局部圖)" if is_partial
                else f"{today_str} 施作進度回報"
            )
            plt.title(title_text, fontsize=24, pad=20)
            plt.legend(loc="upper right", fontsize=12)

            pdf_buf = io.BytesIO()
            plt.savefig(pdf_buf, format="pdf", bbox_inches="tight")
            plt.close(fig_pdf)
            return pdf_buf.getvalue()

        st.sidebar.download_button(
            pdf_btn_text,
            pdf_gen(pdf_df, is_partial=is_partial),
            f"Plan_{datetime.date.today()}.pdf",
            type="primary",
        )
