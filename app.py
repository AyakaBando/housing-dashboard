import io
import re
from datetime import date
from urllib.parse import urljoin

import pandas as pd
import requests
import streamlit as st
import plotly.express as px
from bs4 import BeautifulSoup

# e-Stat official dataset page for monthly housing starts by prefecture.
ESTAT_HOUSING_LIST_URL = (
    "https://www.e-stat.go.jp/stat-search/files"
    "?cycle=1&kikan=00600&layout=dataset&"
    "tclass1=000001048390&tclass2val=0&toukei=00600120&tstat=000001016966"
)
ESTAT_RENOVATION_LIST_URL = (
    "https://www.e-stat.go.jp/stat-search/files"
    "?kikan=00600&layout=dataset&result_page=1&toukei=00600900&tstat=000001031111"
)
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; HousingDashboard/1.0)"}
TIMEOUT = 30


def fetch_latest_excel_url(list_url: str, required_words: list[str] | None = None) -> tuple[str, str]:
    """Find an e-Stat Excel download link and its surrounding dataset label."""
    response = requests.get(list_url, headers=HEADERS, timeout=TIMEOUT)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    candidates = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        label = " ".join(a.get_text(" ", strip=True).split())
        if "file-download" not in href or "EXCEL" not in label.upper():
            continue
        parent = a
        surrounding = ""
        # Dataset title/date is usually in a parent container or table row.
        for _ in range(5):
            if parent.parent:
                parent = parent.parent
            surrounding = " ".join(parent.get_text(" ", strip=True).split())
            if len(surrounding) > 30:
                break
        combined = f"{label} {surrounding}"
        if required_words and not all(word in combined for word in required_words):
            continue
        url = urljoin("https://www.e-stat.go.jp", href)
        candidates.append((url, surrounding))

    if not candidates:
        raise RuntimeError(
            "e-Statの一覧ページからExcelダウンロードリンクを見つけられませんでした。"
            "e-Stat側のページ構造が変更された可能性があります。"
        )
    return candidates[0]


@st.cache_data(ttl=6 * 60 * 60, show_spinner=False)
def download_excel(url: str) -> dict[str, pd.DataFrame]:
    response = requests.get(url, headers=HEADERS, timeout=60)
    response.raise_for_status()
    content = response.content
    if not content.startswith(b"PK"):
        raise RuntimeError("e-StatからExcelファイルを取得できませんでした。")
    workbook = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    return {str(name): df for name, df in workbook.items()}


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", "", str(value)).replace("　", "")


def find_month(value):
    """Convert common Japanese year/month representations into YYYY-MM."""
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, date)):
        return pd.Timestamp(value).to_period("M").to_timestamp()
    s = str(value).strip()
    s = s.replace("年", "/").replace("月", "").replace("．", ".")
    s = s.replace("年度", "年")
    # Reiwa / Heisei era dates
    era = re.search(r"(令和|平成)\s*(\d+)\s*年\s*(\d+)\s*月", s)
    if era:
        year = int(era.group(2)) + (2018 if era.group(1) == "令和" else 1988)
        return pd.Timestamp(year=year, month=int(era.group(3)), day=1)
    m = re.search(r"((?:19|20)\d{2})\s*[/\-\.]\s*(\d{1,2})", s)
    if m and 1 <= int(m.group(2)) <= 12:
        return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=1)
    m = re.search(r"((?:19|20)\d{2})\s*年\s*(\d{1,2})", s)
    if m and 1 <= int(m.group(2)) <= 12:
        return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=1)
    return None


def parse_prefecture_timeseries(sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    Parse e-Stat's prefecture-by-month Excel without assuming a fixed sheet layout.
    The parser looks for a header row containing month-like values and prefecture labels.
    """
    pref_names = [
        "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
        "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
        "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
        "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
        "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
        "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
        "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県",
    ]
    prefectures = set(pref_names)
    results = []

    for sheet_name, raw in sheets.items():
        if raw is None or raw.empty:
            continue
        df = raw.copy().fillna("")
        max_rows = min(len(df), 35)
        for header_idx in range(max_rows):
            row = [normalize_text(v) for v in df.iloc[header_idx].tolist()]
            month_cols = {}
            for col_idx, value in enumerate(row):
                month = find_month(value)
                if month is not None:
                    month_cols[col_idx] = month
            if len(month_cols) < 3:
                continue

            # Determine likely prefecture-name column in the first several columns.
            pref_col = None
            for col_idx in range(min(6, df.shape[1])):
                vals = {normalize_text(v) for v in df.iloc[header_idx + 1:, col_idx].tolist()}
                if len(vals & prefectures) >= 5:
                    pref_col = col_idx
                    break
            if pref_col is None:
                continue

            for row_idx in range(header_idx + 1, len(df)):
                pref = normalize_text(df.iat[row_idx, pref_col])
                if pref not in prefectures:
                    continue
                for col_idx, month in month_cols.items():
                    if col_idx >= df.shape[1]:
                        continue
                    raw_value = df.iat[row_idx, col_idx]
                    if raw_value == "" or pd.isna(raw_value):
                        continue
                    try:
                        value = float(str(raw_value).replace(",", "").replace(" ", ""))
                    except (ValueError, TypeError):
                        continue
                    results.append({
                        "年月": month,
                        "地域": pref,
                        "指標": "住宅着工戸数",
                        "値": value,
                        "単位": "戸",
                        "出典": "e-Stat 建築着工統計調査",
                        "シート": sheet_name,
                    })
            if results:
                break

    if not results:
        raise RuntimeError(
            "Excelは取得できましたが、都道府県別の月次データを自動判定できませんでした。"
            "e-StatのExcel表構成が想定と異なる可能性があります。"
        )
    out = pd.DataFrame(results).drop_duplicates(subset=["年月", "地域", "指標"])
    return out.sort_values(["年月", "地域"]).reset_index(drop=True)


def groq_analysis(data: pd.DataFrame, model: str, api_key: str) -> str:
    from groq import Groq
    client = Groq(api_key=api_key)
    summary = (
        data.groupby("年月", as_index=False)["値"].sum()
        .sort_values("年月")
        .tail(36)
        .to_string(index=False)
    )
    prompt = (
        "あなたは日本の住宅市場データを分析するアナリストです。"
        "次の住宅着工戸数データに基づき、日本語で簡潔に分析してください。"
        "長期傾向、直近の変化、解釈上の注意を分け、データにない原因を断定しないでください。\n\n"
        f"{summary}"
    )
    response = client.chat.completions.create(
        model=model or "openai/gpt-oss-20b",
        messages=[
            {"role": "system", "content": "数値に忠実で、根拠のない推測を避ける分析担当者です。"},
            {"role": "user", "content": prompt},
        ],
        temperature=0.2,
    )
    return response.choices[0].message.content or "分析結果を取得できませんでした。"


st.set_page_config(page_title="住宅着工統計ダッシュボード", page_icon="📊", layout="wide")
st.title("住宅着工統計ダッシュボード")
st.caption("e-Statの公開Excelを自動取得して表示します。データ取得元：総務省統計局・国土交通省 e-Stat")

with st.sidebar:
    st.header("データ・表示設定")
    st.write("データはe-Statの公開Excelから自動取得します。")
    refresh = st.button("データを再取得", use_container_width=True)
    if refresh:
        st.cache_data.clear()
        st.rerun()

try:
    with st.spinner("e-Statから最新のExcelを取得しています..."):
        excel_url, dataset_label = fetch_latest_excel_url(
            ESTAT_HOUSING_LIST_URL, required_words=["都道府県別着工戸数"]
        )
        workbook = download_excel(excel_url)
        data = parse_prefecture_timeseries(workbook)
except Exception as exc:
    st.error("e-Statデータを読み込めませんでした。")
    st.code(str(exc))
    st.markdown(
        f"データ掲載ページ：[{ESTAT_HOUSING_LIST_URL}]({ESTAT_HOUSING_LIST_URL})"
    )
    st.stop()

data["年月"] = pd.to_datetime(data["年月"])
min_date = data["年月"].min().date()
max_date = data["年月"].max().date()

with st.sidebar:
    start_date = st.date_input("開始月", value=min_date, min_value=min_date, max_value=max_date)
    end_date = st.date_input("終了月", value=max_date, min_value=min_date, max_value=max_date)
    all_regions = sorted(data["地域"].unique().tolist())
    selected_regions = st.multiselect("都道府県", all_regions, default=all_regions)
    chart_type = st.selectbox("グラフ形式", ["折れ線", "棒", "面"])
    show_table = st.checkbox("データ表を表示", value=True)

filtered = data[
    (data["年月"].dt.date >= start_date)
    & (data["年月"].dt.date <= end_date)
    & (data["地域"].isin(selected_regions))
].copy()

if filtered.empty:
    st.warning("指定した条件に該当するデータがありません。")
    st.stop()

latest_month = filtered["年月"].max()
latest_total = filtered.loc[filtered["年月"] == latest_month, "値"].sum()
previous_months = sorted(filtered["年月"].unique())
previous_total = None
if len(previous_months) >= 2:
    previous_total = filtered.loc[filtered["年月"] == previous_months[-2], "値"].sum()

c1, c2, c3 = st.columns(3)
c1.metric("表示対象の最新月", latest_month.strftime("%Y年%m月"))
c2.metric("最新月の合計", f"{latest_total:,.0f} 戸")
if previous_total is not None:
    delta = latest_total - previous_total
    c3.metric("前月比（表示地域合計）", f"{delta:+,.0f} 戸")
else:
    c3.metric("対象都道府県数", f"{filtered['地域'].nunique()}")

st.subheader("住宅着工戸数の推移")
if chart_type == "棒":
    fig = px.bar(filtered, x="年月", y="値", color="地域", barmode="group", labels={"年月": "年月", "値": "戸数", "地域": "都道府県"})
elif chart_type == "面":
    fig = px.area(filtered, x="年月", y="値", color="地域", labels={"年月": "年月", "値": "戸数", "地域": "都道府県"})
else:
    fig = px.line(filtered, x="年月", y="値", color="地域", markers=True, labels={"年月": "年月", "値": "戸数", "地域": "都道府県"})
fig.update_layout(legend_title_text="都道府県", xaxis_title="年月", yaxis_title="住宅着工戸数（戸）")
st.plotly_chart(fig, use_container_width=True)

st.subheader("最新月の都道府県比較")
latest = filtered[filtered["年月"] == latest_month].sort_values("値", ascending=False)
bar = px.bar(latest, x="地域", y="値", labels={"地域": "都道府県", "値": "住宅着工戸数（戸）"})
bar.update_layout(xaxis_title="", yaxis_title="住宅着工戸数（戸）")
st.plotly_chart(bar, use_container_width=True)

if show_table:
    st.subheader("データ一覧")
    display = filtered.sort_values(["年月", "地域"], ascending=[False, True]).copy()
    display["年月"] = display["年月"].dt.strftime("%Y-%m")
    st.dataframe(display[["年月", "地域", "値", "単位", "出典"]], use_container_width=True, hide_index=True)

csv_data = filtered.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
st.download_button(
    "表示データをCSVでダウンロード",
    data=csv_data,
    file_name="housing_starts_estat.csv",
    mime="text/csv",
)

with st.expander("データ取得情報"):
    st.write("取得先：e-Stat「住宅着工統計 / 時系列表 / 都道府県別着工戸数」")
    st.write("ExcelダウンロードURL：", excel_url)
    st.write("検出したデータセット情報：", dataset_label)
    st.write("対象期間：", min_date, "～", max_date)
    st.caption("e-Statの公開Excelの表構成が変更された場合、読み取り処理の調整が必要になることがあります。")

st.subheader("AIによる傾向分析")
try:
    groq_key = st.secrets.get("GROQ_API_KEY", "")
    groq_model = st.secrets.get("GROQ_MODEL", "openai/gpt-oss-20b")
except Exception:
    groq_key = ""
    groq_model = "openai/gpt-oss-20b"

if not groq_key:
    st.info("AI分析を使う場合は、Streamlit Cloud の App settings → Secrets に GROQ_API_KEY を設定してください。")
else:
    if st.button("表示データをAI分析する"):
        try:
            with st.spinner("AIがデータを分析しています..."):
                st.write(groq_analysis(filtered, groq_model, groq_key))
        except Exception as exc:
            st.error(f"AI分析に失敗しました: {exc}")
