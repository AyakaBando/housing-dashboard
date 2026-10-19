import io
import re
from datetime import date, datetime
from urllib.parse import urljoin

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from bs4 import BeautifulSoup

# The e-Stat monthly timeseries list. "layout=dataset" is used by the current listing.
ESTAT_HOUSING_LIST_URL = (
    "https://www.e-stat.go.jp/stat-search/files"
    "?layout=dataset&tclass1=000001048390&toukei=00600120&tstat=000001016966"
)
PREFECTURES = [
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
    "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
    "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
    "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
    "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
    "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
    "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県",
]
PREF_SET = set(PREFECTURES)
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}
TIMEOUT = 45


def clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"\s+", "", str(value)).replace("　", "")


def fetch_excel_link() -> tuple[str, str]:
    """
    Find the Excel download link associated with the exact e-Stat table title.
    The title and download link are siblings on e-Stat, so search for their
    nearest common ancestor instead of matching the whole page or the link alone.
    """
    session = requests.Session()
    session.headers.update(HEADERS)
    response = session.get(ESTAT_HOUSING_LIST_URL, timeout=TIMEOUT)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    title_links = []
    for anchor in soup.find_all("a", href=True):
        label = clean_text(anchor.get_text(" ", strip=True))
        if "都道府県別着工戸数" in label:
            title_links.append(anchor)

    for title_anchor in title_links:
        parent = title_anchor
        # The closest ancestor that contains both the table title and its Excel link wins.
        for _ in range(9):
            parent = parent.parent
            if parent is None:
                break
            text = clean_text(parent.get_text(" ", strip=True))
            if "都道府県別着工戸数" not in text:
                continue
            excel_links = []
            for anchor in parent.find_all("a", href=True):
                href = anchor.get("href", "")
                label = clean_text(anchor.get_text(" ", strip=True)).upper()
                if "file-download" in href and ("EXCEL" in label or "XLS" in label):
                    excel_links.append(anchor)
            if len(excel_links) == 1:
                file_url = urljoin("https://www.e-stat.go.jp", excel_links[0]["href"])
                return file_url, text[:700]
            # Sometimes the immediate wrapper contains multiple links; select the
            # Excel link only if this wrapper is still a reasonably small block.
            if len(excel_links) > 1 and len(text) < 1600:
                chosen = excel_links[0]
                file_url = urljoin("https://www.e-stat.go.jp", chosen["href"])
                return file_url, text[:700]

    # Fallback: locate the title's page/statInfId, then look nearby in the HTML source
    # for the download link that has the same title block.
    raise RuntimeError(
        "e-Statのページにはアクセスできましたが、"
        "「〖住宅〗都道府県別着工戸数」に対応するExcelリンクを特定できませんでした。"
        " e-Statのページ構造が変更された可能性があります。"
    )


@st.cache_data(ttl=6 * 60 * 60, show_spinner=False)
def download_workbook(file_url: str) -> dict[str, pd.DataFrame]:
    session = requests.Session()
    session.headers.update(HEADERS)
    response = session.get(file_url, timeout=90)
    response.raise_for_status()
    body = response.content
    content_type = response.headers.get("content-type", "不明")

    # e-Stat files may be modern .xlsx (ZIP/PK signature) or legacy .xls
    # (OLE Compound File signature). Do not treat every non-PK response as an error.
    if body.startswith(b"PK\x03\x04"):
        engine = "openpyxl"
    elif body.startswith(bytes.fromhex("D0 CF 11 E0 A1 B1 1A E1")):
        engine = "xlrd"
    else:
        # Only decode a small preview when this really appears to be text/HTML.
        preview = "(空の応答)"
        if body:
            if "text" in content_type.lower() or "html" in content_type.lower():
                preview = body[:240].decode(response.encoding or "utf-8", errors="replace").replace("\n", " ")
            else:
                preview = f"先頭バイト={body[:16].hex(' ')}; サイズ={len(body)} bytes"
        raise RuntimeError(
            "e-Statのダウンロード先からExcelとして認識できるファイルが返されませんでした。"
            f" Content-Type={content_type}; {preview}"
        )

    try:
        sheets = pd.read_excel(io.BytesIO(body), sheet_name=None, header=None, engine=engine)
    except Exception as exc:
        raise RuntimeError(
            f"e-Statのファイルは取得できましたが、{engine}でExcelを読み込めませんでした: {exc}"
        ) from exc
    if not sheets:
        raise RuntimeError("Excelファイル内に読み取れるシートがありませんでした。")
    return {str(name): frame for name, frame in sheets.items()}


def parse_month(value):
    """Parse Japanese/Western year-month labels, returning the first day of the month."""
    if value is None:
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        try:
            return pd.Timestamp(value).to_period("M").to_timestamp()
        except Exception:
            return None

    s = str(value).strip()
    if not s or s.lower() in {"nan", "nat", "none"}:
        return None
    s = s.replace("　", " ").strip()

    # Japanese era formats.
    era = re.search(r"(令和|平成|昭和)\s*(\d+)\s*年?\s*[/年.\-]\s*(\d{1,2})\s*月?", s)
    if era:
        base = {"令和": 2018, "平成": 1988, "昭和": 1925}[era.group(1)]
        year, month = base + int(era.group(2)), int(era.group(3))
        if 1 <= month <= 12:
            return pd.Timestamp(year=year, month=month, day=1)

    # R8.4 / H31.4 / S63.4 patterns.
    era_short = re.search(r"\b([RHS])\s*(\d{1,2})\s*[./\-年]\s*(\d{1,2})\s*月?\b", s, re.I)
    if era_short:
        base = {"R": 2018, "H": 1988, "S": 1925}[era_short.group(1).upper()]
        year, month = base + int(era_short.group(2)), int(era_short.group(3))
        if 1 <= month <= 12:
            return pd.Timestamp(year=year, month=month, day=1)

    # YYYY年M月, YYYY/M, YYYY-M, YYYY.M, and YYYYMM.
    patterns = [
        r"((?:19|20)\d{2})\s*年\s*(\d{1,2})\s*月",
        r"((?:19|20)\d{2})\s*[/.\-]\s*(\d{1,2})(?:\D|$)",
        r"\b((?:19|20)\d{2})(\d{2})\b",
    ]
    for index, pattern in enumerate(patterns):
        match = re.search(pattern, s)
        if not match:
            continue
        year, month = int(match.group(1)), int(match.group(2))
        if index == 2 and not 1 <= month <= 12:
            continue
        if 1 <= month <= 12:
            return pd.Timestamp(year=year, month=month, day=1)
    return None


def numeric_value(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "").replace("，", "").replace(" ", "")
    if not s or s in {"-", "－", "…", "x", "X", "―"}:
        return None
    # Parentheses are commonly used in Excel for negative values.
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def normalize_pref_name(value) -> str:
    """Normalize Excel cells so full-width or invisible characters do not break matching."""
    import unicodedata
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = unicodedata.normalize("NFKC", str(value))
    for char in ("\u200b", "\u200c", "\u200d", "\ufeff", "\u2060"):
        text = text.replace(char, "")
    return re.sub(r"[\s\u3000]+", "", text)


def match_prefecture(value):
    """Return prefecture name for an exact cell or a cell with a harmless suffix/code."""
    text = normalize_pref_name(value)
    if not text:
        return None
    for pref in PREFECTURES:
        if text == normalize_pref_name(pref):
            return pref
    # A few spreadsheet exports append a prefecture code/comment in the same cell.
    for pref in PREFECTURES:
        normalized_pref = normalize_pref_name(pref)
        if text.startswith(normalized_pref):
            suffix = text[len(normalized_pref):]
            if re.fullmatch(r"[0-9０-９()（）:：._-]*", suffix):
                return pref
    return None


def get_pref_cells_in_row(df, row_idx):
    return {
        col_idx: pref
        for col_idx, value in enumerate(df.iloc[row_idx].tolist())
        if (pref := match_prefecture(value)) is not None
    }


def get_pref_cells_in_col(df, col_idx):
    return {
        row_idx: pref
        for row_idx, value in enumerate(df.iloc[:, col_idx].tolist())
        if (pref := match_prefecture(value)) is not None
    }

def month_headers_for_columns(df, max_rows=35):
    """
    Return column_index -> month for one horizontal month header.
    Also supports a two-row header where years are above month numbers.
    """
    search_rows = min(max_rows, len(df))
    best = {}

    # First, find rows that explicitly contain dates.
    for row_idx in range(search_rows):
        found = {}
        for col_idx, value in enumerate(df.iloc[row_idx].tolist()):
            month = parse_month(value)
            if month is not None:
                found[col_idx] = month
        if len(found) > len(best):
            best = found
    if len(best) >= 3:
        return best

    # Then, try a year row followed by a month row, filling merged year cells across.
    for year_row in range(search_rows - 1):
        years = []
        current_year = None
        for value in df.iloc[year_row].tolist():
            s = str(value).strip()
            western = re.search(r"((?:19|20)\d{2})", s)
            era = re.search(r"(令和|平成|昭和)\s*(\d+)", s)
            short = re.search(r"\b([RHS])\s*(\d{1,2})\b", s, re.I)
            if western:
                current_year = int(western.group(1))
            elif era:
                current_year = {"令和": 2018, "平成": 1988, "昭和": 1925}[era.group(1)] + int(era.group(2))
            elif short:
                current_year = {"R": 2018, "H": 1988, "S": 1925}[short.group(1).upper()] + int(short.group(2))
            years.append(current_year)

        for month_row in [year_row + 1, year_row + 2]:
            if month_row >= search_rows:
                continue
            found = {}
            for col_idx, value in enumerate(df.iloc[month_row].tolist()):
                s = str(value).strip()
                month_match = re.search(r"^\s*(\d{1,2})\s*月?\s*$", s)
                if not month_match:
                    continue
                month = int(month_match.group(1))
                year = years[col_idx] if col_idx < len(years) else None
                if year is not None and 1 <= month <= 12:
                    found[col_idx] = pd.Timestamp(year=year, month=month, day=1)
            if len(found) > len(best):
                best = found
    return best if len(best) >= 3 else {}


def row_month(row, inherited_year=None):
    # Prefer an explicit year-month cell in the first six columns.
    for value in row[:min(6, len(row))]:
        result = parse_month(value)
        if result is not None:
            return result

    # Support separate "year" and "month" fields in adjacent cells.
    text = " ".join(str(value) for value in row[:min(6, len(row))] if str(value).strip())
    parsed_year = None
    western = re.search(r"((?:19|20)\d{2})\s*年?", text)
    era = re.search(r"(令和|平成|昭和)\s*(\d+)", text)
    short = re.search(r"\b([RHS])\s*(\d{1,2})\b", text, re.I)
    if western:
        parsed_year = int(western.group(1))
    elif era:
        parsed_year = {"令和": 2018, "平成": 1988, "昭和": 1925}[era.group(1)] + int(era.group(2))
    elif short:
        parsed_year = {"R": 2018, "H": 1988, "S": 1925}[short.group(1).upper()] + int(short.group(2))
    year = parsed_year or inherited_year

    month = None
    for value in row[:min(6, len(row))]:
        s = str(value).strip()
        mm = re.search(r"^\s*(\d{1,2})\s*月?\s*$", s)
        if mm and 1 <= int(mm.group(1)) <= 12:
            month = int(mm.group(1))
    if year and month:
        return pd.Timestamp(year=year, month=month, day=1)
    return None


def parse_month_from_sheet(sheet_name: str, df: pd.DataFrame):
    """Get the reporting month from the report title or sheet name."""
    import unicodedata

    # The workbook's visible title normally includes e.g. 令和８年８月分.
    candidates = [str(sheet_name)]
    for row_idx in range(min(8, len(df))):
        for value in df.iloc[row_idx].tolist():
            if value is not None and str(value).strip():
                candidates.append(str(value))

    for candidate in candidates:
        text = unicodedata.normalize("NFKC", candidate).replace(" ", "")

        # Japanese era year/month: Reiwa 8 = Gregorian 2026.
        match = re.search(r"(令和|平成|昭和)\s*(元|\d+)年\s*(\d{1,2})月", text)
        if match:
            era, era_year_text, month_text = match.groups()
            era_year = 1 if era_year_text == "元" else int(era_year_text)
            base_year = {"令和": 2018, "平成": 1988, "昭和": 1925}[era]
            month = int(month_text)
            if 1 <= month <= 12:
                return pd.Timestamp(year=base_year + era_year, month=month, day=1)

        # Western year/month, should an updated workbook use that notation.
        match = re.search(r"((?:19|20)\d{2})年\s*(\d{1,2})月", text)
        if match:
            year, month = int(match.group(1)), int(match.group(2))
            if 1 <= month <= 12:
                return pd.Timestamp(year=year, month=month, day=1)

    return None


def find_total_count_column(df: pd.DataFrame, prefecture_col: int) -> int:
    """Find the primary-table column labelled 総数 / 戸数, avoiding the auxiliary table."""
    scan_rows = min(8, len(df))
    scan_cols = min(18, df.shape[1])  # the primary table is at the left of the worksheet
    for col_idx in range(scan_cols):
        has_total = any(clean_text(df.iat[row_idx, col_idx]) == "総数" for row_idx in range(scan_rows))
        has_count = any(clean_text(df.iat[row_idx, col_idx]) == "戸数" for row_idx in range(scan_rows))
        if has_total and has_count:
            return col_idx

    # In the observed e-Stat layout, total starts immediately to the right of the prefecture.
    fallback = prefecture_col + 1
    if fallback < df.shape[1]:
        return fallback
    raise RuntimeError("住宅着工戸数（総数）の列を特定できませんでした。")


def parse_prefecture_timeseries(sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Parse e-Stat workbook: one worksheet per month, prefecture in column B, total in C."""
    records = []
    sheet_debug = []

    for sheet_name, raw in sheets.items():
        if raw is None or raw.empty:
            continue
        df = raw.copy().reset_index(drop=True)
        # Normalize ordinary text whitespace; NFKC handling is done when matching prefecture names.
        df = df.apply(lambda column: column.map(lambda x: x.strip() if isinstance(x, str) else x))
        month = parse_month_from_sheet(sheet_name, df)
        if month is None:
            sheet_debug.append(f"シート「{sheet_name}」: タイトル・シート名から年月を特定できません")
            continue

        # Search the left-side main table. Do not require 10+ matches: some valid sheets
        # may contain fewer rows, and a valid prefecture cell is sufficient evidence.
        prefecture_candidates = []
        for col_idx in range(min(8, df.shape[1])):
            hits = get_pref_cells_in_col(df, col_idx)
            if hits:
                prefecture_candidates.append((col_idx, hits))

        if not prefecture_candidates:
            # The observed e-Stat table places prefecture names in column B (index 1).
            # Explicitly retry this known column before declaring extraction failure.
            if df.shape[1] > 1:
                fallback_hits = get_pref_cells_in_col(df, 1)
                if fallback_hits:
                    prefecture_candidates.append((1, fallback_hits))

        if not prefecture_candidates:
            # Diagnostic: show what Excel actually contains in the candidate columns.
            samples = []
            for col_idx in range(min(4, df.shape[1])):
                values = [normalize_pref_name(v) for v in df.iloc[:, col_idx].tolist() if normalize_pref_name(v)]
                samples.append(f"列{col_idx + 1}={values[:12]}")
            sheet_debug.append(
                f"シート「{sheet_name}」: 都道府県名の列を特定できません（" + "; ".join(samples) + "）"
            )
            continue

        # Prefer the column with the greatest number of distinct prefectures; in ties,
        # prefer the leftmost column, which is the main data table, not the auxiliary table.
        prefecture_col, prefecture_rows = max(
            prefecture_candidates,
            key=lambda item: (len(set(item[1].values())), -item[0]),
        )

        try:
            total_col = find_total_count_column(df, prefecture_col)
        except RuntimeError as exc:
            # In the known e-Stat sheet, column C (index 2) is the left-side 総数 / 戸数.
            if df.shape[1] > prefecture_col + 1:
                total_col = prefecture_col + 1
            else:
                sheet_debug.append(f"シート「{sheet_name}」: {exc}")
                continue

        rows_added = 0
        for row_idx, prefecture in prefecture_rows.items():
            if row_idx < 5 or row_idx >= df.shape[0] or total_col >= df.shape[1]:
                continue
            value = numeric_value(df.iat[row_idx, total_col])
            if value is None:
                continue
            records.append({
                "年月": month,
                "地域": prefecture,
                "指標": "住宅着工戸数",
                "値": value,
                "単位": "戸",
                "出典": "e-Stat 建築着工統計調査",
                "シート": str(sheet_name),
            })
            rows_added += 1

        if rows_added == 0:
            sheet_debug.append(
                f"シート「{sheet_name}」: 都道府県名は検出しましたが、戸数列（列{total_col + 1}）に数値がありません"
            )

    if not records:
        previews = []
        for name, df in list(sheets.items())[:3]:
            if df is not None:
                preview = df.head(6).fillna("").astype(str).values.tolist()
                previews.append(f"シート「{name}」の先頭6行: {preview}")
        diagnostic = "\n".join(sheet_debug[:10] + previews)
        raise RuntimeError(
            "Excelファイルは取得できましたが、都道府県別の住宅着工戸数を抽出できませんでした。\n"
            + diagnostic
        )

    result = pd.DataFrame(records)
    result["年月"] = pd.to_datetime(result["年月"])
    result = result.drop_duplicates(subset=["年月", "地域", "指標"], keep="first")
    result = result[result["値"].notna()]
    result = result.sort_values(["年月", "地域"]).reset_index(drop=True)
    return result


def analyze_with_groq(data: pd.DataFrame, api_key: str, model: str) -> str:
    from groq import Groq
    client = Groq(api_key=api_key)
    monthly = (
        data.groupby("年月", as_index=False)["値"].sum()
        .sort_values("年月")
        .tail(36)
    )
    summary = monthly.to_string(index=False)
    prompt = (
        "以下は日本の住宅着工戸数データです。日本語で、長期傾向、直近の動き、"
        "解釈上の注意を簡潔に分析してください。データにない原因は断定しないでください。\n\n"
        + summary
    )
    result = client.chat.completions.create(
        model=model or "openai/gpt-oss-20b",
        messages=[
            {"role": "system", "content": "数値に忠実で、根拠のない推測を避ける分析担当者です。"},
            {"role": "user", "content": prompt},
        ],
        temperature=0.2,
    )
    return result.choices[0].message.content or "分析結果を取得できませんでした。"


st.set_page_config(page_title="住宅着工統計ダッシュボード", page_icon="📊", layout="wide")
st.title("住宅着工統計ダッシュボード")
st.caption("e-Statの公開Excelを自動取得して表示します。")

with st.sidebar:
    st.header("表示設定")
    if st.button("e-Statからデータを再取得", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    with st.spinner("e-Statの一覧ページを確認し、Excelを取得しています..."):
        excel_url, dataset_info = fetch_excel_link()
        sheets = download_workbook(excel_url)
        data = parse_prefecture_timeseries(sheets)
except Exception as exc:
    st.error("e-Statからデータを読み込めませんでした。")
    st.markdown("**エラーの詳細**")
    st.code(str(exc))
    st.markdown(f"[e-Statの住宅着工統計一覧を開く]({ESTAT_HOUSING_LIST_URL})")
    st.markdown(
        "このエラー表示をそのままコピーして共有すると、次の修正箇所を特定できます。"
    )
    st.stop()

data["年月"] = pd.to_datetime(data["年月"])
min_month = data["年月"].min()
max_month = data["年月"].max()
min_date = min_month.date()
max_date = max_month.date()

with st.sidebar:
    start_date = st.date_input("開始月", value=min_date, min_value=min_date, max_value=max_date)
    end_date = st.date_input("終了月", value=max_date, min_value=min_date, max_value=max_date)
    all_regions = sorted(data["地域"].unique().tolist())
    default_regions = all_regions if len(all_regions) <= 12 else all_regions
    regions = st.multiselect("都道府県", all_regions, default=default_regions)
    chart_type = st.selectbox("推移グラフ", ["折れ線", "棒", "面"])
    show_table = st.checkbox("データ表を表示", value=True)

filtered = data[
    (data["年月"].dt.date >= start_date)
    & (data["年月"].dt.date <= end_date)
    & (data["地域"].isin(regions))
].copy()

if filtered.empty:
    st.warning("指定した条件に該当するデータがありません。")
    st.stop()

latest_month = filtered["年月"].max()
latest_total = filtered.loc[filtered["年月"] == latest_month, "値"].sum()
months = sorted(filtered["年月"].drop_duplicates())
prev_total = None
if len(months) >= 2:
    prev_total = filtered.loc[filtered["年月"] == months[-2], "値"].sum()

c1, c2, c3 = st.columns(3)
c1.metric("表示対象の最新月", latest_month.strftime("%Y年%m月"))
c2.metric("最新月の合計", f"{latest_total:,.0f} 戸")
if prev_total is not None:
    c3.metric("前月比（選択地域合計）", f"{latest_total - prev_total:+,.0f} 戸")
else:
    c3.metric("表示地域数", f"{filtered['地域'].nunique()}")

st.subheader("住宅着工戸数の推移")
if chart_type == "棒":
    fig = px.bar(filtered, x="年月", y="値", color="地域", barmode="group",
                 labels={"年月": "年月", "値": "戸数", "地域": "都道府県"})
elif chart_type == "面":
    fig = px.area(filtered, x="年月", y="値", color="地域",
                  labels={"年月": "年月", "値": "戸数", "地域": "都道府県"})
else:
    fig = px.line(filtered, x="年月", y="値", color="地域", markers=False,
                  labels={"年月": "年月", "値": "戸数", "地域": "都道府県"})
fig.update_layout(legend_title_text="都道府県", xaxis_title="年月", yaxis_title="住宅着工戸数（戸）")
st.plotly_chart(fig, use_container_width=True)

st.subheader("最新月の都道府県比較")
latest = filtered[filtered["年月"] == latest_month].sort_values("値", ascending=False)
compare_fig = px.bar(latest, x="地域", y="値",
                     labels={"地域": "都道府県", "値": "住宅着工戸数（戸）"})
compare_fig.update_layout(xaxis_title="", yaxis_title="住宅着工戸数（戸）")
st.plotly_chart(compare_fig, use_container_width=True)

if show_table:
    st.subheader("データ一覧")
    table = filtered.sort_values(["年月", "地域"], ascending=[False, True]).copy()
    table["年月"] = table["年月"].dt.strftime("%Y-%m")
    st.dataframe(table[["年月", "地域", "値", "単位", "出典"]], use_container_width=True, hide_index=True)

csv_bytes = filtered.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
st.download_button("表示データをCSVでダウンロード", csv_bytes,
                   "housing_starts_estat.csv", "text/csv")

with st.expander("データ取得情報"):
    st.write("取得元：e-Stat「建築着工統計調査 / 住宅着工統計 / 時系列表」")
    st.write("検出した統計表：", dataset_info)
    st.write("ExcelダウンロードURL：", excel_url)
    st.write("取得期間：", min_date, "～", max_date)
    st.caption("e-Stat側のページやExcelの表構成が変更された場合、読み取り処理の調整が必要になることがあります。")

st.subheader("AIによる傾向分析")
try:
    groq_api_key = st.secrets.get("GROQ_API_KEY", "")
    groq_model = st.secrets.get("GROQ_MODEL", "openai/gpt-oss-20b")
except Exception:
    groq_api_key = ""
    groq_model = "openai/gpt-oss-20b"

if not groq_api_key:
    st.info("AI分析には、Streamlit Cloud の App settings → Secrets に GROQ_API_KEY を設定してください。")
elif st.button("表示データをAI分析する"):
    try:
        with st.spinner("AIがデータを分析しています..."):
            st.write(analyze_with_groq(filtered, groq_api_key, groq_model))
    except Exception as exc:
        st.error(f"AI分析に失敗しました: {exc}")
