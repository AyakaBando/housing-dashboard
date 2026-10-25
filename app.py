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
    """
    Return the canonical prefecture name for a cell in e-Stat Excel.

    The main table may write prefecture labels with or without the suffix 県/府/都
    (for example, 青森県 or 青森). Some related tables also include a leading
    five-digit regional code (for example, 02000青森), so support those forms too.
    """
    text = normalize_pref_name(value)
    if not text:
        return None

    # Remove common leading statistical area codes, e.g. 02000青森 or 02 青森.
    text = re.sub(r"^[0-9０-９]{1,5}", "", text)
    text = re.sub(r"^[(（]?[0-9０-９]{1,5}[)）]?", "", text)
    text = re.sub(r"^[：:_-]+", "", text)

    # Build aliases while always returning the canonical full prefecture name.
    for pref in PREFECTURES:
        canonical = normalize_pref_name(pref)
        aliases = {canonical}

        # e-Stat tables sometimes abbreviate 青森県 -> 青森, 東京都 -> 東京, etc.
        if pref.endswith(("県", "府", "都")):
            aliases.add(normalize_pref_name(pref[:-1]))

        if text in aliases:
            return pref

        # Allow only safe trailing codes/markers after a recognized name.
        for alias in aliases:
            if text.startswith(alias):
                suffix = text[len(alias):]
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
    """Parse monthly e-Stat workbook into a long regional dataset with all use/housing-type metrics."""
    records = []
    sheet_debug = []
    # The main table has paired value / year-on-year columns, documented by the sample workbook:
    # C/D=total, E/F=owner-occupied, G/H=rental, I/J=company-owned, K/L=for-sale,
    # M/N=condominiums within for-sale, O/P=detached houses within for-sale.
    metrics = [
        ("住宅着工戸数", 2, 3),
        ("持家", 4, 5),
        ("貸家", 6, 7),
        ("給与住宅", 8, 9),
        ("分譲住宅", 10, 11),
        ("分譲マンション（内数）", 12, 13),
        ("分譲一戸建て（内数）", 14, 15),
    ]

    for sheet_name, raw in sheets.items():
        if raw is None or raw.empty:
            continue
        df = raw.copy().reset_index(drop=True)
        df = df.apply(lambda column: column.map(lambda x: x.strip() if isinstance(x, str) else x))
        month = parse_month_from_sheet(sheet_name, df)
        if month is None:
            sheet_debug.append(f"シート「{sheet_name}」: タイトル・シート名から年月を特定できません")
            continue

        prefecture_candidates = []
        for col_idx in range(min(8, df.shape[1])):
            hits = get_pref_cells_in_col(df, col_idx)
            if hits:
                prefecture_candidates.append((col_idx, hits))
        if not prefecture_candidates and df.shape[1] > 1:
            fallback_hits = get_pref_cells_in_col(df, 1)
            if fallback_hits:
                prefecture_candidates.append((1, fallback_hits))
        if not prefecture_candidates:
            sheet_debug.append(f"シート「{sheet_name}」: 都道府県名の列を特定できません")
            continue

        # Prefer the left-most column with the most regional labels (the primary table).
        prefecture_col, prefecture_rows = max(
            prefecture_candidates,
            key=lambda item: (len(set(item[1].values())), -item[0]),
        )
        rows_added = 0
        for row_idx, prefecture in prefecture_rows.items():
            if row_idx < 5 or row_idx >= df.shape[0]:
                continue
            for metric_name, value_col, yoy_col in metrics:
                if value_col >= df.shape[1]:
                    continue
                value = numeric_value(df.iat[row_idx, value_col])
                yoy = numeric_value(df.iat[row_idx, yoy_col]) if yoy_col < df.shape[1] else None
                if value is None:
                    continue
                records.append({
                    "年月": month,
                    "地域": prefecture,
                    "指標": metric_name,
                    "値": value,
                    "前年比（％）": yoy,
                    "単位": "戸",
                    "頻度": "月次",
                    "出典": "e-Stat 建築着工統計調査 / 住宅・都道府県別着工戸数",
                    "シート": str(sheet_name),
                })
                rows_added += 1
        if rows_added == 0:
            sheet_debug.append(f"シート「{sheet_name}」: 都道府県行から数値を抽出できませんでした")

    if not records:
        previews = []
        for name, df in list(sheets.items())[:3]:
            if df is not None:
                preview = df.head(6).fillna("").astype(str).values.tolist()
                previews.append(f"シート「{name}」の先頭6行: {preview}")
        raise RuntimeError(
            "Excelファイルは取得できましたが、都道府県別の住宅着工戸数を抽出できませんでした。\n"
            + "\n".join(sheet_debug[:10] + previews)
        )

    result = pd.DataFrame(records)
    result["年月"] = pd.to_datetime(result["年月"])
    result = result.drop_duplicates(subset=["年月", "地域", "指標"], keep="first")
    result = result[result["値"].notna()]
    return result.sort_values(["年月", "地域", "指標"]).reset_index(drop=True)


@st.cache_data(ttl=6 * 60 * 60, show_spinner=False)
def load_estat_data() -> tuple[pd.DataFrame, str, str]:
    """Cache list-page fetch + Excel download + parsing for six hours."""
    excel_url, dataset_info = fetch_excel_link()
    sheets = download_workbook(excel_url)
    data = parse_prefecture_timeseries(sheets)
    return data, excel_url, dataset_info


def build_national_series(data: pd.DataFrame) -> pd.DataFrame:
    """Build a national aggregate from all available prefectural rows.

    Counts are summed. National year-over-year percentage is calculated from the
    national metric total, rather than summing/averaging prefectural percentages.
    """
    aggregated = (
        data.groupby(["年月", "指標"], as_index=False)
        .agg(値=("値", "sum"), 集計対象都道府県数=("地域", "nunique"))
        .sort_values(["指標", "年月"])
    )
    # Never present a partial sum as the national value.
    aggregated.loc[aggregated["集計対象都道府県数"] != len(PREFECTURES), "値"] = float("nan")
    aggregated["前年比（％）"] = aggregated.groupby("指標")["値"].transform(
        lambda series: (series / series.shift(12) - 1) * 100
    )
    aggregated["地域"] = "全国"
    aggregated["単位"] = "戸"
    aggregated["頻度"] = "月次"
    aggregated["出典"] = "e-Stat 都道府県別データから47都道府県を合計"
    aggregated["シート"] = "全国集計"
    return aggregated


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
st.caption("e-Statの公開Excelを自動取得。地域を選ぶと、その地域のデータだけを全画面に反映します。")

with st.sidebar:
    st.header("① 地域を選択")
    if st.button("e-Statからデータを再取得", use_container_width=True, help="キャッシュを削除し、次回の実行でExcelを取り直します。"):
        st.cache_data.clear()
        st.rerun()

try:
    with st.spinner("初回のみe-StatのExcelを取得・解析しています..."):
        data, excel_url, dataset_info = load_estat_data()
except Exception as exc:
    st.error("e-Statからデータを読み込めませんでした。")
    st.markdown("**エラーの詳細**")
    st.code(str(exc))
    st.markdown(f"[e-Statの住宅着工統計一覧を開く]({ESTAT_HOUSING_LIST_URL})")
    st.stop()

data["年月"] = pd.to_datetime(data["年月"])
min_month = data["年月"].min()
max_month = data["年月"].max()
min_date = min_month.date()
max_date = max_month.date()
available_regions = set(data["地域"].dropna().unique().tolist())
region_options = ["全国"] + [p for p in PREFECTURES if p in available_regions]

with st.sidebar:
    selected_region = st.selectbox(
        "全国または都道府県",
        options=region_options,
        index=None,
        placeholder="地域を選択してください",
        key="region_selector",
        help="全国を選ぶと、取得した47都道府県の値を合計します。都道府県を選ぶと、その都道府県だけを表示します。",
    )

if selected_region is None:
    st.info("まず、画面左側から「全国」または表示したい都道府県を選択してください。")
    st.stop()

if selected_region == "全国":
    region_data = build_national_series(data)
else:
    region_data = data.loc[data["地域"] == selected_region].copy()
    if region_data.empty:
        st.error(f"{selected_region}のデータを取得できていません。データ抽出状況を確認してください。")
        st.stop()

metrics = [
    "住宅着工戸数", "持家", "貸家", "給与住宅", "分譲住宅",
    "分譲マンション（内数）", "分譲一戸建て（内数）",
]
available_metrics = [m for m in metrics if m in region_data["指標"].unique()]
if not available_metrics:
    st.error("選択した地域で表示できる指標がありません。")
    st.stop()

with st.sidebar:
    st.divider()
    st.header("② 指標・期間を選択")
    selected_metric = st.selectbox("指標", available_metrics, index=0)
    display_measure = st.radio("表示内容", ["着工戸数（戸）", "前年比（％）"], horizontal=False)
    start_date = st.date_input("開始月", value=min_date, min_value=min_date, max_value=max_date)
    end_date = st.date_input("終了月", value=max_date, min_value=min_date, max_value=max_date)
    chart_type = st.selectbox("グラフ形式", ["折れ線", "棒", "面"])
    show_table = st.checkbox("データ表を表示", value=True)

if start_date > end_date:
    st.error("開始月は終了月以前の日付を指定してください。")
    st.stop()

filtered = region_data[
    (region_data["指標"] == selected_metric)
    & (region_data["年月"].dt.date >= start_date)
    & (region_data["年月"].dt.date <= end_date)
].copy().sort_values("年月")
if filtered.empty:
    st.warning("選択した地域・指標・期間に該当するデータがありません。")
    st.stop()

if selected_region == "全国":
    region_label = "全国（47都道府県合計）"
else:
    region_label = selected_region

value_col = "値" if display_measure == "着工戸数（戸）" else "前年比（％）"
filtered = filtered[filtered[value_col].notna()].copy()
if filtered.empty:
    st.warning(f"選択した期間では「{display_measure}」を表示できるデータがありません。")
    st.stop()

latest_row = filtered.iloc[-1]
latest_month = pd.Timestamp(latest_row["年月"])
latest_value = float(latest_row[value_col])
metric_label = "戸" if value_col == "値" else "%"
prev_rows = filtered[filtered["年月"] < latest_month]
month_delta = None
if not prev_rows.empty:
    month_delta = latest_value - float(prev_rows.iloc[-1][value_col])

c1, c2, c3 = st.columns(3)
c1.metric("表示対象", region_label)
c2.metric("最新月", latest_month.strftime("%Y年%m月"))
c3.metric(f"最新値（{display_measure}）", f"{latest_value:,.1f} {metric_label}" if value_col != "値" else f"{latest_value:,.0f} {metric_label}")

st.subheader(f"{region_label}｜{selected_metric}の推移")
plot_df = filtered.copy()
plot_df["グラフ値"] = plot_df[value_col]
if chart_type == "棒":
    fig = px.bar(plot_df, x="年月", y="グラフ値", labels={"年月": "年月", "グラフ値": display_measure})
elif chart_type == "面":
    fig = px.area(plot_df, x="年月", y="グラフ値", labels={"年月": "年月", "グラフ値": display_measure})
else:
    fig = px.line(plot_df, x="年月", y="グラフ値", markers=False, labels={"年月": "年月", "グラフ値": display_measure})
fig.update_layout(xaxis_title="年月", yaxis_title=display_measure)
st.plotly_chart(fig, use_container_width=True)

if value_col == "値":
    yoy_value = latest_row.get("前年比（％）")
    # For national totals the value is calculated from prefectural sums; its YoY is computed from that sum.
    if pd.notna(yoy_value):
        st.metric("前年同月比", f"{float(yoy_value):+.1f}%")
    else:
        st.caption("前年同月比は比較対象月がないため表示していません。")
else:
    st.caption("前年比は、原則としてe-Stat公表表の同月比を表示します。全国では47都道府県合計から算出しています。")

if selected_metric in {"分譲マンション（内数）", "分譲一戸建て（内数）"}:
    st.caption("※「内数」は分譲住宅の内訳です。分譲住宅の総数と加算しないでください。")

if show_table:
    st.subheader("データ一覧")
    table = filtered[["年月", "地域", "指標", "値", "前年比（％）", "単位", "頻度", "出典"]].copy()
    table["年月"] = table["年月"].dt.strftime("%Y-%m")
    st.dataframe(table.sort_values("年月", ascending=False), use_container_width=True, hide_index=True)

csv_bytes = filtered.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
st.download_button(
    "表示データをCSVでダウンロード",
    data=csv_bytes,
    file_name=f"housing_starts_{selected_region}_{selected_metric}.csv",
    mime="text/csv",
)

with st.expander("データ取得情報"):
    st.write("取得元：e-Stat「建築着工統計調査 / 住宅着工統計 / 時系列表」")
    st.write("検出した統計表：", dataset_info)
    st.write("ExcelダウンロードURL：", excel_url)
    st.write("取得期間：", min_date, "～", max_date)
    st.write("表示地域：", region_label)
    st.write("抽出された都道府県数：", data["地域"].nunique())
    st.caption("全国は取得した都道府県別データの合計値です。個別の全国公表値とは集計定義が異なる可能性があります。")
    st.caption("構造別（木造・RC造など）の都道府県別データは、別の年次統計表を追加する必要があります。現時点では、この月次表で確認できる指標を表示しています。")

st.subheader("AIによる傾向分析")
try:
    groq_api_key = st.secrets.get("GROQ_API_KEY", "")
    groq_model = st.secrets.get("GROQ_MODEL", "openai/gpt-oss-20b")
except Exception:
    groq_api_key = ""
    groq_model = "openai/gpt-oss-20b"

if not groq_api_key:
    st.info("AI分析には、Streamlit Cloud の App settings → Secrets に GROQ_API_KEY を設定してください。")
elif st.button("表示中の地域・指標をAI分析する"):
    try:
        with st.spinner("AIが選択地域のデータを分析しています..."):
            st.write(analyze_with_groq(filtered, groq_api_key, groq_model))
    except Exception as exc:
        st.error(f"AI分析に失敗しました: {exc}")
