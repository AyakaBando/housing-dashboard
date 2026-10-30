import io
import json
import re
import unicodedata
from datetime import date, datetime
from urllib.parse import urljoin

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from bs4 import BeautifulSoup

APP_VERSION = "A4-PDF v10・市区町村／都道府県／全国比較"

# The e-Stat monthly timeseries list. "layout=dataset" is used by the current listing.
ESTAT_HOUSING_LIST_URL = (
    "https://www.e-stat.go.jp/stat-search/files"
    "?layout=dataset&tclass1=000001048390&toukei=00600120&tstat=000001016966"
)
ESTAT_HOUSING_ANNUAL_URL = (
    "https://www.e-stat.go.jp/stat-search/files"
    "?cycle=7&layout=datalist&page=1&tclass1=000001048390&"
    "tclass2val=0&toukei=00600120&tstat=000001016966"
)
ESTAT_BUILDING_MONTHLY_URL = (
    "https://www.e-stat.go.jp/stat-search/files"
    "?cycle=1&layout=datalist&tclass1=000001048388&"
    "toukei=00600120&tstat=000001016965"
)
ESTAT_BUILDING_ANNUAL_URL = (
    "https://www.e-stat.go.jp/stat-search/files"
    "?cycle=8&layout=datalist&page=1&tclass1=000001048388&"
    "toukei=00600120&tstat=000001016965"
)
ESTAT_RENOVATION_URL = (
    "https://www.e-stat.go.jp/stat-search/files"
    "?layout=datalist&page=1&toukei=00600900&tstat=000001031111"
)
ESTAT_RENOVATION_TABLE_URL = "https://www.e-stat.go.jp/dbview?sid=0003360983"
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


def fetch_excel_link(
    list_url: str = ESTAT_HOUSING_LIST_URL,
    include_terms: tuple[str, ...] = ("都道府県別着工戸数",),
    exclude_terms: tuple[str, ...] = (),
) -> tuple[str, str]:
    """Find an e-Stat Excel download link under the requested table title."""
    session = requests.Session()
    session.headers.update(HEADERS)
    response = session.get(list_url, timeout=TIMEOUT)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    includes = tuple(clean_text(term) for term in include_terms if term)
    excludes = tuple(clean_text(term) for term in exclude_terms if term)
    title_anchors = []
    for anchor in soup.find_all("a", href=True):
        label = clean_text(anchor.get_text(" ", strip=True))
        if includes and all(term in label for term in includes) and not any(term in label for term in excludes):
            title_anchors.append(anchor)

    for title_anchor in title_anchors:
        parent = title_anchor
        for _ in range(10):
            parent = parent.parent
            if parent is None:
                break
            text = clean_text(parent.get_text(" ", strip=True))
            if includes and not all(term in text for term in includes):
                continue
            excel_links = []
            for anchor in parent.find_all("a", href=True):
                href = anchor.get("href", "")
                label = clean_text(anchor.get_text(" ", strip=True)).upper()
                if "file-download" in href and ("EXCEL" in label or "XLS" in label):
                    excel_links.append(anchor)
            # Prefer the nearest block with exactly one Excel link, which links the title to its file.
            if len(excel_links) == 1:
                return urljoin("https://www.e-stat.go.jp", excel_links[0]["href"]), text[:1000]
            if len(excel_links) > 1 and len(text) < 1800:
                # In these blocks the first Excel link is the closest to the table's title.
                return urljoin("https://www.e-stat.go.jp", excel_links[0]["href"]), text[:1000]

    # Fallback for pages where the title isn't itself an anchor: search compact ancestors
    # of text nodes that identify the table, then select their Excel link.
    for node in soup.find_all(string=True):
        label = clean_text(str(node))
        if not (includes and all(term in label for term in includes)):
            continue
        parent = node.parent
        for _ in range(8):
            if parent is None:
                break
            text = clean_text(parent.get_text(" ", strip=True))
            if all(term in text for term in includes) and not any(term in text for term in excludes):
                links = [
                    a for a in parent.find_all("a", href=True)
                    if "file-download" in a.get("href", "")
                    and ("EXCEL" in clean_text(a.get_text(" ", strip=True)).upper()
                         or "XLS" in clean_text(a.get_text(" ", strip=True)).upper())
                ]
                if len(links) == 1:
                    return urljoin("https://www.e-stat.go.jp", links[0]["href"]), text[:1000]
            parent = parent.parent

    raise RuntimeError(
        "e-Statの一覧から対象Excelのリンクを特定できませんでした。"
        f" 対象語={', '.join(include_terms)}。ページ構造または対象表が変更された可能性があります。"
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


# Precompiled aliases speed up scanning the many cells in multi-sheet Excel workbooks.
PREF_ALIAS_TO_CANONICAL = {}
for _pref in PREFECTURES:
    PREF_ALIAS_TO_CANONICAL[normalize_pref_name(_pref)] = _pref
    if _pref.endswith(("県", "府", "都")):
        PREF_ALIAS_TO_CANONICAL[normalize_pref_name(_pref[:-1])] = _pref
PREF_ALIAS_PATTERN = re.compile(
    r"^(" + "|".join(re.escape(alias) for alias in sorted(PREF_ALIAS_TO_CANONICAL, key=len, reverse=True))
    + r")([0-9０-９()（）:：._-]*)$"
)


def match_prefecture(value):
    """Canonicalize prefecture names in both current and older e-Stat workbook layouts."""
    text = normalize_pref_name(value)
    if not text:
        return None

    # Normalize historical Excel labels such as "01北海道", "北海道計", or "北海道（01）".
    text = re.sub(r"^[0-9０-９]{1,5}", "", text)
    text = re.sub(r"^[(（]?[0-9０-９]{1,5}[)）]?", "", text)
    text = re.sub(r"^(?:都道府県|地域|都道府県名)", "", text)
    text = re.sub(r"^[：:_-]+", "", text)
    if text in PREF_ALIAS_TO_CANONICAL:
        return PREF_ALIAS_TO_CANONICAL[text]
    match = PREF_ALIAS_PATTERN.fullmatch(text)
    if match:
        return PREF_ALIAS_TO_CANONICAL.get(match.group(1))

    # Some annual workbooks append total/area markers to the prefecture label.
    # Accept only known suffixes to avoid treating arbitrary descriptions as a region.
    allowed_suffix = r"(?:合計|計|管内|地区|地方|地域|都道府県|県内|府内|都内|道内|（計）|\(計\))"
    for alias in sorted(PREF_ALIAS_TO_CANONICAL, key=len, reverse=True):
        if not text.startswith(alias):
            continue
        suffix = text[len(alias):]
        suffix = re.sub(r"^[（(]?[0-9０-９]{1,5}[）)]?$", "", suffix)
        if not suffix or re.fullmatch(r"[0-9０-９()（）:：._-]*", suffix) or re.fullmatch(allowed_suffix, suffix):
            return PREF_ALIAS_TO_CANONICAL[alias]
    # Also accept a unique prefecture name embedded alongside a leading code/marker.
    candidates = []
    for alias in sorted(PREF_ALIAS_TO_CANONICAL, key=len, reverse=True):
        start_at = text.find(alias)
        if start_at < 0:
            continue
        prefix, suffix = text[:start_at], text[start_at + len(alias):]
        clean_prefix = re.sub(r"^[0-9０-９()（）:：._-]*(?:都道府県|地域)?", "", prefix)
        clean_suffix = re.sub(r"[0-9０-９()（）:：._-]*$", "", suffix)
        if (not clean_prefix and (not clean_suffix or re.fullmatch(allowed_suffix, clean_suffix))):
            candidates.append(PREF_ALIAS_TO_CANONICAL[alias])
    candidates = list(dict.fromkeys(candidates))
    return candidates[0] if len(candidates) == 1 else None



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


def parse_year_from_text(value):
    """Return Gregorian year from Japanese-era or Western year text."""
    import unicodedata
    text = unicodedata.normalize("NFKC", "" if value is None else str(value)).replace(" ", "")
    era = re.search(r"(令和|平成|昭和)\s*(元|\d+)年?", text)
    if era:
        era_name, year_text = era.groups()
        era_year = 1 if year_text == "元" else int(year_text)
        return {"令和": 2018, "平成": 1988, "昭和": 1925}[era_name] + era_year
    western = re.search(r"(?<!\d)((?:19|20)\d{2})年?", text)
    if western:
        year = int(western.group(1))
        if 1900 <= year <= 2100:
            return year
    return None


def _cell_is_header_noise(value):
    text = normalize_pref_name(value)
    if not text:
        return True
    if len(text) > 45:
        return True
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        return True
    if parse_month(text) is not None or parse_year_from_text(text) is not None:
        return True
    noise = (
        "対前年", "同月比", "前年同月比", "前年比", "増減率", "構成比",
        "調査年", "調査月", "県コード", "単位:", "単位：", "統計表", "表番号",
        "着工新設", "都道府県別", "時系列表", "公表資料", "調査結果",
        "（単位", "(単位", "年月", "年次", "年度計",
    )
    return any(item in text for item in noise)


def _generic_header_labels(df: pd.DataFrame, header_end: int, skip_cols=None):
    """Build descriptive multi-level column labels from merged e-Stat Excel headers."""
    skip_cols = set(skip_cols or [])
    header_end = max(0, min(int(header_end), len(df), 24))
    header = df.iloc[:header_end].copy().apply(
        lambda column: column.map(lambda x: None if clean_text(x) == "" else x)
    )
    # Horizontal fill reproduces Excel merged category headings while avoiding vertical data fill.
    header = header.ffill(axis=1).ffill(axis=0)
    labels = {}
    for col in range(df.shape[1]):
        if col in skip_cols:
            continue
        parts = []
        for row_idx in range(header_end):
            value = header.iat[row_idx, col]
            text = clean_text(value)
            if _cell_is_header_noise(text):
                continue
            if match_prefecture(text):
                continue
            # Header cells such as a lone year code are not metric names.
            if text in {"戸数", "床面積", "建築物の数", "工事費予定額", "棟数", "計", "総数"}:
                if text not in parts:
                    parts.append(text)
                continue
            if text not in parts:
                parts.append(text)
        labels[col] = "／".join(parts) if parts else f"項目（列{col + 1}）"
    return labels


def _year_by_column(df: pd.DataFrame, header_end: int) -> dict[int, int]:
    """Detect year headings in columns for annual tables with multiple years on one sheet."""
    header_end = max(1, min(int(header_end), len(df), 24))
    year_by_col = {}
    year_header_rows = []
    for row_idx in range(header_end):
        row_years = {}
        for col_idx, value in enumerate(df.iloc[row_idx].tolist()):
            year = parse_year_from_text(value)
            if year and 1900 <= year <= 2100:
                row_years[col_idx] = int(year)
        if len(set(row_years.values())) >= 2:
            year_header_rows.append((row_idx, row_years))

    # Choose the row with the most distinct time headings; extend merged year cells
    # rightward only until the next explicit year heading.
    if not year_header_rows:
        return {}
    _, markers = max(year_header_rows, key=lambda item: (len(set(item[1].values())), len(item[1])))
    current_year = None
    for col_idx in range(df.shape[1]):
        if col_idx in markers:
            current_year = markers[col_idx]
        if current_year is not None:
            year_by_col[col_idx] = current_year
    return year_by_col


def _infer_unit(sheets, default="値"):
    joined = " ".join(
        clean_text(v)
        for df in list(sheets.values())[:3]
        if df is not None
        for v in df.head(8).fillna("").values.flatten().tolist()
    )
    if "億円" in joined:
        return "億円"
    if "万円" in joined:
        return "万円"
    if "千平方メートル" in joined or "千㎡" in joined or "千平米" in joined:
        return "千㎡"
    if "平方メートル" in joined or "㎡" in joined or "平米" in joined:
        return "㎡"
    if "戸" in joined:
        return "戸"
    if "棟" in joined:
        return "棟"
    return default


def _unit_for_metric(label, default_unit, workbook_unit):
    """Choose a unit per series when one Excel contains both counts and floor area."""
    text = clean_text(label)
    if any(token in text for token in ("工事費予定額", "工事費", "建設費", "請負金額")):
        if "億円" in text or workbook_unit == "億円":
            return "億円"
        return "万円"
    if any(token in text for token in ("床面積", "面積の合計", "㎡", "平方メートル", "平米")):
        return workbook_unit if workbook_unit in {"㎡", "千㎡"} else "㎡"
    if any(token in text for token in ("戸数", "着工戸数", "住宅戸数")):
        return "戸"
    if any(token in text for token in ("建築物の数", "棟数", "着工棟数")):
        return "棟"
    if "受注高" in text or default_unit == "億円" or workbook_unit == "億円":
        return "億円"
    if default_unit not in {"", "値", None}:
        return default_unit
    return workbook_unit or "値"


def _is_legacy_structure_table(df: pd.DataFrame) -> bool:
    """Recognize the historical e-Stat multi-row header even when Excel merges cells unusually."""
    if df is None or df.empty or df.shape[1] < 8:
        return False
    head = df.head(35).fillna("")
    normalized = [normalize_pref_name(v) for v in head.values.flatten().tolist()]
    col_b = {normalize_pref_name(v) for v in head.iloc[:, 1].tolist()} if head.shape[1] > 1 else set()
    all_text = "".join(normalized)
    # Most editions have the hierarchy labels in the second column and the data measures
    # in the following header rows. Some Excel exports shift the hierarchy labels slightly,
    # so also recognize the complete set anywhere in the upper header area.
    has_structure = "構造別" in col_b or "構造別" in all_text
    has_building_form = "建て方別" in col_b or "建て方別" in all_text
    has_use_relation = "利用関係別" in col_b or "利用関係別" in all_text
    has_count = "戸数" in all_text
    has_floor_area = "床面積" in all_text and "合計" in all_text
    has_wood = "木造" in all_text
    has_prefecture_heading = any(token in all_text for token in ("都道", "府県", "都道府県"))
    return (has_structure and has_building_form and has_use_relation
            and has_count and has_floor_area and has_wood and has_prefecture_heading)


def _parse_legacy_structure_table(df, period, sheet_name, dataset_name, default_unit="戸"):
    """Parse e-Stat's historical annual housing table where prefecture names are vertically split.

    Historical sheets (e.g. 昭和63年, 平成元年) use a multi-row header:
    structure -> building form -> measure (units / total floor area). The region name
    column can be split one character at a time down the five utilization rows
    (計, 持家, 貸家, 給与住宅, 分譲住宅), so ordinary exact-cell prefecture matching fails.
    """
    if df.shape[1] < 8 or len(df) < 8:
        return None, "表の行数または列数が想定より少ない"

    normalized_rows = []
    for row_idx in range(min(30, len(df))):
        # NFKC normalization is important for old e-Stat Excel files, which may contain
        # full-width forms, ideographic spaces, or invisible formatting characters.
        vals = [normalize_pref_name(v) for v in df.iloc[row_idx].tolist()]
        normalized_rows.append((row_idx, vals, "".join(vals)))

    # Anchor on the heading labels in column B, then validate the merged headers. The title
    # row itself also contains words like 利用関係別・戸数・床面積の合計, so it must not
    # be selected as the measure row. Looking at normalized cell values avoids false misses
    # caused by full-width characters or invisible spacing in older workbooks.
    structure_row = next((idx for idx, vals, joined in normalized_rows
                          if len(vals) > 1 and vals[1] == "構造別"
                          and "木造" in joined and "総" in joined and "計" in joined), None)
    building_row = next((idx for idx, vals, joined in normalized_rows
                         if len(vals) > 1 and vals[1] == "建て方別"
                         and "一戸建" in joined and "共同住宅" in joined), None)
    measure_row = next((idx for idx, vals, joined in normalized_rows
                        if len(vals) > 1 and vals[1] == "利用関係別"
                        and "戸数" in joined and "床面積" in joined and "合計" in joined), None)
    if structure_row is None or building_row is None or measure_row is None:
        return None, "構造別・建て方別・利用関係別の複数段ヘッダーを特定できない"

    header_end = max(structure_row, building_row, measure_row)
    if not (structure_row < building_row < measure_row):
        # Some e-Stat revisions change vertical positions; still require distinct rows.
        if len({structure_row, building_row, measure_row}) < 3:
            return None, "統計表の見出し行を区別できない"

    # These headers represent merged Excel cells. Fill horizontally within each header row.
    def expanded_header_row(row_idx):
        output = [""] * df.shape[1]
        current = ""
        for col_idx in range(2, df.shape[1]):
            raw = clean_text(df.iat[row_idx, col_idx])
            if raw:
                current = raw
            output[col_idx] = current
        return output

    structure_headers = expanded_header_row(structure_row)
    building_headers = expanded_header_row(building_row)
    measure_headers = expanded_header_row(measure_row)

    header_labels = {}
    for col_idx in range(2, df.shape[1]):
        structure = structure_headers[col_idx]
        building = building_headers[col_idx]
        measure = measure_headers[col_idx]
        # Skip columns with incomplete or non-data header combinations.
        if not structure or not measure:
            continue
        parts = []
        for part in (structure, building):
            if part and part not in parts:
                parts.append(part)
        if measure:
            parts.append(measure)
        header_labels[col_idx] = "／".join(parts)

    categories = {"計", "持家", "貸家", "給与住宅", "分譲住宅"}
    def category_at(row_idx):
        return clean_text(df.iat[row_idx, 1]) if df.shape[1] > 1 else ""

    # Table begins on rows where the relation column says 計 and there are numeric cells.
    starts = []
    for row_idx in range(header_end + 1, len(df)):
        if category_at(row_idx) != "計":
            continue
        numeric_count = sum(numeric_value(df.iat[row_idx, col_idx]) is not None
                            for col_idx in range(2, df.shape[1]))
        if numeric_count >= 5:
            starts.append(row_idx)

    if not starts:
        return None, "利用関係別のデータ開始行を特定できない"

    records = []
    unresolved = []
    for group_idx, start_row in enumerate(starts):
        stop_row = starts[group_idx + 1] if group_idx + 1 < len(starts) else len(df)
        group_rows = [r for r in range(start_row, stop_row) if category_at(r) in categories]
        if not group_rows:
            continue

        # In these historical tables the prefecture label may be split vertically across rows.
        fragments = []
        for row_idx in group_rows:
            part = normalize_pref_name(df.iat[row_idx, 0]) if df.shape[1] > 0 else ""
            if part:
                fragments.append(part)
        raw_region = "".join(fragments)
        region = match_prefecture(raw_region)

        # Prefer an exact label if a later table version already stores full names.
        if region is None:
            direct = [match_prefecture(df.iat[r, 0]) for r in group_rows]
            direct = [item for item in direct if item]
            if len(set(direct)) == 1:
                region = direct[0]

        # National subtotal may be split across cells as 全国; don't add it as a prefecture.
        if raw_region in {"全国", "全國"}:
            continue
        if region is None:
            unresolved.append(f"行{start_row + 1}: 地域ラベルの分割片={fragments!r}")
            continue

        for row_idx in group_rows:
            relation = category_at(row_idx)
            for col_idx, base_label in header_labels.items():
                value = numeric_value(df.iat[row_idx, col_idx])
                if value is None:
                    continue
                measure = measure_headers[col_idx]
                # Use explicit Japanese names and avoid treating count and floor area as the same unit.
                metric_parts = [p for p in base_label.split("／") if p]
                if relation != "計":
                    metric_parts.insert(max(1, len(metric_parts) - 1), relation)
                label = "／".join(dict.fromkeys(metric_parts))
                unit = "㎡" if ("床面積" in measure or "平方メートル" in measure) else ("戸" if "戸数" in measure else default_unit)
                records.append({
                    "年月": period,
                    "地域": region,
                    "指標": label,
                    "値": value,
                    "単位": unit,
                    "頻度": "年次",
                    "出典": dataset_name,
                    "シート": str(sheet_name),
                })

    if not records:
        detail = "。".join(unresolved[:5]) if unresolved else "抽出条件に一致する数値がありません"
        return None, f"縦方向に分割された都道府県名から数値を抽出できない: {detail}"

    result = pd.DataFrame(records)
    result["年月"] = pd.to_datetime(result["年月"])
    result["値"] = pd.to_numeric(result["値"], errors="coerce")
    result = result.dropna(subset=["値"])
    result = result.drop_duplicates(subset=["年月", "地域", "指標"], keep="first")
    return result, None


def parse_regional_metric_workbook(sheets, frequency="月次", dataset_name="e-Stat統計表", default_unit="値"):
    """Parse regional e-Stat Excel tables in either prefecture-row or prefecture-column orientation.

    Returns long-form data with one record per period, prefecture, and table category. If
    the table structure cannot be identified, it raises with a readable diagnostic instead
    of silently presenting partial or mislabeled values.
    """
    records = []
    debug = []
    unit = _infer_unit(sheets, default=default_unit)

    for sheet_name, raw in sheets.items():
        if raw is None or raw.empty:
            continue
        df = raw.copy().reset_index(drop=True)
        df = df.apply(lambda column: column.map(lambda x: x.strip() if isinstance(x, str) else x))
        period = parse_month_from_sheet(sheet_name, df) if frequency == "月次" else None
        if frequency == "年次":
            year = parse_year_from_text(sheet_name)
            if year is None:
                for v in df.head(12).fillna("").values.flatten().tolist():
                    year = parse_year_from_text(v)
                    if year:
                        break
            if year:
                period = pd.Timestamp(year=int(year), month=1, day=1)
        elif period is None:
            for v in df.head(10).fillna("").values.flatten().tolist():
                period = parse_month_from_sheet(str(v), df.iloc[:10])
                if period is not None:
                    break
        if period is None:
            debug.append(f"{sheet_name}: 年月・年次を判定できない")
            continue

        # Historical annual housing structure tables split prefecture names vertically across
        # utilization rows. Parse that layout explicitly before the generic exact-label parser.
        if frequency == "年次" and _is_legacy_structure_table(df):
            # Parse the historical table before generic region scanning. In these files,
            # one prefecture name can be split vertically over the five utilisation rows:
            # e.g. 鹿 / (blank) / 児 / (blank) / 島 -> 鹿児島県.
            special_data, special_error = _parse_legacy_structure_table(
                df, period, sheet_name, dataset_name, default_unit=default_unit
            )
            if special_data is not None and not special_data.empty:
                records.extend(special_data.to_dict("records"))
            else:
                debug.append(f"{sheet_name}: 旧形式の構造・建て方表を解析できない: {special_error or '原因不明'}")
            continue

        # Find both possible table orientations.
        pref_columns = []
        # Older annual Excel files may place region labels beyond column L.
        for col_idx in range(df.shape[1]):
            cells = get_pref_cells_in_col(df, col_idx)
            if len(set(cells.values())) >= 10:
                pref_columns.append((col_idx, cells))
        pref_rows = []
        # Do not assume the heading ends within the first 45 rows; annual tables vary by era.
        for row_idx in range(df.shape[0]):
            cells = get_pref_cells_in_row(df, row_idx)
            if len(set(cells.values())) >= 10:
                pref_rows.append((row_idx, cells))

        # Choose the orientation with the most distinct prefectures. Ties favor row-oriented data.
        row_best = max(pref_columns, key=lambda item: (len(set(item[1].values())), -item[0])) if pref_columns else None
        col_best = max(pref_rows, key=lambda item: (len(set(item[1].values())), -item[0])) if pref_rows else None
        if row_best is None and col_best is None:
            preview_rows = df.head(12).fillna("").astype(str).values.tolist()
            preview_cols = df.iloc[:, :min(12, df.shape[1])].tail(12).fillna("").astype(str).values.tolist()
            debug.append(
                f"{sheet_name}: 都道府県ラベルを検出できない。先頭12行={preview_rows}; "
                f"先頭12列の末尾12行={preview_cols}"
            )
            continue
        if col_best is not None and (row_best is None or len(set(col_best[1].values())) > len(set(row_best[1].values()))):
            # Orientation B: prefecture names are column headers, categories are in rows.
            pref_header_row, pref_col_map = col_best
            pref_cols = set(pref_col_map.keys())
            first_pref_col = min(pref_cols)
            left = df.iloc[:, :first_pref_col].copy()
            left = left.apply(lambda col: col.map(lambda x: None if clean_text(x) == "" else x)).ffill(axis=0)
            added = 0
            for row_idx in range(pref_header_row + 1, len(df)):
                row_text = []
                for col_idx in range(min(first_pref_col, 14)):
                    value = clean_text(left.iat[row_idx, col_idx]) if col_idx < left.shape[1] else ""
                    if not value or _cell_is_header_noise(value) or match_prefecture(value):
                        continue
                    if value not in row_text:
                        row_text.append(value)
                label = "／".join(row_text) or f"分類（行{row_idx + 1}）"
                if re.search(r"対前年|同月比|前年比|増減率|構成比", label):
                    continue
                row = df.iloc[row_idx].tolist()
                row_year = parse_year_from_text(" ".join(str(v) for v in row[:min(first_pref_col, 8)] if v is not None))
                row_period = period
                if frequency == "年次" and row_year:
                    row_period = pd.Timestamp(year=int(row_year), month=1, day=1)
                for col_idx, pref in pref_col_map.items():
                    value = numeric_value(row[col_idx]) if col_idx < len(row) else None
                    if value is None:
                        continue
                    records.append({"年月": row_period, "地域": pref, "指標": label, "値": value,
                                    "単位": _unit_for_metric(label, default_unit, unit), "頻度": frequency, "出典": dataset_name, "シート": str(sheet_name)})
                    added += 1
            if added == 0:
                debug.append(f"{sheet_name}: 都道府県列は検出したが数値行を抽出できない")
            continue

        # Orientation A: prefecture names are row labels; categories/measures are columns.
        pref_col, pref_row_map = row_best
        first_data_row = min(pref_row_map.keys()) if pref_row_map else 0
        metadata_cols = []
        for col_idx in range(df.shape[1]):
            top_values = [clean_text(df.iat[r, col_idx]) for r in range(min(max(first_data_row, 1), 16))]
            if any(any(token in value for token in ("調査年", "調査月", "県コード")) for value in top_values):
                metadata_cols.append(col_idx)
        aux_start = min(metadata_cols) if metadata_cols else None
        skip_cols = {pref_col}
        if aux_start is not None and aux_start > pref_col:
            skip_cols.update(range(aux_start, df.shape[1]))
        labels = _generic_header_labels(df, first_data_row, skip_cols=skip_cols)
        years_by_column = _year_by_column(df, first_data_row) if frequency == "年次" else {}
        numeric_cols = []
        for col_idx in range(df.shape[1]):
            if col_idx in skip_cols:
                continue
            label = labels.get(col_idx, "")
            raw_header_cells = [clean_text(df.iat[r, col_idx]) for r in range(min(max(first_data_row, 1), 16))]
            if re.search(r"対前年|同月比|前年比|増減率|構成比", label) or any(
                re.search(r"対前年|同月比|前年比|前年同月比|増減率|構成比", value)
                for value in raw_header_cells
            ):
                continue
            numeric_count = sum(
                1 for row_idx in pref_row_map
                if row_idx < len(df) and numeric_value(df.iat[row_idx, col_idx]) is not None
            )
            if numeric_count >= min(5, max(3, len(pref_row_map) // 4)):
                numeric_cols.append(col_idx)
        added = 0
        for row_idx, pref in pref_row_map.items():
            # Ignore header-row occurrences, which the detector may see in title blocks.
            if row_idx < first_data_row:
                continue
            for col_idx in numeric_cols:
                value = numeric_value(df.iat[row_idx, col_idx])
                if value is None:
                    continue
                label = labels.get(col_idx, f"項目（列{col_idx + 1}）")
                record_period = period
                if frequency == "年次" and col_idx in years_by_column:
                    record_period = pd.Timestamp(year=years_by_column[col_idx], month=1, day=1)
                records.append({"年月": record_period, "地域": pref, "指標": label, "値": value,
                                "単位": _unit_for_metric(label, default_unit, unit), "頻度": frequency, "出典": dataset_name, "シート": str(sheet_name)})
                added += 1
        if added == 0:
            debug.append(f"{sheet_name}: 都道府県行は検出したが数値列を抽出できない")

    if not records:
        raise RuntimeError(
            "e-StatのExcelは取得できましたが、都道府県別の数値を抽出できませんでした。\n"
            + "\n".join(debug[:12])
        )
    result = pd.DataFrame(records)
    result["年月"] = pd.to_datetime(result["年月"])
    result["値"] = pd.to_numeric(result["値"], errors="coerce")
    result = result.dropna(subset=["値"]).drop_duplicates(subset=["年月", "地域", "指標"], keep="first")
    # Remove categories that are only metadata/empty headers.
    result = result[result["指標"].astype(str).str.len().between(1, 100)]
    if result.empty:
        raise RuntimeError("数値を抽出できましたが、表示できる統計項目名を特定できませんでした。")
    return result.sort_values(["年月", "地域", "指標"]).reset_index(drop=True)


@st.cache_data(ttl=24 * 60 * 60, show_spinner=False)
def load_regional_source(list_url, include_terms, exclude_terms, frequency, dataset_name, default_unit="値"):
    """Cache the complete acquisition and parsing pipeline for one secondary official table."""
    file_url, table_info = fetch_excel_link(list_url, tuple(include_terms), tuple(exclude_terms))
    sheets = download_workbook(file_url)
    data = parse_regional_metric_workbook(sheets, frequency=frequency, dataset_name=dataset_name, default_unit=default_unit)
    return data, file_url, table_info


@st.cache_data(ttl=24 * 60 * 60, show_spinner=False)
def load_renovation_api_data(app_id: str) -> tuple[pd.DataFrame, str]:
    """Fetch the e-Stat renovation reference table via its official API (requires free app ID)."""
    if not app_id or not str(app_id).strip():
        raise RuntimeError("ESTAT_APP_ID が設定されていません。")

    api_url = "https://api.e-stat.go.jp/rest/2.1/app/json/getStatsData"
    response = requests.get(
        api_url,
        params={"appId": str(app_id).strip(), "statsDataId": "0003360983", "lang": "J"},
        headers=HEADERS,
        timeout=60,
    )
    response.raise_for_status()
    payload = response.json().get("GET_STATS_DATA", {})
    result = payload.get("RESULT", {}) or {}
    if str(result.get("STATUS", "0")) != "0":
        raise RuntimeError(f"e-Stat APIエラー: {result.get('ERROR_MSG', result)}")

    statistical = payload.get("STATISTICAL_DATA", {}) or {}
    class_inf = statistical.get("CLASS_INF", {}) or {}
    class_objs = class_inf.get("CLASS_OBJ", [])
    if isinstance(class_objs, dict):
        class_objs = [class_objs]
    class_maps = {}
    class_names = {}
    class_units = {}
    for obj in class_objs:
        dim_id = str(obj.get("@id", ""))
        class_names[dim_id] = str(obj.get("@name", ""))
        classes = obj.get("CLASS", [])
        if isinstance(classes, dict):
            classes = [classes]
        mapping = {}
        units = {}
        for cls in classes:
            code = str(cls.get("@code", ""))
            mapping[code] = str(cls.get("@name", code))
            if cls.get("@unit"):
                units[code] = str(cls.get("@unit"))
        class_maps[dim_id] = mapping
        class_units[dim_id] = units

    data_inf = statistical.get("DATA_INF", {}) or {}
    values = data_inf.get("VALUE", [])
    if isinstance(values, dict):
        values = [values]
    if not values:
        raise RuntimeError("e-Stat APIから数値データが返されませんでした。統計表IDまたはAPI利用登録を確認してください。")

    # Find region/time/use dimensions by their metadata, with code-name based fallbacks.
    region_dim = next((k for k, n in class_names.items() if "施工地域" in n or "都道府県" in n or n == "地域"), None)
    time_dim = next((k for k, n in class_names.items() if "時間軸" in n or "年度" in n or "年次" in n), None)
    use_dim = next((k for k, n in class_names.items() if "用途" in n), None)
    if region_dim is None:
        region_dim = next((k for k, mp in class_maps.items() if sum(1 for v in mp.values() if v == "全国" or match_prefecture(v)) >= 10), None)
    if time_dim is None:
        time_dim = next((k for k, mp in class_maps.items() if any(parse_year_from_text(v) for v in mp.values())), None)
    if use_dim is None:
        use_dim = next((k for k, mp in class_maps.items() if any(v in {"計", "住宅", "非住宅建築物"} for v in mp.values())), None)
    if region_dim is None or time_dim is None:
        raise RuntimeError(f"e-Stat APIの分類情報から地域・時間軸を特定できません。分類={class_names}")

    records = []
    for item in values:
        value = numeric_value(item.get("$"))
        if value is None:
            continue
        attr = {str(k).lstrip("@"): str(v) for k, v in item.items() if str(k).startswith("@")}
        region_name = class_maps.get(region_dim, {}).get(attr.get(region_dim, ""), "")
        if region_name == "全国":
            region = "全国"
        else:
            region = match_prefecture(region_name)
        if region is None:
            continue
        time_name = class_maps.get(time_dim, {}).get(attr.get(time_dim, ""), "")
        year = parse_year_from_text(time_name)
        if year is None:
            continue
        if use_dim:
            metric = class_maps.get(use_dim, {}).get(attr.get(use_dim, ""), "受注高")
        else:
            metric = "受注高"
        unit = item.get("@unit", "") or "億円"
        records.append({
            "年月": pd.Timestamp(year=int(year), month=1, day=1),
            "地域": region,
            "指標": str(metric),
            "値": value,
            "単位": str(unit),
            "頻度": "年次",
            "出典": "e-Stat 建築物リフォーム・リニューアル調査（参考表2）",
            "シート": "API取得",
        })
    if not records:
        raise RuntimeError("e-Stat APIには接続できましたが、地域・年度別の受注高を抽出できませんでした。")

    result_df = pd.DataFrame(records).drop_duplicates(subset=["年月", "地域", "指標"], keep="first")
    return result_df.sort_values(["年月", "地域", "指標"]).reset_index(drop=True), ESTAT_RENOVATION_TABLE_URL


# -----------------------------------------------------------------------------
# Municipality / housing acquisition environment (e-Stat API)
# -----------------------------------------------------------------------------
ESTAT_MUNICIPAL_BUILDING_STATS_ID = "0004019380"  # 市区町村別・構造別 建築物の数/床面積
ESTAT_HOUSEHOLD_INCOME_STATS_ID = "0004021668"    # 2023住宅・土地統計調査 年間収入階級
ESTAT_HOUSING_COST_STATS_ID = "0003114524"        # 新築住宅の工事費（構造・建て方・利用関係別）
ESTAT_MUNICIPAL_BUILDING_URL = "https://www.e-stat.go.jp/stat-search/database?statdisp_id=0004019380"
ESTAT_HOUSEHOLD_INCOME_URL = "https://www.e-stat.go.jp/stat-search/database?statdisp_id=0004021668"
ESTAT_HOUSING_COST_URL = "https://www.e-stat.go.jp/stat-search/database?statdisp_id=0003114524"
TSUBO_M2 = 3.305785

# e-Stat standard area codes. Table 0004019380 is a municipality-level table;
# unlike many other tables it does not include prefecture rows in CLASS_INF.
PREFECTURE_AREA_PREFIX = {
    "北海道": "01", "青森県": "02", "岩手県": "03", "宮城県": "04", "秋田県": "05",
    "山形県": "06", "福島県": "07", "茨城県": "08", "栃木県": "09", "群馬県": "10",
    "埼玉県": "11", "千葉県": "12", "東京都": "13", "神奈川県": "14", "新潟県": "15",
    "富山県": "16", "石川県": "17", "福井県": "18", "山梨県": "19", "長野県": "20",
    "岐阜県": "21", "静岡県": "22", "愛知県": "23", "三重県": "24", "滋賀県": "25",
    "京都府": "26", "大阪府": "27", "兵庫県": "28", "奈良県": "29", "和歌山県": "30",
    "鳥取県": "31", "島根県": "32", "岡山県": "33", "広島県": "34", "山口県": "35",
    "徳島県": "36", "香川県": "37", "愛媛県": "38", "高知県": "39", "福岡県": "40",
    "佐賀県": "41", "長崎県": "42", "熊本県": "43", "大分県": "44", "宮崎県": "45",
    "鹿児島県": "46", "沖縄県": "47",
}


def _as_estat_list(value):
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


@st.cache_data(ttl=24 * 60 * 60, show_spinner=False)
def load_estat_table_metadata(app_id: str, stats_data_id: str) -> list:
    """Load and normalize e-Stat CLASS_INF metadata for a table."""
    if not app_id:
        raise RuntimeError("ESTAT_APP_ID が設定されていません。")
    url = "https://api.e-stat.go.jp/rest/2.1/app/json/getMetaInfo"
    response = requests.get(
        url,
        params={"appId": app_id, "statsDataId": stats_data_id, "lang": "J"},
        headers=HEADERS,
        timeout=TIMEOUT,
    )
    response.raise_for_status()
    payload = response.json().get("GET_META_INFO", {}) or {}
    result = payload.get("RESULT", {}) or {}
    if str(result.get("STATUS", "0")) != "0":
        raise RuntimeError(f"e-StatメタデータAPIエラー: {result.get('ERROR_MSG', result)}")
    class_inf = (payload.get("METADATA_INF", {}) or {}).get("CLASS_INF", {}) or {}
    class_objs = _as_estat_list(class_inf.get("CLASS_OBJ"))
    if not class_objs:
        raise RuntimeError(f"統計表 {stats_data_id} の分類情報を取得できませんでした。")
    dimensions = []
    for obj in class_objs:
        classes = []
        for cls in _as_estat_list(obj.get("CLASS")):
            classes.append({
                "code": str(cls.get("@code", "")),
                "name": str(cls.get("@name", cls.get("@code", ""))),
                "level": str(cls.get("@level", "")),
                "parent": str(cls.get("@parentCode", "")),
                "unit": str(cls.get("@unit", "")),
            })
        dimensions.append({
            "id": str(obj.get("@id", "")),
            "name": str(obj.get("@name", obj.get("@id", ""))),
            "classes": classes,
        })
    return dimensions


@st.cache_data(ttl=24 * 60 * 60, show_spinner=False)
def load_estat_table_values(app_id: str, stats_data_id: str, area_code: str, filters_json: str = "{}") -> list:
    """Fetch one selected area from e-Stat. Per-area caching keeps the UI responsive."""
    if not app_id:
        raise RuntimeError("ESTAT_APP_ID が設定されていません。")
    import json
    filters = json.loads(filters_json or "{}")
    params = {
        "appId": app_id,
        "statsDataId": stats_data_id,
        "lang": "J",
        "limit": 100000,
        "cdArea": str(area_code),
    }
    params.update({str(k): str(v) for k, v in filters.items() if v not in (None, "")})
    url = "https://api.e-stat.go.jp/rest/2.1/app/json/getStatsData"
    response = requests.get(url, params=params, headers=HEADERS, timeout=90)
    response.raise_for_status()
    payload = response.json().get("GET_STATS_DATA", {}) or {}
    result = payload.get("RESULT", {}) or {}
    if str(result.get("STATUS", "0")) != "0":
        raise RuntimeError(f"e-StatデータAPIエラー: {result.get('ERROR_MSG', result)}")
    statistical = payload.get("STATISTICAL_DATA", {}) or {}
    data_inf = statistical.get("DATA_INF", {}) or {}
    values = _as_estat_list(data_inf.get("VALUE"))
    if not values:
        return []
    return values


@st.cache_data(ttl=24 * 60 * 60, show_spinner=False)
def load_estat_table_values_by_area_range(
    app_id: str, stats_data_id: str, area_from: str, area_to: str, filters_json: str = "{}"
) -> list:
    """Fetch a range of standard area codes, paging if the API returns NEXT_KEY."""
    if not app_id:
        raise RuntimeError("ESTAT_APP_ID が設定されていません。")
    import json
    filters = json.loads(filters_json or "{}")
    start_position = 1
    all_values = []
    for _ in range(5):
        params = {
            "appId": app_id, "statsDataId": stats_data_id, "lang": "J",
            "limit": 100000, "startPosition": start_position,
            "cdAreaFrom": str(area_from), "cdAreaTo": str(area_to),
        }
        params.update({str(k): str(v) for k, v in filters.items() if v not in (None, "")})
        response = requests.get(
            "https://api.e-stat.go.jp/rest/2.1/app/json/getStatsData",
            params=params, headers=HEADERS, timeout=90,
        )
        response.raise_for_status()
        payload = response.json().get("GET_STATS_DATA", {}) or {}
        result = payload.get("RESULT", {}) or {}
        if str(result.get("STATUS", "0")) != "0":
            raise RuntimeError(f"e-Stat地域コード範囲取得エラー: {result.get('ERROR_MSG', result)}")
        statistical = payload.get("STATISTICAL_DATA", {}) or {}
        data_inf = statistical.get("DATA_INF", {}) or {}
        all_values.extend(_as_estat_list(data_inf.get("VALUE")))
        result_inf = statistical.get("RESULT_INF", {}) or {}
        next_key = result_inf.get("NEXT_KEY")
        if not next_key:
            break
        next_position = int(next_key)
        if next_position <= start_position:
            break
        start_position = next_position
    return all_values


def _estat_dim(dimensions, *tokens):
    for dim in dimensions:
        name = clean_text(dim.get("name", ""))
        if all(token in name for token in tokens):
            return dim
    return None


def _estat_class(dim, exact=None, contains=None):
    if not dim:
        return None
    exact_norm = clean_text(exact) if exact else None
    contains_norm = clean_text(contains) if contains else None
    classes = dim.get("classes", [])
    for cls in classes:
        name = clean_text(cls.get("name", ""))
        if exact_norm and name == exact_norm:
            return cls
    for cls in classes:
        name = clean_text(cls.get("name", ""))
        if contains_norm and contains_norm in name:
            return cls
    return None


def _estat_area_catalog(dimensions):
    area_dim = _estat_dim(dimensions, "地域") or _estat_dim(dimensions, "施工地域")
    if not area_dim:
        raise RuntimeError("統計表の分類情報に地域項目がありません。")
    return area_dim


def _find_area_code(dimensions, area_name):
    area_dim = _estat_area_catalog(dimensions)
    target = clean_text(area_name)
    for cls in area_dim.get("classes", []):
        if clean_text(cls.get("name", "")) == target:
            return cls.get("code")
    if target == "全国":
        national = _estat_class(area_dim, exact="全国") or _estat_class(area_dim, contains="全国")
        return national.get("code") if national else "00000"
    return None


def _municipality_options(dimensions, prefecture):
    """Build municipality options using the first two digits of e-Stat standard area codes.

    Table 0004019380 contains municipality/ward areas (e.g. 01100 Sapporo,
    01101 Chuo-ku) but does not publish prefecture-level CLASS entries. Therefore
    do not require a prefecture row to exist in this table.
    """
    area_dim = _estat_area_catalog(dimensions)
    prefix = PREFECTURE_AREA_PREFIX.get(prefecture)
    if not prefix:
        raise RuntimeError(f"都道府県コードの対応表に{prefecture}がありません。")

    options = []
    for cls in area_dim.get("classes", []):
        code = str(cls.get("code", "")).strip()
        name = str(cls.get("name", "")).strip()
        if len(code) != 5 or not code.startswith(prefix) or code[2:] == "000" or not name:
            continue
        # Retain actual municipality/city-ward records; exclude county and
        # subprefecture group headings which are not selectable municipalities.
        if not name.endswith(("市", "区", "町", "村")):
            continue
        if any(token in name for token in ("合計", "その他", "不詳")):
            continue
        options.append({"code": code, "name": name, "label": f"{name}（{code}）"})

    # One code may be represented by a city aggregate and its wards. Keep both
    # selectable; the selected area's published aggregate remains the display value.
    options = list({x["code"]: x for x in options}.values())
    options.sort(key=lambda x: (x["name"], x["code"]))
    if not options:
        # Show enough metadata to diagnose a genuine table-code mismatch.
        examples = [
            f"{c.get('code')}:{c.get('name')}" for c in area_dim.get("classes", [])
            if str(c.get("code", "")).startswith(prefix)
        ][:20]
        raise RuntimeError(
            f"{prefecture}（標準地域コード先頭2桁={prefix}）の市区町村コードが統計表にありません。"
            f" 地域分類の候補={examples}"
        )
    # Conventional five-digit prefecture code. This table may not contain an
    # observation at this code; prefecture values are calculated from municipality leaves.
    return f"{prefix}000", options


def _municipality_leaf_codes(area_dim, prefix=None):
    """Return lowest-level municipality/ward codes, avoiding city/ward double counts."""
    classes = area_dim.get("classes", [])
    valid = []
    for cls in classes:
        code = str(cls.get("code", "")).strip()
        name = str(cls.get("name", "")).strip()
        if len(code) != 5 or (prefix and not code.startswith(prefix)):
            continue
        if not name.endswith(("市", "区", "町", "村")):
            continue
        valid.append(cls)

    code_set = {str(c.get("code", "")) for c in valid}
    parent_codes = {
        str(c.get("parent", "")) for c in valid
        if str(c.get("parent", "")) in code_set
    }
    # Fallback for older metadata that omits @parentCode: designated cities and
    # special-ward aggregates (e.g. 01100 Sapporo) have child ward codes.
    for parent in valid:
        pcode = str(parent.get("code", ""))
        pname = str(parent.get("name", ""))
        if not pname.endswith("市"):
            continue
        if any(
            str(child.get("code", "")).startswith(pcode[:3])
            and str(child.get("code", "")) != pcode
            and str(child.get("name", "")).endswith("区")
            for child in valid
        ):
            parent_codes.add(pcode)
    return [str(c.get("code")) for c in valid if str(c.get("code")) not in parent_codes]


def _parse_estat_api_rows(dimensions, values):
    """Convert e-Stat VALUE rows to a flat frame with all classification labels."""
    maps = {
        dim["id"]: {str(c["code"]): str(c["name"]) for c in dim.get("classes", [])}
        for dim in dimensions
    }
    names = {dim["id"]: dim["name"] for dim in dimensions}
    time_dim = next((d for d in dimensions if any(x in clean_text(d["name"]) for x in ("時間軸", "年次", "年度", "調査年"))), None)
    area_dim = next((d for d in dimensions if "地域" in clean_text(d["name"])), None)
    measure_dim = next((d for d in dimensions if any(t in clean_text(d["name"]) for t in ("表章項目", "項目"))), None)
    records = []
    for item in values:
        val = numeric_value(item.get("$"))
        if val is None:
            continue
        attrs = {str(k).lstrip("@"): str(v) for k, v in item.items() if str(k).startswith("@")}
        row = {"値": float(val), "単位": str(item.get("@unit", ""))}
        for dim_id, dim_name in names.items():
            code = attrs.get(dim_id, "")
            row[dim_name] = maps.get(dim_id, {}).get(code, code)
            row[f"{dim_name}_コード"] = code
        if area_dim:
            area_code = attrs.get(area_dim["id"], "")
            row["地域コード"] = area_code
            row["地域名"] = maps.get(area_dim["id"], {}).get(area_code, area_code)
        if time_dim:
            time_code = attrs.get(time_dim["id"], "")
            time_label = maps.get(time_dim["id"], {}).get(time_code, time_code)
            row["時間"] = time_label
            parsed_month = parse_month(time_label)
            year = parsed_month.year if parsed_month is not None else parse_year_from_text(time_label)
            if year:
                row["年"] = int(year)
                row["年月"] = parsed_month if parsed_month is not None else pd.Timestamp(year=int(year), month=1, day=1)
        if measure_dim:
            measure_code = attrs.get(measure_dim["id"], "")
            row["表章項目"] = maps.get(measure_dim["id"], {}).get(measure_code, measure_code)
        records.append(row)
    return pd.DataFrame(records)


def _api_filter_for_dimension(dim, class_code):
    if not dim or not class_code:
        return {}
    dim_id = str(dim.get("id", ""))
    if dim_id == "tab":
        return {"cdTab": str(class_code)}
    if dim_id == "area":
        return {"cdArea": str(class_code)}
    if dim_id == "time":
        return {"cdTime": str(class_code)}
    if dim_id.startswith("cat"):
        return {"cd" + dim_id.capitalize(): str(class_code)}
    return {}


def load_area_table_frame(app_id, stats_data_id, dimensions, area_codes):
    records = []
    for area_code in dict.fromkeys(str(x) for x in area_codes if x):
        raw = load_estat_table_values(app_id, stats_data_id, area_code)
        if raw:
            frame = _parse_estat_api_rows(dimensions, raw)
            if not frame.empty:
                records.append(frame)
    if not records:
        return pd.DataFrame()
    combined = pd.concat(records, ignore_index=True)
    return combined.drop_duplicates().reset_index(drop=True)


def _canonical_area_name(dimensions, area_code, fallback=""):
    """Resolve an e-Stat area code to its exact published Japanese label."""
    try:
        area_dim = _estat_area_catalog(dimensions)
        for cls in area_dim.get("classes", []):
            if str(cls.get("code", "")) == str(area_code):
                return str(cls.get("name", fallback))
    except Exception:
        pass
    return fallback or str(area_code)


def _find_measure_dimension(dimensions, keywords):
    """Find a classification dimension by inspecting its class labels, not only its title."""
    for dim in dimensions:
        labels = [clean_text(x.get("name", "")) for x in dim.get("classes", [])]
        if any(any(word in label for word in keywords) for label in labels):
            return dim
    return None


def _normalise_income_label(value):
    text = unicodedata.normalize("NFKC", clean_text(value))
    return text.replace(",", "").replace("，", "").replace("〜", "～").replace("~", "～")


def make_income_summary(income_frame, area_names):
    """Return income distribution and midpoint-estimated mean (not an official mean)."""
    if income_frame is None or income_frame.empty:
        return pd.DataFrame(), pd.DataFrame()
    income_dim_col = next((c for c in income_frame.columns if "年間収入階級" in str(c)), None)
    if not income_dim_col:
        # Fall back to the dimension whose class labels look like income bands.
        for col in income_frame.columns:
            vals = income_frame[col].astype(str).head(30).tolist()
            if any("万円" in x for x in vals):
                income_dim_col = col
                break
    tenure_dim_col = next((c for c in income_frame.columns if "所有の関係" in str(c)), None)
    if not income_dim_col:
        raise RuntimeError("年収階級の分類を特定できませんでした。")
    frame = income_frame.copy()
    if "表章項目" in frame.columns:
        frame = frame[frame["表章項目"].map(clean_text).str.contains("主世帯数", na=False)]
    if tenure_dim_col:
        overall_tenure = frame[frame[tenure_dim_col].map(clean_text).str.contains("総数", na=False)]
        if not overall_tenure.empty:
            frame = overall_tenure

    # Income classes are published as ranges. The top open-ended class uses 1,750万円
    # as an explicit midpoint assumption; this is an estimate, not the published mean.
    midpoint = {
        "100万円未満": 50,
        "100～150万円未満": 125, "100～150万円": 125,
        "150～200万円未満": 175, "150～200万円": 175,
        "200～300万円未満": 250, "200～300万円": 250,
        "300～400万円未満": 350, "300～400万円": 350,
        "400～500万円未満": 450, "400～500万円": 450,
        "500～700万円未満": 600, "500～700万円": 600,
        "700～1000万円未満": 850, "700～1000万円": 850,
        "1000～1500万円未満": 1250, "1000～1500万円": 1250,
        "1500万円以上": 1750,
    }
    total = {}
    distribution_rows = []
    for area in area_names:
        area_rows = frame[frame["地域名"].map(clean_text) == clean_text(area)]
        if area_rows.empty:
            continue
        vals = {_normalise_income_label(k): float(v) for k, v in area_rows.groupby(income_dim_col)["値"].sum().items()}
        total_class = next((v for k, v in vals.items() if k == "総数"), None)
        band_values = [(label, vals[label], mid) for label, mid in midpoint.items() if label in vals]
        if not band_values:
            # Accept variants such as 100万円未満/100～150万円未満 after NFKC normalization.
            for label, count in vals.items():
                for standard, mid in midpoint.items():
                    if label == standard or label.replace("未満", "") == standard.replace("未満", ""):
                        band_values.append((standard, count, mid))
                        break
        denom = total_class if total_class and total_class > 0 else sum(x[1] for x in band_values)
        band_total = sum(x[1] for x in band_values)
        if denom <= 0 or band_total <= 0:
            continue
        estimate = sum(count * mid for _, count, mid in band_values) / band_total
        high_income = sum(count for label, count, _ in band_values if any(t in label for t in ("500～700", "700～1000", "1000～1500", "1500万円以上")))
        total[area] = {
            "推計平均世帯年収（万円）": estimate,
            "500万円以上の世帯割合（％）": high_income / denom * 100,
            "世帯数": denom,
        }
        for label, count, mid in band_values:
            distribution_rows.append({
                "地域": area, "年収階級": label, "世帯数": count,
                "構成比": count / denom * 100,
            })
    summary = pd.DataFrame([{"地域": k, **v} for k, v in total.items()])
    distribution = pd.DataFrame(distribution_rows)
    return summary, distribution


def build_municipal_dashboard_data(app_id, prefecture, municipality_code, municipality_name=None):
    """Load municipality/prefecture/national construction, income and cost indicators.

    Scope is stated per source: municipal wood-building statistics cover all uses;
    household income is an income-class estimate; construction costs are prefectural/national.
    One failed source does not stop the other two.
    """
    result = {
        "building": pd.DataFrame(),
        "income_summary": pd.DataFrame(),
        "income_distribution": pd.DataFrame(),
        "cost": pd.DataFrame(),
        "errors": {},
        "sources": {
            "building": ESTAT_MUNICIPAL_BUILDING_URL,
            "income": ESTAT_HOUSEHOLD_INCOME_URL,
            "cost": ESTAT_HOUSING_COST_URL,
        },
    }

    # 1) Municipal wood-building count and floor area, plus prefecture/national comparisons.
    building_meta = None
    city_name = municipality_name or str(municipality_code)
    pref_code = PREFECTURE_AREA_PREFIX.get(prefecture, "") + "000"
    try:
        building_meta = load_estat_table_metadata(app_id, ESTAT_MUNICIPAL_BUILDING_STATS_ID)
        building_area_dim = _estat_area_catalog(building_meta)
        city_name = _canonical_area_name(building_meta, municipality_code, city_name)

        structure_dim = next((d for d in building_meta if "構造" in clean_text(d["name"])), None)
        measure_dim = next((d for d in building_meta if any(t in clean_text(d["name"]) for t in ("表章項目", "項目"))), None)
        if not measure_dim:
            measure_dim = _find_measure_dimension(building_meta, ("建築物の数", "床面積の合計"))
        if not structure_dim or not measure_dim:
            raise RuntimeError("建築物統計の構造または表章項目を特定できませんでした。")

        wood_cls = _estat_class(structure_dim, exact="木造")
        count_cls = _estat_class(measure_dim, contains="建築物の数")
        floor_cls = _estat_class(measure_dim, contains="床面積の合計")
        if not wood_cls or not count_cls or not floor_cls:
            raise RuntimeError("木造・建築物の数・床面積の分類コードを統計表から特定できませんでした。")

        import json
        api_filters = {}
        for dim, selected_codes in ((structure_dim, [wood_cls["code"]]), (measure_dim, [count_cls["code"], floor_cls["code"]])):
            param = _api_filter_for_dimension(dim, ",".join(str(x) for x in selected_codes))
            api_filters.update(param)

        # This table contains municipality/ward areas, not prefecture aggregates.
        # Fetch all five-digit area codes once; calculate prefecture/national totals
        # from the lowest-level municipality rows, avoiding city+ward double counts.
        area_values = load_estat_table_values_by_area_range(
            app_id, ESTAT_MUNICIPAL_BUILDING_STATS_ID, "00000", "47999",
            json.dumps(api_filters, ensure_ascii=False, sort_keys=True),
        )
        building_rows = _parse_estat_api_rows(building_meta, area_values)
        if building_rows.empty:
            raise RuntimeError("市区町村別の木造建築物データが空でした。")

        structure_col = structure_dim.get("name")
        measure_col = measure_dim.get("name")
        building_rows = building_rows[
            (building_rows[structure_col].map(clean_text) == clean_text(wood_cls["name"]))
            & (building_rows[measure_col].map(clean_text).isin([clean_text(count_cls["name"]), clean_text(floor_cls["name"])]))
        ].copy()

        all_leaf_codes = set(_municipality_leaf_codes(building_area_dim))
        pref_leaf_codes = set(_municipality_leaf_codes(building_area_dim, PREFECTURE_AREA_PREFIX.get(prefecture)))
        selected_code = str(municipality_code)
        records = []

        def add_aggregated_region(region_name, codes):
            subset = building_rows[building_rows["地域コード"].astype(str).isin(codes)]
            if subset.empty:
                return
            aggregated = subset.groupby(["年", measure_col], as_index=False)["値"].sum()
            for _, row in aggregated.iterrows():
                records.append({
                    "年": int(row["年"]), "地域": region_name,
                    "地域コード": prefecture if region_name == prefecture else "00000",
                    "指標": str(row[measure_col]), "値": float(row["値"]),
                    "単位": "",
                })

        # Selected municipality/ward uses the exact published area row. If the
        # selected city aggregate is unavailable, sum its immediate child areas.
        city_rows = building_rows[building_rows["地域コード"].astype(str) == selected_code]
        if city_rows.empty:
            child_codes = [
                str(c.get("code", "")) for c in building_area_dim.get("classes", [])
                if str(c.get("parent", "")) == selected_code
            ]
            if child_codes:
                city_rows = building_rows[building_rows["地域コード"].astype(str).isin(child_codes)]
                if not city_rows.empty:
                    city_rows = city_rows.groupby(["年", measure_col], as_index=False).agg({"値": "sum", "単位": "first"})
        if not city_rows.empty:
            for _, row in city_rows.iterrows():
                records.append({
                    "年": int(row["年"]), "地域": city_name,
                    "地域コード": selected_code, "指標": str(row[measure_col]),
                    "値": float(row["値"]), "単位": str(row.get("単位", "")),
                })

        add_aggregated_region(prefecture, pref_leaf_codes)
        add_aggregated_region("全国", all_leaf_codes)

        raw_building = pd.DataFrame(records)
        if not raw_building.empty:
            # Keep the selected city, its prefecture aggregate, and national aggregate distinct.
            count = raw_building[raw_building["指標"].map(clean_text) == clean_text(count_cls["name"])][["年", "地域", "値"]].rename(columns={"値": "木造建築物数"})
            area = raw_building[raw_building["指標"].map(clean_text) == clean_text(floor_cls["name"])][["年", "地域", "値"]].rename(columns={"値": "木造建築物床面積㎡"})
            result["building"] = count.merge(area, on=["年", "地域"], how="outer")
            result["building"]["平均床面積坪"] = (
                result["building"]["木造建築物床面積㎡"]
                / result["building"]["木造建築物数"] / TSUBO_M2
            )
        if result["building"].empty:
            result["errors"]["building"] = "市区町村・都道府県・全国の木造建築物データを集計できませんでした。"
    except Exception as exc:
        result["errors"]["building"] = str(exc)

    # 2) 2023 Housing and Land Survey household annual-income classes.
    try:
        income_meta = load_estat_table_metadata(app_id, ESTAT_HOUSEHOLD_INCOME_STATS_ID)
        income_area_dim = _estat_area_catalog(income_meta)
        income_nat = _estat_class(income_area_dim, exact="全国") or _estat_class(income_area_dim, contains="全国")
        income_pref_code = _find_area_code(income_meta, prefecture)
        income_city_code = _find_area_code(income_meta, city_name)
        if not income_city_code:
            # The municipality JIS code often matches across e-Stat tables; use it only
            # when that code is actually present in this table's area classification.
            if any(str(c.get("code")) == str(municipality_code) for c in income_area_dim.get("classes", [])):
                income_city_code = str(municipality_code)
        income_nat_code = income_nat.get("code") if income_nat else "00000"
        income_codes = [x for x in [income_city_code, income_pref_code, income_nat_code] if x]
        income_rows = load_area_table_frame(app_id, ESTAT_HOUSEHOLD_INCOME_STATS_ID, income_meta, income_codes)
        available_codes = set(income_rows.get("地域コード", pd.Series(dtype=str)).astype(str))
        income_label_map = {}
        if income_city_code and str(income_city_code) in available_codes:
            income_label_map[str(income_city_code)] = city_name
        if income_pref_code and str(income_pref_code) in available_codes:
            income_label_map[str(income_pref_code)] = prefecture
        if income_nat_code and str(income_nat_code) in available_codes:
            income_label_map[str(income_nat_code)] = "全国"
        if not income_rows.empty:
            income_rows["地域名"] = income_rows["地域コード"].astype(str).map(income_label_map).fillna(income_rows["地域名"])
        income_area_names = list(income_label_map.values())
        result["income_summary"], result["income_distribution"] = make_income_summary(income_rows, income_area_names)
        if result["income_summary"].empty:
            result["errors"]["income"] = "年収階級の市区町村データを取得できませんでした。対象地域外の可能性があります。"
    except Exception as exc:
        result["errors"]["income"] = str(exc)

    # 3) Construction costs are published as prefecture/national benchmarks, not city values.
    try:
        cost_meta = load_estat_table_metadata(app_id, ESTAT_HOUSING_COST_STATS_ID)
        cost_area_dim = _estat_area_catalog(cost_meta)
        cost_nat = _estat_class(cost_area_dim, exact="全国") or _estat_class(cost_area_dim, contains="全国")
        cost_pref_code = _find_area_code(cost_meta, prefecture)
        cost_nat_code = cost_nat.get("code") if cost_nat else "00000"
        cost_codes = [x for x in [cost_pref_code, cost_nat_code] if x]
        cost_rows = load_area_table_frame(app_id, ESTAT_HOUSING_COST_STATS_ID, cost_meta, cost_codes)
        if not cost_rows.empty:
            use_dim = next((d for d in cost_meta if "利用関係" in clean_text(d["name"])), None)
            form_dim = next((d for d in cost_meta if "建て方" in clean_text(d["name"])), None)
            structure_dim = next((d for d in cost_meta if "構造" in clean_text(d["name"])), None)
            measure_dim = next((d for d in cost_meta if any(t in clean_text(d["name"]) for t in ("表章項目", "項目"))), None)
            if not measure_dim:
                measure_dim = _find_measure_dimension(cost_meta, ("工事費予定額", "床面積", "戸数"))
            for dim, target in ((use_dim, "持家"), (form_dim, "一戸建"), (structure_dim, "木造")):
                col = dim.get("name") if dim else None
                if col and col in cost_rows.columns:
                    cost_rows = cost_rows[cost_rows[col].map(lambda x: target in clean_text(x))]
            region_labels = {str(code): name for code, name in [(cost_pref_code, prefecture), (cost_nat_code, "全国")] if code}
            metric_col = measure_dim.get("name") if measure_dim else "表章項目"
            records = []
            for _, row in cost_rows.iterrows():
                year = row.get("年")
                if pd.isna(year):
                    continue
                code = str(row.get("地域コード", ""))
                metric = str(row.get(metric_col, ""))
                if not metric:
                    continue
                records.append({
                    "年": int(year),
                    "年月": row.get("年月", pd.Timestamp(year=int(year), month=1, day=1)),
                    "地域": region_labels.get(code, row.get("地域名", code)),
                    "指標": metric,
                    "値": float(row["値"]),
                    "単位": str(row.get("単位", "")),
                })
            result["cost"] = pd.DataFrame(records)
        if result["cost"].empty:
            result["errors"]["cost"] = "住宅工事費統計のAPIデータが空でした。e-Stat表34の提供状況をご確認ください。"
    except Exception as exc:
        result["errors"]["cost"] = str(exc)

    result["municipality_name"] = city_name
    return result



@st.cache_data(ttl=24 * 60 * 60, show_spinner=False)
def cached_municipal_dashboard_data(app_id, prefecture, municipality_code, municipality_name):
    return build_municipal_dashboard_data(app_id, prefecture, municipality_code, municipality_name)


def _unit_to_man_yen(value, unit):
    """Convert monetary amounts to 万円 using the published unit label."""
    u = unicodedata.normalize("NFKC", clean_text(unit)).replace("／", "/")
    v = float(value)
    if "億円" in u:
        return v * 10000
    if "百万円" in u:
        return v * 100
    if "万円" in u:
        return v
    if "千円" in u:
        return v / 10
    if "円" in u:
        return v / 10000
    return None


def _unit_to_sqm(value, unit):
    """Convert published area values to square metres."""
    u = unicodedata.normalize("NFKC", clean_text(unit))
    v = float(value)
    if "億㎡" in u or "億平方メートル" in u:
        return v * 100000000
    if "万㎡" in u or "万平方メートル" in u:
        return v * 10000
    if "千㎡" in u or "千平方メートル" in u:
        return v * 1000
    if "㎡" in u or "m2" in u.lower() or "平方メートル" in u:
        return v
    return None


def _estimate_cost_per_tsubo(row):
    metric = unicodedata.normalize("NFKC", clean_text(row.get("指標", ""))).lower()
    per_sqm_terms = (
        "平米あたり工事費予定額", "平方メートルあたり工事費予定額",
        "m2あたり工事費予定額", "㎡あたり工事費予定額", "m²あたり工事費予定額",
    )
    if not any(term in metric for term in per_sqm_terms):
        return None
    per_m2_man_yen = _unit_to_man_yen(row.get("値", 0), row.get("単位", ""))
    if per_m2_man_yen is None:
        return None
    return per_m2_man_yen * TSUBO_M2


def build_cost_benchmarks(cost):
    """Calculate comparable annual benchmarks from the official cost-table indicators.

    Monetary values are normalized to 万円 or 億円, and area to 坪／戸.
    """
    if cost is None or cost.empty:
        return pd.DataFrame()
    rows = []
    for (year, region), group in cost.groupby(["年", "地域"]):
        metrics = [
            (unicodedata.normalize("NFKC", clean_text(row["指標"])), float(row["値"]), str(row.get("単位", "")))
            for _, row in group.iterrows()
        ]

        def find_value(tokens, excludes=()):
            for label, value, unit in metrics:
                if all(t in label for t in tokens) and not any(t in label for t in excludes):
                    return value, unit
            return None, ""

        dwelling_count, count_unit = find_value(("戸数",), ("１戸", "1戸", "戸あたり"))
        total_area, area_unit = find_value(("床面積",), ("１戸", "1戸", "戸あたり"))
        total_cost, total_cost_unit = find_value(("工事費予定額",), ("１戸", "1戸", "一戸", "戸あたり", "平米あたり", "平方メートルあたり", "m2あたり", "㎡あたり"))
        per_house, per_house_unit = find_value(("戸あたり工事費予定額",))
        per_sqm_row = next((
            (v, u) for label, v, u in metrics
            if any(term in label.lower() for term in (
                "平米あたり工事費予定額", "平方メートルあたり工事費予定額",
                "m2あたり工事費予定額", "㎡あたり工事費予定額", "m²あたり工事費予定額"
            ))
        ), None)

        area_sqm = _unit_to_sqm(total_area, area_unit) if total_area is not None else None
        if dwelling_count and area_sqm is not None and dwelling_count > 0:
            rows.append({
                "年": int(year), "地域": region, "指標": "平均延床面積",
                "値": area_sqm / dwelling_count / TSUBO_M2, "単位": "坪／戸"
            })

        if total_cost is not None:
            total_man_yen = _unit_to_man_yen(total_cost, total_cost_unit)
            if total_man_yen is not None:
                rows.append({
                    "年": int(year), "地域": region, "指標": "工事費予定額（総額）",
                    "値": total_man_yen / 10000, "単位": "億円"
                })

        if per_house is not None:
            per_house_man_yen = _unit_to_man_yen(per_house, per_house_unit)
            if per_house_man_yen is not None:
                rows.append({
                    "年": int(year), "地域": region, "指標": "1戸あたり工事費予定額",
                    "値": per_house_man_yen, "単位": "万円／戸"
                })

        if per_sqm_row:
            raw_value, raw_unit = per_sqm_row
            fake = {"指標": "1m2あたり工事費予定額", "値": raw_value, "単位": raw_unit}
            per_tsubo = _estimate_cost_per_tsubo(fake)
            if per_tsubo is not None:
                rows.append({
                    "年": int(year), "地域": region, "指標": "坪単価",
                    "値": per_tsubo, "単位": "万円／坪"
                })
    return pd.DataFrame(rows)


def municipal_pdf_specs(bundle):
    specs = []
    building = bundle.get("building", pd.DataFrame())
    if not building.empty:
        latest = int(building["年"].max())
        current = building[building["年"] == latest]
        for metric, label, unit in (
            ("木造建築物数", "木造建築物数", "棟"),
            ("木造建築物床面積㎡", "木造建築物の床面積", "㎡"),
            ("平均床面積坪", "平均床面積", "坪／棟"),
        ):
            if metric in current.columns:
                subset = current[["地域", metric]].dropna().rename(columns={"地域": "指標", metric: "値"}).copy()
                if not subset.empty:
                    specs.append({"title": f"{latest}年度・{label}（市区町村／都道府県／全国）", "data": subset, "kind": "bar", "unit": unit})
    income_summary = bundle.get("income_summary", pd.DataFrame())
    if not income_summary.empty:
        for metric, title, unit in (
            ("推計平均世帯年収（万円）", "推計平均世帯年収の比較（2023年）", "万円"),
            ("500万円以上の世帯割合（％）", "世帯年収500万円以上の割合（2023年）", "%"),
        ):
            if metric in income_summary.columns:
                data = income_summary[["地域", metric]].dropna().rename(columns={"地域": "指標", metric: "値"})
                if not data.empty:
                    specs.append({"title": title, "data": data, "kind": "bar", "unit": unit})
    income = bundle.get("income_distribution", pd.DataFrame())
    if not income.empty:
        plot = income.copy()
        plot["指標"] = plot["年収階級"].astype(str) + "・" + plot["地域"].astype(str)
        specs.append({"title": "世帯年収階級別の構成比（2023年）", "data": plot[["指標", "構成比"]].rename(columns={"構成比": "値"}), "kind": "bar", "unit": "%"})
    cost_bench = build_cost_benchmarks(bundle.get("cost", pd.DataFrame()))
    if not cost_bench.empty:
        latest_year = int(cost_bench["年"].max())
        current = cost_bench[cost_bench["年"] == latest_year]
        for metric in ("平均延床面積", "坪単価", "1戸あたり工事費予定額", "工事費予定額（総額）"):
            subset = current[current["指標"] == metric]
            if not subset.empty:
                data = subset[["地域", "値"]].rename(columns={"地域": "指標"})
                specs.append({"title": f"{latest_year}年・{metric}（都道府県／全国参考値）", "data": data, "kind": "bar", "unit": str(subset["単位"].iloc[0])})
    return specs


def render_municipal_market_page(selected_region):
    st.markdown("### 市区町村・住宅取得環境")
    st.markdown("市区町村の建築動向・世帯年収と、都道府県／全国の住宅工事費を比較します。")
    st.caption("この統計表は市区町村別の地域コードを収録しているため、都道府県値・全国値は下位地域の合計から算出します。市区町村の木造建築物数・床面積は住宅以外も含む建築物全体です。")
    try:
        estat_app_id = st.secrets.get("ESTAT_APP_ID", "")
    except Exception:
        estat_app_id = ""
    if not estat_app_id:
        st.warning("この画面の自動取得にはe-Stat APIアプリケーションIDが必要です。既存のリフォーム市場と同じ `ESTAT_APP_ID` を使用します。")
        st.markdown("[e-Stat API利用登録](https://www.e-stat.go.jp/api/)")
        st.markdown(f"[市区町村別の木造建築物統計]({ESTAT_MUNICIPAL_BUILDING_URL})")
        st.markdown(f"[市区町村別の世帯年収分布]({ESTAT_HOUSEHOLD_INCOME_URL})")
        st.markdown(f"[住宅工事費統計]({ESTAT_HOUSING_COST_URL})")
        return

    try:
        with st.spinner("市区町村の選択肢をe-Statから取得しています..."):
            building_meta = load_estat_table_metadata(estat_app_id, ESTAT_MUNICIPAL_BUILDING_STATS_ID)
    except Exception as exc:
        st.error("市区町村の一覧を取得できませんでした。")
        with st.expander("エラー詳細"):
            st.code(str(exc))
        return
    default_pref = selected_region if selected_region in PREF_SET else "北海道"
    pref_index = PREFECTURES.index(default_pref)
    pref = st.selectbox("都道府県", PREFECTURES, index=pref_index, key="municipal_prefecture")
    if st.session_state.get("_municipal_prefecture_for_city") != pref:
        st.session_state.pop("municipal_city", None)
        st.session_state.pop("municipal_selection", None)
        st.session_state.pop("municipal_pdf_bytes", None)
        st.session_state.pop("municipal_pdf_signature", None)
        st.session_state["_municipal_prefecture_for_city"] = pref
    try:
        pref_code, municipalities = _municipality_options(building_meta, pref)
    except Exception as exc:
        st.error(str(exc))
        return
    if not municipalities:
        st.warning(f"{pref}の市区町村コードを取得できませんでした。")
        return
    labels = [x["label"] for x in municipalities]
    option_by_label = {x["label"]: x for x in municipalities}
    default_label = next((x["label"] for x in municipalities if x["name"].endswith("市")), labels[0])
    selected_label = st.selectbox("市区町村", labels, index=labels.index(default_label), key="municipal_city")
    city = option_by_label[selected_label]

    st.caption("市区町村の建築統計は2020年度以降、年次で表示します。費用は都道府県・全国の参考値、世帯年収は2023年調査です。")
    city_name = city["name"]
    st.session_state["municipal_selection"] = {
        "prefecture": pref, "municipality_code": city["code"], "municipality_name": city_name
    }
    try:
        with st.spinner("市区町村・都道府県・全国の統計を取得しています（初回のみ時間がかかります）..."):
            bundle = cached_municipal_dashboard_data(estat_app_id, pref, city["code"], city["name"])
        city_name = bundle.get("municipality_name") or city["name"]
        st.session_state["municipal_selection"] = {
            "prefecture": pref, "municipality_code": city["code"], "municipality_name": city_name
        }
    except Exception as exc:
        st.error("市区町村データを取得できませんでした。")
        with st.expander("エラー詳細"):
            st.code(str(exc))
        st.markdown(f"[市区町村別建築統計]({ESTAT_MUNICIPAL_BUILDING_URL})")
        st.markdown(f"[世帯年収分布]({ESTAT_HOUSEHOLD_INCOME_URL})")
        st.markdown(f"[住宅工事費統計]({ESTAT_HOUSING_COST_URL})")
        return

    building = bundle["building"]
    if not building.empty:
        latest_year = int(building["年"].max())
        latest = building[building["年"] == latest_year]
        city_row = latest[latest["地域"] == city_name]
        pref_row = latest[latest["地域"] == pref]
        national_row = latest[latest["地域"] == "全国"]
        k1, k2, k3, k4 = st.columns(4)
        def fmt_num(df, col, unit=""):
            if df.empty or col not in df.columns or pd.isna(df.iloc[0][col]):
                return "—"
            return f"{df.iloc[0][col]:,.1f}{unit}" if unit == "坪" else f"{df.iloc[0][col]:,.0f}{unit}"
        k1.metric(f"{latest_year}年度・{city_name}", fmt_num(city_row, "木造建築物数", "棟"))
        k2.metric(f"{latest_year}年度・{pref}合計", fmt_num(pref_row, "木造建築物数", "棟"))
        k3.metric(f"{latest_year}年度・全国合計", fmt_num(national_row, "木造建築物数", "棟"))
        k4.metric("市区町村・平均床面積", fmt_num(city_row, "平均床面積坪", "坪"))
        st.markdown("#### 木造建築物の数（用途全体）")
        count_chart = building.melt(id_vars=["年", "地域"], value_vars=["木造建築物数"], var_name="指標", value_name="値")
        count_fig = px.line(count_chart, x="年", y="値", color="地域", markers=True,
                            labels={"年": "年度", "値": "建築物の数（棟）", "地域": "比較地域"})
        count_fig.update_layout(height=360, margin=dict(l=8, r=10, t=18, b=8))
        count_fig.update_xaxes(tickformat="d年")
        count_fig.update_traces(mode="lines+markers")
        st.plotly_chart(count_fig, use_container_width=True, config={"displayModeBar": False})
        st.markdown("#### 木造建築物の床面積・1棟あたり平均坪数")
        area_fig = px.line(building, x="年", y="平均床面積坪", color="地域", markers=True,
                           labels={"年": "年度", "平均床面積坪": "平均床面積（坪／棟）", "地域": "比較地域"})
        area_fig.update_layout(height=340, margin=dict(l=8, r=10, t=18, b=8))
        area_fig.update_xaxes(tickformat="d年")
        st.plotly_chart(area_fig, use_container_width=True, config={"displayModeBar": False})
        st.caption("この統計は住宅に限らず、木造建築物全体（住宅・非住宅を含む）です。木造戸建住宅だけの数値ではありません。")
    else:
        st.warning("選択した市区町村・都道府県の木造建築物データは取得できませんでした。")
        if bundle.get("errors", {}).get("building"):
            with st.expander("建築物統計の取得状況"):
                st.code(bundle["errors"]["building"])

    income_summary = bundle["income_summary"]
    income_distribution = bundle["income_distribution"]
    st.markdown("#### 世帯年収の比較（2023年）")
    if not income_summary.empty:
        cols = st.columns(min(3, len(income_summary)))
        for idx, (_, row) in enumerate(income_summary.iterrows()):
            cols[idx].metric(f"{row['地域']}・推計平均年収", f"{row['推計平均世帯年収（万円）']:,.0f}万円")
        if not income_distribution.empty:
            income_fig = px.bar(income_distribution, x="年収階級", y="構成比", color="地域", barmode="group",
                                labels={"年収階級": "世帯年収階級", "構成比": "世帯構成比（％）", "地域": "比較地域"})
            income_fig.update_layout(height=390, margin=dict(l=8, r=10, t=18, b=8), xaxis_tickangle=-20)
            st.plotly_chart(income_fig, use_container_width=True, config={"displayModeBar": False})
        st.caption("平均年収は公表された平均値ではありません。収入階級の中心値を使った推計で、「1500万円以上」は1750万円と仮定しています。市区町村データは人口1万5千人以上の町村などに限られます。")
    else:
        st.info("この市区町村は2023年住宅・土地統計調査の年収階級表に掲載されていない可能性があります。市区町村の収入データは全町村を網羅していません。")

    cost = bundle["cost"]
    st.markdown("#### 木造・一戸建て・持家の工事費（都道府県・全国の参考値）")
    cost_bench = build_cost_benchmarks(cost)
    if not cost_bench.empty:
        cost_years = sorted(cost_bench["年"].dropna().astype(int).unique().tolist())
        cost_year = st.selectbox("工事費の対象年", cost_years, index=len(cost_years)-1, format_func=lambda y: f"{y}年", key="municipal_cost_year")
        cost_latest = cost_bench[cost_bench["年"] == cost_year]
        benchmark_metrics = [m for m in ("平均延床面積", "坪単価", "1戸あたり工事費予定額", "工事費予定額（総額）") if m in set(cost_latest["指標"])]
        for metric in benchmark_metrics:
            metric_rows = cost_latest[cost_latest["指標"] == metric].copy()
            if metric_rows.empty:
                continue
            unit = str(metric_rows["単位"].iloc[0])
            cost_fig = px.bar(metric_rows, x="地域", y="値", text_auto=",.1f",
                              labels={"地域": "比較地域", "値": f"{metric}（{unit}）"})
            cost_fig.update_traces(textposition="outside", cliponaxis=False)
            cost_fig.update_layout(height=300, margin=dict(l=8, r=24, t=18, b=8), showlegend=False)
            cost_fig.update_yaxes(tickformat=",.0f", exponentformat="none", showexponent="none")
            st.markdown(f"**{metric}**")
            st.plotly_chart(cost_fig, use_container_width=True, config={"displayModeBar": False})
        st.caption("工事費は市区町村別に公表されていないため、選択した都道府県と全国の参考値です。木造・一戸建て・持家に絞っています。")
    else:
        st.info("住宅工事費統計の自動取得ができませんでした。")
        if bundle.get("errors", {}).get("cost"):
            with st.expander("工事費統計の取得状況"):
                st.code(bundle["errors"]["cost"])
        st.markdown(f"[e-Statの住宅工事費統計]({ESTAT_HOUSING_COST_URL})")

    with st.expander("データ出典・算出方法"):
        st.markdown(f"- [市区町村別・構造別の建築物数／床面積]({ESTAT_MUNICIPAL_BUILDING_URL})（2020年度以降）")
        st.markdown(f"- [世帯の年間収入階級別主世帯数]({ESTAT_HOUSEHOLD_INCOME_URL})（2023年調査）")
        st.markdown(f"- [新築住宅の工事費予定額・1㎡あたり工事費]({ESTAT_HOUSING_COST_URL})")
        st.write("坪換算：1坪＝3.305785㎡。坪単価は1㎡あたり工事費予定額×3.305785で算出しています。")
        st.write("平均世帯年収は収入階級の中心値による推計です。1500万円以上の階級は1750万円として計算しています。")
        for key, message in bundle.get("errors", {}).items():
            st.caption(f"{key}：{message}")

    municipal_specs = municipal_pdf_specs(bundle)
    municipal_signature = f"{pref}|{city['code']}|{sorted(building['年'].unique().tolist()) if not building.empty else []}"
    if st.session_state.get("municipal_pdf_signature") != municipal_signature:
        st.session_state.pop("municipal_pdf_bytes", None)
        st.session_state.pop("municipal_pdf_signature", None)
    if st.button("市区町村・住宅取得環境をPDFに出力", type="primary", key="municipal_pdf_button"):
        try:
            st.session_state["municipal_pdf_bytes"] = create_multi_chart_pdf(
                "市区町村・住宅取得環境", f"{pref}・{city_name}", "公表統計の最新期間", municipal_specs)
            st.session_state["municipal_pdf_signature"] = municipal_signature
        except Exception as exc:
            st.error(f"PDFの作成に失敗しました: {exc}")
    if st.session_state.get("municipal_pdf_signature") == municipal_signature and st.session_state.get("municipal_pdf_bytes"):
        st.download_button("市区町村分析PDFをダウンロード", st.session_state["municipal_pdf_bytes"],
                           file_name=f"住宅取得環境_{pref}_{city_name}.pdf", mime="application/pdf", key="municipal_pdf_download")

def with_national_aggregate(data: pd.DataFrame) -> pd.DataFrame:
    """Aggregate 47 prefectures only; never show partial data as a national total."""
    if data.empty:
        return data.copy()
    regional = data[data["地域"].isin(PREFECTURES)].copy()
    grouped = regional.groupby(["年月", "指標", "単位", "頻度"], as_index=False).agg(
        値=("値", "sum"), 地域数=("地域", "nunique")
    )
    grouped.loc[grouped["地域数"] != 47, "値"] = float("nan")
    grouped["地域"] = "全国"
    grouped["出典"] = grouped["頻度"].map(lambda _: "e-Statの都道府県別データから47都道府県を合計")
    grouped["シート"] = "全国集計"
    return pd.concat([data, grouped.drop(columns=["地域数"])], ignore_index=True)


def choose_default_metrics(labels, max_count=6, keywords=()):
    labels = list(dict.fromkeys(str(x) for x in labels if str(x).strip()))
    selected = [label for label in labels if any(keyword in label for keyword in keywords)]
    if not selected:
        selected = labels[:max_count]
    return selected[:max_count]


def render_regional_metric_section(section_title, raw_df, source_url, selected_region, start_date, end_date, key_prefix, default_keywords=()):
    """Show a regional metric source with region/time filtering, charts and a source-specific PDF."""
    st.markdown(f'<div class="panel-title">{section_title}</div>', unsafe_allow_html=True)
    if raw_df is None or raw_df.empty:
        st.warning("この統計表のデータを読み込めませんでした。出典ページから公開表を確認できます。")
        if source_url:
            st.markdown(f"[統計表の公開ページを開く]({source_url})")
        return
    data = raw_df.copy()
    if selected_region == "全国":
        if "全国" in set(data["地域"].unique()):
            region_data = data[data["地域"] == "全国"].copy()
        else:
            region_data = with_national_aggregate(data)
            region_data = region_data[region_data["地域"] == "全国"]
    else:
        region_data = data[data["地域"] == selected_region].copy()
    region_data = region_data[(region_data["年月"].dt.date >= start_date) & (region_data["年月"].dt.date <= end_date)]
    if region_data.empty:
        st.info("選択した地域・期間に該当するデータは、この統計表にはありません。")
        if source_url:
            st.markdown(f"[統計表の公開ページを開く]({source_url})")
        return
    unit_options = sorted(region_data["単位"].dropna().astype(str).unique().tolist())
    if len(unit_options) > 1:
        selected_unit = st.selectbox(
            "表示する単位", unit_options, key=f"{key_prefix}_unit"
        )
        region_data = region_data[region_data["単位"].astype(str) == selected_unit].copy()
    metric_labels = sorted(region_data["指標"].dropna().astype(str).unique().tolist())
    selected_metrics = st.multiselect(
        "表示する項目", metric_labels,
        default=choose_default_metrics(metric_labels, keywords=default_keywords),
        key=f"{key_prefix}_metrics",
    )
    if not selected_metrics:
        st.info("表示する項目を選択してください。")
        return
    chart_df = region_data[region_data["指標"].isin(selected_metrics)].copy()
    unit_list = chart_df["単位"].dropna().astype(str).unique().tolist()
    unit_label = " / ".join(unit_list[:3]) if unit_list else "値"
    fig = px.line(chart_df, x="年月", y="値", color="指標", markers=True,
                  labels={"年月": "年月", "値": f"値（{unit_label}）", "指標": "項目"})
    fig.update_layout(height=390, margin=dict(l=5, r=10, t=12, b=5), legend_title_text="項目",
                      hovermode="x unified", xaxis_title="年月", yaxis_title=f"値（{unit_label}）")
    fig.update_xaxes(tickformat="%Y年%m月", showgrid=True, gridcolor="#e8edf3")
    fig.update_yaxes(tickformat=",.1f", showgrid=True, gridcolor="#e8edf3")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    latest = chart_df[chart_df["年月"] == chart_df["年月"].max()].sort_values("値", ascending=False)
    if not latest.empty:
        latest_fig = px.bar(latest, x="値", y="指標", orientation="h", color="指標", text="値",
                            labels={"値": f"値（{unit_label}）", "指標": "項目"})
        latest_fig.update_traces(texttemplate="%{x:,.1f}", textposition="outside", cliponaxis=False, textfont=dict(size=10))
        latest_fig.update_layout(height=max(240, min(520, 42 * len(latest))), margin=dict(l=4, r=36, t=6, b=5), showlegend=False)
        latest_fig.update_yaxes(autorange="reversed")
        st.plotly_chart(latest_fig, use_container_width=True, config={"displayModeBar": False})
    pdf_specs = [{"title": section_title, "data": chart_df, "kind": "line", "unit": unit_label}]
    if not latest.empty:
        pdf_specs.append({"title": f"{section_title}（最新期の項目比較）", "data": latest, "kind": "bar", "unit": unit_label})
    pdf_signature_key = f"{key_prefix}_pdf_signature"
    pdf_bytes_key = f"{key_prefix}_pdf_bytes"
    pdf_signature = f"{selected_region}|{start_date}|{end_date}|{'|'.join(selected_metrics)}"
    if st.session_state.get(pdf_signature_key) != pdf_signature:
        st.session_state.pop(pdf_bytes_key, None)
        st.session_state.pop(pdf_signature_key, None)
    if st.button(f"{section_title}をPDFに出力", key=f"{key_prefix}_pdf_button"):
        try:
            pdf_bytes = create_multi_chart_pdf(section_title, selected_region, f"{start_date:%Y年%m月}～{end_date:%Y年%m月}", pdf_specs)
            st.session_state[pdf_bytes_key] = pdf_bytes
            st.session_state[pdf_signature_key] = pdf_signature
        except Exception as exc:
            st.error(f"PDFの作成に失敗しました: {exc}")
    pdf_bytes = st.session_state.get(pdf_bytes_key) if st.session_state.get(pdf_signature_key) == pdf_signature else None
    if pdf_bytes:
        st.download_button(f"{section_title}のPDFをダウンロード", pdf_bytes,
                           file_name=f"{key_prefix}_{selected_region}.pdf", mime="application/pdf", key=f"{key_prefix}_pdf_download")
    with st.expander("データ出典"):
        if source_url:
            st.markdown(f"[取得したExcelを開く]({source_url})")

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
    """Analyze the selected region and continue automatically if Groq truncates output."""
    from groq import Groq

    client = Groq(api_key=api_key, timeout=120.0, max_retries=2)
    pivot = (
        data.pivot_table(index="年月", columns="指標", values="値", aggfunc="first")
        .sort_index()
        .tail(36)
    )
    summary = pivot.round(1).to_string()
    prompt = (
        "以下は選択した地域の住宅着工統計です。次の見出しを使い、日本語で具体的に分析してください。\n"
        "1. 全体傾向\n2. 指標ごとの変化\n3. 指標間の比較\n4. 解釈上の注意\n"
        "数値が示すことと推測を区別し、データにない原因は断定しないでください。"
        "住宅着工戸数・分譲住宅と、分譲マンション・分譲一戸建て（内数）を二重計上しないでください。"
        "最後まで書き切り、結論で突然途切れないようにしてください。\n\n"
        + summary
    )
    messages = [
        {"role": "system", "content": "あなたは日本の住宅着工統計を分析するアナリストです。数値に忠実で、簡潔ながら結論まで完結した回答を書きます。"},
        {"role": "user", "content": prompt},
    ]
    parts = []
    # The model API's default completion budget can be too low when reasoning tokens
    # are also counted. Request a larger budget and continue if the server reports length.
    for continuation_index in range(3):
        response = client.chat.completions.create(
            model=model or "openai/gpt-oss-20b",
            messages=messages,
            temperature=0.2,
            max_completion_tokens=4096,
            reasoning_effort="low",
        )
        choice = response.choices[0]
        content = (choice.message.content or "").strip()
        if content:
            parts.append(content)
        if choice.finish_reason != "length":
            break
        messages.append({"role": "assistant", "content": content})
        messages.append({
            "role": "user",
            "content": "前の回答が出力上限で途中終了しました。すでに書いた文章を繰り返さず、途切れた箇所の続きから回答全体を完結させてください。最後に結論を含めてください。",
        })
    final_text = "\n\n".join(parts).strip()
    if not final_text:
        raise RuntimeError("AIから分析文が返されませんでした。Groqのモデル名とAPIキーを確認してください。")
    if response.choices[0].finish_reason == "length":
        final_text += "\n\n※AIの出力上限に達したため、分析の末尾が短くなっている可能性があります。もう一度実行すると続きが改善する場合があります。"
    return final_text


def setup_matplotlib_japanese():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    try:
        import matplotlib_fontja  # noqa: F401
        plt.rcParams["font.family"] = "IPAexGothic"
    except Exception as exc:
        raise RuntimeError(
            "PDFグラフの日本語フォントを読み込めませんでした。"
            " requirements.txt に matplotlib-fontja があることを確認してください。"
        ) from exc
    plt.rcParams["axes.unicode_minus"] = False
    return plt


def _format_data_label(value, percent=False):
    """Format chart data labels compactly without scientific notation."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return ""
    if not pd.notna(v):
        return ""
    if percent:
        return f"{v:,.1f}%"
    if abs(v) >= 1000:
        return f"{v:,.0f}"
    if abs(v) >= 100:
        return f"{v:,.0f}"
    if abs(v) >= 10:
        return f"{v:,.1f}".rstrip("0").rstrip(".")
    if abs(v) >= 1:
        return f"{v:,.1f}".rstrip("0").rstrip(".")
    return f"{v:,.2f}".rstrip("0").rstrip(".")


def _sample_label_indices(n, max_labels):
    """Pick evenly spaced points, always including the last point."""
    if n <= 0:
        return []
    count = min(n, max(1, int(max_labels)))
    if count == 1:
        return [n - 1]
    values = [round(i * (n - 1) / (count - 1)) for i in range(count)]
    return list(dict.fromkeys(values))


def make_pdf_chart_image(kind, trend_df, compare_df, composition_df, figsize=None):
    """Render overview charts with numeric data labels for PDF reading."""
    import io
    import matplotlib.dates as mdates
    from matplotlib.ticker import FuncFormatter

    plt = setup_matplotlib_japanese()
    fig, ax = plt.subplots(figsize=figsize or (5.1, 4.0), dpi=180)

    if kind == "trend":
        if trend_df.empty:
            ax.text(0.5, 0.5, "表示できるデータがありません", ha="center", va="center", transform=ax.transAxes)
            ax.set_axis_off()
        else:
            x = pd.to_datetime(trend_df["年月"])
            series_specs = [
                ("値", "月別実績", "#3478f6", 8),
                ("年度累計", "年度累計", "#32a879", -11),
                ("移動年計", "移動年計（12か月）", "#ef5350", 15),
            ]
            n = len(trend_df)
            label_idx = _sample_label_indices(n, min(8, n))
            for col, label, color, offset in series_specs:
                if col not in trend_df:
                    continue
                y = pd.to_numeric(trend_df[col], errors="coerce")
                ax.plot(x, y, label=label, color=color, linewidth=1.8, marker="o", markersize=2.2)
                # Label a handful of evenly spaced reference values plus the latest month.
                for idx in label_idx:
                    if idx >= len(y) or pd.isna(y.iloc[idx]):
                        continue
                    value = float(y.iloc[idx])
                    ax.annotate(
                        _format_data_label(value), (x.iloc[idx], value),
                        xytext=(0, offset), textcoords="offset points", ha="center", va="center",
                        fontsize=6.5, color=color,
                        bbox={"boxstyle": "round,pad=0.12", "facecolor": "white", "edgecolor": "none", "alpha": 0.80},
                        clip_on=True,
                    )
            ax.set_ylabel("着工戸数（戸）")
            ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:,.0f}"))
            ax.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=4, maxticks=8))
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y年%m月"))
            ax.tick_params(axis="x", labelrotation=25, labelsize=8)
            ax.legend(loc="upper left", fontsize=8, frameon=False)
            ax.grid(True, alpha=0.22)
    elif kind == "compare":
        if compare_df.empty:
            ax.text(0.5, 0.5, "比較できるデータがありません", ha="center", va="center", transform=ax.transAxes)
            ax.set_axis_off()
        else:
            months = list(compare_df["月ラベル"])
            pos = list(range(len(months)))
            width = 0.37
            prev_values = pd.to_numeric(compare_df["前年値"], errors="coerce").fillna(0).tolist()
            latest_values = pd.to_numeric(compare_df["最新年値"], errors="coerce").fillna(0).tolist()
            prev_bars = ax.bar([p - width / 2 for p in pos], prev_values, width=width,
                               label=compare_df.attrs.get("前年ラベル", "前年"), color="#a6b1bd")
            latest_bars = ax.bar([p + width / 2 for p in pos], latest_values, width=width,
                                 label=compare_df.attrs.get("最新年ラベル", "最新年"), color="#3478f6")
            # Numeric labels on every bar; rotate slightly to avoid collisions between years.
            ax.bar_label(prev_bars, labels=[_format_data_label(v) if v else "0" for v in prev_values],
                         padding=1, fontsize=6.2, rotation=45)
            ax.bar_label(latest_bars, labels=[_format_data_label(v) if v else "0" for v in latest_values],
                         padding=1, fontsize=6.2, rotation=45)
            ax.set_xticks(pos, months)
            ax.set_ylabel("着工戸数（戸）")
            ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:,.0f}"))
            ax.legend(loc="upper left", fontsize=8, frameon=False)
            ax.grid(True, axis="y", alpha=0.22)
            ax.margins(y=0.22)
    elif kind == "composition":
        if composition_df.empty or composition_df["値"].sum() <= 0:
            ax.text(0.5, 0.5, "構成比を表示できるデータがありません", ha="center", va="center", transform=ax.transAxes)
            ax.set_axis_off()
        else:
            values = pd.to_numeric(composition_df["値"], errors="coerce").fillna(0).tolist()
            names = composition_df["表示名"].astype(str).tolist()
            labels = [f"{name}\n{value:,.0f}戸" for name, value in zip(names, values)]
            ax.pie(
                values, labels=labels, autopct="%1.1f%%",
                startangle=90, counterclock=False, pctdistance=0.72, labeldistance=1.12,
                wedgeprops={"width": 0.42, "edgecolor": "white"},
                textprops={"fontsize": 7.8},
            )
            ax.set_aspect("equal")

    fig.tight_layout(pad=1.35)
    output = io.BytesIO()
    fig.savefig(output, format="png", dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    output.seek(0)
    return output

def create_dashboard_pdf(region_label, start_date, end_date, latest_month, cards, trend_df, compare_df, composition_df):
    """Export the overview to two consistently sized A4 landscape pages for readability."""
    import io
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
    except Exception:
        pass

    # Use exactly the same physical page size as all other report pages.
    page_width, page_height = landscape(A4)
    output = io.BytesIO()
    pdf = canvas.Canvas(output, pagesize=(page_width, page_height), pageCompression=1)
    pdf.setTitle(f"住宅着工統計ダッシュボード - {region_label}")
    pdf.setAuthor("住宅着工統計ダッシュボード")
    font = "HeiseiKakuGo-W5"
    margin = 24
    navy = colors.HexColor("#263746")
    grey = colors.HexColor("#64748b")
    border = colors.HexColor("#d9e0e8")

    def draw_header(page_subtitle, page_no):
        pdf.setFillColor(navy)
        pdf.setFont(font, 16)
        pdf.drawString(margin, page_height - 28, "住宅着工統計ダッシュボード")
        pdf.setFillColor(grey)
        pdf.setFont(font, 8.5)
        pdf.drawString(
            margin, page_height - 45,
            f"地域：{region_label}　対象期間：{start_date:%Y年%m月}～{end_date:%Y年%m月}　最新月：{latest_month:%Y年%m月}"
        )
        pdf.setFillColor(navy)
        pdf.setFont(font, 10)
        pdf.drawString(margin, page_height - 65, page_subtitle)
        pdf.setFillColor(grey)
        pdf.setFont(font, 7)
        pdf.drawString(margin, 16, "出典：e-Stat 建築着工統計調査")
        pdf.drawRightString(page_width - margin, 16, f"{page_no} / 2　{APP_VERSION}")

    # Page 1: KPI cards in a 3x2 grid, then a large full-width trend chart.
    draw_header("市場概況・住宅着工戸数の推移", 1)
    card_gap_x, card_gap_y = 9, 8
    card_cols = 3
    card_width = (page_width - 2 * margin - card_gap_x * (card_cols - 1)) / card_cols
    card_height = 60
    # Two rows below the header leave space for the full-width trend chart.
    row_y = [page_height - 101 - card_height, page_height - 101 - card_height * 2 - card_gap_y]
    for i, card in enumerate(cards[:6]):
        col = i % card_cols
        row = i // card_cols
        x = margin + col * (card_width + card_gap_x)
        y = row_y[row]
        pdf.setFillColor(colors.white)
        pdf.setStrokeColor(border)
        pdf.roundRect(x, y, card_width, card_height, 4, stroke=1, fill=1)
        pdf.setFillColor(colors.HexColor(card.get("color", "#3478f6")))
        pdf.rect(x, y + card_height - 3, card_width, 3, stroke=0, fill=1)
        pdf.setFillColor(navy)
        pdf.setFont(font, 8)
        pdf.drawString(x + 8, y + 43, str(card.get("label", ""))[:45])
        pdf.setFont(font, 14)
        pdf.drawString(x + 8, y + 25, str(card.get("value", ""))[:35])
        pdf.setFillColor(grey)
        pdf.setFont(font, 7)
        pdf.drawString(x + 8, y + 9, str(card.get("previous", ""))[:55])
        pdf.setFillColor(colors.HexColor(card.get("yoy_color", "#64748b")))
        pdf.drawRightString(x + card_width - 8, y + 9, str(card.get("yoy", ""))[:22])

    trend_image = make_pdf_chart_image("trend", trend_df, compare_df, composition_df, figsize=(10.8, 3.9))
    trend_x = margin
    trend_y = 28
    trend_w = page_width - margin * 2
    trend_h = 300
    pdf.drawImage(ImageReader(trend_image), trend_x, trend_y, width=trend_w, height=trend_h,
                  preserveAspectRatio=True, anchor="c", mask="auto")
    pdf.setStrokeColor(border)
    pdf.roundRect(trend_x, trend_y, trend_w, trend_h, 3, stroke=1, fill=0)
    pdf.showPage()

    # Page 2: monthly comparison and composition are enlarged into two side-by-side panels.
    draw_header("年度比較・利用関係別構成比", 2)
    gap = 14
    panel_width = (page_width - 2 * margin - gap) / 2
    panel_y, panel_h = 52, page_height - 165
    compare_x = margin
    composition_x = margin + panel_width + gap
    pdf.setFillColor(navy)
    pdf.setFont(font, 10)
    pdf.drawString(compare_x, page_height - 91,
                   "月別比較（" + compare_df.attrs.get("前年ラベル", "前年") + "・" + compare_df.attrs.get("最新年ラベル", "最新年") + "）")
    pdf.drawString(composition_x, page_height - 91, f"利用関係別構成比（{latest_month:%Y年%m月}）")
    compare_image = make_pdf_chart_image("compare", trend_df, compare_df, composition_df, figsize=(5.4, 4.8))
    composition_image = make_pdf_chart_image("composition", trend_df, compare_df, composition_df, figsize=(5.0, 5.0))
    pdf.drawImage(ImageReader(compare_image), compare_x, panel_y, width=panel_width, height=panel_h,
                  preserveAspectRatio=True, anchor="c", mask="auto")
    pdf.drawImage(ImageReader(composition_image), composition_x, panel_y, width=panel_width, height=panel_h,
                  preserveAspectRatio=True, anchor="c", mask="auto")
    pdf.setStrokeColor(border)
    pdf.roundRect(compare_x, panel_y, panel_width, panel_h, 3, stroke=1, fill=0)
    pdf.roundRect(composition_x, panel_y, panel_width, panel_h, 3, stroke=1, fill=0)
    pdf.showPage()
    pdf.save()
    return output.getvalue()


def _chart_display_unit(values, unit):
    """Choose human-readable Japanese units and avoid scientific notation on large values."""
    series = pd.to_numeric(pd.Series(values), errors="coerce").dropna()
    max_abs = float(series.abs().max()) if not series.empty else 0.0
    normalized_unit = str(unit or "").replace(" ", "").replace("　", "")
    factor, display_unit = 1.0, normalized_unit or "値"

    is_area = any(token in normalized_unit for token in ("㎡", "m²", "m2", "平方メートル", "平米"))
    is_currency_man = "万円" in normalized_unit
    is_currency_oku = "億円" in normalized_unit
    is_yen = ("円" in normalized_unit) and not (is_currency_man or is_currency_oku)

    if is_area:
        if max_abs >= 100_000_000:
            factor, display_unit = 100_000_000.0, "億㎡"
        elif max_abs >= 10_000:
            factor, display_unit = 10_000.0, "万㎡"
        else:
            display_unit = "㎡"
    elif is_currency_man:
        if max_abs >= 10_000:
            factor, display_unit = 10_000.0, "億円"
        else:
            display_unit = "万円"
    elif is_currency_oku:
        display_unit = "億円"
    elif is_yen:
        if max_abs >= 100_000_000:
            factor, display_unit = 100_000_000.0, "億円"
        elif max_abs >= 10_000:
            factor, display_unit = 10_000.0, "万円"
        else:
            display_unit = "円"
    elif normalized_unit in {"千㎡", "千m²", "千平方メートル"}:
        if max_abs >= 100_000:
            factor, display_unit = 100_000.0, "億㎡"
        elif max_abs >= 10:
            factor, display_unit = 10.0, "万㎡"
    return factor, display_unit


def _compact_category_label(value, max_chars=18):
    """Make e-Stat's long compound labels readable in charts and PDF exports."""
    import unicodedata

    text = unicodedata.normalize("NFKC", str(value or ""))
    text = re.sub(r"[\s\u3000]+", "", text)
    text = text.replace("，", "、").replace("；", "、")

    replacements = [
        ("鉄骨鉄筋コンクリート造", "SRC造"),
        ("鉄筋コンクリート造", "RC造"),
        ("コンクリートブロック造", "CB造"),
        ("鉄骨造", "S造"),
        ("木造", "木造"),
        ("非住宅建築物", "非住宅"),
        ("共同住宅", "共同住宅"),
        ("一戸建て", "戸建"),
        ("一戸建", "戸建"),
        ("長屋建て", "長屋"),
        ("長屋建", "長屋"),
        ("床面積の合計", "床面積"),
        ("建築物の数", "棟数"),
        ("工事費予定額", "工事費"),
    ]
    for old, new in replacements:
        text = text.replace(old, new)

    raw_parts = re.split(r"[／/｜|]+", text)
    # Old e-Stat workbooks prepend the whole table heading to every metric label,
    # e.g. "構造別、用途別―建築物の数、床面積の合計、工事費予定額／S造／事務所／…".
    # Drop that repeated preamble when actual category components follow it.
    if len(raw_parts) > 1 and any(
        token in raw_parts[0] for token in ("構造別", "用途別", "利用関係別", "建築物の数", "新設住宅の戸数")
    ):
        raw_parts = raw_parts[1:]
    parts = []
    measure_words = {
        "構造別", "用途別", "建て方別", "利用関係別", "都道府県別",
        "戸数", "棟数", "床面積", "値", "計", "総数", "総計",
        "平方メートル", "千平方メートル", "平米", "㎡", "m2", "m²",
        "円", "万円", "億円", "戸", "棟", "工事費", "工事費/万円",
    }
    for part in raw_parts:
        part = part.strip("・_－-―—:：,、")
        if not part:
            continue
        # Remove generic table-level descriptions repeated in almost every column.
        for generic in (
            "構造別", "用途別", "建て方別", "利用関係別", "都道府県別",
            "床面積の合計", "建築物の数", "工事費予定額", "新設建築物",
        ):
            part = part.replace(generic, "")
        part = part.strip("・_－-―—:：,、")
        if not part or part in measure_words:
            continue
        # A unit or measure ending is not a category, and should not repeat as a label.
        if part in {"床面積", "棟数", "戸数", "万円", "億円", "平方メートル", "千平方メートル"}:
            continue
        # If one remaining component still carries an e-Stat combined measure description, strip it.
        part = part.replace("床面積", "").replace("工事費", "").replace("棟数", "")
        part = part.strip("・_－-―—:：,、")
        if part and part not in parts:
            parts.append(part)

    if not parts:
        # Fallback to a compact source label rather than producing an empty label.
        fallback = text
        for generic in ("構造別", "用途別", "建て方別", "利用関係別", "都道府県別",
                        "床面積の合計", "建築物の数", "工事費予定額", "平方メートル"):
            fallback = fallback.replace(generic, "")
        fallback = re.sub(r"[／/｜|]+", "・", fallback).strip("・_－-:：,、")
        parts = [fallback or "全体"]

    compact = "・".join(parts)
    if len(compact) <= max_chars:
        return compact
    # Prefer keeping structural and use-type terms at the front; trim only the tail.
    return compact[:max_chars - 1] + "…"


def _render_generic_pdf_chart(spec):
    """Render one legible Japanese chart at full-page resolution for A4 PDF output."""
    import io
    import matplotlib.dates as mdates
    from matplotlib.ticker import FuncFormatter

    plt = setup_matplotlib_japanese()
    df = spec.get("data", pd.DataFrame()).copy()
    kind = spec.get("kind", "line")
    raw_unit = spec.get("unit", "値")
    if "値" in df.columns:
        df["値"] = pd.to_numeric(df["値"], errors="coerce")
    factor, display_unit = _chart_display_unit(df["値"] if "値" in df else [], raw_unit)
    if "値" in df.columns:
        df["_表示値"] = df["値"] / factor

    # A4 landscape's plotting area is wide; keep a fixed canvas size so every page
    # uses the same scaling instead of being shrunk by varying tight bounding boxes.
    fig, ax = plt.subplots(figsize=(10.8, 6.2), dpi=160)
    if df.empty:
        ax.text(0.5, 0.5, "表示できるデータがありません", ha="center", va="center", transform=ax.transAxes)
        ax.set_axis_off()
    elif kind == "bar":
        df = df.sort_values("_表示値", ascending=True)
        full_labels = df["指標"].astype(str).tolist()
        labels = [_compact_category_label(label, max_chars=22) for label in full_labels]
        values = pd.to_numeric(df["_表示値"], errors="coerce").fillna(0).tolist()
        y_positions = list(range(len(labels)))
        bars = ax.barh(y_positions, values, color="#3478f6", height=0.68)
        ax.set_yticks(y_positions)
        ax.set_yticklabels(labels)
        ax.set_xlabel(f"値（{display_unit}）", fontsize=10)
        ax.ticklabel_format(axis="x", style="plain", useOffset=False)
        ax.xaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:,.0f}" if abs(v) >= 10 else f"{v:,.1f}".rstrip("0").rstrip(".")))
        ax.tick_params(axis="y", labelsize=8.5, pad=7)
        ax.tick_params(axis="x", labelsize=8.5)
        ax.grid(True, axis="x", alpha=0.22)
        # Put the numeric value at the end of every bar. Extend x range to leave room for labels.
        label_font = 7.5 if len(values) <= 12 else 6.5
        ax.bar_label(bars, labels=[_format_data_label(v) for v in values], padding=3, fontsize=label_font)
        max_value = max([abs(v) for v in values] + [1])
        ax.set_xlim(min(0, min(values, default=0) * 1.05), max_value * 1.18)
        # More left margin for Japanese multi-level labels; the labels themselves are abbreviated/wrapped.
        fig.subplots_adjust(left=0.29, right=0.96, top=0.90, bottom=0.12)
    else:
        grouped_series = list(df.groupby("指標", sort=False))
        series_count = len(grouped_series)
        # Avoid filling a long timeline with unreadable labels: label all annual/short series,
        # sample long series, and show latest endpoints for charts with many categories.
        for series_index, (label, group) in enumerate(grouped_series):
            group = group.sort_values("年月")
            dates = pd.to_datetime(group["年月"])
            values = pd.to_numeric(group["_表示値"], errors="coerce")
            ax.plot(dates, values, label=_compact_category_label(label, max_chars=24),
                    linewidth=1.8, marker="o" if len(group) <= 24 else None,
                    markersize=3.2 if len(group) <= 24 else 2)
            n_points = len(group)
            if series_count == 1:
                max_labels = n_points if n_points <= 16 else 12
            elif series_count <= 3:
                max_labels = n_points if n_points <= 12 else 7
            elif series_count <= 6:
                max_labels = min(4, n_points)
            else:
                max_labels = 1
            indexes = _sample_label_indices(n_points, max_labels)
            for label_index, point_index in enumerate(indexes):
                if point_index >= n_points or pd.isna(values.iloc[point_index]):
                    continue
                value = float(values.iloc[point_index])
                # Stagger labels so neighboring series and points are less likely to collide.
                if max_labels == 1:
                    offset = (4, 4 + (series_index % 3) * 4)
                else:
                    offset = (0, 7 if (label_index + series_index) % 2 == 0 else -10)
                ax.annotate(
                    _format_data_label(value, percent=("%" in str(spec.get("unit", "")))),
                    (dates.iloc[point_index], value), xytext=offset, textcoords="offset points",
                    ha="center" if max_labels > 1 else "left", va="center", fontsize=6.3,
                    color="#273746",
                    bbox={"boxstyle": "round,pad=0.10", "facecolor": "white", "edgecolor": "none", "alpha": 0.78},
                    clip_on=True,
                )
        ax.set_ylabel(display_unit, fontsize=10)
        ax.ticklabel_format(axis="y", style="plain", useOffset=False)
        ax.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=3, maxticks=8))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y年%m月"))
        ax.tick_params(axis="x", labelrotation=20, labelsize=8.5)
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:,.0f}" if abs(v) >= 10 else f"{v:,.1f}".rstrip("0").rstrip(".")))
        ax.tick_params(axis="y", labelsize=8.5)
        if series_count > 1:
            ax.legend(loc="best", fontsize=8, frameon=False)
        ax.grid(True, alpha=0.22)
        fig.subplots_adjust(left=0.10, right=0.98, top=0.90, bottom=0.16)

    ax.set_title(str(spec.get("title", "")), fontsize=13, loc="left", pad=12)
    output = io.BytesIO()
    fig.savefig(output, format="png", dpi=160, facecolor="white")
    plt.close(fig)
    output.seek(0)
    return output


def create_multi_chart_pdf(title, region_label, period_label, specs):
    """Create a same-size A4 landscape PDF with one full-size chart per page."""
    import io
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
    except Exception:
        pass

    width, height = landscape(A4)
    out = io.BytesIO()
    pdf = canvas.Canvas(out, pagesize=(width, height), pageCompression=1)
    pdf.setTitle(f"{title} - {region_label}")
    font = "HeiseiKakuGo-W5"
    margin = 24
    if not specs:
        specs = [{"title": title, "data": pd.DataFrame(), "kind": "line", "unit": "値"}]
    total_pages = len(specs)

    for page_no, spec in enumerate(specs, start=1):
        pdf.setFillColor(colors.HexColor("#263746"))
        pdf.setFont(font, 15)
        pdf.drawString(margin, height - 28, title)
        pdf.setFillColor(colors.HexColor("#344054"))
        pdf.setFont(font, 11)
        chart_title = str(spec.get("title", title))
        pdf.drawString(margin, height - 48, chart_title[:100])
        pdf.setFont(font, 8.5)
        pdf.setFillColor(colors.HexColor("#64748b"))
        pdf.drawString(margin, height - 64, f"地域：{region_label}　対象期間：{period_label}")

        image = _render_generic_pdf_chart(spec)
        image_x = margin
        image_y = 34
        image_w = width - margin * 2
        image_h = height - 112
        pdf.drawImage(ImageReader(image), image_x, image_y, width=image_w, height=image_h,
                      preserveAspectRatio=True, anchor="c", mask="auto")
        pdf.setStrokeColor(colors.HexColor("#d9e0e8"))
        pdf.roundRect(image_x, image_y, image_w, image_h, 3, stroke=1, fill=0)
        pdf.setFont(font, 7)
        pdf.setFillColor(colors.HexColor("#64748b"))
        pdf.drawString(margin, 16, "出典：e-Stat 公開統計表")
        pdf.drawRightString(width - margin, 16, f"{page_no} / {total_pages}　{APP_VERSION}")
        pdf.showPage()

    pdf.save()
    return out.getvalue()

def prepare_overview_pdf_data(region_data, start_date, end_date, latest_month):
    """Build the three standard overview charts for export independently of the visible tab."""
    total_history = region_data[(region_data["指標"] == "住宅着工戸数") & (region_data["年月"].dt.date <= end_date)][["年月", "値"]].copy()
    total_history["年月"] = pd.to_datetime(total_history["年月"])
    total_history["値"] = pd.to_numeric(total_history["値"], errors="coerce")
    total_history = total_history.dropna(subset=["値"]).sort_values("年月").drop_duplicates("年月", keep="last")
    total_history["年度"] = total_history["年月"].dt.year.where(total_history["年月"].dt.month >= 4, total_history["年月"].dt.year - 1)
    total_history["年度累計"] = total_history.groupby("年度")["値"].cumsum()
    total_history["移動年計"] = total_history["値"].rolling(12, min_periods=12).sum()
    trend_df = total_history[total_history["年月"].dt.date >= start_date].copy()
    last_visible = total_history[total_history["年月"].dt.date <= end_date]
    latest_fiscal_year = int(last_visible.iloc[-1]["年度"])
    previous_fiscal_year = latest_fiscal_year - 1
    compare_source = total_history.copy()
    compare_source["月"] = compare_source["年月"].dt.month
    month_order = [4,5,6,7,8,9,10,11,12,1,2,3]
    compare_rows = []
    for m in month_order:
        prev_year = previous_fiscal_year if m >= 4 else previous_fiscal_year + 1
        current_year = latest_fiscal_year if m >= 4 else latest_fiscal_year + 1
        prev_row = compare_source[(compare_source["年月"].dt.year == prev_year) & (compare_source["月"] == m)]
        cur_row = compare_source[(compare_source["年月"].dt.year == current_year) & (compare_source["月"] == m)]
        compare_rows.append({"月ラベル": f"{m}月", "前年値": float(prev_row["値"].iloc[-1]) if not prev_row.empty else float("nan"),
                             "最新年値": float(cur_row["値"].iloc[-1]) if not cur_row.empty else float("nan")})
    compare_df = pd.DataFrame(compare_rows)
    compare_df.attrs["前年ラベル"] = f"{previous_fiscal_year}年度"
    compare_df.attrs["最新年ラベル"] = f"{latest_fiscal_year}年度"
    composition_metrics = [("持家", "持家"), ("貸家", "貸家"), ("給与住宅", "給与住宅"),
                           ("分譲：マンション", "分譲マンション（内数）"), ("分譲：戸建て", "分譲一戸建て（内数）")]
    composition_rows = []
    latest_rows = region_data[region_data["年月"] == latest_month]
    for name, metric in composition_metrics:
        values = pd.to_numeric(latest_rows.loc[latest_rows["指標"] == metric, "値"], errors="coerce").dropna()
        if not values.empty and values.iloc[0] >= 0:
            composition_rows.append({"表示名": name, "指標": metric, "値": float(values.iloc[0])})
    return trend_df, compare_df, pd.DataFrame(composition_rows)


def create_status_pdf(title, region_label, period_label, message, source_url=None):
    """Make a transparent placeholder page when a source needs configuration or is unavailable."""
    import io
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
    except Exception:
        pass
    width, height = landscape(A4)
    output = io.BytesIO()
    pdf = canvas.Canvas(output, pagesize=(width, height))
    pdf.setTitle(title)
    font = "HeiseiKakuGo-W5"
    pdf.setFillColor(colors.HexColor("#263746")); pdf.setFont(font, 18)
    pdf.drawString(32, height - 42, title)
    pdf.setFont(font, 10); pdf.setFillColor(colors.HexColor("#64748b"))
    pdf.drawString(32, height - 64, f"地域：{region_label}　対象期間：{period_label}")
    pdf.setFillColor(colors.HexColor("#344054")); pdf.setFont(font, 12)
    y = height - 112
    for line in str(message).splitlines():
        pdf.drawString(38, y, line[:100])
        y -= 24
        if y < 70:
            break
    if source_url:
        pdf.setFillColor(colors.HexColor("#3478f6")); pdf.setFont(font, 8)
        pdf.drawString(32, 35, f"出典ページ：{source_url[:130]}")
    pdf.save(); output.seek(0)
    return output.getvalue()


def merge_pdf_documents(documents):
    """Merge all report sections, normalize every page to true A4 landscape, and number globally."""
    import io
    from pypdf import PdfReader, PdfWriter, Transformation
    from pypdf.generic import RectangleObject
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    target_width, target_height = [float(value) for value in landscape(A4)]
    pages = []
    for _title, pdf_bytes in documents:
        if not pdf_bytes:
            continue
        reader = PdfReader(io.BytesIO(pdf_bytes))
        for page in reader.pages:
            # Normalize rotation before measuring the page.
            try:
                page.transfer_rotation_to_content()
            except Exception:
                pass
            source_width = float(page.mediabox.width)
            source_height = float(page.mediabox.height)
            if source_width <= 0 or source_height <= 0:
                continue
            scale = min(target_width / source_width, target_height / source_height)
            offset_x = (target_width - source_width * scale) / 2
            offset_y = (target_height - source_height * scale) / 2
            page.add_transformation(Transformation().scale(scale).translate(tx=offset_x, ty=offset_y))
            page.mediabox = RectangleObject([0, 0, target_width, target_height])
            page.cropbox = RectangleObject([0, 0, target_width, target_height])
            page.trimbox = RectangleObject([0, 0, target_width, target_height])
            pages.append(page)

    if not pages:
        raise RuntimeError("PDFに含めるページがありません。")

    # Add a single, consistent footer after normalization. This also makes it easy to
    # verify that the latest export path is being used and gives a report-wide page count.
    overlay_bytes = io.BytesIO()
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
        footer_font = "HeiseiKakuGo-W5"
    except Exception:
        footer_font = "Helvetica"
    overlay = canvas.Canvas(overlay_bytes, pagesize=(target_width, target_height), pageCompression=1)
    overlay.setTitle(f"住宅市場分析レポート - {APP_VERSION}")
    for page_number in range(1, len(pages) + 1):
        overlay.setFillColor(colors.HexColor("#667085"))
        overlay.setFont(footer_font, 7)
        overlay.drawString(24, 10, "住宅市場分析レポート　A4横")
        overlay.drawRightString(target_width - 24, 10, f"{page_number} / {len(pages)}")
        overlay.showPage()
    overlay.save()
    overlay_bytes.seek(0)
    overlay_reader = PdfReader(overlay_bytes)

    writer = PdfWriter()
    for index, page in enumerate(pages):
        page.merge_page(overlay_reader.pages[index])
        writer.add_page(page)
    writer.add_metadata({
        "/Title": "住宅市場分析レポート",
        "/Author": "住宅市場分析ダッシュボード",
        "/Subject": f"統合PDF・A4横・全{len(pages)}ページ・{APP_VERSION}",
    })
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def regional_dataset_specs(raw_df, selected_region, start_date, end_date, title, default_keywords=(), max_metrics=8):
    """Turn a parsed regional table into one time-series chart and one latest-period comparison."""
    if raw_df is None or raw_df.empty:
        return [], "データを抽出できませんでした。"
    if selected_region == "全国":
        regional = raw_df[raw_df["地域"] == "全国"].copy()
        if regional.empty:
            regional = with_national_aggregate(raw_df)
            regional = regional[regional["地域"] == "全国"].copy()
    else:
        regional = raw_df[raw_df["地域"] == selected_region].copy()
    regional = regional[(regional["年月"].dt.year >= start_date.year) & (regional["年月"].dt.year <= end_date.year)]
    if regional.empty:
        return [], "選択した地域・期間のデータがありません。"
    units = regional["単位"].dropna().astype(str).unique().tolist()
    if len(units) > 1:
        # Avoid comparing different physical units on a shared axis: choose the most common unit.
        unit = regional["単位"].astype(str).value_counts().index[0]
        regional = regional[regional["単位"].astype(str) == unit].copy()
    else:
        unit = units[0] if units else "値"
    labels = sorted(regional["指標"].dropna().astype(str).unique().tolist())
    selected = choose_default_metrics(labels, max_metrics, default_keywords)
    if not selected:
        selected = labels[:max_metrics]
    regional = regional[regional["指標"].isin(selected)].copy()
    if regional.empty:
        return [], "表示できる分類がありません。"
    latest_time = regional["年月"].max()
    latest = regional[regional["年月"] == latest_time].copy()
    specs = [{"title": f"{title}（推移）", "data": regional[["年月", "指標", "値"]].copy(), "kind": "line", "unit": unit}]
    if not latest.empty:
        specs.append({"title": f"{title}（{latest_time.year}年・項目比較）", "data": latest[["指標", "値"]].copy(), "kind": "bar", "unit": unit})
    return specs, ""


def build_all_screens_pdf(region_label, selected_region, start_date, end_date, region_data, cards, latest_month, excel_url):
    """Generate and merge the overview plus every detailed analysis screen into one PDF."""
    documents = []
    period_label = f"{start_date:%Y年%m月}～{end_date:%Y年%m月}"
    notes = []
    trend_df, compare_df, composition_df = prepare_overview_pdf_data(region_data, start_date, end_date, latest_month)
    documents.append(("市場概況", create_dashboard_pdf(region_label, start_date, end_date, latest_month,
                                                        cards, trend_df, compare_df, composition_df)))

    # Screen 2: the seven monthly housing-start series, as displayed by the detailed screen.
    detail_data = region_data[(region_data["年月"].dt.date >= start_date) & (region_data["年月"].dt.date <= end_date)].copy()
    detail_metrics = ["住宅着工戸数", "持家", "貸家", "給与住宅", "分譲住宅", "分譲マンション（内数）", "分譲一戸建て（内数）"]
    detail_specs = []
    detail_value_mode = st.session_state.get("detail_value_mode", "着工戸数")
    ycol = "値" if detail_value_mode == "着工戸数" else "前年比（％）"
    for metric in detail_metrics:
        subset = detail_data[detail_data["指標"] == metric].copy()
        if ycol not in subset:
            continue
        subset[ycol] = pd.to_numeric(subset[ycol], errors="coerce")
        subset = subset.dropna(subset=[ycol]).sort_values("年月")
        if not subset.empty:
            detail_specs.append({"title": metric, "data": subset[["年月", "指標", ycol]].rename(columns={ycol: "値"}),
                                 "kind": "line", "unit": "戸" if ycol == "値" else "%"})
    if detail_specs:
        documents.append(("住宅着工の詳細", create_multi_chart_pdf(f"住宅着工の詳細（{detail_value_mode}）", region_label, period_label, detail_specs)))
    else:
        documents.append(("住宅着工の詳細", create_status_pdf("住宅着工の詳細", region_label, period_label, "選択期間の詳細データを取得できませんでした。", ESTAT_HOUSING_LIST_URL)))

    # Screen 3: annual structure/use/house-type data.
    try:
        structure_raw, structure_file_url, _ = load_regional_source(
            ESTAT_HOUSING_ANNUAL_URL, ("構造別", "建て方別", "都道府県別", "戸数"), (), "年次",
            "e-Stat 住宅利用関係別・構造別・建て方別・都道府県別戸数", "戸")
        specs, message = regional_dataset_specs(structure_raw, selected_region, start_date, end_date,
                                                "住宅の構造・建て方・利用関係", ("木造", "鉄骨造", "鉄筋コンクリート", "鉄骨鉄筋", "共同住宅", "一戸建", "持家", "貸家"), 10)
        if specs:
            documents.append(("構造・建て方分析", create_multi_chart_pdf("構造・建て方分析", region_label, period_label + "（年次）", specs)))
        else:
            documents.append(("構造・建て方分析", create_status_pdf("構造・建て方分析", region_label, period_label, message, structure_file_url)))
    except Exception as exc:
        notes.append(f"構造・建て方分析：{exc}")
        documents.append(("構造・建て方分析", create_status_pdf("構造・建て方分析", region_label, period_label,
                                                                   "この統計表の自動取得または抽出に失敗しました。エラー詳細：" + str(exc)[:180], ESTAT_HOUSING_ANNUAL_URL)))

    # Screen 4a: residential floor area.
    try:
        housing_area, housing_area_url, _ = load_regional_source(
            ESTAT_HOUSING_LIST_URL, ("都道府県別床面積",), (), "月次", "e-Stat 住宅都道府県別床面積", "㎡")
        specs, message = regional_dataset_specs(housing_area, selected_region, start_date, end_date,
                                                "住宅の床面積", ("総数", "床面積", "持家", "貸家"), 6)
        if specs:
            documents.append(("住宅床面積", create_multi_chart_pdf("住宅の床面積", region_label, period_label, specs)))
        else:
            documents.append(("住宅床面積", create_status_pdf("住宅の床面積", region_label, period_label, message, housing_area_url)))
    except Exception as exc:
        notes.append(f"住宅床面積：{exc}")
        documents.append(("住宅床面積", create_status_pdf("住宅の床面積", region_label, period_label,
                                                           "この統計表の自動取得または抽出に失敗しました。" + str(exc)[:180], ESTAT_HOUSING_LIST_URL)))

    # Screen 4b: building floor area.
    try:
        building_area, building_area_url, _ = load_regional_source(
            ESTAT_BUILDING_MONTHLY_URL, ("都道府県別床面積",), (), "月次", "e-Stat 建築物都道府県別床面積", "㎡")
        specs, message = regional_dataset_specs(building_area, selected_region, start_date, end_date,
                                                "建築物の床面積", ("総数", "床面積", "用途"), 6)
        if specs:
            documents.append(("建築物床面積", create_multi_chart_pdf("建築物の床面積", region_label, period_label, specs)))
        else:
            documents.append(("建築物床面積", create_status_pdf("建築物の床面積", region_label, period_label, message, building_area_url)))
    except Exception as exc:
        notes.append(f"建築物床面積：{exc}")
        documents.append(("建築物床面積", create_status_pdf("建築物の床面積", region_label, period_label,
                                                             "この統計表の自動取得または抽出に失敗しました。" + str(exc)[:180], ESTAT_BUILDING_MONTHLY_URL)))

    # Screen 4c: building usage/structure annual data.
    try:
        building_struct, building_struct_url, _ = load_regional_source(
            ESTAT_BUILDING_ANNUAL_URL, ("構造別", "用途別", "都道府県別"), (), "年次",
            "e-Stat 建築物構造別・用途別・都道府県別統計", "値")
        specs, message = regional_dataset_specs(building_struct, selected_region, start_date, end_date,
                                                "建築物の用途・構造", ("非住宅", "店舗", "事務所", "工場", "倉庫", "鉄骨", "鉄筋", "木造"), 10)
        if specs:
            documents.append(("建築物用途・構造", create_multi_chart_pdf("建築物の用途・構造", region_label, period_label + "（年次）", specs)))
        else:
            documents.append(("建築物用途・構造", create_status_pdf("建築物の用途・構造", region_label, period_label, message, building_struct_url)))
    except Exception as exc:
        notes.append(f"建築物用途・構造：{exc}")
        documents.append(("建築物用途・構造", create_status_pdf("建築物の用途・構造", region_label, period_label,
                                                                 "この統計表の自動取得または抽出に失敗しました。" + str(exc)[:180], ESTAT_BUILDING_ANNUAL_URL)))

    # Screen 5: renovation regional reference data is a database table, not an Excel download.
    try:
        try:
            estat_app_id = st.secrets.get("ESTAT_APP_ID", "")
        except Exception:
            estat_app_id = ""
        if estat_app_id:
            renovation, renovation_url = load_renovation_api_data(estat_app_id)
            specs, message = regional_dataset_specs(renovation, selected_region, start_date, end_date,
                                                    "リフォーム・リニューアル受注高", ("計", "住宅", "非住宅建築物"), 3)
            if specs:
                documents.append(("リフォーム市場", create_multi_chart_pdf("リフォーム・リニューアル市場", region_label, period_label + "（年次）", specs)))
            else:
                documents.append(("リフォーム市場", create_status_pdf("リフォーム市場", region_label, period_label, message, renovation_url)))
        else:
            note = "e-Statのこの統計表はデータベース形式です。自動取得には無料のe-Stat APIアプリケーションIDをStreamlit SecretsのESTAT_APP_IDに設定してください。"
            documents.append(("リフォーム市場", create_status_pdf("リフォーム市場", region_label, period_label, note, ESTAT_RENOVATION_TABLE_URL)))
    except Exception as exc:
        notes.append(f"リフォーム市場：{exc}")
        documents.append(("リフォーム市場", create_status_pdf("リフォーム市場", region_label, period_label,
                                                               "e-Stat APIからデータを取得できませんでした。" + str(exc)[:180], ESTAT_RENOVATION_TABLE_URL)))

    # Municipality / housing acquisition environment. The user selects a municipality on that page;
    # preserve that selection in session_state so the combined PDF can include this section too.
    municipal_selection = st.session_state.get("municipal_selection")
    if municipal_selection:
        try:
            try:
                api_key = st.secrets.get("ESTAT_APP_ID", "")
            except Exception:
                api_key = ""
            if api_key:
                bundle = cached_municipal_dashboard_data(
                    api_key,
                    municipal_selection["prefecture"],
                    municipal_selection["municipality_code"],
                    municipal_selection.get("municipality_name"),
                )
                specs = municipal_pdf_specs(bundle)
                if specs:
                    documents.append(("市区町村・住宅取得環境", create_multi_chart_pdf(
                        "市区町村・住宅取得環境",
                        f"{municipal_selection['prefecture']}・{municipal_selection['municipality_name']}",
                        "公表統計の最新期間", specs)))
                else:
                    documents.append(("市区町村・住宅取得環境", create_status_pdf(
                        "市区町村・住宅取得環境", municipal_selection["municipality_name"], period_label,
                        "市区町村分析のPDFデータがありません。", ESTAT_MUNICIPAL_BUILDING_URL)))
            else:
                documents.append(("市区町村・住宅取得環境", create_status_pdf(
                    "市区町村・住宅取得環境", municipal_selection["municipality_name"], period_label,
                    "ESTAT_APP_IDが未設定のため、市区町村分析を取得できません。", ESTAT_MUNICIPAL_BUILDING_URL)))
        except Exception as exc:
            notes.append(f"市区町村・住宅取得環境：{exc}")
            documents.append(("市区町村・住宅取得環境", create_status_pdf(
                "市区町村・住宅取得環境", municipal_selection["municipality_name"], period_label,
                "市区町村データを取得できませんでした。" + str(exc)[:180], ESTAT_MUNICIPAL_BUILDING_URL)))
    else:
        documents.append(("市区町村・住宅取得環境", create_status_pdf(
            "市区町村・住宅取得環境", region_label, period_label,
            "市区町村・住宅取得環境の画面で市区町村を選択してから、一括PDFを作成してください。",
            ESTAT_MUNICIPAL_BUILDING_URL)))

    return merge_pdf_documents(documents), notes


st.set_page_config(page_title="住宅着工統計ダッシュボード", page_icon="📊", layout="wide")
st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 1rem; max-width: 1800px;}
div[data-testid="stHorizontalBlock"] {gap: 0.55rem;}
.kpi-card {background:#ffffff; border:1px solid #d8e0e8; border-top:4px solid var(--accent); border-radius:6px; padding:9px 10px 8px 10px; min-height:117px; box-shadow:0 1px 2px rgba(20,40,60,.08);}
.kpi-label {font-size:13px; font-weight:700; color:#344054; margin-bottom:4px; white-space:nowrap;}
.kpi-value {font-size:23px; line-height:1.15; font-weight:800; color:#263238; white-space:nowrap;}
.kpi-sub {font-size:10px; line-height:1.5; color:#667085; margin-top:5px; white-space:nowrap;}
.kpi-yoy {font-size:10px; font-weight:700; text-align:right;}
.panel-title {font-size:16px; font-weight:750; color:#263746; border-bottom:1px solid #d8e0e8; padding-bottom:7px; margin-bottom:4px;}
div[data-testid="stRadio"] label {font-weight:600;}
@media (max-width: 900px) {.kpi-value {font-size:18px;} .kpi-card {padding:7px; min-height:104px;} .kpi-label {font-size:11px;}}
</style>
""", unsafe_allow_html=True)
st.title("住宅着工統計ダッシュボード")

with st.sidebar:
    st.caption(f"PDF出力形式：{APP_VERSION}")
    st.header("① 地域を選択")
    if st.button("e-Statからデータを再取得", use_container_width=True, help="最新のExcelを取得します"):
        st.cache_data.clear()
        st.rerun()

try:
    with st.spinner("e-Statのデータを確認しています..."):
        data, excel_url, dataset_info = load_estat_data()
except Exception as exc:
    st.error("e-Statからデータを読み込めませんでした。")
    st.code(str(exc))
    st.markdown(f"[e-Statの住宅着工統計一覧を開く]({ESTAT_HOUSING_LIST_URL})")
    st.stop()

data = data.copy()
data["年月"] = pd.to_datetime(data["年月"])
min_month = data["年月"].min()
max_month = data["年月"].max()
available_regions = set(data["地域"].dropna().unique().tolist())
region_options = ["全国"] + [p for p in PREFECTURES if p in available_regions]
with st.sidebar:
    selected_region = st.selectbox(
        "全国または都道府県", options=region_options, index=None,
        placeholder="地域を選択してください", key="region_selector",
    )
if selected_region is None:
    st.info("左側から「全国」または表示したい都道府県を選択してください。")
    st.stop()

if selected_region == "全国":
    region_data = build_national_series(data)
    region_label = "全国"
else:
    region_data = data.loc[data["地域"] == selected_region].copy()
    region_label = selected_region
    if region_data.empty:
        st.error(f"{selected_region}のデータを取得できていません。")
        st.stop()

metrics_order = ["住宅着工戸数", "持家", "貸家", "給与住宅", "分譲住宅", "分譲マンション（内数）", "分譲一戸建て（内数）"]
available_metrics = [m for m in metrics_order if m in set(region_data["指標"].dropna().unique())]
if "住宅着工戸数" not in available_metrics:
    st.error("選択した地域の住宅着工戸数データがありません。")
    st.stop()

available_months = sorted(pd.Timestamp(v).to_period("M").to_timestamp() for v in region_data["年月"].dropna().drop_duplicates().tolist())
period_years = sorted({m.year for m in available_months})
months_by_year = {year: sorted({m.month for m in available_months if m.year == year}) for year in period_years}
default_start_year = 2020 if 2020 in period_years else period_years[0]
default_end_year = period_years[-1]
with st.sidebar:
    st.divider()
    st.header("② 表示期間")
    st.markdown("**開始年月**")
    sc1, sc2 = st.columns(2)
    with sc1:
        start_year = st.selectbox("開始年", period_years, index=period_years.index(default_start_year), format_func=lambda y: f"{y}年", key="start_year")
    with sc2:
        start_month_options = months_by_year[start_year]
        start_month = st.selectbox("開始月", start_month_options, index=start_month_options.index(1) if 1 in start_month_options else 0, format_func=lambda m: f"{m}月", key=f"start_month_{start_year}")
    st.markdown("**終了年月**")
    ec1, ec2 = st.columns(2)
    with ec1:
        end_year = st.selectbox("終了年", period_years, index=period_years.index(default_end_year), format_func=lambda y: f"{y}年", key="end_year")
    with ec2:
        end_month_options = months_by_year[end_year]
        default_end_month = end_month_options.index(12) if 12 in end_month_options else len(end_month_options) - 1
        end_month = st.selectbox("終了月", end_month_options, index=default_end_month, format_func=lambda m: f"{m}月", key=f"end_month_{end_year}")

start_date = date(int(start_year), int(start_month), 1)
end_date = date(int(end_year), int(end_month), 1)
if start_date > end_date:
    st.error("開始年月は終了年月以前を指定してください。")
    st.stop()

with st.sidebar:
    st.divider()
    st.header("③ 分析画面")
    analysis_view = st.radio(
        "表示する分析", 
        ["市場概況", "住宅着工の詳細", "構造・建て方分析", "建築規模・非住宅", "リフォーム市場", "市区町村・住宅取得環境"],
        index=0,
        key="analysis_view",
    )

metric_colors = {
    "住宅着工戸数": "#3478f6", "持家": "#29a66f", "貸家": "#f59a43",
    "給与住宅": "#9149df", "分譲住宅": "#22a9c7",
    "分譲一戸建て（内数）": "#17a8c7", "分譲マンション（内数）": "#ee5058",
}
card_metrics = [
    ("総計", "住宅着工戸数", "#3478f6"),
    ("持家", "持家", "#29a66f"),
    ("貸家", "貸家", "#f59a43"),
    ("給与住宅", "給与住宅", "#9149df"),
    ("分譲：戸建て", "分譲一戸建て（内数）", "#22a9c7"),
    ("分譲：マンション", "分譲マンション（内数）", "#ee5058"),
]

# Select the most recent data month within the chosen period for the KPI cards.
period_rows = region_data[(region_data["年月"].dt.date >= start_date) & (region_data["年月"].dt.date <= end_date)]
total_period = period_rows[period_rows["指標"] == "住宅着工戸数"].dropna(subset=["値"])
if total_period.empty:
    st.warning("選択した地域・期間に表示できるデータがありません。")
    st.stop()
latest_month = pd.Timestamp(total_period["年月"].max())

cards = []
for label, metric, accent in card_metrics:
    current_rows = region_data[(region_data["指標"] == metric) & (region_data["年月"] == latest_month)]
    current_value = pd.to_numeric(current_rows["値"], errors="coerce").dropna() if not current_rows.empty else pd.Series(dtype=float)
    current = float(current_value.iloc[0]) if not current_value.empty else None
    previous_month = latest_month - pd.DateOffset(years=1)
    previous_rows = region_data[(region_data["指標"] == metric) & (region_data["年月"] == previous_month)]
    previous_values = pd.to_numeric(previous_rows["値"], errors="coerce").dropna() if not previous_rows.empty else pd.Series(dtype=float)
    previous = float(previous_values.iloc[0]) if not previous_values.empty else None
    ratio = (current / previous * 100) if current is not None and previous not in (None, 0) else None
    value_text = f"{current:,.0f}戸" if current is not None else "—"
    previous_text = f"前年同月 {previous:,.0f}戸" if previous is not None else "前年同月 —"
    yoy_text = f"{ratio:.1f}%" if ratio is not None else "—"
    yoy_color = "#d92d20" if ratio is not None and ratio > 100 else "#1570ef" if ratio is not None else "#667085"
    cards.append({"label": label, "metric": metric, "color": accent, "value": value_text, "previous": previous_text, "yoy": yoy_text, "yoy_color": yoy_color, "current": current, "previous_value": previous, "ratio": ratio})

# Combined PDF export is available from every analysis screen, including the municipality / housing-acquisition page.
municipal_signature_part = st.session_state.get("municipal_selection", {})
municipal_signature = (
    f"{municipal_signature_part.get('prefecture', '')}|"
    f"{municipal_signature_part.get('municipality_code', '')}"
)
combined_signature = f"{selected_region}|{start_date}|{end_date}|{latest_month}|{municipal_signature}"
if st.session_state.get("all_screens_pdf_signature") != combined_signature:
    st.session_state.pop("all_screens_pdf_bytes", None)
    st.session_state.pop("all_screens_pdf_signature", None)
with st.container(border=True):
    st.markdown("**PDFレポート**")
    st.caption("市場概況・住宅着工詳細・構造／建て方・建築規模／非住宅・リフォーム・市区町村／住宅取得環境をA4 PDFにまとめます。市区町村分析は、先に対象自治体を選択してください。")
    if st.button("全分析画面を1つのPDFにまとめる", type="primary", key="all_screens_pdf_button"):
        try:
            with st.spinner("全分析画面のデータとグラフをまとめています。初回は取得に時間がかかる場合があります..."):
                all_pdf, pdf_notes = build_all_screens_pdf(
                    region_label, selected_region, start_date, end_date, region_data, cards, latest_month, excel_url)
            st.session_state["all_screens_pdf_bytes"] = all_pdf
            st.session_state["all_screens_pdf_signature"] = combined_signature
            st.session_state["all_screens_pdf_notes"] = pdf_notes
        except Exception as exc:
            st.error("一括PDFの作成に失敗しました。")
            st.code(str(exc))
    if st.session_state.get("all_screens_pdf_signature") == combined_signature and st.session_state.get("all_screens_pdf_bytes"):
        safe_region = re.sub(r"[^0-9A-Za-zぁ-んァ-ヶ一-龠_-]+", "_", selected_region)
        st.download_button(
            "全分析画面をまとめたPDFをダウンロード",
            st.session_state["all_screens_pdf_bytes"],
            file_name=f"住宅市場分析_全画面_{safe_region}_{start_date:%Y%m}_{end_date:%Y%m}.pdf",
            mime="application/pdf", key="all_screens_pdf_download", use_container_width=True)
        notes = st.session_state.get("all_screens_pdf_notes", [])
        if notes:
            with st.expander("PDFに含まれる取得状況"):
                st.write("一部の統計表を取得できなかったため、該当ページには出典と状況を記載しています。")
                for note in notes:
                    st.write(f"- {note}")


if analysis_view == "市場概況":
    # KPI row similar to the supplied design reference.
    card_cols = st.columns(6, gap="small")
    for col, card in zip(card_cols, cards):
        with col:
            html = (
                f'<div class="kpi-card" style="--accent:{card["color"]}">'
                f'<div class="kpi-label">{card["label"]}</div>'
                f'<div class="kpi-value">{card["value"]}</div>'
                f'<div class="kpi-sub">{card["previous"]}</div>'
                f'<div class="kpi-yoy" style="color:{card["yoy_color"]}">前年同月比 {card["yoy"]}</div>'
                '</div>'
            )
            st.markdown(html, unsafe_allow_html=True)

    # Build main trend data from the full history up to end_date so fiscal-year accumulations
    # and rolling 12-month totals remain correct even if the chart's start date is later.
    total_history = region_data[(region_data["指標"] == "住宅着工戸数") & (region_data["年月"].dt.date <= end_date)][["年月", "値"]].copy()
    total_history["年月"] = pd.to_datetime(total_history["年月"])
    total_history["値"] = pd.to_numeric(total_history["値"], errors="coerce")
    total_history = total_history.dropna(subset=["値"]).sort_values("年月").drop_duplicates("年月", keep="last")
    total_history["年度"] = total_history["年月"].dt.year.where(total_history["年月"].dt.month >= 4, total_history["年月"].dt.year - 1)
    total_history["年度累計"] = total_history.groupby("年度")["値"].cumsum()
    total_history["移動年計"] = total_history["値"].rolling(12, min_periods=12).sum()
    trend_df = total_history[total_history["年月"].dt.date >= start_date].copy()

    # Compare the latest fiscal year in the selected period against the previous fiscal year.
    last_visible = total_history[total_history["年月"].dt.date <= end_date]
    latest_fiscal_year = int(last_visible.iloc[-1]["年度"])
    previous_fiscal_year = latest_fiscal_year - 1
    compare_source = total_history.copy()
    compare_source["年度"] = compare_source["年月"].dt.year.where(compare_source["年月"].dt.month >= 4, compare_source["年月"].dt.year - 1)
    compare_source["月"] = compare_source["年月"].dt.month
    month_order = [4,5,6,7,8,9,10,11,12,1,2,3]
    month_labels = [f"{m}月" for m in month_order]
    compare_rows = []
    for month_number, month_label in zip(month_order, month_labels):
        previous_calendar_year = previous_fiscal_year if month_number >= 4 else previous_fiscal_year + 1
        latest_calendar_year = latest_fiscal_year if month_number >= 4 else latest_fiscal_year + 1
        prev_row = compare_source[(compare_source["年月"].dt.year == previous_calendar_year) & (compare_source["月"] == month_number)]
        latest_row = compare_source[(compare_source["年月"].dt.year == latest_calendar_year) & (compare_source["月"] == month_number)]
        compare_rows.append({"月ラベル": month_label, "前年値": float(prev_row["値"].iloc[-1]) if not prev_row.empty else float("nan"), "最新年値": float(latest_row["値"].iloc[-1]) if not latest_row.empty else float("nan")})
    compare_df = pd.DataFrame(compare_rows)
    compare_df.attrs["前年ラベル"] = f"{previous_fiscal_year}年度"
    compare_df.attrs["最新年ラベル"] = f"{latest_fiscal_year}年度"

    composition_metrics = [
        ("持家", "持家"), ("貸家", "貸家"), ("給与住宅", "給与住宅"),
        ("分譲：マンション", "分譲マンション（内数）"), ("分譲：戸建て", "分譲一戸建て（内数）"),
    ]
    composition_rows = []
    latest_rows = region_data[region_data["年月"] == latest_month]
    for display_name, metric in composition_metrics:
        vals = latest_rows.loc[latest_rows["指標"] == metric, "値"]
        value = pd.to_numeric(vals, errors="coerce").dropna()
        if not value.empty and value.iloc[0] >= 0:
            composition_rows.append({"表示名": display_name, "指標": metric, "値": float(value.iloc[0])})
    composition_df = pd.DataFrame(composition_rows)

    st.markdown(f"### {region_label}　住宅着工統計")
    st.caption(f"表示期間：{start_date:%Y年%m月}～{end_date:%Y年%m月}")
    chart_cols = st.columns([1.25, 1.05, 0.95], gap="small")

    with chart_cols[0]:
        st.markdown('<div class="panel-title">住宅着工戸数の推移</div>', unsafe_allow_html=True)
        trend_fig = go.Figure()
        if not trend_df.empty:
            def latest_value_text(series):
                values = pd.to_numeric(series, errors="coerce").reset_index(drop=True)
                labels = [""] * len(values)
                valid_positions = [i for i, value in enumerate(values.tolist()) if pd.notna(value)]
                if valid_positions:
                    last_position = valid_positions[-1]
                    labels[last_position] = f"{values.iloc[last_position]:,.0f}"
                return labels
            trend_fig.add_trace(go.Scatter(x=trend_df["年月"], y=trend_df["値"], name="月別実績", mode="lines+markers+text", text=latest_value_text(trend_df["値"]), textposition="top center", textfont=dict(size=10, color="#3478f6"), marker=dict(size=3), line=dict(color="#3478f6", width=2)))
            trend_fig.add_trace(go.Scatter(x=trend_df["年月"], y=trend_df["年度累計"], name="年度累計", mode="lines+markers+text", text=latest_value_text(trend_df["年度累計"]), textposition="bottom center", textfont=dict(size=10, color="#29a66f"), marker=dict(size=3), line=dict(color="#29a66f", width=2)))
            trend_fig.add_trace(go.Scatter(x=trend_df["年月"], y=trend_df["移動年計"], name="移動年計", mode="lines+markers+text", text=latest_value_text(trend_df["移動年計"]), textposition="top center", textfont=dict(size=10, color="#ef5058"), marker=dict(size=3), line=dict(color="#ef5058", width=2)))
        trend_fig.update_layout(height=390, margin=dict(l=0,r=4,t=12,b=5), legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=10)), xaxis_title=None, yaxis_title="戸数（戸）", hovermode="x unified")
        trend_fig.update_xaxes(tickformat="%Y年%m月", nticks=7, showgrid=True, gridcolor="#e8edf3")
        trend_fig.update_yaxes(tickformat=",.0f", showgrid=True, gridcolor="#e8edf3")
        st.plotly_chart(trend_fig, use_container_width=True, config={"displayModeBar": False})

    with chart_cols[1]:
        st.markdown(f'<div class="panel-title">月別比較（{previous_fiscal_year}年度・{latest_fiscal_year}年度）</div>', unsafe_allow_html=True)
        compare_fig = go.Figure()
        compare_fig.add_trace(go.Bar(x=compare_df["月ラベル"], y=compare_df["前年値"], name=f"{previous_fiscal_year}年度", marker_color="#aeb8c3", text=[_format_data_label(v) if pd.notna(v) else "" for v in compare_df["前年値"]], textposition="outside", textfont=dict(size=8), cliponaxis=False))
        compare_fig.add_trace(go.Bar(x=compare_df["月ラベル"], y=compare_df["最新年値"], name=f"{latest_fiscal_year}年度", marker_color="#3478f6", text=[_format_data_label(v) if pd.notna(v) else "" for v in compare_df["最新年値"]], textposition="outside", textfont=dict(size=8), cliponaxis=False))
        compare_fig.update_layout(barmode="group", height=390, margin=dict(l=0,r=4,t=12,b=5), legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=10)), xaxis_title=None, yaxis_title="戸数（戸）")
        compare_fig.update_yaxes(tickformat=",.0f", showgrid=True, gridcolor="#e8edf3")
        compare_fig.update_xaxes(showgrid=False)
        st.plotly_chart(compare_fig, use_container_width=True, config={"displayModeBar": False})

    with chart_cols[2]:
        st.markdown(f'<div class="panel-title">利用関係別構成比（{latest_month:%Y年%m月}）</div>', unsafe_allow_html=True)
        if not composition_df.empty and composition_df["値"].sum() > 0:
            composition_fig = px.pie(composition_df, values="値", names="表示名", hole=0.48, color="表示名", color_discrete_map={"持家":"#35a879", "貸家":"#f59a43", "給与住宅":"#9b59d0", "分譲：マンション":"#e95059", "分譲：戸建て":"#22a6c7"})
            composition_fig.update_traces(textposition="inside", texttemplate="%{label}<br>%{value:,.0f}戸<br>%{percent}", insidetextfont=dict(color="white", size=10), hovertemplate="%{label}<br>%{value:,.0f}戸 (%{percent})<extra></extra>")
            composition_fig.update_layout(height=390, margin=dict(l=0,r=0,t=12,b=5), showlegend=False)
            st.plotly_chart(composition_fig, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("構成比を表示できるデータがありません。")

    # PDF output mirrors the overview: six cards and three purpose-specific charts on one wide page.
    pdf_signature = f"{selected_region}|{start_date}|{end_date}|{latest_month}"
    if st.session_state.get("dashboard_pdf_signature") != pdf_signature:
        st.session_state.pop("dashboard_pdf_bytes", None)
        st.session_state.pop("dashboard_pdf_signature", None)
    st.divider()
    if st.button("グラフ全体をPDFに出力", type="primary"):
        try:
            with st.spinner("グラフと指標カードをPDFにまとめています..."):
                pdf_bytes = create_dashboard_pdf(region_label, start_date, end_date, latest_month, cards, trend_df, compare_df, composition_df)
            st.session_state["dashboard_pdf_bytes"] = pdf_bytes
            st.session_state["dashboard_pdf_signature"] = pdf_signature
        except Exception as exc:
            st.error("PDFの作成に失敗しました。")
            st.code(str(exc))
    if st.session_state.get("dashboard_pdf_signature") == pdf_signature and st.session_state.get("dashboard_pdf_bytes"):
        safe_region = re.sub(r"[^0-9A-Za-zぁ-んァ-ヶ一-龠_-]+", "_", selected_region)
        st.download_button("PDFをダウンロード", st.session_state["dashboard_pdf_bytes"], file_name=f"住宅着工統計_{safe_region}_{start_date:%Y%m}_{end_date:%Y%m}.pdf", mime="application/pdf", use_container_width=True)

    with st.expander("データ出典（e-Stat）"):
        st.markdown(f"[住宅着工統計のExcel一覧を開く]({ESTAT_HOUSING_LIST_URL})")
        st.markdown(f"[今回アプリが取得したExcelを開く]({excel_url})")

    st.subheader("AIによる傾向分析")
    try:
        groq_api_key = st.secrets.get("GROQ_API_KEY", "")
        groq_model = st.secrets.get("GROQ_MODEL", "openai/gpt-oss-20b")
    except Exception:
        groq_api_key = ""
        groq_model = "openai/gpt-oss-20b"
    analysis_signature = f"{selected_region}|{start_date}|{end_date}"
    if st.session_state.get("ai_analysis_signature") != analysis_signature:
        st.session_state.pop("ai_analysis_result", None)
        st.session_state.pop("ai_analysis_signature", None)
    filtered_all = period_rows[(period_rows["年月"].dt.date >= start_date) & (period_rows["年月"].dt.date <= end_date)].copy()
    if not groq_api_key:
        st.info("AI分析を利用するには、Streamlit Cloud の App settings → Secrets に GROQ_API_KEY を設定してください。")
    else:
        if st.button("選択した地域の全指標をAI分析する"):
            try:
                with st.spinner("AIが分析しています..."):
                    analysis_text = analyze_with_groq(filtered_all, groq_api_key, groq_model)
                st.session_state["ai_analysis_result"] = analysis_text
                st.session_state["ai_analysis_signature"] = analysis_signature
            except Exception as exc:
                st.error(f"AI分析に失敗しました: {exc}")
        if st.session_state.get("ai_analysis_signature") == analysis_signature and st.session_state.get("ai_analysis_result"):
            st.markdown(st.session_state["ai_analysis_result"])


elif analysis_view == "住宅着工の詳細":
    st.markdown(f"### {region_label}　住宅着工の詳細")
    value_mode = st.radio("表示する数値", ["着工戸数", "前年比（％）"], horizontal=True, key="detail_value_mode")
    show_df = region_data[(region_data["年月"].dt.date >= start_date) & (region_data["年月"].dt.date <= end_date)].copy()
    metrics_to_show = [m for m in metrics_order if m in set(show_df["指標"].unique())]
    metric_pairs = [("住宅着工戸数", "総計"), ("持家", "持家"), ("貸家", "貸家"),
                    ("給与住宅", "給与住宅"), ("分譲住宅", "分譲住宅"),
                    ("分譲マンション（内数）", "分譲マンション"), ("分譲一戸建て（内数）", "分譲一戸建て")]
    detail_specs = []
    cols = st.columns(2, gap="small")
    for idx, (metric, display) in enumerate(metric_pairs):
        subset = show_df[show_df["指標"] == metric].copy()
        if subset.empty:
            continue
        ycol = "値" if value_mode == "着工戸数" else "前年比（％）"
        if ycol not in subset:
            continue
        subset[ycol] = pd.to_numeric(subset[ycol], errors="coerce")
        subset = subset.dropna(subset=[ycol]).sort_values("年月")
        if subset.empty:
            with cols[idx % 2]:
                st.markdown(f'<div class="panel-title">{display}</div>', unsafe_allow_html=True)
                st.info("この項目は選択期間に表示できる値がありません。")
            continue
        fig = px.line(subset, x="年月", y=ycol, markers=False,
                      labels={"年月": "年月", ycol: "着工戸数（戸）" if ycol == "値" else "前年比（％）"})
        fig.update_traces(line=dict(color=metric_colors.get(metric, "#3478f6"), width=2))
        fig.update_layout(height=280, margin=dict(l=4, r=8, t=8, b=4), showlegend=False,
                          xaxis_title=None, yaxis_title="着工戸数（戸）" if ycol == "値" else "前年比（％）")
        fig.update_xaxes(tickformat="%Y年%m月", showgrid=True, gridcolor="#e8edf3")
        fig.update_yaxes(tickformat=",.0f" if ycol == "値" else ".1f", showgrid=True, gridcolor="#e8edf3")
        with cols[idx % 2]:
            st.markdown(f'<div class="panel-title">{display}</div>', unsafe_allow_html=True)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        pdf_subset = subset[["年月", "指標", ycol]].rename(columns={ycol: "値"}).copy()
        detail_specs.append({"title": display, "data": pdf_subset,
                             "kind": "line", "unit": "戸" if ycol == "値" else "%"})
    detail_pdf_signature = f"{selected_region}|{start_date}|{end_date}|{value_mode}"
    if st.session_state.get("detail_pdf_signature") != detail_pdf_signature:
        st.session_state.pop("detail_pdf_bytes", None)
        st.session_state.pop("detail_pdf_signature", None)
    if st.button("住宅着工の詳細をPDFに出力", type="primary", key="detail_pdf_button"):
        try:
            st.session_state["detail_pdf_bytes"] = create_multi_chart_pdf(
                f"住宅着工の詳細（{value_mode}）", region_label,
                f"{start_date:%Y年%m月}～{end_date:%Y年%m月}", detail_specs)
            st.session_state["detail_pdf_signature"] = detail_pdf_signature
        except Exception as exc:
            st.error(f"PDFの作成に失敗しました: {exc}")
    if st.session_state.get("detail_pdf_signature") == detail_pdf_signature and st.session_state.get("detail_pdf_bytes"):
        st.download_button("詳細グラフPDFをダウンロード", st.session_state["detail_pdf_bytes"],
                           file_name=f"住宅着工詳細_{region_label}.pdf", mime="application/pdf", key="detail_pdf_download")

elif analysis_view == "構造・建て方分析":
    st.markdown(f"### {region_label}　構造・建て方分析")
    st.caption("住宅の利用関係・構造・建て方を組み合わせた都道府県別の年次統計です。")
    try:
        with st.spinner("e-Statから構造・建て方の年次統計を取得しています（初回のみ時間がかかります）..."):
            structure_raw, structure_file_url, structure_info = load_regional_source(
                ESTAT_HOUSING_ANNUAL_URL,
                ("構造別", "建て方別", "都道府県別", "戸数"),
                (), "年次",
                "e-Stat 住宅利用関係別・構造別・建て方別・都道府県別戸数", "戸")
        structure_raw = with_national_aggregate(structure_raw)
        structure_region = structure_raw[structure_raw["地域"] == selected_region].copy()
        structure_region = structure_region[
            (structure_region["年月"].dt.year >= start_year) & (structure_region["年月"].dt.year <= end_year)
        ]
        if structure_region.empty:
            st.warning("選択地域・期間に該当する構造・建て方データがありません。")
        else:
            structure_units = sorted(structure_region["単位"].dropna().astype(str).unique().tolist())
            if len(structure_units) > 1:
                structure_unit = st.selectbox("表示する単位", structure_units, key="structure_unit")
                structure_region = structure_region[structure_region["単位"].astype(str) == structure_unit].copy()
            structure_unit_label = str(structure_region["単位"].iloc[0]) if not structure_region.empty else "値"
            available_structure_years = sorted(structure_region["年月"].dt.year.unique().tolist())
            if available_structure_years:
                selected_structure_year = st.selectbox("対象年", available_structure_years,
                                                       index=len(available_structure_years)-1,
                                                       format_func=lambda y: f"{y}年", key="structure_year")
                latest_structure = structure_region[structure_region["年月"].dt.year == selected_structure_year].copy()
                structure_labels = sorted(latest_structure["指標"].astype(str).unique().tolist())
                likely = choose_default_metrics(structure_labels, max_count=10,
                                                keywords=("木造", "鉄骨造", "鉄筋コンクリート", "鉄骨鉄筋", "持家", "貸家", "分譲", "共同住宅", "一戸建"))
                selected_structure_metrics = st.multiselect("比較する区分", structure_labels, default=likely,
                                                             key="selected_structure_metrics")
                visible = latest_structure[latest_structure["指標"].isin(selected_structure_metrics)].copy()
                if not visible.empty:
                    # Sort horizontal chart labels by value for readable category comparisons.
                    structure_fig = px.bar(visible.sort_values("値", ascending=True), x="値", y="指標", orientation="h",
                                           labels={"値": f"値（{structure_unit_label}）", "指標": "構造・建て方・利用関係"})
                    structure_fig.update_layout(height=max(380, min(760, 35 * len(visible))), margin=dict(l=8, r=12, t=8, b=8), showlegend=False)
                    st.plotly_chart(structure_fig, use_container_width=True, config={"displayModeBar": False})
                if structure_region["年月"].dt.year.nunique() > 1 and selected_structure_metrics:
                    trend_struct = structure_region[structure_region["指標"].isin(selected_structure_metrics)].copy()
                    struct_trend_fig = px.line(trend_struct, x="年月", y="値", color="指標", markers=True,
                                                labels={"年月": "年", "値": f"値（{structure_unit_label}）", "指標": "区分"})
                    struct_trend_fig.update_layout(height=350, margin=dict(l=5,r=8,t=8,b=5), xaxis_title="年", yaxis_title=f"値（{structure_unit_label}）")
                    struct_trend_fig.update_xaxes(tickformat="%Y年")
                    st.plotly_chart(struct_trend_fig, use_container_width=True, config={"displayModeBar": False})
                specs = [{"title": f"構造・建て方別比較（{selected_structure_year}年）", "data": visible, "kind": "bar", "unit": structure_unit_label}]
                if structure_region["年月"].dt.year.nunique() > 1 and selected_structure_metrics:
                    specs.append({"title": "構造・建て方別 年次推移", "data": structure_region[structure_region["指標"].isin(selected_structure_metrics)], "kind": "line", "unit": structure_unit_label})
                structure_pdf_signature = f"{selected_region}|{start_year}|{end_year}|{selected_structure_year}|{structure_unit_label}|{'|'.join(selected_structure_metrics)}"
                if st.session_state.get("structure_pdf_signature") != structure_pdf_signature:
                    st.session_state.pop("structure_pdf_bytes", None)
                    st.session_state.pop("structure_pdf_signature", None)
                if st.button("構造・建て方分析をPDFに出力", type="primary", key="structure_pdf_button"):
                    try:
                        st.session_state["structure_pdf_bytes"] = create_multi_chart_pdf(
                            "住宅の構造・建て方分析", region_label,
                            f"{start_year}年～{end_year}年（年次）", specs)
                        st.session_state["structure_pdf_signature"] = structure_pdf_signature
                    except Exception as exc:
                        st.error(f"PDFの作成に失敗しました: {exc}")
                if st.session_state.get("structure_pdf_signature") == structure_pdf_signature and st.session_state.get("structure_pdf_bytes"):
                    st.download_button("構造・建て方PDFをダウンロード", st.session_state["structure_pdf_bytes"],
                                       file_name=f"構造建て方_{region_label}.pdf", mime="application/pdf", key="structure_pdf_download")
        with st.expander("データ出典"):
            st.markdown(f"[住宅構造・建て方別の年次統計一覧]({ESTAT_HOUSING_ANNUAL_URL})")
            st.markdown(f"[取得したExcelを開く]({structure_file_url})")
    except Exception as exc:
        st.error("構造・建て方の統計を読み込めませんでした。住宅着工の概況画面は引き続き利用できます。")
        with st.expander("エラー詳細"):
            st.code(str(exc))
        st.markdown(f"[e-Statの住宅年次統計一覧を開く]({ESTAT_HOUSING_ANNUAL_URL})")

elif analysis_view == "建築規模・非住宅":
    st.markdown(f"### {region_label}　建築規模・非住宅")
    st.markdown("#### 住宅の床面積（都道府県別）")
    try:
        with st.spinner("e-Statから住宅床面積の統計を取得しています..."):
            housing_area, housing_area_url, _ = load_regional_source(
                ESTAT_HOUSING_LIST_URL, ("都道府県別床面積",), (), "月次",
                "e-Stat 住宅都道府県別床面積", "㎡")
        render_regional_metric_section("住宅の床面積", housing_area, housing_area_url,
                                       selected_region, start_date, end_date, "housing_area",
                                       default_keywords=("総数", "床面積", "持家", "貸家"))
    except Exception as exc:
        st.warning("住宅の床面積データを読み込めませんでした。")
        with st.expander("エラー詳細（住宅床面積）"):
            st.code(str(exc))
        st.markdown(f"[e-Statの住宅床面積一覧を開く]({ESTAT_HOUSING_LIST_URL})")

    st.divider()
    st.markdown("#### 建築物の床面積（都道府県別）")
    try:
        with st.spinner("e-Statから建築物床面積の統計を取得しています..."):
            building_area, building_area_url, _ = load_regional_source(
                ESTAT_BUILDING_MONTHLY_URL, ("都道府県別床面積",), (), "月次",
                "e-Stat 建築物都道府県別床面積", "㎡")
        render_regional_metric_section("建築物の床面積", building_area, building_area_url,
                                       selected_region, start_date, end_date, "building_area",
                                       default_keywords=("総数", "床面積", "用途"))
    except Exception as exc:
        st.warning("建築物の床面積データを読み込めませんでした。")
        with st.expander("エラー詳細（建築物床面積）"):
            st.code(str(exc))
        st.markdown(f"[e-Statの建築物統計一覧を開く]({ESTAT_BUILDING_MONTHLY_URL})")

    st.divider()
    st.markdown("#### 建築物の用途・構造別分析（年度計）")
    try:
        with st.spinner("e-Statから建築物の用途・構造別年次統計を取得しています..."):
            building_struct, building_struct_url, _ = load_regional_source(
                ESTAT_BUILDING_ANNUAL_URL, ("構造別", "用途別", "都道府県別"), (), "年次",
                "e-Stat 建築物構造別・用途別・都道府県別統計", "値")
        building_struct = with_national_aggregate(building_struct)
        b_region = building_struct[building_struct["地域"] == selected_region].copy()
        b_region = b_region[(b_region["年月"].dt.year >= start_year) & (b_region["年月"].dt.year <= end_year)]
        if b_region.empty:
            st.info("選択地域・期間の年次データはありません。")
        else:
            building_units = sorted(b_region["単位"].dropna().astype(str).unique().tolist())
            if len(building_units) > 1:
                building_unit = st.selectbox("表示する単位（建築物）", building_units, key="building_struct_unit")
                b_region = b_region[b_region["単位"].astype(str) == building_unit].copy()
            b_years = sorted(b_region["年月"].dt.year.unique().tolist())
            b_year = st.selectbox("対象年度（建築物）", b_years, index=len(b_years)-1, format_func=lambda y: f"{y}年度", key="building_struct_year")
            b_latest = b_region[b_region["年月"].dt.year == b_year]
            b_labels = sorted(b_latest["指標"].astype(str).unique().tolist())
            b_defaults = choose_default_metrics(b_labels, 8, ("非住宅", "店舗", "事務所", "工場", "倉庫", "鉄骨", "鉄筋", "木造"))
            b_selected = st.multiselect("比較する用途・構造", b_labels, default=b_defaults, key="building_struct_metrics")
            b_visible = b_latest[b_latest["指標"].isin(b_selected)].copy()
            if not b_visible.empty:
                b_unit_label = str(b_visible["単位"].iloc[0]) if "単位" in b_visible.columns else "値"
                scale_factor, display_unit = _chart_display_unit(b_visible["値"], b_unit_label)
                b_visible["表示項目"] = b_visible["指標"].map(lambda value: _compact_category_label(value, max_chars=24))
                b_visible["表示値"] = pd.to_numeric(b_visible["値"], errors="coerce") / scale_factor
                bfig = px.bar(
                    b_visible.sort_values("表示値"), x="表示値", y="表示項目", orientation="h",
                    custom_data=["指標", "値"],
                    labels={"表示値": f"値（{display_unit}）", "表示項目": "用途・構造"},
                )
                bfig.update_traces(
                    text=b_visible["表示値"].map(_format_data_label), texttemplate="%{text}", textposition="outside", cliponaxis=False,
                    textfont=dict(size=10),
                    hovertemplate="%{customdata[0]}<br>値：%{customdata[1]:,.0f} " + b_unit_label + "<extra></extra>"
                )
                bfig.update_xaxes(tickformat=",.0f", exponentformat="none", showexponent="none", separatethousands=True)
                bfig.update_layout(
                    height=max(380, min(720, len(b_visible) * 42)),
                    margin=dict(l=12, r=24, t=12, b=16), showlegend=False,
                    yaxis=dict(automargin=True, tickfont=dict(size=11)),
                )
                st.plotly_chart(bfig, use_container_width=True, config={"displayModeBar": False})
                st.caption("構造名の略記：S造＝鉄骨造、RC造＝鉄筋コンクリート造、SRC造＝鉄骨鉄筋コンクリート造、CB造＝コンクリートブロック造。グラフにカーソルを合わせると正式な項目名を確認できます。")
                building_specs = [{"title": f"建築物の用途・構造別比較（{b_year}年度）", "data": b_visible, "kind": "bar", "unit": b_unit_label}]
                building_pdf_signature = f"{selected_region}|{start_year}|{end_year}|{b_year}|{b_unit_label}|{'|'.join(b_selected)}"
                if st.session_state.get("building_pdf_signature") != building_pdf_signature:
                    st.session_state.pop("building_pdf_bytes", None)
                    st.session_state.pop("building_pdf_signature", None)
                if st.button("建築物の用途・構造別グラフをPDFに出力", key="building_pdf_button"):
                    try:
                        st.session_state["building_pdf_bytes"] = create_multi_chart_pdf(
                            "建築物の用途・構造別分析", region_label,
                            f"{start_year}年度～{end_year}年度", building_specs)
                        st.session_state["building_pdf_signature"] = building_pdf_signature
                    except Exception as exc:
                        st.error(f"PDFの作成に失敗しました: {exc}")
                if st.session_state.get("building_pdf_signature") == building_pdf_signature and st.session_state.get("building_pdf_bytes"):
                    st.download_button("建築物の用途・構造別PDFをダウンロード", st.session_state["building_pdf_bytes"],
                                       file_name=f"建築物用途構造_{region_label}.pdf", mime="application/pdf", key="building_pdf_download")
        with st.expander("データ出典"):
            st.markdown(f"[建築物の年度計統計一覧]({ESTAT_BUILDING_ANNUAL_URL})")
            st.markdown(f"[取得したExcelを開く]({building_struct_url})")
    except Exception as exc:
        st.warning("建築物の用途・構造別年次データを読み込めませんでした。")
        with st.expander("エラー詳細（建築物年次統計）"):
            st.code(str(exc))
        st.markdown(f"[e-Statの建築物年次統計一覧を開く]({ESTAT_BUILDING_ANNUAL_URL})")

elif analysis_view == "リフォーム市場":
    st.markdown(f"### {region_label}　リフォーム市場")
    st.caption("e-Statの施工地域別受注高を表示します。この統計表は参考表です。")
    try:
        try:
            estat_app_id = st.secrets.get("ESTAT_APP_ID", "")
        except Exception:
            estat_app_id = ""
        if not estat_app_id:
            st.info(
                "この統計はe-Statのデータベース形式で公開されており、Excel一覧からの直接取得には対応していません。"
                "自動取得には無料のe-Stat APIアプリケーションIDが必要です。"
            )
            st.markdown(f"[e-Statの施工地域別受注高を開く]({ESTAT_RENOVATION_TABLE_URL})")
            with st.expander("自動取得の設定方法"):
                st.markdown(
                    "1. [e-Stat API](https://www.e-stat.go.jp/api/) で利用登録し、アプリケーションIDを取得します。\n"
                    "2. Streamlit Community Cloud の App settings → Secrets に `ESTAT_APP_ID = \\\"取得したID\\\"` を追加します。\n"
                    "3. アプリを再起動すると、この画面で受注高のグラフを取得できます。"
                )
        else:
            with st.spinner("e-Stat APIからリフォーム市場データを取得しています..."):
                renovation_raw, renovation_source_url = load_renovation_api_data(estat_app_id)
            renovation_region = renovation_raw[renovation_raw["地域"] == selected_region].copy()
            renovation_region = renovation_region[
                (renovation_region["年月"].dt.year >= start_year) & (renovation_region["年月"].dt.year <= end_year)
            ]
            if renovation_region.empty:
                st.info("選択した地域・期間に該当するリフォーム市場データはありません。")
            else:
                years = sorted(renovation_region["年月"].dt.year.unique().tolist())
                selected_renovation_year = st.selectbox(
                    "対象年度", years, index=len(years)-1, format_func=lambda y: f"{y}年度", key="renovation_year")
                latest_reno = renovation_region[renovation_region["年月"].dt.year == selected_renovation_year]
                labels = sorted(latest_reno["指標"].astype(str).unique().tolist())
                defaults = choose_default_metrics(labels, 3, ("計", "住宅", "非住宅建築物"))
                reno_selected = st.multiselect("比較する項目", labels, default=defaults, key="renovation_metrics")
                visible = latest_reno[latest_reno["指標"].isin(reno_selected)]
                if not visible.empty:
                    reno_fig = px.bar(visible.sort_values("値"), x="値", y="指標", orientation="h",
                                      labels={"値": f"受注高（{visible['単位'].iloc[0]}）", "指標": "用途"})
                    reno_fig.update_layout(height=max(300, min(650, 38*len(visible))), margin=dict(l=5,r=8,t=8,b=5), showlegend=False)
                    st.plotly_chart(reno_fig, use_container_width=True, config={"displayModeBar": False})
                if renovation_region["年月"].dt.year.nunique() > 1 and reno_selected:
                    reno_trend = renovation_region[renovation_region["指標"].isin(reno_selected)]
                    rtfig = px.line(reno_trend, x="年月", y="値", color="指標", markers=True,
                                    labels={"年月": "年度", "値": "受注高", "指標": "用途"})
                    rtfig.update_layout(height=330, margin=dict(l=5,r=8,t=8,b=5))
                    rtfig.update_xaxes(tickformat="%Y年")
                    st.plotly_chart(rtfig, use_container_width=True, config={"displayModeBar": False})
                specs = [{"title": f"施工地域別受注高（{selected_renovation_year}年度）", "data": visible, "kind": "bar", "unit": visible["単位"].iloc[0] if not visible.empty else "億円"}]
                if renovation_region["年月"].dt.year.nunique() > 1 and reno_selected:
                    specs.append({"title": "施工地域別受注高の推移", "data": renovation_region[renovation_region["指標"].isin(reno_selected)], "kind": "line", "unit": "億円"})
                renovation_pdf_signature = f"{selected_region}|{start_year}|{end_year}|{selected_renovation_year}|{'|'.join(reno_selected)}"
                if st.session_state.get("renovation_pdf_signature") != renovation_pdf_signature:
                    st.session_state.pop("renovation_pdf_bytes", None)
                    st.session_state.pop("renovation_pdf_signature", None)
                if st.button("リフォーム市場をPDFに出力", type="primary", key="renovation_pdf_button"):
                    try:
                        st.session_state["renovation_pdf_bytes"] = create_multi_chart_pdf(
                            "リフォーム・リニューアル市場", region_label,
                            f"{start_year}年度～{end_year}年度", specs)
                        st.session_state["renovation_pdf_signature"] = renovation_pdf_signature
                    except Exception as exc:
                        st.error(f"PDFの作成に失敗しました: {exc}")
                if st.session_state.get("renovation_pdf_signature") == renovation_pdf_signature and st.session_state.get("renovation_pdf_bytes"):
                    st.download_button("リフォーム市場PDFをダウンロード", st.session_state["renovation_pdf_bytes"],
                                       file_name=f"リフォーム市場_{region_label}.pdf", mime="application/pdf", key="renovation_pdf_download")
            with st.expander("データ出典"):
                st.markdown(f"[e-Statの施工地域別受注高（参考表2）]({ESTAT_RENOVATION_TABLE_URL})")
                st.caption("この地域別表は参考表です。調査対象者を地域別に抽出していないため、数値は参考として扱ってください。")
    except Exception as exc:
        st.error("e-Stat APIからリフォーム統計を取得できませんでした。ほかの分析画面は引き続き利用できます。")
        with st.expander("エラー詳細"):
            st.code(str(exc))
        st.markdown(f"[e-Statの施工地域別受注高（参考表2）を開く]({ESTAT_RENOVATION_TABLE_URL})")
elif analysis_view == "市区町村・住宅取得環境":
    render_municipal_market_page(selected_region)

