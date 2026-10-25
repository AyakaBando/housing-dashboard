import io
import re
from datetime import date, datetime
from urllib.parse import urljoin

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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


def make_pdf_chart_image(kind, trend_df, compare_df, composition_df):
    """Render one of the three overview charts as a Japanese-labelled PNG."""
    import io
    import matplotlib.dates as mdates
    from matplotlib.ticker import FuncFormatter

    plt = setup_matplotlib_japanese()
    fig, ax = plt.subplots(figsize=(5.1, 4.0), dpi=150)

    if kind == "trend":
        if trend_df.empty:
            ax.text(0.5, 0.5, "表示できるデータがありません", ha="center", va="center", transform=ax.transAxes)
            ax.set_axis_off()
        else:
            x = pd.to_datetime(trend_df["年月"])
            ax.plot(x, trend_df["値"], label="月別実績", color="#3478f6", linewidth=1.5)
            ax.plot(x, trend_df["年度累計"], label="年度累計", color="#32a879", linewidth=1.5)
            ax.plot(x, trend_df["移動年計"], label="移動年計（12か月）", color="#ef5350", linewidth=1.5)
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
            ax.bar([p - width / 2 for p in pos], compare_df["前年値"], width=width, label=compare_df.attrs.get("前年ラベル", "前年"), color="#a6b1bd")
            ax.bar([p + width / 2 for p in pos], compare_df["最新年値"], width=width, label=compare_df.attrs.get("最新年ラベル", "最新年"), color="#3478f6")
            ax.set_xticks(pos, months)
            ax.set_ylabel("着工戸数（戸）")
            ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:,.0f}"))
            ax.legend(loc="upper left", fontsize=8, frameon=False)
            ax.grid(True, axis="y", alpha=0.22)
    elif kind == "composition":
        if composition_df.empty or composition_df["値"].sum() <= 0:
            ax.text(0.5, 0.5, "構成比を表示できるデータがありません", ha="center", va="center", transform=ax.transAxes)
            ax.set_axis_off()
        else:
            ax.pie(
                composition_df["値"], labels=composition_df["表示名"], autopct="%1.1f%%",
                startangle=90, counterclock=False, pctdistance=0.77,
                wedgeprops={"width": 0.42, "edgecolor": "white"},
                textprops={"fontsize": 8},
            )
            ax.set_aspect("equal")

    fig.tight_layout(pad=1.2)
    output = io.BytesIO()
    fig.savefig(output, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    output.seek(0)
    return output


def create_dashboard_pdf(region_label, start_date, end_date, latest_month, cards, trend_df, compare_df, composition_df):
    """Export the dashboard overview (six KPI cards and three charts) to one wide PDF page."""
    import io
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
    except Exception:
        pass

    # A wide dashboard page is closer to the on-screen comparison layout than a text report.
    page_width, page_height = 1200, 675
    output = io.BytesIO()
    pdf = canvas.Canvas(output, pagesize=(page_width, page_height), pageCompression=1)
    pdf.setTitle(f"住宅着工統計ダッシュボード - {region_label}")
    pdf.setAuthor("住宅着工統計ダッシュボード")
    font = "HeiseiKakuGo-W5"
    margin = 24

    pdf.setFont(font, 17)
    pdf.drawString(margin, 648, "住宅着工統計ダッシュボード")
    pdf.setFont(font, 9)
    pdf.drawString(margin, 630, f"地域：{region_label}　対象期間：{start_date:%Y年%m月}～{end_date:%Y年%m月}　最新月：{latest_month:%Y年%m月}")

    # Six concise KPI cards across the top.
    card_gap = 8
    card_width = (page_width - 2 * margin - card_gap * 5) / 6
    card_y, card_h = 548, 68
    for i, card in enumerate(cards):
        x = margin + i * (card_width + card_gap)
        pdf.setFillColor(colors.HexColor("#ffffff"))
        pdf.setStrokeColor(colors.HexColor("#d9e0e8"))
        pdf.roundRect(x, card_y, card_width, card_h, 5, stroke=1, fill=1)
        pdf.setFillColor(colors.HexColor(card["color"]))
        pdf.rect(x, card_y + card_h - 4, card_width, 4, stroke=0, fill=1)
        pdf.setFillColor(colors.HexColor("#334155"))
        pdf.setFont(font, 8)
        pdf.drawString(x + 8, card_y + 49, card["label"])
        pdf.setFont(font, 15)
        pdf.drawString(x + 8, card_y + 29, card["value"])
        pdf.setFont(font, 6.8)
        pdf.setFillColor(colors.HexColor("#64748b"))
        pdf.drawString(x + 8, card_y + 13, card["previous"])
        pdf.setFillColor(colors.HexColor(card["yoy_color"]))
        pdf.drawRightString(x + card_width - 7, card_y + 13, card["yoy"])

    panel_gap = 12
    panel_width = (page_width - 2 * margin - panel_gap * 2) / 3
    image_y, image_h = 105, 410
    panel_titles = ["住宅着工戸数の推移", "月別比較", "最新月の構成比"]
    kinds = ["trend", "compare", "composition"]
    for idx, kind in enumerate(kinds):
        x = margin + idx * (panel_width + panel_gap)
        pdf.setFont(font, 10)
        pdf.setFillColor(colors.HexColor("#263746"))
        if kind == "compare":
            title = "月別比較（" + compare_df.attrs.get("前年ラベル", "前年") + "・" + compare_df.attrs.get("最新年ラベル", "最新年") + "）"
        elif kind == "composition":
            title = f"利用関係別構成比（{latest_month:%Y年%m月}）"
        else:
            title = panel_titles[idx]
        pdf.drawString(x, 530, title)
        image = make_pdf_chart_image(kind, trend_df, compare_df, composition_df)
        pdf.drawImage(ImageReader(image), x, image_y, width=panel_width, height=image_h, preserveAspectRatio=True, anchor="c", mask="auto")
        pdf.setStrokeColor(colors.HexColor("#d9e0e8"))
        pdf.roundRect(x, image_y, panel_width, image_h, 3, stroke=1, fill=0)

    pdf.setFont(font, 7)
    pdf.setFillColor(colors.HexColor("#64748b"))
    pdf.drawString(margin, 30, "出典：e-Stat 建築着工統計調査")
    pdf.drawRightString(page_width - margin, 30, "住宅着工統計")
    pdf.save()
    return output.getvalue()


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
@media (max-width: 900px) {.kpi-value {font-size:18px;} .kpi-card {padding:7px; min-height:104px;} .kpi-label {font-size:11px;}}
</style>
""", unsafe_allow_html=True)
st.title("住宅着工統計ダッシュボード")

with st.sidebar:
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
        trend_fig.add_trace(go.Scatter(x=trend_df["年月"], y=trend_df["値"], name="月別実績", mode="lines", line=dict(color="#3478f6", width=2)))
        trend_fig.add_trace(go.Scatter(x=trend_df["年月"], y=trend_df["年度累計"], name="年度累計", mode="lines", line=dict(color="#29a66f", width=2)))
        trend_fig.add_trace(go.Scatter(x=trend_df["年月"], y=trend_df["移動年計"], name="移動年計", mode="lines", line=dict(color="#ef5058", width=2)))
    trend_fig.update_layout(height=390, margin=dict(l=0,r=4,t=12,b=5), legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=10)), xaxis_title=None, yaxis_title="戸数（戸）", hovermode="x unified")
    trend_fig.update_xaxes(tickformat="%Y年%m月", nticks=7, showgrid=True, gridcolor="#e8edf3")
    trend_fig.update_yaxes(tickformat=",.0f", showgrid=True, gridcolor="#e8edf3")
    st.plotly_chart(trend_fig, use_container_width=True, config={"displayModeBar": False})

with chart_cols[1]:
    st.markdown(f'<div class="panel-title">月別比較（{previous_fiscal_year}年度・{latest_fiscal_year}年度）</div>', unsafe_allow_html=True)
    compare_fig = go.Figure()
    compare_fig.add_trace(go.Bar(x=compare_df["月ラベル"], y=compare_df["前年値"], name=f"{previous_fiscal_year}年度", marker_color="#aeb8c3"))
    compare_fig.add_trace(go.Bar(x=compare_df["月ラベル"], y=compare_df["最新年値"], name=f"{latest_fiscal_year}年度", marker_color="#3478f6"))
    compare_fig.update_layout(barmode="group", height=390, margin=dict(l=0,r=4,t=12,b=5), legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=10)), xaxis_title=None, yaxis_title="戸数（戸）")
    compare_fig.update_yaxes(tickformat=",.0f", showgrid=True, gridcolor="#e8edf3")
    compare_fig.update_xaxes(showgrid=False)
    st.plotly_chart(compare_fig, use_container_width=True, config={"displayModeBar": False})

with chart_cols[2]:
    st.markdown(f'<div class="panel-title">利用関係別構成比（{latest_month:%Y年%m月}）</div>', unsafe_allow_html=True)
    if not composition_df.empty and composition_df["値"].sum() > 0:
        composition_fig = px.pie(composition_df, values="値", names="表示名", hole=0.48, color="表示名", color_discrete_map={"持家":"#35a879", "貸家":"#f59a43", "給与住宅":"#9b59d0", "分譲：マンション":"#e95059", "分譲：戸建て":"#22a6c7"})
        composition_fig.update_traces(textposition="inside", textinfo="label+percent", insidetextfont=dict(color="white", size=11), hovertemplate="%{label}<br>%{value:,.0f}戸 (%{percent})<extra></extra>")
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
