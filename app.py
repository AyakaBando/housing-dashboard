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


def make_pdf_chart_image(
    metric_frame: pd.DataFrame,
    display_measure: str,
    chart_type: str,
    start_date,
    end_date,
) -> io.BytesIO:
    """PDF用グラフを日本語ラベルで描画する。"""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from matplotlib.ticker import FuncFormatter
    # requirements.txt includes a package that provides the Japanese IPAex Gothic font.
    # Import after pyplot and explicitly set the family so PDF graph images do not use
    # the default Latin-only font on Streamlit Community Cloud.
    try:
        import matplotlib_fontja  # noqa: F401
        # matplotlib-fontja bundles IPAex Gothic and registers it for Matplotlib.
        plt.rcParams["font.family"] = "IPAexGothic"
    except Exception as exc:
        raise RuntimeError(
            "PDFグラフ用の日本語フォントを読み込めませんでした。"
            "requirements.txt に matplotlib-fontja が含まれていることを確認して再デプロイしてください。"
        ) from exc
    plt.rcParams["axes.unicode_minus"] = False

    value_col = "値" if display_measure == "着工戸数（戸）" else "前年比（％）"
    plot_data = metric_frame[["年月", value_col]].copy()
    plot_data["年月"] = pd.to_datetime(plot_data["年月"])
    plot_data[value_col] = pd.to_numeric(plot_data[value_col], errors="coerce")
    plot_data = plot_data.dropna(subset=["年月", value_col]).sort_values("年月")

    fig, ax = plt.subplots(figsize=(7.2, 3.0), dpi=150)
    if plot_data.empty:
        ax.text(0.5, 0.5, "選択期間にデータがありません", ha="center", va="center", transform=ax.transAxes)
        ax.set_xticks([])
        ax.set_yticks([])
    else:
        x = plot_data["年月"]
        y = plot_data[value_col]
        if chart_type == "棒":
            ax.bar(x, y, width=22)
        elif chart_type == "面":
            ax.fill_between(x, y, alpha=0.25)
            ax.plot(x, y, linewidth=1.4)
        else:
            ax.plot(x, y, linewidth=1.5)
        if value_col == "前年比（％）":
            ax.axhline(0, linewidth=0.9, linestyle="--")
            ax.set_ylabel("前年比（％）")
            ax.yaxis.set_major_formatter(FuncFormatter(lambda value, pos: f"{value:.0f}%"))
        else:
            ax.set_ylabel("着工戸数（戸）")
            ax.yaxis.set_major_formatter(FuncFormatter(lambda value, pos: f"{value:,.0f}"))
        span_months = max(0, (end_date.year - start_date.year) * 12 + end_date.month - start_date.month)
        if span_months > 36:
            ax.xaxis.set_major_locator(mdates.YearLocator())
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y年"))
        elif span_months > 12:
            ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y年%m月"))
        else:
            ax.xaxis.set_major_locator(mdates.MonthLocator(interval=max(1, span_months // 6 or 1)))
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y年%m月"))
        ax.set_xlabel("年月")
        ax.grid(True, axis="both", alpha=0.25)
        ax.tick_params(axis="x", labelrotation=25, labelsize=7)
        ax.tick_params(axis="y", labelsize=7)
    fig.tight_layout()
    output = io.BytesIO()
    fig.savefig(output, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    output.seek(0)
    return output


def create_dashboard_pdf(
    chart_items: list[tuple[str, pd.DataFrame]],
    region_label: str,
    display_measure: str,
    start_date,
    end_date,
    dataset_info: str,
) -> bytes:
    """Create a multi-page A4 landscape PDF with all visible indicator charts."""
    from math import ceil
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
    except Exception:
        # The registration may already exist in the running process.
        pass

    output = io.BytesIO()
    page_width, page_height = landscape(A4)
    pdf = canvas.Canvas(output, pagesize=(page_width, page_height), pageCompression=1)
    pdf.setTitle(f"住宅着工統計ダッシュボード - {region_label}")
    pdf.setAuthor("住宅着工統計ダッシュボード")

    margin = 26
    col_gap = 18
    col_width = (page_width - margin * 2 - col_gap) / 2
    image_height = 160
    charts_per_page = 4
    page_count = max(1, ceil(len(chart_items) / charts_per_page))
    note = "注：分譲マンション・分譲一戸建ては分譲住宅の内数です。総数・分譲住宅に加算しないでください。"

    for page_index in range(page_count):
        pdf.setFont("HeiseiKakuGo-W5", 15)
        pdf.drawString(margin, page_height - 29, "住宅着工統計ダッシュボード")
        pdf.setFont("HeiseiKakuGo-W5", 9)
        pdf.drawString(
            margin,
            page_height - 47,
            f"地域：{region_label}　期間：{start_date.strftime('%Y年%m月')}～{end_date.strftime('%Y年%m月')}　表示：{display_measure}",
        )
        pdf.setLineWidth(0.6)
        pdf.line(margin, page_height - 55, page_width - margin, page_height - 55)

        start_idx = page_index * charts_per_page
        page_charts = chart_items[start_idx:start_idx + charts_per_page]
        for local_index, (metric, frame) in enumerate(page_charts):
            col = local_index % 2
            row = local_index // 2
            x = margin + col * (col_width + col_gap)
            title_y = page_height - 74 - row * 218
            image_y = title_y - image_height - 8

            pdf.setFont("HeiseiKakuGo-W5", 10)
            pdf.drawString(x, title_y, metric)
            if metric in {"分譲マンション（内数）", "分譲一戸建て（内数）"}:
                pdf.setFont("HeiseiKakuGo-W5", 7)
                pdf.drawRightString(x + col_width, title_y, "※分譲住宅の内数")

            metric_chart_type = "棒" if metric in {"分譲マンション（内数）", "分譲一戸建て（内数）"} else "折れ線"
            image_buffer = make_pdf_chart_image(
                frame, display_measure, metric_chart_type, start_date, end_date
            )
            pdf.drawImage(
                ImageReader(image_buffer), x, image_y,
                width=col_width, height=image_height,
                preserveAspectRatio=True, anchor="c", mask="auto",
            )

        pdf.setFont("HeiseiKakuGo-W5", 7)
        pdf.drawString(margin, 35, note)
        pdf.setFont("HeiseiKakuGo-W5", 6.5)
        pdf.drawString(margin, 23, f"出典：e-Stat {dataset_info[:95]}")
        pdf.drawRightString(page_width - margin, 23, f"{page_index + 1} / {page_count} ページ")
        pdf.showPage()

    pdf.save()
    return output.getvalue()


st.set_page_config(page_title="住宅着工統計ダッシュボード", page_icon="📊", layout="wide")
st.title("住宅着工統計ダッシュボード")
st.caption("地域と対象期間を選ぶと、月次統計の各指標を複数のグラフで比較し、PDFで保存できます。")

with st.sidebar:
    st.header("① 地域を選択")
    if st.button("e-Statからデータを再取得", use_container_width=True, help="保存済みキャッシュを削除し、次回の実行でExcelを取り直します。"):
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

data = data.copy()
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
        help="全国を選ぶと47都道府県の値を合計します。都道府県を選ぶと、その地域だけを表示します。",
    )

if selected_region is None:
    st.info("まず、画面左側から「全国」または表示したい都道府県を選択してください。")
    st.stop()

if selected_region == "全国":
    region_data = build_national_series(data)
    region_label = "全国（47都道府県合計）"
else:
    region_data = data.loc[data["地域"] == selected_region].copy()
    region_label = selected_region
    if region_data.empty:
        st.error(f"{selected_region}のデータを取得できていません。")
        st.stop()

metrics = [
    "住宅着工戸数", "持家", "貸家", "給与住宅", "分譲住宅",
    "分譲マンション（内数）", "分譲一戸建て（内数）",
]
available_metrics = [m for m in metrics if m in region_data["指標"].unique()]
if not available_metrics:
    st.error("選択した地域で表示できる指標がありません。")
    st.stop()

# Build explicit Japanese year/month menus from the months actually present in the workbook.
# This makes 2025 a visible year option and avoids browser-specific date-input formatting.
available_months = sorted(
    pd.Timestamp(value).to_period("M").to_timestamp()
    for value in region_data["年月"].dropna().drop_duplicates().tolist()
)
period_years = sorted({month.year for month in available_months})
months_by_year = {
    year: sorted({month.month for month in available_months if month.year == year})
    for year in period_years
}
default_start_year = 2020 if 2020 in period_years else period_years[0]
default_end_year = period_years[-1]

with st.sidebar:
    st.divider()
    st.header("② 表示設定")
    display_measure = st.radio(
        "表示する数値",
        ["着工戸数（戸）", "前年比（％）"],
        index=0,
        help="選択した表示内容を、以下の全グラフとPDFに反映します。",
    )
    st.markdown("**開始年月**")
    start_year_col, start_month_col = st.columns(2)
    with start_year_col:
        start_year = st.selectbox(
            "開始年", options=period_years,
            index=period_years.index(default_start_year),
            format_func=lambda year: f"{year}年",
            key="period_start_year",
        )
    with start_month_col:
        start_month_options = months_by_year[start_year]
        preferred_start_month = 1 if 1 in start_month_options else start_month_options[0]
        start_month = st.selectbox(
            "開始月", options=start_month_options,
            index=start_month_options.index(preferred_start_month),
            format_func=lambda month: f"{month}月",
            key=f"period_start_month_{start_year}",
        )

    st.markdown("**終了年月**")
    end_year_col, end_month_col = st.columns(2)
    with end_year_col:
        end_year = st.selectbox(
            "終了年", options=period_years,
            index=period_years.index(default_end_year),
            format_func=lambda year: f"{year}年",
            key="period_end_year",
        )
    with end_month_col:
        end_month_options = months_by_year[end_year]
        preferred_end_month = 12 if 12 in end_month_options else end_month_options[-1]
        end_month = st.selectbox(
            "終了月", options=end_month_options,
            index=end_month_options.index(preferred_end_month),
            format_func=lambda month: f"{month}月",
            key=f"period_end_month_{end_year}",
        )
    st.caption(
        f"取得済み期間：{min_month:%Y年%m月}～{max_month:%Y年%m月} "
        f"（{period_years[0]}～{period_years[-1]}年）"
    )

start_date = date(int(start_year), int(start_month), 1)
end_date = date(int(end_year), int(end_month), 1)
if start_date > end_date:
    st.error("開始年月は終了年月以前を指定してください。")
    st.stop()

filtered_all = region_data[
    (region_data["年月"].dt.date >= start_date)
    & (region_data["年月"].dt.date <= end_date)
].copy().sort_values(["年月", "指標"])
if filtered_all.empty:
    st.warning("選択した地域・期間に該当するデータがありません。")
    st.stop()

value_col = "値" if display_measure == "着工戸数（戸）" else "前年比（％）"

# KPI cards summarise counts, while chart panels below show all available indicators simultaneously.
def latest_metric(metric: str, col: str = "値"):
    part = filtered_all.loc[filtered_all["指標"] == metric].dropna(subset=[col]).sort_values("年月")
    if part.empty:
        return None, None
    return part.iloc[-1][col], pd.Timestamp(part.iloc[-1]["年月"])

total_latest, latest_month = latest_metric("住宅着工戸数", "値")
rental_latest, _ = latest_metric("貸家", "値")
for_sale_latest, _ = latest_metric("分譲住宅", "値")

k1, k2, k3, k4 = st.columns(4)
k1.metric("表示地域", region_label)
k2.metric("最新月", latest_month.strftime("%Y年%m月") if latest_month is not None else "データなし")
k3.metric("住宅着工戸数", f"{total_latest:,.0f} 戸" if total_latest is not None else "データなし")
k4.metric("貸家", f"{rental_latest:,.0f} 戸" if rental_latest is not None else "データなし")

st.subheader(f"{region_label}｜指標別の推移")
st.caption(
    f"{start_date.strftime('%Y年%m月')}～{end_date.strftime('%Y年%m月')}　・　"
    f"{display_measure}　・　各指標を個別のグラフで表示しています。"
)

chart_items = []
for metric in available_metrics:
    metric_frame = filtered_all.loc[filtered_all["指標"] == metric].copy().sort_values("年月")
    chart_items.append((metric, metric_frame))

# Two-column dashboard layout: all monthly indicators appear in this single dashboard page.
figures = []
for idx in range(0, len(chart_items), 2):
    columns = st.columns(2)
    for col, item in zip(columns, chart_items[idx:idx + 2]):
        metric, metric_frame = item
        plot_data = metric_frame.dropna(subset=[value_col]).copy()
        with col:
            st.markdown(f"#### {metric}")
            if metric in {"分譲マンション（内数）", "分譲一戸建て（内数）"}:
                st.caption("分譲住宅の内数です。分譲住宅に加算しないでください。")
            if plot_data.empty:
                fig = px.line()
                fig.add_annotation(
                    text="選択期間に表示できるデータがありません。",
                    x=0.5, y=0.5, xref="paper", yref="paper", showarrow=False,
                )
            else:
                # 時系列の変化は折れ線、内数の指標は棒グラフで比較しやすく表示
                if metric in {"分譲マンション（内数）", "分譲一戸建て（内数）"}:
                    fig = px.bar(plot_data, x="年月", y=value_col)
                else:
                    fig = px.line(plot_data, x="年月", y=value_col, markers=False)
            span_months = max(0, (end_date.year - start_date.year) * 12 + end_date.month - start_date.month)
            if span_months > 36:
                fig.update_xaxes(tickformat="%Y年", dtick="M12")
            elif span_months > 12:
                fig.update_xaxes(tickformat="%Y年%m月", dtick="M3")
            else:
                fig.update_xaxes(tickformat="%Y年%m月", dtick="M1")
            if display_measure == "前年比（％）":
                fig.add_hline(y=0, line_dash="dash")
            fig.update_layout(
                height=310,
                margin=dict(l=10, r=10, t=15, b=10),
                showlegend=False,
                xaxis_title="年月",
                yaxis_title="前年比（％）" if display_measure == "前年比（％）" else "着工戸数（戸）",
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
            figures.append((metric, metric_frame))

if "分譲マンション（内数）" in available_metrics or "分譲一戸建て（内数）" in available_metrics:
    st.caption("※ 分譲マンション・分譲一戸建ては分譲住宅の内数です。総数や分譲住宅に加算しないでください。")

# Export PDF for all chart panels, spread over as many landscape A4 pages as needed.
pdf_signature = "|".join([
    selected_region, display_measure, str(start_date), str(end_date),
    ",".join(available_metrics),
])
if st.session_state.get("dashboard_pdf_signature") != pdf_signature:
    st.session_state.pop("dashboard_pdf_bytes", None)
    st.session_state.pop("dashboard_pdf_signature", None)

st.divider()
st.subheader("グラフ全体をPDFに保存")
st.write("上に表示されているすべての指標グラフを、横向きA4の複数ページPDFにまとめます。")
if st.button("グラフ全体のPDFを作成", type="primary"):
    try:
        with st.spinner("全グラフをPDFにまとめています..."):
            pdf_bytes = create_dashboard_pdf(
                chart_items=chart_items,
                region_label=region_label,
                display_measure=display_measure,
                start_date=start_date,
                end_date=end_date,
                dataset_info=dataset_info,
            )
        st.session_state["dashboard_pdf_bytes"] = pdf_bytes
        st.session_state["dashboard_pdf_signature"] = pdf_signature
        st.success("PDFを作成しました。下のボタンから保存できます。")
    except Exception as exc:
        st.error("PDFの作成に失敗しました。")
        st.code(str(exc))

if (
    st.session_state.get("dashboard_pdf_signature") == pdf_signature
    and st.session_state.get("dashboard_pdf_bytes")
):
    safe_region = re.sub(r"[^0-9A-Za-zぁ-んァ-ヶ一-龠_-]+", "_", selected_region)
    safe_measure = "戸数" if display_measure == "着工戸数（戸）" else "前年比"
    st.download_button(
        "PDFをダウンロード",
        data=st.session_state["dashboard_pdf_bytes"],
        file_name=f"住宅着工統計_{safe_region}_{safe_measure}_{start_date:%Y%m}_{end_date:%Y%m}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

with st.expander("データ取得情報"):
    st.write("取得元：e-Stat「建築着工統計調査 / 住宅着工統計 / 時系列表」")
    st.write("検出した統計表：", dataset_info)
    st.write("ExcelダウンロードURL：", excel_url)
    st.write("取得期間：", min_date, "～", max_date)
    st.write("表示地域：", region_label)
    st.write("グラフ数：", len(chart_items))
    st.caption("全国は47都道府県のデータを合計しています。都道府県別・月次Excelで確認できる指標を表示しています。")
    st.caption("木造・RC造などの構造別データは別の統計表が必要なため、現時点のグラフにはまだ含めていません。")

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

if not groq_api_key:
    st.info("AI分析を利用するには、Streamlit Cloud の App settings → Secrets に GROQ_API_KEY を設定してください。")
else:
    if st.button("選択した地域の全指標をAI分析する"):
        try:
            with st.spinner("AIが分析しています。文章が途中で切れた場合は自動で続きを取得します..."):
                analysis_text = analyze_with_groq(filtered_all, groq_api_key, groq_model)
            st.session_state["ai_analysis_result"] = analysis_text
            st.session_state["ai_analysis_signature"] = analysis_signature
        except Exception as exc:
            st.error(f"AI分析に失敗しました: {exc}")
    if (
        st.session_state.get("ai_analysis_signature") == analysis_signature
        and st.session_state.get("ai_analysis_result")
    ):
        st.markdown(st.session_state["ai_analysis_result"])
