import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
import xlrd
from bs4 import BeautifulSoup
from urllib.parse import urljoin

st.set_page_config(
    page_title="住宅着工統計ダッシュボード",
    page_icon="🏠",
    layout="wide",
)

SOURCE_PAGE = (
    "https://www.mlit.go.jp/sogoseisaku/jouhouka/"
    "sosei_jouhouka_tk4_000002.html"
)
CACHE_SECONDS = 6 * 60 * 60
START_DATE = pd.Timestamp("2020-01-01")

PREFECTURE_ALIASES = {
    "北海道": "北海道",
    "青森": "青森県", "青森県": "青森県",
    "岩手": "岩手県", "岩手県": "岩手県",
    "宮城": "宮城県", "宮城県": "宮城県",
    "秋田": "秋田県", "秋田県": "秋田県",
    "山形": "山形県", "山形県": "山形県",
    "福島": "福島県", "福島県": "福島県",
    "茨城": "茨城県", "茨城県": "茨城県",
    "栃木": "栃木県", "栃木県": "栃木県",
    "群馬": "群馬県", "群馬県": "群馬県",
    "埼玉": "埼玉県", "埼玉県": "埼玉県",
    "千葉": "千葉県", "千葉県": "千葉県",
    "東京": "東京都", "東京都": "東京都",
    "神奈川": "神奈川県", "神奈川県": "神奈川県",
    "新潟": "新潟県", "新潟県": "新潟県",
    "富山": "富山県", "富山県": "富山県",
    "石川": "石川県", "石川県": "石川県",
    "福井": "福井県", "福井県": "福井県",
    "山梨": "山梨県", "山梨県": "山梨県",
    "長野": "長野県", "長野県": "長野県",
    "岐阜": "岐阜県", "岐阜県": "岐阜県",
    "静岡": "静岡県", "静岡県": "静岡県",
    "愛知": "愛知県", "愛知県": "愛知県",
    "三重": "三重県", "三重県": "三重県",
    "滋賀": "滋賀県", "滋賀県": "滋賀県",
    "京都": "京都府", "京都府": "京都府",
    "大阪": "大阪府", "大阪府": "大阪府",
    "兵庫": "兵庫県", "兵庫県": "兵庫県",
    "奈良": "奈良県", "奈良県": "奈良県",
    "和歌山": "和歌山県", "和歌山県": "和歌山県",
    "鳥取": "鳥取県", "鳥取県": "鳥取県",
    "島根": "島根県", "島根県": "島根県",
    "岡山": "岡山県", "岡山県": "岡山県",
    "広島": "広島県", "広島県": "広島県",
    "山口": "山口県", "山口県": "山口県",
    "徳島": "徳島県", "徳島県": "徳島県",
    "香川": "香川県", "香川県": "香川県",
    "愛媛": "愛媛県", "愛媛県": "愛媛県",
    "高知": "高知県", "高知県": "高知県",
    "福岡": "福岡県", "福岡県": "福岡県",
    "佐賀": "佐賀県", "佐賀県": "佐賀県",
    "長崎": "長崎県", "長崎県": "長崎県",
    "熊本": "熊本県", "熊本県": "熊本県",
    "大分": "大分県", "大分県": "大分県",
    "宮崎": "宮崎県", "宮崎県": "宮崎県",
    "鹿児島": "鹿児島県", "鹿児島県": "鹿児島県",
    "沖縄": "沖縄県", "沖縄県": "沖縄県",
}


def normalize_text(value):
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"\s+", "", text).strip()


def number_or_none(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if pd.isna(value):
            return None
        return int(value)
    text = normalize_text(value).replace(",", "")
    if not text or text in {"-", "－", "…", "－"}:
        return None
    try:
        return int(float(text))
    except (TypeError, ValueError):
        return None


def parse_month_from_sheet(sheet):
    # The report title is normally in the first several rows, e.g. 令和８年８月分...
    pattern = re.compile(r"(令和|平成)\s*(元|\d+)年\s*(\d+)月")
    for row_index in range(min(sheet.nrows, 12)):
        for col_index in range(min(sheet.ncols, 8)):
            value = sheet.cell_value(row_index, col_index)
            if not isinstance(value, str):
                continue
            normalized = unicodedata.normalize("NFKC", value)
            match = pattern.search(normalized)
            if not match:
                continue
            era, year_text, month_text = match.groups()
            era_year = 1 if year_text == "元" else int(year_text)
            month = int(month_text)
            year = (2018 + era_year) if era == "令和" else (1988 + era_year)
            if 1 <= month <= 12:
                return pd.Timestamp(year=year, month=month, day=1)
    return None


def find_latest_workbook_url():
    response = requests.get(
        SOURCE_PAGE,
        timeout=60,
        headers={"User-Agent": "Mozilla/5.0 housing-dashboard/1.0"},
    )
    response.raise_for_status()
    soup = BeautifulSoup(response.content, "html.parser")

    for anchor in soup.find_all("a", href=True):
        label = normalize_text(anchor.get_text(" ", strip=True))
        if "都道府県別着工戸数" in label:
            return urljoin(SOURCE_PAGE, anchor["href"])

    raise RuntimeError(
        "国土交通省の公式ページから「住宅 都道府県別着工戸数」のExcelリンクを"
        "見つけられませんでした。公式ページのリンク名が変更された可能性があります。"
    )


@st.cache_data(ttl=CACHE_SECONDS, show_spinner=False)
def load_housing_data():
    file_url = find_latest_workbook_url()
    response = requests.get(
        file_url,
        timeout=120,
        headers={"User-Agent": "Mozilla/5.0 housing-dashboard/1.0"},
    )
    response.raise_for_status()
    file_bytes = response.content

    if len(file_bytes) < 10_000:
        raise RuntimeError(
            "ダウンロードしたファイルのサイズが小さすぎます。"
            "公式サイトがExcelではなくエラーページを返した可能性があります。"
        )

    workbook = xlrd.open_workbook(file_contents=file_bytes, on_demand=True)
    rows = []
    skipped_sheets = []

    for sheet_name in workbook.sheet_names():
        sheet = workbook.sheet_by_name(sheet_name)
        month = parse_month_from_sheet(sheet)
        if month is None or month < START_DATE:
            continue

        prefecture_values = {}
        national_value = None

        # In the official workbook, prefecture names are in column B (index 1)
        # and total new housing starts are in column C (index 2).
        # This matches the attached kenbetu808.xls workbook.
        for row_index in range(sheet.nrows):
            if sheet.ncols < 3:
                continue
            label = normalize_text(sheet.cell_value(row_index, 1))
            value = number_or_none(sheet.cell_value(row_index, 2))
            if value is None:
                continue

            if label == "合計":
                national_value = value
            elif label in PREFECTURE_ALIASES:
                prefecture = PREFECTURE_ALIASES[label]
                # Keep the first occurrence in case the sheet repeats a region.
                prefecture_values.setdefault(prefecture, value)

        if len(prefecture_values) < 47:
            skipped_sheets.append(
                f"{sheet_name}（都道府県の取得数: {len(prefecture_values)}）"
            )
            continue

        rows.append(
            {
                "年月": month,
                "地域": "全国",
                "住宅着工戸数": national_value
                if national_value is not None
                else sum(prefecture_values.values()),
            }
        )
        for prefecture, value in prefecture_values.items():
            rows.append(
                {
                    "年月": month,
                    "地域": prefecture,
                    "住宅着工戸数": value,
                }
            )

    workbook.release_resources()

    df = pd.DataFrame(rows)
    if df.empty:
        raise RuntimeError(
            "2020年1月以降の月次データを読み取れませんでした。"
            "Excelのシート構成が変更された可能性があります。"
        )

    df = df.drop_duplicates(subset=["年月", "地域"], keep="last")
    df = df.sort_values(["地域", "年月"]).reset_index(drop=True)

    # YoY is calculated against the same month in the previous year.
    df["前年同月戸数"] = df.groupby("地域")["住宅着工戸数"].shift(12)
    df["前年同月比（％）"] = (
        (df["住宅着工戸数"] / df["前年同月戸数"] - 1) * 100
    ).where(df["前年同月戸数"] > 0)

    latest_month = df["年月"].max()
    return df, latest_month, file_url, skipped_sheets


st.title("住宅着工統計ダッシュボード")
st.caption("出典：国土交通省「建築着工統計調査報告 時系列一覧」")
st.caption("対象：2020年1月〜最新公表月／新設住宅着工戸数")

try:
    with st.spinner("国土交通省の最新Excelを確認し、月次データを読み込んでいます..."):
        df, latest_month, source_url, skipped_sheets = load_housing_data()
except Exception as exc:
    st.error(f"データ取得に失敗しました：{exc}")
    st.markdown(f"[国土交通省の公式時系列一覧を開く]({SOURCE_PAGE})")
    st.stop()

national = df[df["地域"] == "全国"].sort_values("年月")
prefectures = df[df["地域"] != "全国"].copy()
latest_pref = prefectures[prefectures["年月"] == latest_month].copy()
latest_national = national[national["年月"] == latest_month]

if latest_national.empty:
    st.warning("最新月の全国計が見つかりません。")
    latest_total = None
    latest_yoy = None
else:
    latest_total = int(latest_national.iloc[0]["住宅着工戸数"])
    latest_yoy = latest_national.iloc[0]["前年同月比（％）"]

metric1, metric2, metric3, metric4 = st.columns(4)
metric1.metric("最新公表月", latest_month.strftime("%Y年%m月"))
metric2.metric("全国の住宅着工戸数", f"{latest_total:,}戸" if latest_total is not None else "—")
metric3.metric(
    "全国・前年同月比",
    f"{latest_yoy:+.1f}%" if pd.notna(latest_yoy) else "—",
)
metric4.metric("取得期間", f"{national['年月'].min():%Y年%m月}〜{latest_month:%Y年%m月}")

st.subheader("全国の月次推移")
fig_national = px.line(
    national,
    x="年月",
    y="住宅着工戸数",
    markers=True,
    labels={"年月": "年月", "住宅着工戸数": "住宅着工戸数（戸）"},
)
fig_national.update_layout(xaxis_title="年月", yaxis_title="住宅着工戸数（戸）")
st.plotly_chart(fig_national, use_container_width=True)

st.subheader("地域別の月次推移")
available_regions = ["全国"] + sorted(prefectures["地域"].unique().tolist())
default_region_index = available_regions.index("東京都") if "東京都" in available_regions else 0
selected_region = st.selectbox(
    "地域を選択",
    available_regions,
    index=default_region_index,
)
region_df = df[df["地域"] == selected_region].sort_values("年月")
fig_region = px.line(
    region_df,
    x="年月",
    y="住宅着工戸数",
    markers=True,
    labels={"年月": "年月", "住宅着工戸数": "住宅着工戸数（戸）"},
    title=f"{selected_region}の住宅着工戸数",
)
fig_region.update_layout(xaxis_title="年月", yaxis_title="住宅着工戸数（戸）")
st.plotly_chart(fig_region, use_container_width=True)

st.subheader("最新月の都道府県別比較")
latest_pref = latest_pref.sort_values("住宅着工戸数", ascending=False)
fig_pref = px.bar(
    latest_pref,
    x="地域",
    y="住宅着工戸数",
    labels={"地域": "都道府県", "住宅着工戸数": "住宅着工戸数（戸）"},
)
fig_pref.update_layout(xaxis_title="都道府県", yaxis_title="住宅着工戸数（戸）")
st.plotly_chart(fig_pref, use_container_width=True)

st.dataframe(
    latest_pref[["地域", "住宅着工戸数", "前年同月比（％）"]].rename(
        columns={"地域": "都道府県"}
    ).assign(
        **{
            "前年同月比（％）": lambda x: x["前年同月比（％）"].round(1)
        }
    ),
    use_container_width=True,
    hide_index=True,
)

with st.expander("取得データ・更新情報"):
    st.write(f"データ最終月：{latest_month:%Y年%m月}")
    st.write(f"キャッシュ有効期間：{CACHE_SECONDS // 3600}時間")
    st.write(f"取得元Excel：[ダウンロード元を開く]({source_url})")
    st.write(f"対象データ件数：{len(df):,}行（全国＋都道府県）")
    if skipped_sheets:
        st.warning("一部のシートを読み飛ばしました。")
        st.write(skipped_sheets)
    st.dataframe(df, use_container_width=True, hide_index=True)

st.caption(
    "注：最新データは国土交通省の公表状況に依存します。"
    "キャッシュ有効期間は6時間です。"
)
