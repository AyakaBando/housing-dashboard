
import os
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from openai import OpenAI


# =========================================================
# 1. 基本設定
# =========================================================

st.set_page_config(
    page_title="住宅着工統計ダッシュボード",
    page_icon="🏠",
    layout="wide",
)

BASE_URL = "https://api.e-stat.go.jp/rest/3.0/app/json"
CACHE_SECONDS = 60 * 60 * 24
MAX_RECORDS = 200_000

st.title("住宅着工統計ダッシュボード")
st.caption(
    "出典：政府統計の総合窓口（e-Stat）"
    "｜住宅着工統計を利用した公開分析ツール"
)


def get_secret(name, default=""):
    """Streamlit Secretsと環境変数から設定を読む。"""
    try:
        value = st.secrets.get(name, default)
    except Exception:
        value = default

    return os.getenv(name, value) or default


ESTAT_APP_ID = get_secret("ESTAT_APP_ID")
GROQ_API_KEY = get_secret("GROQ_API_KEY")
GROQ_MODEL = get_secret(
    "GROQ_MODEL",
    "openai/gpt-oss-20b",
)


# =========================================================
# 2. APIレスポンスを扱う共通関数
# =========================================================

def as_list(value):
    """APIが1件を辞書、複数件をリストで返す違いを吸収する。"""
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def get_text(value):
    """e-StatのJSONに含まれる文字列を取り出す。"""
    if value is None:
        return ""

    if isinstance(value, dict):
        for key in ("$", "#text", "text"):
            if key in value:
                return str(value[key])

        return ""

    return str(value)


def get_api_json(endpoint, params):
    """e-Stat APIを呼び出す。"""
    if not ESTAT_APP_ID:
        raise ValueError(
            "e-StatのアプリケーションIDが未設定です。"
        )

    request_params = {
        "appId": ESTAT_APP_ID,
        "lang": "J",
        **params,
    }

    response = requests.get(
        f"{BASE_URL}/{endpoint}",
        params=request_params,
        timeout=60,
    )
    response.raise_for_status()

    result = response.json()

    # APIがエラーを返した場合
    root = result.get(
        "GET_STATS_LIST",
        result.get(
            "GET_STATS_DATA",
            result.get("GET_META_INFO", {}),
        ),
    )

    status = root.get("RESULT", {}).get("STATUS", 0)

    if str(status) != "0":
        message = root.get("RESULT", {}).get(
            "ERROR_MSG", "APIエラー"
        )
        raise RuntimeError(str(message))

    return result


def get_root(result):
    """APIレスポンスのルート要素を取得する。"""
    for key in (
        "GET_STATS_LIST",
        "GET_STATS_DATA",
        "GET_META_INFO",
    ):
        if key in result:
            return result[key]

    raise ValueError("APIレスポンスの形式を確認できません。")


# =========================================================
# 3. 住宅着工統計の表を検索
# =========================================================

@st.cache_data(ttl=CACHE_SECONDS)
def search_tables(app_id, search_word):
    params = {
        "appId": app_id,
        "lang": "J",
        "searchWord": search_word,
        "limit": 100,
        "searchKind": 1,
    }

    response = requests.get(
        f"{BASE_URL}/getStatsList",
        params=params,
        timeout=60,
    )
    response.raise_for_status()

    result = response.json()
    root = result.get("GET_STATS_LIST", {})

    status = root.get("RESULT", {}).get("STATUS", 0)
    if str(status) != "0":
        raise RuntimeError(
            root.get("RESULT", {}).get(
                "ERROR_MSG", "統計表検索に失敗しました。"
            )
        )

    data = root.get("DATALIST_INF", {})
    table_items = as_list(data.get("TABLE_INF"))

    tables = []

    for item in table_items:
        table_id = str(item.get("@id", ""))

        title = get_text(item.get("TITLE"))
        stat_name = get_text(item.get("STAT_NAME"))

        if not table_id:
            continue

        tables.append({
            "id": table_id,
            "title": title or f"統計表 {table_id}",
            "stat_name": stat_name,
        })

    return tables


# =========================================================
# 4. 統計表の分類情報を取得
# =========================================================

@st.cache_data(ttl=CACHE_SECONDS)
def get_metadata(app_id, table_id):
    result = get_api_json(
        "getMetaInfo",
        {"statsDataId": table_id},
    )

    root = get_root(result)
    meta = root.get("METADATA_INF", {})
    class_info = meta.get("CLASS_INF", {})
    objects = as_list(class_info.get("CLASS_OBJ"))

    dimensions = {}

    for obj in objects:
        dimension_id = obj.get("@id", "")
        dimension_name = obj.get("@name", dimension_id)

        labels = {}

        for item in as_list(obj.get("CLASS")):
            code = str(item.get("@code", ""))
            name = item.get("@name", code)

            if code:
                labels[code] = str(name)

        if dimension_id:
            dimensions[dimension_id] = {
                "name": str(dimension_name),
                "labels": labels,
            }

    return dimensions


# =========================================================
# 5. 統計データを取得
# =========================================================

@st.cache_data(ttl=CACHE_SECONDS)
def get_statistics(app_id, table_id):
    all_values = []
    dimensions = {}
    start_position = 1

    while len(all_values) < MAX_RECORDS:
        result = get_api_json(
            "getStatsData",
            {
                "statsDataId": table_id,
                "metaGetFlg": "Y",
                "cntGetFlg": "N",
                "limit": 100000,
                "startPosition": start_position,
            },
        )

        root = get_root(result)
        stat_data = root.get("STATISTICAL_DATA", {})

        if not dimensions:
            meta = stat_data.get("CLASS_INF", {})
            objects = as_list(meta.get("CLASS_OBJ"))

            for obj in objects:
                dimension_id = obj.get("@id", "")
                labels = {}

                for item in as_list(obj.get("CLASS")):
                    code = str(item.get("@code", ""))
                    name = item.get("@name", code)

                    if code:
                        labels[code] = str(name)

                if dimension_id:
                    dimensions[dimension_id] = {
                        "name": str(
                            obj.get("@name", dimension_id)
                        ),
                        "labels": labels,
                    }

        data_info = stat_data.get("DATA_INF", {})
        values = as_list(data_info.get("VALUE"))

        if not values:
            break

        all_values.extend(values)

        result_info = stat_data.get("RESULT_INF", {})
        next_key = result_info.get("NEXT_KEY")

        if not next_key or len(all_values) >= MAX_RECORDS:
            break

        next_position = int(next_key)

        if next_position <= start_position:
            break

        start_position = next_position

    rows = []

    for item in all_values[:MAX_RECORDS]:
        row = {}

        # 表章事項、地域、時間、分類事項を取得する
        for code, value in item.items():
            if not code.startswith("@"):
                continue

            dimension_id = code[1:]
            raw_code = str(value)

            dimension = dimensions.get(dimension_id, {})
            labels = dimension.get("labels", {})

            row[dimension_id] = labels.get(
                raw_code, raw_code
            )

            row[f"{dimension_id}_code"] = raw_code

        raw_value = get_text(item.get("$"))

        if not raw_value:
            raw_value = str(item.get("$", ""))

        try:
            row["value"] = float(raw_value.replace(",", ""))
        except (ValueError, AttributeError):
            row["value"] = None

        row["unit"] = str(item.get("@unit", ""))
        rows.append(row)

    df = pd.DataFrame(rows)

    if df.empty:
        return df, dimensions

    if "time" in df.columns:
        df["time_code"] = df.get(
            "time_code", df["time"].astype(str)
        )

    return df, dimensions


# =========================================================
# 6. 画面：接続設定
# =========================================================

with st.sidebar:
    st.header("設定")

    if ESTAT_APP_ID:
        st.success("e-Stat API：設定済み")
    else:
        st.error(
            "e-Stat API IDが未設定です。"
            "Secretsを確認してください。"
        )

    search_word = st.text_input(
        "統計表の検索キーワード",
        value="住宅着工",
    )

    search_button = st.button(
        "統計表を検索",
        type="primary",
    )

    st.divider()

    st.caption(
        "統計表の内容によって、地域や分類の項目は異なります。"
    )


# =========================================================
# 7. 統計表を選択
# =========================================================

if not ESTAT_APP_ID:
    st.info(
        "まずe-StatのアプリケーションIDを設定してください。"
    )
    st.stop()

if search_button:
    search_tables.clear()

try:
    tables = search_tables(ESTAT_APP_ID, search_word)
except Exception as e:
    st.error(f"統計表検索に失敗しました：{e}")
    st.stop()

if not tables:
    st.warning(
        "統計表が見つかりません。"
        "検索語を「住宅着工統計」などに変更してください。"
    )
    st.stop()

table_options = {
    f"{t['title']}（ID: {t['id']}）": t["id"]
    for t in tables
}

selected_table_label = st.selectbox(
    "分析する統計表",
    list(table_options.keys()),
)

selected_table_id = table_options[selected_table_label]

st.caption(
    "選択中の統計表ID：" + selected_table_id
)


# =========================================================
# 8. データ取得・分類情報
# =========================================================

try:
    with st.spinner("e-Statから統計データを取得しています..."):
        df, dimensions = get_statistics(
            ESTAT_APP_ID,
            selected_table_id,
        )

except Exception as e:
    st.error(f"統計データを取得できませんでした：{e}")
    st.stop()

if df.empty:
    st.warning(
        "この統計表から数値データを取得できませんでした。"
        "別の統計表を選択してください。"
    )
    st.stop()

st.success(f"取得件数：{len(df):,}件")

st.caption(
    "取得日時：" + datetime.now().strftime("%Y-%m-%d %H:%M")
)

# e-Statの分類項目から地域・時間の列を判定
area_column = "area" if "area" in df.columns else None
time_column = "time" if "time" in df.columns else None

# 地域の選択
filtered = df.copy()

if area_column:
    areas = sorted(
        filtered[area_column].dropna().astype(str).unique()
    )

    preferred_area = next(
        (a for a in areas if a == "全国"),
        areas[0] if areas else "",
    )

    selected_area = st.selectbox(
        "地域（時系列グラフ）",
        areas,
        index=areas.index(preferred_area)
        if preferred_area in areas else 0,
    )

    filtered = filtered[
        filtered[area_column].astype(str) == selected_area
    ]
else:
    selected_area = None
    st.info("この統計表には地域の分類がありません。")


# =========================================================
# 9. 分類項目を選択
# =========================================================

excluded_dimensions = {
    "area", "time", "value", "unit",
}

dimension_columns = []

for dimension_id, info in dimensions.items():
    if dimension_id in excluded_dimensions:
        continue

    if dimension_id not in df.columns:
        continue

    if df[dimension_id].nunique(dropna=True) <= 1:
        continue

    dimension_columns.append(dimension_id)

with st.expander("構造別・利用関係別などの絞り込み", expanded=True):
    for dimension_id in dimension_columns:
        dimension_name = dimensions.get(
            dimension_id, {}
        ).get("name", dimension_id)

        options = sorted(
            df[dimension_id].dropna().astype(str).unique()
        )

        if not options:
            continue

        # 各分類は1つを選び、異なる分類の数値を
        # 意図せず合算しないようにする
        default_index = 0

        # 総数・合計がある場合は初期値にする
        total_index = next(
            (
                i for i, option in enumerate(options)
                if "総数" in option or "合計" in option
            ),
            None,
        )

        if total_index is not None:
            default_index = total_index

        selected_category = st.selectbox(
            dimension_name,
            options,
            index=default_index,
            key=f"filter_{dimension_id}",
        )

        filtered = filtered[
            filtered[dimension_id].astype(str)
            == selected_category
        ]


# =========================================================
# 10. 数値と年月の整形
# =========================================================

if "value" not in filtered.columns:
    st.error("数値列を取得できませんでした。")
    st.stop()

filtered = filtered.dropna(subset=["value"]).copy()

if filtered.empty:
    st.warning(
        "選択した条件に該当する数値がありません。"
    )
    st.stop()

# 時間軸のラベルを表示
if time_column:
    filtered["time_label"] = filtered[time_column].astype(str)

    # 時間軸のコードがあれば、並べ替えに使用する
    if "time_code" in filtered.columns:
        filtered["_time_sort"] = filtered["time_code"].astype(str)
    else:
        filtered["_time_sort"] = filtered["time_label"]

    filtered = filtered.sort_values("_time_sort")

    # 同一月に複数行ある場合は合算
    chart_df = (
        filtered.groupby(
            ["_time_sort", "time_label"],
            as_index=False,
        )["value"]
        .sum()
        .sort_values("_time_sort")
    )

    st.subheader("住宅着工の月次推移")

    fig = px.line(
        chart_df,
        x="time_label",
        y="value",
        markers=True,
        labels={
            "time_label": "年月",
            "value": "統計値",
        },
        title=f"{selected_area or '地域未分類'}の推移",
    )

    fig.update_layout(
        xaxis_tickangle=-45,
        hovermode="x unified",
    )

    st.plotly_chart(fig, use_container_width=True)

    if len(chart_df) >= 2:
        first_value = chart_df.iloc[0]["value"]
        latest_value = chart_df.iloc[-1]["value"]

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "最新の統計値",
            f"{latest_value:,.0f}",
        )

        col2.metric(
            "最初の統計値",
            f"{first_value:,.0f}",
        )

        if first_value != 0:
            change = (
                (latest_value - first_value)
                / abs(first_value)
                * 100
            )
            col3.metric(
                "期間内の変化率",
                f"{change:+.1f}%",
            )
        else:
            col3.metric("期間内の変化率", "算出不可")
else:
    st.info(
        "時間軸がない統計表のため、時系列グラフは表示しません。"
    )


# =========================================================
# 11. 都道府県別比較
# =========================================================

if area_column and time_column:
    st.subheader("地域別の比較")

    comparison_df = df.dropna(
        subset=["value"]
    ).copy()

    # 選択済みの分類条件を適用する
    for dimension_id in dimension_columns:
        selected_category = st.session_state.get(
            f"filter_{dimension_id}"
        )

        if selected_category is not None:
            comparison_df = comparison_df[
                comparison_df[dimension_id].astype(str)
                == str(selected_category)
            ]

    # 最新時点を選ぶ
    if "time_code" in comparison_df.columns:
        latest_code = comparison_df["time_code"].astype(str).max()

        comparison_df = comparison_df[
            comparison_df["time_code"].astype(str) == latest_code
        ]

    else:
        latest_label = comparison_df[time_column].astype(str).max()

        comparison_df = comparison_df[
            comparison_df[time_column].astype(str) == latest_label
        ]

    # 全国は都道府県比較から除外
    comparison_df = comparison_df[
        comparison_df[area_column].astype(str) != "全国"
    ]

    # 同じ都道府県に複数行ある場合の集計
    region_df = (
        comparison_df.groupby(
            area_column,
            as_index=False,
        )["value"]
        .sum()
        .sort_values("value", ascending=False)
    )

    if not region_df.empty:
        fig_region = px.bar(
            region_df.head(47),
            x=area_column,
            y="value",
            labels={
                area_column: "地域",
                "value": "統計値",
            },
            title="都道府県別の統計値",
        )

        fig_region.update_layout(
            xaxis_tickangle=-45,
        )

        st.plotly_chart(
            fig_region,
            use_container_width=True,
        )
    else:
        st.info(
            "都道府県別の比較対象がありません。"
            "地域分類のある統計表を選んでください。"
        )


# =========================================================
# 12. データテーブル・CSVダウンロード
# =========================================================

st.subheader("取得データ")

st.dataframe(
    filtered,
    use_container_width=True,
    hide_index=True,
)

csv_data = filtered.to_csv(
    index=False
).encode("utf-8-sig")

st.download_button(
    "表示中のデータをCSVでダウンロード",
    data=csv_data,
    file_name="housing_starts_analysis.csv",
    mime="text/csv",
)


# =========================================================
# 13. AIによる分析
# =========================================================

st.subheader("AIによる分析")

if not GROQ_API_KEY:
    st.info(
        "Groq APIキーが未設定のため、AI分析は無効です。"
        "グラフと統計データは引き続き利用できます。"
    )

else:
    if st.button("現在のデータをAIで分析する"):
        # AIに大量の生データを送らず、集計結果を送る
        ai_df = filtered.copy()

        if time_column and time_column in ai_df.columns:
            summary = (
                ai_df.groupby(time_column)["value"]
                .sum()
                .tail(36)
                .reset_index()
            )
        else:
            summary = ai_df.head(100)

        summary_text = summary.to_csv(index=False)

        prompt = f"""
あなたは日本の住宅着工統計を分析するアシスタントです。

以下のデータを日本語で分析してください。

条件：
- データに存在しない数値を作らない
- 数値から直接確認できる事実と推測を分ける
- 景気、金利、人口動態などの原因は、
  データだけでは断定できないと明記する
- 時系列の変化、注目点、追加調査の案を説明する
- 数値の単位が不明な場合は、単位を断定しない
- 住宅着工「棟数」ではなく「戸数」の表である場合も
  あるため、表の単位を確認する

統計表：
{selected_table_label}

対象地域：
{selected_area or '地域分類なし'}

データ：
{summary_text}
"""

        try:
            client = OpenAI(
                api_key=GROQ_API_KEY,
                base_url="https://api.groq.com/openai/v1",
                timeout=120.0,
                max_retries=1,
            )

            with st.spinner("AIが分析しています..."):
                response = client.chat.completions.create(
                    model=GROQ_MODEL,
                    messages=[
                        {
                            "role": "system",
                            "content": (
                                "あなたは正確性を重視する"
                                "統計分析アシスタントです。"
                            ),
                        },
                        {
                            "role": "user",
                            "content": prompt,
                        },
                    ],
                    temperature=0.2,
                )

            answer = response.choices[0].message.content

            if answer:
                st.markdown(answer)
            else:
                st.warning("AIから回答が返されませんでした。")

        except Exception as e:
            st.error(
                f"AI分析に失敗しました："
                f"{type(e).__name__}: {e}"
            )


# =========================================================
# 14. 出典・更新に関する表示
# =========================================================

st.divider()

st.markdown(
    """
**データ出典：** [政府統計の総合窓口（e-Stat）](https://www.e-stat.go.jp/)

このページはe-Stat APIを利用しています。
統計表の定義、分類、単位、公開時点を確認したうえでご利用ください。

AIによるコメントは参考情報です。
統計値そのものは、表示している元データを確認してください。
"""
)
