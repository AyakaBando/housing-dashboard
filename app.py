    df["_time_label"] = "時点なし"

# Sidebar filters
with st.sidebar:
    st.header("表示設定")
    if area_dim:
        area_codes = sorted(df[area_dim].dropna().astype(str).unique())
        area_names = {code: label_for(area_dim, code, dimensions) for code in area_codes}
        national_code = next((code for code, name in area_names.items() if name in ("全国", "全国計")), None)
        default_idx = area_codes.index(national_code) if national_code in area_codes else 0
        selected_area = st.selectbox(
            "時系列グラフの地域",
            area_codes,
            index=default_idx,
            format_func=lambda code: area_names.get(code, code),
        )
    else:
        selected_area = None
        st.caption("この表に地域分類はありません。")

    selected_dims = {}
    for dim_id in dimension_ids:
        info = dimensions.get(dim_id, {})
        codes = sorted(df[dim_id].dropna().astype(str).unique())
        if len(codes) > 1:
            selected_dims[dim_id] = st.selectbox(
                info.get("name", dim_id),
                codes,
                format_func=lambda code, d=dim_id: label_for(d, code, dimensions),
            )

# Keep a selected series for time chart.
trend = df.copy()
if area_dim and selected_area is not None:
    trend = trend[trend[area_dim].astype(str) == str(selected_area)]
for dim_id, code in selected_dims.items():
    trend = trend[trend[dim_id].astype(str) == str(code)]

# A stats table may have several records per month due to extra classifications.
# User selections above reduce them where possible; sum the remaining values by time.
trend = (
    trend.groupby(["_time_sort", "_time_label"], as_index=False)["value"]
    .sum()
    .sort_values("_time_sort")
)

latest_label = str(trend.iloc[-1]["_time_label"]) if not trend.empty else "不明"
metric1, metric2, metric3 = st.columns(3)
metric1.metric("最新時点", latest_label)
metric2.metric("最新値", f"{trend.iloc[-1]['value']:,.0f}" if not trend.empty else "—")
metric3.metric("データ件数", f"{len(df):,}")

st.subheader("全国・選択地域の月次推移")
if trend.empty:
    st.warning("選択条件に該当するデータがありません。")
else:
    fig = px.line(trend, x="_time_label", y="value", markers=True, labels={"_time_label": "年月", "value": "値"})
    fig.update_layout(xaxis_title="年月", yaxis_title=f"値（{unit_values[0] if unit_values else ''}）", xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)

# Prefecture comparison: use latest available time and aggregate other dimensions only
st.subheader("都道府県別比較")
if area_dim:
    latest_sort = df["_time_sort"].max()
    latest_df = df[df["_time_sort"] == latest_sort].copy()
    latest_df["_area_label"] = latest_df[area_dim].map(lambda x: label_for(area_dim, x, dimensions))

    # Only display prefectural-level rows, excluding national and regional totals.
    exclude_names = {"全国", "全国計", "全国（計）", "北海道・東北", "関東", "北陸", "中部", "近畿", "中国", "四国", "九州", "沖縄"}
    latest_df = latest_df[~latest_df["_area_label"].isin(exclude_names)]
    prefecture_names = {
        "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
        "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
        "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
        "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
        "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
        "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
        "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"
    }
    prefecture_df = latest_df[latest_df["_area_label"].isin(prefecture_names)]
    if prefecture_df.empty:
        st.info("この統計表の最新時点に都道府県別データが見つかりません。統計表の地域分類を確認してください。")
    else:
        prefecture_df = prefecture_df.groupby("_area_label", as_index=False)["value"].sum().sort_values("value", ascending=False)
        st.caption(f"比較時点：{latest_label}")
        fig2 = px.bar(prefecture_df, x="_area_label", y="value", labels={"_area_label": "都道府県", "value": "値"})
        fig2.update_layout(xaxis_title="都道府県", yaxis_title=f"値（{unit_values[0] if unit_values else ''}）", xaxis_tickangle=-45)
        st.plotly_chart(fig2, use_container_width=True)
        st.dataframe(prefecture_df.rename(columns={"_area_label": "都道府県", "value": "値"}), use_container_width=True, hide_index=True)
else:
    st.info("この統計表に地域分類がないため、都道府県別比較は表示できません。")

with st.expander("取得データの確認"):
    st.write("分類項目")
    st.json({k: v["name"] for k, v in dimensions.items()})
    st.dataframe(df.head(100), use_container_width=True)

st.subheader("AIによる分析")
if not GROQ_API_KEY:
    st.info("AI分析を使う場合は、Streamlit Secretsに GROQ_API_KEY を設定してください。グラフの表示には不要です。")
else:
    if st.button("最新データをAIで分析"):
        try:
            client = Groq(api_key=GROQ_API_KEY)
            recent = trend.tail(24).to_dict(orient="records")
            prompt = (
                "あなたは住宅市場データの分析担当です。以下の住宅着工統計について、"
                "観察できる傾向を日本語で簡潔に説明してください。"
                "データから断定できない原因は推測と明記し、数値を捏造しないでください。\n\n"
                f"統計表ID: {STATS_DATA_ID}\n"
                f"地域: {selected_area if selected_area is not None else '地域分類なし'}\n"
                f"直近データ: {recent}"
            )
            completion = client.chat.completions.create(
                model=GROQ_MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2,
            )
            st.markdown(completion.choices[0].message.content)
        except Exception as exc:
            st.error(f"AI分析に失敗しました：{exc}")

st.caption(f"最終取得：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
st.caption("注意：統計表の分類や単位により、表示できる分析は異なります。")
