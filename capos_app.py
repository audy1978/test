from pathlib import Path

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).parent
DEFAULT_FILE = APP_DIR / "à¸ªà¸£à¸¸à¸›à¸›à¸£à¸´à¸¡à¸²à¸“à¸‡à¸²à¸™ à¸£à¸²à¸¢à¹„à¸”à¹‰ à¸ˆà¸²à¸à¸£à¸°à¸šà¸š New CAPOS à¸£à¸²à¸¢à¹€à¸”à¸·à¸­à¸™ à¸›.xlsx"

MONTHS = [
    "à¸¡.à¸„.",
    "à¸.à¸ž.",
    "à¸¡à¸µ.à¸„.",
    "à¹€à¸¡.à¸¢.",
    "à¸ž.à¸„.",
    "à¸¡à¸´.à¸¢.",
    "à¸.à¸„.",
    "à¸ª.à¸„.",
    "à¸.à¸¢.",
    "à¸•.à¸„.",
    "à¸ž.à¸¢.",
    "à¸˜.à¸„.",
]


st.set_page_config(
    page_title="New CAPOS Monthly Dashboard",
    page_icon="ðŸ“Š",
    layout="wide",
)


@st.cache_data(show_spinner=False)
def load_capos_workbook(file):
    executive = pd.read_excel(file, sheet_name="Executive", header=3)
    detail = pd.read_excel(file, sheet_name="Grouped Detail", header=3)
    monthly = pd.read_excel(file, sheet_name="Monthly Summary", header=3)

    executive = clean_summary_table(executive)
    detail = clean_detail_table(detail)
    monthly = clean_summary_table(monthly)
    long_detail = make_long_detail(detail)
    long_monthly = make_long_monthly(monthly)

    return executive, detail, monthly, long_detail, long_monthly


def clean_summary_table(df):
    df = df.dropna(how="all").copy()
    df = df[df["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"].notna()]
    df["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"] = df["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"].astype(str)

    for col in ["à¸ˆà¸³à¸™à¸§à¸™ à¸›à¸“", "à¸£à¸²à¸¢à¸à¸²à¸£à¸šà¸£à¸´à¸à¸²à¸£", "à¸Šà¸´à¹‰à¸™à¸£à¸§à¸¡ 2568", "à¹€à¸‡à¸´à¸™à¸£à¸§à¸¡ 2568", "à¸£à¸§à¸¡ à¸Šà¸´à¹‰à¸™", "à¸£à¸§à¸¡ à¹€à¸‡à¸´à¸™"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def clean_detail_table(df):
    df = df.dropna(how="all").copy()
    df = df[df["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"].notna()]
    df = df[df["à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£"].notna()]

    text_cols = [
        "à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ",
        "à¹€à¸«à¸•à¸¸à¸œà¸¥à¹€à¸Šà¸´à¸‡à¸¢à¸¸à¸—à¸˜à¸¨à¸²à¸ªà¸•à¸£à¹Œ",
        "à¹€à¸‚à¸•",
        "à¸£à¸«à¸±à¸ªà¹„à¸›à¸£à¸©à¸“à¸µà¸¢à¹Œ",
        "à¸›à¸“",
        "à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£",
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r"\.0$", "", regex=True)

    numeric_cols = [
        col
        for col in df.columns
        if col.endswith("à¸Šà¸´à¹‰à¸™") or col.endswith("à¹€à¸‡à¸´à¸™") or col in ["à¹à¸–à¸§à¸•à¹‰à¸™à¸‰à¸šà¸±à¸š", "à¸¥à¸³à¸”à¸±à¸š"]
    ]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def make_long_detail(detail):
    long_rows = []
    id_cols = [
        "à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ",
        "à¹€à¸«à¸•à¸¸à¸œà¸¥à¹€à¸Šà¸´à¸‡à¸¢à¸¸à¸—à¸˜à¸¨à¸²à¸ªà¸•à¸£à¹Œ",
        "à¹€à¸‚à¸•",
        "à¸£à¸«à¸±à¸ªà¹„à¸›à¸£à¸©à¸“à¸µà¸¢à¹Œ",
        "à¸›à¸“",
        "à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£",
    ]

    for index, month in enumerate(MONTHS, start=1):
        pieces_col = f"{month} à¸Šà¸´à¹‰à¸™"
        revenue_col = f"{month} à¹€à¸‡à¸´à¸™"
        month_df = detail[id_cols + [pieces_col, revenue_col]].copy()
        month_df = month_df.rename(columns={pieces_col: "à¸Šà¸´à¹‰à¸™", revenue_col: "à¹€à¸‡à¸´à¸™"})
        month_df["à¹€à¸”à¸·à¸­à¸™"] = month
        month_df["à¸¥à¸³à¸”à¸±à¸šà¹€à¸”à¸·à¸­à¸™"] = index
        long_rows.append(month_df)

    long_df = pd.concat(long_rows, ignore_index=True)
    long_df["à¸Šà¸´à¹‰à¸™"] = pd.to_numeric(long_df["à¸Šà¸´à¹‰à¸™"], errors="coerce").fillna(0)
    long_df["à¹€à¸‡à¸´à¸™"] = pd.to_numeric(long_df["à¹€à¸‡à¸´à¸™"], errors="coerce").fillna(0)
    long_df = long_df[(long_df["à¸Šà¸´à¹‰à¸™"] != 0) | (long_df["à¹€à¸‡à¸´à¸™"] != 0)]
    long_df["à¸£à¸²à¸¢à¹„à¸”à¹‰à¸•à¹ˆà¸­à¸Šà¸´à¹‰à¸™"] = long_df["à¹€à¸‡à¸´à¸™"].div(long_df["à¸Šà¸´à¹‰à¸™"]).where(long_df["à¸Šà¸´à¹‰à¸™"] > 0, 0)
    return long_df


def make_long_monthly(monthly):
    long_rows = []
    id_cols = ["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ", "à¸ˆà¸³à¸™à¸§à¸™ à¸›à¸“", "à¸£à¸²à¸¢à¸à¸²à¸£à¸šà¸£à¸´à¸à¸²à¸£"]

    for index, month in enumerate(MONTHS, start=1):
        pieces_col = f"{month} à¸Šà¸´à¹‰à¸™"
        revenue_col = f"{month} à¹€à¸‡à¸´à¸™"
        month_df = monthly[id_cols + [pieces_col, revenue_col]].copy()
        month_df = month_df.rename(columns={pieces_col: "à¸Šà¸´à¹‰à¸™", revenue_col: "à¹€à¸‡à¸´à¸™"})
        month_df["à¹€à¸”à¸·à¸­à¸™"] = month
        month_df["à¸¥à¸³à¸”à¸±à¸šà¹€à¸”à¸·à¸­à¸™"] = index
        long_rows.append(month_df)

    long_df = pd.concat(long_rows, ignore_index=True)
    long_df["à¸Šà¸´à¹‰à¸™"] = pd.to_numeric(long_df["à¸Šà¸´à¹‰à¸™"], errors="coerce").fillna(0)
    long_df["à¹€à¸‡à¸´à¸™"] = pd.to_numeric(long_df["à¹€à¸‡à¸´à¸™"], errors="coerce").fillna(0)
    long_df["à¸£à¸²à¸¢à¹„à¸”à¹‰à¸•à¹ˆà¸­à¸Šà¸´à¹‰à¸™"] = long_df["à¹€à¸‡à¸´à¸™"].div(long_df["à¸Šà¸´à¹‰à¸™"]).where(long_df["à¸Šà¸´à¹‰à¸™"] > 0, 0)
    return long_df


def number_format(value):
    return f"{value:,.0f}"


def money_format(value):
    return f"{value:,.2f}"


st.title("à¸ªà¸£à¸¸à¸›à¸›à¸£à¸´à¸¡à¸²à¸“à¸‡à¸²à¸™ à¸£à¸²à¸¢à¹„à¸”à¹‰ à¸ˆà¸²à¸à¸£à¸°à¸šà¸š New CAPOS à¸£à¸²à¸¢à¹€à¸”à¸·à¸­à¸™ à¸›")

with st.sidebar:
    st.header("à¸‚à¹‰à¸­à¸¡à¸¹à¸¥")
    uploaded_file = st.file_uploader("à¹€à¸¥à¸·à¸­à¸à¹„à¸Ÿà¸¥à¹Œ Excel", type=["xlsx"])
    source_file = uploaded_file if uploaded_file is not None else DEFAULT_FILE
    st.caption("à¹„à¸Ÿà¸¥à¹Œà¹€à¸£à¸´à¹ˆà¸¡à¸•à¹‰à¸™: " + (uploaded_file.name if uploaded_file else DEFAULT_FILE.name))

try:
    executive_df, detail_df, monthly_df, long_detail_df, long_monthly_df = load_capos_workbook(source_file)
except FileNotFoundError:
    st.error(f"à¹„à¸¡à¹ˆà¸žà¸šà¹„à¸Ÿà¸¥à¹Œà¹€à¸£à¸´à¹ˆà¸¡à¸•à¹‰à¸™: {DEFAULT_FILE}")
    st.stop()
except Exception as exc:
    st.error(f"à¸­à¹ˆà¸²à¸™à¹„à¸Ÿà¸¥à¹Œà¹„à¸¡à¹ˆà¸ªà¸³à¹€à¸£à¹‡à¸ˆ: {exc}")
    st.stop()

with st.sidebar:
    st.header("à¸•à¸±à¸§à¸à¸£à¸­à¸‡")
    area_options = sorted(long_detail_df["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"].dropna().unique())
    selected_areas = st.multiselect("à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ", area_options, default=area_options)

    area_filtered = long_detail_df[long_detail_df["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"].isin(selected_areas)]

    district_options = sorted(area_filtered["à¹€à¸‚à¸•"].dropna().unique())
    selected_districts = st.multiselect("à¹€à¸‚à¸•", district_options, default=district_options)

    district_filtered = area_filtered[area_filtered["à¹€à¸‚à¸•"].isin(selected_districts)]

    office_options = sorted(district_filtered["à¸›à¸“"].dropna().unique())
    selected_offices = st.multiselect("à¸›à¸“", office_options, default=office_options)

    office_filtered = district_filtered[district_filtered["à¸›à¸“"].isin(selected_offices)]

    service_options = sorted(office_filtered["à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£"].dropna().unique())
    selected_services = st.multiselect("à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£", service_options, default=service_options)

    selected_months = st.multiselect("à¹€à¸”à¸·à¸­à¸™", MONTHS, default=MONTHS)

filtered = long_detail_df[
    long_detail_df["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"].isin(selected_areas)
    & long_detail_df["à¹€à¸‚à¸•"].isin(selected_districts)
    & long_detail_df["à¸›à¸“"].isin(selected_offices)
    & long_detail_df["à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£"].isin(selected_services)
    & long_detail_df["à¹€à¸”à¸·à¸­à¸™"].isin(selected_months)
].copy()

filtered_monthly = long_monthly_df[
    long_monthly_df["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"].isin(selected_areas)
    & long_monthly_df["à¹€à¸”à¸·à¸­à¸™"].isin(selected_months)
].copy()

if filtered.empty:
    st.warning("à¹„à¸¡à¹ˆà¸¡à¸µà¸‚à¹‰à¸­à¸¡à¸¹à¸¥à¸•à¸²à¸¡à¸•à¸±à¸§à¸à¸£à¸­à¸‡à¸—à¸µà¹ˆà¹€à¸¥à¸·à¸­à¸")
    st.stop()

total_pieces = filtered["à¸Šà¸´à¹‰à¸™"].sum()
total_revenue = filtered["à¹€à¸‡à¸´à¸™"].sum()
avg_revenue_per_piece = total_revenue / total_pieces if total_pieces else 0

metric_cols = st.columns(5)
metric_cols[0].metric("à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ", number_format(filtered["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"].nunique()))
metric_cols[1].metric("à¸ˆà¸³à¸™à¸§à¸™ à¸›à¸“", number_format(filtered["à¸›à¸“"].nunique()))
metric_cols[2].metric("à¸ˆà¸³à¸™à¸§à¸™à¸Šà¸´à¹‰à¸™", number_format(total_pieces))
metric_cols[3].metric("à¸£à¸²à¸¢à¹„à¸”à¹‰", money_format(total_revenue))
metric_cols[4].metric("à¸£à¸²à¸¢à¹„à¸”à¹‰à¸•à¹ˆà¸­à¸Šà¸´à¹‰à¸™", money_format(avg_revenue_per_piece))

tab_exec, tab_month, tab_service, tab_office, tab_data = st.tabs(
    ["Executive", "à¸£à¸²à¸¢à¹€à¸”à¸·à¸­à¸™", "à¸šà¸£à¸´à¸à¸²à¸£", "à¸›à¸“/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ", "à¸‚à¹‰à¸­à¸¡à¸¹à¸¥à¸¥à¸°à¹€à¸­à¸µà¸¢à¸”"]
)

with tab_exec:
    executive_view = executive_df[executive_df["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"].isin(selected_areas)].copy()
    executive_view = executive_view[
        [
            "à¸¥à¸³à¸”à¸±à¸š",
            "à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ",
            "à¹€à¸«à¸•à¸¸à¸œà¸¥à¹€à¸Šà¸´à¸‡à¸¢à¸¸à¸—à¸˜à¸¨à¸²à¸ªà¸•à¸£à¹Œ",
            "à¸ˆà¸³à¸™à¸§à¸™ à¸›à¸“",
            "à¸£à¸²à¸¢à¸à¸²à¸£à¸šà¸£à¸´à¸à¸²à¸£",
            "à¸Šà¸´à¹‰à¸™à¸£à¸§à¸¡ 2568",
            "à¹€à¸‡à¸´à¸™à¸£à¸§à¸¡ 2568",
            "à¸ªà¸–à¸²à¸™à¸°à¸•à¸£à¸§à¸ˆà¸ªà¸­à¸š",
        ]
    ]
    st.dataframe(
        executive_view,
        width="stretch",
        hide_index=True,
        column_config={
            "à¸ˆà¸³à¸™à¸§à¸™ à¸›à¸“": st.column_config.NumberColumn("à¸ˆà¸³à¸™à¸§à¸™ à¸›à¸“", format="%,.0f"),
            "à¸£à¸²à¸¢à¸à¸²à¸£à¸šà¸£à¸´à¸à¸²à¸£": st.column_config.NumberColumn("à¸£à¸²à¸¢à¸à¸²à¸£à¸šà¸£à¸´à¸à¸²à¸£", format="%,.0f"),
            "à¸Šà¸´à¹‰à¸™à¸£à¸§à¸¡ 2568": st.column_config.NumberColumn("à¸Šà¸´à¹‰à¸™à¸£à¸§à¸¡ 2568", format="%,.0f"),
            "à¹€à¸‡à¸´à¸™à¸£à¸§à¸¡ 2568": st.column_config.NumberColumn("à¹€à¸‡à¸´à¸™à¸£à¸§à¸¡ 2568", format="%,.2f"),
        },
    )

with tab_month:
    monthly_by_area = (
        filtered_monthly.groupby(["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ"], as_index=False)[["à¸Šà¸´à¹‰à¸™", "à¹€à¸‡à¸´à¸™"]]
        .sum()
        .sort_values("à¹€à¸‡à¸´à¸™", ascending=False)
    )
    monthly_trend = (
        filtered_monthly.groupby(["à¸¥à¸³à¸”à¸±à¸šà¹€à¸”à¸·à¸­à¸™", "à¹€à¸”à¸·à¸­à¸™"], as_index=False)[["à¸Šà¸´à¹‰à¸™", "à¹€à¸‡à¸´à¸™"]]
        .sum()
        .sort_values("à¸¥à¸³à¸”à¸±à¸šà¹€à¸”à¸·à¸­à¸™")
    )

    chart_cols = st.columns(2)
    with chart_cols[0]:
        st.subheader("à¹à¸™à¸§à¹‚à¸™à¹‰à¸¡à¸ˆà¸³à¸™à¸§à¸™à¸Šà¸´à¹‰à¸™")
        st.line_chart(monthly_trend, x="à¹€à¸”à¸·à¸­à¸™", y="à¸Šà¸´à¹‰à¸™")
    with chart_cols[1]:
        st.subheader("à¹à¸™à¸§à¹‚à¸™à¹‰à¸¡à¸£à¸²à¸¢à¹„à¸”à¹‰")
        st.line_chart(monthly_trend, x="à¹€à¸”à¸·à¸­à¸™", y="à¹€à¸‡à¸´à¸™")

    st.subheader("à¸£à¸²à¸¢à¹„à¸”à¹‰à¸•à¸²à¸¡à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ")
    st.bar_chart(monthly_by_area, x="à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ", y="à¹€à¸‡à¸´à¸™")

with tab_service:
    service_summary = (
        filtered.groupby("à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£", as_index=False)[["à¸Šà¸´à¹‰à¸™", "à¹€à¸‡à¸´à¸™"]]
        .sum()
        .sort_values("à¹€à¸‡à¸´à¸™", ascending=False)
    )
    service_summary["à¸£à¸²à¸¢à¹„à¸”à¹‰à¸•à¹ˆà¸­à¸Šà¸´à¹‰à¸™"] = service_summary["à¹€à¸‡à¸´à¸™"].div(
        service_summary["à¸Šà¸´à¹‰à¸™"]
    ).where(service_summary["à¸Šà¸´à¹‰à¸™"] > 0, 0)

    st.subheader("Top 20 à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£à¸•à¸²à¸¡à¸£à¸²à¸¢à¹„à¸”à¹‰")
    st.bar_chart(service_summary.head(20), x="à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£", y="à¹€à¸‡à¸´à¸™")
    st.dataframe(
        service_summary,
        width="stretch",
        hide_index=True,
        column_config={
            "à¸Šà¸´à¹‰à¸™": st.column_config.NumberColumn("à¸Šà¸´à¹‰à¸™", format="%,.0f"),
            "à¹€à¸‡à¸´à¸™": st.column_config.NumberColumn("à¹€à¸‡à¸´à¸™", format="%,.2f"),
            "à¸£à¸²à¸¢à¹„à¸”à¹‰à¸•à¹ˆà¸­à¸Šà¸´à¹‰à¸™": st.column_config.NumberColumn("à¸£à¸²à¸¢à¹„à¸”à¹‰à¸•à¹ˆà¸­à¸Šà¸´à¹‰à¸™", format="%,.2f"),
        },
    )

with tab_office:
    office_summary = (
        filtered.groupby(["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ", "à¹€à¸‚à¸•", "à¸£à¸«à¸±à¸ªà¹„à¸›à¸£à¸©à¸“à¸µà¸¢à¹Œ", "à¸›à¸“"], as_index=False)[["à¸Šà¸´à¹‰à¸™", "à¹€à¸‡à¸´à¸™"]]
        .sum()
        .sort_values("à¹€à¸‡à¸´à¸™", ascending=False)
    )
    area_summary = (
        filtered.groupby("à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ", as_index=False)[["à¸Šà¸´à¹‰à¸™", "à¹€à¸‡à¸´à¸™"]]
        .sum()
        .sort_values("à¹€à¸‡à¸´à¸™", ascending=False)
    )

    st.subheader("à¸ªà¸£à¸¸à¸›à¸•à¸²à¸¡à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ")
    st.dataframe(
        area_summary,
        width="stretch",
        hide_index=True,
        column_config={
            "à¸Šà¸´à¹‰à¸™": st.column_config.NumberColumn("à¸Šà¸´à¹‰à¸™", format="%,.0f"),
            "à¹€à¸‡à¸´à¸™": st.column_config.NumberColumn("à¹€à¸‡à¸´à¸™", format="%,.2f"),
        },
    )

    st.subheader("Top 30 à¸›à¸“ à¸•à¸²à¸¡à¸£à¸²à¸¢à¹„à¸”à¹‰")
    st.dataframe(
        office_summary.head(30),
        width="stretch",
        hide_index=True,
        column_config={
            "à¸Šà¸´à¹‰à¸™": st.column_config.NumberColumn("à¸Šà¸´à¹‰à¸™", format="%,.0f"),
            "à¹€à¸‡à¸´à¸™": st.column_config.NumberColumn("à¹€à¸‡à¸´à¸™", format="%,.2f"),
        },
    )

with tab_data:
    display_df = filtered.sort_values(
        ["à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ", "à¹€à¸‚à¸•", "à¸›à¸“", "à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£", "à¸¥à¸³à¸”à¸±à¸šà¹€à¸”à¸·à¸­à¸™"]
    )[
        [
            "à¸ªà¸²à¸‚à¸²/à¸žà¸·à¹‰à¸™à¸—à¸µà¹ˆ",
            "à¹€à¸‚à¸•",
            "à¸£à¸«à¸±à¸ªà¹„à¸›à¸£à¸©à¸“à¸µà¸¢à¹Œ",
            "à¸›à¸“",
            "à¸à¸¥à¸¸à¹ˆà¸¡à¸šà¸£à¸´à¸à¸²à¸£",
            "à¹€à¸”à¸·à¸­à¸™",
            "à¸Šà¸´à¹‰à¸™",
            "à¹€à¸‡à¸´à¸™",
            "à¸£à¸²à¸¢à¹„à¸”à¹‰à¸•à¹ˆà¸­à¸Šà¸´à¹‰à¸™",
        ]
    ]

    st.dataframe(
        display_df,
        width="stretch",
        hide_index=True,
        column_config={
            "à¸Šà¸´à¹‰à¸™": st.column_config.NumberColumn("à¸Šà¸´à¹‰à¸™", format="%,.0f"),
            "à¹€à¸‡à¸´à¸™": st.column_config.NumberColumn("à¹€à¸‡à¸´à¸™", format="%,.2f"),
            "à¸£à¸²à¸¢à¹„à¸”à¹‰à¸•à¹ˆà¸­à¸Šà¸´à¹‰à¸™": st.column_config.NumberColumn("à¸£à¸²à¸¢à¹„à¸”à¹‰à¸•à¹ˆà¸­à¸Šà¸´à¹‰à¸™", format="%,.2f"),
        },
    )

    csv = display_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "à¸”à¸²à¸§à¸™à¹Œà¹‚à¸«à¸¥à¸” CSV",
        data=csv,
        file_name="new_capos_filtered.csv",
        mime="text/csv",
    )

