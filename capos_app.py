from pathlib import Path

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).parent
DEFAULT_FILE = APP_DIR / "สรุปปริมาณงาน รายได้ จากระบบ New CAPOS รายเดือน ป.xlsx"

MONTHS = [
    "ม.ค.",
    "ก.พ.",
    "มี.ค.",
    "เม.ย.",
    "พ.ค.",
    "มิ.ย.",
    "ก.ค.",
    "ส.ค.",
    "ก.ย.",
    "ต.ค.",
    "พ.ย.",
    "ธ.ค.",
]


st.set_page_config(
    page_title="New CAPOS Monthly Dashboard",
    page_icon="📊",
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
    df = df[df["สาขา/พื้นที่"].notna()]
    df["สาขา/พื้นที่"] = df["สาขา/พื้นที่"].astype(str)

    for col in ["จำนวน ปณ", "รายการบริการ", "ชิ้นรวม 2568", "เงินรวม 2568", "รวม ชิ้น", "รวม เงิน"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def clean_detail_table(df):
    df = df.dropna(how="all").copy()
    df = df[df["สาขา/พื้นที่"].notna()]
    df = df[df["กลุ่มบริการ"].notna()]

    text_cols = [
        "สาขา/พื้นที่",
        "เหตุผลเชิงยุทธศาสตร์",
        "เขต",
        "รหัสไปรษณีย์",
        "ปณ",
        "กลุ่มบริการ",
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r"\.0$", "", regex=True)

    numeric_cols = [
        col
        for col in df.columns
        if col.endswith("ชิ้น") or col.endswith("เงิน") or col in ["แถวต้นฉบับ", "ลำดับ"]
    ]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def make_long_detail(detail):
    long_rows = []
    id_cols = [
        "สาขา/พื้นที่",
        "เหตุผลเชิงยุทธศาสตร์",
        "เขต",
        "รหัสไปรษณีย์",
        "ปณ",
        "กลุ่มบริการ",
    ]

    for index, month in enumerate(MONTHS, start=1):
        pieces_col = f"{month} ชิ้น"
        revenue_col = f"{month} เงิน"
        month_df = detail[id_cols + [pieces_col, revenue_col]].copy()
        month_df = month_df.rename(columns={pieces_col: "ชิ้น", revenue_col: "เงิน"})
        month_df["เดือน"] = month
        month_df["ลำดับเดือน"] = index
        long_rows.append(month_df)

    long_df = pd.concat(long_rows, ignore_index=True)
    long_df["ชิ้น"] = pd.to_numeric(long_df["ชิ้น"], errors="coerce").fillna(0)
    long_df["เงิน"] = pd.to_numeric(long_df["เงิน"], errors="coerce").fillna(0)
    long_df = long_df[(long_df["ชิ้น"] != 0) | (long_df["เงิน"] != 0)]
    long_df["รายได้ต่อชิ้น"] = long_df["เงิน"].div(long_df["ชิ้น"]).where(long_df["ชิ้น"] > 0, 0)
    return long_df


def make_long_monthly(monthly):
    long_rows = []
    id_cols = ["สาขา/พื้นที่", "จำนวน ปณ", "รายการบริการ"]

    for index, month in enumerate(MONTHS, start=1):
        pieces_col = f"{month} ชิ้น"
        revenue_col = f"{month} เงิน"
        month_df = monthly[id_cols + [pieces_col, revenue_col]].copy()
        month_df = month_df.rename(columns={pieces_col: "ชิ้น", revenue_col: "เงิน"})
        month_df["เดือน"] = month
        month_df["ลำดับเดือน"] = index
        long_rows.append(month_df)

    long_df = pd.concat(long_rows, ignore_index=True)
    long_df["ชิ้น"] = pd.to_numeric(long_df["ชิ้น"], errors="coerce").fillna(0)
    long_df["เงิน"] = pd.to_numeric(long_df["เงิน"], errors="coerce").fillna(0)
    long_df["รายได้ต่อชิ้น"] = long_df["เงิน"].div(long_df["ชิ้น"]).where(long_df["ชิ้น"] > 0, 0)
    return long_df


def number_format(value):
    return f"{value:,.0f}"


def money_format(value):
    return f"{value:,.2f}"


st.title("สรุปปริมาณงาน รายได้ จากระบบ New CAPOS รายเดือน ป")

with st.sidebar:
    st.header("ข้อมูล")
    uploaded_file = st.file_uploader("เลือกไฟล์ Excel", type=["xlsx"])
    source_file = uploaded_file if uploaded_file is not None else DEFAULT_FILE
    st.caption("ไฟล์เริ่มต้น: " + (uploaded_file.name if uploaded_file else DEFAULT_FILE.name))

try:
    executive_df, detail_df, monthly_df, long_detail_df, long_monthly_df = load_capos_workbook(source_file)
except FileNotFoundError:
    st.error(f"ไม่พบไฟล์เริ่มต้น: {DEFAULT_FILE}")
    st.stop()
except Exception as exc:
    st.error(f"อ่านไฟล์ไม่สำเร็จ: {exc}")
    st.stop()

with st.sidebar:
    st.header("ตัวกรอง")
    area_options = sorted(long_detail_df["สาขา/พื้นที่"].dropna().unique())
    selected_areas = st.multiselect("สาขา/พื้นที่", area_options, default=area_options)

    area_filtered = long_detail_df[long_detail_df["สาขา/พื้นที่"].isin(selected_areas)]

    district_options = sorted(area_filtered["เขต"].dropna().unique())
    selected_districts = st.multiselect("เขต", district_options, default=district_options)

    district_filtered = area_filtered[area_filtered["เขต"].isin(selected_districts)]

    office_options = sorted(district_filtered["ปณ"].dropna().unique())
    selected_offices = st.multiselect("ปณ", office_options, default=office_options)

    office_filtered = district_filtered[district_filtered["ปณ"].isin(selected_offices)]

    service_options = sorted(office_filtered["กลุ่มบริการ"].dropna().unique())
    selected_services = st.multiselect("กลุ่มบริการ", service_options, default=service_options)

    selected_months = st.multiselect("เดือน", MONTHS, default=MONTHS)

filtered = long_detail_df[
    long_detail_df["สาขา/พื้นที่"].isin(selected_areas)
    & long_detail_df["เขต"].isin(selected_districts)
    & long_detail_df["ปณ"].isin(selected_offices)
    & long_detail_df["กลุ่มบริการ"].isin(selected_services)
    & long_detail_df["เดือน"].isin(selected_months)
].copy()

filtered_monthly = long_monthly_df[
    long_monthly_df["สาขา/พื้นที่"].isin(selected_areas)
    & long_monthly_df["เดือน"].isin(selected_months)
].copy()

if filtered.empty:
    st.warning("ไม่มีข้อมูลตามตัวกรองที่เลือก")
    st.stop()

total_pieces = filtered["ชิ้น"].sum()
total_revenue = filtered["เงิน"].sum()
avg_revenue_per_piece = total_revenue / total_pieces if total_pieces else 0

metric_cols = st.columns(5)
metric_cols[0].metric("สาขา/พื้นที่", number_format(filtered["สาขา/พื้นที่"].nunique()))
metric_cols[1].metric("จำนวน ปณ", number_format(filtered["ปณ"].nunique()))
metric_cols[2].metric("จำนวนชิ้น", number_format(total_pieces))
metric_cols[3].metric("รายได้", money_format(total_revenue))
metric_cols[4].metric("รายได้ต่อชิ้น", money_format(avg_revenue_per_piece))

tab_exec, tab_month, tab_service, tab_office, tab_data = st.tabs(
    ["Executive", "รายเดือน", "บริการ", "ปณ/พื้นที่", "ข้อมูลละเอียด"]
)

with tab_exec:
    executive_view = executive_df[executive_df["สาขา/พื้นที่"].isin(selected_areas)].copy()
    executive_view = executive_view[
        [
            "ลำดับ",
            "สาขา/พื้นที่",
            "เหตุผลเชิงยุทธศาสตร์",
            "จำนวน ปณ",
            "รายการบริการ",
            "ชิ้นรวม 2568",
            "เงินรวม 2568",
            "สถานะตรวจสอบ",
        ]
    ]
    st.dataframe(
        executive_view,
        use_container_width=True,
        hide_index=True,
        column_config={
            "จำนวน ปณ": st.column_config.NumberColumn("จำนวน ปณ", format="%,.0f"),
            "รายการบริการ": st.column_config.NumberColumn("รายการบริการ", format="%,.0f"),
            "ชิ้นรวม 2568": st.column_config.NumberColumn("ชิ้นรวม 2568", format="%,.0f"),
            "เงินรวม 2568": st.column_config.NumberColumn("เงินรวม 2568", format="%,.2f"),
        },
    )

with tab_month:
    monthly_by_area = (
        filtered_monthly.groupby(["สาขา/พื้นที่"], as_index=False)[["ชิ้น", "เงิน"]]
        .sum()
        .sort_values("เงิน", ascending=False)
    )
    monthly_trend = (
        filtered_monthly.groupby(["ลำดับเดือน", "เดือน"], as_index=False)[["ชิ้น", "เงิน"]]
        .sum()
        .sort_values("ลำดับเดือน")
    )

    chart_cols = st.columns(2)
    with chart_cols[0]:
        st.subheader("แนวโน้มจำนวนชิ้น")
        st.line_chart(monthly_trend, x="เดือน", y="ชิ้น")
    with chart_cols[1]:
        st.subheader("แนวโน้มรายได้")
        st.line_chart(monthly_trend, x="เดือน", y="เงิน")

    st.subheader("รายได้ตามสาขา/พื้นที่")
    st.bar_chart(monthly_by_area, x="สาขา/พื้นที่", y="เงิน")

with tab_service:
    service_summary = (
        filtered.groupby("กลุ่มบริการ", as_index=False)[["ชิ้น", "เงิน"]]
        .sum()
        .sort_values("เงิน", ascending=False)
    )
    service_summary["รายได้ต่อชิ้น"] = service_summary["เงิน"].div(
        service_summary["ชิ้น"]
    ).where(service_summary["ชิ้น"] > 0, 0)

    st.subheader("Top 20 กลุ่มบริการตามรายได้")
    st.bar_chart(service_summary.head(20), x="กลุ่มบริการ", y="เงิน")
    st.dataframe(
        service_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "ชิ้น": st.column_config.NumberColumn("ชิ้น", format="%,.0f"),
            "เงิน": st.column_config.NumberColumn("เงิน", format="%,.2f"),
            "รายได้ต่อชิ้น": st.column_config.NumberColumn("รายได้ต่อชิ้น", format="%,.2f"),
        },
    )

with tab_office:
    office_summary = (
        filtered.groupby(["สาขา/พื้นที่", "เขต", "รหัสไปรษณีย์", "ปณ"], as_index=False)[["ชิ้น", "เงิน"]]
        .sum()
        .sort_values("เงิน", ascending=False)
    )
    area_summary = (
        filtered.groupby("สาขา/พื้นที่", as_index=False)[["ชิ้น", "เงิน"]]
        .sum()
        .sort_values("เงิน", ascending=False)
    )

    st.subheader("สรุปตามสาขา/พื้นที่")
    st.dataframe(
        area_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "ชิ้น": st.column_config.NumberColumn("ชิ้น", format="%,.0f"),
            "เงิน": st.column_config.NumberColumn("เงิน", format="%,.2f"),
        },
    )

    st.subheader("Top 30 ปณ ตามรายได้")
    st.dataframe(
        office_summary.head(30),
        use_container_width=True,
        hide_index=True,
        column_config={
            "ชิ้น": st.column_config.NumberColumn("ชิ้น", format="%,.0f"),
            "เงิน": st.column_config.NumberColumn("เงิน", format="%,.2f"),
        },
    )

with tab_data:
    display_df = filtered.sort_values(
        ["สาขา/พื้นที่", "เขต", "ปณ", "กลุ่มบริการ", "ลำดับเดือน"]
    )[
        [
            "สาขา/พื้นที่",
            "เขต",
            "รหัสไปรษณีย์",
            "ปณ",
            "กลุ่มบริการ",
            "เดือน",
            "ชิ้น",
            "เงิน",
            "รายได้ต่อชิ้น",
        ]
    ]

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "ชิ้น": st.column_config.NumberColumn("ชิ้น", format="%,.0f"),
            "เงิน": st.column_config.NumberColumn("เงิน", format="%,.2f"),
            "รายได้ต่อชิ้น": st.column_config.NumberColumn("รายได้ต่อชิ้น", format="%,.2f"),
        },
    )

    csv = display_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "ดาวน์โหลด CSV",
        data=csv,
        file_name="new_capos_filtered.csv",
        mime="text/csv",
    )
