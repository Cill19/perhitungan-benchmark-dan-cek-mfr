
import hashlib
import itertools
import io
import logging
import zipfile
from pathlib import PurePosixPath

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side,
)



# ============================================================
# INDONESIA ENGINE
# ============================================================

from benchmark_engine import (
    ENGINE_VERSION,
    BenchmarkInputError,
    read_order_file,
    process_dataframe,
    calculate_benchmark,
    validate_file_calculation,
    validate_final_result,
    extract_record_ids,
)


# ============================================================
# THAILAND ENGINE
# ============================================================

from thailand_engine import (
    process_shopee_th,
)

from thailand_tiktok_engine import (
    process_tiktok_th,
    parse_tiktok_xlsx,
)


# ============================================================
# MFR ENGINE
# ============================================================

from mfr_engine import (
    MFR_DISPLAY_LABELS,
    MFR_METRIC_COLUMNS,
    MFR_ORDER_DETAIL_COLUMNS,
    MFR_ORDER_DETAIL_DISPLAY_LABELS,
    process_dataframe_mfr,
    calculate_mfr,
    validate_mfr_file,
    create_mfr_excel_report,
)


LOGGER = logging.getLogger(__name__)

MFR_ORDER_DETAIL_CACHE_VERSION = "snake_case_discount_components_v1"


# ============================================================
# SUPPORTED FILE
# ============================================================

SUPPORTED_DATA_EXTENSIONS = {
    ".xlsx",
    ".xls",
    ".xlsm",
    ".csv",
    ".txt",
}

def create_excel_report_single_sheet(
    result,
    active_marketplaces,
    is_thailand,
):

    from io import BytesIO

    output = BytesIO()

    wb = Workbook()

    ws = wb.active
    ws.title = "Benchmark Report"


    currency_format = (
        '"฿"#,##0'
        if is_thailand
        else '"Rp "#,##0'
    )


    # =========================
    # STYLE
    # =========================

    section_fill = PatternFill(
        "solid",
        fgColor="D9EAF7"
    )

    header_fill = PatternFill(
        "solid",
        fgColor="BDD7EE"
    )


    bold = Font(
        bold=True
    )


    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )


    row = 1


    def section(title):

        nonlocal row

        ws.cell(
            row=row,
            column=1,
            value=title
        )

        ws.cell(
            row=row,
            column=1
        ).font = bold

        ws.cell(
            row=row,
            column=1
        ).fill = section_fill

        row += 1



    def table(headers, data):

        nonlocal row


        for i, h in enumerate(headers,1):

            cell = ws.cell(
                row=row,
                column=i,
                value=h
            )

            cell.font = bold
            cell.fill = header_fill
            cell.border = thin


        row += 1


        for item in data:

            for i, value in enumerate(item,1):

                cell = ws.cell(
                    row=row,
                    column=i,
                    value=value
                )

                cell.border = thin


            row += 1


        row += 2



    # =========================
    # TITLE
    # =========================

    ws.merge_cells(
        start_row=1,
        start_column=1,
        end_row=1,
        end_column=5
    )


    ws["A1"] = (
        "Benchmark Brand Superstar Report"
    )

    ws["A1"].font = Font(
        bold=True,
        size=14
    )


    row = 3



    # =========================
    # SUMMARY
    # =========================

    section(
        "Benchmark Summary"
    )


    table(
        [
            "Metric",
            "Value"
        ],
        [
            [
                "Benchmark Average",
                result["benchmark"]
            ],
            [
                "Period",
                f"L{result['valid_months']}M"
            ],
            [
                "Total Month",
                result["valid_months"]
            ],
            [
                "Marketplace",
                ", ".join(active_marketplaces)
            ],
            [
                "Currency",
                "THB" if is_thailand else "IDR"
            ],
        ]
    )


    # =========================
    # MARKETPLACE SUMMARY
    # =========================

    section(
        "Ringkasan Benchmark"
    )


    summary = []


    monthly_marketplace = (
        result["monthly_marketplace"]
    )


    for mp in active_marketplaces:

        data = monthly_marketplace[
            monthly_marketplace["marketplace"]
            .str.contains(mp)
        ]


        value = (
            data["benchmark_gmv"].sum()
            /
            result["valid_months"]
            if result["valid_months"]
            else 0
        )


        summary.append(
            [
                mp,
                value
            ]
        )


    summary.append(
        [
            "Benchmark Avg",
            result["benchmark"]
        ]
    )


    table(
        [
            "Marketplace",
            "Benchmark"
        ],
        summary
    )



    # =========================
    # MONTHLY
    # =========================

    section(
        "Benchmark Per Bulan"
    )


    monthly = (
        result["monthly_combined"]
        .copy()
    )


    monthly["month"] = (
        monthly["month"]
        .dt.strftime(
            "%B %Y"
        )
    )


    table(
        list(monthly.columns),
        monthly.values.tolist()
    )



    # =========================
    # DETAIL
    # =========================

    section(
        "Detail Perhitungan Marketplace"
    )


    detail = (
        result["monthly_marketplace"]
        .copy()
    )


    detail["month"] = (
        detail["month"]
        .dt.strftime(
            "%B %Y"
        )
    )


    table(
        list(detail.columns),
        detail.values.tolist()
    )


    # =========================
    # NUMBER FORMAT
    # =========================

    for row_cells in ws.iter_rows():

        for cell in row_cells:

            if isinstance(
                cell.value,
                (int,float)
            ):

                cell.number_format = (
                    currency_format
                )



    # =========================
    # WIDTH
    # =========================

    for column in ws.columns:

        length = max(
            len(str(cell.value))
            if cell.value
            else 0
            for cell in column
        )

        ws.column_dimensions[
            get_column_letter(
                column[0].column
            )
        ].width = length + 5


    ws.freeze_panes = "A4"


    wb.save(output)

    output.seek(0)

    return output

@st.fragment
def render_seasonal_benchmark(
    result,
    currency_symbol,
    include_empty_months,
):
    monthly = result["monthly_combined"].copy()

    monthly["month"] = (
        pd.to_datetime(monthly["month"])
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    monthly = (
        monthly
        .sort_values("month")
        .reset_index(drop=True)
    )

    # ========================================================
    # MONTHS YANG MEMANG MASUK BENCHMARK
    # ========================================================

    eligible_monthly = monthly.copy()

    if not include_empty_months:
        monthly_marketplace = result.get(
            "monthly_marketplace",
            pd.DataFrame(),
        )

        if (
            not monthly_marketplace.empty
            and "source_rows" in monthly_marketplace.columns
        ):
            monthly_marketplace = monthly_marketplace.copy()

            monthly_marketplace["month"] = (
                pd.to_datetime(monthly_marketplace["month"])
                .dt.to_period("M")
                .dt.to_timestamp()
            )

            month_has_rows = (
                monthly_marketplace
                .groupby("month")["source_rows"]
                .sum()
            )

            active_months = (
                month_has_rows[
                    month_has_rows > 0
                ]
                .index
            )

            eligible_monthly = (
                monthly[
                    monthly["month"].isin(active_months)
                ]
                .copy()
            )

    # ========================================================
    # SEASONAL SELECTOR
    # ========================================================

    st.markdown("---")
    st.subheader("Seasonal Benchmark")

    month_options = (
        eligible_monthly["month"]
        .tolist()
    )

    selected_seasonal = st.multiselect(
        "Take Out Bulan Seasonal",
        options=month_options,
        default=[],
        format_func=lambda x: pd.Timestamp(x).strftime("%B %Y"),
        key="seasonal_month",
        placeholder="Pilih satu atau lebih bulan seasonal",
    )

    # ========================================================
    # CALCULATION
    # ========================================================

    if not selected_seasonal:
        regular_data = eligible_monthly.copy()
        seasonal_data = eligible_monthly.iloc[0:0].copy()

    else:
        selected_periods = {
            pd.Timestamp(month)
            .to_period("M")
            .to_timestamp()
            for month in selected_seasonal
        }

        seasonal_mask = (
            eligible_monthly["month"]
            .isin(selected_periods)
        )

        seasonal_data = (
            eligible_monthly[
                seasonal_mask
            ]
            .copy()
        )

        regular_data = (
            eligible_monthly[
                ~seasonal_mask
            ]
            .copy()
        )

    benchmark_regular = (
        float(
            regular_data["total_gmv"].mean()
        )
        if not regular_data.empty
        else None
    )

    benchmark_seasonal = (
        float(
            seasonal_data["total_gmv"].mean()
        )
        if not seasonal_data.empty
        else None
    )

    # ========================================================
    # DETAIL DATA — PRESENTATION ONLY
    # ========================================================

    regular_months = set(
        pd.to_datetime(regular_data["month"])
        .dt.to_period("M")
        .dt.to_timestamp()
    )
    seasonal_months = set(
        pd.to_datetime(seasonal_data["month"])
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    monthly_detail_source = result["monthly_combined"].copy()
    monthly_detail_source["month"] = (
        pd.to_datetime(monthly_detail_source["month"])
        .dt.to_period("M")
        .dt.to_timestamp()
    )
    regular_monthly_detail = (
        monthly_detail_source[
            monthly_detail_source["month"].isin(regular_months)
        ]
        .sort_values("month")
        .reset_index(drop=True)
        .copy()
    )
    seasonal_monthly_detail = (
        monthly_detail_source[
            monthly_detail_source["month"].isin(seasonal_months)
        ]
        .sort_values("month")
        .reset_index(drop=True)
        .copy()
    )

    marketplace_detail_source = result["monthly_marketplace"].copy()
    marketplace_detail_source["month"] = (
        pd.to_datetime(marketplace_detail_source["month"])
        .dt.to_period("M")
        .dt.to_timestamp()
    )
    regular_marketplace_detail = (
        marketplace_detail_source[
            marketplace_detail_source["month"].isin(regular_months)
        ]
        .sort_values(["month", "marketplace"], kind="stable")
        .reset_index(drop=True)
        .copy()
    )
    seasonal_marketplace_detail = (
        marketplace_detail_source[
            marketplace_detail_source["month"].isin(seasonal_months)
        ]
        .sort_values(["month", "marketplace"], kind="stable")
        .reset_index(drop=True)
        .copy()
    )

    def format_currency(value):
        numeric_value = pd.to_numeric(value, errors="coerce")
        if pd.isna(numeric_value):
            return "-"
        if currency_symbol == "Rp":
            return f"Rp {numeric_value:,.0f}".replace(",", ".")
        return f"{currency_symbol} {numeric_value:,.0f}"

    monthly_labels = {
        "month": "Bulan",
        "shopee": "Shopee",
        "tiktok": "TikTok",
        "tokopedia": "Tokopedia",
        "lazada": "Lazada",
        "total_gmv": "Total Benchmark GMV",
    }
    marketplace_labels = {
        "marketplace": "Marketplace",
        "month": "Bulan",
        "gmv_all": "GMV All",
        "gmv_ex_cancel": "GMV Exc Batal",
        "platform_discount": "Platform Discount",
        "seller_rebate": "Seller Rebate",
        "benchmark_gmv": "Benchmark GMV",
        "source_rows": "Source Rows",
    }

    def prepare_monthly_display(detail):
        display = detail.copy()
        visible_columns = [
            column for column in monthly_labels
            if column in display.columns
        ]
        display = display[visible_columns].copy()
        display["month"] = pd.to_datetime(
            display["month"]
        ).dt.strftime("%B %Y")
        for column in visible_columns:
            if column != "month":
                display[column] = display[column].apply(format_currency)
        return display.rename(columns=monthly_labels)

    def prepare_marketplace_display(detail):
        display = detail.copy()
        visible_columns = [
            column for column in marketplace_labels
            if column in display.columns
        ]
        display = display[visible_columns].copy()
        display["month"] = pd.to_datetime(
            display["month"]
        ).dt.strftime("%B %Y")
        if currency_symbol == "฿":
            display["marketplace"] = display["marketplace"].replace(
                {"SHO_TH": "SHO", "TIK_TH": "TIK"}
            )
        currency_columns = [
            "gmv_all",
            "gmv_ex_cancel",
            "platform_discount",
            "seller_rebate",
            "benchmark_gmv",
        ]
        for column in currency_columns:
            if column in display.columns:
                display[column] = display[column].apply(format_currency)
        if "source_rows" in display.columns:
            display["source_rows"] = (
                pd.to_numeric(display["source_rows"], errors="coerce")
                .fillna(0)
                .astype("int64")
            )
        return display.rename(columns=marketplace_labels)

    regular_monthly_display = prepare_monthly_display(
        regular_monthly_detail
    )
    seasonal_monthly_display = prepare_monthly_display(
        seasonal_monthly_detail
    )
    regular_marketplace_display = prepare_marketplace_display(
        regular_marketplace_detail
    )
    seasonal_marketplace_display = prepare_marketplace_display(
        seasonal_marketplace_detail
    )

    # ========================================================
    # DISPLAY
    # ========================================================

    st.markdown("")
    st.subheader("Detail Perhitungan Benchmark")

    tab_regular, tab_seasonal = st.tabs(
        ["Benchmark Biasa", "Benchmark Seasonal"]
    )

    with tab_regular:
        st.metric(
            "Benchmark Biasa",
            format_currency(benchmark_regular)
            if benchmark_regular is not None
            else "-",
        )
        st.caption(
            f"Average {len(regular_data)} bulan non-seasonal"
        )
        st.subheader("Benchmark per Bulan")
        if regular_monthly_display.empty:
            st.info("Tidak ada bulan non-seasonal yang digunakan.")
        else:
            st.dataframe(
                regular_monthly_display,
                hide_index=True,
                width="stretch",
            )
        st.subheader("Detail Perhitungan Marketplace")
        if regular_marketplace_display.empty:
            st.info("Tidak ada detail marketplace non-seasonal.")
        else:
            st.dataframe(
                regular_marketplace_display,
                hide_index=True,
                width="stretch",
            )
        if benchmark_regular is not None:
            st.info(
                f"Benchmark Biasa = Average {len(regular_data)} bulan = "
                f"{format_currency(benchmark_regular)}"
            )

    with tab_seasonal:
        if seasonal_data.empty:
            st.info("Belum ada bulan seasonal yang dipilih.")
        else:
            st.metric(
                "Benchmark Seasonal",
                format_currency(benchmark_seasonal),
            )
            st.caption(
                f"Average {len(seasonal_data)} bulan seasonal"
            )
            st.subheader("Benchmark per Bulan")
            st.dataframe(
                seasonal_monthly_display,
                hide_index=True,
                width="stretch",
            )
            st.subheader("Detail Perhitungan Marketplace")
            st.dataframe(
                seasonal_marketplace_display,
                hide_index=True,
                width="stretch",
            )
            st.info(
                f"Benchmark Seasonal = Average {len(seasonal_data)} bulan = "
                f"{format_currency(benchmark_seasonal)}"
            )

# ============================================================
# ZIP SAFETY LIMIT
# ============================================================

@st.fragment
def render_benchmark_adjustment(
    result,
    currency_symbol,
    include_empty_months,
):
    # ========================================================
    # PREPARE MONTHLY DATA
    # ========================================================

    monthly = result["monthly_combined"].copy()

    monthly["month"] = (
        pd.to_datetime(monthly["month"])
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    monthly = (
        monthly
        .sort_values("month")
        .reset_index(drop=True)
    )

    # Pastikan total_gmv selalu numeric untuk kalkulasi
    monthly["total_gmv"] = pd.to_numeric(
        monthly["total_gmv"],
        errors="coerce",
    )


    # ========================================================
    # DETERMINE ELIGIBLE MONTHS
    # ========================================================

    eligible_monthly = monthly.copy()

    if not include_empty_months:

        monthly_marketplace = result.get(
            "monthly_marketplace",
            pd.DataFrame(),
        )

        if (
            not monthly_marketplace.empty
            and "source_rows" in monthly_marketplace.columns
        ):
            monthly_marketplace = (
                monthly_marketplace.copy()
            )

            monthly_marketplace["month"] = (
                pd.to_datetime(
                    monthly_marketplace["month"]
                )
                .dt.to_period("M")
                .dt.to_timestamp()
            )

            month_has_rows = (
                monthly_marketplace
                .groupby("month")["source_rows"]
                .sum()
            )

            active_months = (
                month_has_rows[
                    month_has_rows > 0
                ]
                .index
            )

            eligible_monthly = (
                monthly[
                    monthly["month"].isin(
                        active_months
                    )
                ]
                .copy()
            )

    # ========================================================
    # UI
    # ========================================================

    st.markdown("---")
    st.subheader("Adjustment Benchmark")

    st.caption(
        "Pilih satu atau lebih bulan yang ingin "
        "dikeluarkan dari perhitungan benchmark."
    )

    month_options = (
        eligible_monthly["month"]
        .tolist()
    )

    selected_takeout = st.multiselect(
        "Take Out Bulan dari Perhitungan",
        options=month_options,
        default=[],
        format_func=lambda x: (
            pd.Timestamp(x).strftime("%B %Y")
        ),
        key="benchmark_takeout_months",
        placeholder="Pilih satu atau lebih bulan",
    )

    # ========================================================
    # SPLIT INCLUDED / EXCLUDED
    # ========================================================

    if selected_takeout:

        selected_periods = {
            pd.Timestamp(month)
            .to_period("M")
            .to_timestamp()
            for month in selected_takeout
        }

        excluded_mask = (
            eligible_monthly["month"]
            .isin(selected_periods)
        )

        excluded_data = (
            eligible_monthly[
                excluded_mask
            ]
            .copy()
        )

        included_data = (
            eligible_monthly[
                ~excluded_mask
            ]
            .copy()
        )

    else:

        included_data = (
            eligible_monthly.copy()
        )

        excluded_data = (
            eligible_monthly.iloc[0:0]
            .copy()
        )

    # ========================================================
    # CURRENCY FORMATTER
    # ========================================================

    def format_currency(value):

        if value is None:
            return "-"

        if callable(value):
            raise TypeError(
                f"format_currency menerima function object: {value}"
            )

        if isinstance(value, str):

            text = value.strip()

            if not text:
                return "-"

            if text.startswith("Rp") or text.startswith("฿"):
                return text

            try:
                numeric_value = float(text)

            except (ValueError, TypeError):
                raise TypeError(
                    f"format_currency menerima value non-numeric: "
                    f"{text!r} ({type(text).__name__})"
                )

        else:

            try:
                numeric_value = float(value)

            except (TypeError, ValueError):
                raise TypeError(
                    f"format_currency menerima value non-numeric: "
                    f"{value!r} ({type(value).__name__})"
                )

        if pd.isna(numeric_value):
            return "-"

        if currency_symbol == "Rp":
            return (
                f"Rp {numeric_value:,.0f}"
                .replace(",", ".")
            )

        return f"{currency_symbol} {numeric_value:,.0f}"

    # ========================================================
    # ORIGINAL BENCHMARK
    # ========================================================

    original_benchmark = (
        float(
            eligible_monthly["total_gmv"].mean()
        )
        if not eligible_monthly.empty
        else None
    )

    original_month_count = len(
        eligible_monthly
    )

    # ========================================================
    # EDGE CASE:
    # ALL MONTHS TAKEN OUT
    # ========================================================

    if included_data.empty:

        st.warning(
            "Minimal 1 bulan harus tetap digunakan "
            "untuk menghitung benchmark."
        )

        adjusted_benchmark = None

    else:

        adjusted_benchmark = float(
            included_data["total_gmv"].mean()
        )

    # ========================================================
    # METRICS
    # ========================================================

    st.markdown(
        """
        <style>
        .st-key-adjustment_metrics [data-testid="stMetricValue"] {
            font-size: 1.35rem !important;
            line-height: 1.2 !important;
            white-space: nowrap !important;
        }

        .st-key-adjustment_metrics [data-testid="stMetricLabel"] {
            font-size: 0.90rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    with st.container(key="adjustment_metrics"):

        m1, m2, m3, m4, m5 = st.columns(
            [1.35, 1.55, 1, 1, 0.9]
        )

        m1.metric(
            "Benchmark Awal",
            format_currency(original_benchmark),
            border=True,
        )

        m2.metric(
            "Benchmark Setelah Take Out",
            format_currency(adjusted_benchmark),
            border=True,
        )

        m3.metric(
            "Periode Awal",
            f"{original_month_count} bulan",
            border=True,
        )

        m4.metric(
            "Periode Digunakan",
            f"{len(included_data)} bulan",
            border=True,
        )

        m5.metric(
            "Take Out",
            f"{len(excluded_data)} bulan",
            border=True,
        )

    # ========================================================
    # DETAIL CALCULATION
    # ========================================================

    st.markdown("")
    st.subheader("Detail Adjustment Benchmark")

    col_included, col_excluded = st.columns(2)

    with col_included:

        st.markdown("#### Bulan yang Digunakan")

        included_display = (
            included_data[
                ["month", "total_gmv"]
            ]
            .copy()
        )

        included_display["month"] = (
            included_display["month"]
            .dt.strftime("%B %Y")
        )

        included_display = (
            included_display.rename(
                columns={
                    "month": "Bulan",
                    "total_gmv": "Benchmark GMV",
                }
            )
        )

        included_display["Benchmark GMV"] = (
            included_display["Benchmark GMV"]
            .apply(format_currency)
        )

        st.dataframe(
            included_display,
            hide_index=True,
            width="stretch",
        )

        if adjusted_benchmark is not None:

            st.info(
                f"Benchmark = Average "
                f"{len(included_data)} bulan = "
                f"{format_currency(adjusted_benchmark)}"
            )

    # ========================================================
    # EXCLUDED MONTHS
    # ========================================================

    with col_excluded:

        st.markdown(
            "#### Bulan yang Di-take Out"
        )

        if not excluded_data.empty:

            excluded_display = (
                excluded_data[
                    ["month", "total_gmv"]
                ]
                .copy()
            )

            excluded_display["month"] = (
                excluded_display["month"]
                .dt.strftime("%B %Y")
            )

            excluded_display = (
                excluded_display.rename(
                    columns={
                        "month": "Bulan",
                        "total_gmv":
                            "Benchmark GMV",
                    }
                )
            )

            excluded_display[
                "Benchmark GMV"
            ] = (
                excluded_display[
                    "Benchmark GMV"
                ]
                .apply(format_currency)
            )

            st.dataframe(
                excluded_display,
                hide_index=True,
                width="stretch",
            )

            takeout_labels = ", ".join(
                pd.Timestamp(month)
                .strftime("%B %Y")
                for month in selected_takeout
            )

            st.info(
                f"{len(excluded_data)} bulan "
                f"dikeluarkan: "
                f"{takeout_labels}"
            )

        else:

            st.info(
                "Belum ada bulan yang di-take out."
            )

    # ========================================================
    # BENCHMARK PER BULAN SETELAH TAKE OUT
    # ========================================================

    st.markdown("")
    st.subheader("Benchmark per Bulan Setelah Take Out")

    if included_data.empty:

        st.info(
            "Tidak ada bulan yang digunakan. "
            "Pilih minimal 1 bulan untuk menampilkan tabel."
        )

    else:

        adjusted_monthly_display = included_data.copy()

        adjusted_monthly_display["month"] = (
            adjusted_monthly_display["month"]
            .dt.strftime("%B %Y")
        )

        adjusted_monthly_display = (
            adjusted_monthly_display.rename(
                columns={
                    "month": "Bulan",
                    "shopee": "Shopee",
                    "tiktok": "TikTok",
                    "tokopedia": "Tokopedia",
                    "lazada": "Lazada",
                    "total_gmv": "Total Benchmark GMV",
                }
            )
        )

        money_columns = [
            "Shopee",
            "TikTok",
            "Tokopedia",
            "Lazada",
            "Total Benchmark GMV",
        ]

        for col in money_columns:

            if col in adjusted_monthly_display.columns:

                adjusted_monthly_display[col] = (
                    adjusted_monthly_display[col]
                    .apply(format_currency)
                )

        st.dataframe(
            adjusted_monthly_display,
            hide_index=True,
            width="stretch",
        )

        if adjusted_benchmark is not None:

            st.info(
                f"Benchmark Setelah Take Out = "
                f"Average {len(included_data)} bulan = "
                f"{format_currency(adjusted_benchmark)}"
            )

MAX_ZIP_DATA_FILES = 100

MAX_ZIP_MEMBER_BYTES = (
    300 * 1024 * 1024
)

MAX_ZIP_TOTAL_BYTES = (
    1024 * 1024 * 1024
)

MAX_ZIP_COMPRESSION_RATIO = 300.0

# ============================================================
# ZIP HELPER
# ============================================================

def _is_ignored_zip_member(name: str) -> bool:

    normalized = name.replace("\\", "/")

    parts = PurePosixPath(normalized).parts

    basename = PurePosixPath(normalized).name


    if not basename:
        return True


    if "__MACOSX" in parts:
        return True


    if basename in {
        ".DS_Store",
        "Thumbs.db"
    }:
        return True


    if basename.startswith("._"):
        return True


    return False



def _validate_zip_members(
    zf: zipfile.ZipFile,
    zip_name: str
):

    data_members = []

    total_uncompressed = 0

    nested_zip_members = []


    for info in zf.infolist():


        if info.is_dir():
            continue


        member_name = info.filename


        if _is_ignored_zip_member(member_name):
            continue


        suffix = PurePosixPath(
            member_name
        ).suffix.lower()



        if suffix == ".zip":

            nested_zip_members.append(
                member_name
            )

            continue



        if suffix not in SUPPORTED_DATA_EXTENSIONS:

            continue



        if info.file_size > MAX_ZIP_MEMBER_BYTES:

            raise BenchmarkInputError(
                f"ZIP '{zip_name}' terlalu besar."
            )


        total_uncompressed += info.file_size


        if total_uncompressed > MAX_ZIP_TOTAL_BYTES:

            raise BenchmarkInputError(
                f"ZIP '{zip_name}' melebihi batas."
            )


        data_members.append(info)



    if nested_zip_members:

        raise BenchmarkInputError(
            "Nested ZIP tidak didukung."
        )


    if not data_members:

        raise BenchmarkInputError(
            f"ZIP '{zip_name}' tidak berisi file data."
        )


    return data_members

def iter_uploaded_data_files(uploaded_file):

    outer_name = uploaded_file.name

    outer_suffix = PurePosixPath(
        outer_name
    ).suffix.lower()


    outer_bytes = uploaded_file.getvalue()



    # ==========================
    # NORMAL FILE
    # ==========================

    if outer_suffix != ".zip":


        stream = io.BytesIO(
            outer_bytes
        )

        stream.name = outer_name


        yield {

            "stream":
                stream,

            "filename":
                outer_name,

            "display_name":
                outer_name,

            "raw_bytes":
                outer_bytes,

            "container":
                None,

        }

        return



    # ==========================
    # ZIP FILE
    # ==========================

    try:

        zf = zipfile.ZipFile(
            io.BytesIO(
                outer_bytes
            )
        )


    except zipfile.BadZipFile:


        raise BenchmarkInputError(
            "ZIP corrupt."
        )



    with zf:


        members = _validate_zip_members(
            zf,
            outer_name,
        )



        for info in members:


            member_bytes = zf.read(
                info
            )


            inner_name = PurePosixPath(
                info.filename
            ).name



            stream = io.BytesIO(
                member_bytes
            )

            stream.name = inner_name



            yield {

                "stream":
                    stream,

                "filename":
                    inner_name,

                "display_name":
                    f"{outer_name} → {info.filename}",

                "raw_bytes":
                    member_bytes,

                "container":
                    outer_name,

            }

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Benchmark Brand Superstar",
    page_icon="📊",
    layout="wide",
)



# ============================================================
# COUNTRY SWITCH
# ============================================================

st.title(
    "📊 Benchmark Brand Superstar"
)


country = st.radio(
    "Select Market",
    [
        "Indonesia 🇮🇩",
        "Thailand 🇹🇭",
    ],
    horizontal=True,
)


is_thailand = (
    country == "Thailand 🇹🇭"
)


if is_thailand:
    calculation_mode = "Benchmark"
    st.caption("ℹ️ MFR saat ini tersedia untuk Indonesia.")
else:
    calculation_mode = st.radio(
        "Jenis Perhitungan",
        [
            "Benchmark",
            "MFR",
        ],
        horizontal=True,
        key="calc_mode",
    )

is_mfr = (calculation_mode == "MFR") and (not is_thailand)


st.caption(
    f"{'MFR Engine' if is_mfr else f'Benchmark Engine v{ENGINE_VERSION}'} | "
    f"{'Thailand' if is_thailand else 'Indonesia'}"
)

# ============================================================
# LOGIC DESCRIPTION
# ============================================================

with st.expander(
    "Logic " + ("MFR" if is_mfr else "Benchmark"),
    expanded=False
):

    if is_mfr:

        st.markdown(
            """
            - **Prinsip MFR (Monthly Final Revenue):** omzet untuk **1 bulan kalender terpilih**.
            - **Semua pesanan aktif masuk omzet** selama belum/tidak dibatalkan (Diproses, Dikemas, Dikirim, Dalam Transit, Selesai, dll.).
            - **Shopee:** Subtotal non-batal − Voucher Penjual + Diskon Shopee
            - **TikTok:** order non-batal berdasarkan full cancel rule → SKU Subtotal After Discount + SKU Platform Discount
            - **Tokopedia:** order non-batal berdasarkan full cancel rule → SKU Subtotal After Discount + SKU Platform Discount
            - **Lazada:** unitPrice non-cancel + platformDiscountTotal
            - **Final:** Penjumlahan MFR seluruh marketplace pada bulan tersebut (tanpa average multi-bulan).
            """
        )

    elif is_thailand:

        st.markdown(
            """
            - **Shopee Thailand:** GMV exclude cancel − Seller Rebate
            - **Currency:** THB (฿)
            - **Final:** jumlah seluruh marketplace per bulan → average seluruh bulan valid
            """
        )

    else:

        st.markdown(
            """
            - **Shopee:** GMV exclude batal − Voucher Penjual + Diskon Shopee
            - **TikTok:** Order Sub Status = Selesai → SKU Subtotal After Discount + Platform Discount
            - **Tokopedia:** Order Sub Status = Selesai → SKU Subtotal After Discount + Platform Discount
            - **Lazada:** GMV exclude batal + Platform Discount
            - **Final:** jumlah seluruh marketplace per bulan → average seluruh bulan valid
            """
        )


if is_mfr:
    st.info(
        "Zero Mistake Mode aktif: sistem akan memvalidasi data MFR, status pesanan, "
        "rekonsiliasi nominal per marketplace, dan verifikasi status pembatalan."
    )
else:
    st.info(
        "Zero Mistake Mode aktif: sistem akan melakukan reconciliation, "
        "parsing checks, order consistency checks, dan cross-file duplicate checks "
        "sebelum benchmark diloloskan."
    )



# ============================================================
# UPLOAD SECTION
# ============================================================

st.subheader(
    "Upload Data Pesanan"
)


st.caption(
    "Bisa upload file biasa, single ZIP, multiple ZIP, "
    "atau kombinasi keduanya. "
    "Isi ZIP diproses langsung di RAM tanpa diextract ke storage."
)


col1, col2 = st.columns(2)



# ============================================================
# INDONESIA
# ============================================================

if not is_thailand:

    with col1:

        sho_files = st.file_uploader(
            "Shopee Indonesia",
            type=[
                "xlsx",
                "xls",
                "xlsm",
                "csv",
                "txt",
                "zip",
            ],
            accept_multiple_files=True,
            key="sho",
        )


        tik_files = st.file_uploader(
            "TikTok Indonesia",
            type=[
                "xlsx",
                "xls",
                "xlsm",
                "csv",
                "txt",
                "zip",
            ],
            accept_multiple_files=True,
            key="tik",
        )


    with col2:

        tok_files = st.file_uploader(
            "Tokopedia Indonesia",
            type=[
                "xlsx",
                "xls",
                "xlsm",
                "csv",
                "txt",
                "zip",
            ],
            accept_multiple_files=True,
            key="tok",
        )


        laz_files = st.file_uploader(
            "Lazada Indonesia",
            type=[
                "xlsx",
                "xls",
                "xlsm",
                "csv",
                "txt",
                "zip",
            ],
            accept_multiple_files=True,
            key="laz",
        )



# ============================================================
# THAILAND
# ============================================================

else:

    with col1:

        sho_files = st.file_uploader(
            "Shopee Thailand",
            type=[
                "xlsx",
                "xls",
                "xlsm",
                "csv",
                "txt",
                "zip",
            ],
            accept_multiple_files=True,
            key="sho_th",
        )

        tik_files = st.file_uploader(
            "TikTok Thailand",
            type=[
                "xlsx",
                "xls",
                "xlsm",
                "csv",
                "txt",
                "zip",
            ],
            accept_multiple_files=True,
            key="tik_th",
        )


    with col2:

        tok_files = []
        laz_files = []



# ============================================================
# GROUP UPLOAD
# ============================================================

if is_thailand:

    uploaded_groups = {

        "SHO":
            sho_files or [],

        "TIK":
            tik_files or [],

    }


else:

    uploaded_groups = {

        "SHO":
            sho_files or [],

        "TIK":
            tik_files or [],

        "TOK":
            tok_files or [],

        "LAZ":
            laz_files or [],

    }



selected_uploads = sum(
    len(files)
    for files in uploaded_groups.values()
)


active_marketplaces = [
    mp
    for mp, files in uploaded_groups.items()
    if files
]


if selected_uploads:

    st.info(
        f"{selected_uploads} upload item dipilih dari "
        f"{len(active_marketplaces)} marketplace."
    )

else:

    st.info(
        "Upload minimal 1 file untuk mulai menghitung."
    )


# ============================================================
# MFR FLOW
# ============================================================

def run_mfr_flow(
    uploaded_groups: dict[str, list],
    active_marketplaces: list[str],
    selected_uploads: int,
):
    st.markdown("---")
    st.caption("ℹ️ MFR menghitung omzet untuk 1 bulan kalender berdasarkan data pesanan yang diupload.")

    calculate_mfr_btn = st.button(
        "Validasi & Hitung MFR",
        type="primary",
        disabled=selected_uploads == 0,
        key="btn_calc_mfr",
    )

    if calculate_mfr_btn:
        mfr_file_results = []
        processed_info = []
        file_errors = []
        audit_rows = []

        with st.spinner("Memvalidasi dan memproses data MFR..."):
            for marketplace, files in uploaded_groups.items():
                for uploaded_file in files:
                    try:
                        expanded_files = iter_uploaded_data_files(uploaded_file)
                        for expanded in expanded_files:
                            display_name = expanded["display_name"]
                            raw_bytes = expanded["raw_bytes"]
                            stream = expanded["stream"]
                            filename = expanded["filename"]

                            try:
                                stream.seek(0)
                                df, detected_marketplace, header_row = read_order_file(
                                    stream,
                                    filename=filename,
                                    marketplace_hint=(
                                        marketplace
                                        if marketplace in {"TIK", "TOK"}
                                        else None
                                    ),
                                )

                                if (
                                    marketplace in {"SHO", "LAZ"}
                                    and detected_marketplace != marketplace
                                ):
                                    raise BenchmarkInputError(
                                        f"File ini terdeteksi sebagai "
                                        f"{detected_marketplace}, bukan {marketplace}."
                                    )

                                file_res = process_dataframe_mfr(
                                    df,
                                    marketplace,
                                    filename=filename,
                                )
                                mfr_file_results.append(file_res)

                                file_audit = validate_mfr_file(
                                    df,
                                    marketplace,
                                    file_res,
                                )

                                for item in file_audit:
                                    audit_rows.append({
                                        "Scope": "File",
                                        "Marketplace": marketplace,
                                        "File": display_name,
                                        **item,
                                    })

                                m_df = file_res.get("monthly", pd.DataFrame())
                                processed_info.append({
                                    "Status": "✅ OK",
                                    "Marketplace": marketplace,
                                    "File": display_name,
                                    "Rows": int(
                                        m_df["source_rows"].sum()
                                        if "source_rows" in m_df.columns
                                        else len(df)
                                    ),
                                    "Header Row": header_row + 1,
                                    "Periode Awal": (
                                        m_df["month"].dropna().min().strftime("%b %Y")
                                        if not m_df.empty and not m_df["month"].dropna().empty
                                        else "-"
                                    ),
                                    "Periode Akhir": (
                                        m_df["month"].dropna().max().strftime("%b %Y")
                                        if not m_df.empty and not m_df["month"].dropna().empty
                                        else "-"
                                    ),
                                })

                            except Exception as exc:
                                LOGGER.exception(
                                    "MFR file processing failed: marketplace=%s file=%s",
                                    marketplace,
                                    display_name,
                                )
                                file_errors.append({
                                    "Marketplace": marketplace,
                                    "File": display_name,
                                    "Error": str(exc),
                                })
                                processed_info.append({
                                    "Status": "❌ ERROR",
                                    "Marketplace": marketplace,
                                    "File": display_name,
                                    "Rows": "-",
                                    "Header Row": "-",
                                    "Periode Awal": "-",
                                    "Periode Akhir": "-",
                                })

                    except Exception as exc:
                        LOGGER.exception(
                            "MFR upload expansion failed: marketplace=%s file=%s",
                            marketplace,
                            uploaded_file.name,
                        )
                        file_errors.append({
                            "Marketplace": marketplace,
                            "File": uploaded_file.name,
                            "Error": str(exc),
                        })

        st.session_state["mfr_cache"] = {
            "order_detail_cache_version": MFR_ORDER_DETAIL_CACHE_VERSION,
            "order_detail_schema": tuple(MFR_ORDER_DETAIL_COLUMNS),
            "file_results": mfr_file_results,
            "processed_info": processed_info,
            "file_errors": file_errors,
            "audit_rows": audit_rows,
        }

    if "mfr_cache" in st.session_state and st.session_state["mfr_cache"].get("file_results"):
        cache = st.session_state["mfr_cache"]
        cache_schema_is_current = (
            cache.get("order_detail_cache_version")
            == MFR_ORDER_DETAIL_CACHE_VERSION
            and tuple(cache.get("order_detail_schema", ()))
            == tuple(MFR_ORDER_DETAIL_COLUMNS)
        )
        if not cache_schema_is_current:
            del st.session_state["mfr_cache"]
            st.info(
                "Schema Daftar Pesanan telah diperbarui. "
                "Klik kembali Validasi & Hitung MFR untuk memproses ulang upload."
            )
            return
        processed_info = cache["processed_info"]
        file_errors = cache["file_errors"]
        file_results = cache["file_results"]
        audit_rows = cache["audit_rows"]

        st.subheader("Validasi File MFR")
        processed_display = pd.DataFrame(processed_info).copy()
        for column in processed_display.columns:
            processed_display[column] = (
                processed_display[column].astype("string").fillna("")
            )
        st.dataframe(
            processed_display,
            hide_index=True,
            width="stretch",
        )

        if file_errors:
            st.error(f"{len(file_errors)} file gagal divalidasi.")
            for error in file_errors:
                st.error(f"**{error['Marketplace']} — {error['File']}**\n\n{error['Error']}")
            st.stop()

        # Find distinct months
        all_months = set()
        for res in file_results:
            m_df = res.get("monthly")
            if m_df is not None and not m_df.empty:
                all_months.update(m_df["month"].dropna().tolist())

        sorted_months = sorted(all_months)
        if not sorted_months:
            st.warning("Tidak ada bulan transaksi yang valid ditemukan pada file upload.")
            return

        st.markdown("---")
        st.subheader("Pilih Bulan MFR")
        st.caption("MFR menghitung omzet 1 bulan kalender. Silakan pilih bulan yang akan dihitung:")

        selected_month = st.selectbox(
            "Pilih Bulan MFR",
            options=sorted_months,
            index=len(sorted_months) - 1,
            format_func=lambda x: pd.Timestamp(x).strftime("%B %Y"),
            key="mfr_month_picker",
        )

        # Calculate MFR for selected month
        try:
            mfr_result = calculate_mfr(file_results, selected_month)

            st.success(f"MFR berhasil dihitung untuk periode {mfr_result['month_label']}.")

            # Excel download
            excel_bytes = create_mfr_excel_report(mfr_result)
            file_label = mfr_result["month_label"].replace(" ", "_")
            st.download_button(
                label="📥 Download MFR Report (.xlsx)",
                data=excel_bytes,
                file_name=f"MFR_Report_ID_{file_label}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # Metric cards
            st.subheader("Hasil MFR")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(
                "Total MFR",
                f"Rp {mfr_result['mfr_total']:,.0f}".replace(",", "."),
                border=True,
            )
            c2.metric(
                "Periode",
                mfr_result["month_label"],
                border=True,
            )
            c3.metric(
                "Marketplace",
                f"{mfr_result['marketplace_count']}",
                border=True,
            )
            c4.metric(
                "Rows Digunakan",
                f"{mfr_result['total_rows']:,}".replace(",", "."),
                border=True,
            )

            # Summary per marketplace
            st.subheader("Ringkasan MFR per Marketplace")
            sum_df = mfr_result["marketplace_summary"].copy().rename(
                columns=MFR_DISPLAY_LABELS
            )
            for col in [MFR_DISPLAY_LABELS[column] for column in MFR_METRIC_COLUMNS]:
                sum_df[col] = sum_df[col].apply(
                    lambda v: f"Rp {v:,.0f}".replace(",", ".")
                )
            st.dataframe(sum_df, hide_index=True, width="stretch")

            # Detail perhitungan marketplace
            st.subheader("Detail Perhitungan MFR")
            det_df = mfr_result["marketplace_detail"].copy().rename(
                columns=MFR_DISPLAY_LABELS
            )
            for col in [MFR_DISPLAY_LABELS[column] for column in MFR_METRIC_COLUMNS]:
                det_df[col] = det_df[col].apply(
                    lambda v: f"Rp {v:,.0f}".replace(",", ".")
                )
            for col in ["Rows Total", "Rows Included", "Rows Excluded"]:
                det_df[col] = det_df[col].apply(
                    lambda v: f"{int(v):,}".replace(",", ".")
                )
            st.dataframe(det_df, hide_index=True, width="stretch")

            # Line-item detail following the reference Daftar Pesanan layout.
            st.subheader("Daftar Pesanan")
            order_df = mfr_result["order_detail"].copy()
            order_display = order_df.rename(
                columns=MFR_ORDER_DETAIL_DISPLAY_LABELS
            )
            visible_order_columns = list(
                MFR_ORDER_DETAIL_DISPLAY_LABELS.values()
            )
            order_display = order_display.reindex(
                columns=visible_order_columns
            )
            order_display["Tanggal"] = pd.to_datetime(
                order_display["Tanggal"], errors="coerce"
            ).apply(
                lambda value: ""
                if pd.isna(value)
                else f"{value.day} {value.strftime('%b %Y')}"
            )
            for column in [
                "Harga Terbayarkan",
                "Seller Rebate",
                "Diskon Shopee",
                "Platform Discount",
            ]:
                order_display[column] = order_display[column].apply(
                    lambda value: f"Rp {value:,.0f}".replace(",", ".")
                )
            for column in [
                "Marketplace",
                "Nomor Pesanan",
                "SKU",
                "Status Pesanan",
                "Alasan Pembatalan",
            ]:
                order_display[column] = (
                    order_display[column].astype("string").fillna("")
                )
            st.dataframe(order_display, hide_index=True, width="stretch")

            # Breakdown status pesanan
            st.subheader("Breakdown Status Pesanan")
            st.caption(
                "Menampilkan status aktual dari file data untuk bulan terpilih. "
                "Pesanan aktif (belum dibatalkan) berstatus **Included**, "
                "sedangkan pesanan dibatalkan/invalid berstatus **Excluded**."
            )
            st_df = mfr_result["status_breakdown"].copy().rename(
                columns=MFR_DISPLAY_LABELS
            )
            if not st_df.empty:
                st_df["Net GMV"] = st_df["Net GMV"].apply(
                    lambda v: f"Rp {v:,.0f}".replace(",", ".")
                )
                st_df["Rows"] = st_df["Rows"].apply(
                    lambda v: f"{int(v):,}".replace(",", ".")
                )
            st.dataframe(st_df, hide_index=True, width="stretch")

            # Zero mistake audit
            with st.expander("Zero Mistake Audit MFR", expanded=True):
                combined_audit = audit_rows + mfr_result["audit"]
                audit_display = pd.DataFrame(combined_audit)
                for col in ["Actual", "Expected"]:
                    if col in audit_display.columns:
                        audit_display[col] = audit_display[col].apply(
                            lambda value: "" if pd.isna(value) else str(value)
                        )
                st.dataframe(audit_display, hide_index=True, width="stretch")

        except Exception as exc:
            LOGGER.exception("MFR final calculation or rendering failed")
            st.error(f"Gagal menghitung MFR: {exc}")


if is_mfr:
    run_mfr_flow(
        uploaded_groups=uploaded_groups,
        active_marketplaces=active_marketplaces,
        selected_uploads=selected_uploads,
    )
    st.stop()


# ============================================================
# CALCULATION OPTIONS
# ============================================================

st.markdown("---")

include_empty_months = st.checkbox(
    "Include bulan tanpa transaksi dalam perhitungan benchmark",
    value=True,
    key="include_empty_months",
    help=(
        "Jika aktif, bulan tanpa transaksi tetap dihitung sebagai "
        + ("฿0" if is_thailand else "Rp0")
        + " agar average menggunakan seluruh periode yang dipilih."
    ),
)

if include_empty_months:
    st.caption(
        "🟢 **Include bulan kosong: ON** — "
        "Semua bulan yang diupload dihitung, termasuk bulan tanpa order (nilai = 0)."
    )
else:
    st.caption(
        "🟡 **Include bulan kosong: OFF** — "
        "Hanya bulan yang memiliki order masuk yang dihitung dalam rata-rata."
    )

st.markdown("")

calculate = st.button(
    "Validasi & Hitung Benchmark",
    type="primary",
    disabled=selected_uploads == 0,
)

# ============================================================
# PROCESSING
# ============================================================

if calculate:

    monthly_results = []
    processed_info = []
    file_errors = []

    audit_rows = []
    file_identity = []


    with st.spinner(
        "Memvalidasi dan memproses file..."
    ):


        for marketplace, files in uploaded_groups.items():

            for uploaded_file in files:

                try:

                    expanded_files = iter_uploaded_data_files(
                        uploaded_file
                    )


                    for expanded in expanded_files:


                        display_name = expanded["display_name"]

                        raw_bytes = expanded["raw_bytes"]

                        stream = expanded["stream"]

                        filename = expanded["filename"]



                        try:


                            file_hash = hashlib.sha256(
                                raw_bytes
                            ).hexdigest()



                            # ====================================================
                            # READ FILE ROUTING
                            # Indonesia vs Thailand
                            # ====================================================

                            if is_thailand and marketplace == "SHO":

                                stream.seek(0)

                                try:

                                    df = pd.read_excel(
                                        stream,
                                        sheet_name="orders",
                                        header=0,
                                    )

                                except Exception:

                                    stream.seek(0)

                                    df = pd.read_excel(
                                        stream,
                                        header=0,
                                    )

                                detected_marketplace = "SHO_TH"
                                header_row = 0

                            elif is_thailand and marketplace == "TIK":

                                stream.seek(0)
                                df = parse_tiktok_xlsx(stream)
                                detected_marketplace = "TIK_TH"
                                header_row = 0

                            else:

                                stream.seek(0)
                                df, detected_marketplace, header_row = read_order_file(
                                    stream,
                                    filename=filename,
                                    marketplace_hint=(
                                        marketplace
                                        if marketplace in {
                                            "TIK",
                                            "TOK"
                                        }
                                        else None
                                    ),
                                )

                            # ====================================================
                            # PROCESS ENGINE ROUTING
                            # ====================================================

                            if is_thailand and marketplace == "SHO":

                                monthly = process_shopee_th(df, filename=filename)

                            elif is_thailand and marketplace == "TIK":

                                monthly = process_tiktok_th(df, filename=filename)

                            else:

                                if (
                                    marketplace in {
                                        "SHO",
                                        "LAZ"
                                    }
                                    and detected_marketplace != marketplace
                                ):
                                    raise BenchmarkInputError(
                                        f"File ini terdeteksi sebagai "
                                        f"{detected_marketplace}, bukan {marketplace}."
                                    )

                                monthly = process_dataframe(
                                    df,
                                    marketplace,
                                    filename=filename,
                                )

                            monthly_results.append(monthly)

                            # ====================================================
                            # AUDIT
                            # ====================================================

                            if is_thailand and marketplace == "SHO":

                                file_audit = [
                                    {
                                        "Status": "PASS",
                                        "Check": "Shopee Thailand Calculation",
                                        "Detail": "Processed using Thailand Shopee engine.",
                                        "Actual": float(
                                            monthly["benchmark_gmv"].sum()
                                        ),
                                        "Expected": "THB calculation",
                                    }
                                ]

                            elif is_thailand and marketplace == "TIK":

                                file_audit = [
                                    {
                                        "Status": "PASS",
                                        "Check": "TikTok Thailand Calculation",
                                        "Detail": "Processed using Thailand TikTok engine.",
                                        "Actual": float(
                                            monthly["benchmark_gmv"].sum()
                                        ),
                                        "Expected": "THB calculation",
                                    }
                                ]

                            else:

                                file_audit = validate_file_calculation(
                                    df,
                                    marketplace,
                                    monthly,
                                )



                            for item in file_audit:

                                audit_rows.append(
                                    {
                                        "Scope":
                                            "File",

                                        "Marketplace":
                                            marketplace,

                                        "File":
                                            display_name,

                                        **item,

                                    }
                                )



                            # ====================================================
                            # RECORD ID
                            # ====================================================

                            if df is None or df.empty:
                                ids = set()

                            elif is_thailand and marketplace == "SHO":
                                th_col = "หมายเลขคำสั่งซื้อ"
                                ids = (
                                    set(
                                        df[th_col]
                                        .dropna()
                                        .astype(str)
                                    )
                                    if th_col in df.columns
                                    else set()
                                )

                            else:

                                ids = extract_record_ids(
                                    df,
                                    marketplace,
                                )


                            file_identity.append(
                                {

                                    "marketplace":
                                        marketplace,

                                    "file":
                                        display_name,

                                    "sha256":
                                        file_hash,

                                    "ids":
                                        ids,

                                    "months":
                                        set(
                                            monthly[
                                                "month"
                                            ]
                                            .dropna()
                                            .tolist()
                                        ),

                                }
                            )



                            processed_info.append(
                                {

                                    "Status":
                                        "✅ OK",

                                    "Marketplace":
                                        marketplace,

                                    "File":
                                        display_name,

                                    "Rows":
                                        int(
                                            monthly["source_rows"].sum()
                                            if "source_rows" in monthly.columns
                                            else len(df)
                                        ),

                                    "Header Row":
                                        header_row + 1,

                                    "Periode Awal":
                                        (
                                            monthly[
                                                "month"
                                            ]
                                            .dropna()
                                            .min()
                                            .strftime("%b %Y")
                                            if not monthly["month"].dropna().empty
                                            else "-"
                                        ),


                                    "Periode Akhir":
                                        (
                                            monthly[
                                                "month"
                                            ]
                                            .dropna()
                                            .max()
                                            .strftime("%b %Y")
                                            if not monthly["month"].dropna().empty
                                            else "-"
                                        ),

                                }

                            )



                        except Exception as exc:


                            file_errors.append(
                                {

                                    "Marketplace":
                                        marketplace,

                                    "File":
                                        display_name,

                                    "Error":
                                        str(exc),

                                }
                            )


                            processed_info.append(
                                {

                                    "Status":
                                        "❌ ERROR",

                                    "Marketplace":
                                        marketplace,

                                    "File":
                                        display_name,

                                    "Rows":
                                        "-",

                                    "Header Row":
                                        "-",

                                    "Periode Awal":
                                        "-",

                                    "Periode Akhir":
                                        "-",

                                }
                            )



                except Exception as exc:


                    file_errors.append(
                        {

                            "Marketplace":
                                marketplace,

                            "File":
                                uploaded_file.name,

                            "Error":
                                str(exc),

                        }
                    )



    # ============================================================
    # VALIDATION RESULT
    # ============================================================

    st.subheader(
        "Validasi File"
    )


    st.dataframe(
        pd.DataFrame(processed_info),
        hide_index=True,
        width="stretch",
    )



    if file_errors:


        st.error(
            f"{len(file_errors)} file gagal divalidasi."
        )


        for error in file_errors:

            st.error(
                f"**{error['Marketplace']} — {error['File']}**\n\n"
                f"{error['Error']}"
            )


        st.stop()

# ============================================================
# CALCULATE RESULT
# ============================================================
    try:

        # ====================================================
        # THAILAND RESULT
        # ====================================================

        if is_thailand:

            # ====================================================
            # MONTHLY MARKETPLACE
            # ====================================================

            monthly_marketplace = (
                pd.concat(
                    monthly_results,
                    ignore_index=True,
                )
                .groupby(
                    ["marketplace", "month"],
                    as_index=False,
                )
                .agg(
                    gmv_all=("gmv_all", "sum"),
                    gmv_ex_cancel=("gmv_ex_cancel", "sum"),
                    platform_discount=("platform_discount", "sum"),
                    seller_rebate=("seller_rebate", "sum"),
                    benchmark_gmv=("benchmark_gmv", "sum"),
                    source_rows=("source_rows", "sum"),
                )
                .sort_values(
                    ["month", "marketplace"]
                )
                .reset_index(drop=True)
            )

            # ====================================================
            # PIVOT MARKETPLACE → SAME STRUCTURE AS INDONESIA
            # ====================================================

            pivot = (
                monthly_marketplace
                .pivot_table(
                    index="month",
                    columns="marketplace",
                    values="benchmark_gmv",
                    aggfunc="sum",
                    fill_value=0,
                )
                .reset_index()
            )

            if "SHO_TH" not in pivot.columns:
                pivot["SHO_TH"] = 0.0

            if "TIK_TH" not in pivot.columns:
                pivot["TIK_TH"] = 0.0

            monthly_combined = pivot[
                [
                    "month",
                    "SHO_TH",
                    "TIK_TH",
                ]
            ].copy()

            monthly_combined = (
                monthly_combined.rename(
                    columns={
                        "SHO_TH": "shopee",
                        "TIK_TH": "tiktok",
                    }
                )
            )

            # ====================================================
            # TOTAL BENCHMARK GMV
            # ====================================================

            monthly_combined["total_gmv"] = (
                monthly_combined[
                    [
                        "shopee",
                        "tiktok",
                    ]
                ]
                .sum(axis=1)
            )

            monthly_combined = (
                monthly_combined
                .sort_values("month")
                .reset_index(drop=True)
            )

            # ====================================================
            # BENCHMARK AVERAGE
            # ====================================================

            if include_empty_months:

                benchmark_series = (
                    monthly_combined["total_gmv"]
                )

            else:

                month_has_rows = (
                    monthly_marketplace
                    .groupby("month")["source_rows"]
                    .sum()
                )

                active_months = (
                    month_has_rows[
                        month_has_rows > 0
                    ]
                    .index
                )

                benchmark_series = (
                    monthly_combined[
                        monthly_combined["month"]
                        .isin(active_months)
                    ]["total_gmv"]
                )

            if benchmark_series.empty:

                benchmark = 0.0
                valid_months = 0

            else:

                benchmark = float(
                    benchmark_series.mean()
                )

                valid_months = int(
                    len(benchmark_series)
                )

            # ====================================================
            # RESULT
            # ====================================================

            result = {
                "benchmark": benchmark,
                "valid_months": valid_months,
                "monthly_combined": monthly_combined,
                "monthly_marketplace": monthly_marketplace,
                "warnings": [],
            }

        # ====================================================
        # INDONESIA RESULT
        # ====================================================

        else:

            result = calculate_benchmark(
                monthly_results,
                include_empty_months=include_empty_months,
            )



        # ====================================================
        # DISPLAY RESULT
        # ====================================================


        st.success(
            "Benchmark berhasil dihitung."
        )

        excel_file = create_excel_report_single_sheet(
            result,
            active_marketplaces,
            is_thailand,
        )


        st.download_button(
            label="📥 Download Benchmark Report (.xlsx)",
            data=excel_file,
            file_name=(
                "Benchmark_Report_TH.xlsx"
                if is_thailand
                else "Benchmark_Report_ID.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


        st.subheader(
            "Hasil Benchmark"
        )


        currency_symbol = (
            "฿"
            if is_thailand
            else "Rp"
        )


        m1, m2, m3, m4 = st.columns(4)



        m1.metric(
            "Benchmark",

            (
                f"{currency_symbol} "
                f"{result['benchmark']:,.0f}"
                .replace(",", ".")
            ),

            border=True,
        )


        m2.metric(
            "Jumlah Bulan",

            f"{result['valid_months']} bulan",

            border=True,
        )


        m3.metric(
            "Marketplace",

            f"{len(active_marketplaces)}",

            border=True,
        )


        m4.metric(
            "File Diproses",

            f"{len(processed_info)}",

            border=True,
        )



        # ====================================================
        # SUMMARY
        # ====================================================

        st.subheader(
            "Ringkasan Benchmark"
        )

        valid_months = result["valid_months"]
        monthly_combined = result["monthly_combined"]

        summary_rows = []

        if is_thailand:
            monthly_mp = result.get("monthly_marketplace", pd.DataFrame())
            for marketplace in active_marketplaces:
                if marketplace == "SHO":
                    label = "SHO_TH"
                    target_mp = "SHO_TH"
                elif marketplace == "TIK":
                    label = "TIK_TH"
                    target_mp = "TIK_TH"
                else:
                    label = marketplace
                    target_mp = marketplace

                if not monthly_mp.empty and "marketplace" in monthly_mp.columns:
                    val = (
                        monthly_mp[monthly_mp["marketplace"] == target_mp]["benchmark_gmv"].sum() / valid_months
                        if valid_months
                        else 0
                    )
                else:
                    val = result["benchmark"]

                summary_rows.append(
                    {
                        "Ringkasan Benchmark": label,
                        "Nilai": val,
                    }
                )
        else:
            marketplace_columns = {
                "SHO": "shopee",
                "TIK": "tiktok",
                "TOK": "tokopedia",
                "LAZ": "lazada",
            }

            for marketplace in active_marketplaces:
                if marketplace in marketplace_columns:
                    column = marketplace_columns[marketplace]
                    if column in monthly_combined.columns:
                        marketplace_avg = (
                            monthly_combined[column].sum() / valid_months
                            if valid_months
                            else 0
                        )
                        summary_rows.append(
                            {
                                "Ringkasan Benchmark": marketplace,
                                "Nilai": marketplace_avg,
                            }
                        )

        summary_rows.append(
            {
                "Ringkasan Benchmark": f"Benchmark Avg L{valid_months}M",
                "Nilai": result["benchmark"],
            }
        )

        summary_df = pd.DataFrame(summary_rows)

        if is_thailand:
            summary_df["Nilai"] = summary_df["Nilai"].apply(
                lambda x: f"฿{x:,.0f}"
            )
        else:
            summary_df["Nilai"] = summary_df["Nilai"].apply(
                lambda x: f"Rp {x:,.0f}"
            )

        st.dataframe(
            summary_df,
            hide_index=True,
            width="stretch",
        )



        # ====================================================
        # MONTHLY DETAIL
        # ====================================================
        if is_thailand:
            st.subheader("Benchmark per Bulan (THB ฿)")
        else:
            st.subheader("Benchmark per Bulan (IDR Rp)")


        monthly_display = (
            result["monthly_combined"]
            .copy()
        )


        monthly_display["month"] = (
            monthly_display["month"]
            .dt.strftime(
                "%B %Y"
            )
        )
        if is_thailand:

            for col in monthly_display.columns:
                if col != "month":
                    monthly_display[col] = monthly_display[col].apply(
                        lambda x: f"฿{x:,.0f}"
                    )

        else:

            for col in monthly_display.columns:
                if col != "month":
                    monthly_display[col] = monthly_display[col].apply(
                        lambda x: f"Rp {x:,.0f}"
                    )


        st.dataframe(

            monthly_display,

            hide_index=True,

            width="stretch",

        )



        # ====================================================
        # MARKETPLACE DETAIL
        # ====================================================


        if is_thailand:
            st.subheader(
                "Detail Perhitungan Marketplace (THB ฿)"
            )
        else:
            st.subheader(
                "Detail Perhitungan Marketplace (IDR Rp)"
            )


        if is_thailand:

            st.caption(
                """
                **Shopee Thailand calculation:**
                - **GMV All**: Total dari kolom **ราคาขายสุทธิ**
                - **GMV Exclude Cancel**: GMV dengan status pesanan bukan **ยกเลิกแล้ว** (Cancel)
                - **Seller Rebate**: Kolom **โค้ดส่วนลดชำระโดยผู้ขาย** (Order-level deduplicated)
                - **Benchmark**: GMV Exclude Cancel - Seller Rebate

                ---

                **TikTok Thailand calculation:**
                - **GMV All**: Total dari kolom **SKU Subtotal Before Discount**
                - **GMV Exclude Cancel**: GMV dari pesanan dengan status bukan **CANCELED** / **Cancel**
                - **Platform Discount**: Kolom **SKU Platform Discount** (Diskon/voucher disubsidi TikTok)
                - **Seller Rebate**: Kolom **SKU Seller Discount** (Diskon ditanggung Seller)
                - **Benchmark**: GMV Exclude Cancel - SKU Seller Discount (setara dengan SKU Subtotal After Discount + Platform Discount)

                **Currency:** THB (฿)
                """
            )


        detail = (
            result["monthly_marketplace"]
            .copy()
        )

        if is_thailand:
            detail["marketplace"] = (
                detail["marketplace"]
                .replace(
                    {
                        "SHO_TH": "SHO",
                "TIK_TH": "TIK",
            }
        )
    )
    


        detail["month"] = (
            detail["month"]
            .dt.strftime("%B %Y")
        )
        currency_columns = [
    "gmv_all",
    "gmv_ex_cancel",
    "platform_discount",
    "seller_rebate",
    "benchmark_gmv",
]
        for col in currency_columns:

            if col in detail.columns:

                if is_thailand:
                    detail[col] = detail[col].apply(
                        lambda x: f"฿{x:,.0f}"
                    )

                else:
                    detail[col] = detail[col].apply(
                        lambda x: f"Rp {x:,.0f}"
                    )


        st.dataframe(
            detail,
            hide_index=True,
                width="stretch",
        )
        render_seasonal_benchmark(
            result=result,
            currency_symbol=currency_symbol,
            include_empty_months=include_empty_months,
        )
        render_benchmark_adjustment(
            result=result,
            currency_symbol=currency_symbol,
            include_empty_months=include_empty_months,
        )


    except Exception as exc:


        st.error(
            f"Benchmark calculation error: {exc}"
        )
