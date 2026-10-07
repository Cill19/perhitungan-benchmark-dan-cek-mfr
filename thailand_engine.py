import re
from pathlib import Path
import pandas as pd


# ============================================================
# THAILAND HELPERS
# ============================================================

def parse_thb(series):
    """
    Thailand currency parser.

    Example:
    98.00 -> 98
    1,250.50 -> 1250.50
    """

    return (
        series
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("฿", "", regex=False)
        .replace(
            {
                "nan": "0",
                "None": "0",
                "-": "0",
                "": "0",
            }
        )
        .astype(float)
    )


def parse_th_datetime(series):
    """Parse ISO timestamps and Thailand day-first dates without ambiguity."""
    text = series.astype("string").fillna("").str.strip()
    iso_mask = text.str.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}")
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    parsed.loc[iso_mask] = pd.to_datetime(
        text.loc[iso_mask],
        format="mixed",
        yearfirst=True,
        errors="coerce",
    )
    parsed.loc[~iso_mask] = pd.to_datetime(
        text.loc[~iso_mask],
        format="mixed",
        dayfirst=True,
        errors="coerce",
    )
    return parsed


def is_cancelled_th(status):
    """
    Detect Thailand cancelled order.
    """

    text = str(status).lower()

    keywords = [
        "ยกเลิก",
        "cancel",
        "cancelled",
        "canceled",
    ]

    return any(
        keyword in text
        for keyword in keywords
    )


# ============================================================
# SHOPEE TH NORMALIZER
# ============================================================

def normalize_shopee_th(df):

    required_columns = {

        "order_id":
            "หมายเลขคำสั่งซื้อ",

        "status":
            "สถานะการสั่งซื้อ",

        "created_time":
            "วันที่ทำการสั่งซื้อ",

        "gmv":
            "ราคาขายสุทธิ",

        "seller_rebate":
            "โค้ดส่วนลดชำระโดยผู้ขาย",

    }


    missing = [
        column
        for column in required_columns.values()
        if column not in df.columns
    ]


    if missing:
        raise ValueError(
            "Shopee TH missing columns: "
            + ", ".join(missing)
        )


    data = df.rename(
        columns={
            value: key
            for key, value in required_columns.items()
        }
    )


    return data[
        [
            "order_id",
            "status",
            "created_time",
            "gmv",
            "seller_rebate",
        ]
    ].copy()



# ============================================================
# SHOPEE TH BENCHMARK ENGINE
# ============================================================

def process_shopee_th(df, filename=None):

    # --------------------------------------------------------
    # Empty file guard
    # --------------------------------------------------------
    if df is None or df.empty:
        from benchmark_engine import _empty_result
        return _empty_result("SHO_TH", filename)

    data = normalize_shopee_th(df)


    # ========================================================
    # Numeric Conversion
    # ========================================================

    data["gmv"] = parse_thb(
        data["gmv"]
    )


    data["seller_rebate"] = parse_thb(
        data["seller_rebate"]
    )


    # ========================================================
    # Cancel Detection
    # ========================================================

    data["is_cancel"] = (
        data["status"]
        .apply(is_cancelled_th)
    )


    # ========================================================
    # Month Conversion
    # ========================================================

    parsed_date = parse_th_datetime(
        data["created_time"]
    )


    data["month"] = (
        parsed_date
        .dt
        .to_period("M")
        .dt
        .to_timestamp()
    )


    # ========================================================
    # GMV ALL
    # ========================================================

    gmv_all = (
        data
        .groupby("month")["gmv"]
        .sum()
    )


    # ========================================================
    # EXCLUDE CANCEL
    # ========================================================

    valid = data[
        ~data["is_cancel"]
    ].copy()


    gmv_ex_cancel = (
        valid
        .groupby("month")["gmv"]
        .sum()
    )


    # ========================================================
    # SELLER REBATE
    #
    # Voucher is order-level.
    # Prevent duplicate calculation
    # when order contains multiple SKU.
    # ========================================================

    seller_rebate = (
        valid
        .groupby(
            [
                "order_id",
                "month",
            ]
        )["seller_rebate"]
        .max()
        .groupby("month")
        .sum()
    )


    # ========================================================
    # RESULT
    # ========================================================

    result = pd.DataFrame(
        {
            "marketplace":
                "SHO_TH",

            "month":
                gmv_all.index,

            "gmv_all":
                gmv_all.values,

            "gmv_ex_cancel":
                gmv_ex_cancel
                .reindex(
                    gmv_all.index,
                    fill_value=0
                )
                .values,

            "platform_discount":
                0,

            "seller_rebate":
                seller_rebate
                .reindex(
                    gmv_all.index,
                    fill_value=0
                )
                .values,
        }
    )


    # Benchmark formula:
    #
    # GMV Exclude Cancel
    # -
    # Seller Rebate

    result["benchmark_gmv"] = (
        result["gmv_ex_cancel"]
        -
        result["seller_rebate"]
    )


    result["source_rows"] = (
        result["month"]
        .map(
            data
            .groupby("month")
            .size()
        )
        .fillna(0)
        .astype(int)
    )


    return result[
        [
            "marketplace",
            "month",
            "gmv_all",
            "gmv_ex_cancel",
            "platform_discount",
            "seller_rebate",
            "benchmark_gmv",
            "source_rows",
        ]
    ]
