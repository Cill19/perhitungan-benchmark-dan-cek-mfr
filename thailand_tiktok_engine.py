import re

import pandas as pd
from openpyxl import load_workbook


TIKTOK_TH_COLUMN_ALIASES = {
    "Order ID": ("Order ID",),
    "Order Status": ("Order Status",),
    "Order Sub Status": (
        "Order Sub Status",
        "Order Substatus",
        "Order Sub-Status",
    ),
    "Cancelation/Return Type": (
        "Cancelation/Return Type",
        "Cancellation/Return Type",
        "Cancel Type",
    ),
    "Created Time": ("Created Time", "Order Created Time"),
    "SKU Subtotal After Discount": ("SKU Subtotal After Discount",),
    "SKU Platform Discount": ("SKU Platform Discount",),
    "SKU Seller Discount": ("SKU Seller Discount",),
    "SKU Subtotal Before Discount": ("SKU Subtotal Before Discount",),
}

TIKTOK_TH_HEADER_SIGNATURE = {
    "Order ID",
    "Created Time",
    "SKU Subtotal After Discount",
    "SKU Platform Discount",
}


def _normalized_header(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().casefold()


TIKTOK_TH_ALIAS_LOOKUP = {
    _normalized_header(alias): canonical
    for canonical, aliases in TIKTOK_TH_COLUMN_ALIASES.items()
    for alias in aliases
}


def _header_score(row_values):
    matched_headers = {
        TIKTOK_TH_ALIAS_LOOKUP[normalized]
        for value in row_values.values()
        if (normalized := _normalized_header(value))
        in TIKTOK_TH_ALIAS_LOOKUP
    }
    return (
        len(matched_headers & TIKTOK_TH_HEADER_SIGNATURE),
        len(matched_headers),
    )


def _deduplicate_headers(headers):
    seen = {}
    unique_headers = []
    for index, value in enumerate(headers):
        base = str(value).strip() if value is not None else ""
        if not base:
            base = f"col_{index}"
        occurrence = seen.get(base, 0)
        unique_headers.append(
            base if occurrence == 0 else f"{base}.{occurrence}"
        )
        seen[base] = occurrence + 1
    return unique_headers


def parse_tiktok_xlsx(file_path):
    """
    Read a TikTok Shop XLSX export in openpyxl normal mode.

    Normal mode resolves shared strings, inline strings, numeric/boolean
    cells, workbook relationships, and TikTok files with stale worksheet
    dimensions. Every worksheet and the first 50 rows are scored so neither
    the worksheet nor the header row is assumed.
    """
    workbook = load_workbook(
        file_path,
        read_only=False,
        data_only=True,
    )
    try:
        best_match = None
        for sheet_order, worksheet in enumerate(workbook.worksheets):
            scan_limit = min(worksheet.max_row, 50)
            for row_number, row in enumerate(
                worksheet.iter_rows(
                    min_row=1,
                    max_row=scan_limit,
                    values_only=True,
                ),
                start=1,
            ):
                row_values = {
                    column_index: value
                    for column_index, value in enumerate(row)
                    if value is not None
                }
                signature_matches, total_matches = _header_score(
                    row_values
                )
                candidate = (
                    signature_matches,
                    total_matches,
                    -sheet_order,
                    -row_number,
                    worksheet.title,
                    row_number,
                )
                if best_match is None or candidate[:4] > best_match[:4]:
                    best_match = candidate

        if best_match is None:
            raise ValueError("TikTok TH workbook contains no readable rows")

        signature_matches, total_matches = best_match[:2]
        if signature_matches < 3 or total_matches < 4:
            raise ValueError(
                "TikTok TH header row not found in any worksheet. "
                f"Best match contained {signature_matches}/4 signature fields "
                f"and {total_matches} recognized TikTok fields."
            )

        sheet_name = best_match[4]
        header_row_idx = best_match[5]
        worksheet = workbook[sheet_name]
        header_values = next(
            worksheet.iter_rows(
                min_row=header_row_idx,
                max_row=header_row_idx,
                values_only=True,
            )
        )
        last_header_column = max(
            index
            for index, value in enumerate(header_values, start=1)
            if value is not None and str(value).strip()
        )
        headers = _deduplicate_headers(
            list(header_values[:last_header_column])
        )

        rows = [
            list(row)
            for row in worksheet.iter_rows(
                min_row=header_row_idx + 1,
                max_row=worksheet.max_row,
                max_col=last_header_column,
                values_only=True,
            )
            if any(value is not None for value in row)
        ]
        worksheet_path = worksheet.path.lstrip("/")
    finally:
        workbook.close()

    df = pd.DataFrame(rows, columns=headers)

    # TikTok exports place field descriptions immediately below the header.
    order_id_column = next(
        (
            column
            for column in df.columns
            if TIKTOK_TH_ALIAS_LOOKUP.get(_normalized_header(column))
            == "Order ID"
        ),
        None,
    )
    if (
        order_id_column is not None
        and not df.empty
        and _normalized_header(df.iloc[0][order_id_column])
        == _normalized_header("Platform unique order ID.")
    ):
        df = df.iloc[1:].reset_index(drop=True)

    df.attrs["sheet_name"] = sheet_name
    df.attrs["worksheet_path"] = worksheet_path
    df.attrs["header_row"] = header_row_idx
    df.attrs["reader"] = "openpyxl-normal"
    df.attrs["shared_strings_resolved"] = True

    return df


def parse_thb(series):
    return (
        series
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("฿", "", regex=False)
        .replace({"nan": "0", "None": "0", "-": "0", "": "0"})
        .astype(float)
    )


def is_cancelled_tiktok_th(row):
    status = str(row.get("status", "") or row.get("Order Status", "")).upper()
    cancel_type = str(row.get("cancel_type", "") or row.get("Cancelation/Return Type", "")).upper()

    if "CANCEL" in status or "BATAL" in status:
        return True
    if "CANCEL" in cancel_type:
        return True
    return False


def normalize_tiktok_th(df):
    required_columns = {
        "order_id": TIKTOK_TH_COLUMN_ALIASES["Order ID"],
        "status": TIKTOK_TH_COLUMN_ALIASES["Order Status"],
        "cancel_type": TIKTOK_TH_COLUMN_ALIASES["Cancelation/Return Type"],
        "created_time": TIKTOK_TH_COLUMN_ALIASES["Created Time"],
        "subtotal_after_discount": TIKTOK_TH_COLUMN_ALIASES[
            "SKU Subtotal After Discount"
        ],
        "platform_discount": TIKTOK_TH_COLUMN_ALIASES[
            "SKU Platform Discount"
        ],
        "seller_discount": TIKTOK_TH_COLUMN_ALIASES["SKU Seller Discount"],
        "subtotal_before_discount": TIKTOK_TH_COLUMN_ALIASES[
            "SKU Subtotal Before Discount"
        ],
    }

    available_columns = {
        _normalized_header(column): column
        for column in df.columns
    }
    resolved_columns = {}
    missing = []
    for internal_name, aliases in required_columns.items():
        source_column = next(
            (
                available_columns[_normalized_header(alias)]
                for alias in aliases
                if _normalized_header(alias) in available_columns
            ),
            None,
        )
        if source_column is None:
            missing.append(aliases[0])
        else:
            resolved_columns[source_column] = internal_name

    if missing:
        raise ValueError("TikTok TH missing columns: " + ", ".join(missing))

    data = df.rename(columns=resolved_columns).copy()

    return data


def process_tiktok_th(df, filename=None):
    # --------------------------------------------------------
    # Empty file guard
    # --------------------------------------------------------
    if df is None or df.empty:
        from benchmark_engine import _empty_result
        return _empty_result("TIK_TH", filename)

    data = normalize_tiktok_th(df)

    data["subtotal_after_discount"] = parse_thb(data["subtotal_after_discount"])
    data["platform_discount"] = parse_thb(data["platform_discount"])
    data["seller_discount"] = parse_thb(data["seller_discount"])
    data["subtotal_before_discount"] = parse_thb(data["subtotal_before_discount"])

    data["is_cancel"] = data.apply(is_cancelled_tiktok_th, axis=1)

    parsed_date = pd.to_datetime(data["created_time"], dayfirst=True, errors="coerce")

    data["month"] = (
        parsed_date
        .dt
        .to_period("M")
        .dt
        .to_timestamp()
    )

    data["gmv"] = data["subtotal_before_discount"]

    gmv_all = data.groupby("month")["gmv"].sum()

    valid = data[~data["is_cancel"]].copy()

    gmv_ex_cancel = valid.groupby("month")["gmv"].sum()

    platform_discount = valid.groupby("month")["platform_discount"].sum()

    seller_rebate = valid.groupby("month")["seller_discount"].sum()

    # Benchmark formula: GMV Exclude Cancel + Platform Discount - Seller Discount
    valid["benchmark_row"] = valid["subtotal_after_discount"] + valid["platform_discount"]
    benchmark_gmv = valid.groupby("month")["benchmark_row"].sum()

    result = pd.DataFrame(
        {
            "marketplace": "TIK_TH",
            "month": gmv_all.index,
            "gmv_all": gmv_all.values,
            "gmv_ex_cancel": gmv_ex_cancel.reindex(gmv_all.index, fill_value=0).values,
            "platform_discount": platform_discount.reindex(gmv_all.index, fill_value=0).values,
            "seller_rebate": seller_rebate.reindex(gmv_all.index, fill_value=0).values,
            "benchmark_gmv": benchmark_gmv.reindex(gmv_all.index, fill_value=0).values,
        }
    )

    result["source_rows"] = (
        result["month"]
        .map(data.groupby("month").size())
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
