import pandas as pd


LAZADA_TH_COLUMNS = {
    "order_item_id": "orderItemId",
    "order_id": "orderNumber",
    "seller_sku": "sellerSku",
    "created_time": "createTime",
    "status": "status",
    "paid_price": "paidPrice",
    "unit_price": "unitPrice",
    "seller_discount_raw": "sellerDiscountTotal",
    "platform_discount_raw": "platformDiscountTotal",
    "shipping_fee": "shippingFee",
}

LAZADA_TH_CANCEL_STATUSES = frozenset({"canceled", "cancelled"})
LAZADA_TH_DATE_FORMAT = "%d %b %Y %H:%M"
LAZADA_TH_TOLERANCE = 0.01


def parse_lazada_th(file_or_buffer, filename=None):
    """Read a Lazada Thailand export with the shared order-file reader."""
    from benchmark_engine import read_order_file

    if hasattr(file_or_buffer, "seek"):
        file_or_buffer.seek(0)

    data, detected_marketplace, header_row = read_order_file(
        file_or_buffer,
        filename=filename,
        marketplace_hint="LAZ",
    )
    if detected_marketplace != "LAZ":
        raise ValueError(
            "Lazada TH file was detected as "
            f"{detected_marketplace}, not LAZ."
        )

    data.attrs["header_row"] = header_row
    data.attrs["detected_marketplace"] = detected_marketplace
    return data


def normalize_lazada_th(df):
    """Map the exact Lazada Thailand export columns to internal names."""
    missing = [
        source_column
        for source_column in LAZADA_TH_COLUMNS.values()
        if source_column not in df.columns
    ]
    if missing:
        raise ValueError(
            "Lazada TH missing columns: " + ", ".join(missing)
        )

    return (
        df[list(LAZADA_TH_COLUMNS.values())]
        .rename(
            columns={
                source_column: internal_name
                for internal_name, source_column in LAZADA_TH_COLUMNS.items()
            }
        )
        .copy()
    )


def _nonempty(series):
    return (
        series.astype("string")
        .fillna("")
        .str.strip()
        .ne("")
    )


def _numeric_values(series, field_name, allow_empty=True):
    cleaned = (
        series.astype("string")
        .fillna("")
        .str.strip()
        .str.replace(",", "", regex=False)
        .str.replace("฿", "", regex=False)
        .str.replace(r"^-$", "", regex=True)
    )
    parsed = pd.to_numeric(cleaned, errors="coerce")
    invalid = cleaned.ne("") & parsed.isna()
    if invalid.any():
        samples = cleaned[invalid].drop_duplicates().head(3).tolist()
        raise ValueError(
            f"Lazada TH invalid numeric values in {field_name}: {samples}"
        )
    if not allow_empty and cleaned.eq("").any():
        raise ValueError(
            f"Lazada TH empty numeric values in {field_name}: "
            f"{int(cleaned.eq('').sum())} row(s)"
        )
    return parsed.fillna(0.0).astype(float)


def _created_time_values(series):
    parsed = pd.to_datetime(
        series,
        format=LAZADA_TH_DATE_FORMAT,
        errors="coerce",
    )
    invalid = _nonempty(series) & parsed.isna()
    if invalid.any():
        samples = (
            series.astype("string")[invalid]
            .drop_duplicates()
            .head(3)
            .tolist()
        )
        raise ValueError(
            "Lazada TH invalid createTime values: " + str(samples)
        )
    if parsed.isna().any():
        raise ValueError(
            "Lazada TH empty createTime values: "
            f"{int(parsed.isna().sum())} row(s)"
        )
    return parsed


def _normalized_status(series):
    return (
        series.astype("string")
        .fillna("")
        .str.strip()
        .str.casefold()
    )


def _prepare_lazada_th(df):
    data = normalize_lazada_th(df)

    data["order_item_id"] = (
        data["order_item_id"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )
    data["order_id"] = (
        data["order_id"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )
    data["status"] = _normalized_status(data["status"])
    data["created_time"] = _created_time_values(data["created_time"])

    for column in ["paid_price", "unit_price", "shipping_fee"]:
        data[column] = _numeric_values(
            data[column],
            LAZADA_TH_COLUMNS[column],
            allow_empty=False,
        )
    for column in ["seller_discount_raw", "platform_discount_raw"]:
        data[column] = _numeric_values(
            data[column],
            LAZADA_TH_COLUMNS[column],
            allow_empty=True,
        )

    if data["order_item_id"].eq("").any():
        raise ValueError("Lazada TH contains empty orderItemId values")
    duplicate_item_ids = data["order_item_id"].duplicated(keep=False)
    if duplicate_item_ids.any():
        samples = (
            data.loc[duplicate_item_ids, "order_item_id"]
            .drop_duplicates()
            .head(3)
            .tolist()
        )
        raise ValueError(
            "Lazada TH duplicate orderItemId values: " + str(samples)
        )
    if data["order_id"].eq("").any():
        raise ValueError("Lazada TH contains empty orderNumber values")
    if data["status"].eq("").any():
        raise ValueError("Lazada TH contains empty status values")

    cancel_like = data["status"].str.contains("cancel", regex=False)
    unrecognized_cancel = sorted(
        set(
            data.loc[
                cancel_like
                & ~data["status"].isin(LAZADA_TH_CANCEL_STATUSES),
                "status",
            ]
        )
    )
    if unrecognized_cancel:
        raise ValueError(
            "Lazada TH unrecognized cancel statuses: "
            + ", ".join(unrecognized_cancel)
        )

    for column in ["seller_discount_raw", "platform_discount_raw"]:
        positive = data[column] > LAZADA_TH_TOLERANCE
        if positive.any():
            raise ValueError(
                f"Lazada TH {LAZADA_TH_COLUMNS[column]} changed sign; "
                f"expected zero or negative values, found {int(positive.sum())} row(s)"
            )

    paid_reconciliation = (
        data["paid_price"]
        - (
            data["unit_price"]
            + data["seller_discount_raw"]
            + data["platform_discount_raw"]
            + data["shipping_fee"]
        )
    )
    bad_reconciliation = paid_reconciliation.abs() > LAZADA_TH_TOLERANCE
    if bad_reconciliation.any():
        raise ValueError(
            "Lazada TH price components no longer reconcile: "
            "paidPrice must equal unitPrice + sellerDiscountTotal + "
            "platformDiscountTotal + shippingFee. "
            f"Affected rows: {int(bad_reconciliation.sum())}"
        )

    data["month"] = (
        data["created_time"]
        .dt.to_period("M")
        .dt.to_timestamp()
    )
    data["is_cancel"] = data["status"].isin(
        LAZADA_TH_CANCEL_STATUSES
    )
    data["seller_rebate"] = -data["seller_discount_raw"]
    data["platform_discount"] = -data["platform_discount_raw"]
    return data


def process_lazada_th(df, filename=None):
    """
    Calculate Lazada Thailand benchmark at line-item level.

    paidPrice includes shipping and both signed discount fields:
      paidPrice = unitPrice + sellerDiscountTotal
                  + platformDiscountTotal + shippingFee

    Since unitPrice is already the gross value before platform subsidy,
    platform discount must not be added to it again. Therefore:
      benchmark = unitPrice non-cancel - seller rebate non-cancel
                = (paidPrice - shippingFee) + platform discount
    """
    if df is None or df.empty:
        from benchmark_engine import _empty_result

        return _empty_result("LAZ_TH", filename)

    data = _prepare_lazada_th(df)
    valid = ~data["is_cancel"]

    data["gmv_ex_cancel_row"] = data["unit_price"].where(valid, 0.0)
    data["platform_discount_row"] = data["platform_discount"].where(
        valid,
        0.0,
    )
    data["seller_rebate_row"] = data["seller_rebate"].where(
        valid,
        0.0,
    )
    data["benchmark_row"] = (
        data["gmv_ex_cancel_row"] - data["seller_rebate_row"]
    )

    result = (
        data.groupby("month", as_index=False)
        .agg(
            gmv_all=("unit_price", "sum"),
            gmv_ex_cancel=("gmv_ex_cancel_row", "sum"),
            platform_discount=("platform_discount_row", "sum"),
            seller_rebate=("seller_rebate_row", "sum"),
            benchmark_gmv=("benchmark_row", "sum"),
            source_rows=("order_item_id", "size"),
        )
        .sort_values("month")
        .reset_index(drop=True)
    )
    result.insert(0, "marketplace", "LAZ_TH")
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


def _audit(status, check, detail, actual, expected):
    return {
        "Status": status,
        "Check": check,
        "Detail": detail,
        "Actual": actual,
        "Expected": expected,
    }


def validate_lazada_th_calculation(df, monthly):
    """Independently reconcile Lazada Thailand source rows and output."""
    data = _prepare_lazada_th(df)
    tolerance = LAZADA_TH_TOLERANCE
    audit = []

    status_values = sorted(data["status"].unique().tolist())
    audit.append(
        _audit(
            "PASS",
            "Lazada TH required columns and statuses",
            "Required columns are present and cancel matching is controlled.",
            ", ".join(status_values),
            "canceled/cancelled excluded; other observed statuses included",
        )
    )

    paid_difference = (
        data["paid_price"]
        - data["unit_price"]
        - data["seller_discount_raw"]
        - data["platform_discount_raw"]
        - data["shipping_fee"]
    )
    max_paid_difference = float(paid_difference.abs().max())
    audit.append(
        _audit(
            "PASS" if max_paid_difference <= tolerance else "FAIL",
            "Lazada TH raw price reconciliation",
            "paidPrice = unitPrice + sellerDiscountTotal + "
            "platformDiscountTotal + shippingFee for every line item.",
            max_paid_difference,
            0.0,
        )
    )

    valid = ~data["is_cancel"]
    expected_rows = pd.DataFrame(
        {
            "month": data["month"],
            "gmv_all": data["unit_price"],
            "gmv_ex_cancel": data["unit_price"].where(valid, 0.0),
            "platform_discount": data["platform_discount"].where(
                valid,
                0.0,
            ),
            "seller_rebate": data["seller_rebate"].where(valid, 0.0),
            "source_rows": 1,
        }
    )
    expected_rows["benchmark_gmv"] = (
        expected_rows["gmv_ex_cancel"]
        - expected_rows["seller_rebate"]
    )
    expected = (
        expected_rows.groupby("month", as_index=False)
        .agg(
            gmv_all=("gmv_all", "sum"),
            gmv_ex_cancel=("gmv_ex_cancel", "sum"),
            platform_discount=("platform_discount", "sum"),
            seller_rebate=("seller_rebate", "sum"),
            benchmark_gmv=("benchmark_gmv", "sum"),
            source_rows=("source_rows", "sum"),
        )
    )
    actual = monthly[
        [
            "month",
            "gmv_all",
            "gmv_ex_cancel",
            "platform_discount",
            "seller_rebate",
            "benchmark_gmv",
            "source_rows",
        ]
    ].copy()
    reconciled = expected.merge(
        actual,
        on="month",
        how="outer",
        suffixes=("_expected", "_actual"),
    ).fillna(0.0)

    for column, label in [
        ("gmv_all", "GMV All"),
        ("gmv_ex_cancel", "GMV Exclude Cancel"),
        ("platform_discount", "Platform Discount"),
        ("seller_rebate", "Seller Rebate"),
        ("benchmark_gmv", "Benchmark GMV"),
        ("source_rows", "Source Rows"),
    ]:
        difference = (
            reconciled[f"{column}_actual"]
            - reconciled[f"{column}_expected"]
        ).abs()
        max_difference = float(difference.max()) if not difference.empty else 0.0
        audit.append(
            _audit(
                "PASS" if max_difference <= tolerance else "FAIL",
                f"Reconcile Lazada TH {label}",
                f"Maximum monthly difference for {column}.",
                max_difference,
                0.0,
            )
        )

    formula_difference = (
        monthly["benchmark_gmv"]
        - (
            monthly["gmv_ex_cancel"]
            - monthly["seller_rebate"]
        )
    ).abs()
    max_formula_difference = (
        float(formula_difference.max())
        if not formula_difference.empty
        else 0.0
    )
    audit.append(
        _audit(
            "PASS" if max_formula_difference <= tolerance else "FAIL",
            "Lazada TH benchmark formula",
            "benchmark_gmv = gmv_ex_cancel - seller_rebate each month; "
            "platform subsidy is already present in gross unitPrice.",
            max_formula_difference,
            0.0,
        )
    )
    return audit
