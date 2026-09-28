import pandas as pd
import zipfile
import re
from xml.etree import ElementTree as ET


def col_to_index(col_str):
    num = 0
    for c in col_str:
        num = num * 26 + (ord(c.upper()) - ord('A') + 1)
    return num - 1


def parse_tiktok_xlsx(file_path):
    """
    Direct XML parser for TikTok Shop XLSX export files to bypass
    corrupted XML dimension ref (A1 ref bug in TikTok exports).
    """
    with zipfile.ZipFile(file_path) as zf:
        sheet_files = [
            f for f in zf.namelist()
            if f.startswith("xl/worksheets/sheet") and f.endswith(".xml")
        ]
        if not sheet_files:
            raise ValueError("No worksheet XML found in TikTok file")

        xml_bytes = zf.read(sheet_files[0])
        root = ET.fromstring(xml_bytes)
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        data_by_row = {}
        for c in root.findall(".//main:c", ns):
            ref = c.attrib.get("r")
            if not ref:
                continue
            m = re.match(r"([A-Z]+)(\d+)", ref)
            if not m:
                continue
            col_str, row_str = m.group(1), m.group(2)
            row_idx = int(row_str)
            col_idx = col_to_index(col_str)

            v_elem = c.find("main:v", ns)
            if v_elem is not None and v_elem.text is not None:
                val = v_elem.text
            else:
                is_elem = c.find(".//main:t", ns)
                val = is_elem.text if is_elem is not None else None

            if val is not None:
                if row_idx not in data_by_row:
                    data_by_row[row_idx] = {}
                data_by_row[row_idx][col_idx] = val

    if not data_by_row:
        return pd.DataFrame()

    sorted_row_indices = sorted(data_by_row.keys())
    header_row_idx = sorted_row_indices[0]
    header_dict = data_by_row[header_row_idx]
    max_col = max(max(r.keys()) for r in data_by_row.values())

    headers = [header_dict.get(c, f"col_{c}") for c in range(max_col + 1)]

    rows = []
    for r_idx in sorted_row_indices[1:]:
        row_data = data_by_row[r_idx]
        row_list = [row_data.get(c, None) for c in range(max_col + 1)]
        rows.append(row_list)

    df = pd.DataFrame(rows, columns=headers)

    # Drop description row if present (Row 0 of data where Order ID == 'Platform unique order ID.')
    if not df.empty and df.iloc[0].get("Order ID") == "Platform unique order ID.":
        df = df.iloc[1:].reset_index(drop=True)

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
        "order_id": "Order ID",
        "status": "Order Status",
        "cancel_type": "Cancelation/Return Type",
        "created_time": "Created Time",
        "subtotal_after_discount": "SKU Subtotal After Discount",
        "platform_discount": "SKU Platform Discount",
        "seller_discount": "SKU Seller Discount",
        "subtotal_before_discount": "SKU Subtotal Before Discount",
    }

    missing = [
        column for column in required_columns.values()
        if column not in df.columns
    ]

    if missing:
        raise ValueError("TikTok TH missing columns: " + ", ".join(missing))

    data = df.rename(
        columns={v: k for k, v in required_columns.items()}
    ).copy()

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
