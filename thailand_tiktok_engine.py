import pandas as pd
import zipfile
import posixpath
import re
from xml.etree import ElementTree as ET


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


def _xml_local_name(tag):
    return tag.rsplit("}", 1)[-1]


def _normalized_header(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().casefold()


TIKTOK_TH_ALIAS_LOOKUP = {
    _normalized_header(alias): canonical
    for canonical, aliases in TIKTOK_TH_COLUMN_ALIASES.items()
    for alias in aliases
}


def col_to_index(col_str):
    num = 0
    for c in col_str:
        num = num * 26 + (ord(c.upper()) - ord('A') + 1)
    return num - 1


def _read_shared_strings(zf):
    """Decode the XLSX shared-string table, including rich-text runs."""
    path = "xl/sharedStrings.xml"
    if path not in zf.namelist():
        return []

    root = ET.fromstring(zf.read(path))
    shared_strings = []
    for item in root.iter():
        if _xml_local_name(item.tag) != "si":
            continue
        shared_strings.append(
            "".join(
                node.text or ""
                for node in item.iter()
                if _xml_local_name(node.tag) == "t"
            )
        )
    return shared_strings


def _worksheet_candidates(zf):
    """Return worksheets in workbook order, falling back to ZIP order."""
    names = set(zf.namelist())
    candidates = []

    workbook_path = "xl/workbook.xml"
    relationships_path = "xl/_rels/workbook.xml.rels"
    if workbook_path in names and relationships_path in names:
        workbook = ET.fromstring(zf.read(workbook_path))
        relationships = ET.fromstring(zf.read(relationships_path))
        relationship_targets = {
            relationship.attrib.get("Id"): relationship.attrib.get("Target", "")
            for relationship in relationships.iter()
            if _xml_local_name(relationship.tag) == "Relationship"
        }

        for sheet in workbook.iter():
            if _xml_local_name(sheet.tag) != "sheet":
                continue
            relationship_id = next(
                (
                    value
                    for key, value in sheet.attrib.items()
                    if _xml_local_name(key) == "id"
                ),
                None,
            )
            target = relationship_targets.get(relationship_id, "")
            if not target:
                continue
            if target.startswith("/"):
                worksheet_path = target.lstrip("/")
            else:
                worksheet_path = posixpath.normpath(
                    posixpath.join("xl", target)
                )
            if worksheet_path in names:
                candidates.append(
                    (sheet.attrib.get("name", worksheet_path), worksheet_path)
                )

    if candidates:
        return candidates

    sheet_files = sorted(
        name
        for name in names
        if name.startswith("xl/worksheets/") and name.endswith(".xml")
    )
    return [(path.rsplit("/", 1)[-1], path) for path in sheet_files]


def _cell_text(cell):
    return "".join(
        node.text or ""
        for node in cell.iter()
        if _xml_local_name(node.tag) == "t"
    )


def _decode_cell(cell, shared_strings):
    cell_type = cell.attrib.get("t", "")
    value_element = next(
        (
            child
            for child in cell
            if _xml_local_name(child.tag) == "v"
        ),
        None,
    )
    raw_value = (
        value_element.text
        if value_element is not None and value_element.text is not None
        else None
    )

    if cell_type == "inlineStr":
        return _cell_text(cell)
    if cell_type == "s":
        if raw_value is None:
            return None
        try:
            shared_index = int(raw_value)
            return shared_strings[shared_index]
        except (ValueError, IndexError) as exc:
            raise ValueError(
                f"Invalid XLSX shared-string index: {raw_value!r}"
            ) from exc
    if cell_type == "b":
        return raw_value == "1"
    if cell_type in {"str", "e"}:
        return raw_value
    if raw_value is None:
        inline_value = _cell_text(cell)
        return inline_value if inline_value != "" else None

    # Untyped/numeric cells are values, not shared-string indexes.
    try:
        if re.fullmatch(r"[-+]?\d+", raw_value):
            return int(raw_value)
        return float(raw_value)
    except ValueError:
        return raw_value


def _read_worksheet_rows(xml_bytes, shared_strings):
    """Read actual cells and ignore unreliable worksheet dimension metadata."""
    root = ET.fromstring(xml_bytes)
    data_by_row = {}

    for row in root.iter():
        if _xml_local_name(row.tag) != "row":
            continue
        row_number = row.attrib.get("r")
        row_values = {}
        next_column = 0
        for cell in row:
            if _xml_local_name(cell.tag) != "c":
                continue
            reference = cell.attrib.get("r", "")
            match = re.match(r"([A-Za-z]+)(\d+)", reference)
            if match:
                column_index = col_to_index(match.group(1))
                if row_number is None:
                    row_number = match.group(2)
            else:
                column_index = next_column
            next_column = column_index + 1
            value = _decode_cell(cell, shared_strings)
            if value is not None:
                row_values[column_index] = value

        if row_values:
            if row_number is None:
                row_number = str(len(data_by_row) + 1)
            # Some malformed TikTok exports repeat the same <row r="...">
            # element once per cell. Merge those fragments instead of keeping
            # only the final cell for that row.
            data_by_row.setdefault(int(row_number), {}).update(row_values)

    return data_by_row


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
    Direct XML parser for TikTok Shop XLSX export files to bypass
    corrupted XML dimension ref (A1 ref bug in TikTok exports).
    """
    with zipfile.ZipFile(file_path) as zf:
        shared_strings = _read_shared_strings(zf)
        worksheets = _worksheet_candidates(zf)
        if not worksheets:
            raise ValueError("No worksheet XML found in TikTok file")

        best_match = None
        for sheet_order, (sheet_name, worksheet_path) in enumerate(worksheets):
            data_by_row = _read_worksheet_rows(
                zf.read(worksheet_path),
                shared_strings,
            )
            for row_number in sorted(data_by_row)[:50]:
                signature_matches, total_matches = _header_score(
                    data_by_row[row_number]
                )
                candidate = (
                    signature_matches,
                    total_matches,
                    -sheet_order,
                    -row_number,
                    sheet_name,
                    worksheet_path,
                    row_number,
                    data_by_row,
                )
                if best_match is None or candidate[:4] > best_match[:4]:
                    best_match = candidate

    if best_match is None:
        return pd.DataFrame()

    signature_matches, total_matches = best_match[:2]
    if signature_matches < 3 or total_matches < 4:
        raise ValueError(
            "TikTok TH header row not found in any worksheet. "
            f"Best match contained {signature_matches}/4 signature fields "
            f"and {total_matches} recognized TikTok fields."
        )

    sheet_name = best_match[4]
    worksheet_path = best_match[5]
    header_row_idx = best_match[6]
    data_by_row = best_match[7]
    sorted_row_indices = sorted(data_by_row)
    header_dict = data_by_row[header_row_idx]
    max_col = max(max(r.keys()) for r in data_by_row.values())

    headers = _deduplicate_headers(
        [header_dict.get(c, f"col_{c}") for c in range(max_col + 1)]
    )

    rows = []
    for r_idx in sorted_row_indices:
        if r_idx <= header_row_idx:
            continue
        row_data = data_by_row[r_idx]
        row_list = [row_data.get(c, None) for c in range(max_col + 1)]
        rows.append(row_list)

    df = pd.DataFrame(rows, columns=headers)

    # Drop description row if present (Row 0 of data where Order ID == 'Platform unique order ID.')
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
    df.attrs["shared_strings_resolved"] = bool(shared_strings)

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
