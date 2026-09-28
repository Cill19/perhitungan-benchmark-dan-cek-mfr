"""
Benchmark Brand Superstar - Processing Engine v0.13.0

Business logic:
- SHO: GMV Exclude Batal - Voucher Ditanggung Penjual + Diskon Dari Shopee
- TIK: Order Sub Status Selesai → SKU Subtotal After Discount + SKU Platform Discount
- TOK: sama dengan TikTok (Order Sub Status Selesai → SKU Subtotal After Discount + SKU Platform Discount)
- LAZ: tetap (unitPrice Exclude Batal + platformDiscountTotal)
- Final: jumlah benchmark marketplace per calendar month lalu average seluruh bulan eligible
- tidak ada maksimum L6M.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, BinaryIO, Iterable, Mapping
import io
import re

import numpy as np
import pandas as pd


ENGINE_VERSION = "0.13.0"


class BenchmarkInputError(ValueError):
    """Error input yang aman ditampilkan ke user."""


@dataclass
class FileInspection:
    filename: str
    marketplace: str
    header_row: int
    row_count: int
    min_month: pd.Timestamp | None
    max_month: pd.Timestamp | None
    warnings: list[str]


SIGNATURES = {
    "SHO": {
        "No. Pesanan",
        "Status Pesanan",
        "Waktu Pesanan Dibuat",
        "Subtotal Pesanan",
    },
    "TIKTOK_FAMILY": {
        "Order ID",
        "Order Sub Status",
        "SKU Subtotal After Discount",
        "Created Time",
    },
    "LAZ": {
        "orderItemId",
        "createTime",
        "unitPrice",
        "platformDiscountTotal",
        "status",
    },
}


def _rewind(source: Any) -> None:
    if hasattr(source, "seek"):
        source.seek(0)


def _clean_header(value: Any) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _header_score(values: Iterable[Any]) -> tuple[float, str]:
    cols = {_clean_header(v) for v in values if _clean_header(v)}
    best_score = 0.0
    best_template = ""

    for template, required in SIGNATURES.items():
        score = len(required & cols) / len(required)
        if score > best_score:
            best_score = score
            best_template = template

    return best_score, best_template


def _find_header_in_dataframe(
    preview: pd.DataFrame,
) -> tuple[int, str, float]:
    best: tuple[float, int, str] | None = None

    for idx, row in preview.iterrows():
        score, template = _header_score(row.tolist())
        candidate = (score, int(idx), template)
        if best is None or candidate[0] > best[0]:
            best = candidate

    if best is None:
        return 0, "", 0.0

    return best[1], best[2], best[0]


def _resolve_tik_tok(
    filename: str,
    marketplace_hint: str | None,
) -> str:
    if marketplace_hint:
        hint = marketplace_hint.strip().upper()
        if hint in {"TIK", "TIKTOK"}:
            return "TIK"
        if hint in {"TOK", "TOKOPEDIA"}:
            return "TOK"

    name = filename.lower()

    if "tiktok" in name or re.search(r"(^|[_\-\s])tik([_\-\s.]|$)", name):
        return "TIK"

    if "tokopedia" in name or re.search(r"(^|[_\-\s])tok([_\-\s.]|$)", name):
        return "TOK"

    raise BenchmarkInputError(
        f"File '{filename}' memiliki format TikTok/Tokopedia yang sama. "
        "Upload file pada slot TikTok atau Tokopedia yang sesuai."
    )



HEADER_ALIASES = {
    # Shopee
    "No. Pesanan": {
        "No. Pesanan",
        "No Pesanan",
        "Nomor Pesanan",
        "Nomor Order",
    },
    "Status Pesanan": {
        "Status Pesanan",
        "Status Pesanan ",
        "Status Pesanan*",
        "Status Order",
        "Order Status Shopee",
    },
    "Waktu Pesanan Dibuat": {
        "Waktu Pesanan Dibuat",
        "Waktu Pesanan Dibuat ",
        "Tanggal Pesanan Dibuat",
        "Waktu Order Dibuat",
    },
    "Subtotal Pesanan": {
        "Subtotal Pesanan",
        "Subtotal Pesanan ",
        "Subtotal Order",
    },
    "Voucher Ditanggung Penjual": {
        "Voucher Ditanggung Penjual",
        "Voucher Ditanggung Penjual ",
        "Voucher Penjual",
    },
    "Paket Diskon (Diskon dari Penjual)": {
        "Paket Diskon (Diskon dari Penjual)",
        "Paket Diskon - Diskon dari Penjual",
        "Paket Diskon Diskon dari Penjual",
        "Diskon Paket dari Penjual",
    },

    "Diskon Dari Shopee": {
        "Diskon Dari Shopee",
        "Diskon Shopee",
    },

    # TikTok / Tokopedia
    "Order ID": {
        "Order ID",
        "OrderID",
        "order_id",
    },

    "Order Status": {
        "Order Status",
        "order_status",
    },

    "Order Sub Status": {
        "Order Sub Status",
        "Order Substatus",
        "Order Sub-Status",
        "order_sub_status",
    },

    "Cancelation/Return Type": {
        "Cancelation/Return Type",
        "Cancellation/Return Type",
        "Cancellation / Return Type",
        "Cancelation / Return Type",
        "Cancelation Return Type",
        "Cancellation Return Type",
    },
    "Created Time": {
        "Created Time",
        "Order Created Time",
        "created_time",
    },
    "SKU Subtotal After Discount": {
        "SKU Subtotal After Discount",
        "SKU Subtotal after Discount",
        "SKU Subtotal",
    },
    "SKU Platform Discount": {
        "SKU Platform Discount",
        "Platform Discount",
        "SKU Platform discount",
    },

    # Lazada
    "orderItemId": {
        "orderItemId",
        "orderItemID",
        "Order Item ID",
    },
    "createTime": {
        "createTime",
        "Create Time",
    },
    "unitPrice": {
        "unitPrice",
        "Unit Price",
        "unit price",
    },
    "platformDiscountTotal": {
        "platformDiscountTotal",
        "Platform Discount Total",
    },
    "status": {
        "status",
        "Status",
    },
}


def _header_key(value: Any) -> str:
    """
    Header comparison key:
      'Status Pesanan*' -> 'statuspesanan'
      'No. Pesanan' -> 'nopesanan'
    """
    value = _clean_header(value).lower()
    return re.sub(r"[^a-z0-9]+", "", value)


_ALIAS_LOOKUP: dict[str, str] = {}

for _canonical, _aliases in HEADER_ALIASES.items():
    _ALIAS_LOOKUP[_header_key(_canonical)] = _canonical
    for _alias in _aliases:
        _ALIAS_LOOKUP[_header_key(_alias)] = _canonical


def _normalize_header_name(value: Any) -> str:
    cleaned = _clean_header(value)
    if not cleaned:
        return ""

    return _ALIAS_LOOKUP.get(
        _header_key(cleaned),
        cleaned,
    )


def _normalize_dataframe_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    normalized = [
        _normalize_header_name(c)
        for c in df.columns
    ]

    # Avoid duplicate canonical column names after normalization.
    counts: dict[str, int] = {}
    final_cols: list[str] = []

    for col in normalized:
        if col not in counts:
            counts[col] = 1
            final_cols.append(col)
        else:
            counts[col] += 1
            final_cols.append(
                f"{col}__dup{counts[col]}"
            )

    df.columns = final_cols
    return df


def _resolve_column(
    df: pd.DataFrame,
    canonical: str,
    extra_candidates: Iterable[str] = (),
) -> str | None:
    """
    Resolve a logical field even when marketplace export changes punctuation
    or slightly changes its label.
    """
    candidates = [
        canonical,
        *HEADER_ALIASES.get(canonical, set()),
        *extra_candidates,
    ]

    by_key = {
        _header_key(col): col
        for col in df.columns
    }

    for candidate in candidates:
        key = _header_key(candidate)
        if key in by_key:
            return by_key[key]

    return None


def _require_columns(
    df: pd.DataFrame,
    marketplace: str,
    fields: Mapping[str, Iterable[str] | None],
) -> dict[str, str]:
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for canonical, extras in fields.items():
        col = _resolve_column(
            df,
            canonical,
            extras or (),
        )

        if col is None:
            missing.append(canonical)
        else:
            resolved[canonical] = col

    if missing:
        visible_cols = ", ".join(
            str(c)
            for c in list(df.columns)[:25]
        )

        raise BenchmarkInputError(
            f"{marketplace}: kolom wajib tidak ditemukan: "
            + ", ".join(missing)
            + f". Kolom yang terbaca: {visible_cols}"
        )

    return resolved


def _expand_single_column_table(
    preview: pd.DataFrame,
) -> tuple[pd.DataFrame, str | None]:
    """
    Beberapa file .xlsx ternyata menyimpan satu baris CSV/TSV utuh
    di setiap cell kolom A. Jika begitu, expand kembali menjadi tabel.
    """
    if preview.shape[1] != 1 or preview.empty:
        return preview, None

    values = (
        preview.iloc[:, 0]
        .astype("string")
        .fillna("")
        .tolist()
    )

    separators = ["\t", ";", "|", ","]

    best_df = preview
    best_sep = None
    best_width = 1

    for sep in separators:
        split_rows = [
            str(v).split(sep)
            for v in values
        ]
        width = max(
            (len(r) for r in split_rows),
            default=1,
        )

        if width <= best_width:
            continue

        normalized_rows = [
            r + [None] * (width - len(r))
            for r in split_rows
        ]

        candidate = pd.DataFrame(
            normalized_rows
        )

        _, _, score = _find_header_in_dataframe(
            candidate
        )

        # Prefer widest parse, but require at least some header evidence.
        if score > 0 and width > best_width:
            best_df = candidate
            best_sep = sep
            best_width = width

    return best_df, best_sep



def _xlsx_col_index(cell_ref: str) -> int:
    """Convert Excel cell ref like BM12 -> zero-based column index."""
    match = re.match(r"([A-Z]+)", cell_ref.upper())
    if not match:
        return 0

    result = 0
    for ch in match.group(1):
        result = result * 26 + (ord(ch) - 64)

    return result - 1


def _read_xlsx_xml_direct(
    source: Any,
    sheet_name: str,
) -> pd.DataFrame:
    """
    Read XLSX worksheet XML while bypassing incorrect worksheet dimensions.

    Important difference from v0.7:
    we use openpyxl's own package parser to resolve:
      - the real worksheet XML target
      - the real shared-string table

    Some marketplace-generated workbooks contain stale/orphan OOXML parts.
    Hardcoding xl/sharedStrings.xml can therefore decode the correct cell
    indexes with the wrong string table (for example 'Creator Handle'
    instead of 'Order ID').
    """
    from openpyxl.reader.excel import ExcelReader
    from xml.etree import ElementTree as ET

    raw = _source_to_bytes(source)
    stream = io.BytesIO(raw)

    try:
        reader = ExcelReader(
            stream,
            read_only=True,
            data_only=True,
            keep_links=False,
        )
        reader.read_manifest()
        reader.read_strings()
        reader.read_workbook()
    except Exception as exc:
        raise BenchmarkInputError(
            f"OOXML reader gagal membuka workbook: {exc}"
        ) from exc

    target = None

    for sheet, rel in reader.parser.find_sheets():
        if sheet.name == sheet_name:
            target = rel.target
            break

    if not target:
        raise BenchmarkInputError(
            f"Worksheet XML untuk '{sheet_name}' tidak ditemukan."
        )

    try:
        xml_bytes = reader.archive.read(target)
    except Exception as exc:
        raise BenchmarkInputError(
            f"Worksheet XML '{target}' tidak dapat dibaca: {exc}"
        ) from exc

    root = ET.fromstring(xml_bytes)

    # Support transitional and strict SpreadsheetML namespaces.
    if root.tag.startswith("{"):
        main_ns = root.tag.split("}")[0].strip("{")
    else:
        main_ns = (
            "http://schemas.openxmlformats.org/"
            "spreadsheetml/2006/main"
        )

    ns = {"m": main_ns}

    shared_strings = list(
        getattr(reader, "shared_strings", []) or []
    )

    parsed_rows: dict[int, dict[int, Any]] = {}
    max_col = -1
    max_row = -1

    for row_el in root.findall(
        ".//m:sheetData/m:row",
        ns,
    ):
        row_num = int(
            row_el.attrib.get("r", "1")
        ) - 1

        max_row = max(
            max_row,
            row_num,
        )

        row_values: dict[int, Any] = {}

        for cell in row_el.findall(
            "m:c",
            ns,
        ):
            ref = cell.attrib.get(
                "r",
                "A1",
            )

            col_idx = _xlsx_col_index(
                ref
            )

            max_col = max(
                max_col,
                col_idx,
            )

            cell_type = cell.attrib.get(
                "t"
            )

            value = None

            if cell_type == "inlineStr":
                texts = [
                    t.text or ""
                    for t in cell.findall(
                        ".//m:is/m:t",
                        ns,
                    )
                ]
                value = "".join(texts)

            else:
                v = cell.find(
                    "m:v",
                    ns,
                )

                if v is not None:
                    raw_value = v.text or ""

                    if cell_type == "s":
                        try:
                            value = shared_strings[
                                int(raw_value)
                            ]
                        except Exception:
                            value = raw_value

                    elif cell_type == "b":
                        value = (
                            raw_value == "1"
                        )

                    else:
                        value = raw_value

            row_values[
                col_idx
            ] = value

        parsed_rows[
            row_num
        ] = row_values

    if max_row < 0 or max_col < 0:
        return pd.DataFrame()

    rows = []

    for row_idx in range(
        max_row + 1
    ):
        row = [
            None
        ] * (
            max_col + 1
        )

        for (
            col_idx,
            value,
        ) in parsed_rows.get(
            row_idx,
            {},
        ).items():
            row[
                col_idx
            ] = value

        rows.append(
            row
        )

    try:
        reader.archive.close()
    except Exception:
        pass

    return pd.DataFrame(
        rows
    )


def _read_excel_openpyxl(
    source: Any,
    sheet_name: str,
    nrows: int | None = None,
) -> pd.DataFrame:
    """
    Read worksheet using openpyxl NORMAL mode (read_only=False).

    Why:
    Some TikTok/Tokopedia marketplace exports contain a broken worksheet
    dimension such as `A1`, even though the worksheet actually contains
    dozens of columns (e.g. A:BM).

    openpyxl read_only=True trusts that bad dimension and may return only A1.
    Normal mode parses the actual cell records and correctly reconstructs
    all populated rows/columns, matching what Excel itself displays.
    """
    from openpyxl import load_workbook

    raw = _source_to_bytes(source)

    try:
        wb = load_workbook(
            io.BytesIO(raw),
            read_only=False,
            data_only=True,
        )
    except Exception as exc:
        raise BenchmarkInputError(
            f"Workbook '{sheet_name}' gagal dibaca dengan openpyxl normal mode: {exc}"
        ) from exc

    if sheet_name not in wb.sheetnames:
        raise BenchmarkInputError(
            f"Worksheet '{sheet_name}' tidak ditemukan."
        )

    ws = wb[sheet_name]

    max_row = ws.max_row
    max_col = ws.max_column

    if nrows is not None:
        max_row = min(
            max_row,
            nrows,
        )

    rows = []

    for row in ws.iter_rows(
        min_row=1,
        max_row=max_row,
        min_col=1,
        max_col=max_col,
        values_only=True,
    ):
        rows.append(
            list(row)
        )

    if not rows:
        return pd.DataFrame()

    # Remove trailing fully-empty columns.
    width = max(
        (
            max(
                (
                    i + 1
                    for i, value in enumerate(row)
                    if value is not None
                ),
                default=0,
            )
            for row in rows
        ),
        default=0,
    )

    if width == 0:
        return pd.DataFrame()

    rows = [
        row[:width]
        for row in rows
    ]

    return pd.DataFrame(
        rows
    )


def _read_excel_preview_openpyxl(
    source: Any,
    sheet_name: str,
    nrows: int = 40,
) -> pd.DataFrame:
    return _read_excel_openpyxl(
        source,
        sheet_name,
        nrows=nrows,
    )


def _scan_excel(
    source: Any,
    filename: str,
) -> tuple[str, int, str, str | None, str]:
    """
    Scan all worksheets with multiple readers:
    - pandas
    - openpyxl reset_dimensions
    - direct XLSX XML

    Returns:
      sheet_name, header_row, template, embedded_separator, reader_name
    """
    _rewind(source)

    try:
        excel = pd.ExcelFile(source)
    except Exception as exc:
        raise BenchmarkInputError(
            f"File '{filename}' tidak dapat dibuka sebagai Excel: {exc}"
        ) from exc

    best = None
    debug = []

    ext = Path(filename).suffix.lower()

    for sheet_name in excel.sheet_names:
        previews = []

        # Reader 1: pandas
        try:
            _rewind(source)
            pd_preview = pd.read_excel(
                source,
                sheet_name=sheet_name,
                header=None,
                nrows=40,
                dtype=object,
            )
            previews.append(
                ("pandas", pd_preview)
            )
        except Exception:
            pass

        # Reader 2: openpyxl + reset_dimensions
        try:
            ox_preview = _read_excel_preview_openpyxl(
                source,
                sheet_name,
                nrows=40,
            )
            previews.append(
                ("openpyxl-normal", ox_preview)
            )
        except Exception:
            pass

        # Reader 3: direct raw XLSX XML
        if ext in {".xlsx", ".xlsm"}:
            try:
                xml_full = _read_xlsx_xml_direct(
                    source,
                    sheet_name,
                )
                xml_preview = xml_full.head(40)
                previews.append(
                    ("xlsx-xml-direct", xml_preview)
                )
            except Exception as exc:
                debug.append(
                    f"{sheet_name}/xlsx-xml-direct ERROR={exc}"
                )

        for reader_name, preview in previews:
            if preview is None or preview.empty:
                continue

            expanded, embedded_sep = _expand_single_column_table(
                preview
            )

            # Normalize candidate header values before scoring.
            normalized_preview = expanded.copy()
            normalized_preview = normalized_preview.map(
                _normalize_header_name
            )

            header_row, template, score = _find_header_in_dataframe(
                normalized_preview
            )

            width = int(
                expanded.shape[1]
            )

            candidate = (
                score,
                width,
                sheet_name,
                header_row,
                template,
                embedded_sep,
                reader_name,
            )

            if best is None or candidate[:2] > best[:2]:
                best = candidate

            preview_text = []

            for idx, row in expanded.head(4).iterrows():
                vals = [
                    _clean_header(v)
                    for v in row.tolist()
                    if _clean_header(v)
                ][:12]

                if vals:
                    preview_text.append(
                        f"row {idx + 1}: {vals}"
                    )

            debug.append(
                f"{sheet_name}/{reader_name} "
                f"[{width} kolom, score={score:.2f}] "
                + " | ".join(preview_text)
            )

    if best is None:
        raise BenchmarkInputError(
            f"File '{filename}' tidak memiliki worksheet yang dapat dibaca."
        )

    (
        score,
        width,
        sheet_name,
        header_row,
        template,
        embedded_sep,
        reader_name,
    ) = best

    if score < 0.6:
        raise BenchmarkInputError(
            f"Header marketplace belum dikenali pada file '{filename}'. "
            "Diagnostic: "
            + " || ".join(debug[:10])
        )

    return (
        str(sheet_name),
        int(header_row),
        template,
        embedded_sep,
        reader_name,
    )


def _source_to_bytes(source: Any) -> bytes:
    if isinstance(source, (str, Path)):
        return Path(source).read_bytes()

    _rewind(source)
    data = source.read()
    _rewind(source)

    if isinstance(data, str):
        return data.encode("utf-8")

    return bytes(data)


def _scan_text_file(
    source: Any,
    filename: str,
) -> tuple[str, str, int, str]:
    raw = _source_to_bytes(source)

    encodings = [
        "utf-8-sig",
        "utf-8",
        "utf-16",
        "utf-16-le",
        "utf-16-be",
        "cp1252",
        "latin1",
    ]
    separators = [",", ";", "\t", "|"]

    best = None

    for encoding in encodings:
        try:
            text = raw.decode(encoding)
        except Exception:
            continue

        for sep in separators:
            try:
                preview = pd.read_csv(
                    io.StringIO(text),
                    sep=sep,
                    header=None,
                    nrows=40,
                    dtype=object,
                    engine="python",
                    on_bad_lines="skip",
                )
            except Exception:
                continue

            header_row, template, score = _find_header_in_dataframe(
                preview
            )

            width = int(
                preview.shape[1]
            )

            candidate = (
                score,
                width,
                encoding,
                sep,
                header_row,
                template,
            )

            if best is None or candidate[:2] > best[:2]:
                best = candidate

    if best is None:
        raise BenchmarkInputError(
            f"File '{filename}' tidak dapat dibaca sebagai CSV/TXT."
        )

    (
        score,
        width,
        encoding,
        sep,
        header_row,
        template,
    ) = best

    if score < 0.6:
        raise BenchmarkInputError(
            f"Header marketplace belum dikenali pada file '{filename}'. "
            f"Parser terbaik menemukan {width} kolom."
        )

    return (
        encoding,
        sep,
        int(header_row),
        template,
    )


def _drop_tiktok_description_row(
    df: pd.DataFrame,
) -> pd.DataFrame:
    if df.empty or "Order ID" not in df.columns:
        return df

    drop_count = 0

    for i in range(min(5, len(df))):
        row = df.iloc[i]

        order_id = str(
            row.get("Order ID", "")
        ).strip()

        order_status = str(
            row.get("Order Status", "")
        ).strip().lower()

        created_time = str(
            row.get("Created Time", "")
        ).strip().lower()

        looks_like_description = (
            "platform unique order id" in order_id.lower()
            or "current order status" in order_status
            or "order created time" in created_time
        )

        normalized_order_id = re.sub(
            r"\.0$",
            "",
            order_id,
        )

        looks_like_real_order = bool(
            re.fullmatch(
                r"\d{8,}",
                normalized_order_id,
            )
        )

        if looks_like_description:
            drop_count += 1
            continue

        if not looks_like_real_order:
            nonempty = sum(
                1
                for value in row.tolist()
                if str(value).strip().lower()
                not in {"", "nan", "none"}
            )

            if nonempty <= 3:
                drop_count += 1
                continue

        break

    if drop_count:
        df = (
            df.iloc[drop_count:]
            .reset_index(drop=True)
        )

    return df


def read_order_file(
    source: str | Path | BinaryIO,
    *,
    filename: str | None = None,
    marketplace_hint: str | None = None,
) -> tuple[pd.DataFrame, str, int]:
    if isinstance(source, (str, Path)):
        resolved_filename = (
            filename
            or Path(source).name
        )
        source_obj: Any = str(source)
    else:
        resolved_filename = (
            filename
            or getattr(
                source,
                "name",
                "uploaded_file.xlsx",
            )
        )
        source_obj = source

    ext = Path(
        resolved_filename
    ).suffix.lower()

    if ext in {
        ".xlsx",
        ".xls",
        ".xlsm",
    }:
        (
            sheet_name,
            header_row,
            template,
            embedded_sep,
            excel_reader,
        ) = _scan_excel(
            source_obj,
            resolved_filename,
        )

        if ext in {".xlsx", ".xlsm"}:
            if excel_reader == "xlsx-xml-direct":
                raw_table = _read_xlsx_xml_direct(
                    source_obj,
                    sheet_name,
                )
            else:
                raw_table = _read_excel_openpyxl(
                    source_obj,
                    sheet_name,
                    nrows=None,
                )

            if raw_table.empty:
                raise BenchmarkInputError(
                    f"Worksheet '{sheet_name}' pada file "
                    f"'{resolved_filename}' kosong."
                )

            expanded, detected_sep = _expand_single_column_table(
                raw_table
            )

            if embedded_sep is not None and detected_sep is None:
                raise BenchmarkInputError(
                    f"File '{resolved_filename}' tidak dapat diexpand menjadi tabel."
                )

            header_values = (
                expanded.iloc[header_row]
                .tolist()
            )

            data = expanded.iloc[
                header_row + 1:
            ].copy()

            data.columns = header_values
            df = data.reset_index(
                drop=True
            )

        else:
            # Legacy .xls fallback.
            _rewind(source_obj)
            df = pd.read_excel(
                source_obj,
                sheet_name=sheet_name,
                header=header_row,
                dtype=object,
            )

    elif ext in {
        ".csv",
        ".txt",
    }:
        (
            encoding,
            sep,
            header_row,
            template,
        ) = _scan_text_file(
            source_obj,
            resolved_filename,
        )

        raw = _source_to_bytes(
            source_obj
        )

        text = raw.decode(
            encoding
        )

        df = pd.read_csv(
            io.StringIO(text),
            sep=sep,
            header=header_row,
            dtype=object,
            engine="python",
            on_bad_lines="skip",
        )

    else:
        raise BenchmarkInputError(
            f"Format file '{resolved_filename}' belum didukung. "
            "Gunakan .xlsx, .xls, .xlsm, atau .csv."
        )

    df = _normalize_dataframe_headers(
        df
    )

    keep_cols = [
        c
        for c in df.columns
        if c
        and not str(c).lower().startswith(
            "unnamed:"
        )
    ]

    df = df.loc[
        :,
        keep_cols,
    ]

    df = (
        df.dropna(how="all")
        .reset_index(drop=True)
    )

    score, actual_template = _header_score(
        df.columns
    )

    if score < 0.6:
        raise BenchmarkInputError(
            f"File '{resolved_filename}' berhasil dibaca, "
            "tetapi kolom benchmark wajib belum lengkap. "
            f"Kolom terbaca: {list(df.columns)[:20]}"
        )

    if actual_template == "TIKTOK_FAMILY":
        marketplace = _resolve_tik_tok(
            resolved_filename,
            marketplace_hint,
        )

        df = _drop_tiktok_description_row(
            df
        )
    else:
        marketplace = actual_template

    return (
        df,
        marketplace,
        header_row,
    )


def _validate_columns(
    df: pd.DataFrame,
    required: list[str],
    marketplace: str,
) -> None:
    missing = [col for col in required if col not in df.columns]

    if missing:
        raise BenchmarkInputError(
            f"{marketplace}: kolom wajib tidak ditemukan: "
            + ", ".join(missing)
        )


def _to_idr_number(series: pd.Series) -> pd.Series:
    """
    Normalize nominal IDR.
    Examples:
      225.000 -> 225000
      225,000 -> 225000
      Rp 225.000 -> 225000
    """
    text = series.astype("string").fillna("").str.strip()

    cleaned = text.str.replace(
        r"[^0-9\-]",
        "",
        regex=True,
    )

    return pd.to_numeric(
        cleaned.replace({"": "0", "-": "0"}),
        errors="coerce",
    ).fillna(0.0)


def _parse_datetime(
    series: pd.Series,
    *,
    dayfirst: bool = False,
) -> pd.Series:
    parsed = pd.to_datetime(
        series,
        errors="coerce",
        dayfirst=dayfirst,
    )

    # Excel serial fallback
    numeric = pd.to_numeric(series, errors="coerce")
    mask = parsed.isna() & numeric.notna()

    if mask.any():
        parsed.loc[mask] = pd.to_datetime(
            numeric.loc[mask],
            unit="D",
            origin="1899-12-30",
            errors="coerce",
        )

    return parsed


def _month_start(
    series: pd.Series,
    *,
    dayfirst: bool = False,
) -> pd.Series:
    dt = _parse_datetime(
        series,
        dayfirst=dayfirst,
    )
    return dt.dt.to_period("M").dt.to_timestamp()


def process_shopee(df: pd.DataFrame) -> pd.DataFrame:
    """
    SHO benchmark per month:

      Benchmark GMV
      = GMV Exclude Batal
      - Voucher Ditanggung Penjual
      + Diskon Dari Shopee

    GMV base:
      Subtotal Pesanan

    Tidak dihitung:
      - Paket Diskon (Diskon dari Penjual)
        karena sudah tercermin pada harga coret / Subtotal Pesanan.
      - Voucher Ditanggung Shopee
        karena subsidi platform tidak mengurangi omzet produk.
    """

    required = [
        "No. Pesanan",
        "Status Pesanan",
        "Waktu Pesanan Dibuat",
        "Subtotal Pesanan",
        "Voucher Ditanggung Penjual",
        "Diskon Dari Shopee",
    ]
    _validate_columns(df, required, "SHO")

    month = _month_start(
        df["Waktu Pesanan Dibuat"],
        dayfirst=False,
    )

    status = (
        df["Status Pesanan"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    not_cancelled = ~status.str.contains(
        r"BATAL|CANCEL",
        regex=True,
        na=False,
    )

    order_id = (
        df["No. Pesanan"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    subtotal = _to_idr_number(
        df["Subtotal Pesanan"]
    )

    voucher_seller = _to_idr_number(
        df["Voucher Ditanggung Penjual"]
    )

    shopee_discount = _to_idr_number(
        df["Diskon Dari Shopee"]
    )

    # ---------------------------------------------------------
    # GMV
    # ---------------------------------------------------------

    base = pd.DataFrame({
        "month": month,
        "gmv_raw": subtotal,
        "not_cancelled": not_cancelled,
    })

    base = base[
        base["month"].notna()
    ].copy()

    base["gmv_ex_cancel"] = np.where(
        base["not_cancelled"],
        base["gmv_raw"],
        0.0,
    )

    # ---------------------------------------------------------
    # VOUCHER SELLER
    # Voucher Ditanggung Penjual berada di level Order ID.
    # Hindari double count jika satu order memiliki beberapa SKU.
    # ---------------------------------------------------------

    seller_rows = pd.DataFrame({
        "month": month,
        "order_id": order_id,
        "voucher_seller": voucher_seller,
        "not_cancelled": not_cancelled,
    })

    seller_rows = seller_rows[
        seller_rows["month"].notna()
        & seller_rows["not_cancelled"]
        & seller_rows["order_id"].ne("")
    ].copy()

    monthly_seller = (
        seller_rows
        .groupby(
            ["month", "order_id"],
            as_index=False,
        )
        .agg(
            seller_rebate=("voucher_seller", "max")
        )
        .groupby(
            "month",
            as_index=False,
        )
        .agg(
            seller_rebate=("seller_rebate", "sum")
        )
    )

    # ---------------------------------------------------------
    # DISKON SHOPEE
    # Diskon Dari Shopee berada di level item/SKU,
    # sehingga dijumlahkan langsung dari seluruh row non-batal.
    # ---------------------------------------------------------

    platform_rows = pd.DataFrame({
        "month": month,
        "shopee_discount": shopee_discount,
        "not_cancelled": not_cancelled,
    })

    platform_rows = platform_rows[
        platform_rows["month"].notna()
        & platform_rows["not_cancelled"]
    ].copy()

    monthly_platform = (
        platform_rows
        .groupby(
            "month",
            as_index=False,
        )
        .agg(
            platform_discount=(
                "shopee_discount",
                "sum",
            )
        )
    )

    # ---------------------------------------------------------
    # MONTHLY GMV
    # ---------------------------------------------------------

    monthly_gmv = (
        base
        .groupby(
            "month",
            as_index=False,
        )
        .agg(
            gmv_all=("gmv_raw", "sum"),
            gmv_ex_cancel=("gmv_ex_cancel", "sum"),
            source_rows=("gmv_raw", "size"),
        )
    )

    # ---------------------------------------------------------
    # MERGE
    # ---------------------------------------------------------

    out = (
        monthly_gmv
        .merge(
            monthly_seller,
            on="month",
            how="left",
        )
        .merge(
            monthly_platform,
            on="month",
            how="left",
        )
    )

    out["seller_rebate"] = (
        out["seller_rebate"]
        .fillna(0.0)
    )

    out["platform_discount"] = (
        out["platform_discount"]
        .fillna(0.0)
    )

    # FIXED SHOPEE BENCHMARK LOGIC
    out["benchmark_gmv"] = (
        out["gmv_ex_cancel"]
        - out["seller_rebate"]
        + out["platform_discount"]
    )

    out["marketplace"] = "SHO"

    return out[
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


def process_tik_family(
    df: pd.DataFrame,
    marketplace: str,
) -> pd.DataFrame:
    """
    TikTok / Tokopedia benchmark.

    Order valid:
      Order Sub Status = SELESAI
      Order ID tidak kosong

    Benchmark:
      SKU Subtotal After Discount + SKU Platform Discount
    """
    if marketplace not in {"TIK", "TOK"}:
        raise BenchmarkInputError(
            "Marketplace TikTok-family harus TIK atau TOK."
        )

    cols = _require_columns(
        df,
        marketplace,
        {
            "Order ID": (),
            "Order Sub Status": (),
            "SKU Subtotal After Discount": (),
            "SKU Platform Discount": (),
            "Created Time": (),
        },
    )

    order_id = (
        df[cols["Order ID"]]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True,
        )
    )

    sub_status = (
        df[cols["Order Sub Status"]]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    subtotal = _to_idr_number(
        df[cols["SKU Subtotal After Discount"]]
    )

    platform_discount = _to_idr_number(
        df[cols["SKU Platform Discount"]]
    )

    created_time = _parse_datetime(
        df[cols["Created Time"]],
        dayfirst=True,
    )

    raw = pd.DataFrame(
        {
            "order_id": order_id,
            "sub_status": sub_status,
            "subtotal": subtotal,
            "platform_discount": platform_discount,
            "created_time": created_time,
        }
    )

    # Order valid hanya berdasarkan status akhir:
    # Order Sub Status = SELESAI dan Order ID tidak kosong
    valid_orders_raw = raw[
        raw["order_id"].ne("")
        & raw["sub_status"].eq("SELESAI")
    ].copy()

    if valid_orders_raw.empty:
        return pd.DataFrame(
            columns=[
                "marketplace",
                "month",
                "gmv_all",
                "gmv_ex_cancel",
                "platform_discount",
                "seller_rebate",
                "benchmark_gmv",
                "source_rows",
            ]
        )

    grouped = (
        valid_orders_raw
        .groupby(
            "order_id",
            as_index=False,
            sort=False,
        )
        .agg(
            subtotal=("subtotal", "sum"),
            platform_discount=("platform_discount", "sum"),
            created_time=("created_time", "first"),
            source_rows=("order_id", "size"),
        )
    )

    grouped["month"] = (
        grouped["created_time"]
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    grouped = grouped[
        grouped["month"].notna()
    ].copy()

    grouped["benchmark_gmv"] = (
        grouped["subtotal"]
        + grouped["platform_discount"]
    )

    out = (
        grouped
        .groupby(
            "month",
            as_index=False,
        )
        .agg(
            gmv_all=("subtotal", "sum"),
            gmv_ex_cancel=("subtotal", "sum"),
            platform_discount=("platform_discount", "sum"),
            benchmark_gmv=("benchmark_gmv", "sum"),
            source_rows=("source_rows", "sum"),
        )
    )

    out["seller_rebate"] = 0.0
    out["marketplace"] = marketplace

    return out[
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


def process_lazada(df: pd.DataFrame) -> pd.DataFrame:
    """
    Lazada benchmark:
      GMV Benchmark = unitPrice exclude canceled/cancelled
                      + platformDiscountTotal exclude canceled/cancelled

    `unitPrice` is the GMV base. `paidPrice` is intentionally not used.
    """
    cols = _require_columns(
        df,
        "LAZ",
        {
            "createTime": (),
            "unitPrice": (),
            "platformDiscountTotal": (),
            "status": (),
        },
    )

    month = _month_start(
        df[cols["createTime"]],
        dayfirst=True,
    )

    status = (
        df[cols["status"]]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.lower()
    )

    not_cancelled = ~status.isin(
        {"canceled", "cancelled"}
    )

    unit_price = _to_idr_number(
        df[cols["unitPrice"]]
    )

    platform_discount = _to_idr_number(
        df[cols["platformDiscountTotal"]]
    )

    work = pd.DataFrame(
        {
            "month": month,
            "gmv_raw": unit_price,
            "platform_discount_raw": platform_discount,
            "not_cancelled": not_cancelled,
        }
    )

    work = work[
        work["month"].notna()
    ].copy()

    work["gmv_ex_cancel"] = np.where(
        work["not_cancelled"],
        work["gmv_raw"],
        0.0,
    )

    work["platform_discount"] = np.where(
        work["not_cancelled"],
        work["platform_discount_raw"],
        0.0,
    )

    work["benchmark_gmv"] = (
        work["gmv_ex_cancel"]
        + work["platform_discount"]
    )

    out = (
        work.groupby(
            "month",
            as_index=False,
        )
        .agg(
            gmv_all=("gmv_raw", "sum"),
            gmv_ex_cancel=("gmv_ex_cancel", "sum"),
            platform_discount=("platform_discount", "sum"),
            benchmark_gmv=("benchmark_gmv", "sum"),
            source_rows=("gmv_raw", "size"),
        )
    )

    out["seller_rebate"] = 0.0
    out["marketplace"] = "LAZ"

    return out[
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


# ============================================================
# EMPTY FILE HELPERS
# ============================================================

# Month abbreviations for filename parsing (English + Indonesian)
_MONTH_MAP: dict[str, int] = {
    "jan": 1, "january": 1, "januari": 1,
    "feb": 2, "february": 2, "februari": 2,
    "mar": 3, "march": 3, "maret": 3,
    "apr": 4, "april": 4,
    "may": 5, "mei": 5,
    "jun": 6, "june": 6, "juni": 6,
    "jul": 7, "july": 7, "juli": 7,
    "aug": 8, "august": 8, "agustus": 8,
    "sep": 9, "september": 9,
    "oct": 10, "october": 10, "oktober": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12, "desember": 12,
}


def _parse_period_from_filename(
    filename: str,
) -> pd.Timestamp | None:
    """
    Attempt to extract a month/year period from a filename.

    Patterns tried (in order):
      1. YYYYMMDD_YYYYMMDD  →  month of first date
         e.g. "Order.all.20260101_20260131.xlsx"
      2. Word month + 4-digit year
         e.g. "Data Pesanan Jan 2026 SHO TH.xlsx"
      3. MM-YYYY or MM_YYYY
         e.g. "pesanan_01_2026.xlsx"
    """
    name = Path(filename).stem  # strip extension

    # Pattern 1: YYYYMMDD_YYYYMMDD
    m = re.search(
        r"(\d{4})(\d{2})(\d{2})[_\-](\d{4})(\d{2})(\d{2})",
        name,
    )
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        try:
            return pd.Timestamp(year=year, month=month, day=1)
        except Exception:
            pass

    # Pattern 2: month word + 4-digit year
    m = re.search(
        r"([a-zA-Z]+)[\s\-_]+([0-9]{4})",
        name,
    )
    if m:
        word = m.group(1).lower()
        year = int(m.group(2))
        month_num = _MONTH_MAP.get(word)
        if month_num and 2000 <= year <= 2099:
            try:
                return pd.Timestamp(year=year, month=month_num, day=1)
            except Exception:
                pass

    # Pattern 3: MM-YYYY or MM_YYYY
    m = re.search(
        r"\b(0?[1-9]|1[0-2])[_\-](20[0-9]{2})\b",
        name,
    )
    if m:
        month, year = int(m.group(1)), int(m.group(2))
        try:
            return pd.Timestamp(year=year, month=month, day=1)
        except Exception:
            pass

    return None


_EMPTY_COLUMNS = [
    "marketplace",
    "month",
    "gmv_all",
    "gmv_ex_cancel",
    "platform_discount",
    "seller_rebate",
    "benchmark_gmv",
    "source_rows",
]


def _empty_result(
    marketplace: str,
    filename: str | None = None,
) -> pd.DataFrame:
    """
    Return a single-row DataFrame with all monetary values = 0.

    The month is extracted from filename if provided.
    If the period cannot be determined, month is NaT (row still included
    so that average calculation counts this month as 0).
    """
    period = None
    if filename:
        period = _parse_period_from_filename(filename)

    row = {
        "marketplace": marketplace.upper(),
        "month": period,  # may be NaT if not parseable
        "gmv_all": 0.0,
        "gmv_ex_cancel": 0.0,
        "platform_discount": 0.0,
        "seller_rebate": 0.0,
        "benchmark_gmv": 0.0,
        "source_rows": 0,
    }

    return pd.DataFrame([row])[_EMPTY_COLUMNS]


def process_dataframe(
    df: pd.DataFrame,
    marketplace: str,
    filename: str | None = None,
) -> pd.DataFrame:
    marketplace = marketplace.upper()

    # --------------------------------------------------------
    # Empty file guard — return a zeroed row so that this
    # month is counted in the average (value = 0).
    # --------------------------------------------------------
    if df is None or df.empty:
        return _empty_result(marketplace, filename)

    if marketplace == "SHO":
        return process_shopee(df)

    if marketplace in {"TIK", "TOK"}:
        return process_tik_family(
            df,
            marketplace,
        )

    if marketplace == "LAZ":
        return process_lazada(df)

    raise BenchmarkInputError(
        f"Marketplace belum didukung: {marketplace}"
    )



def _audit_row(
    status: str,
    check: str,
    detail: str,
    *,
    actual: Any = "",
    expected: Any = "",
) -> dict[str, Any]:
    return {
        "Status": status,
        "Check": check,
        "Detail": detail,
        "Actual": actual,
        "Expected": expected,
    }


def _nonempty_invalid_numeric_count(series: pd.Series) -> tuple[int, list[str]]:
    """Find non-empty monetary values that would silently collapse to 0."""
    text = series.astype("string").fillna("").str.strip()
    allowed_zero = text.str.lower().isin({"", "-", "0", "0.0", "0,0", "nan", "none"})
    cleaned = text.str.replace(r"[^0-9\-]", "", regex=True)
    invalid = (~allowed_zero) & cleaned.isin({"", "-"})
    samples = text[invalid].drop_duplicates().head(5).tolist()
    return int(invalid.sum()), samples


def _normalized_id_series(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .fillna("")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


def extract_record_ids(df: pd.DataFrame, marketplace: str) -> set[str]:
    """Stable identifiers used only for cross-file duplicate detection."""
    marketplace = marketplace.upper()

    if marketplace == "SHO":
        col = _resolve_column(df, "No. Pesanan")
    elif marketplace in {"TIK", "TOK"}:
        col = _resolve_column(df, "Order ID")
    elif marketplace == "LAZ":
        col = _resolve_column(df, "orderItemId")
    else:
        col = None

    if col is None:
        return set()

    ids = _normalized_id_series(df[col])
    return set(ids[ids.ne("")].tolist())


def validate_file_calculation(
    df: pd.DataFrame,
    marketplace: str,
    monthly: pd.DataFrame,
) -> list[dict[str, Any]]:
    """
    Independent audit checks for one uploaded file.

    FAIL = benchmark must not be treated as final.
    WARNING = calculation can continue, but user should review the anomaly.
    PASS = reconciliation succeeded exactly (within floating tolerance).
    """
    marketplace = marketplace.upper()
    audit: list[dict[str, Any]] = []
    tol = 0.5  # rupiah-level rounding tolerance

    def add(status: str, check: str, detail: str, actual: Any = "", expected: Any = ""):
        audit.append(_audit_row(status, check, detail, actual=actual, expected=expected))

    # ---------------------------------------------------------
    # Common monthly arithmetic
    # ---------------------------------------------------------
    if monthly is None or monthly.empty:
        add("FAIL", "Monthly output tersedia", "Tidak ada baris bulanan yang berhasil dihitung.")
        return audit

    if monthly["month"].isna().any():
        add("FAIL", "Bulan hasil valid", "Output monthly masih mengandung bulan kosong/invalid.")
    else:
        add("PASS", "Bulan hasil valid", "Semua baris output memiliki bulan yang valid.")

    if marketplace == "SHO":
        expected_bench = (
            monthly["gmv_ex_cancel"]
            - monthly["seller_rebate"]
            + monthly["platform_discount"]
        )
    else:
        expected_bench = (
            monthly["gmv_ex_cancel"]
            + monthly["platform_discount"]
        )

    diff = (monthly["benchmark_gmv"] - expected_bench).abs().max()
    diff = float(diff) if pd.notna(diff) else 0.0
    add(
        "PASS" if diff <= tol else "FAIL",
        "Formula GMV Benchmark",
        "GMV Benchmark direkonsiliasi ulang dari komponen detail.",
        actual=f"max selisih Rp {diff:,.0f}",
        expected="Rp 0",
    )

    # ---------------------------------------------------------
    # SHOPEE
    # ---------------------------------------------------------
    if marketplace == "SHO":
        cols = _require_columns(
            df,
            "SHO",
            {
                "No. Pesanan": (),
                "Status Pesanan": ("Order Status", "Status"),
                "Waktu Pesanan Dibuat": (),
                "Subtotal Pesanan": (),
                "Voucher Ditanggung Penjual": (),
                "Diskon Dari Shopee": (),
            },
        )

        ids = _normalized_id_series(df[cols["No. Pesanan"]])
        status = df[cols["Status Pesanan"]].astype("string").fillna("").str.strip().str.lower()
        created = _parse_datetime(df[cols["Waktu Pesanan Dibuat"]], dayfirst=False)

        nonempty_date = df[cols["Waktu Pesanan Dibuat"]].astype("string").fillna("").str.strip().ne("")
        invalid_dates = int((nonempty_date & created.isna()).sum())
        add(
            "PASS" if invalid_dates == 0 else "FAIL",
            "Parsing tanggal Shopee",
            "Semua Waktu Pesanan Dibuat harus dapat diterjemahkan menjadi tanggal.",
            actual=invalid_dates,
            expected=0,
        )

        for field in [
            "Subtotal Pesanan",
            "Voucher Ditanggung Penjual",
            "Diskon Dari Shopee",
        ]:
            count, samples = _nonempty_invalid_numeric_count(df[cols[field]])
            add(
                "PASS" if count == 0 else "FAIL",
                f"Parsing nominal {field}",
                "Nominal non-kosong tidak boleh diam-diam berubah menjadi 0 karena format tidak dikenali."
                + (f" Contoh: {samples}" if samples else ""),
                actual=count,
                expected=0,
            )

        exact_cancel = {"batal", "dibatalkan", "cancelled", "canceled", "cancel"}
        looks_cancel = status.str.contains(r"batal|cancel", regex=True, na=False)
        missed_cancel = sorted(set(status[looks_cancel & ~status.isin(exact_cancel)].tolist()))
        add(
            "PASS" if not missed_cancel else "FAIL",
            "Status batal Shopee",
            "Semua status yang terlihat seperti batal harus masuk rule cancel."
            + (f" Status berisiko: {missed_cancel[:10]}" if missed_cancel else ""),
            actual=len(missed_cancel),
            expected=0,
        )

        order_meta = pd.DataFrame({
            "id": ids,
            "status": status,
            "month": created.dt.to_period("M").astype("string"),
        })
        order_meta = order_meta[order_meta["id"].ne("")]
        status_conflicts = int((order_meta.groupby("id")["status"].nunique(dropna=True) > 1).sum())
        month_conflicts = int((order_meta.groupby("id")["month"].nunique(dropna=True) > 1).sum())
        add(
            "PASS" if status_conflicts == 0 else "FAIL",
            "Konsistensi status per No. Pesanan",
            "Satu No. Pesanan tidak boleh memiliki status berbeda antar-baris SKU.",
            actual=status_conflicts,
            expected=0,
        )
        add(
            "PASS" if month_conflicts == 0 else "FAIL",
            "Konsistensi bulan per No. Pesanan",
            "Satu No. Pesanan tidak boleh tersebar ke bulan Created Time yang berbeda.",
            actual=month_conflicts,
            expected=0,
        )

        voucher = _to_idr_number(df[cols["Voucher Ditanggung Penjual"]])
        shopee_discount = _to_idr_number(df[cols["Diskon Dari Shopee"]])
        subtotal = _to_idr_number(df[cols["Subtotal Pesanan"]])

        not_cancel = ~looks_cancel
        parseable = created.notna()

        # Ambiguity check: hanya periksa Voucher Ditanggung Penjual per Order ID
        rebate_rows = pd.DataFrame({"id": ids, "voucher": voucher})
        rebate_rows = rebate_rows[rebate_rows["id"].ne("") & not_cancel & parseable]
        pair_counts = rebate_rows.drop_duplicates().groupby("id").size()
        ambiguous_rebate_orders = int((pair_counts > 1).sum())
        add(
            "PASS" if ambiguous_rebate_orders == 0 else "FAIL",
            "Seller Rebate tidak ambigu",
            "Voucher Ditanggung Penjual per No. Pesanan harus konsisten agar dedupe rebate tidak overcount.",
            actual=ambiguous_rebate_orders,
            expected=0,
        )

        # Reconcile GMV All Shopee
        expected_all = float(subtotal[parseable].sum())
        actual_all = float(monthly["gmv_all"].sum())
        add(
            "PASS" if abs(actual_all - expected_all) <= tol else "FAIL",
            "Reconcile GMV All Shopee",
            "SUM Subtotal Pesanan raw dibandingkan hasil monthly.",
            actual=f"Rp {actual_all:,.0f}",
            expected=f"Rp {expected_all:,.0f}",
        )

        # Reconcile GMV Exclude Batal Shopee
        expected_ex = float(subtotal[parseable & not_cancel].sum())
        actual_ex = float(monthly["gmv_ex_cancel"].sum())
        add(
            "PASS" if abs(actual_ex - expected_ex) <= tol else "FAIL",
            "Reconcile GMV Exclude Batal Shopee",
            "SUM Subtotal Pesanan non-batal raw dibandingkan hasil monthly.",
            actual=f"Rp {actual_ex:,.0f}",
            expected=f"Rp {expected_ex:,.0f}",
        )

        # Reconcile Seller Rebate Shopee:
        # MAX Voucher Ditanggung Penjual per unique No. Pesanan non-batal, lalu SUM.
        seller_check_df = pd.DataFrame({
            "month": created.dt.to_period("M").astype("string"),
            "id": ids,
            "voucher": voucher,
        })
        seller_check_df = seller_check_df[seller_check_df["id"].ne("") & not_cancel & parseable]
        expected_seller = float(
            seller_check_df.groupby(["month", "id"])["voucher"].max().groupby("month").sum().sum()
        ) if not seller_check_df.empty else 0.0
        actual_seller = float(monthly["seller_rebate"].sum())
        add(
            "PASS" if abs(actual_seller - expected_seller) <= tol else "FAIL",
            "Reconcile Seller Rebate Shopee",
            "MAX Voucher Ditanggung Penjual per unique No. Pesanan non-batal dibandingkan hasil monthly.",
            actual=f"Rp {actual_seller:,.0f}",
            expected=f"Rp {expected_seller:,.0f}",
        )

        # Reconcile Diskon Dari Shopee:
        # SUM Diskon Dari Shopee untuk row non-batal.
        expected_platform = float(shopee_discount[parseable & not_cancel].sum())
        actual_platform = float(monthly["platform_discount"].sum())
        add(
            "PASS" if abs(actual_platform - expected_platform) <= tol else "FAIL",
            "Reconcile Diskon Dari Shopee",
            "SUM Diskon Dari Shopee non-batal raw dibandingkan hasil monthly.",
            actual=f"Rp {actual_platform:,.0f}",
            expected=f"Rp {expected_platform:,.0f}",
        )

    # ---------------------------------------------------------
    # TIKTOK / TOKOPEDIA
    # ---------------------------------------------------------
    elif marketplace in {"TIK", "TOK"}:
        cols = _require_columns(
            df,
            marketplace,
            {
                "Order ID": (),
                "Order Sub Status": (),
                "SKU Subtotal After Discount": (),
                "SKU Platform Discount": (),
                "Created Time": (),
            },
        )

        ids = _normalized_id_series(df[cols["Order ID"]])
        sub_status = df[cols["Order Sub Status"]].astype("string").fillna("").str.strip().str.upper()
        created = _parse_datetime(df[cols["Created Time"]], dayfirst=True)
        subtotal = _to_idr_number(df[cols["SKU Subtotal After Discount"]])
        platform = _to_idr_number(df[cols["SKU Platform Discount"]])

        nonempty_date = df[cols["Created Time"]].astype("string").fillna("").str.strip().ne("")
        invalid_dates = int((nonempty_date & created.isna()).sum())
        add(
            "PASS" if invalid_dates == 0 else "FAIL",
            f"Parsing tanggal {marketplace}",
            "Semua Created Time harus dapat diterjemahkan menjadi tanggal.",
            actual=invalid_dates,
            expected=0,
        )

        for field in ["SKU Subtotal After Discount", "SKU Platform Discount"]:
            count, samples = _nonempty_invalid_numeric_count(df[cols[field]])
            add(
                "PASS" if count == 0 else "FAIL",
                f"Parsing nominal {field}",
                "Nominal non-kosong tidak boleh diam-diam berubah menjadi 0 karena format tidak dikenali."
                + (f" Contoh: {samples}" if samples else ""),
                actual=count,
                expected=0,
            )

        eligibility = sub_status.eq("SELESAI") & ids.ne("")
        valid_mask = eligibility & created.notna()

        order_months = pd.DataFrame({
            "id": ids,
            "month": created.dt.to_period("M").astype("string"),
        })
        order_months = order_months[order_months["id"].ne("") & valid_mask]
        month_conflicts = int((order_months.groupby("id")["month"].nunique(dropna=True) > 1).sum())
        add(
            "PASS" if month_conflicts == 0 else "FAIL",
            "Konsistensi bulan per Order ID",
            "Satu Order ID tidak boleh memiliki Created Time pada bulan yang berbeda.",
            actual=month_conflicts,
            expected=0,
        )

        expected_subtotal = float(subtotal[valid_mask].sum())
        expected_platform = float(platform[valid_mask].sum())
        actual_subtotal = float(monthly["gmv_ex_cancel"].sum())
        actual_platform = float(monthly["platform_discount"].sum())

        add(
            "PASS" if abs(actual_subtotal - expected_subtotal) <= tol else "FAIL",
            f"Reconcile GMV Exclude Batal {marketplace}",
            f"SUM SKU Subtotal After Discount untuk Order Sub Status SELESAI dibandingkan output monthly.",
            actual=f"Rp {actual_subtotal:,.0f}",
            expected=f"Rp {expected_subtotal:,.0f}",
        )
        add(
            "PASS" if abs(actual_platform - expected_platform) <= tol else "FAIL",
            f"Reconcile Platform Discount {marketplace}",
            f"SUM SKU Platform Discount untuk Order Sub Status SELESAI dibandingkan output monthly.",
            actual=f"Rp {actual_platform:,.0f}",
            expected=f"Rp {expected_platform:,.0f}",
        )

    # ---------------------------------------------------------
    # LAZADA
    # ---------------------------------------------------------
    elif marketplace == "LAZ":
        cols = _require_columns(
            df,
            "LAZ",
            {
                "orderItemId": (),
                "createTime": (),
                "unitPrice": (),
                "platformDiscountTotal": (),
                "status": (),
            },
        )

        ids = _normalized_id_series(df[cols["orderItemId"]])
        status = df[cols["status"]].astype("string").fillna("").str.strip().str.lower()
        created = _parse_datetime(df[cols["createTime"]], dayfirst=True)
        unit_price = _to_idr_number(df[cols["unitPrice"]])
        platform = _to_idr_number(df[cols["platformDiscountTotal"]])

        nonempty_date = df[cols["createTime"]].astype("string").fillna("").str.strip().ne("")
        invalid_dates = int((nonempty_date & created.isna()).sum())
        add(
            "PASS" if invalid_dates == 0 else "FAIL",
            "Parsing tanggal Lazada",
            "Semua createTime harus dapat diterjemahkan menjadi tanggal.",
            actual=invalid_dates,
            expected=0,
        )

        for field in ["unitPrice", "platformDiscountTotal"]:
            count, samples = _nonempty_invalid_numeric_count(df[cols[field]])
            add(
                "PASS" if count == 0 else "FAIL",
                f"Parsing nominal {field}",
                "Nominal non-kosong tidak boleh diam-diam berubah menjadi 0 karena format tidak dikenali."
                + (f" Contoh: {samples}" if samples else ""),
                actual=count,
                expected=0,
            )

        exact_cancel = {"canceled", "cancelled"}
        looks_cancel = status.str.contains("cancel", regex=False, na=False)
        missed_cancel = sorted(set(status[looks_cancel & ~status.isin(exact_cancel)].tolist()))
        add(
            "PASS" if not missed_cancel else "FAIL",
            "Status cancel Lazada",
            "Status yang mengandung kata cancel harus cocok dengan rule exclude saat ini."
            + (f" Status berisiko: {missed_cancel[:10]}" if missed_cancel else ""),
            actual=len(missed_cancel),
            expected=0,
        )

        duplicate_item_ids = int(ids[ids.ne("")].duplicated(keep=False).sum())
        add(
            "PASS" if duplicate_item_ids == 0 else "WARNING",
            "Duplikasi orderItemId dalam file",
            "orderItemId Lazada idealnya unik per line item. Duplikasi perlu ditinjau agar tidak double count.",
            actual=duplicate_item_ids,
            expected=0,
        )

        parseable = created.notna()
        not_cancel = ~status.isin(exact_cancel)
        expected_all = float(unit_price[parseable].sum())
        expected_ex = float(unit_price[parseable & not_cancel].sum())
        expected_platform = float(platform[parseable & not_cancel].sum())
        actual_all = float(monthly["gmv_all"].sum())
        actual_ex = float(monthly["gmv_ex_cancel"].sum())
        actual_platform = float(monthly["platform_discount"].sum())

        add(
            "PASS" if abs(actual_all - expected_all) <= tol else "FAIL",
            "Reconcile GMV All Lazada",
            "SUM unitPrice raw dibandingkan output monthly.",
            actual=f"Rp {actual_all:,.0f}",
            expected=f"Rp {expected_all:,.0f}",
        )
        add(
            "PASS" if abs(actual_ex - expected_ex) <= tol else "FAIL",
            "Reconcile GMV Exclude Batal Lazada",
            "SUM unitPrice non-cancel dibandingkan output monthly.",
            actual=f"Rp {actual_ex:,.0f}",
            expected=f"Rp {expected_ex:,.0f}",
        )
        add(
            "PASS" if abs(actual_platform - expected_platform) <= tol else "FAIL",
            "Reconcile Platform Discount Lazada",
            "SUM platformDiscountTotal non-cancel dibandingkan output monthly.",
            actual=f"Rp {actual_platform:,.0f}",
            expected=f"Rp {expected_platform:,.0f}",
        )

    # Generic negative-value sanity check.
    negative_cols = [
        c for c in ["gmv_all", "gmv_ex_cancel", "platform_discount", "seller_rebate", "benchmark_gmv"]
        if c in monthly.columns and (monthly[c] < 0).any()
    ]
    add(
        "PASS" if not negative_cols else "WARNING",
        "Nilai bulanan non-negatif",
        "Komponen benchmark umumnya tidak diharapkan negatif."
        + (f" Kolom negatif: {negative_cols}" if negative_cols else ""),
        actual=len(negative_cols),
        expected=0,
    )

    return audit


def validate_final_result(result: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Audit independent for combined monthly and final benchmark arithmetic."""
    audit: list[dict[str, Any]] = []
    tol = 0.5

    def add(status: str, check: str, detail: str, actual: Any = "", expected: Any = ""):
        audit.append(_audit_row(status, check, detail, actual=actual, expected=expected))

    combined = result["monthly_combined"].copy()
    marketplace = result["monthly_marketplace"].copy()

    expected_total = combined[["shopee", "tiktok", "tokopedia", "lazada"]].sum(axis=1)
    total_diff = float((combined["total_gmv"] - expected_total).abs().max()) if len(combined) else 0.0
    add(
        "PASS" if total_diff <= tol else "FAIL",
        "Total GMV per bulan",
        "Total GMV harus sama dengan SHO + TIK + TOK + LAZ pada setiap bulan.",
        actual=f"max selisih Rp {total_diff:,.0f}",
        expected="Rp 0",
    )

    # Determine if include_empty_months was active:
    month_has_rows = (
        marketplace.groupby("month")["source_rows"].sum()
    ) if "source_rows" in marketplace.columns else pd.Series(dtype=float)
    active_months = month_has_rows[month_has_rows > 0].index

    valid_months_result = int(result["valid_months"])
    total_months_combined = int(combined["month"].nunique())

    if valid_months_result == len(active_months) and valid_months_result != total_months_combined:
        # include_empty_months = False
        active_series = combined[combined["month"].isin(active_months)]["total_gmv"]
        expected_benchmark = float(active_series.mean()) if len(active_series) else 0.0
        expected_months = valid_months_result
        contribution_sum = 0.0
        if expected_months:
            for col in ["shopee", "tiktok", "tokopedia", "lazada"]:
                active_col_sum = combined[combined["month"].isin(active_months)][col].sum()
                contribution_sum += float(active_col_sum) / expected_months
    else:
        # include_empty_months = True
        expected_benchmark = float(expected_total.mean()) if len(combined) else 0.0
        expected_months = total_months_combined
        contribution_sum = 0.0
        if expected_months:
            for col in ["shopee", "tiktok", "tokopedia", "lazada"]:
                contribution_sum += float(combined[col].sum()) / expected_months

    bench_diff = abs(float(result["benchmark"]) - expected_benchmark)
    add(
        "PASS" if bench_diff <= tol else "FAIL",
        "Final Benchmark Average",
        "Benchmark final dihitung ulang sebagai average Total GMV seluruh bulan valid.",
        actual=f"Rp {float(result['benchmark']):,.0f}",
        expected=f"Rp {expected_benchmark:,.0f}",
    )

    add(
        "PASS" if int(result["valid_months"]) == expected_months else "FAIL",
        "Jumlah bulan benchmark",
        "Jumlah bulan final harus sama dengan jumlah distinct calendar month di tabel combined.",
        actual=int(result["valid_months"]),
        expected=expected_months,
    )

    contribution_diff = abs(contribution_sum - float(result["benchmark"]))
    add(
        "PASS" if contribution_diff <= tol else "FAIL",
        "Ringkasan marketplace = Benchmark",
        "Jumlah average contribution SHO/TIK/TOK/LAZ harus sama dengan final benchmark.",
        actual=f"Rp {contribution_sum:,.0f}",
        expected=f"Rp {float(result['benchmark']):,.0f}",
    )

    # Validate marketplace formula by row again after multi-file consolidation.
    formula_failures = 0
    for _, row in marketplace.iterrows():
        mp = row["marketplace"]
        if mp == "SHO":
            expected = row["gmv_ex_cancel"] - row["seller_rebate"] + row["platform_discount"]
        else:
            expected = row["gmv_ex_cancel"] + row["platform_discount"]
        if abs(float(row["benchmark_gmv"]) - float(expected)) > tol:
            formula_failures += 1
    add(
        "PASS" if formula_failures == 0 else "FAIL",
        "Formula marketplace setelah merge file",
        "Setiap baris marketplace-bulan direkonsiliasi ulang setelah seluruh file digabung.",
        actual=formula_failures,
        expected=0,
    )

    # Detect internal month gaps (e.g. Jan, Mar without Feb).
    months = sorted(pd.to_datetime(combined["month"].dropna().unique()))
    missing_months: list[str] = []
    if months:
        expected_range = pd.date_range(months[0], months[-1], freq="MS")
        actual_set = {pd.Timestamp(m) for m in months}
        missing_months = [m.strftime("%b %Y") for m in expected_range if m not in actual_set]
    add(
        "PASS" if not missing_months else "WARNING",
        "Kontinuitas periode",
        "Tidak ada bulan yang lompat di antara periode awal dan akhir."
        + (f" Bulan yang tidak ada: {missing_months}" if missing_months else ""),
        actual=len(missing_months),
        expected=0,
    )

    return audit


def calculate_benchmark(
    monthly_frames: Iterable[pd.DataFrame],
    include_empty_months: bool = True,
) -> dict[str, Any]:
    frames = [
        frame.copy()
        for frame in monthly_frames
        if frame is not None and not frame.empty
    ]

    if not frames:
        raise BenchmarkInputError(
            "Tidak ada data marketplace valid untuk dihitung."
        )

    monthly = pd.concat(
        frames,
        ignore_index=True,
    )

    monthly["month"] = pd.to_datetime(
        monthly["month"],
        errors="coerce",
    )

    monthly = monthly[
        monthly["month"].notna()
    ].copy()

    numeric_cols = [
        "gmv_all",
        "gmv_ex_cancel",
        "platform_discount",
        "seller_rebate",
        "benchmark_gmv",
        "source_rows",
    ]

    for col in numeric_cols:
        if col not in monthly.columns:
            monthly[col] = 0.0

    monthly_marketplace = (
        monthly.groupby(
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
        .sort_values(["month", "marketplace"])
        .reset_index(drop=True)
    )

    pivot = monthly_marketplace.pivot_table(
        index="month",
        columns="marketplace",
        values="benchmark_gmv",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    for mp in ["SHO", "TIK", "TOK", "LAZ"]:
        if mp not in pivot.columns:
            pivot[mp] = 0.0

    monthly_combined = pivot[
        ["month", "SHO", "TIK", "TOK", "LAZ"]
    ].copy()

    monthly_combined = monthly_combined.rename(
        columns={
            "SHO": "shopee",
            "TIK": "tiktok",
            "TOK": "tokopedia",
            "LAZ": "lazada",
        }
    )

    monthly_combined["total_gmv"] = (
        monthly_combined[
            [
                "shopee",
                "tiktok",
                "tokopedia",
                "lazada",
            ]
        ].sum(axis=1)
    )

    monthly_combined = (
        monthly_combined
        .sort_values("month")
        .reset_index(drop=True)
    )

    # --------------------------------------------------------
    # Average calculation — respects include_empty_months flag
    # --------------------------------------------------------
    if include_empty_months:
        # Count ALL months (incl. zero-order months)
        benchmark_series = monthly_combined["total_gmv"]
    else:
        # Count only months with at least 1 order row
        # Determine which months have any source_rows > 0
        month_has_rows = (
            monthly_marketplace
            .groupby("month")["source_rows"]
            .sum()
        )
        active_months = month_has_rows[month_has_rows > 0].index
        benchmark_series = monthly_combined[
            monthly_combined["month"].isin(active_months)
        ]["total_gmv"]

    if benchmark_series.empty:
        benchmark = 0.0
        valid_months = 0
    else:
        benchmark = float(benchmark_series.mean())
        valid_months = int(len(benchmark_series))

    marketplace_summary = (
        monthly_marketplace.groupby(
            "marketplace",
            as_index=False,
        )
        .agg(
            valid_months=("month", "nunique"),
            total_benchmark_gmv=("benchmark_gmv", "sum"),
        )
    )

    warnings: list[str] = []

    month_sets = (
        monthly_marketplace.groupby(
            "marketplace"
        )["month"]
        .apply(
            lambda s: tuple(sorted(set(s)))
        )
        .tolist()
    )

    if len(set(month_sets)) > 1:
        warnings.append(
            "Periode data antar-marketplace tidak sepenuhnya sama. "
            "Marketplace yang tidak memiliki data pada suatu bulan dianggap 0 "
            "pada bulan tersebut."
        )

    return {
        "benchmark": benchmark,
        "valid_months": valid_months,
        "monthly_marketplace": monthly_marketplace,
        "monthly_combined": monthly_combined,
        "marketplace_summary": marketplace_summary,
        "warnings": warnings,
    }


def process_uploaded_files(
    files: Iterable[Mapping[str, Any]],
) -> dict[str, Any]:
    monthly_frames = []
    inspections = []

    for item in files:
        source = item["source"]
        filename = item.get(
            "filename",
            getattr(
                source,
                "name",
                "uploaded_file.xlsx",
            ),
        )
        hint = item.get("marketplace_hint")

        df, marketplace, header_row = read_order_file(
            source,
            filename=filename,
            marketplace_hint=hint,
        )

        monthly = process_dataframe(
            df,
            marketplace,
        )

        monthly_frames.append(monthly)

        inspections.append(
            FileInspection(
                filename=filename,
                marketplace=marketplace,
                header_row=header_row,
                row_count=len(df),
                min_month=(
                    monthly["month"].min()
                    if not monthly.empty
                    else None
                ),
                max_month=(
                    monthly["month"].max()
                    if not monthly.empty
                    else None
                ),
                warnings=[],
            )
        )

    result = calculate_benchmark(
        monthly_frames
    )
    result["files"] = inspections

    return result
