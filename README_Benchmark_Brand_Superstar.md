HOw to run in terminal: 
1. cd "/Users/fahroni/Downloads/web app benchmark "
2. source .venv/bin/activate
3. python -m streamlit run app.py

"""
Benchmark Brand Superstar - Processing Engine v0.1

Tujuan:
- Membaca raw order marketplace tanpa helper formula Google Sheets.
- Menghasilkan GMV bulanan per marketplace.
- Maksimal menggunakan 6 bulan terbaru.
- Mengikuti logic spreadsheet Indonesia saat ini untuk parity awal.

Current parity logic:
- SHO: benchmark GMV = Subtotal Pesanan dengan Waktu Pembayaran Dilakukan terisi.
- TIK/TOK: benchmark GMV = SKU Subtotal After Discount dengan Paid Time terisi.
- LAZ: benchmark GMV = paidPrice + platformDiscountTotal, exclude status == canceled.
- Final benchmark = SUM(average benchmark GMV per marketplace).

Catatan:
TIK dan TOK menggunakan struktur export yang sama. Jika nama file tidak cukup untuk
membedakan, marketplace_hint wajib diberikan ("TIK" atau "TOK").
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, BinaryIO, Iterable, Mapping
import io
import re

import numpy as np
import pandas as pd


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


# -----------------------------
# Header / marketplace detection
# -----------------------------

SIGNATURES = {
    "SHO": {
        "No. Pesanan",
        "Status Pesanan",
        "Waktu Pesanan Dibuat",
        "Waktu Pembayaran Dilakukan",
        "Subtotal Pesanan",
    },
    "TIKTOK_FAMILY": {
        "Order ID",
        "Order Status",
        "SKU Subtotal After Discount",
        "Created Time",
        "Paid Time",
    },
    "LAZ": {
        "orderItemId",
        "createTime",
        "paidPrice",
        "platformDiscountTotal",
        "status",
    },
}


def _clean_header(value: Any) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _detect_template(columns: Iterable[Any]) -> str:
    cols = {_clean_header(c) for c in columns if _clean_header(c)}
    scores = {
        name: len(required & cols) / len(required)
        for name, required in SIGNATURES.items()
    }
    best = max(scores, key=scores.get)
    if scores[best] < 0.8:
        missing = {
            name: sorted(required - cols)
            for name, required in SIGNATURES.items()
        }
        raise BenchmarkInputError(
            "Format file belum dikenali. Required header tidak cukup cocok. "
            f"Missing candidates: {missing}"
        )
    return best


def _resolve_tik_tok(filename: str, marketplace_hint: str | None) -> str:
    if marketplace_hint:
        hint = marketplace_hint.strip().upper()
        if hint in {"TIK", "TIKTOK"}:
            return "TIK"
        if hint in {"TOK", "TOKOPEDIA"}:
            return "TOK"
        raise BenchmarkInputError(
            "marketplace_hint untuk format TikTok/Tokopedia harus TIK atau TOK."
        )

    name = filename.lower()
    if "tokopedia" in name or re.search(r"(^|[^a-z])tok([^a-z]|$)", name):
        return "TOK"
    if "tiktok" in name or "tik tok" in name or re.search(r"(^|[^a-z])tik([^a-z]|$)", name):
        return "TIK"

    raise BenchmarkInputError(
        f"File '{filename}' memakai format TikTok/Tokopedia yang sama. "
        "Marketplace tidak dapat dibedakan dari header saja; berikan marketplace_hint='TIK' atau 'TOK'."
    )


def _rewind(source: Any) -> None:
    if hasattr(source, "seek"):
        source.seek(0)


def _score_header_values(values: Iterable[Any]) -> tuple[float, str]:
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
    """
    Cari header marketplace pada beberapa baris awal.

    TikTok/Tokopedia sering berbentuk:
      row 1 = header
      row 2 = deskripsi field
      row 3+ = data

    Posisi header tidak diasumsikan selalu row pertama.
    """
    best: tuple[float, int, str] | None = None

    for idx, row in preview.iterrows():
        values = row.tolist()
        cols = {_clean_header(v) for v in values if _clean_header(v)}

        if {
            "Order ID",
            "Order Status",
            "Created Time",
            "SKU Subtotal After Discount",
        }.issubset(cols):
            return int(idx), "TIKTOK_FAMILY", 1.0

        if {
            "No. Pesanan",
            "Status Pesanan",
            "Waktu Pesanan Dibuat",
            "Subtotal Pesanan",
        }.issubset(cols):
            return int(idx), "SHO", 1.0

        if {
            "orderItemId",
            "createTime",
            "paidPrice",
            "status",
        }.issubset(cols):
            return int(idx), "LAZ", 1.0

        score, template = _score_header_values(values)
        candidate = (score, int(idx), template)

        if best is None or candidate[0] > best[0]:
            best = candidate

    if best is None:
        return 0, "", 0.0

    return best[1], best[2], best[0]


def _scan_excel_workbook(
    source: Any,
    filename: str,
    nrows: int = 30,
) -> tuple[str | int, int, str]:
    """
    Scan semua worksheet dalam file Excel.

    Ini penting karena beberapa export marketplace memiliki sheet awal
    yang bukan tabel order utama.
    """
    _rewind(source)

    try:
        excel = pd.ExcelFile(source)
    except Exception as exc:
        raise BenchmarkInputError(
            f"File Excel '{filename}' tidak dapat dibuka: {exc}"
        ) from exc

    best_candidate = None
    debug_previews = []

    for sheet_name in excel.sheet_names:
        try:
            preview = pd.read_excel(
                excel,
                sheet_name=sheet_name,
                header=None,
                nrows=nrows,
                dtype=object,
            )
        except Exception:
            continue

        header_row, template, score = _find_header_in_dataframe(preview)
        width = int(preview.shape[1])

        candidate = {
            "score": score,
            "width": width,
            "sheet_name": sheet_name,
            "header_row": header_row,
            "template": template,
        }

        if (
            best_candidate is None
            or (candidate["score"], candidate["width"])
            > (best_candidate["score"], best_candidate["width"])
        ):
            best_candidate = candidate

        rows_preview = []
        for idx, row in preview.head(5).iterrows():
            vals = [_clean_header(v) for v in row.tolist()]
            vals = [v for v in vals if v][:8]
            if vals:
                rows_preview.append(
                    f"row {idx + 1}: {vals}"
                )

        debug_previews.append(
            f"sheet '{sheet_name}' [{width} kolom] "
            + " | ".join(rows_preview[:3])
        )

    if best_candidate is None:
        raise BenchmarkInputError(
            f"Tidak ada worksheet yang dapat dibaca dari '{filename}'."
        )

    if (
        best_candidate["score"] < 0.6
        or best_candidate["width"] < 4
    ):
        raise BenchmarkInputError(
            "Header marketplace belum berhasil ditemukan. "
            "Sistem sudah memeriksa semua worksheet. "
            "Detail: "
            + " || ".join(debug_previews[:8])
        )

    return (
        best_candidate["sheet_name"],
        best_candidate["header_row"],
        best_candidate["template"],
    )


def _read_csv_best(
    source: Any,
    *,
    header: int | None,
    nrows: int | None = None,
) -> pd.DataFrame:
    """
    Coba beberapa delimiter dan pilih hasil dengan jumlah kolom terbanyak.
    """
    attempts = [
        {"sep": None, "engine": "python"},
        {"sep": ",", "engine": "python"},
        {"sep": ";", "engine": "python"},
        {"sep": "\t", "engine": "python"},
    ]

    best_df = None
    best_width = 0

    for kwargs in attempts:
        try:
            _rewind(source)
            df = pd.read_csv(
                source,
                header=header,
                nrows=nrows,
                dtype=object,
                **kwargs,
            )

            if df.shape[1] > best_width:
                best_df = df
                best_width = df.shape[1]

        except Exception:
            continue

    if best_df is None:
        raise BenchmarkInputError(
            "File CSV tidak dapat dibaca."
        )

    return best_df


def _find_header_row(preview: pd.DataFrame) -> tuple[int, str]:
    header_row, template, score = _find_header_in_dataframe(preview)

    if score < 0.6:
        preview_rows = []

        for idx, row in preview.head(10).iterrows():
            values = [_clean_header(v) for v in row.tolist()]
            values = [v for v in values if v][:8]

            if values:
                preview_rows.append(
                    f"row {idx + 1}: {values}"
                )

        context = " | ".join(preview_rows[:5])

        raise BenchmarkInputError(
            "Header utama tidak ditemukan pada beberapa baris pertama file. "
            f"Preview terbaca: {context}"
        )

    return header_row, template


def read_order_file(
    source: str | Path | BinaryIO,
    *,
    filename: str | None = None,
    marketplace_hint: str | None = None,
) -> tuple[pd.DataFrame, str, int]:
    """
    Membaca raw order marketplace.

    Mendukung:
    - Excel multi-sheet
    - TikTok/Tokopedia dengan header + baris deskripsi
    - header yang tidak selalu berada di row pertama
    - CSV dengan delimiter berbeda
    """
    if isinstance(source, (str, Path)):
        source_obj: Any = str(source)
        resolved_filename = filename or Path(source).name
    else:
        source_obj = source
        resolved_filename = filename or getattr(
            source,
            "name",
            "uploaded_file.xlsx",
        )

    ext = Path(resolved_filename).suffix.lower()

    if ext in {".xlsx", ".xls", ".xlsm"}:
        sheet_name, header_row, detected_template = _scan_excel_workbook(
            source_obj,
            resolved_filename,
            nrows=30,
        )

        _rewind(source_obj)

        df = pd.read_excel(
            source_obj,
            sheet_name=sheet_name,
            header=header_row,
            dtype=object,
        )

    elif ext in {".csv", ".txt"}:
        preview = _read_csv_best(
            source_obj,
            header=None,
            nrows=30,
        )

        header_row, detected_template = _find_header_row(preview)

        df = _read_csv_best(
            source_obj,
            header=header_row,
            nrows=None,
        )

    else:
        raise BenchmarkInputError(
            f"Format '{ext or 'unknown'}' belum didukung. "
            "Gunakan .xlsx, .xls, .xlsm, atau .csv."
        )

    df.columns = [_clean_header(c) for c in df.columns]

    # Drop kolom kosong/Unnamed.
    keep_cols = [
        c
        for c in df.columns
        if c
        and not str(c).lower().startswith("unnamed:")
    ]
    df = df.loc[:, keep_cols]

    df = df.dropna(how="all").reset_index(drop=True)

    detected_template = _detect_template(df.columns)

    if detected_template == "TIKTOK_FAMILY":
        marketplace = _resolve_tik_tok(
            resolved_filename,
            marketplace_hint,
        )
        df = _drop_tiktok_description_row(df)
    else:
        marketplace = detected_template

    return df, marketplace, header_row


def _drop_tiktok_description_row(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Hapus baris deskripsi bawaan TikTok/Tokopedia setelah header.
    """
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

        order_id_lower = order_id.lower()

        looks_like_description = (
            "platform unique order id" in order_id_lower
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

        # Kadang terdapat satu row metadata/kosong sebelum data.
        if not looks_like_real_order:
            nonempty_count = sum(
                1
                for value in row.tolist()
                if str(value).strip().lower()
                not in {"", "nan", "none"}
            )

            if nonempty_count <= 3:
                drop_count += 1
                continue

        break

    if drop_count:
        df = df.iloc[drop_count:].reset_index(drop=True)

    if not df.empty:
        first_order_id = str(
            df.iloc[0].get("Order ID", "")
        ).strip()

        first_order_id = re.sub(
            r"\.0$",
            "",
            first_order_id,
        )

        if not re.fullmatch(
            r"\d{8,}",
            first_order_id,
        ):
            raise BenchmarkInputError(
                "Header TikTok/Tokopedia sudah ditemukan, "
                "tetapi baris data pertama belum terbaca sebagai Order ID. "
                f"Nilai pertama: '{first_order_id}'."
            )

    return df


def _to_idr_number(series: pd.Series) -> pd.Series:
    """
    Benchmark Indonesia memakai nominal integer.
    Mengikuti spirit formula lama yang membuang separator ribuan '.' / ','.
    """
    text = series.fillna("").astype(str).str.strip()
    negative_parentheses = text.str.match(r"^\(.*\)$")

    cleaned = (
        text.str.replace(r"[^0-9-]", "", regex=True)
            .replace({"": np.nan, "-": np.nan})
    )
    out = pd.to_numeric(cleaned, errors="coerce").fillna(0.0)
    out.loc[negative_parentheses] = -out.loc[negative_parentheses].abs()
    return out.astype(float)


def _has_value(series: pd.Series) -> pd.Series:
    text = series.fillna("").astype(str).str.strip().str.lower()
    return ~text.isin({"", "-", "nan", "nat", "none"})


def _parse_datetime(series: pd.Series, *, dayfirst: bool = False) -> pd.Series:
    # Preserve actual Timestamp/datetime values where possible.
    result = pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)

    # Excel serial date fallback for numeric-looking values.
    missing = result.isna()
    if missing.any():
        numeric = pd.to_numeric(series[missing], errors="coerce")
        serial_mask = numeric.between(20000, 80000, inclusive="both")
        if serial_mask.any():
            serial_dates = pd.to_datetime(
                numeric[serial_mask],
                unit="D",
                origin="1899-12-30",
                errors="coerce",
            )
            result.loc[serial_dates.index] = serial_dates

    return result


def _month_start(series: pd.Series, *, dayfirst: bool = False) -> pd.Series:
    dt = _parse_datetime(series, dayfirst=dayfirst)
    return dt.dt.to_period("M").dt.to_timestamp()


def _validate_columns(df: pd.DataFrame, required: Iterable[str], marketplace: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise BenchmarkInputError(
            f"{marketplace}: kolom wajib tidak ditemukan: {', '.join(missing)}"
        )


# -----------------------------
# Marketplace processors
# -----------------------------
# -----------------------------
# Marketplace processors
# -----------------------------

def process_shopee(df: pd.DataFrame) -> pd.DataFrame:
    """
    SHO benchmark per month:
      GMV Exclude Batal - Seller Rebate

    GMV base:
      Subtotal Pesanan

    Seller Rebate:
      Voucher Ditanggung Penjual + Paket Diskon (Diskon dari Penjual)
      Dihitung unik per kombinasi No. Pesanan + kedua komponen rebate,
      mengikuti helper spreadsheet existing agar tidak double count per SKU.
    """
    required = [
        "No. Pesanan",
        "Status Pesanan",
        "Waktu Pesanan Dibuat",
        "Subtotal Pesanan",
        "Voucher Ditanggung Penjual",
        "Paket Diskon (Diskon dari Penjual)",
    ]
    _validate_columns(df, required, "SHO")

    month = _month_start(df["Waktu Pesanan Dibuat"], dayfirst=False)
    status = df["Status Pesanan"].fillna("").astype(str).str.strip().str.lower()
    not_cancelled = status.ne("batal")

    base = pd.DataFrame({
        "month": month,
        "gmv_raw": _to_idr_number(df["Subtotal Pesanan"]),
        "not_cancelled": not_cancelled,
    })
    base = base[base["month"].notna()].copy()
    base["gmv_ex_cancel"] = np.where(
        base["not_cancelled"], base["gmv_raw"], 0.0
    )

    # Seller rebate hanya untuk order non-batal.
    rebate_rows = pd.DataFrame({
        "month": month,
        "order_id": df["No. Pesanan"].fillna("").astype(str).str.strip(),
        "voucher_seller": _to_idr_number(df["Voucher Ditanggung Penjual"]),
        "bundle_seller": _to_idr_number(df["Paket Diskon (Diskon dari Penjual)"]),
        "not_cancelled": not_cancelled,
    })
    rebate_rows = rebate_rows[
        rebate_rows["month"].notna()
        & rebate_rows["not_cancelled"]
        & rebate_rows["order_id"].ne("")
    ].copy()

    # Mirror spreadsheet helper: UNIQUE(HSTACK(order_id, voucher, bundle_discount)).
    rebate_rows = rebate_rows.drop_duplicates(
        subset=["month", "order_id", "voucher_seller", "bundle_seller"]
    )
    rebate_rows["seller_rebate"] = (
        rebate_rows["voucher_seller"] + rebate_rows["bundle_seller"]
    )

    monthly_gmv = (
        base.groupby("month", as_index=False)
        .agg(
            gmv_all=("gmv_raw", "sum"),
            gmv_ex_cancel=("gmv_ex_cancel", "sum"),
            source_rows=("gmv_raw", "size"),
        )
    )

    monthly_rebate = (
        rebate_rows.groupby("month", as_index=False)
        .agg(seller_rebate=("seller_rebate", "sum"))
    )

    out = monthly_gmv.merge(monthly_rebate, on="month", how="left")
    out["seller_rebate"] = out["seller_rebate"].fillna(0.0)
    out["benchmark_gmv"] = out["gmv_ex_cancel"] - out["seller_rebate"]
    out["platform_discount"] = 0.0
    out["marketplace"] = "SHO"

    return out[
        [
            "marketplace", "month", "gmv_all", "gmv_ex_cancel",
            "platform_discount", "seller_rebate", "benchmark_gmv", "source_rows"
        ]
    ]


def process_tik_family(df: pd.DataFrame, marketplace: str) -> pd.DataFrame:
    """
    TIK/TOK benchmark per month:
      GMV Exclude Batal + SKU Platform Discount

    GMV base:
      SKU Subtotal After Discount

    Exclude:
      Order Status == Dibatalkan
    """
    if marketplace not in {"TIK", "TOK"}:
        raise BenchmarkInputError("Marketplace TikTok-family harus TIK atau TOK.")

    required = [
        "Order Status",
        "Created Time",
        "SKU Subtotal After Discount",
        "SKU Platform Discount",
    ]
    _validate_columns(df, required, marketplace)

    month = _month_start(df["Created Time"], dayfirst=True)
    status = df["Order Status"].fillna("").astype(str).str.strip().str.lower()
    not_cancelled = ~status.isin({"dibatalkan", "cancelled", "canceled"})

    gmv = _to_idr_number(df["SKU Subtotal After Discount"])
    platform_discount = _to_idr_number(df["SKU Platform Discount"])

    work = pd.DataFrame({
        "month": month,
        "gmv_raw": gmv,
        "platform_discount_raw": platform_discount,
        "not_cancelled": not_cancelled,
    })
    work = work[work["month"].notna()].copy()

    work["gmv_ex_cancel"] = np.where(
        work["not_cancelled"], work["gmv_raw"], 0.0
    )
    work["platform_discount"] = np.where(
        work["not_cancelled"], work["platform_discount_raw"], 0.0
    )
    work["benchmark_gmv"] = (
        work["gmv_ex_cancel"] + work["platform_discount"]
    )

    out = (
        work.groupby("month", as_index=False)
        .agg(
            gmv_all=("gmv_raw", "sum"),
            gmv_ex_cancel=("gmv_ex_cancel", "sum"),
            platform_discount=("platform_discount", "sum"),
            benchmark_gmv=("benchmark_gmv", "sum"),
            source_rows=("gmv_raw", "size"),
        )
    )
    out["seller_rebate"] = 0.0
    out["marketplace"] = marketplace

    return out[
        [
            "marketplace", "month", "gmv_all", "gmv_ex_cancel",
            "platform_discount", "seller_rebate", "benchmark_gmv", "source_rows"
        ]
    ]


def process_lazada(df: pd.DataFrame) -> pd.DataFrame:
    """
    LAZ benchmark per month:
      GMV Exclude Batal + Platform Discount

    GMV base:
      paidPrice

    Platform discount:
      platformDiscountTotal

    Exclude:
      status == canceled
    """
    required = [
        "createTime",
        "paidPrice",
        "platformDiscountTotal",
        "status",
    ]
    _validate_columns(df, required, "LAZ")

    month = _month_start(df["createTime"], dayfirst=True)
    status = df["status"].fillna("").astype(str).str.strip().str.lower()
    not_cancelled = ~status.isin({"canceled", "cancelled"})

    paid_price = _to_idr_number(df["paidPrice"])
    platform_discount = _to_idr_number(df["platformDiscountTotal"])

    work = pd.DataFrame({
        "month": month,
        "gmv_raw": paid_price,
        "platform_discount_raw": platform_discount,
        "not_cancelled": not_cancelled,
    })
    work = work[work["month"].notna()].copy()

    work["gmv_ex_cancel"] = np.where(
        work["not_cancelled"], work["gmv_raw"], 0.0
    )
    work["platform_discount"] = np.where(
        work["not_cancelled"], work["platform_discount_raw"], 0.0
    )
    work["benchmark_gmv"] = (
        work["gmv_ex_cancel"] + work["platform_discount"]
    )

    out = (
        work.groupby("month", as_index=False)
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
            "marketplace", "month", "gmv_all", "gmv_ex_cancel",
            "platform_discount", "seller_rebate", "benchmark_gmv", "source_rows"
        ]
    ]


def process_dataframe(df: pd.DataFrame, marketplace: str) -> pd.DataFrame:
    marketplace = marketplace.upper()
    if marketplace == "SHO":
        return process_shopee(df)
    if marketplace in {"TIK", "TOK"}:
        return process_tik_family(df, marketplace)
    if marketplace == "LAZ":
        return process_lazada(df)
    raise BenchmarkInputError(f"Marketplace belum didukung: {marketplace}")


# -----------------------------
# Benchmark combination
# -----------------------------

def calculate_benchmark(
    monthly_frames: Iterable[pd.DataFrame],
) -> dict[str, Any]:
    """
    Final benchmark:
      1. Jumlahkan benchmark_gmv seluruh marketplace per bulan.
      2. Rata-ratakan semua bulan valid yang tersedia.

    Tidak ada cap 6 bulan.
    Jika input berisi 3 bulan valid, benchmark = average 3 bulan.
    Jika 8 bulan valid, benchmark = average 8 bulan.
    """
    frames = [x.copy() for x in monthly_frames if x is not None and not x.empty]
    if not frames:
        raise BenchmarkInputError("Tidak ada data marketplace valid untuk dihitung.")

    monthly = pd.concat(frames, ignore_index=True)
    monthly["month"] = pd.to_datetime(monthly["month"], errors="coerce")
    monthly = monthly[monthly["month"].notna()].copy()

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

    # Jika marketplace yang sama di-upload lebih dari sekali,
    # gabungkan per marketplace-bulan.
    monthly_marketplace = (
        monthly.groupby(["marketplace", "month"], as_index=False)
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

    # Inilah tabel utama benchmark:
    # semua marketplace dijumlahkan dulu per bulan.
    monthly_combined = (
        monthly_marketplace.groupby("month", as_index=False)
        .agg(
            shopee_gmv=("benchmark_gmv", lambda s: 0.0),
            total_benchmark_gmv=("benchmark_gmv", "sum"),
            marketplace_count=("marketplace", "nunique"),
        )
    )

    # Buat breakdown marketplace sebagai kolom agar enak ditampilkan di UI.
    pivot = (
        monthly_marketplace.pivot_table(
            index="month",
            columns="marketplace",
            values="benchmark_gmv",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )
    for mp in ["SHO", "TIK", "TOK", "LAZ"]:
        if mp not in pivot.columns:
            pivot[mp] = 0.0

    monthly_combined = pivot[["month", "SHO", "TIK", "TOK", "LAZ"]].copy()
    monthly_combined = monthly_combined.rename(
        columns={
            "SHO": "shopee",
            "TIK": "tiktok",
            "TOK": "tokopedia",
            "LAZ": "lazada",
        }
    )
    monthly_combined["total_gmv"] = monthly_combined[
        ["shopee", "tiktok", "tokopedia", "lazada"]
    ].sum(axis=1)
    monthly_combined = monthly_combined.sort_values("month").reset_index(drop=True)

    benchmark = float(monthly_combined["total_gmv"].mean())
    valid_months = int(monthly_combined["month"].nunique())

    marketplace_summary = (
        monthly_marketplace.groupby("marketplace", as_index=False)
        .agg(
            valid_months=("month", "nunique"),
            total_benchmark_gmv=("benchmark_gmv", "sum"),
        )
        .sort_values("marketplace")
        .reset_index(drop=True)
    )

    warnings: list[str] = []
    month_sets = (
        monthly_marketplace.groupby("marketplace")["month"]
        .apply(lambda s: tuple(sorted(set(s))))
        .tolist()
    )
    if len(set(month_sets)) > 1:
        warnings.append(
            "Periode data antar-marketplace tidak sepenuhnya sama. "
            "Bulan yang tidak memiliki data pada suatu marketplace dianggap 0 untuk marketplace tersebut."
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
    """
    Convenience API untuk UI Streamlit nantinya.

    files example:
    [
        {"source": uploaded_file, "filename": "shopee.xlsx"},
        {"source": uploaded_file2, "filename": "tiktok.xlsx", "marketplace_hint": "TIK"},
    ]
    """
    monthly_frames: list[pd.DataFrame] = []
    inspections: list[FileInspection] = []

    for item in files:
        source = item["source"]
        filename = item.get("filename") or getattr(source, "name", "uploaded_file.xlsx")
        hint = item.get("marketplace_hint")

        df, marketplace, header_row = read_order_file(
            source,
            filename=filename,
            marketplace_hint=hint,
        )
        monthly = process_dataframe(df, marketplace)
        monthly_frames.append(monthly)

        inspections.append(
            FileInspection(
                filename=filename,
                marketplace=marketplace,
                header_row=header_row,
                row_count=len(df),
                min_month=monthly["month"].min() if not monthly.empty else None,
                max_month=monthly["month"].max() if not monthly.empty else None,
                warnings=[],
            )
        )

    result = calculate_benchmark(monthly_frames)
    result["files"] = inspections
    return result

