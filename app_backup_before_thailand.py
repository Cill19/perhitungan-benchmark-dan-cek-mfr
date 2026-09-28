import hashlib
import itertools
import io
import zipfile
from pathlib import PurePosixPath

import streamlit as st
import pandas as pd

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


SUPPORTED_DATA_EXTENSIONS = {
    ".xlsx",
    ".xls",
    ".xlsm",
    ".csv",
    ".txt",
}

# Safety limits. ZIP is processed entirely in RAM; nothing is extracted to disk.
MAX_ZIP_DATA_FILES = 100
MAX_ZIP_MEMBER_BYTES = 300 * 1024 * 1024       # 300 MB / file
MAX_ZIP_TOTAL_BYTES = 1024 * 1024 * 1024        # 1 GB / ZIP
MAX_ZIP_COMPRESSION_RATIO = 300.0


def _is_ignored_zip_member(name: str) -> bool:
    normalized = name.replace("\\", "/")
    parts = PurePosixPath(normalized).parts
    basename = PurePosixPath(normalized).name

    if not basename:
        return True

    if "__MACOSX" in parts:
        return True

    if basename in {".DS_Store", "Thumbs.db"}:
        return True

    if basename.startswith("._"):
        return True

    return False


def _validate_zip_members(zf: zipfile.ZipFile, zip_name: str):
    """
    Validate the complete archive BEFORE any member is processed.
    This prevents a partially processed ZIP if a later member is unsafe.
    """
    data_members = []
    total_uncompressed = 0
    nested_zip_members = []

    for info in zf.infolist():
        if info.is_dir():
            continue

        member_name = info.filename

        if _is_ignored_zip_member(member_name):
            continue

        suffix = PurePosixPath(member_name).suffix.lower()

        if suffix == ".zip":
            nested_zip_members.append(member_name)
            continue

        if suffix not in SUPPORTED_DATA_EXTENSIONS:
            # Other files (images, PDFs, notes, etc.) are ignored.
            continue

        if info.file_size > MAX_ZIP_MEMBER_BYTES:
            raise BenchmarkInputError(
                f"ZIP '{zip_name}' ditolak: file '{member_name}' berukuran "
                f"{info.file_size / (1024**2):,.1f} MB, melebihi batas "
                f"{MAX_ZIP_MEMBER_BYTES / (1024**2):,.0f} MB per file."
            )

        total_uncompressed += info.file_size

        if total_uncompressed > MAX_ZIP_TOTAL_BYTES:
            raise BenchmarkInputError(
                f"ZIP '{zip_name}' ditolak: total ukuran hasil extract "
                f"melebihi {MAX_ZIP_TOTAL_BYTES / (1024**3):,.0f} GB."
            )

        if info.file_size > 1024 * 1024:
            if info.compress_size <= 0:
                raise BenchmarkInputError(
                    f"ZIP '{zip_name}' ditolak karena rasio kompresi "
                    f"file '{member_name}' tidak wajar."
                )

            ratio = info.file_size / info.compress_size

            if ratio > MAX_ZIP_COMPRESSION_RATIO:
                raise BenchmarkInputError(
                    f"ZIP '{zip_name}' ditolak karena rasio kompresi "
                    f"file '{member_name}' terlalu tinggi ({ratio:,.1f}x)."
                )

        data_members.append(info)

    if nested_zip_members:
        examples = ", ".join(nested_zip_members[:3])
        raise BenchmarkInputError(
            f"ZIP '{zip_name}' memiliki nested ZIP yang tidak didukung. "
            f"Contoh: {examples}. Extract nested ZIP terlebih dahulu."
        )

    if not data_members:
        raise BenchmarkInputError(
            f"ZIP '{zip_name}' tidak berisi file data yang didukung "
            "(.xlsx, .xls, .xlsm, .csv, .txt)."
        )

    if len(data_members) > MAX_ZIP_DATA_FILES:
        raise BenchmarkInputError(
            f"ZIP '{zip_name}' berisi {len(data_members)} file data. "
            f"Batas maksimal adalah {MAX_ZIP_DATA_FILES} file per ZIP."
        )

    return data_members


def iter_uploaded_data_files(uploaded_file):
    """
    Yield raw data files from either:
    - a normal supported file, or
    - a ZIP containing one/many supported files.

    Everything remains in memory. No ZIP content is written to disk.
    """
    outer_name = uploaded_file.name
    outer_suffix = PurePosixPath(outer_name).suffix.lower()
    outer_bytes = uploaded_file.getvalue()

    if outer_suffix != ".zip":
        if outer_suffix not in SUPPORTED_DATA_EXTENSIONS:
            raise BenchmarkInputError(
                f"Format file '{outer_name}' belum didukung."
            )

        stream = io.BytesIO(outer_bytes)
        stream.name = outer_name

        yield {
            "stream": stream,
            "filename": outer_name,
            "display_name": outer_name,
            "raw_bytes": outer_bytes,
            "container": None,
        }
        return

    try:
        zf = zipfile.ZipFile(io.BytesIO(outer_bytes))
    except zipfile.BadZipFile as exc:
        raise BenchmarkInputError(
            f"ZIP '{outer_name}' corrupt atau bukan ZIP yang valid."
        ) from exc

    with zf:
        members = _validate_zip_members(
            zf,
            outer_name,
        )

        for info in members:
            try:
                member_bytes = zf.read(info)
            except Exception as exc:
                raise BenchmarkInputError(
                    f"Gagal membaca '{info.filename}' dari ZIP '{outer_name}': {exc}"
                ) from exc

            inner_name = PurePosixPath(
                info.filename.replace("\\", "/")
            ).name

            stream = io.BytesIO(member_bytes)
            stream.name = inner_name

            yield {
                "stream": stream,
                "filename": inner_name,
                "display_name": f"{outer_name} → {info.filename}",
                "raw_bytes": member_bytes,
                "container": outer_name,
            }


st.set_page_config(
    page_title="Benchmark Brand Superstar",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Benchmark Brand Superstar")
st.caption(f"Benchmark Engine v{ENGINE_VERSION}")

with st.expander("Logic Benchmark", expanded=False):
    st.markdown(
        """
        - **Shopee:** GMV exclude batal − Seller Rebate
        - **TikTok:** GMV exclude batal + Platform Discount
        - **Tokopedia:** GMV exclude batal + Platform Discount
        - **Lazada:** GMV exclude batal + Platform Discount
        - **Final:** jumlah seluruh marketplace per bulan → average seluruh bulan valid
        """
    )

st.info(
    "Zero Mistake Mode aktif: sistem akan melakukan reconciliation, parsing checks, "
    "order consistency checks, dan cross-file duplicate checks sebelum benchmark diloloskan."
)

st.subheader("Upload Data Pesanan")
st.caption(
    "Bisa upload file biasa, single ZIP, multiple ZIP, atau kombinasi keduanya. "
    "Isi ZIP diproses langsung di RAM tanpa diextract ke storage."
)

col1, col2 = st.columns(2)

with col1:
    sho_files = st.file_uploader(
        "Shopee Indonesia",
        type=["xlsx", "xls", "xlsm", "csv", "txt", "zip"],
        accept_multiple_files=True,
        key="sho",
    )

    tik_files = st.file_uploader(
        "TikTok Indonesia",
        type=["xlsx", "xls", "xlsm", "csv", "txt", "zip"],
        accept_multiple_files=True,
        key="tik",
    )

with col2:
    tok_files = st.file_uploader(
        "Tokopedia Indonesia",
        type=["xlsx", "xls", "xlsm", "csv", "txt", "zip"],
        accept_multiple_files=True,
        key="tok",
    )

    laz_files = st.file_uploader(
        "Lazada Indonesia",
        type=["xlsx", "xls", "xlsm", "csv", "txt", "zip"],
        accept_multiple_files=True,
        key="laz",
    )

uploaded_groups = {
    "SHO": sho_files or [],
    "TIK": tik_files or [],
    "TOK": tok_files or [],
    "LAZ": laz_files or [],
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

calculate = st.button(
    "Validasi & Hitung Benchmark",
    type="primary",
    disabled=selected_uploads == 0,
)

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

                            monthly = process_dataframe(
                                df,
                                marketplace,
                            )

                            monthly_results.append(
                                monthly
                            )

                            file_audit = validate_file_calculation(
                                df,
                                marketplace,
                                monthly,
                            )

                            for item in file_audit:
                                audit_rows.append(
                                    {
                                        "Scope": "File",
                                        "Marketplace": marketplace,
                                        "File": display_name,
                                        **item,
                                    }
                                )

                            file_identity.append(
                                {
                                    "marketplace": marketplace,
                                    "file": display_name,
                                    "sha256": file_hash,
                                    "ids": extract_record_ids(
                                        df,
                                        marketplace,
                                    ),
                                    "months": set(
                                        monthly["month"]
                                        .dropna()
                                        .tolist()
                                    ),
                                }
                            )

                            processed_info.append(
                                {
                                    "Status": "✅ OK",
                                    "Marketplace": marketplace,
                                    "File": display_name,
                                    "Rows": len(df),
                                    "Header Row": header_row + 1,
                                    "Periode Awal": (
                                        monthly["month"].min().strftime("%b %Y")
                                        if not monthly.empty
                                        else "-"
                                    ),
                                    "Periode Akhir": (
                                        monthly["month"].max().strftime("%b %Y")
                                        if not monthly.empty
                                        else "-"
                                    ),
                                }
                            )

                        except Exception as exc:
                            message = str(exc)

                            file_errors.append(
                                {
                                    "Marketplace": marketplace,
                                    "File": display_name,
                                    "Error": message,
                                }
                            )

                            processed_info.append(
                                {
                                    "Status": "❌ ERROR",
                                    "Marketplace": marketplace,
                                    "File": display_name,
                                    "Rows": "-",
                                    "Header Row": "-",
                                    "Periode Awal": "-",
                                    "Periode Akhir": "-",
                                }
                            )

                except Exception as exc:
                    # ZIP/container-level error: corrupt ZIP, nested ZIP,
                    # empty ZIP, excessive size/compression, etc.
                    message = str(exc)

                    file_errors.append(
                        {
                            "Marketplace": marketplace,
                            "File": uploaded_file.name,
                            "Error": message,
                        }
                    )

                    processed_info.append(
                        {
                            "Status": "❌ ERROR",
                            "Marketplace": marketplace,
                            "File": uploaded_file.name,
                            "Rows": "-",
                            "Header Row": "-",
                            "Periode Awal": "-",
                            "Periode Akhir": "-",
                        }
                    )

    st.subheader("Validasi File")

    st.dataframe(
        pd.DataFrame(processed_info),
        hide_index=True,
        width="stretch",
    )

    if file_errors:
        st.error(
            f"{len(file_errors)} file gagal divalidasi. "
            "Benchmark belum dihitung agar hasil tidak parsial."
        )

        for error in file_errors:
            st.error(
                f"**{error['Marketplace']} — {error['File']}**\n\n"
                f"{error['Error']}"
            )

        st.stop()

    # ---------------------------------------------------------
    # Cross-file Zero Mistake checks
    # ---------------------------------------------------------
    hash_groups = {}
    for item in file_identity:
        hash_groups.setdefault(item["sha256"], []).append(item)

    duplicate_file_groups = [
        items
        for items in hash_groups.values()
        if len(items) > 1
    ]

    if duplicate_file_groups:
        for items in duplicate_file_groups:
            labels = [
                f"{x['marketplace']} — {x['file']}"
                for x in items
            ]
            audit_rows.append(
                {
                    "Scope": "Cross-file",
                    "Marketplace": "ALL",
                    "File": " | ".join(labels),
                    "Status": "FAIL",
                    "Check": "File identik ter-upload lebih dari sekali",
                    "Detail": "SHA-256 file sama persis. Jika diteruskan, transaksi akan double count.",
                    "Actual": len(items),
                    "Expected": 1,
                }
            )
    else:
        audit_rows.append(
            {
                "Scope": "Cross-file",
                "Marketplace": "ALL",
                "File": "-",
                "Status": "PASS",
                "Check": "Tidak ada file identik",
                "Detail": "Tidak ditemukan file dengan isi byte yang sama persis.",
                "Actual": 0,
                "Expected": 0,
            }
        )

    overlap_details = []
    for marketplace in ["SHO", "TIK", "TOK", "LAZ"]:
        mp_files = [
            x for x in file_identity
            if x["marketplace"] == marketplace
        ]

        for left, right in itertools.combinations(mp_files, 2):
            overlap = left["ids"] & right["ids"]
            if overlap:
                overlap_details.append(
                    (
                        marketplace,
                        left["file"],
                        right["file"],
                        len(overlap),
                        list(overlap)[:5],
                    )
                )

    if overlap_details:
        for mp, left_file, right_file, count, samples in overlap_details:
            audit_rows.append(
                {
                    "Scope": "Cross-file",
                    "Marketplace": mp,
                    "File": f"{left_file} ↔ {right_file}",
                    "Status": "FAIL",
                    "Check": "Order overlap antar-file",
                    "Detail": f"Identifier transaksi yang sama muncul di dua file. Contoh: {samples}",
                    "Actual": count,
                    "Expected": 0,
                }
            )
    else:
        audit_rows.append(
            {
                "Scope": "Cross-file",
                "Marketplace": "ALL",
                "File": "-",
                "Status": "PASS",
                "Check": "Tidak ada Order ID overlap antar-file",
                "Detail": "Tidak ditemukan identifier transaksi yang sama pada dua file marketplace yang sama.",
                "Actual": 0,
                "Expected": 0,
            }
        )

    try:
        result = calculate_benchmark(
            monthly_results
        )

        final_audit = validate_final_result(result)
        for item in final_audit:
            audit_rows.append(
                {
                    "Scope": "Final",
                    "Marketplace": "ALL",
                    "File": "-",
                    **item,
                }
            )

        # ---------------------------------------------------------
        # Zero Mistake Audit UI
        # ---------------------------------------------------------
        st.subheader("Zero Mistake Audit")

        audit_df = pd.DataFrame(audit_rows)

        pass_count = int((audit_df["Status"] == "PASS").sum())
        warning_count = int((audit_df["Status"] == "WARNING").sum())
        fail_count = int((audit_df["Status"] == "FAIL").sum())

        qa1, qa2, qa3 = st.columns(3)
        qa1.metric("PASS", pass_count, border=True)
        qa2.metric("WARNING", warning_count, border=True)
        qa3.metric("FAIL", fail_count, border=True)

        with st.expander(
            "Lihat seluruh hasil audit",
            expanded=fail_count > 0 or warning_count > 0,
        ):
            st.dataframe(
                audit_df,
                hide_index=True,
                width="stretch",
            )

        audit_csv = audit_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Download Validation Report (.csv)",
            data=audit_csv,
            file_name="benchmark_validation_report.csv",
            mime="text/csv",
        )

        if fail_count > 0:
            st.error(
                f"Zero Mistake Audit menemukan {fail_count} FAIL. "
                "Benchmark tidak diloloskan sebagai hasil final sampai seluruh FAIL diselesaikan."
            )
            st.stop()

        if warning_count > 0:
            st.warning(
                f"Perhitungan reconcile tanpa FAIL, tetapi ada {warning_count} WARNING yang perlu direview."
            )
        else:
            st.success(
                "Zero Mistake Audit PASS: seluruh reconciliation utama menghasilkan selisih Rp0 dan tidak ada konflik kritis."
            )

        st.success(
            "Semua file valid, audit kritis lolos, dan benchmark berhasil dihitung."
        )

        st.subheader("Hasil Benchmark")
        st.caption(
            "Metode perhitungan: **Exclude Batal** — transaksi berstatus batal/cancel "
            "tidak dimasukkan ke GMV benchmark sesuai logic masing-masing marketplace."
        )

        m1, m2, m3, m4 = st.columns(4)

        m1.metric(
            "Benchmark",
            f"Rp {result['benchmark']:,.0f}".replace(",", "."),
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

        processed_file_count = sum(
            1
            for item in processed_info
            if item["Status"] == "✅ OK"
        )

        m4.metric(
            "File Diproses",
            f"{processed_file_count}",
            border=True,
        )

        for warning in result["warnings"]:
            st.warning(warning)

        # ---------------------------------------------------------
        # Ringkasan Benchmark
        # ---------------------------------------------------------
        st.subheader("Ringkasan Benchmark")
        st.caption(
            "Seluruh nilai di bawah merupakan **Benchmark Exclude Batal**."
        )

        valid_months = result["valid_months"]
        monthly_combined = result["monthly_combined"]

        marketplace_columns = {
            "SHO": "shopee",
            "TIK": "tiktok",
            "TOK": "tokopedia",
            "LAZ": "lazada",
        }

        summary_rows = []

        # Kontribusi average masing-masing marketplace memakai
        # denominator bulan yang sama dengan Final Benchmark.
        # Dengan begitu jumlah SHO + TIK + TOK + LAZ = Final Benchmark.
        for marketplace in ["SHO", "TIK", "TOK", "LAZ"]:
            if marketplace not in active_marketplaces:
                continue

            column = marketplace_columns[marketplace]

            marketplace_avg = (
                monthly_combined[column].sum()
                / valid_months
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

        # Bold baris final + format nominal dengan pemisah ribuan.
        final_row_index = len(summary_df) - 1

        summary_styled = (
            summary_df.style
            .format(
                {
                    "Nilai": "{:,.0f}",
                }
            )
            .apply(
                lambda row: [
                    "font-weight: bold"
                    if row.name == final_row_index
                    else ""
                    for _ in row
                ],
                axis=1,
            )
        )

        st.dataframe(
            summary_styled,
            hide_index=True,
            width="stretch",
        )

        st.subheader("Benchmark per Bulan")

        monthly_display = (
            result["monthly_combined"]
            .copy()
        )

        monthly_display["month"] = (
            monthly_display["month"]
            .dt.strftime("%B %Y")
        )

        monthly_display = monthly_display.rename(
            columns={
                "month": "Bulan",
                "shopee": "Shopee",
                "tiktok": "TikTok",
                "tokopedia": "Tokopedia",
                "lazada": "Lazada",
                "total_gmv": "Total GMV",
            }
        )

        st.dataframe(
            monthly_display,
            hide_index=True,
            width="stretch",
        )

        st.subheader(
            "Detail Perhitungan Marketplace"
        )

        with st.expander(
            "ℹ️ Rincian sumber perhitungan tiap marketplace",
            expanded=False,
        ):
            st.markdown(
                """
| Marketplace | GMV All | GMV Exclude Batal | Platform Discount | Seller Rebate | GMV Benchmark |
|---|---|---|---|---|---|
| **SHO** | `Subtotal Pesanan` seluruh data | `Subtotal Pesanan` untuk pesanan dengan `Status Pesanan` bukan Batal | Tidak digunakan (`0`) | `Voucher Ditanggung Penjual` + `Paket Diskon (Diskon dari Penjual)`, dihitung per `No. Pesanan` agar tidak double count | `GMV Exclude Batal - Seller Rebate` |
| **TIK** | `SKU Subtotal After Discount` dari Order ID valid yang masuk ke helper Exclude Batal | Total `SKU Subtotal After Discount` per `Order ID` yang `Order Status` tidak mengandung BATAL/CANCEL dan `Cancelation/Return Type` tidak mengandung CANCEL | Total `SKU Platform Discount` dari Order ID valid yang sama | Tidak digunakan (`0`) | `GMV Exclude Batal + Platform Discount` |
| **TOK** | `SKU Subtotal After Discount` dari Order ID valid yang masuk ke helper Exclude Batal | Total `SKU Subtotal After Discount` per `Order ID` yang `Order Status` tidak mengandung BATAL/CANCEL dan `Cancelation/Return Type` tidak mengandung CANCEL | Total `SKU Platform Discount` dari Order ID valid yang sama | Tidak digunakan (`0`) | `GMV Exclude Batal + Platform Discount` |
| **LAZ** | `unitPrice` seluruh data | `unitPrice` untuk baris dengan `status` bukan `canceled/cancelled` | `platformDiscountTotal` untuk baris yang bukan `canceled/cancelled` | Tidak digunakan (`0`) | `GMV Exclude Batal + Platform Discount` |
                """
            )

            st.caption(
                "Catatan: untuk TIK/TOK, perhitungan mengikuti helper "
                "'Breakdown (Exc Batal ID)' pada template benchmark lama, "
                "sehingga agregasi dilakukan per Order ID terlebih dahulu sebelum dijumlahkan per bulan."
            )

        detail = (
            result["monthly_marketplace"]
            .copy()
        )

        detail["month"] = (
            detail["month"]
            .dt.strftime("%B %Y")
        )

        detail = detail.rename(
            columns={
                "marketplace": "Marketplace",
                "month": "Bulan",
                "gmv_all": "GMV All",
                "gmv_ex_cancel": "GMV Exclude Batal",
                "platform_discount": "Platform Discount",
                "seller_rebate": "Seller Rebate",
                "benchmark_gmv": "GMV Benchmark",
                "source_rows": "Source Rows",
            }
        )

        st.dataframe(
            detail,
            hide_index=True,
            width="stretch",
        )

    except BenchmarkInputError as exc:
        st.error(str(exc))


