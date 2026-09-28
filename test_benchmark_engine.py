import pandas as pd
from benchmark_engine import (
    process_shopee,
    process_tik_family,
    process_lazada,
    calculate_benchmark,
    validate_file_calculation,
    validate_final_result,
)

# SHO
sho = pd.DataFrame({
    "No. Pesanan": ["A", "B", "C", "C"],
    "Status Pesanan": ["Selesai", "Batal", "Selesai", "Selesai"],
    "Waktu Pesanan Dibuat": [
        "2026-01-01 10:00", "2026-01-02 10:00",
        "2026-02-01 10:00", "2026-02-01 10:00"
    ],
    "Subtotal Pesanan": ["100.000", "200.000", "150.000", "50.000"],
    "Voucher Ditanggung Penjual": ["10.000", "20.000", "5.000", "5.000"],
    "Diskon Dari Shopee": ["5.000", "10.000", "3.000", "2.000"],
})
sho_m = process_shopee(sho)

jan = sho_m[sho_m["month"] == pd.Timestamp("2026-01-01")].iloc[0]
assert jan["gmv_ex_cancel"] == 100000
assert jan["seller_rebate"] == 10000
assert jan["platform_discount"] == 5000
assert jan["benchmark_gmv"] == 95000

feb = sho_m[sho_m["month"] == pd.Timestamp("2026-02-01")].iloc[0]
# order C rebate deduplicated (max per order = 5000)
assert feb["gmv_ex_cancel"] == 200000
assert feb["seller_rebate"] == 5000
assert feb["platform_discount"] == 5000
assert feb["benchmark_gmv"] == 200000

sho_audit = validate_file_calculation(sho, "SHO", sho_m)
assert all(r["Status"] == "PASS" for r in sho_audit), f"SHO audit failed: {sho_audit}"

# TIK
tik = pd.DataFrame({
    "Order ID": ["T1", "T2", "T3"],
    "Order Sub Status": ["Selesai", "Dibatalkan", "Selesai"],
    "Created Time": [
        "01/01/2026 10:00:00",
        "02/01/2026 10:00:00",
        "01/02/2026 10:00:00",
    ],
    "SKU Subtotal After Discount": [50000, 25000, 100000],
    "SKU Platform Discount": [10000, 5000, 20000],
})
tik_m = process_tik_family(tik, "TIK")
jan_tik = tik_m[tik_m["month"] == pd.Timestamp("2026-01-01")].iloc[0]
assert jan_tik["gmv_ex_cancel"] == 50000
assert jan_tik["platform_discount"] == 10000
assert jan_tik["benchmark_gmv"] == 60000

tik_audit = validate_file_calculation(tik, "TIK", tik_m)
assert all(r["Status"] == "PASS" for r in tik_audit), f"TIK audit failed: {tik_audit}"

# LAZ
laz = pd.DataFrame({
    "orderItemId": ["L1", "L2", "L3"],
    "createTime": ["01/01/2026 10:00:00", "02/01/2026 10:00:00", "01/03/2026 10:00:00"],
    "unitPrice": [100000, 200000, 300000],
    "platformDiscountTotal": [10000, 20000, 30000],
    "status": ["confirmed", "canceled", "confirmed"],
})
laz_m = process_lazada(laz)
jan_laz = laz_m[laz_m["month"] == pd.Timestamp("2026-01-01")].iloc[0]
assert jan_laz["benchmark_gmv"] == 110000

# Combined monthly logic:
# Jan = SHO 95k + TIK 60k + LAZ 110k = 265k
# Feb = SHO 200k + TIK 120k = 320k
# Mar = LAZ 330k
# Benchmark = (265 + 320 + 330) / 3 = 305k
result = calculate_benchmark([sho_m, tik_m, laz_m])
assert result["valid_months"] == 3
assert result["benchmark"] == 305000

final_audit = validate_final_result(result)
assert all(r["Status"] in {"PASS", "WARNING"} for r in final_audit), f"Final audit failed: {final_audit}"

print("All synthetic tests and audits passed!")
print(result["monthly_combined"].to_string(index=False))
print("Benchmark:", result["benchmark"])
