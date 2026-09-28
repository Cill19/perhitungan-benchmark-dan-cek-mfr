from pathlib import Path
import pandas as pd
from thailand_tiktok_engine import parse_tiktok_xlsx, process_tiktok_th

def run_test():
    tik_dir = Path("/Users/fahroni/Downloads/web app benchmark /TIK")
    if not tik_dir.exists() or not list(tik_dir.glob("*.xlsx")):
        tik_dir = Path("/Users/fahroni/Downloads/TH.Hibal/TIK")
    files = sorted(tik_dir.glob("*.xlsx"))

    if not files:
        print("No TikTok TH files found in /TIK")
        return

    frames = []
    total_rows = 0
    num_cols = 0

    for f in files:
        df_raw = parse_tiktok_xlsx(f)
        total_rows += len(df_raw)
        num_cols = len(df_raw.columns)
        
        m_result = process_tiktok_th(df_raw)
        frames.append(m_result)

    combined = pd.concat(frames, ignore_index=True)
    
    # Combine monthly benchmark
    monthly_combined = (
        combined
        .groupby("month", as_index=False)
        .agg({
            "gmv_all": "sum",
            "gmv_ex_cancel": "sum",
            "platform_discount": "sum",
            "seller_rebate": "sum",
            "benchmark_gmv": "sum",
            "source_rows": "sum"
        })
    )

    valid_months = monthly_combined["month"].nunique()
    total_benchmark = monthly_combined["benchmark_gmv"].sum()
    benchmark_avg = total_benchmark / valid_months if valid_months else 0

    print("================================================")
    print("REAL TIKTOK TH TEST")
    print("================================================")
    print("\nRows:")
    print(total_rows)
    print("\nColumns:")
    print(num_cols)
    print("\n\nRESULT\n")

    display_df = monthly_combined[["month", "benchmark_gmv"]].copy()
    display_df["month"] = display_df["month"].dt.strftime("%Y-%m")
    display_df["benchmark_gmv"] = display_df["benchmark_gmv"].apply(lambda x: f"฿{x:,.2f}")

    print(display_df.to_string(index=False))

    print("\n\nTOTAL BENCHMARK THB:")
    print(f"฿{total_benchmark:,.2f}")
    print(f"\nBENCHMARK AVG L{valid_months}M THB:")
    print(f"฿{benchmark_avg:,.2f}")

if __name__ == "__main__":
    run_test()
