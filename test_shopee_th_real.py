import io
import zipfile
from pathlib import Path

import pandas as pd

from thailand_engine import process_shopee_th


FOLDER_PATH = Path(
    "SHO"
)


SUPPORTED_EXT = {
    ".xlsx",
    ".xls",
    ".xlsm",
}


DATE_COLUMN = "วันที่ทำการสั่งซื้อ"


def read_excel_stream(stream):

    return pd.read_excel(
        stream,
        sheet_name="orders",
        header=0
    )



def diagnose_date(df, filename):

    print("\n" + "-" * 80)
    print("DATE DIAGNOSTIC:", filename)
    print("-" * 80)


    if DATE_COLUMN not in df.columns:

        print(
            f"Column {DATE_COLUMN} tidak ditemukan"
        )

        return


    print("\nRAW SAMPLE:")
    print(
        df[DATE_COLUMN]
        .head(5)
        .to_string()
    )


    parsed_date = pd.to_datetime(
        df[DATE_COLUMN],
        errors="coerce"
    )


    print("\nPARSED MIN:")
    print(
        parsed_date.min()
    )


    print("\nPARSED MAX:")
    print(
        parsed_date.max()
    )


    print("\nMONTH DISTRIBUTION:")

    print(
        parsed_date
        .dt
        .to_period("M")
        .value_counts()
        .sort_index()
    )



def process_file(df, filename):

    print("\nPROCESS FILE:")
    print(filename)

    print(
        "Rows:",
        len(df)
    )


    diagnose_date(
        df,
        filename
    )


    result = process_shopee_th(
        df
    )


    print("\nENGINE RESULT:")
    print(result)


    return result



def load_folder_files(folder):

    all_results = []


    for file in sorted(folder.iterdir()):


        # ======================================
        # EXCEL
        # ======================================

        if file.suffix.lower() in SUPPORTED_EXT:


            df = pd.read_excel(
                file,
                sheet_name="orders",
                header=0
            )


            result = process_file(
                df,
                file.name
            )


            all_results.append(
                result
            )



        # ======================================
        # ZIP
        # ======================================

        elif file.suffix.lower() == ".zip":


            with zipfile.ZipFile(file) as z:


                for name in z.namelist():


                    if not name.lower().endswith(
                        ".xlsx"
                    ):
                        continue



                    data = z.read(
                        name
                    )


                    stream = io.BytesIO(
                        data
                    )


                    df = read_excel_stream(
                        stream
                    )


                    result = process_file(
                        df,
                        f"{file.name} → {name}"
                    )


                    all_results.append(
                        result
                    )


    return all_results



def main():


    print("=" * 80)

    print(
        "REAL SHOPEE TH FOLDER TEST"
    )

    print("=" * 80)



    results = load_folder_files(
        FOLDER_PATH
    )


    combined = pd.concat(
        results,
        ignore_index=True
    )


    print("\n\nRAW RESULT PER FILE")

    print("=" * 80)

    print(
        combined
    )



    monthly = (

        combined

        .groupby(
            "month",
            as_index=False
        )

        ["benchmark_gmv"]

        .sum()

    )


    print("\n\nFINAL MONTHLY")

    print("=" * 80)

    print(
        monthly
    )


    print("\nTOTAL MONTH:")

    print(
        monthly["month"]
        .nunique()
    )


    print("\nAVERAGE BENCHMARK THB:")

    print(
        monthly["benchmark_gmv"]
        .mean()
    )



if __name__ == "__main__":

    main()