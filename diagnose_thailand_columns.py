import pandas as pd


FILE_PATH = "data pesanan sho th.xlsx"


df = pd.read_excel(
    FILE_PATH,
    sheet_name="orders",
    header=0
)


print("="*80)
print("COLUMN MAPPING")
print("="*80)


for index, col in enumerate(df.columns, start=1):

    print(
        index,
        "→",
        col
    )