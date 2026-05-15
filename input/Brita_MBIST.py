import pandas as pd
import pathlib

input_file   = pathlib.Path(r".\output\MBIST\MBIST_units_batch_summary.csv")
output_file  = input_file.with_name(input_file.stem + "_filtered.csv")
agg_file     = input_file.with_name(input_file.stem + "_agg.csv")

df = pd.read_csv(input_file, dtype=str)

# Strip leading/trailing whitespace from column names (BOM-safe)
df.columns = df.columns.str.strip().str.lstrip("\ufeff")

filtered = df[
    df["Indicator"].str.endswith("::Y", na=False) &
    (df["Level"].str.strip() == "2") &
    (df["Result"].str.strip() == "PASS")
].copy()

filtered.to_csv(output_file, index=False)
print(f"Filtered {len(filtered)} row(s) → {output_file}")

# Aggregation: sum(%Impact_on_Tested_no_Untested) / wafer count per Indicator
impact_col = "%Impact_on_Tested_no_Untested"
filtered[impact_col] = (
    filtered[impact_col]
    .str.rstrip("%")
    .replace("NaN", None)
    .astype(float)
)

agg = (
    filtered.groupby("Indicator", sort=False)
    .apply(
        lambda g: pd.Series({
            "Wafer_Count":      g["WAFER"].nunique(),
            "Sum_%Impact":      g[impact_col].sum(),
            "Avg_%Impact":      g[impact_col].sum() / g["WAFER"].nunique(),
        }),
        include_groups=False,
    )
    .reset_index()
)

agg.to_csv(agg_file, index=False)
print(f"Aggregated {len(agg)} indicator(s) → {agg_file}")

