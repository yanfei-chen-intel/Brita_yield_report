"""
SCN Yield Analysis Script
Reads units_batch_summary.csv and computes per-indicator actual yield.
"""

import argparse
import sys
from pathlib import Path

import pandas as pd


def compute_yield(input_path: str) -> None:
    input_file = Path(input_path)
    if not input_file.exists():
        print(f"Error: input file not found: {input_file}")
        sys.exit(1)

    df = pd.read_csv(input_file)

    # Normalize column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Step 2: keep only Level==3 and Result in ['DFT', 'PASS']
    df = df[df["Level"] == 3]
    df = df[df["Result"].isin(["DFT", "PASS"])]

    # Step 3: split Indicator by '::' into four new columns
    split_cols = df["Indicator"].str.split("::", expand=True)
    split_cols.columns = ["Module", "RG", "EID", "FLOPCOUNT"]
    df = pd.concat([df.reset_index(drop=True), split_cols.reset_index(drop=True)], axis=1)

    # Step 4: pivot so each indicator row carries both PASS and DFT values.
    # Group keys that uniquely identify one indicator instance (per lot/wafer).
    group_keys = ["LOT", "WAFER", "Indicator", "Module", "RG", "EID", "FLOPCOUNT"]

    pass_df = (
        df[df["Result"] == "PASS"][group_keys + ["Die_Tested", "Die_Per_Wafer"]]
        .rename(columns={"Die_Tested": "pass_value", "Die_Per_Wafer": "die_per_wafer_value"})
    )

    dft_df = (
        df[df["Result"] == "DFT"][group_keys + ["Die_Per_Wafer"]]
        .rename(columns={"Die_Per_Wafer": "dft_value"})
    )

    merged = pd.merge(pass_df, dft_df, on=group_keys, how="inner")

    # Step 4d-e: compute actual_tested_die_count and actual_yield
    merged["actual_tested_die_count"] = merged["pass_value"] - merged["dft_value"]
    merged["actual_yield"] = merged["die_per_wafer_value"] / merged["actual_tested_die_count"]

    # Step 4f: build output with required columns
    output_cols = [
        "LOT", "WAFER",
        "Indicator", "Module", "RG", "EID", "FLOPCOUNT",
        "pass_value", "dft_value", "die_per_wafer_value",
        "actual_tested_die_count", "actual_yield",
    ]
    result = merged[output_cols]

    output_file = input_file.parent / "SCN_yield_result.csv"
    result.to_csv(output_file, index=False)
    print(f"Done. {len(result)} rows written to: {output_file}")


def main():
    parser = argparse.ArgumentParser(description="Compute SCN indicator actual yield.")
    parser.add_argument(
        "input",
        nargs="?",
        default=r".\output\SCN_GT\units_batch_summary.csv",
        help="Path to units_batch_summary.csv (default: .\\output\\SCN_GT\\units_batch_summary.csv)",
    )
    args = parser.parse_args()
    compute_yield(args.input)


if __name__ == "__main__":
    main()
