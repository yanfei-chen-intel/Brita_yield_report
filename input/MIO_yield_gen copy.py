"""
MIO_yield_gen.py
Reads ttd_result.csv and produces MIO_yield.csv with per-test-instance yield stats.
Test instance columns = all columns ending with 'P/F' (values: 1=pass, 0=fail, NaN=untested).

Arguments:
  -output_dir   Path to the task output directory (e.g. output\\MIO). Required.
  -filter_dict  JSON string mapping category keys to lists of allowed condition values.
                Rows whose Name contains both a key AND one of its values are kept.
                Filtered output is saved to <output_dir>/yield.csv.
"""

import argparse
import json
import pathlib
import pandas as pd

# ---------------------------------------------------------------------------
# Argument parsing
# ---------------------------------------------------------------------------
parser = argparse.ArgumentParser()
parser.add_argument("-output_dir", type=str, required=True,
                    help="Path to the task output directory")
parser.add_argument("-filter_dict", type=str, default=None,
                    help="JSON string: {key: [values]} filter applied to Name column")
args, _ = parser.parse_known_args()

filter_dict = json.loads(args.filter_dict) if args.filter_dict else {}

output_dir  = pathlib.Path(args.output_dir)
INPUT_FILE  = output_dir / "ttd_result.csv"
OUTPUT_FILE = output_dir / "MIO_yield.csv"

# ---------------------------------------------------------------------------
# Load data
# ---------------------------------------------------------------------------
df = pd.read_csv(INPUT_FILE)
total = len(df)   # data rows (header excluded by pandas)

# ---------------------------------------------------------------------------
# Material  →  SORT_LOT_wafer1_wafer2_...
# ---------------------------------------------------------------------------
lots = sorted(df["SORT_LOT"].dropna().unique().tolist())
lot_str = "_".join(str(l) for l in lots)

wafers = sorted(df["SORT_WAFER"].dropna().astype(int).unique().tolist())
wafer_str = "_".join(str(w) for w in wafers)

material = lot_str + "_" + wafer_str

# ---------------------------------------------------------------------------
# TestProgram  →  deduplicated program names across all slot columns
# ---------------------------------------------------------------------------
program_cols = [c for c in df.columns if c.startswith("Program Name_")]
all_programs = set()
for col in program_cols:
    all_programs.update(df[col].dropna().unique().tolist())
test_program = ", ".join(sorted(all_programs))

# ---------------------------------------------------------------------------
# Test instance columns  →  all columns ending with 'P/F'
# (Port columns contain socket numbers, not pass/fail, so are excluded)
# ---------------------------------------------------------------------------
pf_cols = [c for c in df.columns if c.endswith("P/F")]

# ---------------------------------------------------------------------------
# Build output rows
# ---------------------------------------------------------------------------
rows = []
for col in pf_cols:
    series = df[col]
    n_pass  = int((series == 1).sum())
    n_fail  = int((series == 0).sum())
    n_unk   = int(series.isna().sum())

    # Sanity check
    if n_pass + n_fail + n_unk != total:
        print(f"[WARN] {col}: #P({n_pass}) + #F({n_fail}) + #U({n_unk}) "
              f"= {n_pass+n_fail+n_unk} ≠ Total({total})")

    denom   = total - n_unk
    pct_pass = round(n_pass / denom * 100, 4) if denom > 0 else float("nan")

    # Name: strip "_P/F", split by "/", take last field
    stripped = col[:-4] if col.endswith("_P/F") else col   # remove trailing _P/F
    name = stripped.split("/")[-1]

    # Module: from original col, split by "/", take 3rd field (index 2),
    #         then take everything before the last "_"
    parts = col.split("/")
    field3 = parts[2] if len(parts) > 2 else parts[-1]
    module = field3.rsplit("_", 1)[0] if "_" in field3 else field3

    rows.append({
        "Material":    material,
        "TestProgram": test_program,
        "Name":        name,
        "Module":      module,
        "Total":       total,
        "#P":          n_pass,
        "#F":          n_fail,
        "#U":          n_unk,
        "%P":          pct_pass,
    })

# ---------------------------------------------------------------------------
# Save output
# ---------------------------------------------------------------------------
out_df = pd.DataFrame(rows, columns=["Material", "TestProgram", "Name", "Module",
                                      "Total", "#P", "#F", "#U", "%P"])
out_df.to_csv(OUTPUT_FILE, index=False)
print(f"Done. {len(out_df)} test instances -> {OUTPUT_FILE}")
print(f"  Material:    {material}")
print(f"  TestProgram: {test_program}")
print(f"  Total rows:  {total}")

# ---------------------------------------------------------------------------
# Filtered yield output  (only when -filter_dict is provided)
# Row is kept if Name contains a filter key AND one of its associated values.
# ---------------------------------------------------------------------------
if filter_dict:
    def _matches(name: str) -> bool:
        for key, values in filter_dict.items():
            if key in name and any(v in name for v in values):
                return True
        return False

    filtered_df = out_df[out_df["Name"].apply(_matches)].reset_index(drop=True)
    yield_file  = pathlib.Path(OUTPUT_FILE).parent / "MIO_filtered_yield.csv"
    filtered_df.to_csv(yield_file, index=False)
    print(f"Filtered {len(filtered_df)} / {len(out_df)} rows -> {yield_file}")
