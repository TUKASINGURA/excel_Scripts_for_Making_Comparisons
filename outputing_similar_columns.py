import pandas as pd
import csv
import re

# --- INPUT FILE (single file mode) ---
input_file = "filtered_by_date_original_text.xlsx"
sheet = "Sheet1"

# --- COMPARISON MODE ---
compare_mode = "single_column"   # Set to "single_column" to focus on one column

# --- SINGLE-COLUMN MODE ---
single_column = "subscriptions_packages_id"            # Column to look for repeated / similar values
single_min_count = 1                # Minimum occurrences to consider "similar" (set to 1 to include all)
single_filter_value = "18"     # Only consider values containing this substring
single_filter_regex = False         # Treat single_filter_value as regex if True

# --- OPTIONAL: output only these columns (None => keep full matched rows) ---
output_columns = None      # Set to None to save full matched rows

# --- HELPER: detect CSV delimiter ---
def detect_sep(path):
    try:
        with open(path, "r", encoding="utf-8", newline="") as f:
            sample = f.read(4096)
            return csv.Sniffer().sniff(sample).delimiter
    except Exception:
        return ","

# --- READ INPUT (CSV or Excel) ---
def read_table(path, sheet=None):
    lower = path.lower()
    if lower.endswith(('.xls', '.xlsx', '.xlsm')):
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", dtype=str)
    else:
        sep = detect_sep(path)
        df = pd.read_csv(path, sep=sep, engine="python", encoding="utf-8", dtype=str, skipinitialspace=True)
    df.columns = df.columns.astype(str).str.replace("\ufeff", "").str.strip()
    return df

# --- LOAD DATA ---
df = read_table(input_file, sheet=sheet)
print(f"{input_file} columns:", list(df.columns))

# --- SINGLE COLUMN MODE ---
if single_column not in df.columns:
    raise KeyError(f"Column '{single_column}' not found in {input_file}. Available: {list(df.columns)}")

# Normalize the column values for reliable grouping
series_full = df[single_column].fillna("").astype(str).str.strip()

# Filter by substring (e.g., "android")
mask_filter = series_full.str.contains(single_filter_value, case=False, na=False, regex=False)
df_masked = df[mask_filter].copy()

# Find all values in the filtered DataFrame
matches = df_masked.copy()

# --- SELECT OUTPUT COLUMNS ---
if output_columns:
    missing = [c for c in output_columns if c not in matches.columns]
    if missing:
        raise KeyError(f"Requested output columns not present: {missing}")
    out_df = matches[output_columns].copy()
else:
    out_df = matches.copy()

# --- SAVE RESULTS ---
output_file = f"matched_rows_containing_{single_filter_value}.xlsx"
out_df.to_excel(output_file, index=False, engine="openpyxl")
print(f"Done! {len(out_df)} rows saved to: {output_file}")
