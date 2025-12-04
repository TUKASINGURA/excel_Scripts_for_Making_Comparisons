import pandas as pd
import csv

# Purpose:
# Read the input file (CSV or Excel), remove duplicate rows based on one or more
# specified columns, and save the deduplicated result. No date filtering is performed.

# --- INPUT FILE ---
file_path = "filtered_by_date_original_text.xlsx"
sheet_name = "Sheet1"

# --- DEDUPLICATION CONFIG ---
# List the column names to deduplicate by (example: ["User"] to keep one "Richard" row)
# Set to [] or None to disable deduplication.
dedup_columns = ["devices_customers_id"]

# Keep strategy for duplicates: "first", "last", or False/None to disable.
dedup_keep = "first"

# If True perform case-insensitive comparison for deduplication.
dedup_case_insensitive = True

# --- HELPER: detect delimiter for CSV ---
def detect_sep(path):
    try:
        with open(path, "r", encoding="utf-8", newline="") as f:
            sample = f.read(4096)
            return csv.Sniffer().sniff(sample).delimiter
    except Exception:
        return ","

# --- HELPER: read Excel or CSV as raw strings (preserve original cell text) ---
def read_table_as_strings(path, sheet=None):
    """
    Read the file and return a DataFrame where ALL values are read as strings.
    Try several encodings for CSV input (utf-8, utf-8-sig, cp1252, latin-1).
    Fall back to reading with errors='replace' if decode fails.
    """
    lower = path.lower()
    if lower.endswith(('.xls', '.xlsx', '.xlsm')):
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", dtype=str)
        df.columns = df.columns.astype(str).str.replace("\ufeff", "").str.strip()
        return df

    sep = detect_sep(path)
    encodings = ['utf-8', 'utf-8-sig', 'cp1252', 'latin-1']
    read_kwargs = dict(
        sep=sep,
        engine="python",
        dtype=str,
        skipinitialspace=True,
        keep_default_na=False,
        na_filter=False
    )
    for enc in encodings:
        try:
            df = pd.read_csv(path, encoding=enc, **read_kwargs)
            df.columns = df.columns.astype(str).str.replace("\ufeff", "").str.strip()
            return df
        except UnicodeDecodeError:
            # try next encoding
            continue

    # Last-resort: read with replacement of invalid bytes so load never fails
    df = pd.read_csv(path, encoding='utf-8', errors='replace', **read_kwargs)
    df.columns = df.columns.astype(str).str.replace("\ufeff", "").str.strip()
    return df

# --- LOAD THE TABLE (raw strings) ---
df_raw = read_table_as_strings(file_path, sheet=sheet_name)
print("Columns:", list(df_raw.columns))
print(f"Total rows before deduplication: {len(df_raw)}")

# --- REMOVE DUPLICATES (if configured) ---
if dedup_columns:
    missing = [c for c in dedup_columns if c not in df_raw.columns]
    if missing:
        raise KeyError(f"Dedup columns not found in data: {missing}")

    # create temporary normalized columns for comparison
    tmp_cols = []
    for c in dedup_columns:
        tmp = f"__dedup_tmp__{c}"
        if dedup_case_insensitive:
            df_raw[tmp] = df_raw[c].fillna("").astype(str).str.strip().str.lower()
        else:
            df_raw[tmp] = df_raw[c].fillna("").astype(str).str.strip()
        tmp_cols.append(tmp)

    before_count = len(df_raw)
    df_deduped = df_raw.drop_duplicates(subset=tmp_cols, keep=dedup_keep).copy()

    # remove temporary helper columns
    df_deduped.drop(columns=tmp_cols, inplace=True, errors="ignore")

    removed = before_count - len(df_deduped)
    print(f"Removed {removed} duplicate rows based on columns {dedup_columns} (keep='{dedup_keep}')")
else:
    df_deduped = df_raw.copy()
    print("Deduplication disabled (dedup_columns is empty).")

# --- SAVE RESULT ---
output_file = "deduped_output.xlsx"
df_deduped.to_excel(output_file, index=False, engine="openpyxl")

print(f"Done! {len(df_deduped)} rows saved to: {output_file}")
