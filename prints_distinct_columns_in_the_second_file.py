import pandas as pd
import csv

# --- SCRIPT PURPOSE ---
# Read two tables (CSV or Excel), compare rows on the set of columns they share,
# and produce exactly two output files:
#  - rows_common_on_shared_columns.xlsx  -> rows from file2 that match some row in file1 on shared columns
#  - rows_distinct_on_shared_columns.xlsx -> rows from file2 that do not match any row in file1 on shared columns

# --- INPUT FILES ---
# file1: reference file (rows here are treated as "existing")
# file2: file to check (rows here will be classified as common or distinct)
file1 = "Actors.csv"
file2 = "title.csv"

# --- SHEETS (optional, used only for Excel files) ---
# If your inputs are Excel workbooks, set these to the sheet names to read.
sheet1 = "Sheet1"
sheet2 = "Sheet1"

# --- LOAD THE DATA ---
def detect_sep(path):
    """
    Try to detect CSV delimiter using csv.Sniffer on a sample of the file.
    If detection fails, return comma as a sensible default.
    This function is only used when the path is not an Excel file.
    """
    try:
        with open(path, "r", encoding="utf-8", newline="") as f:
            sample = f.read(4096)
            return csv.Sniffer().sniff(sample).delimiter
    except Exception:
        # Fallback delimiter
        return ","

def read_table(path, sheet=None):
    """
    Read a file as Excel (openpyxl) if it ends with an Excel extension,
    otherwise treat it as CSV and try to detect the separator.
    After reading, normalize the column names by removing BOM and trimming whitespace.
    """
    lower = path.lower()
    if lower.endswith(('.xls', '.xlsx', '.xlsm')):
        # Read Excel sheet using openpyxl engine (works for .xlsx)
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    else:
        # Read CSV with detected separator. skipinitialspace helps with spaces after delimiters.
        sep = detect_sep(path)
        df = pd.read_csv(path, sep=sep, engine="python", encoding="utf-8", skipinitialspace=True)
    # Normalize column names: remove BOM (\ufeff) and trim whitespace for reliable matching
    df.columns = df.columns.astype(str).str.replace("\ufeff", "").str.strip()
    return df

# Load both input tables (function handles Excel vs CSV automatically)
df1 = read_table(file1, sheet=sheet1)
df2 = read_table(file2, sheet=sheet2)

# --- COMMON / DISTINCT ON SHARED COLUMNS ---
# Determine which columns exist in both files; comparison uses only these shared columns.
common_cols = list(df1.columns.intersection(df2.columns))

if common_cols:
    def row_sig(df, cols):
        """
        Build a stable string signature for each row using the specified columns.
        - fillna("") avoids NaN vs empty-string mismatches
        - astype(str) ensures consistent type
        - strip whitespace from each value to avoid trivial differences
        - join with a separator unlikely to appear in actual data ("||")
        """
        return (
            df[cols]
            .fillna("")
            .astype(str)
            .apply(lambda r: "||".join(v.strip() for v in r.values), axis=1)
        )

    # create signature series for both dataframes using only shared columns
    sig1 = row_sig(df1, common_cols)
    sig2 = row_sig(df2, common_cols)

    # Rows in df2 whose signature appears in df1 are "common"
    common_rows = df2[sig2.isin(sig1)]

    # Rows in df2 whose signature does not appear in df1 are "distinct"
    distinct_rows = df2[~sig2.isin(sig1)]
else:
    # If there are no shared columns, nothing can be considered "common".
    common_rows = df2.head(0).copy()
    distinct_rows = df2.copy()

# --- SAVE ONLY TWO FILES ---
# Write exactly two Excel files: one with the common rows and one with the distinct rows.
common_out = "rows_common_on_shared_columns.xlsx"
distinct_out = "rows_distinct_on_shared_columns_with_diffcols.xlsx"

# For each distinct row, determine which of the shared columns contain values
# that do not appear in the corresponding column of file1. This helps identify
# which columns make the row 'distinct'. The result is saved as a new column
# `distinct_columns` and the full distinct rows are written to a single output file.
if not distinct_rows.empty and common_cols:
    # Build sets of normalized values present in file1 for each shared column
    value_sets = {}
    for col in common_cols:
        s = df1[col].fillna("").astype(str).str.strip().str.lower()
        value_sets[col] = set(s.unique())

    def _distinct_columns_for_row(row):
        cols = []
        for col in common_cols:
            val = row.get(col, "")
            vnorm = ("" if pd.isna(val) else str(val)).strip().lower()
            if vnorm not in value_sets.get(col, set()):
                cols.append(col)
        return ",".join(cols)

    distinct_rows = distinct_rows.copy()
    distinct_rows["distinct_columns"] = distinct_rows.apply(_distinct_columns_for_row, axis=1)

common_rows.to_excel(common_out, index=False)
distinct_rows.to_excel(distinct_out, index=False)

# --- SUMMARY / FEEDBACK ---
print(f"Shared columns used for comparison: {common_cols}")
print(f"Common rows (saved): {len(common_rows)} -> {common_out}")
print(f"Distinct rows (saved): {len(distinct_rows)} -> {distinct_out}")
