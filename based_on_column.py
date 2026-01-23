import pandas as pd
import csv

# --- SCRIPT PURPOSE ---
# Read two tables (CSV or Excel), compare rows based on specific columns you choose,
# and produce exactly two output files:
#  - rows_common_on_shared_columns.xlsx  -> entire rows from file2 that match some row in file1 on selected columns
#  - rows_distinct_on_shared_columns.xlsx -> entire rows from file2 that do not match any row in file1 on selected columns

# --- INPUT FILES ---
# file1: reference file (rows here are treated as "existing")
# file2: file to check (rows here will be classified as common or distinct)
file1 = "customer_data_organized.xlsx"
file2 = "staff.xlsx"

# --- SHEETS (optional, used only for Excel files) ---
# If your inputs are Excel workbooks, set these to the sheet names to read.
sheet1 = "Sheet1"
sheet2 = "Sheet1"

# --- COLUMNS TO COMPARE ON ---
# Specify which column(s) you want to use for comparison.
# You can use:
#   - A single column as a string: 'customer_login'
#   - Multiple columns as a list: ['Name', 'Email']
#   - None to use all shared columns between the two files
compare_columns = 'customer_login'

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

# --- COMMON / DISTINCT ON SELECTED COLUMNS ---
# Determine which columns to use for comparison
if compare_columns is None:
    # Use all shared columns
    cols_to_compare = list(df1.columns.intersection(df2.columns))
else:
    # Use the specified columns (convert to list if it's a string)
    cols_to_compare = [compare_columns] if isinstance(compare_columns, str) else compare_columns

if cols_to_compare:
    def row_sig(df, cols):
        """
        Build a stable string signature for each row using the specified columns.
        - fillna("") avoids NaN vs empty-string mismatches
        - astype(str) ensures consistent type
        - strip whitespace from each value to avoid trivial differences
        - join with a separator unlikely to appear in actual data ("||")
        """
        if isinstance(cols, list) and len(cols) > 1:
            # Multiple columns: use apply with axis=1
            return (
                df[cols]
                .fillna("")
                .astype(str)
                .apply(lambda r: "||".join(v.strip() for v in r.values), axis=1)
            )
        else:
            # Single column: convert directly to string and strip
            col = cols[0] if isinstance(cols, list) else cols
            return df[col].fillna("").astype(str).str.strip()

    # create signature series for both dataframes using only the selected columns
    sig1 = row_sig(df1, cols_to_compare)
    sig2 = row_sig(df2, cols_to_compare)

    # Rows in df2 whose signature appears in df1 are "common"
    # (Output contains the ENTIRE row from df2, not just the comparison columns)
    common_rows = df2[sig2.isin(sig1)]

    # Rows in df2 whose signature does not appear in df1 are "distinct"
    # (Output contains the ENTIRE row from df2, not just the comparison columns)
    distinct_rows = df2[~sig2.isin(sig1)]
else:
    # If there are no shared columns, nothing can be considered "common".
    common_rows = df2.head(0).copy()
    distinct_rows = df2.copy()

# --- SAVE ONLY TWO FILES ---
# Write exactly two Excel files: one with the common rows and one with the distinct rows.
# Output contains ENTIRE rows (all columns) from file2, but comparison was based on selected columns only.
common_out = "rows_common_on_shared_columns.xlsx"
distinct_out = "rows_distinct_on_shared_columns.xlsx"

common_rows.to_excel(common_out, index=False)
distinct_rows.to_excel(distinct_out, index=False)

# --- SUMMARY / FEEDBACK ---
print(f"Columns used for comparison: {cols_to_compare}")
print(f"Common rows (saved): {len(common_rows)} -> {common_out}")
print(f"Distinct rows (saved): {len(distinct_rows)} -> {distinct_out}")
