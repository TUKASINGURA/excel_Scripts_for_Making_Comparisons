import pandas as pd
import csv
import re
import difflib

# --- INPUT FILES ---
file1 = "customer_data_organized.xlsx"   # base file
file2 = "staff.xlsx"   # file providing extra columns

sheet1 = "Sheet1"
sheet2 = "Sheet1"

# --- JOIN COLUMN ---
join_column = "customer_login"   # column to join on

# --- LOAD FUNCTIONS ---
def detect_sep(path):
    try:
        with open(path, "r", encoding="utf-8", newline="") as f:
            sample = f.read(4096)
            return csv.Sniffer().sniff(sample).delimiter
    except Exception:
        return ","

def read_table(path, sheet=None):
    lower = path.lower()
    if lower.endswith(('.xls', '.xlsx', '.xlsm')):
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    else:
        sep = detect_sep(path)
        df = pd.read_csv(path, sep=sep, engine="python", encoding="utf-8", skipinitialspace=True)

    # Normalize column names
    df.columns = df.columns.astype(str).str.replace("\ufeff", "").str.strip()
    return df

# --- LOAD DATA ---
df1 = read_table(file1, sheet1)
df2 = read_table(file2, sheet2)

# --- NORMALIZE / RESOLVE JOIN COLUMN ---
def _normalize(s: str) -> str:
    return re.sub(r"\W+", "_", s.strip().lower()).strip("_")

def _find_column(df, desired):
    desired_n = _normalize(desired)
    mapping = { _normalize(c): c for c in df.columns }
    # exact normalized match
    if desired_n in mapping:
        return mapping[desired_n]

    # try close matches on normalized names
    close = difflib.get_close_matches(desired_n, mapping.keys(), n=1, cutoff=0.6)
    if close:
        return mapping[close[0]]

    # try token overlap (common words like 'id' or 'vods')
    desired_tokens = set(desired_n.split("_"))
    best = None
    best_score = 0
    for norm, orig in mapping.items():
        tokens = set(norm.split("_"))
        score = len(desired_tokens & tokens)
        if score > best_score:
            best_score = score
            best = orig
    if best_score > 0:
        return best

    return None

col1 = _find_column(df1, join_column)
col2 = _find_column(df2, join_column)

if not col1:
    raise ValueError(f"'{join_column}' not found in {file1}. Available columns: {list(df1.columns[:20])}...")
if not col2:
    raise ValueError(f"'{join_column}' not found in {file2}. Available columns: {list(df2.columns[:20])}...")

if col1 != join_column:
    print(f"Using '{col1}' from {file1} as join column (matched to '{join_column}').")
if col2 != join_column:
    print(f"Using '{col2}' from {file2} as join column (matched to '{join_column}').")

# unify join column name in both dataframes
UNIFIED_JOIN = "__join_col__"
if col1 != UNIFIED_JOIN:
    df1 = df1.rename(columns={col1: UNIFIED_JOIN})
if col2 != UNIFIED_JOIN:
    df2 = df2.rename(columns={col2: UNIFIED_JOIN})

# use unified name from now on
join_column = UNIFIED_JOIN

# --- REMOVE DUPLICATES IN FILE2 (VERY IMPORTANT) ---
df2 = df2.drop_duplicates(subset=[join_column])

# Coerce join keys to strings and normalize empty/missing values to empty string
def _coerce_join_col(df, col):
    df[col] = df[col].apply(lambda x: '' if pd.isna(x) else str(x).strip())
    return df

df1 = _coerce_join_col(df1, join_column)
df2 = _coerce_join_col(df2, join_column)

# --- SELECT COLUMNS TO APPEND ---
# keep VOD_ID + all other columns from file2 except duplicates
columns_to_add = [c for c in df2.columns if c != join_column]

df2_subset = df2[[join_column] + columns_to_add]

# --- MERGE (APPEND DATA FROM FILE2 INTO FILE1) ---
merged_df = df1.merge(
    df2_subset,
    on=join_column,
    how="left",              # keep all rows from file1
    suffixes=("", "_from_file2")
)

# --- SAVE FINAL OUTPUT ---
output_file = "final_merged_output.xlsx"
merged_df.to_excel(output_file, index=False)

# --- SUMMARY ---
print("Merge completed successfully")
print(f"Join column: {join_column}")
print(f"Rows in file1: {len(df1)}")
print(f"Rows enriched from file2: {merged_df[join_column].isin(df2[join_column]).sum()}")
print(f"Final output saved as: {output_file}")
