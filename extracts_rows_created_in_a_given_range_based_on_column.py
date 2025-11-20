import pandas as pd
import csv
import re
from datetime import datetime

# --- INPUT FILE ---
file_path = "users.csv"
sheet_name = "Sheet1"

# --- DATE COLUMN NAME ---
date_column = "Created"    # <-- change to your actual column name

# --- DATE RANGE (CHANGE AS NEEDED) ---
#the Date format is DD/MM/YYYY 
#example1 :  start_date_str = "1/11/2025" , end_date_str   = "18/11/2025" 
#example2 :  start_date_str = "17/11/2025 2:39:44" , end_date_str   = "18/11/2025 14:20:00"

start_date_str = "28/8/2025 21:00:00"
end_date_str   = "28/8/2025 23:59:00"

#Note any data that is found to be at or after the end date/time will be excluded.

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
    This preserves the original text exactly as in the file so the filtered rows
    can be returned "the way they are".
    """
    lower = path.lower()
    if lower.endswith(('.xls', '.xlsx', '.xlsm')):
        # read Excel with dtype=str to preserve original cell text
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", dtype=str)
    else:
        sep = detect_sep(path)
        # read CSV with dtype=str and skipinitialspace to preserve exact text
        df = pd.read_csv(path, sep=sep, engine="python", encoding="utf-8", dtype=str, skipinitialspace=True)
    # normalize header names (remove BOM and surrounding whitespace)
    df.columns = df.columns.astype(str).str.replace("\ufeff", "").str.strip()
    return df

# --- LOAD THE TABLE (raw strings) ---
df_raw = read_table_as_strings(file_path, sheet=sheet_name)
print("Columns:", list(df_raw.columns))

if date_column not in df_raw.columns:
    raise KeyError(f"Column '{date_column}' not found. Available columns: {list(df_raw.columns)}")

# --- PARSE DATES FROM THE RAW TEXT (for filtering only) ---
# Convert the raw date/text values into datetimes for comparison; errors -> NaT
def parse_date_flexible(series):
    """
    Try fast vectorized parse first (dayfirst=True). For any remaining NaT
    try to extract a date substring and parse with several common formats,
    then fall back to pd.to_datetime with dayfirst.
    """
    parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
    if not parsed.isna().any():
        return parsed

    # formats to attempt when automatic parse fails
    fmt_candidates = [
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y",
    ]

    # try per-value fallback only for those still NaT
    for idx, raw in series[parsed.isna()].items():
        s = "" if raw is None else str(raw).strip()
        # extract a simple date substring if present (e.g. "05/11/2025")
        m = re.search(r"\d{1,2}[/-]\d{1,2}[/-]\d{4}", s)
        candidate = m.group(0) if m else s
        dt = pd.NaT
        for fmt in fmt_candidates:
            try:
                dt = pd.to_datetime(candidate, format=fmt, dayfirst=True, errors="raise")
                break
            except Exception:
                continue
        if pd.isna(dt):
            # last resort: let pandas try (handles more weird inputs)
            dt = pd.to_datetime(candidate, dayfirst=True, errors="coerce")
        parsed.at[idx] = dt
    return parsed

parsed_dates = parse_date_flexible(df_raw[date_column])

# Convert the input strings into datetime objects for the range
start_date = pd.to_datetime(start_date_str, dayfirst=True)
end_date = pd.to_datetime(end_date_str, dayfirst=True)

# --- FILTER BETWEEN START & END USING PARSED DATES ---
mask = (parsed_dates >= start_date) & (parsed_dates <= end_date)
filtered_raw = df_raw.loc[mask].copy()   # preserve original text/format exactly

# --- SAVE RESULT (rows returned the way they are) ---
output_file = "filtered_by_date_original_text.xlsx"
filtered_raw.to_excel(output_file, index=False, engine="openpyxl")

print(f"Done! {len(filtered_raw)} rows saved to: {output_file}")
