import pandas as pd
import csv
import re
from datetime import datetime

# --- INPUT FILE ---                                                                                                                   
file_path = "customer_data_organized.xlsx"
sheet_name = "Sheet1"

# --- DATE COLUMN NAME --- is the Column where the Date Ranges are located. This should be the same as the column in your input file that contains the start dates.
date_column = "latest_subscription_started"

# --- DATE RANGE ---
# Enhanced slightly to ensure we capture the full days
# Note: The original date format is "2026-02-04 00:00:00", so you can use ISO8601 parsing
start_date_str = "2026-02-04 00:00:00"
end_date_str   = "2026-02-04 23:59:59"

# --- HELPER: detect delimiter for CSV ---
# This is a simple heuristic that reads a sample of the file and uses csv.Sniffer to guess the delimiter. It defaults to comma if it fails.
def detect_sep(path):
    try:
        with open(path, "r", encoding="utf-8", newline="") as f:
            sample = f.read(4096)
            return csv.Sniffer().sniff(sample).delimiter
    except Exception:
        return ","

# --- HELPER: read Excel or CSV as raw strings ---
def read_table_as_strings(path, sheet=None):
    lower = path.lower()
    if lower.endswith(('.xls', '.xlsx', '.xlsm','.xlsb')):
        # We read as strings to keep the output format identical to the input
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", dtype=str)
    else:
        sep = detect_sep(path)
        df = pd.read_csv(path, sep=sep, engine="python", encoding="utf-8", dtype=str, skipinitialspace=True)
    
    df.columns = df.columns.astype(str).str.replace("\ufeff", "").str.strip()
    return df

# --- LOAD THE TABLE ---
# Reads the data as raw strings to preserve formatting/whitespace
df_raw = read_table_as_strings(file_path, sheet=sheet_name)

if date_column not in df_raw.columns:
    raise KeyError(f"Column '{date_column}' not found. Available columns: {list(df_raw.columns)}")

# --- PARSE DATES (Modern Robust Way) ---
# 'format="ISO8601"' is the fastest/safest for YYYY-MM-DD formats
# Removes 'dayfirst=True' because your year comes first
parsed_dates = pd.to_datetime(df_raw[date_column], errors="coerce", format="ISO8601")

# Convert filter strings to datetime objects
start_date = pd.to_datetime(start_date_str)
end_date = pd.to_datetime(end_date_str)

# --- FILTER ---
mask = (parsed_dates >= start_date) & (parsed_dates <= end_date)
filtered_raw = df_raw.loc[mask].copy()

# --- SAVE RESULT ---
output_file = "filtered_by_date_original_text.xlsx"
if not filtered_raw.empty:
    filtered_raw.to_excel(output_file, index=False, engine="openpyxl")
    print(f"Success! {len(filtered_raw)} rows found and saved to: {output_file}")
else:
    # Create an empty file with headers so the NEXT script doesn't crash
    df_raw.iloc[0:0].to_excel(output_file, index=False, engine="openpyxl")
    print("Warning: No rows matched that date range. An empty file with headers was created.")

# --- DEBUG SUMMARY ---
# This summary will help you understand how many rows were processed and if the date parsing worked correctly.
print(f"Total rows scanned: {len(df_raw)}")
print(f"Rows with valid dates: {parsed_dates.notna().sum()}")
if parsed_dates.notna().sum() > 0:
    print("Earliest date found:", parsed_dates.min())
    print("Latest date found:", parsed_dates.max())
